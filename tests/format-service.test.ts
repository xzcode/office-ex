// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { OfficeError, type Diagnostic } from '@office-ex/contracts';
import { createDocument } from '@office-ex/core';
import { BoundedReader, planPreservingSave, sniffContainer, sortedDiagnostics } from '@office-ex/format-common';
import { inspectFile, parseDraft, serializeDraft } from '@office-ex/format-service';
import { writeCfb } from '@office-ex/cfb';
import { docxParts, xlsxParts, writeZip } from '@office-ex/ooxml';

const encoder = new TextEncoder();
const parts = (strings: Record<string, string>): Map<string, Uint8Array> => new Map(Object.entries(strings).map(([name, value]) => [name, encoder.encode(value)]));
const cfb = (streams: Record<string, Uint8Array>): Uint8Array => writeCfb({ entries: Object.entries(streams).map(([name, data]) => ({ type: 'stream', path: [name], data })) });
const view = (bytes: Uint8Array): DataView => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
const code = (value: string) => (error: unknown): boolean => error instanceof OfficeError && error.code === value;
const digest = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex');

/** Header-only synthetic recognizer fixtures. NOT valid Office documents or interoperability evidence. */
function docHeader(nFib = 0xc1): Uint8Array {
  const bytes = new Uint8Array(32), v = view(bytes);
  v.setUint16(0, 0xa5ec, true); v.setUint16(2, nFib, true); v.setUint16(10, 0x1000, true); v.setUint16(12, 0xbf, true);
  return bytes;
}
function xlsGlobals(records: { type: number; data: Uint8Array }[] = []): Uint8Array {
  const bytes = new Uint8Array(24 + records.reduce((size, record) => size + 4 + record.data.length, 0)), v = view(bytes);
  v.setUint16(0, 0x0809, true); v.setUint16(2, 16, true); v.setUint16(4, 0x600, true); v.setUint16(6, 5, true);
  v.setUint16(10, 0x7cc, true); v.setUint32(12, 1, true); v.setUint32(16, 6, true);
  let offset = 20;
  for (const record of records) { v.setUint16(offset, record.type, true); v.setUint16(offset + 2, record.data.length, true); bytes.set(record.data, offset + 4); offset += 4 + record.data.length; }
  v.setUint16(offset, 0x000a, true);
  return bytes;
}
function pptHeaders(encrypted = false): { main: Uint8Array; user: Uint8Array } {
  const user = new Uint8Array(32), u = view(user), main = new Uint8Array(encrypted ? 48 : 44), m = view(main), edit = 8;
  u.setUint16(2, 0x0ff6, true); u.setUint32(4, 24, true); u.setUint32(8, 20, true);
  u.setUint32(12, encrypted ? 0xf3d1c4df : 0xe391c05f, true); u.setUint32(16, edit, true); u.setUint16(22, 0x3f4, true); u.setUint8(24, 3); u.setUint32(28, 8, true);
  m.setUint16(edit + 2, 0x0ff5, true); m.setUint32(edit + 4, encrypted ? 32 : 28, true); m.setUint8(edit + 15, 3); m.setUint32(edit + 24, 1, true);
  return { main, user };
}
function pptxHeader(): Uint8Array {
  return writeZip(parts({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/custom/deck.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>',
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rMain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/deck.xml"/></Relationships>',
    'custom/deck.xml': '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
  }));
}
test('FMT-01/M1: detects six format envelopes, preserves source hashes, leaves semantic capabilities C0', async () => {
  const ppt = pptHeaders();
  const fixtures = [
    ['docx', writeZip(parts(docxParts([{ runs: [{ text: 'test' }] }])))],
    ['xlsx', writeZip(parts(xlsxParts([{ name: 'Sheet1', cells: [] }])))],
    ['pptx', pptxHeader()],
    ['doc', cfb({ WordDocument: docHeader(), '0Table': new Uint8Array() })],
    ['xls', cfb({ Workbook: xlsGlobals() })],
    ['ppt', cfb({ 'PowerPoint Document': ppt.main, 'Current User': ppt.user })],
  ] as const;
  for (const [format, bytes] of fixtures) {
    const before = digest(bytes), result = await inspectFile(bytes, `fixture.${format}`);
    assert.equal(result.detection.format, format);
    assert.equal(result.sha256, before); assert.equal(digest(bytes), before);
    assert.equal(result.report.readOnly, true);
    const semantic = result.report.features.find(feature => feature.featureId === 'FMT-04')!;
    assert.equal(semantic.level, 'C0'); assert.equal(semantic.read, false); assert.equal(semantic.write, false); assert.equal(semantic.blocked, true);
    assert.ok(result.report.diagnostics.some(item => item.code === 'SEMANTIC_ADAPTER_PENDING'));
    assert.ok(result.entries.every(entry => /^[a-f0-9]{64}$/.test(entry.sha256)));
  }
});
test('FMT-01: extension does not select parser; CFB lookup respects case-equivalent root names', async () => {
  const xlsx = writeZip(parts(xlsxParts([{ name: 'Sheet1', cells: [] }])));
  const renamed = await inspectFile(xlsx, 'misleading.doc');
  assert.equal(renamed.detection.format, 'xlsx');
  assert.ok(renamed.detection.diagnostics.some(item => item.code === 'EXTENSION_MISMATCH'));
  const result = await inspectFile(cfb({ worddocument: docHeader(), '0table': new Uint8Array() }));
  assert.equal(result.detection.format, 'doc'); assert.ok(result.entries.some(entry => entry.name === 'worddocument'));
});
test('DOC detector rejects unknown/pre97 FIB versions, missing table streams and ambiguous root types', async () => {
  for (const version of [0x65, 0xffff]) await assert.rejects(inspectFile(cfb({ WordDocument: docHeader(version), '0Table': new Uint8Array() })), code('unsupported'));
  await assert.rejects(inspectFile(cfb({ WordDocument: docHeader() })), code('invalid'));
  await assert.rejects(inspectFile(cfb({ WordDocument: docHeader(), Workbook: xlsGlobals(), '0Table': new Uint8Array() })), code('invalid'));
});
test('XLS detector bounds BOF and EOF records and identifies encryption without executing content', async () => {
  const bof = xlsGlobals(); view(bof).setUint16(2, 17, true);
  await assert.rejects(inspectFile(cfb({ Workbook: bof })), code('unsupported'));
  const eof = new Uint8Array(25); eof.set(xlsGlobals()); view(eof).setUint16(22, 1, true);
  await assert.rejects(inspectFile(cfb({ Workbook: eof })), code('invalid'));
  await assert.rejects(inspectFile(cfb({ Workbook: xlsGlobals().subarray(0, 20) })), code('invalid'));
  const encrypted = await inspectFile(cfb({ Workbook: xlsGlobals([{ type: 0x2f, data: new Uint8Array(4) }]) }));
  assert.equal(encrypted.detection.encrypted, true);
  assert.ok(encrypted.report.diagnostics.some(item => item.code === 'ENCRYPTED_UNSUPPORTED'));
});
test('PPT detector enforces CurrentUser and UserEdit record lengths and versions', async () => {
  const make = (edit: (main: Uint8Array, user: Uint8Array) => void): Uint8Array => { const { main, user } = pptHeaders(); edit(main, user); return cfb({ 'PowerPoint Document': main, 'Current User': user }); };
  for (const bytes of [
    make((_main, user) => view(user).setUint32(4, 1, true)),
    make((_main, user) => view(user).setUint16(20, 256, true)),
    make((_main, user) => view(user).setUint32(28, 0, true)),
    make((main) => view(main).setUint32(12, 0, true)),
    make((main) => view(main).setUint8(23, 4)),
  ]) await assert.rejects(inspectFile(bytes), code('invalid'));
  const encrypted = pptHeaders(true);
  assert.equal((await inspectFile(cfb({ 'PowerPoint Document': encrypted.main, 'Current User': encrypted.user }))).detection.encrypted, true);
});
test('encrypted OOXML requires both CFB root streams and enough cleartext header bytes', async () => {
  const incomplete: Record<string, Uint8Array>[] = [{ EncryptionInfo: new Uint8Array(8) }, { EncryptedPackage: new Uint8Array(8) }, { EncryptionInfo: new Uint8Array(7), EncryptedPackage: new Uint8Array(8) }];
  for (const streams of incomplete) {
    await assert.rejects(inspectFile(cfb(streams)), code('invalid'));
  }
  const result = await inspectFile(cfb({ EncryptionInfo: new Uint8Array(8), EncryptedPackage: new Uint8Array(8) }), 'encrypted.docx');
  assert.equal(result.detection.format, 'cfb'); assert.equal(result.detection.encrypted, true);
  assert.equal(result.detection.version, 'encrypted OOXML');
});
test('Office EX draft saves validated models and rejects unknown schema or envelope fields', async () => {
  const model = createDocument('writer', 'draft:test', 'Draft');
  const bytes = serializeDraft(model);
  assert.deepEqual(parseDraft(bytes), model);
  const result = await inspectFile(bytes, 'test.oex');
  assert.equal(result.report.readOnly, false); assert.equal(result.detection.format, 'oex');
  const envelope = JSON.parse(new TextDecoder().decode(bytes));
  envelope.version = 2;
  assert.throws(() => parseDraft(encoder.encode(JSON.stringify(envelope))), code('unsupported'));
  envelope.version = 1; envelope.extra = true;
  assert.throws(() => parseDraft(encoder.encode(JSON.stringify(envelope))), code('invalid'));
});
test('format boundary utilities reject out-of-bounds binary input and unsafe save plans', () => {
  const reader = new BoundedReader(new Uint8Array(4));
  for (const [offset, length] of [[-1, 1], [0, 5], [Number.MAX_SAFE_INTEGER, 2], [1.5, 1]]) assert.throws(() => reader.range(offset!, length!), code('invalid'));
  assert.equal(sniffContainer(encoder.encode('fake.docx')), 'unknown');
  const input = { sourceFormat: 'docx' as const, targetFormat: 'docx' as const, sourceContentId: 'a', currentContentId: 'a', canSerialize: false, hasUnknownDependencies: true };
  assert.equal(planPreservingSave(input).mode, 'copy-original'); assert.equal(planPreservingSave(input).allowed, true);
  assert.equal(planPreservingSave({ ...input, currentContentId: 'b' }).allowed, false);
  assert.equal(planPreservingSave({ ...input, targetFormat: 'doc' }).allowed, false);
});
test('diagnostics share the contracts DTO and sort without locale-sensitive collation', () => {
  const diagnostics: Diagnostic[] = ['z', 'A', 'a'].map(code => ({ code, featureId: 'FMT-01', severity: 'warning', messageKey: 'test', safeAction: 'none', sourceLocation: { part: 'word/document.xml', offset: 3 } }));
  assert.deepEqual(sortedDiagnostics(diagnostics).map(item => item.code), ['A','a','z']);
  assert.deepEqual(diagnostics.map(item => item.code), ['z','A','a']);
});
