// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { OfficeError, type Diagnostic } from '@office-ex/contracts';
import { createDocument } from '@office-ex/core';
import { BoundedReader, planPreservingSave, sniffContainer, sortedDiagnostics } from '@office-ex/format-common';
import { inspectFile, inspectForEditor, parseDraft, serializeDraft, prepareWriterDocxSave } from '@office-ex/format-service';
import { writeCfb } from '@office-ex/cfb';
import { docxParts, xlsxParts, writeZip, writeWriterDocx, openWriterDocx, openOpc } from '@office-ex/ooxml';

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

function editorFixture(id:string,text:string){
  const model=createDocument('writer',id,'Editor fixture');
  model.sections[0]!.blocks=[{id:'p',type:'paragraph',runs:[{id:'r',type:'text',text,style:{bold:true}}]}];
  return model;
}
test('WRT-04 desktop inspection enables only the narrow DOCX profile and correct source extension',async()=>{
  const bytes=writeWriterDocx(editorFixture('source','中文 日本語 😀'));
  const structural=await inspectFile(bytes,'file.docx');
  assert.equal(structural.report.readOnly,true);assert.equal(structural.editableWriter,undefined);
  const result=await inspectForEditor(bytes,'file.docx','tab-1');
  assert.equal(result.editableWriter?.documentId,'tab-1');
  assert.equal(result.report.readOnly,false);
  assert.ok(result.report.features.some(f=>f.featureId==='WRT-04'&&f.read&&f.write));
  assert.equal(result.report.features.find(f=>f.featureId==='FMT-04')?.level,'C0');
  assert.equal((await inspectForEditor(bytes,'file.doc','tab-2')).editableWriter,undefined);
  assert.equal((await inspectForEditor(bytes,'file','tab-2')).report.readOnly,true);
  const unknown=writeZip(parts(docxParts([{runs:[{text:'outside narrow profile'}]}])));
  assert.equal((await inspectForEditor(unknown,'unknown.docx','tab-3')).report.readOnly,true);
});
test('WRT-04 save verifier creates real DOCX from snapshots; unchanged source is byte-exact and edits preserve the OPC envelope',async()=>{
  const model=editorFixture('tab-1','before 😀');
  const generated=prepareWriterDocxSave(serializeDraft(model),'tab-1');
  const result=await inspectForEditor(generated,'source.docx','tab-2');
  assert.ok(result.editableWriter);
  const current=structuredClone(result.editableWriter);
  const unchanged=prepareWriterDocxSave(serializeDraft(current),'tab-2',generated);
  assert.deepEqual(unchanged,generated);
  const p=current.sections[0]!.blocks[0]!;
  if(p.type!=='paragraph'||p.runs[0]?.type!=='text')throw Error('expected run');
  p.runs[0].text='edited 中文';
  const updated=prepareWriterDocxSave(serializeDraft(current),'tab-2',generated);
  const a=openOpc(generated),b=openOpc(updated);
  for(const [path,data] of a.parts)if(path!==a.mainPart)assert.deepEqual(b.parts.get(path),data);
  const read=openWriterDocx(updated,'tab-2','source.docx').document.sections[0]!.blocks[0]!;
  assert.equal(read.type,'paragraph');if(read.type!=='paragraph'||read.runs[0]?.type!=='text')throw Error();
  assert.equal(read.runs[0].text,'edited 中文');assert.equal(read.runs[0].style?.bold,true);
});
test('WRT-04 save verifier rejects raw archive input, mismatched model IDs, other document kinds and unsupported source',()=>{
  const model=editorFixture('tab-1','safe');
  const snapshot=serializeDraft(model);
  assert.throws(()=>prepareWriterDocxSave(snapshot,'other'),code('invalid'));
  assert.throws(()=>prepareWriterDocxSave(serializeDraft(createDocument('sheets','tab-1')),'tab-1'),code('invalid'));
  const archive=writeWriterDocx(model);
  assert.throws(()=>prepareWriterDocxSave(archive,'tab-1'),code('invalid'));
  const unsupported=writeZip(parts(docxParts([{runs:[{text:'unsupported original'}]}])));
  const before=digest(unsupported);
  assert.throws(()=>prepareWriterDocxSave(snapshot,'tab-1',unsupported),code('unsupported'));
  assert.equal(digest(unsupported),before);
});

test('M7.5 OPC inspection pins limits, reuses only one private validated graph, and retains exact hashes',async()=>{
 const {prepareOpcInspection,inspectOpc,writeOpc}=await import('@office-ex/ooxml');
 const source=writeWriterDocx(createDocument('writer','inspection-scope'));
 const scope=prepareOpcInspection(source);
 assert.equal(scope.summary.sha256,digest(source));assert.ok(Object.isFrozen(scope.options)&&Object.isFrozen(scope.options.limits));
 const graph=openOpc(source,scope.options);
 for(const e of scope.summary.entries){const part=graph.parts.get(e.name)!;assert.equal(e.size,part.length);assert.equal(e.sha256,digest(part));}
 assert.deepEqual(writeOpc(graph),source);
 graph.parts.delete(graph.mainPart!);
 // A consumed scope never shares a caller-mutated package with a later open.
 const fresh=openOpc(source,scope.options);assert.ok(fresh.parts.has(fresh.mainPart!));
 assert.deepEqual(inspectOpc(source).entries,scope.summary.entries);scope.release();
 const changed=source.slice(),other=prepareOpcInspection(changed);changed[0]^=255;
 assert.throws(()=>openOpc(changed,other.options),code('conflict'));
 const controller=new AbortController(),cancelled=prepareOpcInspection(source,{signal:controller.signal});controller.abort();
 assert.throws(()=>openOpc(source,cancelled.options),code('cancelled'));cancelled.release();
 assert.throws(()=>prepareOpcInspection(source,{limits:{maxFileBytes:1}}),code('limit'));
});
test('M7.5 scoped editor inspection reports the exact source digest and structural part digests',async()=>{
 const {writeSheetsXlsx,writeSlidesPptx}=await import('@office-ex/ooxml');
 const docs=[['docx',writeWriterDocx(createDocument('writer','m75-word'))],['xlsx',writeSheetsXlsx(createDocument('sheets','m75-sheet'))],['pptx',writeSlidesPptx(createDocument('slides','m75-slides'))]] as const;
 for(const [ext,b]of docs){
  const structural=await inspectFile(b,'test.'+ext),editor=await inspectForEditor(b,'test.'+ext,'m75-open');
  assert.equal(editor.report.readOnly,false,editor.semanticFailure?.detail);
  assert.equal(editor.sha256,digest(b));assert.deepEqual(editor.entries,structural.entries);
 }
 // Changing bytes while an async inspection is yielding must not let a
 // previously validated private graph authorize a different byte stream.
 const source=docs[0][1].slice(),pending=inspectForEditor(source,'test.docx','m75-changed');
 source[0]^=255;await assert.rejects(pending,code('conflict'));
});
