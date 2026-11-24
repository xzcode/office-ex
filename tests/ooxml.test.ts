// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { Zip, ZipDeflate } from 'fflate';
import {
  readZip, writeZip, crc32, parseXml, serializeXml, docxParts, xlsxParts,
  openOpc, writeOpc, resolvePartTarget,
} from '@office-ex/ooxml';
import { OfficeError } from '@office-ex/contracts';
const encode = (text: string): Uint8Array => new TextEncoder().encode(text);
const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);
const mapParts = (parts: Record<string, string>): Map<string, Uint8Array> => new Map(Object.entries(parts).map(([path, text]) => [path, encode(text)]));
const isCode = (code: string) => (error: unknown): boolean => error instanceof OfficeError && error.code === code;
function centralOffset(bytes: Uint8Array): number { return new DataView(bytes.buffer, bytes.byteOffset).getUint32(bytes.length - 6, true); }
function docx(): Map<string, Uint8Array> { return mapParts(docxParts([{ runs: [{ text: '中文 & 😀', bold: true }] }])); }

test('ZIP STORE/DEFLATE round trip and CRC fixture', () => {
  const parts = new Map([['foo/中文.txt', encode('中文 & hello'.repeat(100))], ['empty', new Uint8Array()], ['folder/', new Uint8Array()]]);
  for (const compress of [false, true]) assert.deepEqual(readZip(writeZip(parts, { compress })), parts);
  assert.equal(crc32(encode('123456789')), 0xcbf43926);
});
test('Python zipfile independently reads our archive and creates a DEFLATE input', () => {
  const input = writeZip(new Map([['中文/a.txt', encode('hello 中文')], ['empty', new Uint8Array()]]));
  const script = `import sys,io,zipfile,base64\nb=base64.b64decode(sys.stdin.buffer.read())\nz=zipfile.ZipFile(io.BytesIO(b))\nassert z.testzip() is None\nassert z.read('中文/a.txt').decode()=='hello 中文'\nassert z.read('empty')==b''\nout=io.BytesIO()\nwith zipfile.ZipFile(out,'w',compression=zipfile.ZIP_DEFLATED) as z:\n z.writestr('from-python.txt','Independent ZIP writer 中文')\nprint(base64.b64encode(out.getvalue()).decode())\n`;
  const result = spawnSync(process.env.PYTHON ?? 'python3', ['-c', script], { input: Buffer.from(input).toString('base64'), encoding: 'utf8' });
  assert.equal(result.status, 0, result.stderr);
  assert.equal(decode(readZip(new Uint8Array(Buffer.from(result.stdout.trim(), 'base64'))).get('from-python.txt')!), 'Independent ZIP writer 中文');
});
test('ZIP data descriptors are checked rather than skipped', () => {
  const chunks: Uint8Array[] = [];
  const archive = new Zip((error, chunk) => { assert.ifError(error); chunks.push(chunk); });
  const item = new ZipDeflate('data.txt'); archive.add(item); item.push(encode('descriptor data'), true); archive.end();
  const bytes = new Uint8Array(Buffer.concat(chunks));
  assert.equal(decode(readZip(bytes).get('data.txt')!), 'descriptor data');
  const corrupt = bytes.slice(), descriptor = centralOffset(bytes) - 12;
  corrupt[descriptor] = corrupt[descriptor]! ^ 1;
  assert.throws(() => readZip(corrupt), isCode('invalid'));
});
test('ZIP path, truncation, headers, symlinks and CRC attacks are rejected', () => {
  for (const path of ['../escape', '/absolute', 'a/../b', 'a\\b', 'C:/x', 'a//b', './x']) assert.throws(() => writeZip({ [path]: encode('x') }), isCode('unsafe'));
  const bytes = writeZip({ 'a.txt': encode('original') }, { compress: false });
  assert.throws(() => readZip(bytes.subarray(0, bytes.length - 1)), isCode('invalid'));
  const data = bytes.slice(); data[35] = data[35]! ^ 1;
  assert.throws(() => readZip(data), isCode('invalid'));
  const mismatch = bytes.slice(); new DataView(mismatch.buffer).setUint32(18, 2, true);
  assert.throws(() => readZip(mismatch), isCode('invalid'));
  const symlink = bytes.slice(); new DataView(symlink.buffer).setUint32(centralOffset(bytes) + 38, 0xa1ff0000, true);
  assert.throws(() => readZip(symlink), isCode('unsafe'));
  const encrypted = bytes.slice(); new DataView(encrypted.buffer).setUint16(centralOffset(bytes) + 8, 0x801, true);
  assert.throws(() => readZip(encrypted), isCode('unsafe'));
});
test('ZIP checks declared and actual expansion, cumulative budgets and cancellation', () => {
  const bytes = writeZip({ 'a.txt': encode('abc'.repeat(300)), 'b.txt': encode('def'.repeat(200)) });
  assert.throws(() => readZip(bytes, { limits: { maxEntryBytes: 800 } }), isCode('limit'));
  assert.throws(() => readZip(bytes, { limits: { maxTotalBytes: 1400 } }), isCode('limit'));
  assert.throws(() => readZip(bytes, { limits: { maxCompressionRatio: 1 } }), isCode('limit'));
  assert.throws(() => readZip(bytes, { signal: { aborted: true } }), isCode('cancelled'));
  const forged = bytes.slice(), view = new DataView(forged.buffer);
  view.setUint32(22, 10, true); view.setUint32(centralOffset(bytes) + 24, 10, true);
  assert.throws(() => readZip(forged), isCode('invalid'));
  let polls = 0;
  assert.throws(() => readZip(writeZip({ 'many.bin': new Uint8Array(200_000) }, { compress: false }), { signal: { get aborted() { return ++polls > 4; } } }), isCode('cancelled'));
});
test('ZIP overlapping records, duplicate names and hidden bytes are rejected', () => {
  const bytes = writeZip({ 'a.txt': encode('x'), 'b.txt': encode('y') }, { compress: false });
  const start = centralOffset(bytes), second = start + 51;
  const overlap = bytes.slice(); new DataView(overlap.buffer).setUint32(second + 42, 0, true);
  assert.throws(() => readZip(overlap), isCode('invalid'));
  const duplicate = bytes.slice(); duplicate[second + 46] = 97;
  assert.throws(() => readZip(duplicate), isCode('invalid'));
  const hidden = new Uint8Array(bytes.length + 1); hidden.set(bytes, 1);
  const hiddenView = new DataView(hidden.buffer); hiddenView.setUint32(hidden.length - 6, start + 1, true);
  hiddenView.setUint32(start + 1 + 42, 1, true); hiddenView.setUint32(second + 1 + 42, 37, true);
  assert.throws(() => readZip(hidden), isCode('invalid'));
});
test('XML namespace identity, UTF-16 and serializer preserve unknown XML', () => {
  const xml = '<p:root xmlns:p="urn:test" xmlns:q="urn:extension" q:flag="a &amp; b"><q:unknown>中文😀</q:unknown></p:root>';
  const document = parseXml(xml);
  assert.equal(document.documentElement!.namespaceURI, 'urn:test');
  assert.equal(document.getElementsByTagNameNS('urn:extension', 'unknown').item(0)!.textContent, '中文😀');
  assert.equal(parseXml(serializeXml(document)).documentElement!.getAttributeNS('urn:extension', 'flag'), 'a & b');
  const utf16 = Buffer.from('\ufeff<?xml version="1.0" encoding="UTF-16"?><a>中文</a>', 'utf16le');
  assert.equal(parseXml(utf16).documentElement!.textContent, '中文');
  assert.equal(parseXml(serializeXml(parseXml(utf16))).documentElement!.textContent, '中文');
  assert.throws(() => parseXml(encode('<?xml version="1.0" encoding="UTF-16"?><a/>')), isCode('invalid'));
});
test('XML rejects DTD, custom entities, XInclude, invalid encodings and recovered malformed documents', () => {
  for (const xml of ['<!DOCTYPE a [<!ENTITY x "boom">]><a>&x;</a>', '<a>&custom;</a>', '<a xmlns:x="http://www.w3.org/2001/XInclude"><x:include href="file:///etc/passwd"/></a>']) {
    assert.throws(() => parseXml(xml), isCode('unsafe'));
  }
  for (const xml of ['<a><b></a>', '<a/><b/>', '<a x="1" x="2"/>', '<a></a junk>', '<a>&#0;</a>', '<a>\u0000</a>', '<a/><!---->text', '<p:a/>', '<a>abc]]></a>', '<a / >', '<a xmlns:p="urn:x" xmlns:q="urn:x" p:b="1" q:b="2"/>']) assert.throws(() => parseXml(xml), isCode('invalid'));
  assert.throws(() => parseXml('<?xml version="1.0" encoding="ISO-8859-1"?><a/>'), isCode('unsupported'));
  assert.throws(() => parseXml(new Uint8Array([60, 97, 62, 0xff, 60, 47, 97, 62])), isCode('invalid'));
  // Harmless strings in a comment are not confused with declarations.
  assert.equal(parseXml('<a><!-- <!DOCTYPE not-a-declaration> --><![CDATA[&not-an-entity;]]></a>').documentElement!.textContent, '&not-an-entity;');
});
test('XML enforces depth, attribute/node, byte and cancellation budgets before DOM construction', () => {
  assert.throws(() => parseXml('<a>'.repeat(5) + '</a>'.repeat(5), { limits: { maxXmlDepth: 4 } }), isCode('limit'));
  assert.throws(() => parseXml('<a x="1" y="2"/>', { limits: { maxXmlNodes: 2 } }), isCode('limit'));
  assert.throws(() => parseXml('<a>中文</a>', { limits: { maxXmlBytes: 10 } }), isCode('limit'));
  assert.throws(() => parseXml('<a/>', { signal: { aborted: true } }), isCode('cancelled'));
});
test('OPC detects document kind from relationships and content types, independent of prefixes and main path', () => {
  const parts = docx();
  const main = decode(parts.get('word/document.xml')!).replaceAll('w:', 'alt:').replace('xmlns:w=', 'xmlns:alt=');
  parts.delete('word/document.xml'); parts.set('custom/main.xml', encode(main));
  parts.set('[Content_Types].xml', encode(decode(parts.get('[Content_Types].xml')!).replace('word/document.xml', 'custom/main.xml')));
  parts.set('_rels/.rels', encode(decode(parts.get('_rels/.rels')!).replace('word/document.xml', 'custom/main.xml')));
  const package_ = openOpc(writeZip(parts));
  assert.equal(package_.kind, 'docx'); assert.equal(package_.mainPart, 'custom/main.xml'); assert.equal(package_.conformance, 'transitional');
  assert.equal(openOpc(writeZip(mapParts(xlsxParts([{ name: 'Sheet', cells: [] }])))).kind, 'xlsx');
});
test('OPC resolves relative and case-equivalent targets and reports external relationships without fetching', () => {
  const parts = docx();
  parts.set('word/_rels/document.xml.rels', encode('<r:Relationships xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships"><r:Relationship Id="rIdImage" Type="urn:image" Target="../assets/IMAGE.bin"/><r:Relationship Id="rIdExternal" Type="urn:link" TargetMode="External" Target="https://invalid.example/never-request"/></r:Relationships>'));
  parts.set('assets/image.bin', new Uint8Array([1, 2, 3]));
  parts.set('[Content_Types].xml', encode(decode(parts.get('[Content_Types].xml')!).replace('</Types>', '<Default Extension="bin" ContentType="application/octet-stream"/></Types>')));
  const package_ = openOpc(writeZip(parts));
  assert.equal(package_.relationships.get('word/document.xml')![0]!.targetPart, 'assets/image.bin');
  assert.equal(package_.externalRelationships, 1);
  assert.deepEqual(resolvePartTarget('word/document.xml', '../assets/image.bin#anchor'), { path: 'assets/image.bin', fragment: 'anchor' });
  assert.throws(() => resolvePartTarget('', '../outside'), isCode('unsafe'));
  assert.throws(() => resolvePartTarget('', '%2e%2e/escape'), isCode('unsafe'));
  assert.throws(() => resolvePartTarget('', 'https://example.com/x'), isCode('unsafe'));
});
test('OPC rejects relationship collisions, missing targets, equivalent names and malformed XML', () => {
  const base = docx();
  const missing = new Map(base); missing.delete('word/document.xml'); assert.throws(() => openOpc(writeZip(missing)), isCode('invalid'));
  const duplicateIds = new Map(base);
  duplicateIds.set('_rels/.rels', encode(decode(base.get('_rels/.rels')!).replace('</Relationships>', '<Relationship Id="rId1" Type="urn:duplicate" Target="word/document.xml"/></Relationships>')));
  assert.throws(() => openOpc(writeZip(duplicateIds)), isCode('invalid'));
  const names = new Map(base); names.set('WORD/document.xml', base.get('word/document.xml')!);
  assert.throws(() => openOpc(writeZip(names)), isCode('invalid'));
  const malformed = new Map(base); malformed.set('word/document.xml', encode('<unclosed>'));
  assert.throws(() => openOpc(writeZip(malformed)), isCode('invalid'));
});
test('OPC identifies Strict and preserves untouched unknown parts and the no-edit archive', () => {
  const parts = docx();
  parts.set('word/vendor.xml', encode('<v:opaque xmlns:v="urn:vendor" mark="x"><v:data> untouched </v:data></v:opaque>'));
  const bytes = writeZip(parts), package_ = openOpc(bytes);
  assert.deepEqual(writeOpc(package_), bytes);
  const modified = encode(decode(parts.get('word/document.xml')!).replace('中文', '修改'));
  const saved = readZip(writeOpc(package_, new Map([['word/document.xml', modified]])));
  assert.deepEqual(saved.get('word/vendor.xml'), parts.get('word/vendor.xml'));
  assert.deepEqual(saved.get('[Content_Types].xml'), parts.get('[Content_Types].xml'));
  const strict = new Map(parts);
  strict.set('_rels/.rels', encode(decode(strict.get('_rels/.rels')!).replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument')));
  strict.set('word/document.xml', encode(decode(strict.get('word/document.xml')!).replace('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'http://purl.oclc.org/ooxml/wordprocessingml/main')));
  assert.equal(openOpc(writeZip(strict)).conformance, 'strict');
});
test('OPC blocks uncontrolled source mutation and signed-package edits', () => {
  const package_ = openOpc(writeZip(docx())); package_.parts.get('word/document.xml')![0] = 0;
  assert.throws(() => writeOpc(package_), isCode('conflict'));
  const mutatedBytes = openOpc(writeZip(docx())); mutatedBytes.originalBytes[0] = 0;
  assert.throws(() => writeOpc(mutatedBytes), isCode('conflict'));
  const parts = docx(); parts.set('_xmlsignatures/origin.sigs', new Uint8Array());
  parts.set('[Content_Types].xml', encode(decode(parts.get('[Content_Types].xml')!).replace('</Types>', '<Default Extension="sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/></Types>')));
  const signed = openOpc(writeZip(parts)); assert.equal(signed.hasSignatures, true);
  signed.hasSignatures = false; // Caller mutation must not bypass the adapter's original evidence.
  assert.throws(() => writeOpc(signed, new Map([['word/document.xml', parts.get('word/document.xml')!]])), isCode('unsafe'));
});
