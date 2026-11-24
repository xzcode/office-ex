import test from 'node:test';
import assert from 'node:assert/strict';
import { OfficeError } from '@office-ex/contracts';
import { readCfb, writeCfb, findCfbEntry, isCfb, compareCfbNames, type CfbDocument } from '@office-ex/cfb';

const END = 0xfffffffe;
const FREE = 0xffffffff;
const pattern = (size: number, seed = 17): Uint8Array => Uint8Array.from({ length: size }, (_, index) => (index * seed + 31) & 255);
const fixture = (): CfbDocument => ({
  root: { clsid: '0102030405060708090a0b0c0d0e0f10', stateBits: 0x12345678, modifiedTime: '01bab44b13921e80' },
  entries: [
    { type: 'storage', path: ['Unknown'], metadata: { clsid: '102030405060708090a0b0c0d0e0f000', stateBits: 8, creationTime: '01bab44b13921e80', modifiedTime: '01bab44b13921e81' } },
    { type: 'storage', path: ['Unknown', '空目录'] },
    { type: 'stream', path: ['Unknown', 'Small'], data: pattern(65) },
    { type: 'stream', path: ['Unknown', 'Boundary'], data: pattern(4095) },
    { type: 'stream', path: ['Large'], data: pattern(8193) },
    { type: 'stream', path: ['Exact'], data: pattern(4096) },
    { type: 'stream', path: ['Empty'], data: new Uint8Array() },
    { type: 'stream', path: ['\u0005SummaryInformation'], data: pattern(5) },
  ],
});
function equivalent(source: CfbDocument, result: CfbDocument): void {
  assert.equal(result.entries.length, source.entries.length);
  for (const entry of source.entries) {
    const actual = findCfbEntry(result, entry.path);
    assert.ok(actual);
    assert.equal(actual.type, entry.type);
    if (actual.type === 'stream' && entry.type === 'stream') assert.deepEqual(actual.data, entry.data);
    for (const [key, value] of Object.entries(entry.metadata ?? {})) assert.equal(actual.metadata?.[key as keyof NonNullable<typeof actual.metadata>], value);
  }
  for (const [key, value] of Object.entries(source.root ?? {})) assert.equal(result.root?.[key as keyof NonNullable<typeof result.root>], value);
}
function rejected(work: () => unknown, code: string = 'invalid'): void {
  assert.throws(work, (error: unknown) => error instanceof OfficeError && error.code === code);
}
function dv(bytes: Uint8Array): DataView { return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength); }
function directoryOffset(bytes: Uint8Array, id: number): number {
  const view = dv(bytes), size = 2 ** view.getUint16(30, true);
  // Fixtures' directory stream is one sector, except tests that supply an explicit ID.
  return (view.getUint32(48, true) + 1) * size + id * 128;
}
function fatOffset(bytes: Uint8Array, sector: number): number {
  const view = dv(bytes), size = 2 ** view.getUint16(30, true), perFat = size / 4;
  const fatId = view.getUint32(76 + Math.floor(sector / perFat) * 4, true);
  return (fatId + 1) * size + (sector % perFat) * 4;
}

for (const majorVersion of [3, 4] as const) {
  test(`LEG-01: v${majorVersion} nested storages, metadata, mini and normal streams round-trip`, () => {
    const source = fixture();
    const bytes = writeCfb(source, { majorVersion });
    assert.equal(isCfb(bytes), true);
    const result = readCfb(bytes);
    assert.equal(result.majorVersion, majorVersion);
    equivalent(source, result);
    equivalent(source, readCfb(writeCfb(result)));
    assert.deepEqual(writeCfb(source, { majorVersion }), bytes, 'deterministic new serialization');
    assert.equal(findCfbEntry(result, ['unknown', 'small'])?.type, 'stream');
    const size = majorVersion === 3 ? 512 : 4096;
    const entriesUsed = source.entries.length + 1;
    const slots = Math.ceil(entriesUsed * 128 / size) * size / 128;
    for (let id = entriesUsed; id < slots; id++) {
      const offset = directoryOffset(bytes, id);
      const expected = new Uint8Array(128);
      new DataView(expected.buffer).setUint32(68, FREE, true);
      new DataView(expected.buffer).setUint32(72, FREE, true);
      new DataView(expected.buffer).setUint32(76, FREE, true);
      assert.deepEqual(bytes.slice(offset, offset + 128), expected, 'unused directory entries obey MS-CFB 2.6.3');
    }
  });
}

test('LEG-01: empty container and only empty streams need no MiniFAT', () => {
  for (const entries of [[], [{ type: 'stream' as const, path: ['Empty'], data: new Uint8Array() }]]) {
    const bytes = writeCfb({ entries });
    assert.equal(dv(bytes).getUint32(64, true), 0);
    equivalent({ entries }, readCfb(bytes));
  }
});

test('LEG-01: extended DIFAT covers more than 109 FAT sectors', () => {
  const source: CfbDocument = { entries: [{ type: 'stream', path: ['Large'], data: pattern(8 * 1024 * 1024 + 7) }] };
  const bytes = writeCfb(source);
  assert.ok(dv(bytes).getUint32(44, true) > 109);
  assert.ok(dv(bytes).getUint32(72, true) > 0);
  equivalent(source, readCfb(bytes));
  const corrupted = bytes.slice();
  const difat = (dv(corrupted).getUint32(68, true) + 1) * 512;
  dv(corrupted).setUint32(difat + 508, dv(corrupted).getUint32(68, true), true);
  rejected(() => readCfb(corrupted));
});

test('LEG-01: sibling ordering uses UTF-16 length then simple case, with unique names', () => {
  assert.ok(compareCfbNames('Z', 'aa') < 0);
  assert.ok(compareCfbNames('ab', 'AC') < 0);
  assert.equal(compareCfbNames('ᾀ', 'ᾈ'), 0);
  assert.equal(compareCfbNames('ß', 'SS') < 0, true, 'full uppercase expansion does not change name length');
  rejected(() => writeCfb({ entries: [{ type: 'storage', path: ['A'] }, { type: 'storage', path: ['a'] }] }));
  rejected(() => writeCfb({ entries: [{ type: 'storage', path: ['ᾀ'] }, { type: 'storage', path: ['ᾈ'] }] }));
  equivalent({ entries: [{ type: 'storage', path: ['字'.repeat(31)] }] }, readCfb(writeCfb({ entries: [{ type: 'storage', path: ['字'.repeat(31)] }] })));
  for (const name of ['', 'a/b', 'a\\b', 'a:b', 'a!b', 'a\0b', 'x'.repeat(32), '\ud800']) {
    rejected(() => writeCfb({ entries: [{ type: 'storage', path: [name] }] }));
  }
  rejected(() => writeCfb({ entries: [{ type: 'storage', path: ['Missing', 'Child'] }] }));
});

test('SEC-02: rejects truncated files, invalid header geometry and unsupported versions', () => {
  const original = writeCfb(fixture());
  for (const length of [0, 7, 511, original.length - 1, original.length - 512]) rejected(() => readCfb(original.slice(0, length)));
  for (const [offset, value] of [[28, 0xfeff], [30, 12], [32, 9]] as const) {
    const bytes = original.slice(); dv(bytes).setUint16(offset, value, true);
    rejected(() => readCfb(bytes));
  }
  const version = original.slice(); dv(version).setUint16(26, 2, true);
  rejected(() => readCfb(version), 'unsupported');
  const v4 = writeCfb(fixture(), { majorVersion: 4 }); v4[512] = 1;
  rejected(() => readCfb(v4));
});

test('SEC-02: rejects FAT self-cycles, sector overlap and wrong stream length', () => {
  const source: CfbDocument = { entries: [
    { type: 'stream', path: ['One'], data: pattern(5000) },
    { type: 'stream', path: ['Two'], data: pattern(6000) },
  ] };
  const original = writeCfb(source);
  const one = directoryOffset(original, 1), two = directoryOffset(original, 2);
  const start = dv(original).getUint32(one + 116, true);
  const cyclic = original.slice(); dv(cyclic).setUint32(fatOffset(cyclic, start), start, true);
  rejected(() => readCfb(cyclic));
  const overlap = original.slice(); dv(overlap).setUint32(two + 116, start, true);
  rejected(() => readCfb(overlap));
  const short = original.slice(); dv(short).setUint32(fatOffset(short, start), END, true);
  rejected(() => readCfb(short));
  const size = original.slice(); dv(size).setBigUint64(one + 120, 4096n, true);
  rejected(() => readCfb(size));
  const outOfRange = original.slice(); dv(outOfRange).setUint32(fatOffset(outOfRange, start), 0x10000000, true);
  rejected(() => readCfb(outOfRange));
});

test('SEC-02: rejects MiniFAT overlap, cycles and truncated declared allocations', () => {
  const source: CfbDocument = { entries: [
    { type: 'stream', path: ['One'], data: pattern(129) },
    { type: 'stream', path: ['Two'], data: pattern(65) },
  ] };
  const original = writeCfb(source);
  const one = directoryOffset(original, 1), two = directoryOffset(original, 2);
  const start = dv(original).getUint32(one + 116, true);
  const miniFat = (dv(original).getUint32(60, true) + 1) * 512;
  const cyclic = original.slice(); dv(cyclic).setUint32(miniFat + start * 4, start, true);
  rejected(() => readCfb(cyclic));
  const overlap = original.slice(); dv(overlap).setUint32(two + 116, start, true);
  rejected(() => readCfb(overlap));
  const short = original.slice(); dv(short).setUint32(miniFat + start * 4, END, true);
  rejected(() => readCfb(short));
});

test('SEC-02: rejects directory graph cycles, orphans, stream children, bad order and names', () => {
  const original = writeCfb({ entries: [
    { type: 'stream', path: ['AA'], data: pattern(17) },
    { type: 'stream', path: ['BB'], data: pattern(19) },
  ] });
  const root = directoryOffset(original, 0), first = directoryOffset(original, 1), second = directoryOffset(original, 2);
  const cycle = original.slice(); dv(cycle).setUint32(second + 68, 2, true);
  rejected(() => readCfb(cycle));
  const orphan = original.slice(); dv(orphan).setUint32(root + 76, FREE, true);
  rejected(() => readCfb(orphan));
  const child = original.slice(); dv(child).setUint32(first + 76, 2, true);
  rejected(() => readCfb(child));
  const order = original.slice(); dv(order).setUint16(first, 'Z'.charCodeAt(0), true);
  rejected(() => readCfb(order));
  const name = original.slice(); dv(name).setUint16(first + 64, 65, true);
  rejected(() => readCfb(name));
});

test('SEC-03: file/entry/stream/materialization/depth budgets and cancellation are enforced', () => {
  const original = writeCfb(fixture());
  rejected(() => readCfb(original, { limits: { maxFileBytes: 512 } }), 'limit');
  rejected(() => readCfb(original, { limits: { maxStreamBytes: 64 } }), 'limit');
  rejected(() => readCfb(original, { limits: { maxMaterializedBytes: original.length + 1 } }), 'limit');
  rejected(() => readCfb(original, { limits: { maxDepth: 1 } }), 'limit');
  rejected(() => writeCfb(fixture(), { limits: { maxEntries: 3 } }), 'limit');
  rejected(() => writeCfb(fixture(), { limits: { maxStreamBytes: 64 } }), 'limit');
  rejected(() => writeCfb(fixture(), { limits: { maxFileBytes: 4096 } }), 'limit');
  rejected(() => readCfb(original, { signal: { aborted: true } }), 'cancelled');
  rejected(() => writeCfb(fixture(), { signal: { aborted: true } }), 'cancelled');
  let polls = 0;
  rejected(() => readCfb(original, { signal: { get aborted() { return ++polls > 10; } } }), 'cancelled');
  assert.ok(polls > 10, 'cancellation is observed during parsing');
});

test('LEG-01: v4 rejects unsafe 64-bit sizes; v3 follows recommended legacy high-word handling', () => {
  const source: CfbDocument = { entries: [{ type: 'stream', path: ['Tiny'], data: pattern(2) }] };
  const v4 = writeCfb(source, { majorVersion: 4 });
  dv(v4).setBigUint64(directoryOffset(v4, 1) + 120, 1n << 63n, true);
  rejected(() => readCfb(v4), 'limit');
  const v3 = writeCfb(source);
  dv(v3).setUint32(directoryOffset(v3, 1) + 124, 0xabcdef01, true);
  equivalent(source, readCfb(v3));
});


test('LEG-01: reserved root/stream metadata and storage allocations obey MS-CFB 2.6', () => {
  rejected(() => writeCfb({ root: { creationTime: '0000000000000001' }, entries: [] }));
  rejected(() => writeCfb({ entries: [{ type: 'stream', path: ['A'], data: new Uint8Array(), metadata: { clsid: '1'.repeat(32) } }] }));
  rejected(() => writeCfb({ entries: [{ type: 'stream', path: ['A'], data: new Uint8Array(), metadata: { modifiedTime: '0000000000000001' } }] }));
  const output = writeCfb({ entries: [{ type: 'storage', path: ['Storage'], metadata: { creationTime: '0000000000000001', modifiedTime: '0000000000000002' } }] });
  assert.equal(dv(output).getUint32(directoryOffset(output, 1) + 116, true), 0);
  assert.equal(dv(output).getBigUint64(directoryOffset(output, 1) + 120, true), 0n);
});
