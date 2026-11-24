import test, { type TestContext } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, realpath, readFile, writeFile, readdir, rm, mkdir, rename, symlink, unlink, lstat, open } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { OfficeError } from '@office-ex/contracts';
import { RecoveryStore } from '../apps/desktop/src/recovery.js';

const bytes = (value: string) => new TextEncoder().encode(value);
const text = (value: Uint8Array) => new TextDecoder().decode(value);
async function temporary(t: TestContext): Promise<string> {
  const root = await realpath(await mkdtemp(join(tmpdir(), 'office-ex-recovery-')));
  t.after(() => rm(root, { recursive: true, force: true })); return root;
}
async function rejectsCode(work: Promise<unknown>, code: string): Promise<void> { await assert.rejects(work, (error: unknown) => error instanceof OfficeError && error.code === code); }

test('FS-03: core content IDs, immutable bytes and safe filenames survive a new store instance', async t => {
  const parent = await temporary(t), directory = join(parent, 'recovery'), store = new RecoveryStore(directory);
  const data = bytes('original'); const pending = store.put('doc:example.1', 'doc:example.1:content:0', 0, data); data.fill(0);
  const saved = await pending;
  assert.equal(saved.id.includes(':'), false); assert.equal(saved.id.includes('doc:example.1'), false);
  const reopened = new RecoveryStore(directory); const loaded = await reopened.load(saved.id);
  assert.equal(text(loaded.bytes), 'original'); assert.equal(loaded.metadata.contentId, 'doc:example.1:content:0');
  const duplicate = await reopened.put('doc:example.1', 'doc:example.1:content:0', 7, bytes('original'));
  assert.equal(duplicate.id, saved.id); assert.equal((await reopened.list()).entries.length, 1);
  if (process.platform !== 'win32') { assert.equal((await lstat(directory)).mode & 0o777, 0o700); assert.equal((await lstat(join(directory, saved.id))).mode & 0o777, 0o600); }
});

test('FS-03: retains the newest three versions per document; backup hash IDs are accepted', async t => {
  const store = new RecoveryStore(join(await temporary(t), 'recovery'));
  for (let revision = 1; revision <= 5; revision++) await store.put('doc', `doc:content:${revision}`, revision, bytes(`revision ${revision}`));
  await store.put('other', 'backup:' + 'a'.repeat(64), 0, bytes('previous file'));
  const listing = await store.list();
  assert.deepEqual(listing.entries.filter(e => e.documentId === 'doc').map(e => e.revision), [5, 4, 3]);
  assert.equal(listing.entries.filter(e => e.documentId === 'other').length, 1);
  assert.equal(listing.corrupt, 0);
});

test('FS-03: disk quota pruning keeps the last good version if a new write fails', async t => {
  const directory = join(await temporary(t), 'recovery'); let fail = false;
  const store = new RecoveryStore(directory, { maxBytes: 15_000, maxSnapshotBytes: 5000, fault: async point => { if (fail && point === 'after-write') throw Object.assign(new Error('disk full'), { code: 'ENOSPC' }); } });
  for (let i = 1; i <= 3; i++) await store.put('doc', `doc:content:${i}`, i, new Uint8Array(4096).fill(i));
  fail = true; await rejectsCode(store.put('doc', 'doc:content:4', 4, new Uint8Array(4096).fill(4)), 'io');
  const afterFailure = await store.list();
  assert.equal(afterFailure.entries[0]!.revision, 3); assert.ok(afterFailure.warnings.includes('quota-pruned'));
  assert.deepEqual((await store.load(afterFailure.entries[0]!.id)).bytes, new Uint8Array(4096).fill(3));
  fail = false; await store.put('doc', 'doc:content:4', 4, new Uint8Array(4096).fill(4));
  assert.deepEqual((await store.list()).entries.map(e => e.revision), [4, 3, 2]);
  const names = await readdir(directory); let used = 0; for (const name of names) used += (await lstat(join(directory, name))).size;
  assert.ok(used <= 15_000); assert.equal(names.some(name => name.startsWith('.pending-')), false);
});

test('FS-03: insufficient peak budget preserves the only good snapshot and enforces payload limits', async t => {
  const store = new RecoveryStore(join(await temporary(t), 'recovery'), { maxBytes: 8000, maxSnapshotBytes: 5000 });
  const previous = await store.put('doc', 'doc:content:1', 1, new Uint8Array(4096).fill(1));
  await rejectsCode(store.put('doc', 'doc:content:2', 2, new Uint8Array(4096).fill(2)), 'limit');
  await rejectsCode(store.put('doc', 'doc:content:2', 2, new Uint8Array(5001)), 'limit');
  assert.equal((await store.list()).entries[0]!.id, previous.id); assert.equal((await store.load(previous.id)).bytes[0], 1);
});

test('FS-03: corrupted payload, malformed envelope and unknown schema remain available for diagnosis', async t => {
  const directory = join(await temporary(t), 'recovery'), store = new RecoveryStore(directory);
  const a = await store.put('a', 'a:content:1', 1, bytes('payload'));
  const path = join(directory, a.id), corrupt = await readFile(path); corrupt[corrupt.length - 1] ^= 255; await writeFile(path, corrupt);
  await rejectsCode(store.load(a.id), 'invalid');
  const b = await store.put('b', 'b:content:1', 1, bytes('future'));
  const futurePath = join(directory, b.id), future = await readFile(futurePath);
  const length = future.readUInt32LE(0), metadata = JSON.parse(future.subarray(4, 4 + length).toString()); metadata.version = 2;
  const header = Buffer.from(JSON.stringify(metadata)); assert.equal(header.length, length); header.copy(future, 4); await writeFile(futurePath, future);
  await rejectsCode(store.load(b.id), 'invalid');
  const c = await store.put('c', 'c:content:1', 1, bytes('short'));
  await writeFile(join(directory, c.id), Uint8Array.of(1, 2)); await rejectsCode(store.load(c.id), 'invalid');
  const listing = await store.list(); assert.equal(listing.corrupt, 3); assert.equal(listing.entries.length, 0);
  assert.equal((await readdir(directory)).length, 3, 'diagnosis never silently deletes corrupt recovery files');
});

test('FS-03: tampered staging bytes are detected before retention can delete old snapshots', async t => {
  const directory = join(await temporary(t), 'recovery'); let tamper = false;
  const store = new RecoveryStore(directory, { fault: async point => {
    if (!tamper || point !== 'before-commit') return;
    const pending = (await readdir(directory)).find(name => name.startsWith('.pending-'))!;
    const file = await open(join(directory, pending), 'r+');
    try { const size = (await file.stat()).size; await file.write(Uint8Array.of(99), 0, 1, size - 1); } finally { await file.close(); }
  } });
  const previous = await store.put('doc', 'doc:content:1', 1, bytes('original'));
  tamper = true; await rejectsCode(store.put('doc', 'doc:content:2', 2, bytes('newbytes')), 'invalid');
  assert.equal(text((await store.load(previous.id)).bytes), 'original');
  assert.deepEqual((await store.list()).entries.map(e => e.id), [previous.id]);
});

test('FS-03: saved revision cleanup is serialized with an in-flight checkpoint and preserves later edits', async t => {
  const directory = join(await temporary(t), 'recovery'); let entered!: () => void, unblock!: () => void;
  const ready = new Promise<void>(resolve => { entered = resolve; }), gate = new Promise<void>(resolve => { unblock = resolve; }); let block = true;
  const store = new RecoveryStore(directory, { fault: async point => { if (block && point === 'after-write') { entered(); await gate; } } });
  const checkpoint = store.put('doc', 'doc:content:1', 1, bytes('one')); await ready;
  const cleanup = store.clearSaved('doc', 1); block = false; unblock(); await checkpoint; await cleanup;
  assert.equal((await store.list()).entries.length, 0);
  await rejectsCode(store.put('doc', 'doc:content:1', 1, bytes('late stale checkpoint')), 'conflict');
  const later = await store.put('doc', 'doc:content:2', 2, bytes('two'));
  await store.put('other', 'other:content:1', 1, bytes('other'));
  await store.clearSaved('doc', 1);
  assert.ok((await store.list()).entries.some(e => e.id === later.id)); assert.equal((await store.list()).entries.length, 2);
});

test('FS-03: path traversal and symlinked recovery payloads are rejected without touching external files', async t => {
  const parent = await temporary(t), directory = join(parent, 'recovery'), store = new RecoveryStore(directory);
  const saved = await store.put('doc', 'doc:content:1', 1, bytes('original'));
  await rejectsCode(store.load('../outside'), 'invalid'); await rejectsCode(store.discard('../outside'), 'invalid');
  const outside = join(parent, 'outside'); await writeFile(outside, 'private outside'); await unlink(join(directory, saved.id));
  try { await symlink(outside, join(directory, saved.id), 'file'); }
  catch (error) { if (process.platform === 'win32' && (error as NodeJS.ErrnoException).code === 'EPERM') { t.skip('Host does not grant symlink creation'); return; } throw error; }
  await rejectsCode(store.load(saved.id), 'invalid'); assert.equal(await readFile(outside, 'utf8'), 'private outside'); assert.equal((await store.list()).corrupt, 1);
});

test('FS-03: replacing the private directory invalidates the store identity', async t => {
  const parent = await temporary(t), directory = join(parent, 'recovery'), store = new RecoveryStore(directory);
  const previous = await store.put('doc', 'doc:content:1', 1, bytes('original'));
  await rename(directory, join(parent, 'moved')); await mkdir(directory);
  await rejectsCode(store.put('doc', 'doc:content:2', 2, bytes('new')), 'conflict');
  const movedStore = new RecoveryStore(join(parent, 'moved')); assert.equal(text((await movedStore.load(previous.id)).bytes), 'original');
});

test('FS-03: startup staging cleanup preserves committed and unrelated files', async t => {
  const directory = join(await temporary(t), 'recovery'), store = new RecoveryStore(directory);
  const previous = await store.put('doc', 'doc:content:1', 1, bytes('original'));
  const staging = '.pending-' + 'a'.repeat(32) + '-12345678-1234-1234-1234-123456789abc';
  await writeFile(join(directory, staging), 'incomplete', { mode: 0o600 });
  await writeFile(join(directory, 'unrelated'), 'leave alone');
  assert.equal(await store.cleanupStaging(), 1);
  assert.equal(text((await store.load(previous.id)).bytes), 'original');
  assert.equal(await readFile(join(directory, 'unrelated'), 'utf8'), 'leave alone');
});

test('FS-03: trusted parent aliases canonicalize once while linked recovery directories are rejected', async t => {
  const parent = await temporary(t), actual = join(parent, 'actual'), alias = join(parent, 'alias'); await mkdir(actual);
  try { await symlink(actual, alias, process.platform === 'win32' ? 'junction' : 'dir'); }
  catch (error) { if (process.platform === 'win32' && (error as NodeJS.ErrnoException).code === 'EPERM') { t.skip('Host does not grant symlink creation'); return; } throw error; }
  const store = new RecoveryStore(join(alias, 'recovery'));
  const previous = await store.put('doc', 'doc:content:1', 1, bytes('original'));
  assert.equal(store.directory, join(actual, 'recovery')); assert.equal(text((await store.load(previous.id)).bytes), 'original');
  const leafAlias = join(parent, 'leaf'); await symlink(join(actual, 'recovery'), leafAlias, process.platform === 'win32' ? 'junction' : 'dir');
  await rejectsCode(new RecoveryStore(leafAlias).list(), 'unsafe');
});
