import test, { type TestContext } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, realpath, writeFile, readFile, readdir, rm, mkdir, rename, symlink, link, lstat, truncate } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { OfficeError } from '@office-ex/contracts';
import { FileAuthority, readLocalFile, MAX_FILE_BYTES, type CommitOptions, type SaveFaultPoint } from '../apps/desktop/src/file-service.js';

const bytes = (value: string) => new TextEncoder().encode(value);
const asText = (value: Uint8Array) => new TextDecoder().decode(value);
const validate = async (value: Uint8Array): Promise<void> => { assert.ok(value instanceof Uint8Array); };
const backup: CommitOptions['backup'] = async () => undefined;
async function temporary(t: TestContext): Promise<string> {
  const root = await realpath(await mkdtemp(join(tmpdir(), 'office-ex-files-')));
  t.after(() => rm(root, { recursive: true, force: true })); return root;
}
async function rejectsCode(work: Promise<unknown>, code: string): Promise<void> { await assert.rejects(work, (error: unknown) => error instanceof OfficeError && error.code === code); }
async function noTemps(directory: string): Promise<void> { assert.equal((await readdir(directory)).some(name => name.startsWith('.office-ex-')), false); }
async function makeSymlink(t: TestContext, target: string, path: string, type: 'file' | 'dir' = 'file'): Promise<boolean> {
  try { await symlink(target, path, process.platform === 'win32' && type === 'dir' ? 'junction' : type); return true; }
  catch (error) { if (process.platform === 'win32' && (error as NodeJS.ErrnoException).code === 'EPERM') { t.skip('Host does not grant symlink creation'); return false; } throw error; }
}

test('FS-01: capabilities are window/document bound, duplicate-aware and revocable', async t => {
  const directory = await temporary(t), path = join(directory, 'a.oex'); await writeFile(path, 'original');
  const authority = new FileAuthority();
  const opened = await authority.authorizeOpen(path, 1, 'doc:one');
  assert.equal(opened.duplicate, false);
  const duplicate = await authority.authorizeOpen(path, 1, 'doc:two');
  assert.equal(duplicate.duplicate, true); assert.equal(duplicate.documentId, 'doc:one');
  assert.throws(() => authority.get(opened.token, 2, 'doc:one'), { code: 'unsafe' });
  assert.throws(() => authority.get(opened.token, 1, 'doc:two'), { code: 'unsafe' });
  const grant = authority.get(opened.token, 1, 'doc:one'); grant.path = '/changed'; grant.identity!.sha256 = 'changed';
  assert.equal(authority.get(opened.token, 1, 'doc:one').path, path);
  authority.revokeDocument(1, 'doc:one');
  assert.throws(() => authority.get(opened.token, 1, 'doc:one'), { code: 'unsafe' });
});

test('FS-02: Save As requires a distinct path and an explicit overwrite grant', async t => {
  const directory = await temporary(t), source = join(directory, 'a.oex'), target = join(directory, 'b.oex');
  await writeFile(source, 'source'); await writeFile(target, 'target');
  const authority = new FileAuthority(), opened = await authority.authorizeOpen(source, 1, 'doc');
  await rejectsCode(authority.authorizeTarget(source, 1, 'doc', true, opened.token), 'conflict');
  await rejectsCode(authority.authorizeTarget(target, 1, 'doc', false), 'conflict');
  const token = await authority.authorizeTarget(target, 1, 'doc', true, opened.token);
  const backed: string[] = [];
  const result = await authority.commit(token, 1, 'doc', bytes('new'), { validate, backup: async old => { backed.push(asText(old)); } });
  assert.equal(result.backupCreated, true); assert.deepEqual(backed, ['target']);
  assert.equal(await readFile(source, 'utf8'), 'source'); assert.equal(await readFile(target, 'utf8'), 'new');
});

test('FS-02: a new atomic save, zero-byte save, and subsequent overwrite preserve identities', async t => {
  const directory = await temporary(t), path = join(directory, 'new.oex');
  const authority = new FileAuthority(), token = await authority.authorizeTarget(path, 1, 'doc', false);
  const first = await authority.commit(token, 1, 'doc', bytes('one'), { validate });
  assert.equal(first.backupCreated, false); assert.equal(first.cleanupPending, false);
  const previous: string[] = [];
  const second = await authority.commit(token, 1, 'doc', new Uint8Array(), { validate, backup: async old => { previous.push(asText(old)); } });
  assert.deepEqual(previous, ['one']); assert.equal(second.identity.size, 0); assert.equal((await readFile(path)).length, 0);
  assert.equal(authority.get(token, 1, 'doc').identity!.sha256, second.identity.sha256);
  if (process.platform !== 'win32') assert.equal((await lstat(path)).mode & 0o777, 0o600);
  await noTemps(directory);
});

test('FS-02: overwrite without a backup or with a failed backup leaves original intact', async t => {
  const directory = await temporary(t), path = join(directory, 'a.oex'); await writeFile(path, 'original');
  const authority = new FileAuthority(), source = await authority.authorizeOpen(path, 1, 'doc');
  await rejectsCode(authority.commit(source.token, 1, 'doc', bytes('changed'), { validate }), 'unsafe');
  await rejectsCode(authority.commit(source.token, 1, 'doc', bytes('changed'), { validate, backup: async () => { throw Object.assign(new Error('disk full'), { code: 'ENOSPC' }); } }), 'io');
  assert.equal((await readLocalFile(path)).identity.sha256, source.identity.sha256); await noTemps(directory);
});

for (const point of ['after-temp-open', 'after-write', 'after-validate', 'before-commit'] as SaveFaultPoint[]) {
  test(`FS-02: injected ${point} failure does not replace the original`, async t => {
    const directory = await temporary(t), path = join(directory, 'a.oex'); await writeFile(path, 'original');
    const authority = new FileAuthority(), source = await authority.authorizeOpen(path, 1, 'doc');
    await rejectsCode(authority.commit(source.token, 1, 'doc', bytes('changed'), { validate, backup, fault: async actual => { if (actual === point) throw Object.assign(new Error('injected IO failure'), { code: 'ENOSPC' }); } }), 'io');
    assert.equal((await readLocalFile(path)).identity.sha256, source.identity.sha256); await noTemps(directory);
  });
}

test('FS-02: validator rejection and cancellation after writing preserve original', async t => {
  const directory = await temporary(t), path = join(directory, 'a.oex'); await writeFile(path, 'original');
  const authority = new FileAuthority(), source = await authority.authorizeOpen(path, 1, 'doc');
  await rejectsCode(authority.commit(source.token, 1, 'doc', bytes('changed'), { backup, validate: async () => { throw new OfficeError('invalid', 'semantic verification failed'); } }), 'invalid');
  const controller = new AbortController();
  await rejectsCode(authority.commit(source.token, 1, 'doc', bytes('changed'), { validate, backup, signal: controller.signal, fault: async point => { if (point === 'after-write') controller.abort(); } }), 'cancelled');
  assert.equal((await readLocalFile(path)).identity.sha256, source.identity.sha256); await noTemps(directory);
});

test('FS-02: external changes before save or during backup cause a conflict', async t => {
  const directory = await temporary(t), path = join(directory, 'a.oex'); await writeFile(path, 'original');
  const authority = new FileAuthority(), source = await authority.authorizeOpen(path, 1, 'doc');
  await writeFile(path, 'foreign');
  await rejectsCode(authority.commit(source.token, 1, 'doc', bytes('changed'), { validate, backup }), 'conflict');
  assert.equal(await readFile(path, 'utf8'), 'foreign');
  const fresh = new FileAuthority(), opened = await fresh.authorizeOpen(path, 1, 'doc');
  await rejectsCode(fresh.commit(opened.token, 1, 'doc', bytes('changed'), { validate, backup: async () => { await writeFile(path, 'racing writer'); } }), 'conflict');
  assert.equal(await readFile(path, 'utf8'), 'racing writer'); await noTemps(directory);
});

test('FS-02: queued saves use refreshed identity and an immutable input snapshot', async t => {
  const directory = await temporary(t), path = join(directory, 'a.oex'); await writeFile(path, 'original');
  const authority = new FileAuthority(), source = await authority.authorizeOpen(path, 1, 'doc');
  const backed: string[] = []; let unblock!: () => void, entered!: () => void;
  const gate = new Promise<void>(resolve => { unblock = resolve; }); const ready = new Promise<void>(resolve => { entered = resolve; });
  const first = authority.commit(source.token, 1, 'doc', bytes('first'), { validate: async () => { entered(); await gate; }, backup: async old => { backed.push(asText(old)); } });
  await ready;
  const mutable = bytes('second');
  const second = authority.commit(source.token, 1, 'doc', mutable, { validate, backup: async old => { backed.push(asText(old)); } });
  mutable.fill(0); unblock(); await Promise.all([first, second]);
  assert.equal(await readFile(path, 'utf8'), 'second'); assert.deepEqual(backed, ['original', 'first']);
});

test('FS-02: independent grants to one path cannot both overwrite stale contents', async t => {
  const directory = await temporary(t), path = join(directory, 'a.oex'); await writeFile(path, 'original');
  const authority = new FileAuthority();
  const a = await authority.authorizeTarget(path, 1, 'a', true), b = await authority.authorizeTarget(path, 1, 'b', true);
  const results = await Promise.allSettled([
    authority.commit(a, 1, 'a', bytes('first'), { validate, backup }),
    authority.commit(b, 1, 'b', bytes('second'), { validate, backup }),
  ]);
  assert.equal(results[0]!.status, 'fulfilled'); assert.equal(results[1]!.status, 'rejected');
  if (results[1]!.status === 'rejected') assert.equal(results[1]!.reason.code, 'conflict');
  assert.equal(await readFile(path, 'utf8'), 'first');
});

test('FS-02: a racing new target is never overwritten; revoked in-flight grant cannot commit', async t => {
  const directory = await temporary(t), path = join(directory, 'a.oex'), authority = new FileAuthority();
  const token = await authority.authorizeTarget(path, 1, 'doc', false);
  await rejectsCode(authority.commit(token, 1, 'doc', bytes('ours'), { validate, fault: async point => { if (point === 'before-commit') await writeFile(path, 'foreign', { flag: 'wx' }); } }), 'conflict');
  assert.equal(await readFile(path, 'utf8'), 'foreign');
  const other = await authority.authorizeTarget(join(directory, 'b.oex'), 1, 'doc2', false);
  await rejectsCode(authority.commit(other, 1, 'doc2', bytes('ours'), { validate, fault: async point => { if (point === 'before-commit') authority.revokeDocument(1, 'doc2'); } }), 'unsafe');
  await assert.rejects(lstat(join(directory, 'b.oex')), { code: 'ENOENT' }); await noTemps(directory);
});

test('FS-02: a post-link cleanup failure still reports a successful verified new save', async t => {
  const directory = await temporary(t), path = join(directory, 'a.oex'), authority = new FileAuthority();
  const token = await authority.authorizeTarget(path, 1, 'doc', false);
  const result = await authority.commit(token, 1, 'doc', bytes('ours'), { validate, fault: async point => { if (point === 'before-cleanup') throw new Error('cleanup failed'); } });
  assert.equal(await readFile(path, 'utf8'), 'ours'); assert.equal(result.identity.sha256, (await readLocalFile(path)).identity.sha256);
  assert.equal(result.cleanupPending, false, 'finally successfully retried cleanup'); await noTemps(directory);
});

test('FS-02: changed parent directory is rejected and cleanup cannot unlink a foreign file', async t => {
  const directory = await temporary(t), parent = join(directory, 'parent'), moved = join(directory, 'moved'); await mkdir(parent);
  const path = join(parent, 'a.oex'); await writeFile(path, 'original');
  const authority = new FileAuthority(), source = await authority.authorizeOpen(path, 1, 'doc'); let foreignTemp = '';
  await rejectsCode(authority.commit(source.token, 1, 'doc', bytes('ours'), { validate, backup, fault: async point => {
    if (point !== 'before-commit') return;
    const tempName = (await readdir(parent)).find(name => name.startsWith('.office-ex-'))!;
    await rename(parent, moved); await mkdir(parent); foreignTemp = join(parent, tempName); await writeFile(foreignTemp, 'foreign');
  } }), 'conflict');
  assert.equal(await readFile(join(moved, 'a.oex'), 'utf8'), 'original'); assert.equal(await readFile(foreignTemp, 'utf8'), 'foreign');
});

test('FS-01/02: symlinks, hardlinks, directories and oversized sparse sources are rejected', async t => {
  const directory = await temporary(t), path = join(directory, 'source'); await writeFile(path, 'source');
  const soft = join(directory, 'soft'); if (!await makeSymlink(t, path, soft)) return;
  await rejectsCode(readLocalFile(soft), 'unsafe');
  const authority = new FileAuthority(); await rejectsCode(authority.authorizeTarget(soft, 1, 'doc', true), 'unsafe');
  const hard = join(directory, 'hard'); await link(path, hard); await rejectsCode(readLocalFile(hard), 'unsafe');
  await rejectsCode(readLocalFile(directory), 'unsafe');
  const huge = join(directory, 'huge'); await writeFile(huge, ''); await truncate(huge, MAX_FILE_BYTES + 1); await rejectsCode(readLocalFile(huge), 'limit');
  assert.equal(await readFile(path, 'utf8'), 'source');
});
