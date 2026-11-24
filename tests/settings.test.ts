// SPDX-License-Identifier: MIT
import test, { type TestContext } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, readdir, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { OfficeError } from '@office-ex/contracts';
import { SettingsStore } from '../apps/desktop/src/settings.js';

const defaults = { version: 1, theme: 'system', language: 'zh-CN', recent: [] };
const fails = (code: string) => (error: unknown): boolean => error instanceof OfficeError && error.code === code;
async function directory(t: TestContext): Promise<string> {
  const path = await mkdtemp(join(tmpdir(), 'office-ex-settings-'));
  t.after(() => rm(path, { recursive: true, force: true }));
  return path;
}

test('Settings defaults, theme and language survive an atomic write and reload', async t => {
  const path = await directory(t), store = new SettingsStore(path);
  assert.deepEqual(await store.load(), defaults);
  await store.update({ theme: 'dark', language: 'en' });
  const loaded = await new SettingsStore(path).load();
  assert.deepEqual(loaded, { ...defaults, theme: 'dark', language: 'en' });
  assert.deepEqual((await readdir(path)).sort(), ['settings.json']);
});

test('A rejected settings transaction changes neither memory nor persisted preferences', async t => {
  const path = await directory(t), store = new SettingsStore(path);
  await store.update({ theme: 'light' });
  const before = await readFile(join(path, 'settings.json'));
  for (const invalid of [
    { theme: 'dark', language: 'invalid' }, { theme: ['dark'] },
    { language: ['en'] }, { clearRecent: 'true' }, { theme: null },
  ]) await assert.rejects(store.update(invalid), fails('invalid'));
  assert.deepEqual(store.value(), { ...defaults, theme: 'light' });
  assert.deepEqual(await readFile(join(path, 'settings.json')), before);
});

test('Concurrent settings changes merge against the last committed transaction', async t => {
  const path = await directory(t), store = new SettingsStore(path);
  const first = resolve(path, 'one.docx'), second = resolve(path, 'two.xls');
  await Promise.all([
    store.update({ theme: 'dark' }), store.addRecent(first, 'one.docx'),
    store.update({ language: 'en' }), store.addRecent(second, 'two.xls'),
  ]);
  assert.equal(store.value().theme, 'dark');
  assert.equal(store.value().language, 'en');
  assert.deepEqual(store.value().recent.map(item => item.path), [second, first]);
  assert.deepEqual(await new SettingsStore(path).load(), store.value());
});

test('A failed preference commit preserves memory, cleans staging and does not poison the queue', async t => {
  const path = await directory(t), store = new SettingsStore(path);
  await store.update({ theme: 'dark' });
  await rm(join(path, 'settings.json'));
  await mkdir(join(path, 'settings.json')); // A directory cannot be replaced by the settings file.
  await assert.rejects(store.update({ theme: 'light', language: 'en' }), fails('io'));
  assert.deepEqual(store.value(), { ...defaults, theme: 'dark' });
  assert.deepEqual(await readdir(path), ['settings.json']);
  await rm(join(path, 'settings.json'), { recursive: true });
  await store.update({ language: 'en' });
  assert.deepEqual(await new SettingsStore(path).load(), { ...defaults, theme: 'dark', language: 'en' });
});

test('Malformed, oversized and coercible persisted settings are ignored as a whole', async t => {
  const path = await directory(t);
  const validRecent = { id: '00000000-0000-0000-0000-000000000000', path: resolve(path, 'one.docx'), name: 'one.docx' };
  for (const invalid of [
    { ...defaults, theme: ['dark'] }, { ...defaults, language: ['en'] },
    { ...defaults, theme: 'dark', extra: true },
    { ...defaults, theme: 'dark', recent: [{ ...validRecent, path: 'relative.docx' }] },
    { ...defaults, theme: 'dark', recent: [{ ...validRecent, name: '' }] },
    { ...defaults, theme: 'dark', recent: [validRecent, validRecent] },
  ]) {
    await writeFile(join(path, 'settings.json'), JSON.stringify(invalid));
    assert.deepEqual(await new SettingsStore(path).load(), defaults);
  }
  for (const bytes of [Buffer.from('{malformed'), Buffer.from([0xff, 0xfe]), Buffer.alloc(128 * 1024 + 1, 0x20)]) {
    await writeFile(join(path, 'settings.json'), bytes);
    assert.deepEqual(await new SettingsStore(path).load(), defaults);
  }
});

test('Recent-file capabilities withhold paths and return defensive state copies', async t => {
  const path = await directory(t), store = new SettingsStore(path), source = resolve(path, 'private.docx');
  await store.addRecent(source, 'private.docx');
  const copy = store.value(), item = copy.recent[0]!;
  assert.equal(store.resolveRecent(item.id), source);
  assert.deepEqual(store.publicValue().recent, [{ id: item.id, name: 'private.docx' }]);
  copy.recent[0]!.path = '/changed'; copy.theme = 'dark';
  const publicCopy = store.publicValue(); publicCopy.recent[0]!.name = 'changed';
  assert.equal(store.resolveRecent(item.id), source);
  assert.equal(store.value().theme, 'system');
  assert.equal(store.publicValue().recent[0]!.name, 'private.docx');
  await assert.rejects(store.addRecent('relative.docx', 'relative.docx'), fails('invalid'));
  await assert.rejects(store.addRecent(source, 'bad\0name'), fails('invalid'));
  await store.update({ clearRecent: true });
  assert.throws(() => store.resolveRecent(item.id), fails('invalid'));
  assert.deepEqual((await new SettingsStore(path).load()).recent, []);
});

test('A settings symlink is never followed during load', { skip: process.platform === 'win32' }, async t => {
  const path = await directory(t), target = join(path, 'outside.json');
  await writeFile(target, JSON.stringify({ ...defaults, theme: 'dark' }));
  await symlink(target, join(path, 'settings.json'));
  assert.deepEqual(await new SettingsStore(path).load(), defaults);
});
