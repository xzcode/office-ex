// SPDX-License-Identifier: MIT
// Build-time only. The distributed application uses local, verified font bytes.
import { readFile, writeFile, rename, rm } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { resolve, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
const directory = resolve(dirname(fileURLToPath(import.meta.url)), '../assets/fonts');
const manifest = JSON.parse(await readFile(join(directory, 'manifest.json'), 'utf8'));
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const valid = (bytes, entry) => bytes.length === entry.bytes && sha(bytes) === entry.sha256;
if (manifest.version !== 1 || manifest.files.length !== 8 || manifest.notices.length !== 2) throw new Error('Unexpected font manifest');
for (const entry of [...manifest.files, ...manifest.notices]) {
  if (!/^[A-Za-z0-9_.-]+$/.test(entry.file) || !/^[0-9a-f]{64}$/.test(entry.sha256) || !Number.isSafeInteger(entry.bytes) || entry.bytes <= 0 || entry.bytes > 20_000_000 || !/^https:\/\/raw\.githubusercontent\.com\/notofonts\/noto-(?:fonts|cjk)\/[0-9a-f]{40}\/[A-Za-z0-9/_.-]+$/.test(entry.source)) throw new Error('Invalid pinned font entry');
  const target = join(directory, entry.file);
  const local = await readFile(target).catch(error => { if(error.code === 'ENOENT') return undefined; throw error; });
  if (local) { if (!valid(local, entry)) throw new Error(`Unexpected existing resource: ${entry.file}; remove it and retry`); continue; }
  const response = await fetch(entry.source, { redirect: 'error', signal: AbortSignal.timeout(120_000) });
  if (!response.ok || !response.body) throw new Error(`Font resource unavailable: ${entry.file}: ${response.status}`);
  const chunks = []; let length = 0;
  for await (const chunk of response.body) { length += chunk.byteLength; if(length > entry.bytes) throw new Error(`Resource exceeds pinned size: ${entry.file}`); chunks.push(chunk); }
  const bytes = Buffer.concat(chunks);
  if (!valid(bytes, entry)) throw new Error(`Resource hash mismatch: ${entry.file}`);
  const temporary = `${target}.${process.pid}.tmp`;
  try { await writeFile(temporary, bytes, { flag:'wx' }); await rename(temporary, target); } finally { await rm(temporary, { force:true }); }
  console.log(`Verified ${entry.file}`);
}
console.log('All eight offline font resources and both upstream licenses verified.');
