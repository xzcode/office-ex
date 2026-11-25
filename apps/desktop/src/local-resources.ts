// SPDX-License-Identifier: MIT
/** Enumerate only build-owned assets at startup; requests never become file paths. */
import { readFile, readdir, lstat } from 'node:fs/promises';
import { join, extname } from 'node:path';
import { OfficeError } from '@office-ex/contracts';
export interface LocalResource { data: Uint8Array; type: string }
export type LocalResources = ReadonlyMap<string, LocalResource>;
const TYPES: Readonly<Record<string, string>> = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.ttf': 'font/ttf', '.ttc': 'font/collection', '.otf': 'font/otf', '.woff': 'font/woff', '.woff2': 'font/woff2', '.wasm': 'application/wasm', '.bcmap': 'application/octet-stream', '.pfb': 'application/octet-stream', '.icc': 'application/octet-stream' };
export async function loadLocalResources(bundle: string, entryFiles: readonly string[]): Promise<Map<string, LocalResource>> {
  const resources = new Map<string, LocalResource>(); let total = 0;
  async function add(relative: string): Promise<void> {
    const path = join(bundle, relative), stat = await lstat(path);
    if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 64 * 1024 * 1024 || (total += stat.size) > 160 * 1024 * 1024) throw new OfficeError('limit', 'Packaged asset budget exceeded');
    const type = TYPES[extname(relative)]; if (!type) throw new OfficeError('invalid', 'Unexpected packaged asset type');
    resources.set('/' + relative, { data: new Uint8Array(await readFile(path)), type });
  }
  async function walk(directory: string, depth = 0): Promise<void> {
    if (depth > 3) throw new OfficeError('limit', 'Packaged asset hierarchy is too deep');
    const stat = await lstat(join(bundle, directory));
    if (!stat.isDirectory() || stat.isSymbolicLink()) throw new OfficeError('unsafe', 'Unexpected packaged asset directory');
    const entries = await readdir(join(bundle, directory), { withFileTypes: true });
    for (const entry of entries) {
      if (!/^[A-Za-z0-9_.-]+$/.test(entry.name) || entry.isSymbolicLink()) throw new OfficeError('unsafe', 'Unexpected packaged asset path');
      const relative = directory + '/' + entry.name;
      if (entry.isDirectory()) await walk(relative, depth + 1);
      else if (entry.isFile() && TYPES[extname(entry.name)]) await add(relative);
    }
  }
  for (const file of entryFiles) await add(file);
  for (const directory of ['pdf-assets', 'fonts']) await walk(directory);
  return resources;
}
export function findLocalResource(resources: LocalResources, value: string): LocalResource | undefined {
  try {
    const url = new URL(value);
    if (url.protocol !== 'office-ex:' || url.hostname !== 'app' || url.port || url.username || url.password || url.search || url.hash) return undefined;
    return resources.get(url.pathname === '/' ? '/index.html' : url.pathname);
  } catch { return undefined; }
}
