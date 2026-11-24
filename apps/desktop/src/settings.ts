import { mkdir, readFile, rename, writeFile, lstat } from 'node:fs/promises';
import { join } from 'node:path';
import { randomUUID } from 'node:crypto';
import { OfficeError } from '@office-ex/contracts';
export interface Settings { version: 1; theme: 'system' | 'light' | 'dark'; language: 'zh-CN' | 'en'; recent: { id: string; path: string; name: string }[] }
const defaults = (): Settings => ({ version: 1, theme: 'system', language: 'zh-CN', recent: [] });
export class SettingsStore {
  private settings: Settings = defaults();
  private queue: Promise<void> = Promise.resolve();
  constructor(private readonly directory: string) {}
  async load(): Promise<Settings> {
    try {
      const path = join(this.directory, 'settings.json');
      const s = await lstat(path);
      if (!s.isFile() || s.isSymbolicLink() || s.size > 128 * 1024) return this.value();
      const x: unknown = JSON.parse(await readFile(path, 'utf8'));
      if (x && typeof x === 'object' && !Array.isArray(x)) {
        const a = x as Record<string, unknown>;
        if (a.version === 1 && ['system', 'light', 'dark'].includes(String(a.theme)) && ['zh-CN', 'en'].includes(String(a.language)) && Array.isArray(a.recent) && a.recent.length <= 20) {
          const recent = a.recent.filter((r): r is Settings['recent'][number] => r && typeof r === 'object' && typeof r.id === 'string' && /^[a-f0-9-]{36}$/.test(r.id) && typeof r.path === 'string' && r.path.length <= 8192 && typeof r.name === 'string' && r.name.length <= 1024);
          this.settings = { version: 1, theme: a.theme as Settings['theme'], language: a.language as Settings['language'], recent };
        }
      }
    } catch { /* Missing/corrupt settings never make documents unavailable. */ }
    return this.value();
  }
  value(): Settings { return structuredClone(this.settings); }
  publicValue(): Omit<Settings,'recent'> & {recent: {id:string;name:string}[]} { return { ...this.settings, recent: this.settings.recent.map(({id,name}) => ({id,name})) }; }
  async update(patch: {theme?:unknown;language?:unknown;clearRecent?:unknown}): Promise<void> {
    if (patch.theme !== undefined) {
      if (!['system','light','dark'].includes(String(patch.theme))) throw new OfficeError('invalid','Invalid theme');
      this.settings.theme = patch.theme as Settings['theme'];
    }
    if (patch.language !== undefined) {
      if (!['zh-CN','en'].includes(String(patch.language))) throw new OfficeError('invalid','Invalid language');
      this.settings.language = patch.language as Settings['language'];
    }
    if (patch.clearRecent === true) this.settings.recent = [];
    await this.persist();
  }
  async addRecent(path: string, name: string): Promise<void> {
    this.settings.recent = [{ id: randomUUID(), path, name }, ...this.settings.recent.filter(r => r.path !== path)].slice(0,20);
    await this.persist();
  }
  resolveRecent(id: string): string {
    const item = this.settings.recent.find(r => r.id === id);
    if (!item) throw new OfficeError('invalid', 'Unknown recent file');
    return item.path;
  }
  private async persist(): Promise<void> {
    const data = JSON.stringify(this.settings);
    this.queue = this.queue.catch(() => undefined).then(async () => {
      await mkdir(this.directory, {recursive: true, mode: 0o700});
      const temp = join(this.directory, `.settings-${randomUUID()}.tmp`);
      await writeFile(temp, data, {flag:'wx',mode:0o600});
      await rename(temp, join(this.directory,'settings.json'));
    });
    await this.queue;
  }
}
