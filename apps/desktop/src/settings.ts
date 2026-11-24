// SPDX-License-Identifier: MIT
import { mkdir, open, rename, unlink, lstat } from 'node:fs/promises';
import { constants } from 'node:fs';
import { join, isAbsolute } from 'node:path';
import { randomUUID } from 'node:crypto';
import { OfficeError, isRecord } from '@office-ex/contracts';
import { isLanguagePreference, isRegionalPreference, isDocumentLanguagePreference, type LanguagePreference, type RegionalPreference, type DocumentLanguagePreference } from '@office-ex/i18n';
import { parseSettingsPatch } from './ipc-policy.js';
export interface Settings { version: 2; theme: 'system' | 'light' | 'dark'; language: LanguagePreference; region: RegionalPreference; documentLanguage: DocumentLanguagePreference; recent: { id: string; path: string; name: string }[] }
const MAX_SETTINGS_BYTES = 128 * 1024;
const defaults = (): Settings => ({ version: 2, theme: 'system', language: 'system', region: 'system', documentLanguage: 'system', recent: [] });
function validPath(path: unknown): path is string { return typeof path === 'string' && path.length > 0 && path.length <= 8192 && !path.includes('\0') && isAbsolute(path); }
function validName(name: unknown): name is string { return typeof name === 'string' && name.length > 0 && name.length <= 1024 && !name.includes('\0'); }
function validate(value: unknown): Settings | null {
  if (!isRecord(value)) return null;
  const legacy=value.version===1;
  if(Object.keys(value).some(key=>!(legacy?['version','theme','language','recent']:['version','theme','language','region','documentLanguage','recent']).includes(key)))return null;
  if(legacy&&(value.language!=='zh-CN'&&value.language!=='en'))return null;
  const language=legacy&&value.language==='en'?'en-US':value.language;
  const region=legacy?'system':value.region, documentLanguage=legacy?'system':value.documentLanguage;
  if((!legacy&&value.version!==2)||(value.theme!=='system'&&value.theme!=='light'&&value.theme!=='dark')||
    !isLanguagePreference(language)||!isRegionalPreference(region)||!isDocumentLanguagePreference(documentLanguage)||
    !Array.isArray(value.recent)||value.recent.length>20)return null;
  const recent: Settings['recent'] = [], ids = new Set<string>(), paths = new Set<string>();
  for (const item of value.recent) {
    if (!isRecord(item) || Object.keys(item).some(key => !['id','path','name'].includes(key)) || typeof item.id !== 'string' || !/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/.test(item.id) ||
        !validPath(item.path) || !validName(item.name) || ids.has(item.id) || paths.has(item.path)) return null;
    ids.add(item.id); paths.add(item.path); recent.push({id:item.id,path:item.path,name:item.name});
  }
  return {version:2,theme:value.theme,language,region,documentLanguage,recent};
}
export class SettingsStore {
  private settings: Settings = defaults();
  private queue: Promise<void> = Promise.resolve();
  constructor(private readonly directory: string) {}
  async load(): Promise<Settings> {
    let file: Awaited<ReturnType<typeof open>> | undefined;
    try {
      const path = join(this.directory, 'settings.json'), before = await lstat(path, {bigint:true});
      if (!before.isFile() || before.isSymbolicLink() || before.size > BigInt(MAX_SETTINGS_BYTES)) return this.value();
      file = await open(path, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0));
      const first = await file.stat({bigint:true});
      if (!first.isFile() || first.dev !== before.dev || first.ino !== before.ino || first.size > BigInt(MAX_SETTINGS_BYTES)) return this.value();
      const bytes = new Uint8Array(Number(first.size)); let offset = 0;
      while (offset < bytes.length) {
        const result = await file.read(bytes,offset,bytes.length-offset,offset);
        if (!result.bytesRead) return this.value();
        offset += result.bytesRead;
      }
      const last = await file.stat({bigint:true});
      if (last.size !== first.size || last.mtimeNs !== first.mtimeNs || last.ctimeNs !== first.ctimeNs) return this.value();
      const candidate: unknown = JSON.parse(new TextDecoder('utf-8',{fatal:true}).decode(bytes));
      const checked = validate(candidate); if (checked) this.settings = checked;
    } catch { /* Missing/corrupt settings never make documents unavailable. */ }
    finally { await file?.close().catch(() => undefined); }
    return this.value();
  }
  value(): Settings { return structuredClone(this.settings); }
  publicValue(): Omit<Settings,'recent'> & {recent: {id:string;name:string}[]} { return { ...this.settings, recent: this.settings.recent.map(({id,name}) => ({id,name})) }; }
  async update(patch: {theme?:unknown;language?:unknown;region?:unknown;documentLanguage?:unknown;clearRecent?:unknown}): Promise<void> {
    // Validate the whole transaction before queuing; never coerce objects to enum strings.
    const checked = parseSettingsPatch(patch);
    await this.change(current => ({...current,
      ...(checked.theme === undefined ? {} : {theme:checked.theme}),
      ...(checked.language === undefined ? {} : {language:checked.language}),
      ...(checked.region === undefined ? {} : {region:checked.region}),
      ...(checked.documentLanguage === undefined ? {} : {documentLanguage:checked.documentLanguage}),
      ...(checked.clearRecent === true ? {recent:[]} : {}),
    }));
  }
  async addRecent(path: string, name: string): Promise<void> {
    if (!validPath(path) || !validName(name)) throw new OfficeError('invalid','Invalid recent file metadata');
    await this.change(current => ({...current,recent:[{id:randomUUID(),path,name},...current.recent.filter(item => item.path !== path)].slice(0,20)}));
  }
  resolveRecent(id: string): string {
    const item = this.settings.recent.find(value => value.id === id);
    if (!item) throw new OfficeError('invalid', 'Unknown recent file');
    return item.path;
  }
  private async change(transform: (current: Settings) => Settings): Promise<void> {
    const operation = this.queue.catch(() => undefined).then(async () => {
      const next = transform(this.value());
      await this.persist(next);
      this.settings = next; // Publish only the state which actually reached the settings file.
    });
    this.queue = operation;
    await operation;
  }
  private async persist(next: Settings): Promise<void> {
    const bytes = new TextEncoder().encode(JSON.stringify(next));
    if (bytes.length > MAX_SETTINGS_BYTES) throw new OfficeError('limit','Settings exceed the local metadata budget');
    let temp: string | undefined;
    let file: Awaited<ReturnType<typeof open>> | undefined;
    try {
      await mkdir(this.directory,{recursive:true,mode:0o700});
      temp = join(this.directory, `.settings-${randomUUID()}.tmp`);
      file = await open(temp,'wx',0o600); await file.writeFile(bytes); await file.sync(); await file.close(); file = undefined;
      await rename(temp,join(this.directory,'settings.json')); temp = undefined;
      // Directory sync is best effort for preferences; the document save path reports its durability separately.
      const directory = await open(this.directory,constants.O_RDONLY).catch(() => undefined);
      if (directory) { try { await directory.sync(); } catch {/* unsupported directory sync */} finally { await directory.close().catch(() => undefined); } }
    } catch (error) {
      if (error instanceof OfficeError) throw error;
      throw new OfficeError('io','Settings could not be saved');
    } finally { await file?.close().catch(() => undefined); if (temp) await unlink(temp).catch(() => undefined); }
  }
}
