/** Main-process only. FS-01/02: bounded file capabilities and atomic local commits.
 * Node does not expose a portable rename-if-inode-unchanged or directory-relative
 * handle API. We recheck identity/parent/temp immediately before replacement and
 * keep a durable previous-version backup. An external writer racing the final
 * check and rename cannot be excluded; this is not an OS-level compare-and-swap.
 */
import { createHash, randomUUID } from 'node:crypto';
import { constants } from 'node:fs';
import { open, lstat, realpath, rename, link, unlink, statfs, type FileHandle } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { OfficeError } from '@office-ex/contracts';

export const MAX_FILE_BYTES = 100 * 1024 * 1024;
export interface FileIdentity { dev: string; ino: string; size: number; mtimeNs: string; sha256: string }
export interface FileGrant { token: string; windowId: number; documentId: string; path: string; parent: string; parentIdentity: string; identity: FileIdentity | null; mode: 'source' | 'target' }
export interface OpenedFile { token: string; name: string; bytes: Uint8Array; identity: FileIdentity; documentId: string; duplicate: boolean }
export type SaveFaultPoint = 'after-temp-open' | 'after-write' | 'after-validate' | 'before-commit' | 'before-cleanup';
export interface CommitOptions {
  signal?: AbortSignal;
  validate: (bytes: Uint8Array) => Promise<void>;
  /** Must durably retain exact previous bytes before an existing target is replaced. */
  backup?: (bytes: Uint8Array, identity: FileIdentity) => Promise<void>;
  /** Test injection only; never exposed across IPC. */
  fault?: (point: SaveFaultPoint) => Promise<void>;
}
export interface CommitResult { identity: FileIdentity; name: string; durability: 'synced' | 'file-synced'; cleanupPending: boolean; backupCreated: boolean }
const digest = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex');
const abort = (signal?: AbortSignal): void => { if (signal?.aborted) throw new OfficeError('cancelled', 'Save cancelled'); };
function sameIdentity(a: FileIdentity, b: FileIdentity): boolean { return a.dev === b.dev && a.ino === b.ino && a.size === b.size && a.mtimeNs === b.mtimeNs && a.sha256 === b.sha256; }
function sameFile(a: FileIdentity, b: FileIdentity): boolean { return a.dev === b.dev && a.ino === b.ino; }
function fsError(error: unknown): never {
  if (error instanceof OfficeError) throw error;
  throw new OfficeError('io', 'Local file operation failed');
}
function errno(error: unknown): string | undefined { return error && typeof error === 'object' && 'code' in error ? String(error.code) : undefined; }
function assertSession(windowId: number, documentId: string): void {
  if (!Number.isSafeInteger(windowId) || windowId < 0 || typeof documentId !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9_.:-]{0,127}$/.test(documentId)) throw new OfficeError('invalid', 'Invalid file session');
}
async function canonicalTarget(path: string): Promise<{path:string;parent:string;parentIdentity:string}> {
  if (typeof path !== 'string' || path.length === 0 || path.length > 8192 || path.includes('\0')) throw new OfficeError('invalid', 'Invalid file selection');
  if (process.platform === 'win32' && /^(?:\\\\|\/\/)/.test(path)) throw new OfficeError('unsupported', 'Safe save does not support network shares');
  const absolute = resolve(path);
  const parent = await realpath(dirname(absolute));
  const s = await lstat(parent, { bigint: true });
  if (!s.isDirectory() || s.isSymbolicLink() || !basename(absolute)) throw new OfficeError('unsafe', 'Target parent is not a directory');
  return { path: join(parent, basename(absolute)), parent, parentIdentity: `${s.dev}:${s.ino}` };
}
async function assertParent(grant: Pick<FileGrant, 'parent' | 'parentIdentity'>): Promise<void> {
  try {
    const p = await realpath(grant.parent);
    const s = await lstat(grant.parent, { bigint: true });
    if (p !== grant.parent || s.isSymbolicLink() || !s.isDirectory() || `${s.dev}:${s.ino}` !== grant.parentIdentity) throw new OfficeError('conflict', 'Target directory changed');
  } catch (error) { if (error instanceof OfficeError) throw error; throw new OfficeError('conflict', 'Target directory changed'); }
}
async function assertFilesystem(grant: FileGrant): Promise<void> {
  try {
    const fs = await statfs(grant.parent);
    // Known remote/passthrough Linux types: NFS, CIFS, SMB, 9P, Coda, FUSE,
    // AFS, Ceph. This is conservative detection, not proof of filesystem behavior.
    if ([0x6969, 0xff534d42, 0xfe534d42, 0x517b, 0x01021997, 0x73757245, 0x65735546, 0x5346414f, 0x00c36400].includes(Number(fs.type) >>> 0)) throw new OfficeError('unsupported', 'Safe save requires a local filesystem');
  } catch (error) {
    // libuv statfs may be absent on Windows. UNC shares were rejected at grant
    // creation. Native platform CI verifies ordinary local-drive behavior.
    if (process.platform === 'win32' && ['ENOSYS', 'ENOTSUP', 'EINVAL'].includes(errno(error) ?? '')) return;
    throw error;
  }
}
async function scanCheckedFile(path:string,expectedLinkedInode:FileIdentity|undefined,retainBytes:boolean):Promise<{bytes?:Uint8Array;identity:FileIdentity}>{
  let handle: FileHandle | undefined;
  try {
    const before = await lstat(path, { bigint: true });
    const allowedLinks = expectedLinkedInode && String(before.dev) === expectedLinkedInode.dev && String(before.ino) === expectedLinkedInode.ino ? 2n : 1n;
    if (before.isSymbolicLink() || !before.isFile() || before.nlink > allowedLinks) throw new OfficeError('unsafe', 'Only regular, non-linked files are accepted');
    if (before.size < 0n || before.size > BigInt(MAX_FILE_BYTES)) throw new OfficeError('limit', 'File exceeds the 100 MiB budget');
    handle = await open(path, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0));
    const first = await handle.stat({ bigint: true });
    if (!first.isFile() || first.ino !== before.ino || first.dev !== before.dev || first.nlink > allowedLinks) throw new OfficeError('conflict', 'File changed while opening');
    if (first.size < 0n || first.size > BigInt(MAX_FILE_BYTES)) throw new OfficeError('limit', 'File exceeds the 100 MiB budget');
    const size=Number(first.size),chunkSize=1024*1024;
    const buffer=new Uint8Array(retainBytes?size:Math.min(size,chunkSize));
    const hash=createHash('sha256');
    let offset=0;
    while(offset<size){
      const start=retainBytes?offset:0;
      const {bytesRead}=await handle.read(buffer,start,Math.min(chunkSize,size-offset),offset);
      if(bytesRead===0)throw new OfficeError('conflict','File changed while reading');
      hash.update(buffer.subarray(start,start+bytesRead));offset+=bytesRead;
    }
    const last = await handle.stat({ bigint: true });
    if (last.size !== first.size || last.mtimeNs !== first.mtimeNs || last.ctimeNs !== first.ctimeNs || last.nlink > allowedLinks) throw new OfficeError('conflict', 'File changed while reading');
    return {...(retainBytes?{bytes:buffer}:{}),identity:{dev:String(last.dev),ino:String(last.ino),size,mtimeNs:String(last.mtimeNs),sha256:hash.digest('hex')}};
  } catch (e) { return fsError(e); } finally { await handle?.close().catch(() => undefined); }
}
async function readCheckedFile(path:string,expectedLinkedInode?:FileIdentity):Promise<{bytes:Uint8Array;identity:FileIdentity}>{
 const result=await scanCheckedFile(path,expectedLinkedInode,true);return {bytes:result.bytes!,identity:result.identity};
}
/** Full-content identity verification without retaining a file-sized copy.
 * The same handle, link, size, mtime/ctime and SHA-256 checks are applied. */
async function readCheckedIdentity(path:string,expectedLinkedInode?:FileIdentity):Promise<FileIdentity>{
 return (await scanCheckedFile(path,expectedLinkedInode,false)).identity;
}
export async function readLocalFile(path:string):Promise<{bytes:Uint8Array;identity:FileIdentity}>{return readCheckedFile(path);}

export class FileAuthority {
  private readonly grants = new Map<string, FileGrant>();
  // One local commit queue also covers case/alias spellings of the same inode
  // on macOS/Windows. Revisions stay independent; only physical commits serialize.
  private writes: Promise<void> = Promise.resolve();
  /** The caller must get path from a native dialog, OS file-open event, or preload-validated File drop. */
  async authorizeOpen(path: string, windowId: number, documentId: string): Promise<OpenedFile> {
    try {
      assertSession(windowId, documentId);
      const target = await canonicalTarget(path);
      const read = await readLocalFile(target.path);
      await assertParent(target);
      const existing = [...this.grants.values()].find(g => g.windowId === windowId && g.mode === 'source' && (g.path === target.path || (g.identity && sameFile(g.identity, read.identity))));
      if (existing) return { token: existing.token, name: basename(target.path), ...read, documentId: existing.documentId, duplicate: true };
      const token = randomUUID();
      this.grants.set(token, { ...target, token, windowId, documentId, identity: read.identity, mode: 'source' });
      return { token, name: basename(target.path), ...read, documentId, duplicate: false };
    } catch (e) { return fsError(e); }
  }
  /** Main-only lookup: a second window must focus the existing source
   * capability, not create an independently writable copy of that file. */
  async findOpenSource(path:string):Promise<FileGrant|undefined>{
    const target=await canonicalTarget(path),stat=await lstat(target.path,{bigint:true});
    const existing=[...this.grants.values()].find(g=>g.mode==='source'&&(g.path===target.path||g.identity?.dev===String(stat.dev)&&g.identity.ino===String(stat.ino)));
    if(!existing)return undefined;
    const identity=await readCheckedIdentity(target.path);
    await assertParent(target);
    if(!existing.identity||!sameIdentity(existing.identity,identity))throw new OfficeError('conflict','The already-open file changed outside Office EX');
    return {...existing,identity:{...existing.identity}};
  }
  async authorizeTarget(path: string, windowId: number, documentId: string, replaceConfirmed: boolean, sourceToken?: string): Promise<string> {
    try {
      assertSession(windowId, documentId);
      const target = await canonicalTarget(path);
      const source = sourceToken ? this.get(sourceToken, windowId, documentId) : undefined;
      if (source && source.path === target.path) throw new OfficeError('conflict', 'Save As requires a different target');
      let exists = true;
      try { await lstat(target.path); } catch (error) { if (errno(error) === 'ENOENT') exists = false; else throw error; }
      const identity=exists?await readCheckedIdentity(target.path):null;
      if(identity&&[...this.grants.values()].some(g=>g.mode==='source'&&g.identity&&(g.windowId!==windowId||g.documentId!==documentId)&&sameFile(g.identity,identity)))throw new OfficeError('conflict','Target is open in another document session');
      if (identity && source?.identity && sameFile(source.identity, identity)) throw new OfficeError('conflict', 'Save As requires a different target');
      if (identity && !replaceConfirmed) throw new OfficeError('conflict', 'Existing target needs explicit overwrite confirmation');
      await assertParent(target);
      const token = randomUUID();
      this.grants.set(token, { ...target, token, windowId, documentId, identity, mode: 'target' });
      return token;
    } catch (e) { return fsError(e); }
  }
  get(token: string, windowId: number, documentId: string): FileGrant {
    const grant = this.grants.get(token);
    if (!grant || grant.windowId !== windowId || grant.documentId !== documentId) throw new OfficeError('unsafe', 'File capability is invalid for this session');
    return { ...grant, identity: grant.identity ? { ...grant.identity } : null };
  }
  revokeToken(token: string, windowId: number, documentId: string): void {
    this.get(token, windowId, documentId);
    this.grants.delete(token);
  }
  revokeDocument(windowId: number, documentId: string): void {
    for (const [token, grant] of this.grants) if (grant.windowId === windowId && grant.documentId === documentId) this.grants.delete(token);
  }
  hasToken(token:string):boolean{return this.grants.has(token);}
  revokeWindow(windowId: number): void { for (const [token, grant] of this.grants) if (grant.windowId === windowId) this.grants.delete(token); }
  async commit(token: string, windowId: number, documentId: string, bytes: Uint8Array, options: CommitOptions): Promise<CommitResult> {
    this.get(token, windowId, documentId);
    if (!(bytes instanceof Uint8Array)) throw new OfficeError('invalid', 'Save output must be bytes');
    if (bytes.length > MAX_FILE_BYTES) throw new OfficeError('limit', 'Save output exceeds budget');
    if (!options || typeof options.validate !== 'function') throw new OfficeError('invalid', 'Save validation is required');
    const immutable = bytes.slice();
    const result = this.writes.then(() => this.commitNow(token, windowId, documentId, immutable, options));
    this.writes = result.then(() => undefined, () => undefined);
    return result;
  }
  private async commitNow(token: string, windowId: number, documentId: string, immutable: Uint8Array, options: CommitOptions): Promise<CommitResult> {
    const grant = this.get(token, windowId, documentId);
    let temporary: string | undefined;
    let handle: FileHandle | undefined;
    let cleanupPending = false;
    let backupCreated = false;
    let temporaryIdentity: string | undefined;
    let result: CommitResult | undefined;
    const cleanup = async (): Promise<void> => {
      if (!temporary) return;
      await assertParent(grant);
      const s = await lstat(temporary, { bigint: true });
      if (!s.isFile() || s.isSymbolicLink() || `${s.dev}:${s.ino}` !== temporaryIdentity) throw new OfficeError('conflict', 'Temporary save identity changed');
      await unlink(temporary);
      temporary = undefined;
    };
    try {
      abort(options.signal);
      await assertParent(grant);
      await assertFilesystem(grant);
      if (grant.identity && typeof options.backup !== 'function') throw new OfficeError('unsafe', 'Overwriting requires a durable previous-version backup');
      temporary = join(grant.parent, `.office-ex-${randomUUID()}.tmp`);
      handle = await open(temporary, constants.O_CREAT | constants.O_EXCL | constants.O_WRONLY | (constants.O_NOFOLLOW ?? 0), 0o600);
      const openedTemp = await handle.stat({ bigint: true });
      temporaryIdentity = `${openedTemp.dev}:${openedTemp.ino}`;
      await options.fault?.('after-temp-open');
      let offset = 0;
      while (offset < immutable.length) {
        abort(options.signal);
        const { bytesWritten } = await handle.write(immutable, offset, Math.min(1024 * 1024, immutable.length - offset));
        if (!bytesWritten) throw new OfficeError('io', 'Save made no progress');
        offset += bytesWritten;
      }
      await options.fault?.('after-write');
      await handle.sync();
      await handle.close(); handle = undefined;
      const reread = await readLocalFile(temporary);
      if (reread.identity.dev !== String(openedTemp.dev) || reread.identity.ino !== String(openedTemp.ino) || reread.identity.sha256 !== digest(immutable)) throw new OfficeError('io', 'Saved bytes failed verification');
      await options.validate(reread.bytes);
      await options.fault?.('after-validate');
      abort(options.signal);
      if (grant.identity) {
        const previous = await readLocalFile(grant.path);
        if (!sameIdentity(grant.identity, previous.identity)) throw new OfficeError('conflict', 'File was modified outside Office EX');
        await options.backup!(previous.bytes, previous.identity);
        backupCreated = true;
      }
      await options.fault?.('before-commit');
      abort(options.signal);
      this.get(token, windowId, documentId); // A closed/revoked tab cannot commit later.
      await assertParent(grant);
      const lastTemp=await readCheckedIdentity(temporary);
      if (!sameIdentity(reread.identity,lastTemp)) throw new OfficeError('conflict', 'Temporary save changed before commit');
      if (grant.identity) {
        const current=await readCheckedIdentity(grant.path);
        if (!sameIdentity(grant.identity,current)) throw new OfficeError('conflict', 'File was modified outside Office EX');
        await rename(temporary, grant.path); // Same-directory atomic replace; never delete first.
        temporary = undefined;
      } else {
        try { await link(temporary, grant.path); }
        catch (e) { if (errno(e) === 'EEXIST') throw new OfficeError('conflict', 'Target appeared while saving'); throw e; }
        // The commit has succeeded. A cleanup error must not report the save as
        // failed or remove the committed destination. Finally retries cleanup.
        try { await options.fault?.('before-cleanup'); await cleanup(); }
        catch { cleanupPending = true; }
      }
      let durability: 'synced' | 'file-synced' = 'synced';
      let directory: FileHandle | undefined;
      try { directory = await open(grant.parent, constants.O_RDONLY); await directory.sync(); }
      catch { durability = 'file-synced'; }
      finally { await directory?.close().catch(() => undefined); }
      const saved=await readCheckedIdentity(grant.path,cleanupPending?reread.identity:undefined);
      if (!sameFile(saved,reread.identity)||saved.sha256!==reread.identity.sha256) throw new OfficeError('conflict', 'Saved file changed after commit; previous-version backup was retained');
      if (this.grants.has(token)) this.grants.set(token, { ...grant, identity: saved, mode: 'source' });
      result = { identity: saved, name: basename(grant.path), durability, cleanupPending, backupCreated };
      return result;
    } catch (e) { return fsError(e); }
    finally {
      await handle?.close().catch(() => undefined);
      if (temporary) {
        try { await cleanup(); if (result) result.cleanupPending = false; }
        catch { if (result) result.cleanupPending = true; }
      }
    }
  }
}
