/** Private local recovery. Main only checks bounded envelopes/hashes; models are
 * parsed and schema-checked in the sandbox when the user chooses a recovery. */
import { createHash, randomUUID } from 'node:crypto';
import { constants } from 'node:fs';
import { mkdir, open, readdir, rename, unlink, lstat, realpath, chmod, type FileHandle } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { OfficeError } from '@office-ex/contracts';
import { MAX_FILE_BYTES } from './file-service.js';

export const RECOVERY_INTERVAL_MS = 30_000;
export const RECOVERY_MAX_BYTES = 256 * 1024 * 1024;
export const RECOVERY_VERSIONS = 3;
const HEADER_BYTES = 4096;
const MAX_RECOVERY_FILES = 10_000;
const validDocumentId = (id: unknown): id is string => typeof id === 'string' && /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,127}$/.test(id);
const validContentId = (id: unknown): id is string => typeof id === 'string' && /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,255}$/.test(id);
const hash = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex');
const prefix = (documentId: string): string => createHash('sha256').update(documentId).digest('hex').slice(0, 32);
const namePattern = /^[a-f0-9]{32}\.[0-9]+\.[a-f0-9-]{36}\.oex-recovery$/;
export interface RecoveryMetadata { version: 1; id: string; documentId: string; contentId: string; revision: number; createdAt: number; sha256: string; byteLength: number }
export interface RecoveryEntry { metadata: RecoveryMetadata; bytes: Uint8Array }
export type RecoveryWarning = 'quota-pruned' | 'cleanup-failed' | 'directory-sync-unavailable';
export interface RecoveryListing { entries: RecoveryMetadata[]; corrupt: number; warnings: RecoveryWarning[] }
export interface RecoveryOptions {
  /** Lower limits are useful in tests/managed environments; cannot raise budgets. */
  maxBytes?: number;
  maxSnapshotBytes?: number;
  versions?: number;
  fault?: (point: 'after-write' | 'before-commit') => Promise<void>;
}
function failure(error: unknown): never { if (error instanceof OfficeError) throw error; throw new OfficeError('io', 'Local recovery operation failed'); }
function errno(error: unknown): string | undefined { return error && typeof error === 'object' && 'code' in error ? String(error.code) : undefined; }

export class RecoveryStore {
  private storedDirectory: string;
  get directory(): string { return this.storedDirectory; }
  private directoryIdentity?: string;
  private queue: Promise<void> = Promise.resolve();
  private readonly savedThrough = new Map<string, number>();
  private readonly warningCodes = new Set<RecoveryWarning>();
  private readonly maxBytes: number;
  private readonly maxSnapshotBytes: number;
  private readonly versions: number;
  constructor(directory: string, private readonly options: RecoveryOptions = {}) {
    this.storedDirectory = resolve(directory);
    this.maxBytes = options.maxBytes ?? RECOVERY_MAX_BYTES;
    this.maxSnapshotBytes = options.maxSnapshotBytes ?? MAX_FILE_BYTES;
    this.versions = options.versions ?? RECOVERY_VERSIONS;
    for (const [value, maximum] of [[this.maxBytes, RECOVERY_MAX_BYTES], [this.maxSnapshotBytes, MAX_FILE_BYTES], [this.versions, RECOVERY_VERSIONS]]) {
      if (!Number.isSafeInteger(value) || value < 1 || value > maximum) throw new OfficeError('invalid', 'Invalid recovery budget');
    }
  }
  private mutation<T>(work: () => Promise<T>): Promise<T> {
    const result = this.queue.then(work);
    this.queue = result.then(() => undefined, () => undefined);
    return result.catch(failure);
  }
  private async prepare(): Promise<void> {
    // The constructor receives trusted app configuration, never renderer input.
    // Canonicalize its parent once (macOS /var aliases /private/var). The leaf
    // itself must not be a link, and its inode is pinned after initialization.
    if (this.directoryIdentity === undefined) {
      this.storedDirectory = join(await realpath(dirname(this.directory)), basename(this.directory));
    }
    try { await mkdir(this.directory, { mode: 0o700 }); } catch (error) { if (errno(error) !== 'EEXIST') throw error; }
    const stats = await lstat(this.directory, { bigint: true });
    if (stats.isSymbolicLink() || !stats.isDirectory() || await realpath(this.directory) !== this.directory) throw new OfficeError('unsafe', 'Recovery location is not a private local directory');
    const identity = `${stats.dev}:${stats.ino}`;
    if (this.directoryIdentity !== undefined && identity !== this.directoryIdentity) throw new OfficeError('conflict', 'Recovery directory changed');
    if (process.platform !== 'win32' && (Number(stats.mode) & 0o077) !== 0) await chmod(this.directory, 0o700);
    this.directoryIdentity = identity;
  }
  private async assertDirectory(): Promise<void> {
    const stats = await lstat(this.directory, { bigint: true });
    if (stats.isSymbolicLink() || !stats.isDirectory() || await realpath(this.directory) !== this.directory || `${stats.dev}:${stats.ino}` !== this.directoryIdentity) throw new OfficeError('conflict', 'Recovery directory changed');
  }
  private async names(): Promise<string[]> {
    await this.assertDirectory();
    const names = await readdir(this.directory);
    if (names.length > MAX_RECOVERY_FILES) throw new OfficeError('limit', 'Recovery entry limit reached');
    return names;
  }
  private async decode(name: string, pendingPath?: string): Promise<RecoveryEntry> {
    if (!namePattern.test(name)) throw new OfficeError('invalid', 'Invalid recovery identifier');
    await this.assertDirectory();
    const path = pendingPath ?? join(this.directory, name);
    let handle: FileHandle | undefined;
    try {
      const before = await lstat(path, { bigint: true });
      if (!before.isFile() || before.isSymbolicLink() || before.nlink !== 1n || before.size < 4n || before.size > BigInt(this.maxSnapshotBytes + HEADER_BYTES)) throw new OfficeError('invalid', 'Invalid recovery file');
      handle = await open(path, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0));
      const first = await handle.stat({ bigint: true });
      if (!first.isFile() || first.nlink !== 1n || first.dev !== before.dev || first.ino !== before.ino || first.size < 4n || first.size > BigInt(this.maxSnapshotBytes + HEADER_BYTES)) throw new OfficeError('invalid', 'Recovery file changed');
      const data = new Uint8Array(Number(first.size));
      let offset = 0;
      while (offset < data.length) {
        const { bytesRead } = await handle.read(data, offset, Math.min(1024 * 1024, data.length - offset), offset);
        if (bytesRead === 0) throw new OfficeError('invalid', 'Truncated recovery file');
        offset += bytesRead;
      }
      const last = await handle.stat({ bigint: true });
      if (last.size !== first.size || last.mtimeNs !== first.mtimeNs || last.ctimeNs !== first.ctimeNs || last.nlink !== 1n) throw new OfficeError('conflict', 'Recovery file changed');
      await this.assertDirectory();
      const length = new DataView(data.buffer).getUint32(0, true);
      if (length > HEADER_BYTES - 4 || length > data.length - 4) throw new OfficeError('invalid', 'Invalid recovery header');
      let candidate: unknown;
      try { candidate = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(data.subarray(4, 4 + length))); } catch { throw new OfficeError('invalid', 'Invalid recovery metadata'); }
      if (!candidate || typeof candidate !== 'object' || Array.isArray(candidate)) throw new OfficeError('invalid', 'Invalid recovery metadata');
      const m = candidate as Record<string, unknown>;
      if (m.version !== 1 || m.id !== name || !validDocumentId(m.documentId) || !validContentId(m.contentId)
        || !name.startsWith(`${prefix(m.documentId)}.${m.revision}.`) || !Number.isSafeInteger(m.revision) || Number(m.revision) < 0
        || !Number.isSafeInteger(m.createdAt) || Number(m.createdAt) < 0 || typeof m.sha256 !== 'string' || !/^[0-9a-f]{64}$/.test(m.sha256)
        || !Number.isSafeInteger(m.byteLength) || Number(m.byteLength) < 0 || Number(m.byteLength) > this.maxSnapshotBytes) throw new OfficeError('invalid', 'Invalid recovery metadata');
      const bytes = data.slice(4 + length);
      if (bytes.length !== m.byteLength || hash(bytes) !== m.sha256) throw new OfficeError('invalid', 'Recovery integrity check failed');
      return { metadata: candidate as RecoveryMetadata, bytes };
    } finally { await handle?.close().catch(() => undefined); }
  }
  async list(): Promise<RecoveryListing> {
    try {
      await this.prepare();
      const entries: RecoveryMetadata[] = []; let corrupt = 0;
      for (const name of await this.names()) {
        if (!namePattern.test(name)) continue;
        try { entries.push((await this.decode(name)).metadata); } catch { corrupt++; }
      }
      entries.sort((a, b) => b.createdAt - a.createdAt || b.revision - a.revision || (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));
      return { entries, corrupt, warnings: [...this.warningCodes] };
    } catch (error) { return failure(error); }
  }
  async put(documentId: string, contentId: string, revision: number, bytes: Uint8Array): Promise<RecoveryMetadata> {
    if (!validDocumentId(documentId) || !validContentId(contentId) || !Number.isSafeInteger(revision) || revision < 0) throw new OfficeError('invalid', 'Invalid recovery session');
    if (!(bytes instanceof Uint8Array)) throw new OfficeError('invalid', 'Recovery snapshot must be bytes');
    if (bytes.length > this.maxSnapshotBytes) throw new OfficeError('limit', 'Recovery snapshot exceeds budget');
    const immutable = bytes.slice();
    return this.mutation(async () => {
      const saved = this.savedThrough.get(documentId);
      if (saved !== undefined && revision <= saved) throw new OfficeError('conflict', 'Snapshot is already covered by a saved revision');
      await this.prepare();
      const previous = (await this.list()).entries.filter(e => e.documentId === documentId);
      const digest = hash(immutable);
      const existing = previous.find(e => e.contentId === contentId && e.sha256 === digest);
      if (existing) return existing;
      const id = `${prefix(documentId)}.${revision}.${randomUUID()}.oex-recovery`;
      const metadata: RecoveryMetadata = { version: 1, id, documentId, contentId, revision, createdAt: Math.max(Date.now(), (previous[0]?.createdAt ?? 0) + 1), sha256: digest, byteLength: immutable.length };
      const header = new TextEncoder().encode(JSON.stringify(metadata));
      if (header.length > HEADER_BYTES - 4) throw new OfficeError('limit', 'Recovery header exceeds budget');
      const envelopeBytes = 4 + header.length + immutable.length;
      const costs = new Map<string, number>();
      let used = 0;
      for (const name of await this.names()) {
        if (!namePattern.test(name) || !name.startsWith(`${prefix(documentId)}.`)) continue;
        const size = Number((await lstat(join(this.directory, name))).size);
        used += size; costs.set(name, size);
      }
      // When peak space requires pruning, keep the newest known-good snapshot
      // until the replacement is durable. Never remove the last good version.
      const remaining = [...previous];
      while (used + envelopeBytes > this.maxBytes && remaining.length > 1) {
        const oldest = remaining.pop()!;
        await this.assertDirectory();
        await unlink(join(this.directory, oldest.id));
        used -= costs.get(oldest.id) ?? 0;
        this.warningCodes.add('quota-pruned');
      }
      if (used + envelopeBytes > this.maxBytes) throw new OfficeError('limit', 'Recovery budget reached; newest snapshot was preserved');
      let temp: string | undefined = join(this.directory, `.pending-${prefix(documentId)}-${randomUUID()}`);
      let openedIdentity: string | undefined;
      try {
        await this.assertDirectory();
        const file = await open(temp, constants.O_CREAT | constants.O_EXCL | constants.O_WRONLY | (constants.O_NOFOLLOW ?? 0), 0o600);
        try {
          const s = await file.stat({ bigint: true }); openedIdentity = `${s.dev}:${s.ino}`;
          const lead = new Uint8Array(4 + header.length);
          new DataView(lead.buffer).setUint32(0, header.length, true); lead.set(header, 4);
          for (const part of [lead, immutable]) {
            let offset = 0;
            while (offset < part.length) {
              const { bytesWritten } = await file.write(part, offset, Math.min(1024 * 1024, part.length - offset));
              if (!bytesWritten) throw new OfficeError('io', 'Recovery write made no progress');
              offset += bytesWritten;
            }
          }
          await this.options.fault?.('after-write');
          await file.sync();
        } finally { await file.close(); }
        await this.options.fault?.('before-commit');
        await this.assertDirectory();
        const pending = await lstat(temp, { bigint: true });
        if (!pending.isFile() || pending.isSymbolicLink() || pending.nlink !== 1n || `${pending.dev}:${pending.ino}` !== openedIdentity || pending.size !== BigInt(envelopeBytes)) throw new OfficeError('conflict', 'Pending recovery changed');
        const verified = await this.decode(id, temp);
        if (verified.metadata.sha256 !== digest || verified.metadata.contentId !== contentId || verified.metadata.createdAt !== metadata.createdAt) throw new OfficeError('invalid', 'Pending recovery failed verification');
        // Unpredictable identifier in our private directory; no document path is
        // used. Existing valid snapshots remain available until this rename.
        await rename(temp, join(this.directory, id)); temp = undefined;
        let directory: FileHandle | undefined;
        try { directory = await open(this.directory, constants.O_RDONLY); await directory.sync(); }
        catch { this.warningCodes.add('directory-sync-unavailable'); }
        finally { await directory?.close().catch(() => undefined); }
        for (const old of remaining.slice(this.versions - 1)) {
          try { await this.assertDirectory(); await unlink(join(this.directory, old.id)); }
          catch { this.warningCodes.add('cleanup-failed'); }
        }
        return metadata;
      } finally {
        if (temp) {
          try {
            await this.assertDirectory();
            const s = await lstat(temp, { bigint: true });
            if (s.isFile() && !s.isSymbolicLink() && `${s.dev}:${s.ino}` === openedIdentity) await unlink(temp);
          } catch { this.warningCodes.add('cleanup-failed'); }
        }
      }
    });
  }
  /** Call once during startup, after the app acquired its single-instance lock
   * and before accepting checkpoints. Only uncommitted staging files are removed. */
  async cleanupStaging(): Promise<number> {
    return this.mutation(async () => {
      await this.prepare();
      let removed = 0;
      for (const name of await this.names()) {
        if (!/^\.pending-(?:[a-f0-9]{32}-)?[a-f0-9-]{36}$/.test(name)) continue;
        try {
          await this.assertDirectory();
          const path = join(this.directory, name), s = await lstat(path, { bigint: true });
          if (!s.isFile() || s.isSymbolicLink() || s.nlink !== 1n || (process.platform !== 'win32' && s.uid !== BigInt(process.getuid!()))) throw new OfficeError('unsafe', 'Unrecognized recovery staging file');
          await unlink(path); removed++;
        } catch { this.warningCodes.add('cleanup-failed'); }
      }
      return removed;
    });
  }
  async load(id: string): Promise<RecoveryEntry> { try { await this.prepare(); return await this.decode(id); } catch (error) { return failure(error); } }
  async discard(id: string): Promise<void> {
    if (!namePattern.test(id)) throw new OfficeError('invalid', 'Invalid recovery identifier');
    return this.mutation(async () => { await this.prepare(); await this.assertDirectory(); await unlink(join(this.directory, id)); });
  }
  /** Serialize against checkpoints and remove only revisions actually committed. */
  async clearSaved(documentId: string, savedRevision: number): Promise<void> {
    if (!validDocumentId(documentId) || !Number.isSafeInteger(savedRevision) || savedRevision < 0) throw new OfficeError('invalid', 'Invalid saved revision');
    this.savedThrough.set(documentId, Math.max(this.savedThrough.get(documentId) ?? -1, savedRevision));
    return this.mutation(async () => {
      for (const entry of (await this.list()).entries) if (entry.documentId === documentId && entry.revision <= savedRevision) {
        await this.assertDirectory(); await unlink(join(this.directory, entry.id));
      }
    });
  }
}
