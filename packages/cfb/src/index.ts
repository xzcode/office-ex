/**
 * Original bounded CFB container implementation. No Office format semantics.
 * Standards baseline: [MS-CFB] revision 13.0 (2026-10-26), sections 2.1–2.6.
 * https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b
 * Reference PDF SHA-256: dc5c9924fb9b1eaf074e319511fb99121c22124e868461f1de95849fdd3d6919.
 * Directory ordering follows 2.6.4 (UTF-16 length, then simple uppercase units).
 * No COM activation, external IO, or extraction to filesystem paths occurs here.
 */
import { OfficeError } from '@office-ex/contracts';

const FREE = 0xffffffff;
const END = 0xfffffffe;
const FAT_SECTOR = 0xfffffffd;
const DIFAT_SECTOR = 0xfffffffc;
const MINI_SIZE = 64;
const MINI_CUTOFF = 4096;
const MAGIC = Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1);
const ZERO_CLSID = '00000000000000000000000000000000';
const ZERO_TIME = '0000000000000000';

export interface CfbLimits {
  maxFileBytes: number;
  maxEntries: number;
  maxDepth: number;
  maxStreamBytes: number;
  maxMaterializedBytes: number;
}
export const CFB_DEFAULT_LIMITS: Readonly<CfbLimits> = Object.freeze({
  maxFileBytes: 100 * 1024 * 1024,
  maxEntries: 10_000,
  maxDepth: 128,
  maxStreamBytes: 64 * 1024 * 1024,
  maxMaterializedBytes: 512 * 1024 * 1024,
});
export interface CfbOptions {
  /** Callers may lower product limits; raising a product budget requires an ADR. */
  limits?: Partial<CfbLimits>;
  signal?: { readonly aborted: boolean };
}
export interface CfbWriteOptions extends CfbOptions {
  majorVersion?: 3 | 4;
}
export interface CfbMetadata {
  /** Raw on-disk 16 CLSID bytes as 32 hexadecimal digits, not a formatted GUID. */
  clsid?: string;
  stateBits?: number;
  /** Unsigned FILETIME value as 16 hexadecimal digits; preserves all 64 bits. */
  creationTime?: string;
  modifiedTime?: string;
}
interface CfbEntryBase {
  /** Components are CFB names, never host filesystem paths. */
  path: readonly string[];
  metadata?: CfbMetadata;
}
export type CfbEntry =
  | (CfbEntryBase & { type: 'storage' })
  | (CfbEntryBase & { type: 'stream'; data: Uint8Array });
export interface CfbDocument {
  majorVersion?: 3 | 4;
  minorVersion?: number;
  transactionSignature?: number;
  root?: CfbMetadata;
  entries: readonly CfbEntry[];
}
export interface ParsedCfbDocument extends CfbDocument {
  majorVersion: 3 | 4;
  minorVersion: number;
  transactionSignature: number;
  root: Required<CfbMetadata>;
}

function invalid(message: string): never { throw new OfficeError('invalid', `CFB: ${message}`); }
function limited(message: string): never { throw new OfficeError('limit', `CFB: ${message}`); }
function checkCancel(options: CfbOptions): void {
  if (options.signal?.aborted) throw new OfficeError('cancelled', 'CFB operation cancelled');
}
function limitsFor(options: CfbOptions): CfbLimits {
  const limits = { ...CFB_DEFAULT_LIMITS };
  for (const key of Object.keys(CFB_DEFAULT_LIMITS) as (keyof CfbLimits)[]) {
    const value = options.limits?.[key];
    if (value !== undefined) {
      if (!Number.isSafeInteger(value) || value < 1 || value > CFB_DEFAULT_LIMITS[key]) {
        invalid(`invalid ${key} budget`);
      }
      limits[key] = value;
    }
  }
  return limits;
}
function metadata(value: CfbMetadata | undefined, type: 'root' | 'storage' | 'stream'): Required<CfbMetadata> {
  const result = {
    clsid: value?.clsid ?? ZERO_CLSID,
    stateBits: value?.stateBits ?? 0,
    creationTime: value?.creationTime ?? ZERO_TIME,
    modifiedTime: value?.modifiedTime ?? ZERO_TIME,
  };
  if (typeof result.clsid !== 'string' || !/^[0-9a-f]{32}$/i.test(result.clsid)
    || typeof result.creationTime !== 'string' || !/^[0-9a-f]{16}$/i.test(result.creationTime)
    || typeof result.modifiedTime !== 'string' || !/^[0-9a-f]{16}$/i.test(result.modifiedTime)
    || !Number.isInteger(result.stateBits) || result.stateBits < 0 || result.stateBits > FREE) {
    invalid('invalid directory metadata');
  }
  result.clsid = result.clsid.toLowerCase();
  result.creationTime = result.creationTime.toLowerCase();
  result.modifiedTime = result.modifiedTime.toLowerCase();
  if ((type === 'stream' || type === 'root') && result.creationTime !== ZERO_TIME) invalid('nonzero creation time on stream/root');
  if (type === 'stream' && (result.modifiedTime !== ZERO_TIME || result.clsid !== ZERO_CLSID)) invalid('nonzero reserved stream metadata');
  return result;
}
function nameValid(name: string): void {
  if (typeof name !== 'string' || name.length === 0 || name.length > 31 || /[\u0000/\\:!]/u.test(name)) invalid('invalid directory name');
  for (let i = 0; i < name.length; i++) {
    const c = name.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff) {
      const next = name.charCodeAt(++i);
      if (!(next >= 0xdc00 && next <= 0xdfff)) invalid('invalid UTF-16 directory name');
    } else if (c >= 0xdc00 && c <= 0xdfff) invalid('invalid UTF-16 directory name');
  }
}
/** Unicode simple uppercase is one code unit. Greek ypogegrammeni has a simple
 * one-unit mapping even though ECMAScript's full uppercase expands it.
 * Verified against Unicode 17.0 UnicodeData simple-uppercase field:
 * https://www.unicode.org/Public/17.0.0/ucd/UnicodeData.txt */
function upperUnit(unit: number): number {
  if ((unit >= 0x1f80 && unit <= 0x1f87) || (unit >= 0x1f90 && unit <= 0x1f97) || (unit >= 0x1fa0 && unit <= 0x1fa7)) return unit + 8;
  if (unit === 0x1fb3) return 0x1fbc;
  if (unit === 0x1fc3) return 0x1fcc;
  if (unit === 0x1ff3) return 0x1ffc;
  const uppercase = String.fromCharCode(unit).toUpperCase();
  return uppercase.length === 1 ? uppercase.charCodeAt(0) : unit;
}
export function compareCfbNames(a: string, b: string): number {
  if (a.length !== b.length) return a.length - b.length;
  for (let i = 0; i < a.length; i++) {
    const diff = upperUnit(a.charCodeAt(i)) - upperUnit(b.charCodeAt(i));
    if (diff !== 0) return diff;
  }
  return 0;
}
function pathKey(path: readonly string[]): string {
  return JSON.stringify(path.map(name => {
    let upper = '';
    for (let i = 0; i < name.length; i++) upper += String.fromCharCode(upperUnit(name.charCodeAt(i)));
    return upper;
  }));
}
export function findCfbEntry(document: CfbDocument, path: readonly string[]): CfbEntry | undefined {
  const key = pathKey(path);
  return document.entries.find(entry => pathKey(entry.path) === key);
}
export function isCfb(bytes: Uint8Array): boolean {
  return bytes.length >= 8 && MAGIC.every((value, index) => bytes[index] === value);
}

interface RawEntry {
  id: number;
  type: 0 | 1 | 2 | 5;
  name: string;
  left: number;
  right: number;
  child: number;
  color: number;
  start: number;
  size: number;
  meta: Required<CfbMetadata>;
}

/** Parses both 512-byte (v3) and 4096-byte (v4) containers. */
export function readCfb(bytes: Uint8Array, options: CfbOptions = {}): ParsedCfbDocument {
  checkCancel(options);
  const limits = limitsFor(options);
  if (!(bytes instanceof Uint8Array)) invalid('input must be Uint8Array');
  if (bytes.length > limits.maxFileBytes) limited('file bytes exceeded');
  if (bytes.length < 512 || !isCfb(bytes)) invalid('missing compound file signature/header');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (offset: number) => view.getUint16(offset, true);
  const u32 = (offset: number) => view.getUint32(offset, true);
  const major = u16(26);
  if (major !== 3 && major !== 4) throw new OfficeError('unsupported', 'CFB major version is not 3 or 4');
  const sectorSize = major === 3 ? 512 : 4096;
  const perFat = sectorSize / 4;
  const perDirectory = sectorSize / 128;
  if (u16(28) !== 0xfffe || u16(30) !== (major === 3 ? 9 : 12) || u16(32) !== 6 || u32(56) !== MINI_CUTOFF) invalid('invalid header geometry');
  if (bytes.length < sectorSize * 3 || bytes.length % sectorSize !== 0) invalid('truncated or unaligned sector file');
  for (let i = 8; i < 24; i++) if (bytes[i] !== 0) invalid('nonzero header CLSID');
  for (let i = 34; i < 40; i++) if (bytes[i] !== 0) invalid('nonzero reserved header');
  if (major === 4) for (let i = 512; i < sectorSize; i++) if (bytes[i] !== 0) invalid('nonzero v4 header padding');
  const sectorCount = bytes.length / sectorSize - 1;
  const dirCount = u32(40), fatCount = u32(44), miniFatCount = u32(64), difatCount = u32(72);
  if (major === 3 && dirCount !== 0) invalid('v3 directory sector count must be zero');
  if (major === 4 && (dirCount < 1 || dirCount > Math.ceil(limits.maxEntries / perDirectory))) limited('directory sector count');
  if (fatCount < 1 || fatCount > sectorCount || difatCount > sectorCount || miniFatCount > sectorCount) invalid('allocation counts exceed physical sectors');
  let materialized = bytes.length;
  const account = (size: number): void => {
    if (!Number.isSafeInteger(size) || size < 0 || size > limits.maxMaterializedBytes - materialized) limited('materialized bytes exceeded');
    materialized += size;
  };
  account(sectorCount);
  const owners = new Uint8Array(sectorCount);
  const sectorOffset = (sid: number): number => {
    if (!Number.isInteger(sid) || sid < 0 || sid >= sectorCount) invalid('sector index out of range');
    return (sid + 1) * sectorSize;
  };
  const claim = (sid: number, owner: number): void => {
    sectorOffset(sid);
    if (owners[sid] !== 0) invalid('sector overlap or cyclic chain');
    owners[sid] = owner;
  };
  const fatIds: number[] = [];
  let reachedFree = false;
  const addFatId = (sid: number): void => {
    if (sid === FREE) { reachedFree = true; return; }
    if (reachedFree || fatIds.length >= fatCount) invalid('invalid DIFAT padding/count');
    sectorOffset(sid);
    fatIds.push(sid);
  };
  for (let i = 0; i < 109; i++) addFatId(u32(76 + i * 4));
  let difat = u32(68);
  const difatIds: number[] = [];
  for (let i = 0; i < difatCount; i++) {
    checkCancel(options);
    claim(difat, 2);
    difatIds.push(difat);
    const offset = sectorOffset(difat);
    for (let j = 0; j < perFat - 1; j++) addFatId(u32(offset + j * 4));
    difat = u32(offset + sectorSize - 4);
  }
  if (difat !== END || fatIds.length !== fatCount) invalid('DIFAT terminator/count mismatch');
  account(fatCount * sectorSize);
  const fat = new Uint32Array(fatCount * perFat);
  if (fat.length < sectorCount) invalid('FAT does not cover file');
  for (let i = 0; i < fatIds.length; i++) {
    checkCancel(options);
    const sid = fatIds[i]!;
    claim(sid, 1);
    const offset = sectorOffset(sid);
    for (let j = 0; j < perFat; j++) fat[i * perFat + j] = u32(offset + j * 4);
  }
  for (let i = 0; i < fat.length; i++) {
    const next = fat[i]!;
    if (i >= sectorCount) {
      if (next !== FREE) invalid('allocated FAT entry beyond file');
      continue;
    }
    if (next < DIFAT_SECTOR && next >= sectorCount) invalid('FAT successor out of range');
    if (owners[i] === 1 && next !== FAT_SECTOR) invalid('FAT sector is not marked FATSECT');
    if (owners[i] === 2 && next !== DIFAT_SECTOR) invalid('DIFAT sector is not marked DIFSECT');
    if (next === FAT_SECTOR && owners[i] !== 1) invalid('unlisted FAT sector');
    if (next === DIFAT_SECTOR && owners[i] !== 2) invalid('unlisted DIFAT sector');
  }
  const chain = (start: number, owner: number, expected: number | undefined, max: number): number[] => {
    const ids: number[] = [];
    if (expected === 0) {
      if (start !== END && start !== FREE) invalid('empty chain has allocated start');
      return ids;
    }
    let current = start;
    while (current !== END) {
      checkCancel(options);
      if (ids.length >= max) limited('chain resource budget exceeded');
      if (expected !== undefined && ids.length >= expected) invalid('chain longer than declared');
      claim(current, owner);
      ids.push(current);
      current = fat[current]!;
    }
    if (expected !== undefined && ids.length !== expected) invalid('chain shorter than declared');
    return ids;
  };
  const readChain = (ids: readonly number[], size: number): Uint8Array => {
    if (size > ids.length * sectorSize) invalid('stream exceeds chain');
    account(size);
    const output = new Uint8Array(size);
    for (let i = 0; i < ids.length && i * sectorSize < size; i++) {
      checkCancel(options);
      const offset = sectorOffset(ids[i]!);
      output.set(bytes.subarray(offset, offset + Math.min(sectorSize, size - i * sectorSize)), i * sectorSize);
    }
    return output;
  };
  const directoryIds = chain(u32(48), 3, major === 4 ? dirCount : undefined, Math.ceil(limits.maxEntries / perDirectory));
  if (directoryIds.length === 0) invalid('missing directory');
  const directory = readChain(directoryIds, directoryIds.length * sectorSize);
  const dv = new DataView(directory.buffer, directory.byteOffset, directory.byteLength);
  const raw: RawEntry[] = [];
  let activeEntries = 0;
  for (let offset = 0; offset < directory.length; offset += 128) {
    checkCancel(options);
    const type = directory[offset + 66]!;
    if (type !== 0 && type !== 1 && type !== 2 && type !== 5) invalid('invalid directory object type');
    if (type === 0) {
      raw.push({ id: raw.length, type, name: '', left: FREE, right: FREE, child: FREE, color: 1, start: END, size: 0, meta: metadata(undefined, 'storage') });
      continue;
    }
    if (++activeEntries > limits.maxEntries) limited('directory entries exceeded');
    const nameBytes = dv.getUint16(offset + 64, true);
    if (nameBytes < 4 || nameBytes > 64 || nameBytes % 2 !== 0 || dv.getUint16(offset + nameBytes - 2, true) !== 0) invalid('invalid directory name length/termination');
    let name = '';
    for (let i = 0; i < nameBytes - 2; i += 2) name += String.fromCharCode(dv.getUint16(offset + i, true));
    nameValid(name);
    const color = directory[offset + 67]!;
    if (color > 1) invalid('invalid directory tree color');
    let clsid = '';
    for (let i = 80; i < 96; i++) clsid += directory[offset + i]!.toString(16).padStart(2, '0');
    const meta = metadata({
      clsid,
      stateBits: dv.getUint32(offset + 96, true),
      creationTime: dv.getBigUint64(offset + 100, true).toString(16).padStart(16, '0'),
      modifiedTime: dv.getBigUint64(offset + 108, true).toString(16).padStart(16, '0'),
    }, type === 5 ? 'root' : type === 1 ? 'storage' : 'stream');
    // 2.6.1 explicitly recommends ignoring uninitialized high size words in v3.
    const sizeBig = major === 3 ? BigInt(dv.getUint32(offset + 120, true)) : dv.getBigUint64(offset + 120, true);
    if (sizeBig > BigInt(Number.MAX_SAFE_INTEGER)) limited('unsafe 64-bit stream size');
    const size = Number(sizeBig);
    const start = dv.getUint32(offset + 116, true);
    if (type === 1 && (start !== 0 || size !== 0)) invalid('storage has stream allocation');
    if (type === 2 && size > limits.maxStreamBytes) limited('stream bytes exceeded');
    if (type === 5 && size > limits.maxFileBytes) limited('mini stream bytes exceeded');
    const entry: RawEntry = {
      id: raw.length, type, name, color, meta, start, size,
      left: dv.getUint32(offset + 68, true),
      right: dv.getUint32(offset + 72, true),
      child: dv.getUint32(offset + 76, true),
    };
    if (type === 2 && entry.child !== FREE) invalid('stream has directory children');
    raw.push(entry);
  }
  const root = raw[0]!;
  if (root.type !== 5 || root.name !== 'Root Entry' || root.left !== FREE || root.right !== FREE) invalid('invalid root entry');
  if (root.size % MINI_SIZE !== 0) invalid('unaligned root mini stream');
  const miniStreamIds = chain(root.start, 5, Math.ceil(root.size / sectorSize), sectorCount);
  const miniStream = readChain(miniStreamIds, root.size);
  const miniCount = root.size / MINI_SIZE;
  // Some producers leave additional all-free MiniFAT capacity. Preserve logical
  // streams; every unused MiniFAT entry is checked below, never followed.
  const miniFatIds = chain(u32(60), 4, miniFatCount, sectorCount);
  account(miniFatCount * sectorSize + miniCount);
  const miniFat = new Uint32Array(miniFatCount * perFat);
  const miniOwners = new Uint8Array(miniCount);
  if (miniFat.length < miniCount) invalid('MiniFAT does not cover mini stream');
  for (let i = 0; i < miniFatIds.length; i++) {
    const offset = sectorOffset(miniFatIds[i]!);
    for (let j = 0; j < perFat; j++) miniFat[i * perFat + j] = u32(offset + j * 4);
  }
  for (let i = 0; i < miniFat.length; i++) {
    const next = miniFat[i]!;
    if (i >= miniCount) {
      if (next !== FREE) invalid('allocated mini sector beyond mini stream');
    } else if (next !== FREE && next !== END && next >= miniCount) invalid('mini FAT successor out of range');
  }
  const readMini = (entry: RawEntry): Uint8Array => {
    account(entry.size);
    const output = new Uint8Array(entry.size);
    let current = entry.start;
    const expected = Math.ceil(entry.size / MINI_SIZE);
    if (expected === 0) {
      if (current !== END && current !== FREE) invalid('empty stream has allocated start');
      return output;
    }
    for (let i = 0; i < expected; i++) {
      checkCancel(options);
      if (current >= miniCount || miniOwners[current] !== 0) invalid('mini sector overlap, cycle or out of range');
      miniOwners[current] = 1;
      const length = Math.min(MINI_SIZE, entry.size - i * MINI_SIZE);
      output.set(miniStream.subarray(current * MINI_SIZE, current * MINI_SIZE + length), i * MINI_SIZE);
      current = miniFat[current]!;
    }
    if (current !== END) invalid('mini chain longer than declared');
    return output;
  };
  const entries: CfbEntry[] = [];
  const seen = new Set<number>([0]);
  const containers: { child: number; path: string[] }[] = [{ child: root.child, path: [] }];
  while (containers.length > 0) {
    checkCancel(options);
    const container = containers.pop()!;
    if (container.child === FREE) continue;
    const treeRoot = raw[container.child];
    if (!treeRoot || treeRoot.color !== 1) invalid('invalid child tree root');
    const stack: { id: number; lower?: string; upper?: string; redParent: boolean }[] = [{ id: container.child, redParent: false }];
    while (stack.length > 0) {
      checkCancel(options);
      const node = stack.pop()!;
      const entry = raw[node.id];
      if (!entry || (entry.type !== 1 && entry.type !== 2) || seen.has(node.id)) invalid('directory alias, cycle or invalid child');
      seen.add(node.id);
      if ((node.lower !== undefined && compareCfbNames(entry.name, node.lower) <= 0)
        || (node.upper !== undefined && compareCfbNames(entry.name, node.upper) >= 0)) invalid('directory order or duplicate name');
      if (node.redParent && entry.color === 0) invalid('adjacent red directory nodes');
      const path = [...container.path, entry.name];
      if (path.length > limits.maxDepth) limited('storage hierarchy depth exceeded');
      if (entry.left !== FREE) stack.push({ id: entry.left, lower: node.lower, upper: entry.name, redParent: entry.color === 0 });
      if (entry.right !== FREE) stack.push({ id: entry.right, lower: entry.name, upper: node.upper, redParent: entry.color === 0 });
      if (entry.type === 1) {
        entries.push({ type: 'storage', path, metadata: entry.meta });
        containers.push({ child: entry.child, path });
      } else {
        const data = entry.size < MINI_CUTOFF ? readMini(entry)
          : readChain(chain(entry.start, 6, Math.ceil(entry.size / sectorSize), sectorCount), entry.size);
        entries.push({ type: 'stream', path, metadata: entry.meta, data });
      }
    }
  }
  if (seen.size !== activeEntries) invalid('unreachable allocated directory entry');
  for (let i = 0; i < sectorCount; i++) if (owners[i] === 0 && fat[i] !== FREE) invalid('orphan allocated sector');
  for (let i = 0; i < miniCount; i++) if (miniOwners[i] === 0 && miniFat[i] !== FREE) invalid('orphan allocated mini sector');
  return { majorVersion: major, minorVersion: u16(24), transactionSignature: u32(52), root: root.meta, entries };
}

interface WriteEntry {
  type: 1 | 2 | 5;
  name: string;
  path: readonly string[];
  meta: Required<CfbMetadata>;
  data?: Uint8Array;
  children: number[];
  left: number;
  right: number;
  child: number;
  start: number;
  size: number;
}

/** Writes deterministic sector layout and balanced all-black sibling trees.
 * MS-CFB 2.6.4 expressly permits all-black trees; no Office engine is involved. */
export function writeCfb(document: CfbDocument, options: CfbWriteOptions = {}): Uint8Array {
  checkCancel(options);
  const limits = limitsFor(options);
  if (!document || !Array.isArray(document.entries)) invalid('document entries must be an array');
  if (document.entries.length + 1 > limits.maxEntries) limited('directory entries exceeded');
  const major = options.majorVersion ?? document.majorVersion ?? 3;
  if (major !== 3 && major !== 4) invalid('invalid output major version');
  const minor = document.minorVersion ?? 0x003e;
  const transaction = document.transactionSignature ?? 0;
  if (!Number.isInteger(minor) || minor < 0 || minor > 0xffff || !Number.isInteger(transaction) || transaction < 0 || transaction > FREE) invalid('invalid header metadata');
  const sectorSize = major === 3 ? 512 : 4096;
  const perFat = sectorSize / 4;
  const entries: WriteEntry[] = [{ type: 5, name: 'Root Entry', path: [], meta: metadata(document.root, 'root'), children: [], left: FREE, right: FREE, child: FREE, start: END, size: 0 }];
  const ids = new Map<string, number>([[pathKey([]), 0]]);
  let inputBytes = 0;
  let miniCount = 0;
  let largeSectors = 0;
  for (const source of document.entries) {
    checkCancel(options);
    if (!source || (source.type !== 'storage' && source.type !== 'stream') || !Array.isArray(source.path) || source.path.length === 0) invalid('invalid directory entry');
    if (source.path.length > limits.maxDepth) limited('storage hierarchy depth exceeded');
    for (const component of source.path) nameValid(component);
    const key = pathKey(source.path);
    if (ids.has(key)) invalid('duplicate case-insensitive path');
    let size = 0;
    if (source.type === 'stream') {
      if (!(source.data instanceof Uint8Array)) invalid('stream data must be Uint8Array');
      size = source.data.length;
      if (size > limits.maxStreamBytes) limited('stream bytes exceeded');
      inputBytes += size;
      if (inputBytes > limits.maxFileBytes) limited('stream total exceeds file budget');
      if (size < MINI_CUTOFF) miniCount += Math.ceil(size / MINI_SIZE);
      else largeSectors += Math.ceil(size / sectorSize);
    }
    ids.set(key, entries.length);
    entries.push({ type: source.type === 'storage' ? 1 : 2, name: source.path[source.path.length - 1]!, path: source.path,
      meta: metadata(source.metadata, source.type), ...(source.type === 'stream' ? { data: source.data } : {}),
      children: [], left: FREE, right: FREE, child: FREE, start: source.type === 'storage' ? 0 : END, size });
  }
  for (let i = 1; i < entries.length; i++) {
    const entry = entries[i]!;
    const parentId = ids.get(pathKey(entry.path.slice(0, -1)));
    if (parentId === undefined || entries[parentId]!.type === 2) invalid('missing parent storage');
    // Retain one canonical spelling of every storage path.
    const parent = entries[parentId]!;
    if (parent.path.some((name, index) => name !== entry.path[index])) invalid('parent storage spelling mismatch');
    parent.children.push(i);
  }
  const buildTree = (sorted: number[], first: number, end: number): number => {
    if (first >= end) return FREE;
    const middle = Math.floor((first + end) / 2);
    const id = sorted[middle]!;
    entries[id]!.left = buildTree(sorted, first, middle);
    entries[id]!.right = buildTree(sorted, middle + 1, end);
    return id;
  };
  for (const entry of entries) {
    entry.children.sort((a, b) => compareCfbNames(entries[a]!.name, entries[b]!.name));
    entry.child = buildTree(entry.children, 0, entry.children.length);
  }
  const miniBytes = miniCount * MINI_SIZE;
  const miniStreamSectors = Math.ceil(miniBytes / sectorSize);
  const miniFatSectors = Math.ceil(miniCount / perFat);
  const directorySectors = Math.ceil(entries.length * 128 / sectorSize);
  const baseSectors = largeSectors + miniStreamSectors + miniFatSectors + directorySectors;
  let fatSectors = 0, difatSectors = 0;
  for (;;) {
    const newFat = Math.ceil((baseSectors + fatSectors + difatSectors) / perFat);
    const newDifat = Math.ceil(Math.max(0, newFat - 109) / (perFat - 1));
    if (newFat === fatSectors && newDifat === difatSectors) break;
    fatSectors = newFat; difatSectors = newDifat;
  }
  const totalSectors = baseSectors + fatSectors + difatSectors;
  const outputBytes = (totalSectors + 1) * sectorSize;
  if (outputBytes > limits.maxFileBytes) limited('serialized file bytes exceeded');
  const materialized = inputBytes + outputBytes + fatSectors * sectorSize + miniFatSectors * sectorSize;
  if (materialized > limits.maxMaterializedBytes) limited('materialized bytes exceeded');
  const output = new Uint8Array(outputBytes);
  const view = new DataView(output.buffer);
  const set16 = (offset: number, value: number): void => view.setUint16(offset, value, true);
  const set32 = (offset: number, value: number): void => view.setUint32(offset, value, true);
  const fat = new Uint32Array(fatSectors * perFat).fill(FREE);
  const miniFat = new Uint32Array(miniFatSectors * perFat).fill(FREE);
  const offsetOf = (sid: number) => (sid + 1) * sectorSize;
  const setChain = (start: number, count: number): void => {
    for (let i = 0; i < count; i++) fat[start + i] = i + 1 < count ? start + i + 1 : END;
  };
  let sectorCursor = 0;
  const miniStreamStart = largeSectors;
  let miniCursor = 0;
  for (const entry of entries) {
    checkCancel(options);
    if (entry.type !== 2 || !entry.data || entry.size === 0) continue;
    if (entry.size < MINI_CUTOFF) {
      entry.start = miniCursor;
      const count = Math.ceil(entry.size / MINI_SIZE);
      for (let i = 0; i < count; i++) miniFat[miniCursor + i] = i + 1 < count ? miniCursor + i + 1 : END;
      output.set(entry.data, offsetOf(miniStreamStart) + miniCursor * MINI_SIZE);
      miniCursor += count;
    } else {
      entry.start = sectorCursor;
      const count = Math.ceil(entry.size / sectorSize);
      setChain(sectorCursor, count);
      output.set(entry.data, offsetOf(sectorCursor));
      sectorCursor += count;
    }
  }
  entries[0]!.start = miniStreamSectors > 0 ? miniStreamStart : END;
  entries[0]!.size = miniBytes;
  setChain(miniStreamStart, miniStreamSectors);
  const miniFatStart = miniStreamStart + miniStreamSectors;
  setChain(miniFatStart, miniFatSectors);
  for (let i = 0; i < miniFat.length; i++) set32(offsetOf(miniFatStart) + i * 4, miniFat[i]!);
  const directoryStart = miniFatStart + miniFatSectors;
  setChain(directoryStart, directorySectors);
  // MS-CFB 2.6.3: even unallocated directory slots have NOSTREAM pointers;
  // the remaining bytes (including the object type) are zero.
  for (let slot = 0; slot < directorySectors * sectorSize / 128; slot++) {
    const offset = offsetOf(directoryStart) + slot * 128;
    set32(offset + 68, FREE); set32(offset + 72, FREE); set32(offset + 76, FREE);
  }
  for (let id = 0; id < entries.length; id++) {
    checkCancel(options);
    const entry = entries[id]!;
    const offset = offsetOf(directoryStart) + id * 128;
    for (let i = 0; i < entry.name.length; i++) set16(offset + i * 2, entry.name.charCodeAt(i));
    set16(offset + 64, (entry.name.length + 1) * 2);
    output[offset + 66] = entry.type;
    output[offset + 67] = 1;
    set32(offset + 68, entry.left); set32(offset + 72, entry.right); set32(offset + 76, entry.child);
    for (let i = 0; i < 16; i++) output[offset + 80 + i] = Number.parseInt(entry.meta.clsid.slice(i * 2, i * 2 + 2), 16);
    set32(offset + 96, entry.meta.stateBits);
    view.setBigUint64(offset + 100, BigInt(`0x${entry.meta.creationTime}`), true);
    view.setBigUint64(offset + 108, BigInt(`0x${entry.meta.modifiedTime}`), true);
    set32(offset + 116, entry.start);
    view.setBigUint64(offset + 120, BigInt(entry.size), true);
  }
  const fatStart = directoryStart + directorySectors;
  for (let i = 0; i < fatSectors; i++) fat[fatStart + i] = FAT_SECTOR;
  const difatStart = fatStart + fatSectors;
  for (let i = 0; i < difatSectors; i++) fat[difatStart + i] = DIFAT_SECTOR;
  for (let i = 0; i < fat.length; i++) set32(offsetOf(fatStart) + i * 4, fat[i]!);
  for (let i = 0; i < difatSectors; i++) {
    checkCancel(options);
    const offset = offsetOf(difatStart + i);
    for (let j = 0; j < perFat - 1; j++) {
      const fatIndex = 109 + i * (perFat - 1) + j;
      set32(offset + j * 4, fatIndex < fatSectors ? fatStart + fatIndex : FREE);
    }
    set32(offset + sectorSize - 4, i + 1 < difatSectors ? difatStart + i + 1 : END);
  }
  output.set(MAGIC);
  set16(24, minor); set16(26, major); set16(28, 0xfffe); set16(30, major === 3 ? 9 : 12); set16(32, 6);
  set32(40, major === 3 ? 0 : directorySectors); set32(44, fatSectors); set32(48, directoryStart);
  set32(52, transaction); set32(56, MINI_CUTOFF); set32(60, miniFatSectors > 0 ? miniFatStart : END);
  set32(64, miniFatSectors); set32(68, difatSectors > 0 ? difatStart : END); set32(72, difatSectors);
  for (let i = 0; i < 109; i++) set32(76 + i * 4, i < fatSectors ? fatStart + i : FREE);
  return output;
}
