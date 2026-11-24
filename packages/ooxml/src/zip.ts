// SPDX-License-Identifier: MIT
// Original ZIP envelope/CRC validation; generic DEFLATE codec: fflate (MIT).
// Reference: PKWARE APPNOTE 6.3.10 (2022-11-01), 4.3 / 4.4.
// https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT
import { Deflate, Inflate } from 'fflate';
import { OfficeError } from '@office-ex/contracts';
import { bytesEqual, checkCancelled, containerLimits, joinBytes, type ContainerOptions } from './limits.js';

const encoder = new TextEncoder();
const decoder = new TextDecoder('utf-8', { fatal: true });
const CRC_TABLE = Uint32Array.from({ length: 256 }, (_, value) => {
  let c = value;
  for (let i = 0; i < 8; i++) c = (c >>> 1) ^ ((c & 1) ? 0xedb88320 : 0);
  return c >>> 0;
});
function updateCrc(crc: number, bytes: Uint8Array): number {
  for (const byte of bytes) crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ byte) & 255]!;
  return crc >>> 0;
}
export function crc32(bytes: Uint8Array): number { return (updateCrc(0xffffffff, bytes) ^ 0xffffffff) >>> 0; }

/** Archive paths stay in memory: they are never mapped to host filesystem paths. */
export function validateZipPath(name: string): string {
  if (!name || name.length > 65_535 || /[\\\u0000-\u001f\u007f]/.test(name) || name.startsWith('/') || /^[a-zA-Z]:/.test(name)) {
    throw new OfficeError('unsafe', 'Unsafe ZIP entry path');
  }
  const segments = name.replace(/\/$/, '').split('/');
  if (segments.some(segment => !segment || segment === '.' || segment === '..')) {
    throw new OfficeError('unsafe', 'Non-canonical ZIP entry path');
  }
  return name;
}
interface Entry {
  name: string; rawName: Uint8Array; flags: number; method: number;
  compressed: number; expanded: number; crc: number; local: number; data: number; end: number;
}
function invalid(message: string): never { throw new OfficeError('invalid', message); }
function extraFields(bytes: Uint8Array, start: number, length: number): void {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const end = start + length;
  for (let pos = start; pos < end;) {
    if (pos + 4 > end) invalid('Truncated ZIP extra field');
    const id = view.getUint16(pos, true), size = view.getUint16(pos + 2, true);
    if (pos + 4 + size > end) invalid('Truncated ZIP extra field');
    if (id === 1) throw new OfficeError('unsupported', 'ZIP64 is outside the M1 container profile');
    // A second filename can produce parser differentials; UTF-8 header names are supported instead.
    if (id === 0x7075) throw new OfficeError('unsupported', 'Alternate ZIP filename encoding is unsupported');
    pos += 4 + size;
  }
}
function decodeName(bytes: Uint8Array, flags: number): string {
  if (!(flags & 0x800) && bytes.some(byte => byte > 127)) {
    throw new OfficeError('unsupported', 'Non-UTF-8 ZIP names are unsupported');
  }
  try { return validateZipPath(decoder.decode(bytes)); }
  catch (error) { if (error instanceof OfficeError) throw error; return invalid('Invalid UTF-8 ZIP name'); }
}
function directory(bytes: Uint8Array, options: ContainerOptions): Entry[] {
  const limits = containerLimits(options);
  if (bytes.length > limits.maxFileBytes) throw new OfficeError('limit', 'ZIP file budget exceeded');
  if (bytes.length < 22) invalid('Truncated ZIP archive');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = -1;
  for (let pos = bytes.length - 22; pos >= Math.max(0, bytes.length - 65_557); pos--) {
    if (view.getUint32(pos, true) === 0x06054b50 && pos + 22 + view.getUint16(pos + 20, true) === bytes.length) { end = pos; break; }
  }
  if (end < 0) invalid('Missing ZIP end record');
  const count = view.getUint16(end + 10, true), size = view.getUint32(end + 12, true), offset = view.getUint32(end + 16, true);
  if (count === 0xffff || size === 0xffffffff || offset === 0xffffffff) throw new OfficeError('unsupported', 'ZIP64 is outside the M1 container profile');
  if (view.getUint16(end + 4, true) || view.getUint16(end + 6, true) || view.getUint16(end + 8, true) !== count) {
    throw new OfficeError('unsupported', 'Multi-volume ZIP is unsupported');
  }
  if (count > limits.maxEntries) throw new OfficeError('limit', 'ZIP entry count exceeded');
  if (offset + size !== end) invalid('ZIP directory range is invalid');
  const entries: Entry[] = [], names = new Set<string>();
  let pos = offset, total = 0;
  for (let index = 0; index < count; index++) {
    checkCancelled(options);
    if (pos + 46 > end || view.getUint32(pos, true) !== 0x02014b50) invalid('Truncated ZIP directory');
    const flags = view.getUint16(pos + 8, true), method = view.getUint16(pos + 10, true);
    const compressed = view.getUint32(pos + 20, true), expanded = view.getUint32(pos + 24, true);
    const nameLength = view.getUint16(pos + 28, true), extraLength = view.getUint16(pos + 30, true), commentLength = view.getUint16(pos + 32, true);
    const local = view.getUint32(pos + 42, true), crc = view.getUint32(pos + 16, true);
    if ((flags & 0x2041) !== 0) throw new OfficeError('unsafe', 'Encrypted ZIP entries are unsupported');
    if (flags & ~0x080e) throw new OfficeError('unsupported', 'Unsupported ZIP flags');
    if (method !== 0 && method !== 8) throw new OfficeError('unsupported', 'Unsupported ZIP compression');
    if (view.getUint16(pos + 6, true) > 20 || view.getUint16(pos + 34, true) !== 0) throw new OfficeError('unsupported', 'Unsupported ZIP version or disk');
    if (((view.getUint32(pos + 38, true) >>> 16) & 0xf000) === 0xa000) throw new OfficeError('unsafe', 'ZIP symlinks are forbidden');
    if (pos + 46 + nameLength + extraLength + commentLength > end) invalid('Truncated ZIP directory fields');
    const rawName = bytes.subarray(pos + 46, pos + 46 + nameLength), name = decodeName(rawName, flags);
    if (names.has(name)) invalid('Duplicate ZIP entry path');
    names.add(name);
    if (expanded > limits.maxEntryBytes || (total += expanded) > limits.maxTotalBytes || expanded > Math.max(1, compressed) * limits.maxCompressionRatio) {
      throw new OfficeError('limit', 'ZIP expansion budget exceeded');
    }
    if (compressed === 0xffffffff || expanded === 0xffffffff || local === 0xffffffff) throw new OfficeError('unsupported', 'ZIP64 is unsupported');
    if ((method === 0 && compressed !== expanded) || (name.endsWith('/') && expanded !== 0)) invalid('Invalid ZIP stored length or directory');
    extraFields(bytes, pos + 46 + nameLength, extraLength);
    if (local + 30 > offset || view.getUint32(local, true) !== 0x04034b50) invalid('Invalid ZIP local header');
    const localNameLength = view.getUint16(local + 26, true), localExtraLength = view.getUint16(local + 28, true);
    const data = local + 30 + localNameLength + localExtraLength;
    if (data + compressed > offset || view.getUint16(local + 6, true) !== flags || view.getUint16(local + 8, true) !== method ||
        !bytesEqual(rawName, bytes.subarray(local + 30, local + 30 + localNameLength))) invalid('Conflicting ZIP headers');
    if (view.getUint16(local + 4, true) > 20) throw new OfficeError('unsupported', 'Unsupported ZIP local version');
    extraFields(bytes, local + 30 + localNameLength, localExtraLength);
    let rangeEnd = data + compressed;
    if (flags & 8) {
      const localCrc = view.getUint32(local + 14, true), localCompressed = view.getUint32(local + 18, true), localExpanded = view.getUint32(local + 22, true);
      if ((localCrc && localCrc !== crc) || (localCompressed && localCompressed !== compressed) || (localExpanded && localExpanded !== expanded)) invalid('Conflicting ZIP descriptor header');
      if (rangeEnd + 12 > offset) invalid('Truncated ZIP descriptor');
      const hasSignature = view.getUint32(rangeEnd, true) === 0x08074b50;
      const descriptor = rangeEnd + (hasSignature ? 4 : 0);
      if (descriptor + 12 > offset || view.getUint32(descriptor, true) !== crc || view.getUint32(descriptor + 4, true) !== compressed || view.getUint32(descriptor + 8, true) !== expanded) invalid('Invalid ZIP descriptor');
      rangeEnd = descriptor + 12;
    } else if (view.getUint32(local + 14, true) !== crc || view.getUint32(local + 18, true) !== compressed || view.getUint32(local + 22, true) !== expanded) invalid('Conflicting ZIP lengths or CRC');
    entries.push({ name, rawName, flags, method, compressed, expanded, crc, local, data, end: rangeEnd });
    pos += 46 + nameLength + extraLength + commentLength;
  }
  if (pos !== end) invalid('ZIP directory count mismatch');
  let previousEnd = 0;
  for (const entry of [...entries].sort((a, b) => a.local - b.local)) {
    // No hidden local records or executable preambles; this is a document ZIP profile.
    if (entry.local !== previousEnd) invalid('Overlapping or unaccounted ZIP bytes');
    previousEnd = entry.end;
  }
  if (previousEnd !== offset) invalid('Unaccounted ZIP payload bytes');
  return entries;
}

/** STORE/DEFLATE, single-volume, non-ZIP64. Limits use actual emitted bytes. */
export function readZip(bytes: Uint8Array, options: ContainerOptions = {}): Map<string, Uint8Array> {
  checkCancelled(options);
  const limits = containerLimits(options), entries = directory(bytes, options), result = new Map<string, Uint8Array>();
  let total = 0;
  for (const entry of entries) {
    checkCancelled(options);
    const chunks: Uint8Array[] = [];
    let count = 0, crc = 0xffffffff, finished = false;
    const accept = (chunk: Uint8Array, final: boolean): void => {
      checkCancelled(options);
      count += chunk.length; total += chunk.length;
      if (count > entry.expanded) invalid('ZIP actual size exceeds declared size');
      if (count > limits.maxEntryBytes || total > limits.maxTotalBytes || count > Math.max(1, entry.compressed) * limits.maxCompressionRatio) throw new OfficeError('limit', 'ZIP actual expansion budget exceeded');
      crc = updateCrc(crc, chunk);
      chunks.push(chunk.slice()); finished = final;
    };
    if (entry.method === 0) {
      for (let offset = 0; offset < entry.compressed; offset += 64 * 1024) {
        accept(bytes.subarray(entry.data + offset, entry.data + Math.min(entry.compressed, offset + 64 * 1024)), offset + 64 * 1024 >= entry.compressed);
      }
      if (!entry.compressed) accept(new Uint8Array(), true);
    } else {
      try {
        const inflater = new Inflate(accept);
        // 1 KiB compressed slices bound a codec emission before its callback can stop it.
        for (let offset = 0; offset < entry.compressed; offset += 1024) {
          checkCancelled(options);
          inflater.push(bytes.subarray(entry.data + offset, entry.data + Math.min(entry.compressed, offset + 1024)), offset + 1024 >= entry.compressed);
        }
        if (!entry.compressed) invalid('Missing DEFLATE data');
      } catch (error) {
        if (error instanceof OfficeError) throw error;
        invalid('Invalid DEFLATE stream');
      }
    }
    if (!finished || count !== entry.expanded || ((crc ^ 0xffffffff) >>> 0) !== entry.crc) invalid('ZIP CRC or expanded size mismatch');
    result.set(entry.name, joinBytes(chunks, count));
  }
  return result;
}
export interface ZipWriteOptions extends ContainerOptions { compress?: boolean }

/** Deterministic headers (1980-01-01), UTF-8 filenames, no hidden filesystem metadata. */
export function writeZip(parts: ReadonlyMap<string, Uint8Array> | Record<string, Uint8Array>, options: ZipWriteOptions = {}): Uint8Array {
  checkCancelled(options);
  const limits = containerLimits(options), pairs = parts instanceof Map ? [...parts] : Object.entries(parts);
  if (pairs.length > limits.maxEntries) throw new OfficeError('limit', 'ZIP entry count exceeded');
  const locals: Uint8Array[] = [], central: Uint8Array[] = [], names = new Set<string>();
  let offset = 0, total = 0, centralSize = 0;
  for (const [path, source] of pairs) {
    checkCancelled(options);
    const name = validateZipPath(path), rawName = encoder.encode(name);
    if (names.has(name) || rawName.length > 65535) invalid('Duplicate or oversized ZIP name');
    names.add(name);
    if (!(source instanceof Uint8Array)) invalid('ZIP part must be bytes');
    if (source.length > limits.maxEntryBytes || (total += source.length) > limits.maxTotalBytes) throw new OfficeError('limit', 'ZIP materialization budget exceeded');
    if (name.endsWith('/') && source.length) invalid('ZIP directory contains data');
    let compressed = source, method = 0;
    if (options.compress !== false && source.length > 0) {
      const compressedChunks: Uint8Array[] = []; let compressedLength = 0;
      const deflater = new Deflate({ level: 6 }, chunk => {
        checkCancelled(options);
        compressedLength += chunk.length;
        if (compressedLength > limits.maxFileBytes) throw new OfficeError('limit', 'ZIP output budget exceeded');
        compressedChunks.push(chunk);
      });
      for (let pos = 0; pos < source.length; pos += 64 * 1024) {
        checkCancelled(options);
        deflater.push(source.subarray(pos, Math.min(source.length, pos + 64 * 1024)), pos + 64 * 1024 >= source.length);
      }
      // Never create an archive rejected by our ratio policy: STORE unusually repetitive data.
      if (compressedLength < source.length && source.length <= Math.max(1, compressedLength) * limits.maxCompressionRatio) { compressed = joinBytes(compressedChunks, compressedLength); method = 8; }
    }
    const local = new Uint8Array(30 + rawName.length), header = new DataView(local.buffer), crc = crc32(source);
    header.setUint32(0, 0x04034b50, true); header.setUint16(4, 20, true); header.setUint16(6, 0x800, true); header.setUint16(8, method, true);
    header.setUint16(12, 0x21, true); header.setUint32(14, crc, true); header.setUint32(18, compressed.length, true); header.setUint32(22, source.length, true); header.setUint16(26, rawName.length, true); local.set(rawName, 30);
    const directoryEntry = new Uint8Array(46 + rawName.length), directoryView = new DataView(directoryEntry.buffer);
    directoryView.setUint32(0, 0x02014b50, true); directoryView.setUint16(4, 20, true); directoryView.setUint16(6, 20, true); directoryView.setUint16(8, 0x800, true); directoryView.setUint16(10, method, true);
    directoryView.setUint16(14, 0x21, true); directoryView.setUint32(16, crc, true); directoryView.setUint32(20, compressed.length, true); directoryView.setUint32(24, source.length, true); directoryView.setUint16(28, rawName.length, true); directoryView.setUint32(42, offset, true); directoryEntry.set(rawName, 46);
    locals.push(local, compressed); central.push(directoryEntry); offset += local.length + compressed.length; centralSize += directoryEntry.length;
    if (offset + centralSize + 22 > limits.maxFileBytes) throw new OfficeError('limit', 'ZIP output budget exceeded');
  }
  const end = new Uint8Array(22), endView = new DataView(end.buffer);
  endView.setUint32(0, 0x06054b50, true); endView.setUint16(8, pairs.length, true); endView.setUint16(10, pairs.length, true); endView.setUint32(12, centralSize, true); endView.setUint32(16, offset, true);
  checkCancelled(options);
  return joinBytes([...locals, ...central, end], offset + centralSize + 22);
}
