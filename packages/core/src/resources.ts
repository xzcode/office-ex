import { assert, assertId, type Cancellation, type Diagnostic } from '@office-ex/contracts';
import { cloneData, freezeData, validateTextStyle, type NamedStyle, type ResourceRef, type TextStyle } from './model.js';

// SHA-256 is a format-neutral resource identity, not a signature or trust decision.
const ROUND_CONSTANTS = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);
const rotate = (value: number, bits: number): number => (value >>> bits) | (value << (32 - bits));
/** Deterministic synchronous implementation for pure domain and worker runtimes. */
export function sha256(bytes: Uint8Array, cancellation?: Cancellation): string {
  assert(bytes.length <= 512 * 1024 * 1024, 'limit', 'Digest input exceeds budget');
  cancellation?.throwIfCancelled();
  const state = new Uint32Array([0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19]);
  const words = new Uint32Array(64);
  const compress = (view: DataView, offset: number): void => {
    for (let i = 0; i < 16; i++) words[i] = view.getUint32(offset + i * 4);
    for (let i = 16; i < 64; i++) {
      const first = words[i - 15]!; const second = words[i - 2]!;
      const small0 = rotate(first, 7) ^ rotate(first, 18) ^ (first >>> 3); const small1 = rotate(second, 17) ^ rotate(second, 19) ^ (second >>> 10);
      words[i] = (words[i - 16]! + small0 + words[i - 7]! + small1) >>> 0;
    }
    let [a, b, c, d, e, f, g, h] = state as Uint32Array & [number, number, number, number, number, number, number, number];
    for (let i = 0; i < 64; i++) {
      const big1 = rotate(e, 6) ^ rotate(e, 11) ^ rotate(e, 25); const choose = (e & f) ^ (~e & g);
      const first = (h + big1 + choose + ROUND_CONSTANTS[i]! + words[i]!) >>> 0;
      const big0 = rotate(a, 2) ^ rotate(a, 13) ^ rotate(a, 22); const majority = (a & b) ^ (a & c) ^ (b & c);
      const second = (big0 + majority) >>> 0;
      h = g; g = f; f = e; e = (d + first) >>> 0; d = c; c = b; b = a; a = (first + second) >>> 0;
    }
    for (const [index, value] of [a, b, c, d, e, f, g, h].entries()) state[index] = (state[index]! + value) >>> 0;
  };
  // Only the final block(s) are copied. Large OPC parts do not double their peak memory.
  const completeLength = Math.floor(bytes.length / 64) * 64;
  const input = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let offset = 0; offset < completeLength; offset += 64) {
    if ((offset & 0xffff) === 0) cancellation?.throwIfCancelled();
    compress(input, offset);
  }
  const remainder = bytes.length - completeLength;
  const tailLength = remainder < 56 ? 64 : 128;
  const tail = new Uint8Array(tailLength); tail.set(bytes.subarray(completeLength)); tail[remainder] = 0x80;
  const view = new DataView(tail.buffer); const bitLength = BigInt(bytes.length) * 8n;
  view.setUint32(tailLength - 8, Number(bitLength >> 32n)); view.setUint32(tailLength - 4, Number(bitLength & 0xffffffffn));
  for (let offset = 0; offset < tailLength; offset += 64) compress(view, offset);
  cancellation?.throwIfCancelled();
  return Array.from(state, (value) => value.toString(16).padStart(8, '0')).join('');
}

/** Blob ownership is separate from document DTOs; clients receive defensive byte copies. */
export class ResourceStore {
  #resources = new Map<string, { metadata: ResourceRef; bytes: Uint8Array }>();
  #bytes = 0;
  constructor(readonly byteLimit = 128 * 1024 * 1024) { assert(Number.isSafeInteger(byteLimit) && byteLimit >= 0, 'invalid', 'Invalid resource budget'); }
  put(id: string, bytes: Uint8Array, mimeType: string, expectedSha256?: string): ResourceRef {
    assertId(id); assert(bytes instanceof Uint8Array, 'invalid', 'Expected resource bytes');
    assert(bytes.byteLength <= this.byteLimit, 'limit', 'Resource budget exceeded');
    assert(/^[a-z0-9!#$&^_.+-]+\/[a-z0-9!#$&^_.+-]+$/i.test(mimeType), 'invalid', 'Invalid resource media type');
    const existing = this.#resources.get(id);
    assert(this.#bytes + (existing ? 0 : bytes.byteLength) <= this.byteLimit, 'limit', 'Resource budget exceeded');
    const digest = sha256(bytes); if (expectedSha256 !== undefined) assert(digest === expectedSha256, 'invalid', 'Resource hash mismatch');
    if (existing) { assert(existing.metadata.sha256 === digest && existing.metadata.mimeType === mimeType && existing.metadata.byteLength === bytes.length, 'conflict', 'Resource ID already names different content'); return existing.metadata; }
    const metadata = Object.freeze({ id, sha256: digest, mimeType, byteLength: bytes.byteLength });
    this.#resources.set(id, { metadata, bytes: bytes.slice() }); this.#bytes += bytes.byteLength; return metadata;
  }
  get(id: string): Uint8Array { const resource = this.#resources.get(id); assert(resource, 'invalid', 'Missing resource'); return resource.bytes.slice(); }
  metadata(id: string): ResourceRef { const resource = this.#resources.get(id); assert(resource, 'invalid', 'Missing resource'); return resource.metadata; }
  get byteLength(): number { return this.#bytes; }
  get size(): number { return this.#resources.size; }
  entries(): readonly ResourceRef[] { return Object.freeze(Array.from(this.#resources.values(), (value) => value.metadata)); }
  /** No implicit garbage collection: the adapter must have proven the full reference closure. */
  removeUnreferenced(ids: readonly string[], referenceProof: { readonly complete: boolean; readonly referencedIds: readonly string[] }): void {
    assert(referenceProof.complete === true, 'unsupported', 'Unknown references prohibit resource cleanup');
    const referenced = new Set(referenceProof.referencedIds);
    for (const id of ids) assert(!referenced.has(id), 'conflict', 'A resource is still referenced');
    for (const id of ids) { const resource = this.#resources.get(id); if (resource) { this.#bytes -= resource.bytes.byteLength; this.#resources.delete(id); } }
  }
}

export function resolveTextStyle(defaults: TextStyle, styles: Readonly<Record<string, NamedStyle>>, namedIds: readonly string[] = [], direct: TextStyle = {}): { style: Readonly<TextStyle>; diagnostics: readonly Diagnostic[] } {
  validateTextStyle(defaults); validateTextStyle(direct);
  let style: TextStyle = cloneData(defaults); const diagnostics: Diagnostic[] = [];
  const append = (id: string, ancestors: Set<string>): void => {
    assert(!ancestors.has(id), 'invalid', 'Style inheritance cycle'); assert(ancestors.size < 128, 'limit', 'Style inheritance budget exceeded');
    const named = styles[id];
    if (!named) { diagnostics.push({ code: 'style.missing', severity: 'warning', featureId: 'WRT-02', sourceLocation: { nodeId: id }, messageKey: 'style.missing', safeAction: 'read-only' }); return; }
    ancestors.add(id); if (named.basedOn) append(named.basedOn, ancestors); ancestors.delete(id);
    validateTextStyle(named.properties); style = { ...style, ...cloneData(named.properties) };
    for(const key of named.toggles??[]){assert((key==='bold'||key==='italic')&&!Object.hasOwn(named.properties,key),'invalid','Invalid style toggle');style[key]=!style[key];}
  };
  for (const id of namedIds) append(id, new Set());
  return freezeData({ style: { ...style, ...cloneData(direct) }, diagnostics });
}
