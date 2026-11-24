/** Format-independent contracts. Implementations live in container/domain adapters. */
import { OfficeError, sortDiagnostics, type Diagnostic } from '@office-ex/contracts';

export type FormatId = 'doc' | 'docx' | 'xls' | 'xlsx' | 'ppt' | 'pptx' | 'oex' | 'pdf' | 'cfb' | 'zip' | 'unknown';
export type CompatibilityLevel = 'C0' | 'C1' | 'C2' | 'C3' | 'C4';
/** One diagnostic DTO across core, workers and UI; no format-specific fork. */
export type FormatDiagnostic = Diagnostic;
export interface Detection {
  format: FormatId;
  container: 'zip' | 'cfb' | 'json' | 'pdf' | 'unknown';
  version: string;
  encrypted: boolean;
  diagnostics: FormatDiagnostic[];
}
export interface FeatureCapability {
  featureId: string;
  read: boolean;
  write: boolean;
  preserve: boolean;
  blocked: boolean;
  level: CompatibilityLevel;
}
export interface CompatibilityReport {
  format: FormatId;
  version: string;
  readOnly: boolean;
  features: FeatureCapability[];
  diagnostics: FormatDiagnostic[];
}
export interface SourceProvenance {
  format: FormatId;
  sourceSha256: string;
  originalBytes: Uint8Array;
  /** Hashes cover logical uncompressed parts/streams, not container sector layout. */
  parts: ReadonlyMap<string, { sha256: string; bytes: Uint8Array }>;
}
export interface SavePlan {
  allowed: boolean;
  mode: 'copy-original' | 'serialize';
  targetFormat: FormatId;
  sourceContentId: string;
  blockers: FormatDiagnostic[];
  warnings: FormatDiagnostic[];
}
export interface Adapter<Model> {
  readonly format: FormatId;
  open(bytes: Uint8Array, signal?: AbortSignal): Promise<{ model: Model; provenance: SourceProvenance; capabilities: CompatibilityReport }>;
  serialize(model: Model, plan: SavePlan, signal?: AbortSignal): Promise<Uint8Array>;
}
export function assertNotCancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new OfficeError('cancelled', 'Operation cancelled');
}
/** Header only; never use this result as proof of valid document content. */
export function sniffContainer(bytes: Uint8Array): Detection['container'] {
  if (bytes.length >= 8 && [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((v, i) => bytes[i] === v)) return 'cfb';
  if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && ((bytes[2] === 3 && bytes[3] === 4) || (bytes[2] === 5 && bytes[3] === 6))) return 'zip';
  if (bytes.length >= 5 && new TextDecoder().decode(bytes.subarray(0, 5)) === '%PDF-') return 'pdf';
  let i = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  while (i < Math.min(bytes.length, 1024) && [9, 10, 13, 32].includes(bytes[i]!)) i++;
  return bytes[i] === 0x7b ? 'json' : 'unknown';
}
export class BoundedReader {
  readonly bytes: Uint8Array;
  private readonly view: DataView;
  constructor(bytes: Uint8Array) { this.bytes = bytes; this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength); }
  range(offset: number, length: number): void {
    if (!Number.isSafeInteger(offset) || !Number.isSafeInteger(length) || offset < 0 || length < 0 || offset > this.bytes.length - length) throw new OfficeError('invalid', 'Binary range outside input');
  }
  u8(offset: number): number { this.range(offset, 1); return this.view.getUint8(offset); }
  u16(offset: number): number { this.range(offset, 2); return this.view.getUint16(offset, true); }
  u32(offset: number): number { this.range(offset, 4); return this.view.getUint32(offset, true); }
  slice(offset: number, length: number): Uint8Array { this.range(offset, length); return this.bytes.slice(offset, offset + length); }
}
export function planPreservingSave(input: {
  sourceFormat: FormatId;
  targetFormat: FormatId;
  sourceContentId: string;
  currentContentId: string;
  canSerialize: boolean;
  hasUnknownDependencies: boolean;
}): SavePlan {
  const unchanged = input.sourceContentId === input.currentContentId;
  const copy = unchanged && input.sourceFormat === input.targetFormat;
  const blockers: FormatDiagnostic[] = [];
  if (!copy && (!input.canSerialize || input.hasUnknownDependencies)) blockers.push({
    code: input.hasUnknownDependencies ? 'SAVE_UNKNOWN_DEPENDENCY' : 'SAVE_ADAPTER_UNAVAILABLE', severity: 'error', featureId: 'FMT-03',
    messageKey: input.hasUnknownDependencies ? 'format.unknownDependency' : 'format.adapterUnavailable', safeAction: 'save-copy',
  });
  return { allowed: blockers.length === 0, mode: copy ? 'copy-original' : 'serialize', targetFormat: input.targetFormat,
    sourceContentId: input.currentContentId, blockers, warnings: [] };
}
export function sortedDiagnostics(items: readonly FormatDiagnostic[]): FormatDiagnostic[] {
  const rank = { error: 0, warning: 1, info: 2 };
  // Stable sort groups severity after the shared deterministic code-unit ordering.
  return sortDiagnostics(items).sort((a, b) => rank[a.severity] - rank[b.severity]);
}
