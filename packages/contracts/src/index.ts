/** Shared, runtime-independent contracts. Geometry is stored in integer EMU. */
export type DocumentKind = 'writer' | 'sheets' | 'slides';
export type OfficeFormat = 'doc' | 'docx' | 'xls' | 'xlsx' | 'ppt' | 'pptx';
export type ErrorCode = 'invalid' | 'unsupported' | 'unsafe' | 'limit' | 'conflict' | 'cancelled' | 'io' | 'internal';
export interface Failure { readonly code: ErrorCode; readonly message: string }
export type Result<T> = { readonly ok: true; readonly value: T } | { readonly ok: false; readonly error: Failure };

export class OfficeError extends Error {
  constructor(readonly code: ErrorCode, message: string) {
    super(message);
    this.name = 'OfficeError';
  }
}

export function assert(condition: unknown, code: ErrorCode, message: string): asserts condition {
  if (!condition) throw new OfficeError(code, message);
}

export function asFailure(error: unknown): Failure {
  return error instanceof OfficeError ? { code: error.code, message: error.message } : { code: 'internal', message: 'Internal operation failed' };
}

export function attempt<T>(work: () => T): Result<T> {
  try { return { ok: true, value: work() }; } catch (error) { return { ok: false, error: asFailure(error) }; }
}

export interface Diagnostic {
  readonly code: string;
  readonly severity: 'info' | 'warning' | 'error';
  readonly featureId: string;
  readonly sourceLocation?: { readonly part?: string; readonly nodeId?: string; readonly offset?: number };
  readonly messageKey: string;
  readonly safeAction: 'none' | 'read-only' | 'save-copy' | 'reject';
}

export function sortDiagnostics(items: readonly Diagnostic[]): Diagnostic[] {
  // Code-unit order is deterministic across locales and platforms.
  const key = (item: Diagnostic): string => `${item.featureId}\0${item.code}\0${item.sourceLocation?.part ?? ''}\0${item.sourceLocation?.nodeId ?? ''}\0${String(item.sourceLocation?.offset ?? 0).padStart(16, '0')}`;
  return [...items].sort((a, b) => key(a) < key(b) ? -1 : key(a) > key(b) ? 1 : 0);
}

export interface Cancellation { readonly cancelled: boolean; throwIfCancelled(): void }
export class CancellationToken implements Cancellation {
  #cancelled = false;
  get cancelled(): boolean { return this.#cancelled; }
  cancel(): void { this.#cancelled = true; }
  throwIfCancelled(): void { if (this.#cancelled) throw new OfficeError('cancelled', 'Operation cancelled'); }
}

export const EMU_PER_INCH = 914_400;
export const EMU_PER_POINT = 12_700;
export const EMU_PER_TWIP = 635;
export const EMU_PER_CSS_PIXEL = 9_525;
export const ROTATION_UNITS_PER_DEGREE = 60_000;

export function toEmu(value: number, unit: 'inch' | 'point' | 'twip' | 'css-pixel'): number {
  const scale = { inch: EMU_PER_INCH, point: EMU_PER_POINT, twip: EMU_PER_TWIP, 'css-pixel': EMU_PER_CSS_PIXEL }[unit];
  const result = Math.round(value * scale);
  assert(Number.isFinite(value) && Number.isSafeInteger(result), 'invalid', 'Invalid or out-of-range geometry');
  return result;
}

export function emuToCssPixels(value: number): number {
  assert(Number.isSafeInteger(value), 'invalid', 'Expected integer EMU');
  return value / EMU_PER_CSS_PIXEL;
}

export function degreesToRotation(value: number): number {
  const result = Math.round(value * ROTATION_UNITS_PER_DEGREE);
  assert(Number.isFinite(value) && Number.isSafeInteger(result), 'invalid', 'Invalid rotation');
  return result;
}

export const CELL_LIMITS = Object.freeze({ xls: Object.freeze({ rows: 65_536, columns: 256 }), xlsx: Object.freeze({ rows: 1_048_576, columns: 16_384 }) });
export function assertCellAddress(row: number, column: number, format: 'xls' | 'xlsx' = 'xlsx'): void {
  assert(Number.isInteger(row) && row > 0 && row <= CELL_LIMITS[format].rows && Number.isInteger(column) && column > 0 && column <= CELL_LIMITS[format].columns, 'invalid', 'Cell address is outside format limits');
}

export function assertId(value: unknown): asserts value is string {
  assert(typeof value === 'string' && /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,127}$/.test(value), 'invalid', 'Invalid stable identifier');
}

export function assertUnicode(value: string): void {
  for (let i = 0; i < value.length; i++) {
    const code = value.charCodeAt(i);
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(++i);
      assert(next >= 0xdc00 && next <= 0xdfff, 'invalid', 'Unpaired Unicode surrogate');
    } else assert(code < 0xdc00 || code > 0xdfff, 'invalid', 'Unpaired Unicode surrogate');
  }
}

export function graphemeBoundaries(text: string): number[] {
  assertUnicode(text);
  const boundaries = Array.from(new Intl.Segmenter('und', { granularity: 'grapheme' }).segment(text), (segment) => segment.index);
  boundaries.push(text.length);
  return boundaries;
}

export function assertGraphemeRange(text: string, start: number, end: number): void {
  assert(Number.isInteger(start) && Number.isInteger(end) && start >= 0 && start <= end && end <= text.length, 'invalid', 'Invalid UTF-16 text range');
  const boundaries = new Set(graphemeBoundaries(text));
  assert(boundaries.has(start) && boundaries.has(end), 'invalid', 'Text edit splits a grapheme');
}

export function replaceGraphemeRange(text: string, start: number, end: number, replacement: string): string {
  assertUnicode(replacement);
  assertGraphemeRange(text, start, end);
  return text.slice(0, start) + replacement + text.slice(end);
}

export function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value) && (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null);
}
