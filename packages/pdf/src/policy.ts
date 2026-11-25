// SPDX-License-Identifier: MIT
import { assert, OfficeError } from '@office-ex/contracts';
import { PDF_FONT_COVERAGE } from './font-coverage.js';

/** Application budgets, not limits of the PDF format. */
export const PDF_LIMITS = Object.freeze({
  fileBytes: 100 * 1024 * 1024, pages: 500, pagePixels: 40_000_000,
  pageTextItems: 100_000, pageTextCharacters: 2_000_000,
  searchCharacters: 8_000_000, searchMatches: 10_000, queryCharacters: 512,
  operationMilliseconds: 30_000, printPixels: 100_000_000,
});

export function validatePdfBytes(bytes: Uint8Array): void {
  assert(bytes instanceof Uint8Array && bytes.byteLength > 8, 'invalid', 'Invalid PDF bytes');
  assert(!(typeof SharedArrayBuffer !== 'undefined' && bytes.buffer instanceof SharedArrayBuffer), 'unsafe', 'Shared PDF bytes are not accepted');
  assert(bytes.byteLength <= PDF_LIMITS.fileBytes, 'limit', 'PDF file budget exceeded');
  const prefix = new TextDecoder('latin1').decode(bytes.subarray(0, Math.min(1024, bytes.length)));
  assert(/%PDF-\d\.\d/.test(prefix), 'invalid', 'Missing PDF header');
}

export function validatePdfPage(number: number, total: number): void {
  assert(Number.isSafeInteger(total) && total > 0 && total <= PDF_LIMITS.pages, 'limit', 'PDF page budget exceeded');
  assert(Number.isSafeInteger(number) && number >= 1 && number <= total, 'invalid', 'Invalid PDF page');
}

export function validatePdfViewport(width: number, height: number, scale: number): void {
  assert(Number.isFinite(scale) && scale >= 0.1 && scale <= 6, 'invalid', 'Invalid PDF zoom');
  assert(Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0 && width <= 20000 && height <= 20000 && Math.ceil(width) * Math.ceil(height) <= PDF_LIMITS.pagePixels, 'limit', 'PDF page raster budget exceeded');
}

export interface PdfMatch { readonly start: number; readonly end: number }
/** Literal, case-sensitive search preserves UTF-16 indices and source Unicode.
 * There is no regex evaluation, normalization or replacement of document data. */
export function findPdfText(text: string, query: string, maxMatches: number = PDF_LIMITS.searchMatches): PdfMatch[] {
  assert(text.length <= PDF_LIMITS.pageTextCharacters, 'limit', 'PDF text budget exceeded');
  assert(query.length <= PDF_LIMITS.queryCharacters, 'limit', 'PDF query budget exceeded');
  assert(Number.isSafeInteger(maxMatches) && maxMatches > 0 && maxMatches <= PDF_LIMITS.searchMatches, 'invalid', 'Invalid PDF search budget');
  if (!query) return [];
  const result: PdfMatch[] = [];
  for (let start = text.indexOf(query); start >= 0 && result.length < maxMatches; start = text.indexOf(query, start + Math.max(1, query.length))) result.push({ start, end: start + query.length });
  return result;
}

const FAMILIES = ['Noto Sans', 'Noto Serif', 'Noto Sans Mono', 'Noto Naskh Arabic', 'Noto Sans Hebrew', 'Noto Sans Devanagari', 'Noto Sans Thai', 'Noto Sans CJK SC'] as const;
export const PDF_EMBEDDABLE_FONT_FAMILIES: readonly string[] = Object.freeze([...FAMILIES]);
export function approvedPdfFontFamily(family: string | undefined): boolean {
  return family === undefined || ['sans-serif', 'serif', 'monospace', 'system-ui'].includes(family) || PDF_EMBEDDABLE_FONT_FAMILIES.includes(family);
}
export function assertPdfFontFamily(family: string | undefined): void {
  assert(approvedPdfFontFamily(family), 'unsupported', 'pdf.font-not-approved');
}

export function bundledFontFamily(family = 'sans-serif', language = 'und'): string {
  const concrete = family === 'serif' ? 'Noto Serif' : family === 'monospace' ? 'Noto Sans Mono' : family === 'sans-serif' || family === 'system-ui' ? 'Noto Sans' : family;
  // The browser's shaping backend uses each DOM run's preserved lang for CJK
  // locl features. TTC face fragments are not supported reliably by Chromium;
  // the bundle therefore supplies one real, named SC OpenType face, never
  // misleading aliases that all decode as the collection's first JP face.
  void language;
  return [concrete, 'Noto Sans', 'Noto Naskh Arabic', 'Noto Sans Hebrew', 'Noto Sans Devanagari', 'Noto Sans Thai', 'Noto Sans CJK SC'].filter((value, i, all) => all.indexOf(value) === i).map(value => JSON.stringify(value)).join(', ');
}

/** Never silently export missing-glyph rectangles through a system fallback. */
export function assertPdfTextCoverage(text: string): void {
  for (const character of text) {
    const value = character.codePointAt(0)!;
    if (value <= 0x20 || (value >= 0x200b && value <= 0x200f) || (value >= 0x2028 && value <= 0x202e) || (value >= 0x2060 && value <= 0x2069) || (value >= 0xfe00 && value <= 0xfe0f)) continue;
    let low = 0, high = PDF_FONT_COVERAGE.length - 1, found = false;
    while (low <= high) { const middle = (low + high) >>> 1, range = PDF_FONT_COVERAGE[middle]!; if (value < range[0]) high = middle - 1; else if (value > range[1]) low = middle + 1; else { found = true; break; } }
    if (!found) throw new OfficeError('unsupported', 'pdf.font-glyph-unavailable');
  }
}
