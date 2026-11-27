// SPDX-License-Identifier: MIT
/** PDF marked-content ActualText is the author's logical text replacement,
 * especially for ligatures, pre-base vowels, bidi and shared CJK glyphs. It is
 * separate from the glyph stream used to draw the page. */
import { assert, isRecord } from '@office-ex/contracts';
import type { TextItem } from 'pdfjs-dist/types/src/display/api.js';
import { PDF_LIMITS } from './policy.js';

interface Frame { tag: string | null; actualText: string | null; items: TextItem[] }
const RTL = /[\p{Script=Arabic}\p{Script=Hebrew}]/u;
function merged(items: readonly TextItem[], text: string, rtl = false): TextItem[] {
  if (!items.length) { assert(text.length === 0, 'unsupported', 'PDF ActualText has no measurable glyphs'); return []; }
  const first = items[0]!, matrix = [...first.transform] as number[];
  assert(matrix.length === 6 && matrix.every(Number.isFinite), 'invalid', 'Invalid PDF text geometry');
  const length = Math.hypot(matrix[0]!, matrix[1]!); assert(length > 0, 'invalid', 'Invalid PDF text scale');
  const ux = matrix[0]! / length, uy = matrix[1]! / length;
  let left = 0, right = 0, height = 0;
  for (const item of items) {
    assert(item.transform.length === 6 && item.transform.every(Number.isFinite) && Number.isFinite(item.width) && Number.isFinite(item.height), 'invalid', 'Invalid PDF text bounds');
    const delta = (item.transform[4] - matrix[4]!) * ux + (item.transform[5] - matrix[5]!) * uy;
    left = Math.min(left, delta); right = Math.max(right, delta + Math.abs(item.width)); height = Math.max(height, Math.abs(item.height));
  }
  matrix[4]! += ux * left; matrix[5]! += uy * left;
  return [{ ...first, str: text, dir: rtl || RTL.test(text) ? 'rtl' : first.dir, width: right - left, height, transform: matrix, hasEOL: items.some(item => item.hasEOL) }];
}

export class PdfActualTextBuilder {
  private readonly root: TextItem[] = [];
  private readonly stack: Frame[] = [];
  private inputItems = 0;
  private characters = 0;
  private records = 0;
  private append(items: readonly TextItem[]): void { (this.stack.at(-1)?.items ?? this.root).push(...items); }
  push(value: unknown): void {
    assert(++this.records <= 1_000_000 && isRecord(value), 'limit', 'PDF text record budget exceeded');
    if ('str' in value) {
      assert(typeof value.str === 'string' && ++this.inputItems <= PDF_LIMITS.pageTextItems && (this.characters += value.str.length) <= PDF_LIMITS.pageTextCharacters, 'limit', 'PDF text budget exceeded');
      assert(Array.isArray(value.transform) && value.transform.length === 6 && value.transform.every(x => typeof x === 'number' && Number.isFinite(x)) && typeof value.width === 'number' && typeof value.height === 'number' && typeof value.fontName === 'string' && typeof value.dir === 'string' && typeof value.hasEOL === 'boolean', 'invalid', 'Invalid PDF text item');
      this.append([value as unknown as TextItem]); return;
    }
    if (value.type === 'beginMarkedContent' || value.type === 'beginMarkedContentProps') {
      assert(this.stack.length < 128 && (value.tag === null || typeof value.tag === 'string'), 'limit', 'PDF marked-content depth exceeded');
      const actualText = typeof value.actualText === 'string' ? value.actualText : null;
      if (actualText !== null) assert((this.characters += actualText.length) <= PDF_LIMITS.pageTextCharacters, 'limit', 'PDF ActualText budget exceeded');
      this.stack.push({ tag: value.tag, actualText, items: [] }); return;
    }
    assert(value.type === 'endMarkedContent' && this.stack.length > 0, 'invalid', 'Unbalanced PDF marked content');
    const frame = this.stack.pop()!;
    if (frame.actualText !== null) this.append(merged(frame.items, frame.actualText));
    else if (frame.tag === 'ReversedChars') { const items = frame.items.slice().reverse(); this.append(merged(frame.items, items.map(item => item.str).join(''), true)); }
    else this.append(frame.items);
  }
  finish(): TextItem[] { assert(this.stack.length === 0, 'invalid', 'Unclosed PDF marked content'); return this.root; }
}
