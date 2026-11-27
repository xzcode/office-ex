// SPDX-License-Identifier: MIT
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { PdfActualTextBuilder } from '@office-ex/pdf';

const glyph = (str: string, x: number, width = 8) => ({ str, dir: 'ltr', width, height: 16, transform: [16, 0, 0, 16, x, 100], fontName: 'f1', hasEOL: false });
const begin = (actualText: string | null, tag = 'Span') => ({ type: 'beginMarkedContentProps', tag, actualText });
const end = { type: 'endMarkedContent' };
test('PDF ActualText restores exact CJK and Devanagari logical source without changing page glyphs', () => {
  const builder = new PdfActualTextBuilder();
  builder.push(begin('文')); builder.push(glyph('⽂', 10)); builder.push(end);
  builder.push(begin('हि')); builder.push(glyph('\0', 20, 4)); builder.push(glyph('ह', 24, 9)); builder.push(end);
  const items = builder.finish();
  assert.deepEqual(items.map(item => item.str), ['文', 'हि']); assert.equal(items[1]!.width, 13); assert.equal(items[1]!.transform[4], 20);
});
test('PDF ReversedChars preserves logical Arabic order across nested ActualText spans', () => {
  const builder = new PdfActualTextBuilder(); builder.push({ type: 'beginMarkedContent', tag: 'ReversedChars' });
  builder.push(begin('ب')); builder.push(glyph('ﺏ', 10)); builder.push(end);
  builder.push(begin('ا')); builder.push(glyph('ﺍ', 18)); builder.push(end); builder.push(end);
  const items = builder.finish(); assert.equal(items.length, 1); assert.equal(items[0]!.str, 'اب'); assert.equal(items[0]!.dir, 'rtl'); assert.equal(items[0]!.width, 16);
});
test('PDF ActualText cannot introduce unbounded, unbalanced or unmeasurable content', () => {
  assert.throws(() => new PdfActualTextBuilder().push(end), { code: 'invalid' });
  const unclosed = new PdfActualTextBuilder(); unclosed.push(begin(null)); assert.throws(() => unclosed.finish(), { code: 'invalid' });
  const deep = new PdfActualTextBuilder(); for (let i = 0; i < 128; i++) deep.push(begin(null)); assert.throws(() => deep.push(begin(null)), { code: 'limit' });
  const noGlyph = new PdfActualTextBuilder(); noGlyph.push(begin('must not invent invisible searchable text')); assert.throws(() => noGlyph.push(end), { code: 'unsupported' });
  const long = new PdfActualTextBuilder(); assert.throws(() => long.push(begin('a'.repeat(2_000_001))), { code: 'limit' });
});
