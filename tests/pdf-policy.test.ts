// SPDX-License-Identifier: MIT
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { validatePdfBytes, validatePdfPage, validatePdfViewport, findPdfText, assertPdfFontFamily, assertPdfTextCoverage, bundledFontFamily, PDF_LIMITS } from '@office-ex/pdf';

test('PDF input rejects wrong signatures, shared bytes and over-budget page or viewport values', () => {
  validatePdfBytes(new TextEncoder().encode('%PDF-1.7\n%%EOF'));
  assert.throws(() => validatePdfBytes(new TextEncoder().encode('not a PDF document')), { code: 'invalid' });
  assert.throws(() => validatePdfBytes(new Uint8Array(new SharedArrayBuffer(20))), { code: 'unsafe' });
  assert.throws(() => validatePdfPage(1, PDF_LIMITS.pages + 1), { code: 'limit' });
  assert.throws(() => validatePdfPage(2, 1), { code: 'invalid' });
  assert.throws(() => validatePdfViewport(10000, 10000, 1), { code: 'limit' });
  assert.throws(() => validatePdfViewport(612, 792, NaN), { code: 'invalid' });
});
test('PDF search is bounded, literal and preserves UTF-16 offsets across scripts', () => {
  const text = '😀中文 a.b العربية a.b';
  assert.deepEqual(findPdfText(text, '中文'), [{ start: 2, end: 4 }]);
  assert.deepEqual(findPdfText(text, 'a.b'), [{ start: 5, end: 8 }, { start: 17, end: 20 }]);
  assert.deepEqual(findPdfText('.*', '.*'), [{ start: 0, end: 2 }]);
  assert.deepEqual(findPdfText('aaa', ''), []);
  assert.equal(findPdfText('aaaa', 'a', 2).length, 2);
  assert.throws(() => findPdfText('x', 'x'.repeat(513)), { code: 'limit' });
});
test('PDF embedding uses audited families and complete glyph coverage rather than system font names', () => {
  for (const family of ['sans-serif', 'serif', 'monospace', 'Noto Sans', 'Noto Sans CJK SC', 'Noto Naskh Arabic']) assertPdfFontFamily(family);
  assert.throws(() => assertPdfFontFamily('Unreviewed Commercial Font'), { code: 'unsupported' });
  assertPdfTextCoverage('English 中文 繁體 日本語 한국어 العربية עברית हिन्दी ไทย Tiếng Việt Русский');
  assert.throws(() => assertPdfTextCoverage('\u{10FFFF}'), { code: 'unsupported' });
  assert.match(bundledFontFamily('sans-serif', 'ja-JP'), /Noto Sans CJK SC/);
  assert.match(bundledFontFamily('sans-serif', 'zh-TW'), /Noto Sans CJK SC/);
  assert.throws(() => assertPdfFontFamily('Noto Sans CJK JP'), { code: 'unsupported' });
  assert.match(bundledFontFamily('monospace', 'ko-KR'), /^"Noto Sans Mono"/);
});
test('Bundled font bytes match reviewed OFL manifest and have unrestricted embedding flags', async () => {
  const manifest = JSON.parse(await readFile(new URL('../assets/fonts/manifest.json', import.meta.url), 'utf8')) as { files: { file: string; sha256: string; bytes: number; license: string; faces: { fsType: number; license: string }[] }[] };
  assert.equal(manifest.files.length, 8);
  for (const entry of manifest.files) {
    const bytes = await readFile(new URL('../assets/fonts/' + entry.file, import.meta.url));
    assert.equal(bytes.length, entry.bytes); assert.equal(createHash('sha256').update(bytes).digest('hex'), entry.sha256); assert.equal(entry.license, 'OFL-1.1');
    assert(entry.faces.length > 0); for (const face of entry.faces) { assert.equal(face.fsType, 0); assert.match(face.license, /Open Font License/); }
  }
});
