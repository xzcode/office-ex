import assert from 'node:assert/strict';
import { docxParts, xlsxParts } from '../packages/ooxml/src/export.ts';

const doc = docxParts([{runs: [{text: ' 中文 & <Office> 😀 ', bold: true}, {text: '\n第二行\t尾部'}]}]);
const wb = xlsxParts([{name: '数据 & A', cells: [
  {row: 1, col: 1, value: '=plain text'}, {row: 2, col: 2, value: 12},
  {row: 3, col: 2, value: 30}, {row: 4, col: 2, value: {formula: '=SUM(B2:B3)'}},
  {row: 1, col: 27, value: '中文 <&>'}
]}, {name: '第二页', cells: [{row: 1, col: 1, value: 42}]}]);
assert.throws(() => docxParts([{runs: [{text: '\u0000'}]}]), /Invalid XML/);
assert.throws(() => xlsxParts([{name: 'A', cells: []}, {name: 'a', cells: []}]), /duplicate/);
assert.throws(() => xlsxParts([{name: 'A', cells: [{row: 1, col: 16385, value: 0}]}]), /bounds/);
assert.throws(() => xlsxParts([{name: 'A', cells: [{row: 1, col: 1, value: NaN}]}]), /Non-finite/);
process.stdout.write(JSON.stringify({'writer.docx': doc, 'sheets.xlsx': wb}));
