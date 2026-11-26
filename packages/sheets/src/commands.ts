/** Recovered atomic spreadsheet transactions; source model is never mutated. */
import { assert, assertCellAddress } from '@office-ex/contracts';
import { patchCommand, cloneData, type Cell, type CellStyle, type SheetConditionalRule, type SheetChart, type Sheet, type SheetView, type SheetsDocument, type DocumentSnapshot, type DocumentCommand, type PatchOperation } from '@office-ex/core';
import { parseFormula, serializeFormula, formulaIsSupported, type FormulaNode, type FormulaReference, type FormulaRange } from './formula.js';
import { cellFromInput, formatCell, type CellInputOptions } from './formatting.js';
import type { WorkbookSnapshot, CalculationResult } from './calculation.js';
export interface CellRange { row:number;column:number;rowCount:number;columnCount:number }
export interface CellEdit { row:number;column:number;cell:Cell|null }
type SheetSnapshot = WorkbookSnapshot['sheets'][number];
type Axis = 'rows'|'columns';
interface Dimension {height?:number;width?:number;hidden?:boolean}
export function cellKey(row:number, column:number):string {
  assertCellAddress(row, column);
  return `${row}:${column}`;
}
export function assertCellRange(range:CellRange, maxCells = Number.MAX_SAFE_INTEGER):void {
  assertCellAddress(range.row, range.column);
  assert(Number.isSafeInteger(range.rowCount) && range.rowCount > 0 && Number.isSafeInteger(range.columnCount) && range.columnCount > 0, "invalid", "Invalid cell range");
  assertCellAddress(range.row + range.rowCount - 1, range.column + range.columnCount - 1);
  assert(range.rowCount * range.columnCount <= maxCells, "limit", "Selected cell operation exceeds budget");
}
export const inside = (range:CellRange, row:number, column:number) => row >= range.row && row < range.row + range.rowCount && column >= range.column && column < range.column + range.columnCount;
export const overlap = (a:CellRange, b:CellRange) => a.row < b.row + b.rowCount && b.row < a.row + a.rowCount && a.column < b.column + b.columnCount && b.column < a.column + a.columnCount;
export const contains = (a:CellRange, b:CellRange) => inside(a, b.row, b.column) && inside(a, b.row + b.rowCount - 1, b.column + b.columnCount - 1);
function book(snapshot:DocumentSnapshot):WorkbookSnapshot {
  assert(snapshot.document.kind === "sheets", "invalid", "Expected sheets document");
  return snapshot.document;
}
function sheetIndex(snapshot:DocumentSnapshot, id:string):number {
  const index = book(snapshot).sheets.findIndex((s) => s.id === id);
  assert(index >= 0, "invalid", "Sheet does not exist");
  return index;
}
function edit(snapshot:DocumentSnapshot, operations:PatchOperation[]):DocumentCommand {
  return patchCommand(snapshot, "SHT-02", operations);
}
function cellsCommand(snapshot:DocumentSnapshot, index:number, cells:Record<string,Cell>):DocumentCommand {
  return edit(snapshot, [{ op: "set", path: ["sheets", index, "cells"], value: cells }]);
}
function assertEditableCoordinate(sheet:SheetSnapshot, row:number, column:number):void {
  const merged = sheet.merges.find((m) => inside(m, row, column));
  assert(!merged || row === merged.row && column === merged.column, "unsupported", "Edit the merged cell anchor");
}
export function setCells(snapshot:DocumentSnapshot, sheetId:string, edits:readonly CellEdit[]):DocumentCommand {
  assert(edits.length <= 1e5, "limit", "Cell edit budget exceeded");
  const index = sheetIndex(snapshot, sheetId), sheet = book(snapshot).sheets[index], keys = /* @__PURE__ */ new Set();
  if (edits.length <= 16) return edit(snapshot, edits.map((e) => {
    const key = cellKey(e.row, e.column);
    assert(!keys.has(key), "invalid", "Duplicate cell in edit batch");
    keys.add(key);
    assertEditableCoordinate(sheet, e.row, e.column);
    return e.cell === null ? { op: "delete", path: ["sheets", index, "cells", key] } : { op: "set", path: ["sheets", index, "cells", key], value: e.cell };
  }));
  const cells = { ...sheet.cells };
  for (const e of edits) {
    const key = cellKey(e.row, e.column);
    assert(!keys.has(key), "invalid", "Duplicate cell in edit batch");
    keys.add(key);
    assertEditableCoordinate(sheet, e.row, e.column);
    if (e.cell === null) delete cells[key];
    else cells[key] = e.cell;
  }
  return cellsCommand(snapshot, index, cells);
}
function checkedAst(source:string):FormulaNode {
  const ast = parseFormula(source);
  assert(formulaIsSupported(ast), "unsupported", "Unknown formulas prevent structural edits");
  return ast;
}
function transform(ast:FormulaNode, ref:(r:FormulaReference)=>FormulaNode, range?:(r:FormulaRange)=>FormulaNode):FormulaNode {
  if (ast.kind === "reference") return ref(ast);
  if (ast.kind === "range") {
    if (range) return range(ast);
    const a = ref(ast.start), b = ref(ast.end);
    return a.kind !== "reference" || b.kind !== "reference" ? { kind: "error", value: "#REF!" } : { ...ast, start: a, end: b };
  }
  if (ast.kind === "binary") return { ...ast, left: transform(ast.left, ref, range), right: transform(ast.right, ref, range) };
  if (ast.kind === "unary") return { ...ast, operand: transform(ast.operand, ref, range) };
  if (ast.kind === "call") return { ...ast, args: ast.args.map((a) => transform(a, ref, range)) };
  return ast;
}
function rewrite(sheets:readonly SheetSnapshot[], apply:(ast:FormulaNode,sheet:SheetSnapshot)=>FormulaNode):Sheet[] {
  return sheets.map((sheet) => {
    const cells:Record<string,Cell> = {};
    for (const [key, cell] of Object.entries(sheet.cells)) {
      if (cell.value.kind === "formula") {
        const ast = checkedAst(cell.value.source), after = apply(ast, sheet);
        const same = JSON.stringify(ast) === JSON.stringify(after);
        cells[key] = { ...cloneData(cell), value: { kind: "formula", source: same ? cell.value.source : serializeFormula(after) } };
      } else cells[key] = cloneData(cell);
    }
    return { ...cloneData(sheet), cells } as Sheet;
  });
}
function setSheets(snapshot:DocumentSnapshot, sheets:Sheet[]):DocumentCommand {
  return edit(snapshot, [{ op: "set", path: ["sheets"], value: sheets }]);
}
export function renameSheet(snapshot:DocumentSnapshot, sheetId:string, name:string):DocumentCommand {
  const index = sheetIndex(snapshot, sheetId), old = book(snapshot).sheets[index].name.toLowerCase();
  const sheets = rewrite(book(snapshot).sheets, (ast) => transform(ast, (r) => r.sheet?.toLowerCase() === old ? { ...r, sheet: name } : r));
  sheets[index].name = name;
  return setSheets(snapshot, sheets);
}
export function addSheet(snapshot:DocumentSnapshot, name:string, id?:string, index?:number):DocumentCommand {
  const document2 = book(snapshot);
  const sheets = rewrite(document2.sheets, (ast) => ast);
  let next = 1;
  while (sheets.some((s) => s.id === `sheet-${next}`)) next++;
  const identifier = id ?? `sheet-${next}`, target = index ?? sheets.length;
  assert(Number.isInteger(target) && target >= 0 && target <= sheets.length, "invalid", "Invalid sheet insertion index");
  assert(!sheets.some((s) => s.id === identifier), "invalid", "Duplicate sheet ID");
  sheets.splice(target, 0, { id: identifier, name, cells: {}, rows: {}, columns: {}, merges: [] });
  return setSheets(snapshot, sheets);
}
export function deleteSheet(snapshot:DocumentSnapshot, sheetId:string):DocumentCommand {
  const document2 = book(snapshot), index = sheetIndex(snapshot, sheetId);
  assert(document2.sheets.length > 1, "invalid", "A workbook must retain one sheet");
  const old = document2.sheets[index].name.toLowerCase();
  const sheets = rewrite(document2.sheets.filter((s) => s.id !== sheetId), (ast) => transform(ast, (r) => r.sheet?.toLowerCase() === old ? { kind: "error", value: "#REF!" } : r));
  return setSheets(snapshot, sheets);
}
export function moveSheet(snapshot:DocumentSnapshot, sheetId:string, index:number):DocumentCommand {
  const previous = sheetIndex(snapshot, sheetId), sheets = rewrite(book(snapshot).sheets, (ast) => ast);
  assert(Number.isInteger(index) && index >= 0 && index < sheets.length, "invalid", "Invalid sheet order");
  const [sheet] = sheets.splice(previous, 1);
  sheets.splice(index, 0, sheet);
  return setSheets(snapshot, sheets);
}
function shiftInterval(first:number, last:number, start:number, count:number, removing:boolean):[number,number]|undefined {
  if (!removing) return [first >= start ? first + count : first, last >= start ? last + count : last];
  const end = start + count - 1;
  if (last < start) return [first, last];
  if (first > end) return [first - count, last - count];
  if (first >= start && last <= end) return void 0;
  const a = first < start ? first : start, b = last > end ? last - count : start - 1;
  return a <= b ? [a, b] : void 0;
}
function dimensions2(snapshot:DocumentSnapshot, sheetId:string, axis:Axis, start:number, count:number, removing:boolean):DocumentCommand {
  const index = sheetIndex(snapshot, sheetId), target = book(snapshot).sheets[index], dimension = axis === "rows" ? "row" : "column", size = axis === "rows" ? "rowCount" : "columnCount", maximum = axis === "rows" ? 1048576 : 16384;
  assert(Number.isSafeInteger(start) && start >= 1 && Number.isSafeInteger(count) && count >= 1 && start + count - 1 <= maximum, "invalid", "Invalid row or column operation");
  const refers = (ref:FormulaReference, sheet2:SheetSnapshot) => ref.sheet === void 0 ? sheet2.id === sheetId : ref.sheet.toLowerCase() === target.name.toLowerCase();
  const sheets = rewrite(book(snapshot).sheets, (ast, sheet2) => transform(ast, (r) => {
    if (!refers(r, sheet2)) return r;
    const value = r[dimension];
    if (removing && value >= start && value < start + count) return { kind: "error", value: "#REF!" };
    const next = value >= start ? value + (removing ? -count : count) : value;
    return next > maximum ? { kind: "error", value: "#REF!" } : { ...r, [dimension]: next };
  }, (r) => {
    if (!refers(r.start, sheet2) || r.axis === "columns" && axis === "rows" || r.axis === "rows" && axis === "columns") return r;
    const backwards = r.start[dimension] > r.end[dimension], bounds = shiftInterval(Math.min(r.start[dimension], r.end[dimension]), Math.max(r.start[dimension], r.end[dimension]), start, count, removing);
    if (!bounds || bounds[1] > maximum) return { kind: "error", value: "#REF!" };
    return { ...r, start: { ...r.start, [dimension]: bounds[backwards ? 1 : 0] }, end: { ...r.end, [dimension]: bounds[backwards ? 0 : 1] } };
  }));
  const sheet = sheets[index], cells:Record<string,Cell> = {};
  // All chart ranges are local, bounded single-column vectors. Rebase
  // anchors and source references atomically. Reject an edit that would
  // destroy a source vector, rather than silently remove a chart or series.
  if(sheet.charts){
    const shift=(range:CellRange):CellRange=>{
      const bounds=shiftInterval(range[dimension],range[dimension]+range[size]-1,start,count,removing);
      assert(bounds&&bounds[1]<=maximum,'unsupported','Operation would delete or overflow a chart data range');
      return {...range,[dimension]:bounds[0],[size]:bounds[1]-bounds[0]+1};
    };
    const position=(at:number,anchor=false)=>{
      if(removing&&at>=start&&at<start+count){assert(anchor,'unsupported','Operation would delete a chart series label');return Math.min(start,maximum);}
      const result=at>=start?at+(removing?-count:count):at;
      assert(result<=maximum,'unsupported','Operation would overflow a chart reference');return result;
    };
    sheet.charts=sheet.charts.map(chart=>{
      const categories=shift(chart.categories),series=chart.series.map(s=>({...s,values:shift(s.values),...(s.label?{label:{...s.label,[dimension]:position(s.label[dimension])}}:{})}));
      assert(categories.columnCount===1&&series.every(s=>s.values.columnCount===1&&s.values.rowCount===categories.rowCount),'unsupported','Operation would misalign chart data vectors');
      return {...chart,categories,series,anchor:{...chart.anchor,[dimension]:position(chart.anchor[dimension],true)}};
    });
  }
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const [row, column] = key.split(":").map(Number), value = axis === "rows" ? row : column;
    if (removing && value >= start && value < start + count) continue;
    const next = value >= start ? value + (removing ? -count : count) : value;
    assert(next <= maximum, "unsupported", "Insertion would discard cells at the sheet boundary");
    cells[cellKey(axis === "rows" ? next : row, axis === "columns" ? next : column)] = cell;
  }
  sheet.cells = cells;
  const metadata:Record<string,Dimension> = {};
  for (const [key, value] of Object.entries(sheet[axis])) {
    const at = Number(key);
    if (removing && at >= start && at < start + count) continue;
    const next = at >= start ? at + (removing ? -count : count) : at;
    assert(next <= maximum, "unsupported", "Insertion would discard dimension metadata");
    metadata[String(next)] = value;
  }
  sheet[axis] = metadata;
  sheet.merges = sheet.merges.flatMap((merge) => {
    const bounds = shiftInterval(merge[dimension], merge[dimension] + merge[size] - 1, start, count, removing);
    if (!bounds) return [];
    assert(bounds[1] <= maximum, "unsupported", "Insertion would discard a merged range");
    return [{ ...merge, [dimension]: bounds[0], [size]: bounds[1] - bounds[0] + 1 }];
  }).filter((m) => m.rowCount * m.columnCount > 1);
  if(sheet.conditionalRules){
    sheet.conditionalRules=sheet.conditionalRules.flatMap(rule=>{
      const ranges=rule.ranges.flatMap(r=>{
        const bounds=shiftInterval(r[dimension],r[dimension]+r[size]-1,start,count,removing);
        if(!bounds)return [];
        assert(bounds[1]<=maximum,'unsupported','Insertion would discard conditional range');
        return [{...r,[dimension]:bounds[0],[size]:bounds[1]-bounds[0]+1}];
      });
      return ranges.length?[{...rule,ranges}]:[];
    });
    if(!sheet.conditionalRules.length)delete sheet.conditionalRules;
  }
  if (sheet.view) {
    const field = axis === "rows" ? "frozenRows" : "frozenColumns", value = sheet.view[field];
    if (value !== void 0 && start <= value) sheet.view[field] = removing ? Math.max(start - 1, value - count) : Math.min(maximum - 1, value + count);
  }
  return setSheets(snapshot, sheets);
}
export function insertRows(snapshot:DocumentSnapshot, sheetId:string, start:number, count = 1):DocumentCommand {
  return dimensions2(snapshot, sheetId, "rows", start, count, false);
}
export function deleteRows(snapshot:DocumentSnapshot, sheetId:string, start:number, count = 1):DocumentCommand {
  return dimensions2(snapshot, sheetId, "rows", start, count, true);
}
export function insertColumns(snapshot:DocumentSnapshot, sheetId:string, start:number, count = 1):DocumentCommand {
  return dimensions2(snapshot, sheetId, "columns", start, count, false);
}
export function deleteColumns(snapshot:DocumentSnapshot, sheetId:string, start:number, count = 1):DocumentCommand {
  return dimensions2(snapshot, sheetId, "columns", start, count, true);
}
export function mergeCells(snapshot:DocumentSnapshot, sheetId:string, range:CellRange):DocumentCommand {
  assertCellRange(range);
  assert(range.rowCount * range.columnCount > 1, "invalid", "A merge needs at least two cells");
  const index = sheetIndex(snapshot, sheetId), sheet = book(snapshot).sheets[index];
  assert(!sheet.merges.some((m) => overlap(m, range)), "unsupported", "Merged ranges overlap");
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const [row, column] = key.split(":").map(Number);
    if (inside(range, row, column) && (row !== range.row || column !== range.column)) assert(cell.value.kind === "blank", "unsupported", "Merge would discard a value");
  }
  return edit(snapshot, [{ op: "set", path: ["sheets", index, "merges"], value: [...sheet.merges, range] }]);
}
export function unmergeCells(snapshot:DocumentSnapshot, sheetId:string, range:CellRange):DocumentCommand {
  assertCellRange(range);
  const index = sheetIndex(snapshot, sheetId), sheet = book(snapshot).sheets[index];
  assert(sheet.merges.every((m) => !overlap(m, range) || contains(range, m)), "unsupported", "Select the complete merged range");
  return edit(snapshot, [{ op: "set", path: ["sheets", index, "merges"], value: sheet.merges.filter((m) => !overlap(m, range)) }]);
}
export function shiftFormula(source:string, rowDelta:number, columnDelta:number):string {
  const ast = checkedAst(source);
  const shifted = transform(ast, (r) => {
    const row = r.row + (r.rowAbsolute ? 0 : rowDelta), column = r.column + (r.columnAbsolute ? 0 : columnDelta);
    return row < 1 || row > 1048576 || column < 1 || column > 16384 ? { kind: "error", value: "#REF!" } : { ...r, row, column };
  }, (range) => {
    const shift = (r:FormulaReference):FormulaReference => ({ ...r, row: r.row + (r.rowAbsolute || range.axis === "columns" ? 0 : rowDelta), column: r.column + (r.columnAbsolute || range.axis === "rows" ? 0 : columnDelta) });
    const a = shift(range.start), b = shift(range.end);
    if ([a, b].some((r) => r.row < 1 || r.row > 1048576 || r.column < 1 || r.column > 16384)) return { kind: "error", value: "#REF!" };
    return { ...range, start: a, end: b };
  });
  return serializeFormula(shifted);
}
export function copyCells(snapshot:DocumentSnapshot, sourceSheetId:string, range:CellRange, targetSheetId:string, row:number, column:number):DocumentCommand {
  assertCellRange(range, 1e5);
  assertCellRange({ row, column, rowCount: range.rowCount, columnCount: range.columnCount });
  const source = book(snapshot).sheets[sheetIndex(snapshot, sourceSheetId)], targetIndex = sheetIndex(snapshot, targetSheetId), target = book(snapshot).sheets[targetIndex], destination = { ...range, row, column };
  assert(source.merges.every((m) => !overlap(m, range) || contains(range, m)), "unsupported", "Select entire merged cells to copy");
  assert(target.merges.every((m) => !overlap(m, destination) || contains(destination, m)), "unsupported", "Paste intersects part of a merged cell");
  const cells = { ...target.cells };
  for (const key of Object.keys(cells)) {
    const [r, c] = key.split(":").map(Number);
    if (inside(destination, r, c)) delete cells[key];
  }
  for (const [key, cell] of Object.entries(source.cells)) {
    const [r, c] = key.split(":").map(Number);
    if (!inside(range, r, c)) continue;
    cells[cellKey(r + row - range.row, c + column - range.column)] = cell.value.kind === "formula" ? { ...cell, value: { kind: "formula", source: shiftFormula(cell.value.source, row - range.row, column - range.column) } } : cell;
  }
  const merges = [...target.merges.filter((m) => !overlap(m, destination)), ...source.merges.filter((m) => contains(range, m)).map((m) => ({ ...m, row: m.row + row - range.row, column: m.column + column - range.column }))];
  return edit(snapshot, [{ op: "set", path: ["sheets", targetIndex, "cells"], value: cells }, { op: "set", path: ["sheets", targetIndex, "merges"], value: merges }]);
}
export function clearCells(snapshot:DocumentSnapshot, sheetId:string, range:CellRange):DocumentCommand {
  assertCellRange(range);
  const index = sheetIndex(snapshot, sheetId), cells = { ...book(snapshot).sheets[index].cells };
  for (const [key, cell] of Object.entries(cells)) {
    const [row, column] = key.split(":").map(Number);
    if (!inside(range, row, column)) continue;
    const {richText,...rest}=cell;
    if (Object.keys(rest).length === 1) delete cells[key];
    else cells[key] = { ...rest, value: { kind: "blank" } };
  }
  return cellsCommand(snapshot, index, cells);
}
export function formatCells(snapshot:DocumentSnapshot, sheetId:string, range:CellRange, format:{style?:CellStyle;numberFormat?:string}):DocumentCommand {
  assertCellRange(range, 1e5);
  const index = sheetIndex(snapshot, sheetId), sheet = book(snapshot).sheets[index], cells = { ...sheet.cells };
  for (let r = range.row; r < range.row + range.rowCount; r++) for (let c = range.column; c < range.column + range.columnCount; c++) {
    const key = cellKey(r, c), before = cells[key] ?? { value: { kind: "blank" } };
    const style=format.style!==undefined?{...before.style,...format.style}:undefined;
    if(style&&format.style?.border!==undefined)delete style.borderEdges;
    if(style&&format.style?.borderEdges!==undefined)delete style.border;
    const font=Object.fromEntries(Object.entries(format.style??{}).filter(([key])=>['fontFamily','fontSize','bold','italic','underline','color'].includes(key)));
    const richText=before.richText&&Object.keys(font).length?before.richText.map(r=>({...r,style:{...r.style,...font}})):before.richText;
    cells[key] = { ...before, ...(style?{style}:{}), ...(richText?{richText}:{}), ...format.numberFormat !== void 0 ? { numberFormat: format.numberFormat } : {} };
  }
  return cellsCommand(snapshot, index, cells);
}
export function setDimensions(snapshot:DocumentSnapshot, sheetId:string, axis:Axis, start:number, count:number, dimension:Dimension):DocumentCommand {
  const index = sheetIndex(snapshot, sheetId), maximum = axis === "rows" ? 1048576 : 16384;
  assert(Number.isSafeInteger(start) && start >= 1 && Number.isSafeInteger(count) && count > 0 && count <= 1e5 && start + count - 1 <= maximum, "invalid", "Invalid dimension selection");
  assert(axis === "rows" ? dimension.width === void 0 : dimension.height === void 0, "invalid", "Wrong dimension property");
  const values = { ...book(snapshot).sheets[index][axis] };
  for (let i = start; i < start + count; i++) values[String(i)] = { ...values[String(i)], ...dimension };
  return edit(snapshot, [{ op: "set", path: ["sheets", index, axis], value: values }]);
}
export function fillCells(snapshot:DocumentSnapshot, sheetId:string, range:CellRange, direction:"down"|"right"):DocumentCommand {
  assertCellRange(range, 1e5);
  const index = sheetIndex(snapshot, sheetId), sheet = book(snapshot).sheets[index];
  assert(!sheet.merges.some((m) => overlap(m, range)), "unsupported", "Unmerge the fill range first");
  const cells = { ...sheet.cells };
  for (let r = range.row; r < range.row + range.rowCount; r++) for (let c = range.column; c < range.column + range.columnCount; c++) {
    const sr = direction === "down" ? range.row : r, sc = direction === "right" ? range.column : c, source = sheet.cells[cellKey(sr, sc)], key = cellKey(r, c);
    if (!source) delete cells[key];
    else cells[key] = source.value.kind === "formula" ? { ...source, value: { kind: "formula", source: shiftFormula(source.value.source, r - sr, c - sc) } } : source;
  }
  return cellsCommand(snapshot, index, cells);
}
export function parseTsv(text2:string):string[][] {
  assert(text2.length <= 8 * 1024 * 1024, "limit", "Clipboard text budget exceeded");
  const rows:string[][] = [];
  let row:string[] = [], field = "", quoted = false, closed = false, count = 0;
  const nextField = () => {
    assert(++count <= 1e5 && field.length <= 32767, "limit", "Clipboard cell budget exceeded");
    row.push(field);
    field = "";
    closed = false;
  };
  const nextRow = () => {
    nextField();
    rows.push(row);
    row = [];
  };
  for (let i = 0; i < text2.length; i++) {
    const ch = text2[i];
    if (quoted) {
      if (ch === '"') {
        if (text2[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          quoted = false;
          closed = true;
        }
      } else field += ch;
      continue;
    }
    if (ch === '"' && field === "" && !closed) {
      quoted = true;
      continue;
    }
    if (ch === "	") {
      nextField();
      continue;
    }
    if (ch === "\r" || ch === "\n") {
      if (ch === "\r" && text2[i + 1] === "\n") i++;
      nextRow();
      continue;
    }
    assert(!closed, "invalid", "Unexpected text after quoted clipboard cell");
    field += ch;
  }
  assert(!quoted, "invalid", "Unterminated clipboard field");
  if (field.length || row.length || !rows.length || !/[\r\n]$/.test(text2)) nextRow();
  const width = rows.reduce((max, row2) => Math.max(max, row2.length), 0);
  assert(width * rows.length <= 1e5, "limit", "Rectangular clipboard cell budget exceeded");
  for (const row2 of rows) while (row2.length < width) row2.push("");
  return rows;
}
export function pasteTsv(snapshot:DocumentSnapshot, sheetId:string, row:number, column:number, text2:string, options:CellInputOptions = {}):DocumentCommand {
  const matrix = parseTsv(text2);
  assertCellRange({ row, column, rowCount: matrix.length, columnCount: matrix[0].length }, 1e5);
  const sheet = book(snapshot).sheets[sheetIndex(snapshot, sheetId)], edits = [];
  for (let r = 0; r < matrix.length; r++) for (let c = 0; c < matrix[r].length; c++) {
    const text3 = matrix[r][c], key = cellKey(row + r, column + c), previous = sheet.cells[key];
    edits.push({ row: row + r, column: column + c, cell: cellFromInput(text3, options, previous) });
  }
  return setCells(snapshot, sheetId, edits);
}
export function exportTsv(document2:WorkbookSnapshot, sheetId:string, range:CellRange, options:{formulas?:boolean;calculation?:CalculationResult} = {}):string {
  assertCellRange(range, 1e5);
  const sheet = document2.sheets.find((s) => s.id === sheetId);
  assert(sheet, "invalid", "Sheet does not exist");
  const rows = [];
  const quote = (text2:string) => /[\t\r\n"]/.test(text2) ? `"${text2.replaceAll('"', '""')}"` : text2;
  for (let r = range.row; r < range.row + range.rowCount; r++) {
    const cells = [];
    for (let c = range.column; c < range.column + range.columnCount; c++) {
      const cell = sheet.cells[cellKey(r, c)];
      const text2 = !cell ? "" : cell.value.kind === "formula" && options.formulas ? cell.value.source : formatCell(cell, options.calculation?.get(sheetId, r, c), { dateSystem: document2.dateSystem });
      cells.push(quote(text2));
    }
    rows.push(cells.join("	"));
  }
  return rows.join("\r\n");
}

export function setSheetView(snapshot:DocumentSnapshot,sheetId:string,view:SheetView):DocumentCommand{
  const index=sheetIndex(snapshot,sheetId), sheet=book(snapshot).sheets[index];
  return edit(snapshot,[{op:'set',path:['sheets',index,'view'],value:{...sheet.view,...view}}]);
}

/** A rule is an address-bound worksheet property, not direct cell formatting. */
export function addConditionalRule(snapshot:DocumentSnapshot,sheetId:string,rule:Omit<SheetConditionalRule,'priority'>):DocumentCommand{
 const index=sheetIndex(snapshot,sheetId),rules=book(snapshot).sheets[index]!.conditionalRules??[];
 const priority=rules.reduce((n,r)=>Math.max(n,r.priority),0)+1;
 return edit(snapshot,[{op:'set',path:['sheets',index,'conditionalRules'],value:[...rules,{...rule,priority}]}]);
}
export function removeConditionalRule(snapshot:DocumentSnapshot,sheetId:string,priority:number):DocumentCommand{
 const index=sheetIndex(snapshot,sheetId),rules=book(snapshot).sheets[index]!.conditionalRules??[],remaining=rules.filter(r=>r.priority!==priority);
 assert(remaining.length<rules.length,'invalid','Conditional rule not found');
 return edit(snapshot,[remaining.length?{op:'set',path:['sheets',index,'conditionalRules'],value:remaining}:{op:'delete',path:['sheets',index,'conditionalRules']}]);
}

/** Create a standard chart from one category column and adjacent value
 * columns. First row is the header; the remaining rows are data. */
export function addSheetChart(snapshot:DocumentSnapshot,sheetId:string,range:CellRange,kind:SheetChart['kind']):DocumentCommand{
 const index=sheetIndex(snapshot,sheetId),sheet=book(snapshot).sheets[index]!;
 assertCellRange(range,170000);assert(range.rowCount>=2&&range.rowCount<=10001&&range.columnCount>=2&&range.columnCount<=(kind==='pie'?2:17),'invalid','Select headers, one category column and value columns');
 const categories={row:range.row+1,column:range.column,rowCount:range.rowCount-1,columnCount:1};
 const series=Array.from({length:range.columnCount-1},(_,i)=>({label:{row:range.row,column:range.column+i+1},values:{...categories,column:range.column+i+1}}));
 const column=Math.min(16384,range.column+range.columnCount+1),height=2700000,width=5400000,rowHeight=sheet.defaultGeometry?.rowHeight??24*9525;
 let row=range.row;
 // Place successive charts below overlapping chart anchors using the
 // existing default geometry approximation. Never stack at one point.
 for(const existing of sheet.charts??[]){const a=existing.anchor,columns=Math.ceil(a.width/(sheet.defaultGeometry?.columnWidth??104*9525)),currentColumns=Math.ceil(width/(sheet.defaultGeometry?.columnWidth??104*9525));
  if(column<a.column+columns&&a.column<column+currentColumns)row=Math.max(row,a.row+Math.ceil((a.offsetY+a.height)/rowHeight)+1);
 }
 assert(row<=1048576,'unsupported','No room for a new chart anchor');
 const chart:SheetChart={kind,categories,series,anchor:{row,column,offsetX:0,offsetY:0,width,height}};
 return edit(snapshot,[{op:'set',path:['sheets',index,'charts'],value:[...sheet.charts??[],chart]}]);
}
export function removeSheetChart(snapshot:DocumentSnapshot,sheetId:string,chartIndex:number):DocumentCommand{
 const index=sheetIndex(snapshot,sheetId),charts=book(snapshot).sheets[index]!.charts??[];
 assert(Number.isSafeInteger(chartIndex)&&chartIndex>=0&&chartIndex<charts.length,'invalid','Chart not found');
 const next=charts.filter((_,i)=>i!==chartIndex);
 return edit(snapshot,[next.length?{op:'set',path:['sheets',index,'charts'],value:next}:{op:'delete',path:['sheets',index,'charts']}]);
}

/** Apply a source-defined named style. Its current effective properties
 * are materialized in the cell while the native style association remains. */
export function applyNamedCellStyle(snapshot:DocumentSnapshot,sheetId:string,range:CellRange,name:string):DocumentCommand{
 assertCellRange(range,100000);
 const d=book(snapshot),index=sheetIndex(snapshot,sheetId),sheet=d.sheets[index]!,definition=d.namedCellStyles?.find(s=>s.name===name);
 assert(name==='Normal'||definition,'invalid','Unknown named cell style');
 const cells={...sheet.cells};
 for(let row=range.row;row<range.row+range.rowCount;row++)for(let column=range.column;column<range.column+range.columnCount;column++){
  const key=cellKey(row,column),prior=cells[key]??{value:{kind:'blank' as const}};
  // Named styles are cell-level styles; retain explicit inline rich runs.
  const {style,styleId,namedStyle,numberFormat,...rest}=prior;
  cells[key]={...rest,...(definition?{namedStyle:definition.name,...(definition.style?{style:definition.style}:{}),...(definition.numberFormat?{numberFormat:definition.numberFormat}:{})}:{})};
 }
 return cellsCommand(snapshot,index,cells);
}
