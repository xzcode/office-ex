/**
 * SHT-03 / SHT-04: reconstructed from the archived M3 recalculation worker.
 * Input must be an immutable workbook snapshot. All results are derived state,
 * never source cell values. No third-party spreadsheet engine or runtime eval.
 */
import { assert, assertCellAddress, OfficeError, type Cancellation } from '@office-ex/contracts';
import type { CellValue, DocumentSnapshot } from '@office-ex/core';
import { parseFormula, walkFormula, isSupportedFunction, type FormulaNode, type FormulaReference, type FormulaRange, type FormulaError } from './formula.js';

export type WorkbookSnapshot = Extract<DocumentSnapshot['document'], {kind:'sheets'}>;
type SheetSnapshot = WorkbookSnapshot['sheets'][number];
export type CalculatedValue = Exclude<CellValue, {kind:'formula'}>;
type ErrorValue = Extract<CalculatedValue, {kind:'error'}>;
type BooleanValue = Extract<CalculatedValue, {kind:'boolean'}>;
type NumberValue = Extract<CalculatedValue, {kind:'number'}>;
export type CalculationState = 'valid' | 'error' | 'unsupported' | 'stale';
export type CalculationReason = 'dependency' | 'unsupported-syntax' | 'unknown-function' | 'cycle';
export interface CalculationCell {
  readonly state: CalculationState;
  readonly value?: CalculatedValue;
  readonly reason?: CalculationReason;
}
type MutableResult = { -readonly [K in keyof CalculationCell]: CalculationCell[K] };
export interface CalculationOptions {
  maxSteps?: number;
  maxDepth?: number;
  cancellation?: Cancellation;
  /** Supply a monotonic clock together with the deadline. */
  now?: () => number;
  deadline?: number;
  cachedValues?: ReadonlyMap<string, CalculatedValue>;
  cachedSources?: ReadonlyMap<string, string>;
}
export interface CalculationResult {
  readonly cells: ReadonlyMap<string, CalculationCell>;
  readonly computedCells: number;
  readonly steps: number;
  get(sheetId: string, row: number, column: number): CalculationCell;
}
interface IndexedCell { key: string; row: number; column: number }
interface Rectangle { sheetId: string; row1: number; row2: number; column1: number; column2: number; owner: string }
interface RangeTree { center: number; crossing: Rectangle[]; left?: RangeTree; right?: RangeTree }
interface PendingValue { kind: 'pending'; reason: CalculationReason; state: 'stale' | 'unsupported' }
type EvaluationValue = CalculatedValue | PendingValue | {kind:'range-value'; range:Rectangle};
interface ParsedCell { source: string; ast?: FormulaNode; unsupported?: boolean }
interface FormulaEntry { sheetId: string; row: number; column: number; parsed: ParsedCell }
interface SheetIndexState { keys: string[]; formulaKeys: string[]; all: SparseSheetIndex; formulas: SparseSheetIndex }
type Dependencies = Map<string, Set<string>>;

export function calculationKey(sheetId: string, row: number, column: number): string {
  return `${sheetId}!${row}:${column}`;
}
class SparseSheetIndex {
  readonly rows: number[];
  #byRow = new Map<number, IndexedCell[]>();
  constructor(sheet: SheetSnapshot, tick: () => void = () => {}) {
    for (const key of Object.keys(sheet.cells)) {
      tick();
      const [row, column] = key.split(":").map(Number);
      const entries = this.#byRow.get(row) ?? [];
      entries.push({ key, row, column });
      this.#byRow.set(row, entries);
    }
    this.rows = [...this.#byRow.keys()].sort((a, b) => a - b);
    for (const entries of this.#byRow.values()) entries.sort((a, b) => a.column - b.column);
  }
  *range(firstRow: number, firstColumn: number, lastRow: number, lastColumn: number): IterableIterator<IndexedCell> {
    for (let r = lowerBound(this.rows, firstRow); r < this.rows.length && this.rows[r] <= lastRow; r++) {
      const entries = this.#byRow.get(this.rows[r])!;
      let lo = 0, hi = entries.length;
      while (lo < hi) {
        const mid = lo + hi >>> 1;
        if (entries[mid].column < firstColumn) lo = mid + 1;
        else hi = mid;
      }
      for (let c = lo; c < entries.length && entries[c].column <= lastColumn; c++) yield entries[c];
    }
  }
}
function lowerBound(values: readonly number[], target: number): number {
  let lo = 0, hi = values.length;
  while (lo < hi) {
    const mid = lo + hi >>> 1;
    if (values[mid] < target) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}
class RangeDependencyIndex {
  #sheets = new Map<string, RangeTree | undefined>();
  constructor(ranges: readonly Rectangle[], tick: () => void = () => {}) {
    const groups = new Map<string, Rectangle[]>();
    for (const range of ranges) {
      tick();
      const group = groups.get(range.sheetId) ?? [];
      group.push(range);
      groups.set(range.sheetId, group);
    }
    const build = (items: Rectangle[]): RangeTree | undefined => {
      if (!items.length) return void 0;
      const ends = items.flatMap((i) => [i.row1, i.row2]).sort((a, b) => a - b), center = ends[Math.floor(ends.length / 2)];
      const left: Rectangle[] = [], right: Rectangle[] = [], crossing: Rectangle[] = [];
      for (const item of items) {
        tick();
        if (item.row2 < center) left.push(item);
        else if (item.row1 > center) right.push(item);
        else crossing.push(item);
      }
      return { center, crossing, ...left.length ? { left: build(left) } : {}, ...right.length ? { right: build(right) } : {} };
    };
    for (const [id, ranges2] of groups) this.#sheets.set(id, build(ranges2));
  }
  query(sheetId: string, row: number, column: number, tick?: () => void): string[] {
    let node = this.#sheets.get(sheetId);
    const owners: string[] = [];
    while (node) {
      for (const range of node.crossing) {
        tick?.();
        if (row >= range.row1 && row <= range.row2 && column >= range.column1 && column <= range.column2) owners.push(range.owner);
      }
      node = row < node.center ? node.left : row > node.center ? node.right : void 0;
    }
    return owners;
  }
}
const blank: CalculatedValue = Object.freeze({ kind: "blank" });
const num = (value: number): NumberValue | ErrorValue => Number.isFinite(value) ? { kind: "number", value: Object.is(value, -0) ? 0 : value } : error("#NUM!");
const error = (value: FormulaError): ErrorValue => ({ kind: "error", value });
const bool = (value: boolean): BooleanValue => ({ kind: "boolean", value });
const pending = (reason: CalculationReason = "dependency", state: PendingValue["state"] = "stale"): PendingValue => ({ kind: "pending", reason, state });
function scalar(value: EvaluationValue): CalculatedValue | PendingValue {
  return value.kind === "range-value" ? pending("unsupported-syntax", "unsupported") : value;
}
function numeric(value: CalculatedValue): number | ErrorValue {
  if (value.kind === "number") return value.value;
  if (value.kind === "blank") return 0;
  if (value.kind === "boolean") return Number(value.value);
  if (value.kind === "error") return value;
  const text = value.value.trim();
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text)) return error("#VALUE!");
  const number = Number(text);
  return Number.isFinite(number) ? number : error("#NUM!");
}
function logical(value: CalculatedValue): boolean | ErrorValue {
  if (value.kind === "boolean") return value.value;
  if (value.kind === "blank") return false;
  if (value.kind === "number") return value.value !== 0;
  if (value.kind === "error") return value;
  if (value.value.toUpperCase() === "TRUE") return true;
  if (value.value.toUpperCase() === "FALSE") return false;
  return error("#VALUE!");
}
function textValue(value: CalculatedValue): string {
  return value.kind === "blank" ? "" : value.kind === "boolean" ? value.value ? "TRUE" : "FALSE" : String(value.value);
}
function comparison(a: CalculatedValue, b: CalculatedValue): number {
  if (a.kind === "blank") a = b.kind === "string" ? { kind: "string", value: "" } : b.kind === "boolean" ? bool(false) : num(0);
  if (b.kind === "blank") b = a.kind === "string" ? { kind: "string", value: "" } : a.kind === "boolean" ? bool(false) : num(0);
  const rank = (v: CalculatedValue) => v.kind === "number" ? 0 : v.kind === "string" ? 1 : 2;
  if (a.kind !== b.kind) return rank(a) - rank(b);
  const av = a.kind === "string" ? a.value.toUpperCase() : "value" in a ? a.value : 0;
  const bv = b.kind === "string" ? b.value.toUpperCase() : "value" in b ? b.value : 0;
  return av === bv ? 0 : av < bv ? -1 : 1;
}
function decimalRound(value: number, digits: number, mode:'nearest'|'up'|'down'='nearest'): NumberValue | ErrorValue {
  const precision = Math.trunc(digits);
  if (precision > 308) return num(value);
  if (precision < -308) return mode==='up'&&value!==0?error('#NUM!'):num(0);
  const shift = (n: number, exponent: number): number => {
    const [base, exp = "0"] = String(n).split("e");
    return Number(`${base}e${Number(exp) + exponent}`);
  };
  const shifted = shift(Math.abs(value), precision);
  if (!Number.isFinite(shifted)) return num(value);
  const rounded=mode==='up'?Math.ceil(shifted)||Number(value!==0):mode==='down'?Math.floor(shifted):Math.floor(shifted+0.5);
  return num(Math.sign(value) * shift(rounded, -precision));
}
function sameValue(a: CellValue | undefined, b: CellValue | undefined): boolean {
  if (a === b) return true;
  if (!a || !b || a.kind !== b.kind) return false;
  return a.kind === "blank" || (a.kind === "formula" && b.kind === "formula" ? a.source === b.source : "value" in a && "value" in b && a.value === b.value);
}
export class WorkbookCalculator {
  #previous: WorkbookSnapshot | undefined;
  #results = new Map<string, CalculationCell>();
  #parsed = new Map<string, ParsedCell>();
  #index = new RangeDependencyIndex([]);
  #sheetIndexes = new Map<string, SheetIndexState>();
  #dependencies: Dependencies = new Map();
  #reverse: Dependencies = new Map();
  calculate(document: WorkbookSnapshot, options: CalculationOptions = {}): CalculationResult {
    assert(document.kind === "sheets", "invalid", "Expected workbook");
    let steps = 0, computedCells = 0;
    const maximum = options.maxSteps ?? 1e6, maxDepth = options.maxDepth ?? 256;
    assert(Number.isSafeInteger(maximum) && maximum > 0 && maximum <= 1e8, "invalid", "Invalid calculation budget");
    assert(Number.isInteger(maxDepth) && maxDepth >= 1 && maxDepth <= 512, "invalid", "Invalid calculation depth budget");
    assert(options.deadline === undefined || (Number.isFinite(options.deadline) && typeof options.now === "function"), "invalid", "A finite deadline requires a monotonic clock");
    const tick = () => {
      assert(++steps <= maximum, "limit", "Formula calculation step budget exceeded");
      if ((steps & 127) === 1) {
        options.cancellation?.throwIfCancelled();
        if (options.deadline !== void 0 && options.now) assert(options.now() <= options.deadline, "limit", "Calculation deadline exceeded");
      }
    };
    tick();
    const sheets = new Map(document.sheets.map((sheet) => [sheet.id, sheet]));
    const names = new Map(document.sheets.map((sheet) => [sheet.name.toLowerCase(), sheet.id]));
    const structureChanged = !this.#previous || this.#previous.documentId !== document.documentId || this.#previous.sheets.length !== document.sheets.length || this.#previous.sheets.some((s, i) => s.id !== document.sheets[i]?.id || s.name !== document.sheets[i]?.name);
    const changed = new Set<string>();
    const formulaEntries = new Map<string, FormulaEntry>();
    const parsed = new Map<string, ParsedCell>(), sheetIndexes = new Map<string, SheetIndexState>();
    const indexes = new Map<string, SparseSheetIndex>(), formulaIndexes = new Map<string, SparseSheetIndex>();
    let formulasChanged = structureChanged;
    for (const sheet of document.sheets) {
      const before = this.#previous?.sheets.find((s) => s.id === sheet.id), prior = this.#sheetIndexes.get(sheet.id);
      let state: SheetIndexState;
      if (prior && before?.cells === sheet.cells) state = prior;
      else {
        const keys = Object.keys(sheet.cells), formulaKeys: string[] = [];
        let coordinatesChanged = !prior || keys.length !== prior.keys.length, formulaCoordinatesChanged = !prior;
        for (let index = 0; index < keys.length; index++) {
          if ((index & 1023) === 0) tick();
          const key = keys[index], cell = sheet.cells[key], old = before?.cells[key];
          if (!old) coordinatesChanged = true;
          if (!sameValue(old?.value, cell.value)) changed.add(`${sheet.id}!${key}`);
          if (cell.value.kind === "formula") formulaKeys.push(key);
          if (cell.value.kind === "formula" !== (old?.value.kind === "formula")) formulaCoordinatesChanged = true;
        }
        if (before && coordinatesChanged) for (const key of prior?.keys ?? Object.keys(before.cells)) {
          tick();
          if (Object.hasOwn(sheet.cells, key)) continue;
          changed.add(`${sheet.id}!${key}`);
          if (before.cells[key].value.kind === "formula") formulaCoordinatesChanged = true;
        }
        const all = !coordinatesChanged && prior ? prior.all : new SparseSheetIndex(sheet, tick);
        const formulaCells: Record<string, SheetSnapshot["cells"][string]> = {};
        if (formulaCoordinatesChanged) for (const key of formulaKeys) formulaCells[key] = sheet.cells[key];
        const formulas = !formulaCoordinatesChanged && prior ? prior.formulas : new SparseSheetIndex({ ...sheet, cells: formulaCells }, tick);
        state = { keys, formulaKeys, all, formulas };
      }
      sheetIndexes.set(sheet.id, state);
      indexes.set(sheet.id, state.all);
      formulaIndexes.set(sheet.id, state.formulas);
      for (const cellKey of state.formulaKeys) {
        const value = sheet.cells[cellKey].value;
        assert(value.kind === "formula", "invalid", "Stale formula index");
        tick();
        const key = `${sheet.id}!${cellKey}`;
        let item = this.#parsed.get(key);
        if (!item || item.source !== value.source) {
          formulasChanged = true;
          item = { source: value.source };
          try {
            item.ast = parseFormula(value.source, { cancellation: options.cancellation });
          } catch (failure) {
            if (!(failure instanceof OfficeError) || !["unsupported", "invalid"].includes(failure.code)) throw failure;
            item.unsupported = true;
          }
        }
        parsed.set(key, item);
        const [row, column] = cellKey.split(":").map(Number);
        formulaEntries.set(key, { sheetId: sheet.id, row, column, parsed: item });
      }
    }
    if (parsed.size !== this.#parsed.size) formulasChanged = true;
    const resolve = (reference: FormulaReference, current: string): string | undefined => reference.sheet === void 0 ? current : names.get(reference.sheet.toLowerCase());
    const rectangle = (node: FormulaReference | FormulaRange, current: string, owner: string): Rectangle | undefined => {
      const a = node.kind === "range" ? node.start : node, b = node.kind === "range" ? node.end : node;
      const sheetId = resolve(a, current), endId = resolve(b, sheetId ?? current);
      if (!sheetId || endId !== sheetId) return void 0;
      return { sheetId, row1: Math.min(a.row, b.row), column1: Math.min(a.column, b.column), row2: Math.max(a.row, b.row), column2: Math.max(a.column, b.column), owner };
    };
    const ranges: Rectangle[] = [];
    const deps: Dependencies = formulasChanged ? new Map() : this.#dependencies;
    const reverse: Dependencies = formulasChanged ? new Map() : this.#reverse;
    if (formulasChanged) for (const [key, entry] of formulaEntries) {
      const dependencies = new Set<string>();
      deps.set(key, dependencies);
      if (entry.parsed.ast) walkFormula(entry.parsed.ast, (node) => {
        if (node.kind !== "reference" && node.kind !== "range") return;
        tick();
        const range = rectangle(node, entry.sheetId, key);
        if (!range) return;
        ranges.push(range);
        for (const item of formulaIndexes.get(range.sheetId)!.range(range.row1, range.column1, range.row2, range.column2)) {
          tick();
          const dependency = `${range.sheetId}!${item.key}`;
          dependencies.add(dependency);
          const dependents = reverse.get(dependency) ?? new Set<string>();
          dependents.add(key);
          reverse.set(dependency, dependents);
        }
      });
    }
    const rangeIndex = formulasChanged ? new RangeDependencyIndex(ranges, tick) : this.#index, dirty = new Set<string>();
    if (structureChanged) for (const key of formulaEntries.keys()) dirty.add(key);
    else {
      for (const key of changed) {
        if (formulaEntries.has(key)) dirty.add(key);
        const [sheetId, coordinate] = key.split("!"), [row, col] = coordinate.split(":").map(Number);
        for (const owner of this.#index === rangeIndex ? rangeIndex.query(sheetId, row, col, tick) : [...this.#index.query(sheetId, row, col, tick), ...rangeIndex.query(sheetId, row, col, tick)]) if (formulaEntries.has(owner)) dirty.add(owner);
      }
      for (const [key, entry] of formulaEntries) {
        const prior = this.#results.get(key);
        if (!prior) dirty.add(key);
        else if (prior.state === "unsupported" || prior.state === "stale") {
          const supplied = options.cachedSources?.get(key) === entry.parsed.source || !options.cachedSources ? options.cachedValues?.get(key) : undefined;
          if (supplied !== undefined && !sameValue(supplied, prior.value)) dirty.add(key);
        }
      }
    }
    const queue = [...dirty];
    for (let head = 0; head < queue.length; head++) for (const owner of reverse.get(queue[head]) ?? []) {
      tick();
      if (!dirty.has(owner)) { dirty.add(owner); queue.push(owner); }
    }
    const results = new Map<string, CalculationCell>();
    for (const key of formulaEntries.keys()) if (!dirty.has(key) && this.#results.has(key)) results.set(key, this.#results.get(key)!);
    const active = new Set<string>();
    const stack: string[] = [];
    const cycles = new Set<string>();
    const fromResult = (result: CalculationCell): CalculatedValue | PendingValue => result.state === "valid" || result.state === "error" ? result.value ?? blank : pending();
    const getCell = (sheetId: string, row: number, column: number, depth: number): CalculatedValue | PendingValue => {
      tick();
      const cell = sheets.get(sheetId)?.cells[`${row}:${column}`];
      if (!cell) return blank;
      if (cell.value.kind !== "formula") return cell.value;
      return fromResult(calculateCell(calculationKey(sheetId, row, column), depth + 1));
    };
    const evaluate = (node: FormulaNode, sheetId: string, owner: string, depth: number): EvaluationValue => {
      tick();
      assert(depth <= maxDepth, "limit", "Formula evaluation depth exceeded");
      if (node.kind === "number" || node.kind === "string" || node.kind === "boolean" || node.kind === "error") return node;
      if (node.kind === "missing") return num(0);
      if (node.kind === "reference") {
        const target = resolve(node, sheetId);
        return target ? getCell(target, node.row, node.column, depth) : error("#REF!");
      }
      if (node.kind === "range") {
        const range = rectangle(node, sheetId, owner);
        return range ? { kind: "range-value", range } : error("#REF!");
      }
      if (node.kind === "unary") {
        const value = scalar(evaluate(node.operand, sheetId, owner, depth + 1));
        if (value.kind === "pending") return value;
        const number = numeric(value);
        if (typeof number !== "number") return number;
        return num(node.operator === "-" ? -number : node.operator === "%" ? number / 100 : number);
      }
      if (node.kind === "binary") {
        const left = scalar(evaluate(node.left, sheetId, owner, depth + 1));
        if (left.kind === "pending" || left.kind === "error") return left;
        const right = scalar(evaluate(node.right, sheetId, owner, depth + 1));
        if (right.kind === "pending" || right.kind === "error") return right;
        if (node.operator === "&") {
          const text = textValue(left) + textValue(right);
          return text.length > 32767 ? error("#VALUE!") : { kind: "string", value: text };
        }
        if (["=", "<>", "<", ">", "<=", ">="].includes(node.operator)) {
          const c = comparison(left, right);
          return bool(node.operator === "=" ? c === 0 : node.operator === "<>" ? c !== 0 : node.operator === "<" ? c < 0 : node.operator === ">" ? c > 0 : node.operator === "<=" ? c <= 0 : c >= 0);
        }
        const a = numeric(left), b = numeric(right);
        if (typeof a !== "number") return a;
        if (typeof b !== "number") return b;
        if (node.operator === "/" && b === 0) return error("#DIV/0!");
        if (node.operator === "^" && a === 0 && b < 0) return error("#DIV/0!");
        if (node.operator === "^" && a === 0 && b === 0) return error("#NUM!");
        return num(node.operator === "+" ? a + b : node.operator === "-" ? a - b : node.operator === "*" ? a * b : node.operator === "/" ? a / b : a ** b);
      }
      const name = node.name.toUpperCase();
      if (!isSupportedFunction(name)) return pending("unknown-function", "unsupported");
      if (node.args.length === 0 || node.args.length > 255) return error("#VALUE!");
      if (name === "IF") {
        if (node.args.length < 2 || node.args.length > 3) return error("#VALUE!");
        const condition = scalar(evaluate(node.args[0], sheetId, owner, depth + 1));
        if (condition.kind === "pending") return condition;
        const logicalValue = logical(condition);
        if (typeof logicalValue !== "boolean") return logicalValue;
        const chosen = logicalValue ? node.args[1] : node.args[2];
        return chosen ? evaluate(chosen, sheetId, owner, depth + 1) : bool(false);
      }
      // Conditional error handling must not turn unsupported/stale source
      // formulas into valid values. Only an actual calculated error is caught.
      if(name==='IFERROR'){
        if(node.args.length!==2)return error('#VALUE!');
        const value=scalar(evaluate(node.args[0],sheetId,owner,depth+1));
        if(value.kind==='pending')return value;
        if(value.kind==='error')return scalar(evaluate(node.args[1],sheetId,owner,depth+1));
        return value;
      }
      if(['ISNUMBER','ISTEXT','ISLOGICAL','ISBLANK','ISERROR','ISERR','ISNA'].includes(name)){
        if(node.args.length!==1)return error('#VALUE!');
        const value=scalar(evaluate(node.args[0],sheetId,owner,depth+1));
        if(value.kind==='pending')return value;
        return bool(name==='ISNUMBER'?value.kind==='number':name==='ISTEXT'?value.kind==='string':name==='ISLOGICAL'?value.kind==='boolean':name==='ISBLANK'?value.kind==='blank':name==='ISERROR'?value.kind==='error':name==='ISNA'?value.kind==='error'&&value.value==='#N/A':value.kind==='error'&&value.value!=='#N/A');
      }
      if(name==='COUNTBLANK'){
        if(node.args.length!==1||!['reference','range'].includes(node.args[0]!.kind))return error('#VALUE!');
        const value=evaluate(node.args[0],sheetId,owner,depth+1);
        if(value.kind==='pending')return value;
        if(value.kind!=='range-value'){
          const arg=node.args[0]!;
          if(arg.kind==='range'||arg.kind==='reference'&&!resolve(arg,sheetId))return error('#REF!');
          return num(value.kind==='blank'||value.kind==='string'&&value.value===''?1:0);
        }
        const r=value.range;
        let count=(r.row2-r.row1+1)*(r.column2-r.column1+1);
        // Count implicit empty cells arithmetically; even an entire worksheet
        // is sparse and never expanded into millions of blank allocations.
        for(const entry of indexes.get(r.sheetId)!.range(r.row1,r.column1,r.row2,r.column2)){
          tick();const v=getCell(r.sheetId,entry.row,entry.column,depth+1);
          if(v.kind==='pending')return v;
          if(v.kind!=='blank'&&!(v.kind==='string'&&v.value===''))count--;
        }
        return num(count);
      }
      if(['INT','MOD','ROUNDUP','ROUNDDOWN','TRUNC','SIGN','SQRT','POWER'].includes(name)){
        const two=['MOD','ROUNDUP','ROUNDDOWN','POWER'].includes(name);
        if(name==='TRUNC'?node.args.length<1||node.args.length>2:node.args.length!==(two?2:1))return error('#VALUE!');
        const first=scalar(evaluate(node.args[0],sheetId,owner,depth+1));
        if(first.kind==='pending')return first;const a=numeric(first);if(typeof a!=='number')return a;
        if(name==='INT')return num(Math.floor(a));
        if(name==='SIGN')return num(Math.sign(a));
        if(name==='SQRT')return a<0?error('#NUM!'):num(Math.sqrt(a));
        let b=0;
        if(node.args.length===2){const second=scalar(evaluate(node.args[1],sheetId,owner,depth+1));if(second.kind==='pending')return second;const n=numeric(second);if(typeof n!=='number')return n;b=n;}
        if(name==='MOD'){
          if(b===0)return error('#DIV/0!');
          // Remainder with the divisor's sign, without an overflowing
          // quotient for very different magnitudes.
          const r=a%b;return num(r!==0&&(r<0)!==(b<0)?r+b:r);
        }
        if(name==='POWER'){if(a===0&&b<0)return error('#DIV/0!');if(a===0&&b===0)return error('#NUM!');return num(a**b);}
        return decimalRound(a,b,name==='ROUNDUP'?'up':'down');
      }
      if(['LEN','LEFT','RIGHT','MID','UPPER','LOWER','TRIM','CONCATENATE'].includes(name)){
        const count=node.args.length;
        if(name==='CONCATENATE'?count<1:name==='MID'?count!==3:name==='LEFT'||name==='RIGHT'?count<1||count>2:count!==1)return error('#VALUE!');
        const values:CalculatedValue[]=[];
        for(const arg of node.args){const v=scalar(evaluate(arg,sheetId,owner,depth+1));if(v.kind==='pending'||v.kind==='error')return v;values.push(v);}
        const out=(text:string):CalculatedValue=>text.length>32767?error('#VALUE!'):{kind:'string',value:text};
        const text=textValue(values[0]!);
        if(name==='UPPER')return out(text.toUpperCase());
        if(name==='LOWER')return out(text.toLowerCase());
        if(name==='TRIM')return out(text.replace(/^ +| +$/g,'').replace(/ +/g,' '));
        if(name==='CONCATENATE'){
          let joined='';for(const v of values){tick();joined+=textValue(v);if(joined.length>32767)return error('#VALUE!');}return out(joined);
        }
        // Unicode code-point boundaries avoid manufacturing invalid
        // surrogate halves. Exact legacy Excel compatibility modes vary.
        const chars=Array.from(text);
        if(name==='LEN')return num(chars.length);
        const readInt=(index:number,defaultValue=1)=>{const v=values[index];if(!v)return defaultValue;const n=numeric(v);return typeof n==='number'?Math.trunc(n):n;};
        const n=readInt(1);if(typeof n!=='number')return n;
        if(name==='MID'){const size=readInt(2,0);if(typeof size!=='number')return size;return n<1||size<0?error('#VALUE!'):out(chars.slice(n-1,n-1+size).join(''));}
        if(n<0)return error('#VALUE!');
        return out((name==='LEFT'?chars.slice(0,n):chars.slice(Math.max(0,chars.length-n))).join(''));
      }
      if (name === "ABS" || name === "ROUND" || name === "NOT") {
        if (node.args.length !== (name === "ROUND" ? 2 : 1)) return error("#VALUE!");
        const first = scalar(evaluate(node.args[0], sheetId, owner, depth + 1));
        if (first.kind === "pending") return first;
        if (name === "NOT") {
          const value = logical(first);
          return typeof value === "boolean" ? bool(!value) : value;
        }
        const n = numeric(first);
        if (typeof n !== "number") return n;
        if (name === "ABS") return num(Math.abs(n));
        const second = scalar(evaluate(node.args[1], sheetId, owner, depth + 1));
        if (second.kind === "pending") return second;
        const places = numeric(second);
        return typeof places === "number" ? decimalRound(n, places) : places;
      }
      let count = 0, sum = 0, product = 1, minimum = Infinity, maximumValue = -Infinity, logicalCount = 0, and = true, or = false;
      const accept = (value: CalculatedValue | PendingValue, reference: boolean): ErrorValue | PendingValue | undefined => {
        tick();
        if (value.kind === "pending") return value;
        if (name === "COUNTA") {
          if (value.kind !== "blank") count++;
          return;
        }
        if (name === "COUNT") {
          if (value.kind === "number" || !reference && (value.kind === "boolean" || value.kind === "string" && typeof numeric(value) === "number")) count++;
          return;
        }
        if (value.kind === "error") return value;
        if (name === "AND" || name === "OR") {
          if (reference && (value.kind === "blank" || value.kind === "string")) return;
          const v = logical(value);
          if (typeof v !== "boolean") return v;
          logicalCount++;
          and = and && v;
          or = or || v;
          return;
        }
        if (reference && value.kind !== "number") return;
        const n = numeric(value);
        if (typeof n !== "number") return n;
        count++;
        sum += n;
        if(name==='PRODUCT')product*=n;
        minimum = Math.min(minimum, n);
        maximumValue = Math.max(maximumValue, n);
        return;
      };
      for (const argument of node.args) {
        const result = evaluate(argument, sheetId, owner, depth + 1);
        if (result.kind === "range-value") {
          const r = result.range;
          for (const entry of indexes.get(r.sheetId)!.range(r.row1, r.column1, r.row2, r.column2)) {
            const failure = accept(getCell(r.sheetId, entry.row, entry.column, depth + 1), true);
            if (failure) return failure;
          }
        } else {
          const failure = accept(result, argument.kind === "reference");
          if (failure) return failure;
        }
      }
      if (name === "COUNT" || name === "COUNTA") return num(count);
      if (name === "AND" || name === "OR") return logicalCount ? bool(name === "AND" ? and : or) : error("#VALUE!");
      if (name === "AVERAGE") return count ? num(sum / count) : error("#DIV/0!");
      if (name === "MIN") return num(count ? minimum : 0);
      if (name === "MAX") return num(count ? maximumValue : 0);
      if(name==='PRODUCT')return num(count?product:0);
      return num(sum);
    };
    const calculateCell = (key: string, depth = 0): CalculationCell => {
      const found = results.get(key);
      if (found) return found;
      if (active.has(key)) {
        for (let i = stack.indexOf(key); i < stack.length; i++) cycles.add(stack[i]);
        return { state: "error", value: error("#REF!"), reason: "cycle" };
      }
      assert(depth <= maxDepth, "limit", "Calculation dependency depth exceeded");
      const entry = formulaEntries.get(key);
      assert(entry, "invalid", "Missing formula entry");
      active.add(key);
      stack.push(key);
      computedCells++;
      let result: MutableResult;
      if (!entry.parsed.ast) result = { state: "unsupported", reason: "unsupported-syntax" };
      else {
        const value = scalar(evaluate(entry.parsed.ast, entry.sheetId, key, depth));
        result = value.kind === "pending" ? { state: value.state, reason: value.reason } : { state: value.kind === "error" ? "error" : "valid", value: value.kind === "blank" ? num(0) : value };
        if (cycles.has(key)) result = { state: "error", value: error("#REF!"), reason: "cycle" };
      }
      if (result.state === "unsupported" || result.state === "stale") {
        const importedCache = !options.cachedSources || options.cachedSources.get(key) === entry.parsed.source ? options.cachedValues?.get(key) : void 0;
        const retainedCache = this.#previous?.documentId === document.documentId && this.#parsed.get(key)?.source === entry.parsed.source ? this.#results.get(key)?.value : void 0;
        const cache = importedCache ?? retainedCache;
        if (cache !== void 0) result.value = cache;
      }
      if (result.value) result.value = Object.freeze({ ...result.value });
      stack.pop();
      active.delete(key);
      results.set(key, Object.freeze(result));
      return result;
    };
    const counts = new Map<string, number>(), ready: string[] = [];
    for (const key of dirty) {
      const count = [...deps.get(key) ?? []].filter((d) => { tick(); return dirty.has(d); }).length;
      counts.set(key, count);
      if (count === 0) ready.push(key);
    }
    for (let head = 0; head < ready.length; head++) {
      const key = ready[head];
      calculateCell(key);
      for (const next of reverse.get(key) ?? []) {
        if (!counts.has(next)) continue;
        tick();
        const count = counts.get(next)! - 1;
        counts.set(next, count);
        if (count === 0) ready.push(next);
      }
    }
    for (const key of dirty) if (!results.has(key)) calculateCell(key);
    options.cancellation?.throwIfCancelled();
    if (options.deadline !== undefined) assert(options.now!() <= options.deadline, "limit", "Calculation deadline exceeded");
    this.#previous = document;
    this.#parsed = parsed;
    this.#results = new Map(results);
    this.#index = rangeIndex;
    this.#sheetIndexes = sheetIndexes;
    this.#dependencies = deps;
    this.#reverse = reverse;
    return Object.freeze({ cells: new Map(results), computedCells, steps, get(sheetId: string, row: number, column: number): CalculationCell {
      assertCellAddress(row, column);
      const key = calculationKey(sheetId, row, column);
      if (results.has(key)) return results.get(key)!;
      if (!sheets.has(sheetId)) return { state: "error", value: error("#REF!") };
      const value = sheets.get(sheetId)?.cells[`${row}:${column}`]?.value ?? blank;
      return value.kind === "formula" ? { state: "stale" } : { state: value.kind === "error" ? "error" : "valid", value };
    } });
  }
  clear() {
    this.#previous = void 0;
    this.#results.clear();
    this.#parsed.clear();
    this.#index = new RangeDependencyIndex([]);
    this.#sheetIndexes.clear();
    this.#dependencies.clear();
    this.#reverse.clear();
  }
}
