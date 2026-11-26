/**
 * SHT-03 / I18N-03: recovered from the 2026-11-25 M3 browser bundle.
 * Type declarations and invalid end-of-input handling are reconstructed.
 * This is a candidate parser, not the complete formula engine or original source.
 */
import { assert, assertCellAddress, OfficeError, type Cancellation } from '@office-ex/contracts';
import type { CellValue } from '@office-ex/core';
import { numberFormatter } from './intl-cache.js';

export type FormulaError = Extract<CellValue, {kind:'error'}>['value'];
export type BinaryOperator = '=' | '<>' | '<' | '>' | '<=' | '>=' | '&' | '+' | '-' | '*' | '/' | '^';
export interface FormulaReference { kind: 'reference'; sheet?: string; row: number; column: number; rowAbsolute: boolean; columnAbsolute: boolean }
export interface FormulaRange { kind: 'range'; start: FormulaReference; end: FormulaReference; axis?: 'rows' | 'columns' }
export type FormulaNode =
  | FormulaReference | FormulaRange
  | { kind: 'number'; value: number }
  | { kind: 'string'; value: string }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'error'; value: FormulaError }
  | { kind: 'missing' }
  | { kind: 'unary'; operator: '+' | '-' | '%'; operand: FormulaNode }
  | { kind: 'binary'; operator: BinaryOperator; left: FormulaNode; right: FormulaNode }
  | { kind: 'call'; name: string; args: FormulaNode[] };
export interface ParseFormulaOptions { mode?: 'canonical' | 'regional'; locale?: string; maxTokens?: number; maxDepth?: number; cancellation?: Cancellation }
type TokenKind = 'number' | 'word' | 'quoted' | 'string' | 'symbol' | 'error' | 'end';
interface Token { kind: TokenKind; text: string }
interface Separators { readonly decimal: string; readonly argument: string }

export const SUPPORTED_FUNCTIONS = Object.freeze(["SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "IF", "AND", "OR", "NOT", "ROUND", "ABS", "IFERROR", "ISNUMBER", "ISTEXT", "ISLOGICAL", "ISBLANK", "ISERROR", "ISERR", "ISNA", "PRODUCT", "INT", "MOD", "ROUNDUP", "ROUNDDOWN", "TRUNC", "SIGN", "SQRT", "POWER", "LEN", "LEFT", "RIGHT", "MID", "UPPER", "LOWER", "TRIM", "CONCATENATE", "COUNTBLANK", "COUNTIF", "COUNTIFS", "SUMIF", "SUMIFS", "AVERAGEIF", "AVERAGEIFS", "VLOOKUP", "HLOOKUP", "MATCH", "INDEX"]);
var functions = new Set(SUPPORTED_FUNCTIONS);
export function isSupportedFunction(name: string): boolean {
  return functions.has(name.toUpperCase());
}
export const FORMULA_ALIASES: Readonly<Record<string, Readonly<Record<string, string>>>> = Object.freeze({
  de: Object.freeze({ SUMME: "SUM", MITTELWERT: "AVERAGE", ANZAHL: "COUNT", ANZAHL2: "COUNTA", WENN: "IF", UND: "AND", ODER: "OR", NICHT: "NOT", RUNDEN: "ROUND" }),
  fr: Object.freeze({ SOMME: "SUM", MOYENNE: "AVERAGE", NB: "COUNT", NBVAL: "COUNTA", SI: "IF", ET: "AND", OU: "OR", NON: "NOT", ARRONDI: "ROUND" })
});
export function columnName(column: number): string {
  assertCellAddress(1, column);
  let text2 = "";
  for (let current = column; current; current = Math.floor((current - 1) / 26)) text2 = String.fromCharCode(65 + (current - 1) % 26) + text2;
  return text2;
}
export function columnNumber(name: string): number {
  assert(/^[A-Za-z]{1,3}$/.test(name), "invalid", "Invalid column label");
  let result = 0;
  for (const ch of name.toUpperCase()) result = result * 26 + ch.charCodeAt(0) - 64;
  assertCellAddress(1, result);
  return result;
}
const separatorsCache = new Map<string, Separators>();
export function formulaSeparators(locale: string): Separators {
  const found = separatorsCache.get(locale);
  if (found) return found;
  const decimal = numberFormatter(locale, { useGrouping: false }).formatToParts(1.1).find((p) => p.type === "decimal")?.value ?? ".";
  const result = Object.freeze({ decimal, argument: decimal === "," ? ";" : "," });
  if (separatorsCache.size >= 128) { const oldest = separatorsCache.keys().next().value; if (oldest !== undefined) separatorsCache.delete(oldest); }
  separatorsCache.set(locale, result);
  return result;
}
var errorPattern = /^(?:#NULL!|#DIV\/0!|#VALUE!|#REF!|#NAME\?|#NUM!|#N\/A)/i;
function lex(source: string, options: ParseFormulaOptions): Token[] {
  assert(source.startsWith("=") && source.length > 1 && source.length <= 8192, "invalid", "Invalid formula length or prefix");
  const separators = options.mode === "regional" ? formulaSeparators(options.locale ?? "en-US") : { decimal: ".", argument: "," };
  const tokens: Token[] = [];
  const max = options.maxTokens ?? 8192;
  assert(Number.isInteger(max) && max > 0 && max <= 32768, "invalid", "Invalid token budget");
  const push = (kind: TokenKind, text2: string) => {
    assert(tokens.length < max, "limit", "Formula token budget exceeded");
    tokens.push({ kind, text: text2 });
  };
  for (let index = 1; index < source.length; ) {
    options.cancellation?.throwIfCancelled();
    const ch = source[index];
    if (/\s/u.test(ch)) {
      index++;
      continue;
    }
    if (ch === '"' || ch === "'") {
      const quote = ch;
      index++;
      let text2 = "";
      let closed = false;
      while (index < source.length) {
        const c = source[index++];
        if (c === quote) {
          if (source[index] === quote) {
            text2 += quote;
            index++;
          } else {
            closed = true;
            break;
          }
        } else text2 += c;
      }
      assert(closed, "invalid", "Unterminated formula literal");
      push(quote === '"' ? "string" : "quoted", text2);
      continue;
    }
    if (ch === "#") {
      const match = errorPattern.exec(source.slice(index));
      assert(match, "unsupported", "Unsupported error or spilled reference");
      push("error", match[0].toUpperCase());
      index += match[0].length;
      continue;
    }
    if (/[0-9]/.test(ch) || ch === separators.decimal && /[0-9]/.test(source[index + 1] ?? "")) {
      const start = index;
      let decimal = false;
      while (index < source.length && (/[0-9]/.test(source[index]) || !decimal && source[index] === separators.decimal)) {
        if (source[index] === separators.decimal) decimal = true;
        index++;
      }
      if (/[Ee]/.test(source[index] ?? "") && /^[Ee][+-]?\d/.test(source.slice(index))) {
        index++;
        if (/[+-]/.test(source[index] ?? "")) index++;
        while (/\d/.test(source[index] ?? "")) index++;
      }
      const text2 = source.slice(start, index).replace(separators.decimal, ".");
      assert(Number.isFinite(Number(text2)), "invalid", "Non-finite formula literal");
      push("number", text2);
      continue;
    }
    if (/[\p{L}_$\\]/u.test(ch)) {
      const start = index++;
      while (index < source.length && /[\p{L}\p{N}_.$\\]/u.test(source[index])) index++;
      push("word", source.slice(start, index));
      continue;
    }
    const pair = source.slice(index, index + 2);
    if (["<>", "<=", ">="].includes(pair)) {
      push("symbol", pair);
      index += 2;
      continue;
    }
    if (ch === separators.argument) {
      push("symbol", ",");
      index++;
      continue;
    }
    assert("+-*/^&=<>%():!".includes(ch), "unsupported", "Unsupported formula syntax");
    push("symbol", ch);
    index++;
  }
  tokens.push({ kind: "end", text: "" });
  return tokens;
}
const precedence: Record<string, number> = { "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5 };
export function parseFormula(source: string, options: ParseFormulaOptions = {}): FormulaNode {
  const tokens = lex(source, options);
  let cursor = 0;
  const maxDepth = options.maxDepth ?? 128;
  assert(Number.isInteger(maxDepth) && maxDepth >= 1 && maxDepth <= 256, "invalid", "Invalid formula depth budget");
  const peek = (): Token => tokens[cursor] ?? { kind: "end", text: "" };
  const take = (): Token => { const token = peek(); assert(token.kind !== "end", "invalid", "Unexpected end of formula"); cursor++; return token; };
  const symbol = (text: string): boolean => peek().kind === "symbol" && peek().text === text;
  const expect = (text2: string) => {
    const token = take();
    assert(token.kind === "symbol" && token.text === text2, "invalid", "Unexpected formula token");
  };
  const reference = (token: Token, sheet?: string): FormulaReference | undefined => {
    if (token.kind !== "word") return undefined;
    const match = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/.exec(token.text);
    if (!match) return void 0;
    const row = Number(match[4]), column = columnNumber(match[2]);
    assertCellAddress(row, column);
    return { kind: "reference", ...sheet !== void 0 ? { sheet } : {}, row, column, rowAbsolute: !!match[3], columnAbsolute: !!match[1] };
  };
  const parseReference = (first: Token, sheet?: string): FormulaReference | FormulaRange | undefined => {
    if (first.kind !== "word" && first.kind !== "number") return undefined;
    const start = reference(first, sheet);
    if (start) {
      if (!symbol(":")) return start;
      take();
      let endToken2 = take();
      let endSheet = sheet;
      if (symbol("!")) {
        assert(endToken2.kind === "word" || endToken2.kind === "quoted", "invalid", "Invalid range sheet");
        endSheet = endToken2.text;
        take();
        endToken2 = take();
      }
      const end = reference(endToken2, endSheet);
      assert(end, "unsupported", "Unsupported range endpoint");
      assert((end.sheet ?? sheet) === sheet, "unsupported", "Three-dimensional range is unsupported");
      return { kind: "range", start, end };
    }
    if (!symbol(":")) return void 0;
    const col = /^(\$?)([A-Za-z]{1,3})$/.exec(first.text), row = /^(\$?)([1-9]\d*)$/.exec(first.text);
    assert(col || row, "unsupported", "Unsupported reference range");
    take();
    const endToken = take();
    assert(endToken.kind === "word" || endToken.kind === "number", "invalid", "Invalid range endpoint token");
    if (col) {
      const other2 = /^(\$?)([A-Za-z]{1,3})$/.exec(endToken.text);
      assert(other2, "invalid", "Invalid whole column range");
      return { kind: "range", axis: "columns", start: { kind: "reference", ...sheet !== void 0 ? { sheet } : {}, row: 1, column: columnNumber(col[2]), rowAbsolute: false, columnAbsolute: !!col[1] }, end: { kind: "reference", ...sheet !== void 0 ? { sheet } : {}, row: 1048576, column: columnNumber(other2[2]), rowAbsolute: false, columnAbsolute: !!other2[1] } };
    }
    const other = /^(\$?)([1-9]\d*)$/.exec(endToken.text);
    assert(other, "invalid", "Invalid whole row range");
    assert(row, "invalid", "Invalid whole row range");
    const firstRow = Number(row[2]), lastRow = Number(other[2]);
    assertCellAddress(firstRow, 1);
    assertCellAddress(lastRow, 1);
    return { kind: "range", axis: "rows", start: { kind: "reference", ...sheet !== void 0 ? { sheet } : {}, row: firstRow, column: 1, rowAbsolute: !!row[1], columnAbsolute: false }, end: { kind: "reference", ...sheet !== void 0 ? { sheet } : {}, row: lastRow, column: 16384, rowAbsolute: !!other[1], columnAbsolute: false } };
  };
  const primary = (depth: number): FormulaNode => {
    assert(depth <= maxDepth, "limit", "Formula AST depth exceeded");
    options.cancellation?.throwIfCancelled();
    const token = take();
    if (token.kind === "symbol" && (token.text === "+" || token.text === "-")) return { kind: "unary", operator: token.text, operand: primary(depth + 1) };
    if (token.kind === "symbol" && token.text === "(") {
      const value = expression(1, depth + 1);
      expect(")");
      return value;
    }
    if (token.kind === "string") return { kind: "string", value: token.text };
    if (token.kind === "error") return { kind: "error", value: token.text as FormulaError };
    if ((token.kind === "word" || token.kind === "quoted") && symbol("!")) {
      take();
      const ref2 = parseReference(take(), token.text);
      assert(ref2, "unsupported", "Only cell and range sheet references are supported");
      return ref2;
    }
    if (token.kind === "word" && symbol("(")) {
      take();
      const args: FormulaNode[] = [];
      if (!symbol(")")) {
        while (true) {
          assert(args.length < 255, "limit", "Function argument budget exceeded");
          args.push(symbol(",") || symbol(")") ? { kind: "missing" } : expression(1, depth + 1));
          if (!symbol(",")) break;
          take();
        }
      }
      expect(")");
      const upper = token.text.toUpperCase();
      const name = options.mode === "regional" ? FORMULA_ALIASES[(options.locale ?? "en-US").split("-")[0]]?.[upper] ?? upper : upper;
      return { kind: "call", name, args };
    }
    const ref = parseReference(token);
    if (ref) return ref;
    if (token.kind === "number") return { kind: "number", value: Number(token.text) };
    if (token.kind === "word" && /^(TRUE|FALSE)$/i.test(token.text)) return { kind: "boolean", value: token.text.toUpperCase() === "TRUE" };
    throw new OfficeError("unsupported", "Unsupported named reference or formula token");
  };
  const expression = (minimum: number, depth: number): FormulaNode => {
    assert(depth <= maxDepth, "limit", "Formula AST depth exceeded");
    let left = primary(depth);
    while (symbol("%")) {
      take();
      left = { kind: "unary", operator: "%", operand: left };
    }
    while (peek().kind === "symbol" && (precedence[peek().text] ?? 0) >= minimum) {
      const operator = take().text as BinaryOperator;
      const right = expression(precedence[operator] + 1, depth + 1);
      left = { kind: "binary", operator, left, right };
    }
    return left;
  };
  const ast = expression(1, 1);
  assert(peek().kind === "end", "unsupported", "Unsupported trailing formula syntax");
  walkFormula(ast, () => {
  }, maxDepth);
  return ast;
}
export function walkFormula(ast: FormulaNode, visit: (node: FormulaNode) => void, maxDepth = 128): void {
  const stack = [{ node: ast, depth: 1 }];
  let count = 0;
  while (stack.length) {
    const { node, depth } = stack.pop()!;
    assert(depth <= maxDepth && ++count <= 8192, "limit", "Formula AST budget exceeded");
    visit(node);
    if (node.kind === "binary") stack.push({ node: node.right, depth: depth + 1 }, { node: node.left, depth: depth + 1 });
    else if (node.kind === "unary") stack.push({ node: node.operand, depth: depth + 1 });
    else if (node.kind === "call") for (let i = node.args.length - 1; i >= 0; i--) stack.push({ node: node.args[i], depth: depth + 1 });
  }
}
export function formulaIsSupported(ast: FormulaNode): boolean {
  let supported = true;
  walkFormula(ast, (n) => {
    if (n.kind === "call" && !isSupportedFunction(n.name)) supported = false;
  });
  return supported;
}
function refText(ref: FormulaReference, includeSheet = true): string {
  return `${includeSheet && ref.sheet !== void 0 ? `'${ref.sheet.replaceAll("'", "''")}'!` : ""}${ref.columnAbsolute ? "$" : ""}${columnName(ref.column)}${ref.rowAbsolute ? "$" : ""}${ref.row}`;
}
export function serializeFormula(ast: FormulaNode): string {
  walkFormula(ast, () => {
  });
  const part = (node: FormulaNode): string => {
    switch (node.kind) {
      case "number":
        assert(Number.isFinite(node.value), "invalid", "Invalid formula number");
        return String(node.value);
      case "string":
        return `"${node.value.replaceAll('"', '""')}"`;
      case "boolean":
        return node.value ? "TRUE" : "FALSE";
      case "error":
        return node.value;
      case "missing":
        return "";
      case "reference":
        return refText(node);
      case "range": {
        const prefix = node.start.sheet === void 0 ? "" : `'${node.start.sheet.replaceAll("'", "''")}'!`;
        if (node.axis === "columns") return `${prefix}${node.start.columnAbsolute ? "$" : ""}${columnName(node.start.column)}:${node.end.columnAbsolute ? "$" : ""}${columnName(node.end.column)}`;
        if (node.axis === "rows") return `${prefix}${node.start.rowAbsolute ? "$" : ""}${node.start.row}:${node.end.rowAbsolute ? "$" : ""}${node.end.row}`;
        return `${refText(node.start)}:${refText(node.end, node.end.sheet !== node.start.sheet)}`;
      }
      case "unary":
        return node.operator === "%" ? `(${part(node.operand)})%` : `${node.operator}(${part(node.operand)})`;
      case "binary":
        return `(${part(node.left)}${node.operator}${part(node.right)})`;
      case "call":
        return `${node.name.toUpperCase()}(${node.args.map(part).join(",")})`;
    }
  };
  const source = `=${part(ast)}`;
  assert(source.length <= 8192, "limit", "Serialized formula exceeds length limit");
  return source;
}
export function localizeFormula(source: string, locale: string): string {
  const ast = parseFormula(source);
  const aliases = FORMULA_ALIASES[locale.split("-")[0]] ?? {};
  walkFormula(ast, (n) => {
    if (n.kind === "call") n.name = Object.entries(aliases).find(([, canonical3]) => canonical3 === n.name)?.[0] ?? n.name;
  });
  const canonical2 = serializeFormula(ast), separators = formulaSeparators(locale);
  let out = "", quoted = "";
  for (let i = 0; i < canonical2.length; i++) {
    const ch = canonical2[i];
    if (quoted) {
      out += ch;
      if (ch === quoted) {
        if (canonical2[i + 1] === quoted) out += canonical2[++i];
        else quoted = "";
      }
    } else if (ch === '"' || ch === "'") {
      quoted = ch;
      out += ch;
    } else if (ch === ",") out += separators.argument;
    else if (ch === "." && /\d/.test(canonical2[i - 1] ?? "") && /\d/.test(canonical2[i + 1] ?? "")) out += separators.decimal;
    else out += ch;
  }
  return out;
}
