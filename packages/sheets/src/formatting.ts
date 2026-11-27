/** Reconstructed number/date formatting from the M3 browser artifact. */
import { assert,graphemeBoundaries } from '@office-ex/contracts';
import type { Cell, CellValue, SheetRichRun, SheetConditionalRule, ConditionalCellStyle, DeepReadonly } from '@office-ex/core';
import { parseFormula, serializeFormula, formulaSeparators, type ParseFormulaOptions, type FormulaError } from './formula.js';
import { numberFormatter, dateFormatter } from './intl-cache.js';
import type { CalculationCell, CalculatedValue } from './calculation.js';
export interface DisplayOptions { locale?:string; dateSystem?:'1900'|'1904' }
export interface CellInputOptions extends ParseFormulaOptions, DisplayOptions { type?:'text'|'date'|'number' }
interface NumericRegion { decimal:string;group?:string;digits:Map<string,string> }
var dayMilliseconds = 864e5;
function utcDate(year:number, month:number, day:number): Date {
  const date = /* @__PURE__ */ new Date(0);
  date.setUTCFullYear(year, month - 1, day);
  date.setUTCHours(0, 0, 0, 0);
  return date;
}
export function dateToSerial(year:number, month:number, day:number, dateSystem: "1900"|"1904" = "1900"):number {
  assert([year, month, day].every(Number.isInteger) && year >= 1 && year <= 9999 && month >= 1 && month <= 12 && day >= 1 && day <= 31, "invalid", "Invalid calendar date");
  if (dateSystem === "1900" && year === 1900 && month === 2 && day === 29) return 60;
  const date = utcDate(year, month, day);
  assert(date.getUTCFullYear() === year && date.getUTCMonth() + 1 === month && date.getUTCDate() === day, "invalid", "Invalid calendar date");
  if (dateSystem === "1904") return (date.getTime() - utcDate(1904, 1, 1).getTime()) / dayMilliseconds;
  const serial = (date.getTime() - utcDate(1899, 12, 31).getTime()) / dayMilliseconds;
  return serial >= 60 ? serial + 1 : serial;
}
export function serialToDate(serial:number, dateSystem:"1900"|"1904" = "1900") {
  assert(Number.isFinite(serial) && serial >= 0 && serial <= (dateSystem === "1904" ? 295700399999999e-8 : 295846599999999e-8), "invalid", "Excel date is outside supported display range");
  let whole = Math.floor(serial), ms = Math.round((serial - whole) * dayMilliseconds);
  if (ms === dayMilliseconds) {
    whole++;
    ms = 0;
  }
  const fictitiousLeapDay = dateSystem === "1900" && whole === 60;
  const date = dateSystem === "1904" ? new Date(utcDate(1904, 1, 1).getTime() + whole * dayMilliseconds) : new Date(utcDate(1899, 12, 31).getTime() + (whole > 60 ? whole - 1 : fictitiousLeapDay ? 59 : whole) * dayMilliseconds);
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: fictitiousLeapDay ? 29 : date.getUTCDate(), hour: Math.floor(ms / 36e5), minute: Math.floor(ms / 6e4) % 60, second: Math.floor(ms / 1e3) % 60, millisecond: ms % 1e3, fictitiousLeapDay };
}
const numericRegions = new Map<string,NumericRegion>();
function numericRegion(locale:string):NumericRegion {
  const found = numericRegions.get(locale);
  if (found) return found;
  const parts = numberFormatter(locale).formatToParts(123456789e-2);
  const data = { decimal: formulaSeparators(locale).decimal, group: parts.find((p) => p.type === "group")?.value, digits: new Map(Array.from({ length: 10 }, (_, d) => [numberFormatter(locale, { useGrouping: false }).format(d), String(d)])) };
  if (numericRegions.size >= 128) numericRegions.delete(numericRegions.keys().next().value!);
  numericRegions.set(locale, data);
  return data;
}
function numericInput(text2:string, options:CellInputOptions):number|undefined {
  let normalized = text2, percent = false;
  if (normalized.endsWith("%")) {
    percent = true;
    normalized = normalized.slice(0, -1).trim();
  }
  if (options.mode === "regional") {
    const locale = options.locale ?? "en-US", { decimal, group, digits } = numericRegion(locale);
    normalized = [...normalized].map((c) => digits.get(c) ?? c).join("");
    if (group && normalized.includes(group)) {
      const sign = /^[+-]/.test(normalized) ? normalized[0] : "", unsigned = sign ? normalized.slice(1) : normalized;
      const chunks = unsigned.split(decimal);
      if (chunks.length > 2 || chunks[1]?.includes(group)) return void 0;
      const integer2 = chunks[0], raw = integer2.split(group).join("");
      if (!/^\d+$/.test(raw)) return void 0;
      const canonical2 = numberFormatter(locale, { maximumFractionDigits: 0 }).format(Number(raw));
      const mapped = [...canonical2].map((c) => digits.get(c) ?? c).join("");
      if (mapped !== integer2) return void 0;
      normalized = sign + raw + (chunks[1] !== void 0 ? "." + chunks[1] : "");
    } else if (decimal !== ".") {
      if (normalized.includes(".")) return void 0;
      normalized = normalized.replace(decimal, ".");
    }
  }
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(normalized)) return void 0;
  const value = Number(normalized) / (percent ? 100 : 1);
  assert(Number.isFinite(value), "invalid", "Numeric input is outside binary64 range");
  return value;
}
export function parseCellInput(input:string, options:CellInputOptions = {}):CellValue {
  assert(input.length <= 32767, "limit", "Cell text length exceeded");
  if (options.type === "text" || input.startsWith("'")) return { kind: "string", value: input.startsWith("'") && options.type !== "text" ? input.slice(1) : input };
  const text2 = input.trim();
  if (text2 === "") return input === "" ? { kind: "blank" } : { kind: "string", value: input };
  if (text2.startsWith("=")) {
    const ast = parseFormula(text2, options);
    return { kind: "formula", source: options.mode === "regional" ? serializeFormula(ast) : text2 };
  }
  if (options.type === "date") {
    const date = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(text2);
    assert(date, "invalid", "Use unambiguous ISO date input");
    const hour = Number(date[4] ?? 0), minute = Number(date[5] ?? 0), second = Number(date[6] ?? 0);
    assert(hour <= 23 && minute <= 59 && second <= 59, "invalid", "Invalid date time");
    return { kind: "number", value: dateToSerial(Number(date[1]), Number(date[2]), Number(date[3]), options.dateSystem) + (hour * 3600 + minute * 60 + second) / 86400 };
  }
  const value = numericInput(text2, options);
  if (value !== void 0) return { kind: "number", value };
  assert(options.type !== "number", "invalid", "Ambiguous or invalid regional number");
  if (/^(TRUE|FALSE)$/i.test(text2)) return { kind: "boolean", value: text2.toUpperCase() === "TRUE" };
  if (/^(#NULL!|#DIV\/0!|#VALUE!|#REF!|#NAME\?|#NUM!|#N\/A)$/i.test(text2)) return { kind: "error", value: text2.toUpperCase() as FormulaError };
  return { kind: "string", value: input };
}
/** Preserve unchanged prefix/suffix runs using grapheme boundaries. New
 * text inherits the preceding run (or first run at the beginning). */
function replaceRichText(runs:readonly SheetRichRun[],before:string,after:string):SheetRichRun[]|undefined{
  if(!after)return undefined;if(before===after)return runs.slice();
  const a=graphemeBoundaries(before),b=graphemeBoundaries(after);let head=0,tail=0;
  while(head<a.length-1&&head<b.length-1&&before.slice(a[head],a[head+1])===after.slice(b[head],b[head+1]))head++;
  while(tail<a.length-1-head&&tail<b.length-1-head&&before.slice(a[a.length-2-tail],a[a.length-1-tail])===after.slice(b[b.length-2-tail],b[b.length-1-tail]))tail++;
  const from=a[head]!,end=a[a.length-1-tail]!,out:SheetRichRun[]=[];
  const append=(text:string,style:SheetRichRun['style'])=>{
    if(!text)return;const last=out.at(-1);
    if(last&&JSON.stringify(last.style??{})===JSON.stringify(style??{}))last.text+=text;
    else out.push({text,...(style?{style}:{})});
  };
  let cursor=0,inserted=false,inherit=runs[0]?.style;
  for(const run of runs){
    const next=cursor+run.text.length;
    if(cursor<from){append(run.text.slice(0,Math.min(run.text.length,from-cursor)),run.style);inherit=run.style;}
    if(!inserted&&next>=from){append(after.slice(b[head],b[b.length-1-tail]),inherit);inserted=true;}
    if(next>end)append(run.text.slice(Math.max(0,end-cursor)),run.style);
    cursor=next;
  }
  if(!inserted)append(after.slice(b[head],b[b.length-1-tail]),inherit);
  assert(out.length<=256,'limit','Rich cell run budget exceeded');
  return out.length?out:undefined;
}
export function cellFromInput(input:string, options:CellInputOptions = {}, before?:Cell):Cell {
  const value = parseCellInput(input, options);
  const numberFormat = options.type === "date" ? /[ T]\d\d:/.test(input) ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd" : value.kind === "number" && input.trim().endsWith("%") ? "0.00%" : before?.numberFormat;
  const result:Cell={...before,value,...(numberFormat!==undefined?{numberFormat}:{})};
  if(before?.richText&&before.value.kind==='string'&&value.kind==='string'){
    const rich=replaceRichText(before.richText,before.value.value,value.value);
    if(rich)result.richText=rich;else delete result.richText;
  }else delete result.richText;
  return result;
}
function splitSections(code:string):string[] {
  const sections = [];
  let text2 = "", quote = false, escaped = false;
  for (const ch of code) {
    if (escaped) {
      text2 += ch;
      escaped = false;
    } else if (ch === "\\") {
      text2 += ch;
      escaped = true;
    } else if (ch === '"') {
      quote = !quote;
      text2 += ch;
    } else if (ch === ";" && !quote) {
      sections.push(text2);
      text2 = "";
    } else text2 += ch;
  }
  sections.push(text2);
  return sections;
}
function unquoted(code:string):string {
  return code.replace(/"(?:[^"]|"")*"/g, "").replace(/\\./g, "").replace(/\[[^\]]*\]/g, "");
}
function literal(code:string):string {
  return code.replace(/\[\$([^\]\-]+)(?:-[0-9a-f]+)?\]/gi, "$1").replace(/"([^"]*)"/g, "$1").replace(/\\(.)/g, "$1").replace(/_.|\*./g, "").replace(/\[[^\]]*\]/g, "");
}
export function isSupportedNumberFormat(format = "General"):boolean {
  if (format === "General" || format === "@") return true;
  const clean = unquoted(format);
  if ((clean.match(/\.[0#]+/g) ?? []).some((part) => part.length > 21)) return false;
  if (/\[(?:[<>=][^\]]*|[hms]+)\]/i.test(format) || /\?/.test(clean)) return false;
  if (/[ymdhs]/i.test(clean)) return !/[^ymdhsapm0:/ .,+\-()]/i.test(clean);
  return !/[^0#.,%Ee+\-();$€£¥₹₩₽@ ]/.test(clean);
}
function formatDate(serial:number, code:string, options:DisplayOptions):string {
  if (serial < 0 || serial > (options.dateSystem === "1904" ? 295700399999999e-8 : 295846599999999e-8)) return "########";
  const date = serialToDate(serial, options.dateSystem), locale = options.locale ?? "en-US", native = utcDate(date.year, date.month, date.fictitiousLeapDay ? 28 : date.day), hasPeriod = /AM\/PM/i.test(code);
  const pad = (n:number, count = 2) => numberFormatter(locale, { useGrouping: false, minimumIntegerDigits: count, maximumFractionDigits: 0 }).format(n);
  const tokens = code.match(/"[^"]*"|\\.|AM\/PM|yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|\.0+|[^ymdhs]/gi) ?? [];
  return tokens.map((token, index) => {
    if (token.startsWith('"') || token.startsWith("\\")) return literal(token);
    const t = token.toLowerCase();
    if (t === "yyyy") return pad(date.year, 4);
    if (t === "yy") return pad(date.year % 100);
    if (t === "am/pm") return date.hour >= 12 ? "PM" : "AM";
    if (/^h{1,2}$/.test(t)) return pad(hasPeriod ? date.hour % 12 || 12 : date.hour, t.length);
    if (/^s{1,2}$/.test(t)) return pad(date.second, t.length);
    if (/^d{1,2}$/.test(t)) return pad(date.day, t.length);
    if (t === "ddd" || t === "dddd") return dateFormatter(locale, { weekday: t === "ddd" ? "short" : "long", timeZone: "UTC", calendar: "gregory" }).format(native);
    if (t === "mmm" || t === "mmmm") return dateFormatter(locale, { month: t === "mmm" ? "short" : "long", timeZone: "UTC", calendar: "gregory" }).format(native);
    if (/^m{1,2}$/.test(t)) {
      const before = tokens.slice(0, index).filter((s) => /[a-z]/i.test(s)).at(-1) ?? "", after = tokens.slice(index + 1).find((s) => /[a-z]/i.test(s)) ?? "";
      const minutes = /^h/i.test(before) || /^s/i.test(after);
      return pad(minutes ? date.minute : date.month, t.length);
    }
    if (/^\.0+$/.test(t)) return formulaSeparators(locale).decimal + pad(date.millisecond, 3).slice(0, t.length - 1).padEnd(t.length - 1, "0");
    return token;
  }).join("");
}
export function formatCellValue(value:CalculatedValue, numberFormat = "General", options:DisplayOptions = {}):string {
  if (value.kind === "blank") return "";
  if (value.kind === "string") {
    const sections2 = splitSections(numberFormat);
    return sections2.length > 3 ? literal(sections2[3] ?? "").replace(/@/g, () => value.value) : value.value;
  }
  if (value.kind === "boolean") return value.value ? "TRUE" : "FALSE";
  if (value.kind === "error") return value.value;
  const locale = options.locale ?? "en-US", n = value.value;
  if (numberFormat === "General" || numberFormat === "@" || !isSupportedNumberFormat(numberFormat)) return numberFormatter(locale, { maximumSignificantDigits: 15, useGrouping: false }).format(n);
  const sections = splitSections(numberFormat);
  let code = sections[n < 0 && sections.length > 1 ? 1 : n === 0 && sections.length > 2 ? 2 : 0] ?? "";
  code = code.replace(/\[(?:Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\d+)\]/gi, "");
  if (/[ymdhs]/i.test(unquoted(code))) return formatDate(n, code, options);
  const tokens = code.match(/"[^"]*"|\[[^\]]*\]|\\.|[^]/g) ?? [];
  const isPlaceholder = (token:string) => /^[0#]$/.test(token);
  const first = tokens.findIndex(isPlaceholder);
  let last = tokens.length - 1;
  while (last >= 0 && !isPlaceholder(tokens[last])) last--;
  if (first < 0) return literal(code);
  let tail = last + 1, scale = 0;
  while (tokens[tail] === ",") {
    scale++;
    tail++;
  }
  const body = tokens.slice(first, last + 1).join(""), percentCount = (unquoted(code).match(/%/g) ?? []).length;
  const magnitude = Math.abs(n) * 100 ** percentCount / 1e3 ** scale;
  if (!Number.isFinite(magnitude)) return "########";
  const decimalPart = body.split(".")[1]?.split(/[Ee]/)[0] ?? "", minimumFractionDigits = (decimalPart.match(/0/g) ?? []).length, maximumFractionDigits = (decimalPart.match(/[0#]/g) ?? []).length;
  assert(maximumFractionDigits <= 20, "unsupported", "Number format precision exceeds display budget");
  let formatted;
  if (/[Ee][+-]?0+/.test(body)) {
    const [coefficient, exponent] = magnitude.toExponential(maximumFractionDigits).split("e");
    const sign = Number(exponent) >= 0 ? "+" : "-", exponentDigits = body.split(/[Ee][+-]?/)[1]?.length ?? 2;
    formatted = numberFormatter(locale, { minimumFractionDigits, maximumFractionDigits, useGrouping: false }).format(Number(coefficient)) + "E" + sign + String(Math.abs(Number(exponent))).padStart(exponentDigits, "0");
  } else formatted = numberFormatter(locale, { minimumFractionDigits, maximumFractionDigits, minimumIntegerDigits: Math.max(1, Math.min(21, (body.split(".")[0].replaceAll(",", "").match(/0/g) ?? []).length)), useGrouping: body.split(".")[0].includes(",") }).format(magnitude);
  const negative = n < 0 && sections.length < 2 ? "-" : "";
  return negative + literal(tokens.slice(0, first).join("")) + formatted + literal(tokens.slice(tail).join(""));
}
export function formatCell(cell:Readonly<Cell>|undefined, calculation?:CalculationCell, options:DisplayOptions = {}):string {
  if (!cell) return "";
  if (cell.value.kind !== "formula") return formatCellValue(cell.value, cell.numberFormat, options);
  return calculation?.value ? formatCellValue(calculation.value, cell.numberFormat, options) : cell.value.source;
}

/** Conditional paint is derived only from current numeric values. The
 * highest-priority matching rule wins each property independently. */
export function conditionalCellStyle(rules:readonly DeepReadonly<SheetConditionalRule>[],row:number,column:number,value:CalculatedValue|undefined):DeepReadonly<ConditionalCellStyle>|undefined{
 if(value?.kind!=='number')return undefined;
 const n=value.value,result:ConditionalCellStyle={};
 for(const rule of rules){
  if(!rule.ranges.some(r=>row>=r.row&&row<r.row+r.rowCount&&column>=r.column&&column<r.column+r.columnCount))continue;
  const a=rule.values[0]!,b=rule.values[1]??a;
  const match=rule.operator==='equal'?n===a:rule.operator==='notEqual'?n!==a:rule.operator==='lessThan'?n<a:rule.operator==='lessThanOrEqual'?n<=a:rule.operator==='greaterThan'?n>a:rule.operator==='greaterThanOrEqual'?n>=a:rule.operator==='between'?n>=a&&n<=b:n<a||n>b;
  if(!match)continue;
  for(const [key,v]of Object.entries(rule.style))if(!Object.hasOwn(result,key))Object.assign(result,{[key]:v});
  if(rule.stopIfTrue)break;
 }
 return Object.keys(result).length?result:undefined;
}
