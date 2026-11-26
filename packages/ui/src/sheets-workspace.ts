import {installRibbon,ribbonNodes,type RibbonState,decorateOfficeButton} from './office-ribbon.js';
/** Recovered M3 virtualized worksheet workspace. */
import { OfficeError, EMU_PER_CSS_PIXEL } from '@office-ex/contracts';
import { patchCommand, type DocumentSnapshot, type DocumentCommand, type Color, type CellStyle, type SheetView,type ConditionalOperator,type SheetChart,type DeepReadonly } from '@office-ex/core';
import { addSheet, renameSheet, deleteSheet, moveSheet, insertRows, deleteRows, insertColumns, deleteColumns, mergeCells, unmergeCells, clearCells, copyCells, formatCells, applyNamedCellStyle, addConditionalRule, removeConditionalRule, conditionalCellStyle, addSheetChart, removeSheetChart, setDimensions, fillCells, pasteTsv, exportTsv, setCells, cellFromInput, formatCell, formatCellValue, calculationKey, localizeFormula, formulaSeparators, type CellRange, type WorkbookSnapshot, type CalculationCell, type CalculatedValue, type CellInputOptions } from '@office-ex/sheets';
import type { MessageKey } from '@office-ex/i18n';
import { GridAxis, columnLabel, cellLabel } from './sheets-geometry.js';

type CellSnapshot = WorkbookSnapshot['sheets'][number]['cells'][string];
export interface SheetsViewState extends RibbonState {
  sheetId?:string; row:number; column:number; anchorRow:number; anchorColumn:number;
  scrollLeft:number; scrollTop:number; inputMode:'canonical'|'regional';
  calculation?:Map<string,CalculationCell>; calculatedRevision?:number;
  calculating?:boolean; calculationError?:boolean; onCalculation?:()=>void;
  cachedValues?:[string,CalculatedValue][]; cachedSources?:[string,string][];
  unsupportedFormulas?:string[]; flush?:()=>boolean;
  navigate?: (row:number, column:number, endRow:number, endColumn:number, focusGrid:boolean)=>void;
}
export interface SheetsHost {
  snapshot():DocumentSnapshot;
  execute(command:DocumentCommand):boolean;
  render():void; error(error:unknown):void;
  composing():boolean; setComposing(value:boolean):void;
  t(key:MessageKey,args?:Readonly<Record<string,string|number>>):string;
  locale():string; isCurrent():boolean; readOnly?:boolean; format?:'xls'|'xlsx'|'oex';
}
interface RunningCalculation { worker:Worker;documentId:string;request:number;revision:number;busy:boolean;state:SheetsViewState;timer?:ReturnType<typeof setTimeout> }
interface RenderedCell {
  node:HTMLDivElement;text:Text;row?:number;column?:number;result?:CalculationCell;revision?:number;
  styleId?:string;directStyle?:CellSnapshot['style'];styleApplied?:boolean;conditionalSignature?:string;
}
let running:RunningCalculation|undefined;
var requestId = 0;
let copied:{nonce:string;documentId:string;contentId:string;sheetId:string;range:CellRange}|undefined;
export function releaseSheetCalculation(documentId?:string):void {
  if (!running || documentId && running.documentId !== documentId) return;
  running.worker.terminate();
  clearTimeout(running.timer);
  running.state.calculating = false;
  running = void 0;
}
function recalculate(host:SheetsHost, state2:SheetsViewState, force = false):void {
  const snapshot = host.snapshot();
  if (snapshot.document.kind !== "sheets") return;
  if (host.readOnly && state2.unsupportedFormulas?.length) {
    releaseSheetCalculation();
    state2.calculation = new Map((state2.cachedValues ?? []).map(([key, value]) => [key, { state: "unsupported", reason: "unsupported-syntax", value }]));
    state2.calculatedRevision = snapshot.revision;
    state2.calculating = false;
    state2.calculationError = false;
    return;
  }
  if (!force && state2.calculatedRevision === snapshot.revision) return;
  if (running?.documentId === snapshot.document.documentId && running.revision === snapshot.revision && running.busy) return;
  if (running && (running.documentId !== snapshot.document.documentId || running.busy || force)) releaseSheetCalculation();
  if (!running) {
    const worker = new Worker(new URL("./sheets-recalc-worker.js", import.meta.url), { type: "module" });
    running = { worker, documentId: snapshot.document.documentId, request: 0, revision: -1, busy: false, state: state2 };
  }
  const current = running, id = ++requestId;
  current.request = id;
  current.revision = snapshot.revision;
  current.busy = true;
  current.state = state2;
  state2.calculating = true;
  state2.calculationError = false;
  current.worker.onmessage = (event) => {
    if (running !== current || event.data.id !== id || event.data.revision !== snapshot.revision) return;
    clearTimeout(current.timer);
    current.busy = false;
    state2.calculating = false;
    if (host.isCurrent() && host.snapshot().revision === snapshot.revision) {
      state2.calculationError = !event.data.ok;
      state2.calculatedRevision = event.data.ok ? snapshot.revision : void 0;
      if (event.data.ok) state2.calculation = new Map(event.data.cells);
    }
    state2.onCalculation?.();
  };
  current.worker.onerror = () => {
    if (running !== current) return;
    releaseSheetCalculation();
    state2.calculationError = true;
    state2.onCalculation?.();
  };
  current.timer = setTimeout(() => {
    if (running !== current || !current.busy) return;
    releaseSheetCalculation();
    state2.calculationError = true;
    state2.onCalculation?.();
  }, 12e3);
  current.worker.postMessage({ id, revision: snapshot.revision, document: snapshot.document, cachedValues: state2.cachedValues, cachedSources: state2.cachedSources });
}
const styleColor = (color:Color|undefined) => color?.kind === "rgba" ? `rgba(${color.red},${color.green},${color.blue},${color.alpha / 255})` : "var(--text)";
const rgb = (value:string):Color => ({ kind: "rgba", red: parseInt(value.slice(1, 3), 16), green: parseInt(value.slice(3, 5), 16), blue: parseInt(value.slice(5, 7), 16), alpha: 255 });
function colorHex(color:Color|undefined, fallback:string):string {
  return color?.kind === "rgba" ? "#" + [color.red, color.green, color.blue].map((x) => x.toString(16).padStart(2, "0")).join("") : fallback;
}
function rangeOf(state2:SheetsViewState):CellRange {
  return { row: Math.min(state2.row, state2.anchorRow), column: Math.min(state2.column, state2.anchorColumn), rowCount: Math.abs(state2.row - state2.anchorRow) + 1, columnCount: Math.abs(state2.column - state2.anchorColumn) + 1 };
}
function contains2(range:CellRange, row:number, column:number):boolean {
  return row >= range.row && row < range.row + range.rowCount && column >= range.column && column < range.column + range.columnCount;
}
export function sheetsWorkspace(host:SheetsHost, state2:SheetsViewState):HTMLElement {
  const snapshot = host.snapshot();
  if (snapshot.document.kind !== "sheets") throw new OfficeError("invalid", "Expected a workbook");
  const model:WorkbookSnapshot = snapshot.document;
  if (state2.cachedValues && !state2.cachedSources) state2.cachedSources = model.sheets.flatMap((s) => Object.entries(s.cells).filter(([, cell]) => cell.value.kind === "formula").map(([key, cell]) => [s.id + "!" + key, cell.value.kind === "formula" ? cell.value.source : ""]));
  const sheet = model.sheets.find((x) => x.id === state2.sheetId) ?? model.sheets[0];
  state2.sheetId = sheet.id;
  const opaque = new Set(state2.unsupportedFormulas ?? []);
  const limits = host.format === "xls" ? { rows: 65536, columns: 256 } : { rows: 1048576, columns: 16384 };
  state2.row = Math.max(1, Math.min(state2.row, limits.rows));
  state2.column = Math.max(1, Math.min(state2.column, limits.columns));
  state2.anchorRow = Math.max(1, Math.min(state2.anchorRow, limits.rows));
  state2.anchorColumn = Math.max(1, Math.min(state2.anchorColumn, limits.columns));
  const t = host.t, root2 = document.createElement("section");
  root2.className = "sheets-workspace";
  root2.id = "sheets-workspace";
  const make = <K extends keyof HTMLElementTagNameMap>(tag:K, value = "", className = ""):HTMLElementTagNameMap[K] => {
    const node = document.createElement(tag);
    node.textContent = value;
    if (className) node.className = className;
    return node;
  };
  const guard = (action:()=>unknown) => {
    try {
      if (!flush()) return;
      action();
    } catch (error) {
      host.error(error);
    }
  };
  const changed = (command:DocumentCommand):boolean => {
    if (host.readOnly) return false;
    const ok = host.execute(command);
    if (ok) {
      state2.calculatedRevision = void 0;
      host.render();
    }
    return ok;
  };
  const button = (key:MessageKey, id:string, action:()=>unknown, disabled = Boolean(host.readOnly)):HTMLButtonElement => {
    const node = make("button", t(key));
    node.type = "button";
    node.id = id;
    node.disabled = disabled;
    node.onclick = () => guard(action);
    return node;
  };
  const label = (key:MessageKey, control:HTMLElement):HTMLLabelElement => {
    const wrapper = make("label");
    wrapper.append(make("span", t(key)), control);
    control.setAttribute("aria-label", t(key));
    return wrapper;
  };
  const toolbar = make("div", "", "sheets-formatbar");
  toolbar.setAttribute("aria-label", t("sheetFormatting"));
  const selectionRange = () => {
    let range = rangeOf(state2), changed2 = true;
    while (changed2) {
      changed2 = false;
      for (const merge of sheet.merges) {
        if (merge.row >= range.row + range.rowCount || range.row >= merge.row + merge.rowCount || merge.column >= range.column + range.columnCount || range.column >= merge.column + merge.columnCount) continue;
        const row = Math.min(range.row, merge.row), column = Math.min(range.column, merge.column), rowEnd = Math.max(range.row + range.rowCount, merge.row + merge.rowCount), columnEnd = Math.max(range.column + range.columnCount, merge.column + merge.columnCount);
        if (row !== range.row || column !== range.column || rowEnd !== range.row + range.rowCount || columnEnd !== range.column + range.columnCount) {
          range = { row, column, rowCount: rowEnd - row, columnCount: columnEnd - column };
          changed2 = true;
        }
      }
    }
    return range;
  };
  const selected = sheet.cells[`${state2.row}:${state2.column}`];
  const currentStyle = ():CellStyle => {
    const cell = sheet.cells[`${state2.row}:${state2.column}`];
    return { ...cell?.styleId ? model.styles[cell.styleId]?.properties : {}, ...cell?.style };
  };
  const selectedStyle = currentStyle();
  for (const [key, property, id] of [["writerBold", "bold", "sheet-bold"], ["writerItalic", "italic", "sheet-italic"], ["writerUnderline", "underline", "sheet-underline"]] as const) {
    const b = button(key, id, () => changed(formatCells(host.snapshot(), sheet.id, selectionRange(), { style: { [property]: !currentStyle()[property] } })));
    b.setAttribute("aria-pressed", String(Boolean(selectedStyle[property])));
    toolbar.append(b);
  }
  const family = make("input");
  family.value = selectedStyle.fontFamily ?? model.defaultTextStyle.fontFamily ?? "Arial";
  family.id = "sheet-font";
  family.maxLength = 128;
  family.disabled = Boolean(host.readOnly);
  family.onchange = () => guard(() => changed(formatCells(host.snapshot(), sheet.id, selectionRange(), { style: { fontFamily: family.value } })));
  toolbar.append(label("writerFontFamily", family));
  const fontSize = make("input");
  fontSize.type = "number";
  fontSize.min = "1";
  fontSize.max = "409";
  fontSize.step = ".5";
  fontSize.value = String((selectedStyle.fontSize ?? model.defaultTextStyle.fontSize ?? 139700) / 12700);
  fontSize.id = "sheet-font-size";
  fontSize.disabled = Boolean(host.readOnly);
  fontSize.onchange = () => guard(() => changed(formatCells(host.snapshot(), sheet.id, selectionRange(), { style: { fontSize: Math.round(Number(fontSize.value) * 12700) } })));
  toolbar.append(label("writerFontSize", fontSize));
  for (const [key, property, id, fallback] of [["writerTextColor", "color", "sheet-text-color", "#222735"], ["sheetFillColor", "fill", "sheet-fill-color", "#ffffff"]] as const) {
    const picker = make("input");
    picker.type = "color";
    picker.id = id;
    picker.value = colorHex(selectedStyle[property], fallback);
    picker.disabled = Boolean(host.readOnly);
    picker.onchange = () => guard(() => changed(formatCells(host.snapshot(), sheet.id, selectionRange(), { style: { [property]: rgb(picker.value) } })));
    toolbar.append(label(key, picker));
  }
  for (const [key, value, id] of [["writerAlignLeft", "start", "sheet-align-start"], ["writerAlignCenter", "center", "sheet-align-center"], ["writerAlignRight", "end", "sheet-align-end"]] as const) toolbar.append(button(key, id, () => changed(formatCells(host.snapshot(), sheet.id, selectionRange(), { style: { alignment: value } }))));
  toolbar.append(button("sheetWrap", "sheet-wrap", () => changed(formatCells(host.snapshot(), sheet.id, selectionRange(), { style: { wrapText: !currentStyle().wrapText } }))));
  toolbar.append(button("sheetBorder", "sheet-border", () => changed(formatCells(host.snapshot(), sheet.id, selectionRange(), { style: { border: { width: 9525, color: { kind: "rgba", red: 110, green: 120, blue: 135, alpha: 255 } } } }))));
  const namedStyle=make('select');namedStyle.id='sheet-named-style';
  const chooseStyle=make('option','—');chooseStyle.value='';namedStyle.append(chooseStyle);
  for(const name of ['Normal',...(model.namedCellStyles??[]).map(s=>s.name)]){const option=make('option',name);option.value=name;namedStyle.append(option);}
  namedStyle.value=selected?.namedStyle??'Normal';
  namedStyle.disabled=!!host.readOnly||host.format==='xls'||!model.namedCellStyles?.length;
  namedStyle.onchange=()=>guard(()=>{if(namedStyle.value)changed(applyNamedCellStyle(host.snapshot(),sheet.id,selectionRange(),namedStyle.value));});
  if(model.namedCellStyles?.length)toolbar.append(label('ribbonStyles',namedStyle));
  root2.append(toolbar);
  const second = make("div", "", "sheets-formatbar");
  const formats = make("select");
  formats.id = "sheet-number-format";
  formats.disabled = Boolean(host.readOnly);
  for (const [key, value] of [["sheetGeneral", "General"], ["sheetNumber", "0.00"], ["sheetPercent", "0.00%"], ["sheetCurrency", "$#,##0.00"], ["sheetDate", "yyyy-mm-dd"], ["sheetTime", "hh:mm:ss"], ["sheetScientific", "0.00E+00"]] as const) {
    const option = make("option", t(key));
    option.value = value;
    formats.append(option);
  }
  if (selected?.numberFormat && !Array.from(formats.options).some((x) => x.value === selected.numberFormat)) {
    const option = make("option", selected.numberFormat);
    option.value = selected.numberFormat;
    formats.append(option);
  }
  formats.value = selected?.numberFormat ?? "General";
  formats.onchange = () => guard(() => changed(formatCells(host.snapshot(), sheet.id, selectionRange(), { numberFormat: formats.value })));
  second.append(label("sheetNumberFormat", formats));
  const custom = make("input");
  custom.id = "sheet-custom-format";
  custom.value = selected?.numberFormat ?? "General";
  custom.maxLength = 1024;
  custom.disabled = Boolean(host.readOnly);
  custom.onchange = () => guard(() => changed(formatCells(host.snapshot(), sheet.id, selectionRange(), { numberFormat: custom.value })));
  second.append(label("sheetCustomFormat", custom));
  const mode = make("select");
  mode.id = "sheet-formula-mode";
  for (const [value, key] of [["canonical", "sheetCanonical"], ["regional", "sheetRegional"]] as const) {
    const o = make("option", t(key));
    o.value = value;
    mode.append(o);
  }
  mode.value = state2.inputMode;
  mode.onchange = () => {
    if (!flush()) {
      mode.value = state2.inputMode;
      return;
    }
    state2.inputMode = mode.value === "regional" ? "regional" : "canonical";
    host.render();
  };
  second.append(label("sheetFormulaMode", mode));
  second.append(button("sheetMerge", "sheet-merge", () => changed(mergeCells(host.snapshot(), sheet.id, selectionRange()))), button("sheetUnmerge", "sheet-unmerge", () => changed(unmergeCells(host.snapshot(), sheet.id, selectionRange()))));
  second.append(button("sheetClear", "sheet-clear", () => changed(clearCells(host.snapshot(), sheet.id, selectionRange()))), button("sheetFillDown", "sheet-fill-down", () => changed(fillCells(host.snapshot(), sheet.id, selectionRange(), "down"))), button("sheetFillRight", "sheet-fill-right", () => changed(fillCells(host.snapshot(), sheet.id, selectionRange(), "right"))));
  root2.append(second);
  const conditionalDialog=make('dialog');conditionalDialog.id='sheet-conditional-dialog';
  const conditionalTitle=make('header');conditionalTitle.append(make('h2',t('sheetConditionalRules')),button('close','sheet-conditional-close',()=>conditionalDialog.close(),false));conditionalDialog.append(conditionalTitle);
  const ruleList=make('select');ruleList.id='sheet-conditional-rules';
  const sortedRules=[...(sheet.conditionalRules??[])].sort((a,b)=>a.priority-b.priority);
  for(const rule of sortedRules){const option=make('option',rule.priority+' · '+rule.ranges.map(r=>cellLabel(r.row,r.column)+':'+cellLabel(r.row+r.rowCount-1,r.column+r.columnCount-1)).join(', '));option.value=String(rule.priority);ruleList.append(option);}
  const ruleDelete=button('sheetConditionalRemove','sheet-conditional-remove',()=>{if(changed(removeConditionalRule(host.snapshot(),sheet.id,Number(ruleList.value))))conditionalDialog.close();},!!host.readOnly||!sortedRules.length||host.format==='xls');
  conditionalDialog.append(label('sheetConditionalRules',ruleList),ruleDelete);
  const op=make('select');op.id='sheet-conditional-operator';
  for(const [value,symbol]of [['greaterThan','>'],['greaterThanOrEqual','≥'],['lessThan','<'],['lessThanOrEqual','≤'],['equal','='],['notEqual','≠'],['between','…'],['notBetween','¬ …']] as const){const option=make('option',symbol);option.value=value;op.append(option);}
  const threshold=make('input'),upper=make('input');for(const [input,id]of [[threshold,'sheet-conditional-value'],[upper,'sheet-conditional-upper']] as const){input.type='number';input.step='any';input.value='0';input.id=id;}
  const upperLabel=label('sheetConditionalUpper',upper);upperLabel.hidden=true;
  op.onchange=()=>{upperLabel.hidden=op.value!=='between'&&op.value!=='notBetween';};
  const highlight=make('input');highlight.type='color';highlight.value='#ffdca0';highlight.id='sheet-conditional-color';
  const stop=make('input');stop.type='checkbox';stop.id='sheet-conditional-stop';
  conditionalDialog.append(label('sheetConditionalOperator',op),label('sheetConditionalValue',threshold),upperLabel,label('sheetFillColor',highlight),label('sheetConditionalStop',stop));
  const addRule=button('sheetConditionalAdd','sheet-conditional-add',()=>{
    const operator=op.value as ConditionalOperator,values=upperLabel.hidden?[threshold.valueAsNumber]:[threshold.valueAsNumber,upper.valueAsNumber];
    if(!values.every(Number.isFinite))throw new OfficeError('invalid','Invalid numeric conditional threshold');
    if(changed(addConditionalRule(host.snapshot(),sheet.id,{ranges:[selectionRange()],operator,values,style:{fill:rgb(highlight.value)},...(stop.checked?{stopIfTrue:true}:{})})))conditionalDialog.close();
  },!!host.readOnly||host.format==='xls');conditionalDialog.append(addRule);root2.append(conditionalDialog);
  const chartDialog=make('dialog');chartDialog.id='sheet-chart-dialog';
  const chartHead=make('header');chartHead.append(make('h2',t('sheetCharts')),button('close','sheet-chart-close',()=>chartDialog.close(),false));chartDialog.append(chartHead);
  const chartList=make('select');chartList.id='sheet-chart-list';
  const chartType=make('select');chartType.id='sheet-chart-type';
  for(const [value,key]of [['column','sheetChartColumn'],['line','sheetChartLine'],['pie','sheetChartPie']] as const){const o=make('option',t(key));o.value=value;chartType.append(o);}
  for(const [i,chart]of (sheet.charts??[]).entries()){const o=make('option',(i+1)+' · '+t(chart.kind==='column'?'sheetChartColumn':chart.kind==='line'?'sheetChartLine':'sheetChartPie')+' · '+cellLabel(chart.anchor.row,chart.anchor.column));o.value=String(i);chartList.append(o);}
  chartDialog.append(label('sheetCharts',chartList),button('sheetChartRemove','sheet-chart-remove',()=>{if(changed(removeSheetChart(host.snapshot(),sheet.id,Number(chartList.value))))chartDialog.close();},!!host.readOnly||!sheet.charts?.length||host.format==='xls'));
  chartDialog.append(make('p',t('sheetChartSourceNote')),label('sheetCharts',chartType),button('sheetChartAdd','sheet-chart-add',()=>{if(changed(addSheetChart(host.snapshot(),sheet.id,selectionRange(),chartType.value as SheetChart['kind'])))chartDialog.close();},!!host.readOnly||host.format==='xls'));
  root2.append(chartDialog);
  toolbar.append(button('sheetCharts','sheet-chart-open',()=>chartDialog.showModal(),!!host.readOnly||host.format==='xls'));
  toolbar.append(button('sheetConditionalRules','sheet-conditional-open',()=>conditionalDialog.showModal(),!!host.readOnly||host.format==='xls'));
  const controls2 = make("details", "", "sheets-advanced");
  controls2.append(make("summary", t("sheetManage")));
  const management = make("div", "", "sheets-formatbar");
  const sheetName = make("input");
  sheetName.id = "sheet-name";
  sheetName.value = sheet.name;
  sheetName.maxLength = 31;
  sheetName.disabled = Boolean(host.readOnly);
  management.append(label("sheetName", sheetName));
  management.append(button("sheetRename", "sheet-rename", () => changed(renameSheet(host.snapshot(), sheet.id, sheetName.value))), button("sheetDelete", "sheet-delete", () => {
    state2.row = state2.anchorRow = state2.column = state2.anchorColumn = 1;
    changed(deleteSheet(host.snapshot(), sheet.id));
  }), button("sheetMoveLeft", "sheet-move-left", () => changed(moveSheet(host.snapshot(), sheet.id, Math.max(0, model.sheets.findIndex((s) => s.id === sheet.id) - 1)))), button("sheetMoveRight", "sheet-move-right", () => changed(moveSheet(host.snapshot(), sheet.id, Math.min(model.sheets.length - 1, model.sheets.findIndex((s) => s.id === sheet.id) + 1)))));
  for (const [key, id, fn, axis] of [["sheetInsertRows", "sheet-insert-rows", insertRows, "rows"], ["sheetDeleteRows", "sheet-delete-rows", deleteRows, "rows"], ["sheetInsertColumns", "sheet-insert-columns", insertColumns, "columns"], ["sheetDeleteColumns", "sheet-delete-columns", deleteColumns, "columns"]] as const) management.append(button(key, id, () => {
    const range = selectionRange();
    changed(fn(host.snapshot(), sheet.id, axis === "rows" ? range.row : range.column, axis === "rows" ? range.rowCount : range.columnCount));
  }));
  for (const [key, id, axis, size, defaultPx] of [["sheetRowHeight", "sheet-row-height", "rows", "height", 24], ["sheetColumnWidth", "sheet-column-width", "columns", "width", 104]] as const) {
    const input = make("input");
    input.type = "number";
    input.id = id;
    input.min = "2";
    input.max = "2000";
    input.step = "1";
    const current:{height?:number;width?:number}|undefined = axis === "rows" ? sheet.rows[state2.row] : sheet.columns[state2.column];
    const sourceDefault=axis==='rows'?sheet.defaultGeometry?.rowHeight:sheet.defaultGeometry?.columnWidth;
    input.value = String(Math.round((current?.[size] ?? sourceDefault ?? defaultPx * 9525) / 9525));
    input.disabled = Boolean(host.readOnly);
    input.onchange = () => guard(() => {
      const range = selectionRange();
      changed(setDimensions(host.snapshot(), sheet.id, axis, axis === "rows" ? range.row : range.column, axis === "rows" ? range.rowCount : range.columnCount, { [size]: Math.round(Number(input.value) * EMU_PER_CSS_PIXEL) }));
    });
    management.append(label(key, input));
  }
  const dates = make("select");
  dates.id = "sheet-date-system";
  dates.disabled = Boolean(host.readOnly);
  for (const system of ["1900", "1904"]) {
    const o = make("option", system);
    o.value = system;
    dates.append(o);
  }
  dates.value = model.dateSystem;
  dates.onchange = () => guard(() => changed(patchCommand(host.snapshot(), "SHT-02", [{ op: "set", path: ["dateSystem"], value: dates.value }])));
  management.append(label("sheetDateSystem", dates));
  const updateView = (patch:SheetView) => {
    const snap = host.snapshot();
    if (snap.document.kind !== "sheets") return;
    const index = snap.document.sheets.findIndex((s) => s.id === sheet.id);
    changed(patchCommand(snap, "SHT-02", [{ op: "set", path: ["sheets", index, "view"], value: { ...sheet.view, ...patch } }]));
  };
  management.append(button("sheetFreeze", "sheet-freeze", () => updateView({ frozenRows: state2.row - 1, frozenColumns: state2.column - 1 })), button("sheetUnfreeze", "sheet-unfreeze", () => updateView({ frozenRows: 0, frozenColumns: 0 })), button("sheetRtl", "sheet-rtl", () => updateView({ rightToLeft: !sheet.view?.rightToLeft })), button("sheetShowGrid", "sheet-grid-lines", () => updateView({ showGridLines: sheet.view?.showGridLines === false })));
  controls2.append(management);
  root2.append(controls2);
  const bar = make("div", "", "sheets-formula-bar");
  const address = make("input");
  address.id = "sheets-address";
  address.setAttribute("aria-label", t("sheetAddress"));
  address.spellcheck = false;
  const formula = make("input");
  formula.id = "sheets-formula";
  formula.setAttribute("aria-label", t("sheetFormula"));
  formula.dir = "auto";
  formula.spellcheck = false;
  formula.autocomplete = "off";
  formula.readOnly = Boolean(host.readOnly);
  bar.append(address, make("span", "\u0192x", "sheets-fx"), formula);
  root2.append(bar);
  const scroll = make("div", "", "sheets-scroll");
  scroll.id = "sheets-scroll";
  scroll.dataset.rtl = String(Boolean(sheet.view?.rightToLeft));
  scroll.tabIndex = 0;
  scroll.dir = "ltr";
  scroll.setAttribute("role", "grid");
  scroll.setAttribute("aria-label", t("sheetGrid"));
  scroll.setAttribute("aria-rowcount", String(limits.rows));
  scroll.setAttribute("aria-colcount", String(limits.columns));
  scroll.setAttribute("aria-multiselectable", "true");
  scroll.setAttribute("aria-readonly", String(Boolean(host.readOnly)));
  const plane = make("div", "", "sheets-plane");
  scroll.append(plane);
  root2.append(scroll);
  const rows = new GridAxis(limits.rows, (sheet.defaultGeometry?.rowHeight??24*9525)/9525, sheet.rows, "height"), columns = new GridAxis(limits.columns, (sheet.defaultGeometry?.columnWidth??104*9525)/9525, sheet.columns, "width");
  const headerWidth = 46, headerHeight = 26;
  const totalHeight = Math.min(3e7, rows.total + headerHeight), totalWidth = Math.min(3e7, columns.total + headerWidth);
  plane.style.height = totalHeight + "px";
  plane.style.width = totalWidth + "px";
  const displayRowScale = (rows.total + headerHeight) / totalHeight, displayColScale = (columns.total + headerWidth) / totalWidth;
  const layer = make("div", "", "sheets-cell-layer");
  // The scroll plane establishes the logical scrollbar range; the painted
  // surface itself is bounded to the actual viewport. Keep editor coordinates
  // on the plane and translate only the recycled drawing layer.
  Object.assign(layer.style,{inset:"auto",overflow:"hidden",contain:"layout style paint"});
  plane.append(layer);
  const chartLayer=make('div','','sheets-chart-layer');Object.assign(chartLayer.style,{position:'absolute',left:'0',top:'0',pointerEvents:'none',overflow:'hidden',zIndex:'3'});plane.append(chartLayer);
  let chartRevision:string|undefined;
  const chartNodes=new Map<number,HTMLDivElement>();
  const svgNS='http://www.w3.org/2000/svg';
  function chartElement(chart:DeepReadonly<SheetChart>,revision:number|undefined):HTMLDivElement{
    const box=make('div','','sheets-chart');box.dataset.chartType=chart.kind;
    Object.assign(box.style,{position:'absolute',background:'#ffffff',color:'#20252d',border:'1px solid #a8b2bf',overflow:'hidden',boxSizing:'border-box'});
    const get=(row:number,col:number):CalculatedValue|undefined=>{
      const cell=sheet.cells[row+':'+col];if(!cell)return undefined;
      if(cell.value.kind!=='formula')return cell.value;
      const r=revision===snapshot.revision?state2.calculation?.get(calculationKey(sheet.id,row,col)):undefined;
      return r?.state==='valid'?r.value:undefined;
    };
    const labelValue=(row:number,col:number)=>{const v=get(row,col);return v&&'value'in v?String(v.value):'';};
    const points=Math.min(200,chart.categories.rowCount),indices=Array.from({length:points},(_,i)=>i).filter(i=>!sheet.rows[chart.categories.row+i]?.hidden);
    const names=indices.map(i=>labelValue(chart.categories.row+i,chart.categories.column));
    const series=chart.series.filter(s=>!sheet.columns[s.values.column]?.hidden).map((s,i)=>({label:s.label?labelValue(s.label.row,s.label.column):columnLabel(s.values.column),values:indices.map(i=>{const v=get(s.values.row+i,s.values.column);return v?.kind==='number'?v.value:null;}),color:['#4679bc','#df8844','#60a276','#a56bb0','#d4aa38','#4c9da7','#bb6064','#71859c'][i%8]!}));
    const svg=document.createElementNS(svgNS,'svg');svg.setAttribute('viewBox','0 0 560 280');svg.setAttribute('width','100%');svg.setAttribute('height','100%');svg.setAttribute('role','img');svg.setAttribute('aria-label',t(chart.kind==='column'?'sheetChartColumn':chart.kind==='line'?'sheetChartLine':'sheetChartPie'));box.append(svg);
    const n=(tag:string,attrs:Record<string,string|number>,text?:string)=>{const e=document.createElementNS(svgNS,tag);for(const [k,v]of Object.entries(attrs))e.setAttribute(k,String(v));if(text!==undefined)e.textContent=text;svg.append(e);return e;};
    const label=(x:number,y:number,text:string,anchor='start')=>n('text',{x,y,fill:'#334155','font-size':11,'text-anchor':anchor},text.slice(0,45));
    const vals=series.flatMap(s=>s.values.filter((v):v is number=>v!==null));
    if(!vals.length||chart.kind==='pie'&&(vals.some(v=>v<0)||!vals.some(v=>v>0))){label(280,140,t('sheetChartNoData'),'middle');return box;}
    if(chart.categories.rowCount>points)label(10,15,t('sheetChartLimited'));
    if(chart.kind==='pie'){
      const s=series[0]!;const total=s.values.reduce<number>((sum,v)=>sum+(v??0),0);
      if(!Number.isFinite(total)||total<=0){label(280,140,t('sheetChartNoData'),'middle');return box;}
      let angle=-Math.PI/2;
      for(const [i,v]of s.values.entries()){
        if(!v)continue;const delta=v/total*2*Math.PI,end=angle+delta,color=['#4679bc','#df8844','#60a276','#a56bb0','#d4aa38','#4c9da7','#bb6064','#71859c'][i%8]!;
        if(delta>=2*Math.PI-1e-10)n('circle',{cx:180,cy:140,r:108,fill:color});
        else n('path',{d:'M 180 140 L '+(180+108*Math.cos(angle))+' '+(140+108*Math.sin(angle))+' A 108 108 0 '+(delta>Math.PI?1:0)+' 1 '+(180+108*Math.cos(end))+' '+(140+108*Math.sin(end))+' Z',fill:color,stroke:'#fff','stroke-width':1});
        if(i<11){n('rect',{x:315,y:30+i*20,width:10,height:10,fill:color});label(331,39+i*20,(names[i]||String(i+1))+' '+(v/total*100).toFixed(1)+'%');}
        angle=end;
      }
    }else{
      const left=64,right=535,top=30,bottom=220,w=right-left,h=bottom-top;
      // Normalize first, so opposite large finite inputs do not overflow.
      const scale=Math.max(1,...vals.map(Math.abs)),lo=Math.min(0,...vals.map(v=>v/scale)),hi=Math.max(0,...vals.map(v=>v/scale)),span=hi-lo||1,y=(v:number)=>bottom-(v/scale-lo)/span*h;
      const baseline=y(0),step=w/Math.max(1,names.length);
      for(let i=0;i<=4;i++){const yy=top+h*i/4;n('line',{x1:left,y1:yy,x2:right,y2:yy,stroke:'#e5e7eb'});const number=(hi-(hi-lo)*i/4)*scale;label(left-6,yy+4,Number.isFinite(number)?String(Number(number.toPrecision(4))):'','end');}
      n('line',{x1:left,y1:baseline,x2:right,y2:baseline,stroke:'#526174'});
      const stride=Math.max(1,Math.ceil(names.length/8));
      for(const [i,name]of names.entries())if(i%stride===0)label(left+(i+.5)*step,240,name||String(i+1),'middle');
      for(const [j,s]of series.entries()){
        if(chart.kind==='column'){
          const width=Math.max(.4,step*.72/series.length);
          for(const [i,v]of s.values.entries())if(v!==null)n('rect',{x:left+i*step+step*.14+j*width,y:Math.min(baseline,y(v)),width,height:Math.max(.5,Math.abs(y(v)-baseline)),fill:s.color});
        }else{
          let path='',started=false;for(const [i,v]of s.values.entries()){if(v===null){started=false;continue;}path+=(started?' L ':' M ')+(left+(i+.5)*step)+' '+y(v);started=true;}n('path',{d:path,stroke:s.color,'stroke-width':2,fill:'none'});
          if(s.values.filter(v=>v!==null).length===1){const i=s.values.findIndex(v=>v!==null);n('circle',{cx:left+(i+.5)*step,cy:y(s.values[i]!),r:3,fill:s.color});}
        }
        if(j<4){n('rect',{x:left+j*115,y:260,width:9,height:9,fill:s.color});label(left+14+j*115,269,s.label.slice(0,18));}
      }
    }
    return box;
  }
  const renderedCells = new Map<string,RenderedCell>();
  const rowHeaders = new Map<number,HTMLDivElement>(), columnHeaders = new Map<number,HTMLDivElement>();
  const corner = make("div", "", "sheets-corner");
  corner.style.width = headerWidth + "px";
  corner.style.height = headerHeight + "px";
  layer.append(corner);
  const editor = make("input");
  editor.id = "sheets-cell-editor";
  editor.className = "sheets-cell-editor";
  editor.hidden = true;
  editor.dir = "auto";
  editor.spellcheck = false;
  editor.autocomplete = "off";
  plane.append(editor);
  const bottom = make("div", "", "sheets-bottom");
  const sheetTabs = make("div", "", "sheet-tabs");
  sheetTabs.setAttribute("role", "tablist");
  for (const current of model.sheets) {
    const b = make("button", current.name);
    b.type = "button";
    b.dataset.sheetId = current.id;
    b.setAttribute("role", "tab");
    b.setAttribute("aria-selected", String(current.id === sheet.id));
    b.onclick = () => {
      if (host.composing() || !flush()) return;
      state2.sheetId = current.id;
      state2.row = state2.column = state2.anchorRow = state2.anchorColumn = 1;
      state2.scrollLeft = state2.scrollTop = 0;
      host.render();
    };
    sheetTabs.append(b);
  }
  const add = button("sheetAdd", "sheet-add", () => {
    let number = model.sheets.length + 1;
    let name = "Sheet" + number;
    while (model.sheets.some((x) => x.name.toLowerCase() === name.toLowerCase())) name = "Sheet" + ++number;
    const id = crypto.randomUUID();
    state2.sheetId = id;
    state2.row = state2.column = state2.anchorRow = state2.anchorColumn = 1;
    state2.scrollTop = state2.scrollLeft = 0;
    changed(addSheet(host.snapshot(), name, id));
  });
  sheetTabs.append(add);
  bottom.append(sheetTabs);
  root2.append(bottom);
  const status = make("div", "", "sheets-status");
  status.setAttribute("aria-live", "polite");
  const selectedSummary = make("span"), calcStatus = make("span");
  status.append(selectedSummary, calcStatus, button("sheetRecalculate", "sheet-recalculate", () => {
    state2.calculatedRevision = void 0;
    recalculate(host, state2, true);
    paint();
  }, false), button("sheetCancelCalculation", "sheet-cancel-calculation", () => {
    releaseSheetCalculation();
    state2.calculationError = true;
    paint();
  }, false));
  root2.append(status, make("p", t(host.readOnly ? "sheetSourceProtected" : "sheetHelp"), "muted sheets-help"));
  let editing = false, formulaDirty = false, addressDirty = false, frame = 0, dragging = false, summaryKey = "", controlsKey = "";
  const options = ():CellInputOptions => ({ mode: state2.inputMode, locale: host.locale(), dateSystem: model.dateSystem });
  const rangeAddress = () => {
    const range = selectionRange(), first = cellLabel(range.row, range.column);
    return range.rowCount === 1 && range.columnCount === 1 ? first : first + ":" + cellLabel(range.row + range.rowCount - 1, range.column + range.columnCount - 1);
  };
  const cellSource = () => {
    const value = sheet.cells[`${state2.row}:${state2.column}`]?.value;
    if (!value || value.kind === "blank") return "";
    if (value.kind === "string") return "'" + value.value;
    if (value.kind === "formula") {
      if (state2.inputMode === "regional") {
        try {
          return localizeFormula(value.source, host.locale());
        } catch {
          return value.source;
        }
      }
      return value.source;
    }
    if (value.kind === "number" && state2.inputMode === "regional") return String(value.value).replace(".", formulaSeparators(host.locale()).decimal);
    return String(value.value);
  };
  const activeMerge = () => sheet.merges.find((m) => contains2(m, state2.row, state2.column));
  function applyInput(value:string, advance?:string):boolean {
    if (host.readOnly) return true;
    try {
      const old = sheet.cells[`${state2.row}:${state2.column}`], isDate = Boolean(old?.numberFormat && /[ymd]/i.test(old.numberFormat) && /^\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2})?)?$/.test(value));
      const unchanged = value === cellSource();
      const cell = unchanged ? void 0 : cellFromInput(value, { ...options(), ...isDate ? { type: "date" } : {} }, old);
      if (cell && !host.execute(setCells(host.snapshot(), sheet.id, [{ row: state2.row, column: state2.column, cell }]))) return false;
      editing = false;
      formulaDirty = false;
      editor.hidden = true;
      if (!unchanged) state2.calculatedRevision = void 0;
      if (advance) {
        if (advance === "Enter") state2.row = Math.min(limits.rows, state2.row + (activeMerge()?.rowCount ?? 1));
        else state2.column = Math.max(1, Math.min(limits.columns, state2.column + (advance === "BackTab" ? -1 : activeMerge()?.columnCount ?? 1)));
        state2.anchorRow = state2.row;
        state2.anchorColumn = state2.column;
      }
      host.render();
      queueMicrotask(() => {
        if (host.isCurrent()) document.querySelector<HTMLElement>("#sheets-scroll")?.focus();
      });
      return true;
    } catch (error) {
      host.error(error);
      return false;
    }
  }
  function flush() {
    if (host.composing()) return false;
    const input = editing ? editor : formulaDirty ? formula : void 0;
    if (!input) return true;
    const accepted = applyInput(input.value);
    if (!accepted) input.focus();
    return accepted;
  }
  state2.flush = flush;
  function beginEdit(initial?:string):void {
    if (host.readOnly) return;
    editing = true;
    formulaDirty = false;
    editor.value = initial ?? cellSource();
    editor.hidden = false;
    editor.setAttribute("aria-label", cellLabel(state2.row, state2.column));
    positionEditor();
    editor.focus();
    if (initial === void 0) editor.select();
    else editor.setSelectionRange(editor.value.length, editor.value.length);
  }
  function positionEditor() {
    const merge = activeMerge(), row = merge?.row ?? state2.row, col = merge?.column ?? state2.column;
    const x = screenX(col), y = screenY(row);
    editor.style.left = x + "px";
    editor.style.top = y + "px";
    editor.style.width = Math.max(80, columns.start(col + (merge?.columnCount ?? 1)) - columns.start(col)) + "px";
    editor.style.height = Math.max(24, rows.start(row + (merge?.rowCount ?? 1)) - rows.start(row)) + "px";
  }
  function scrollToCell() {
    const top = rows.start(state2.row), left = columns.start(state2.column), height = scroll.clientHeight - headerHeight, width = scroll.clientWidth - headerWidth;
    const y = scroll.scrollTop * displayRowScale, x = scroll.scrollLeft * displayColScale;
    if (top < y) scroll.scrollTop = top / displayRowScale;
    else if (top + rows.size(state2.row) > y + height) scroll.scrollTop = (top + rows.size(state2.row) - height) / displayRowScale;
    if (left < x) scroll.scrollLeft = left / displayColScale;
    else if (left + columns.size(state2.column) > x + width) scroll.scrollLeft = (left + columns.size(state2.column) - width) / displayColScale;
    state2.scrollTop = scroll.scrollTop;
    state2.scrollLeft = scroll.scrollLeft;
  }
  function select(row:number, col:number, extend = false, ensure = true):void {
    state2.row = Math.max(1, Math.min(row, limits.rows));
    state2.column = Math.max(1, Math.min(col, limits.columns));
    const merged = sheet.merges.find((m) => contains2(m, state2.row, state2.column));
    if (merged) {
      state2.row = merged.row;
      state2.column = merged.column;
    }
    if (!extend) {
      state2.anchorRow = state2.row;
      state2.anchorColumn = state2.column;
    }
    if (ensure) scrollToCell();
    paint();
  }
  function screenX(col:number, offset = scroll.scrollLeft):number {
    const base = columns.start(col);
    const fixed = col <= (sheet.view?.frozenColumns ?? 0);
    return headerWidth + (fixed ? base + offset : base - offset * (displayColScale - 1));
  }
  function screenY(row:number, offset = scroll.scrollTop):number {
    const base = rows.start(row);
    const fixed = row <= (sheet.view?.frozenRows ?? 0);
    return headerHeight + (fixed ? base + offset : base - offset * (displayRowScale - 1));
  }
  // The workbook snapshot is immutable within this mount. A recycled cell
  // can retain its visual CSS only when the named and direct style match.
  const styleProbe=document.createElement('div');
  const namedCss=new Map<string|undefined,string>();
  const directCss=new WeakMap<object,Map<string|undefined,string>>();
  function visualCss(cell:CellSnapshot|undefined):string {
    const named=cell?.styleId,direct=cell?.style;
    let cache:Map<string|undefined,string>;
    if(direct){cache=directCss.get(direct)??new Map();directCss.set(direct,cache);}
    else cache=namedCss;
    const old=cache.get(named);if(old!==undefined)return old;
    const style={...model.defaultTextStyle,...(named?model.styles[named]?.properties:{}),...direct};
    const css=styleProbe.style;css.cssText='';
    css.fontFamily=style.fontFamily??"sans-serif";
    css.fontSize=Math.min(200,(style.fontSize??139700)/9525)+"px";
    css.fontWeight=style.bold?"700":"400";
    css.fontStyle=style.italic?"italic":"normal";
    css.textDecoration=style.underline?"underline":"none";
    css.color=style.color?styleColor(style.color):"var(--text)";
    if(style.fill)css.backgroundColor=styleColor(style.fill);
    if(style.alignment)css.textAlign=style.alignment;
    if(style.wrapText){css.whiteSpace="pre-wrap";css.overflowWrap="anywhere";}
    if(style.border)css.border=Math.max(.5,style.border.width/9525)+"px solid "+styleColor(style.border.color);
    if(style.borderEdges)for(const [side,property]of [['left','borderLeft'],['right','borderRight'],['top','borderTop'],['bottom','borderBottom']] as const){
      const edge=style.borderEdges[side];if(!edge)continue;
      const width=edge.line==='hair'?.5:edge.line==='double'?3:edge.line==='thick'?4:edge.line.startsWith('medium')?2.67:1;
      // CSS approximates dash-dot strokes; the actual XLSX line type
      // and color are retained in the semantic model on save.
      const line=edge.line==='double'?'double':edge.line==='dotted'?'dotted':/dash/i.test(edge.line)?'dashed':'solid';
      css[property]=width+'px '+line+' '+styleColor(edge.color);
    }
    css.alignContent=style.verticalAlignment==="top"?"start":style.verticalAlignment==="center"?"center":"end";
    const result=css.cssText;cache.set(named,result);return result;
  }
  function paint() {
    if (!root2.isConnected) return;
    const scrollTop = scroll.scrollTop, scrollLeft = scroll.scrollLeft, height = scroll.clientHeight, width = scroll.clientWidth;
    if(sheet.charts?.length){
      Object.assign(chartLayer.style,{left:scrollLeft+'px',top:scrollTop+'px',width:width+'px',height:height+'px'});
      const signature=String(state2.calculatedRevision)+':'+String(state2.calculation?.size);
      if(signature!==chartRevision){chartLayer.replaceChildren();chartNodes.clear();chartRevision=signature;}
      for(const [i,chart]of sheet.charts.entries()){
        const a=chart.anchor,x=screenX(a.column,scrollLeft)-scrollLeft+a.offsetX/9525,y=screenY(a.row,scrollTop)-scrollTop+a.offsetY/9525,w=a.width/9525,h=a.height/9525;
        let node=chartNodes.get(i);
        if(x+w<=0||y+h<=0||x>=width||y>=height){if(node){node.remove();chartNodes.delete(i);}continue;}
        if(!node){node=chartElement(chart,state2.calculatedRevision);chartNodes.set(i,node);chartLayer.append(node);}
        Object.assign(node.style,{left:x+'px',top:y+'px',width:w+'px',height:h+'px'});
      }
    }
    const y = scrollTop * displayRowScale, x = scrollLeft * displayColScale;
    const layerLeft=scrollLeft+"px",layerTop=scrollTop+"px",layerWidth=width+"px",layerHeight=height+"px";
    if(layer.style.left!==layerLeft)layer.style.left=layerLeft;
    if(layer.style.top!==layerTop)layer.style.top=layerTop;
    if(layer.style.width!==layerWidth)layer.style.width=layerWidth;
    if(layer.style.height!==layerHeight)layer.style.height=layerHeight;
    const frozenRows = sheet.view?.frozenRows ?? 0, frozenColumns = sheet.view?.frozenColumns ?? 0;
    const fixedRows = rows.visible(0, Math.min(height / 2, 200), 40).filter((r) => r <= frozenRows), fixedColumns = columns.visible(0, Math.min(width / 2, 400), 20).filter((c) => c <= frozenColumns);
    const visibleRows = [.../* @__PURE__ */ new Set([...fixedRows, ...rows.visible(y, height, 160)])].filter((r) => rows.size(r) > 0);
    const visibleColumns = [.../* @__PURE__ */ new Set([...fixedColumns, ...columns.visible(x, width, 60)])].filter((c) => columns.size(c) > 0);
    const fragment = document.createDocumentFragment(), range = selectionRange();
    let count = 0, sum = 0;
    const activeId = "sheet-cell-" + state2.row + "-" + state2.column;
    // First plan only unique visible anchors. This lets outgoing nodes be
    // rebound in place rather than creating a new DOM tree on fast scroll.
    const planned=new Map<string,{row:number;column:number;merge?:CellRange}>();
    const visibleMerges=sheet.merges.filter(m=>visibleRows.some(r=>r>=m.row&&r<m.row+m.rowCount)&&visibleColumns.some(c=>c>=m.column&&c<m.column+m.columnCount));
    for(const row of visibleRows)for(const column of visibleColumns){
      const merge=visibleMerges.find(m=>contains2(m,row,column)),r=merge?.row??row,c=merge?.column??column,key=r+":"+c;
      if(!planned.has(key))planned.set(key,{row:r,column:c,...(merge?{merge}:{})});
    }
    const recycled=new Set<RenderedCell>(),byColumn=new Map<number,RenderedCell[]>();
    for(const [key,rendered]of renderedCells)if(!planned.has(key)){
      renderedCells.delete(key);recycled.add(rendered);
      if(rendered.column!==undefined){let pool=byColumn.get(rendered.column);if(!pool){pool=[];byColumn.set(rendered.column,pool);}pool.push(rendered);}
    }
    // Reuse within the same column first: horizontal geometry and colindex
    // remain unchanged for the common vertical-scroll case.
    function takeCell(column:number):RenderedCell|undefined{
      const pool=byColumn.get(column);while(pool?.length){const cell=pool.shift()!;if(recycled.delete(cell))return cell;}
      const cell=recycled.values().next().value as RenderedCell|undefined;if(cell)recycled.delete(cell);return cell;
    }
    const locale=host.locale(),formatOptions={locale,dateSystem:model.dateSystem},currentRevision=state2.calculatedRevision;
    const rowPosition=new Map<number,number>(),columnPosition=new Map<number,number>();
    for(const row of visibleRows)rowPosition.set(row,screenY(row,scrollTop)-scrollTop);
    for(const col of visibleColumns)columnPosition.set(col,screenX(col,scrollLeft)-scrollLeft);
    for(const [key,{row,column,merge}]of planned){
      const cell=sheet.cells[key],needsFormula=cell?.value.kind==="formula",formulaKey=needsFormula||opaque.size?calculationKey(sheet.id,row,column):"",isOpaque=opaque.has(formulaKey),result=needsFormula||isOpaque?state2.calculation?.get(formulaKey):undefined;
      let rendered=renderedCells.get(key);const rebound=!rendered;
      if(!rendered){
        rendered=takeCell(column);
        if(!rendered){
          const node=make("div","","sheets-cell"),text=document.createTextNode("");
          node.append(text);node.setAttribute("role","gridcell");node.dir="auto";fragment.append(node);rendered={node,text};
        }
        renderedCells.set(key,rendered);
        const node=rendered.node,addressLabel=cellLabel(row,column);
        node.id="sheet-cell-"+row+"-"+column;
        if(rendered.row!==row){const value=String(row);node.dataset.row=value;node.setAttribute("aria-rowindex",value);rendered.row=row;}
        if(rendered.column!==column){const value=String(column);node.dataset.column=value;node.setAttribute("aria-colindex",value);rendered.column=column;}
        node.dataset.address=addressLabel;node.setAttribute("aria-label",addressLabel);
        if(merge){node.setAttribute("aria-rowspan",String(merge.rowCount));node.setAttribute("aria-colspan",String(merge.columnCount));}
        else if(node.hasAttribute("aria-rowspan")){node.removeAttribute("aria-rowspan");node.removeAttribute("aria-colspan");}
        // Reset geometry, style and transient classes on every different key.
        // Setting cssText clears an old background, border, wrap or alignment.
        if(!rendered.styleApplied||rendered.styleId!==cell?.styleId||rendered.directStyle!==cell?.style){
          node.style.cssText=visualCss(cell);
          rendered.styleId=cell?.styleId;rendered.directStyle=cell?.style;rendered.styleApplied=true;
        }
        const cellWidth=columns.start(column+(merge?.columnCount??1))-columns.start(column),cellHeight=rows.start(row+(merge?.rowCount??1))-rows.start(row);
        const w=cellWidth+"px",h=cellHeight+"px";
        if(node.style.width!==w)node.style.width=w;if(node.style.height!==h)node.style.height=h;
        node.classList.toggle("frozen",row<=frozenRows||column<=frozenColumns);
        node.classList.toggle("frozen-intersection",row<=frozenRows&&column<=frozenColumns);
        node.classList.toggle("no-grid",sheet.view?.showGridLines===false);
      }
      const current=rendered.node,left=(columnPosition.get(column)??(screenX(column,scrollLeft)-scrollLeft))+"px",top=(rowPosition.get(row)??(screenY(row,scrollTop)-scrollTop))+"px",selected2=String(contains2(range,row,column));
      if(sortedRules.length){
        const value=needsFormula?(currentRevision===snapshot.revision&&result?.state==='valid'?result.value:undefined):cell?.value;
        const rule=conditionalCellStyle(sortedRules,row,column,value);
        const base=visualCss(cell),signature=base+JSON.stringify(rule??null);
        if(rebound||rendered.conditionalSignature!==signature){
          styleProbe.style.cssText=base;
          current.style.backgroundColor=rule?.fill?styleColor(rule.fill):styleProbe.style.backgroundColor;
          current.style.color=rule?.color?styleColor(rule.color):styleProbe.style.color;
          current.style.fontWeight=rule?.bold!==undefined?rule.bold?'700':'400':styleProbe.style.fontWeight;
          current.style.fontStyle=rule?.italic!==undefined?rule.italic?'italic':'normal':styleProbe.style.fontStyle;
          current.style.textDecoration=rule?.underline!==undefined?rule.underline?'underline':'none':styleProbe.style.textDecoration;
          rendered.conditionalSignature=signature;
        }
      }
      if(current.style.left!==left)current.style.left=left;
      if(current.style.top!==top)current.style.top=top;
      if(current.getAttribute("aria-selected")!==selected2)current.setAttribute("aria-selected",selected2);
      current.classList.toggle("active",row===state2.row&&column===state2.column);
      if(rebound||rendered.result!==result||(needsFormula||isOpaque)&&rendered.revision!==currentRevision){
        if(current.title)current.title="";if(current.dataset.calculationState!==undefined)delete current.dataset.calculationState;
        let text:string;
        if(isOpaque){text=result?.value?formatCellValue(result.value,"General",formatOptions):"…";current.title=t("sheetUnsupported");current.dataset.calculationState="unsupported";}
        else if(needsFormula&&!result){text="…";current.title=t("sheetFormulaPending");current.dataset.calculationState="stale";}
        else{
          text=formatCell(cell,result,formatOptions);
          if(result){current.dataset.calculationState=currentRevision===snapshot.revision?result.state:"stale";
            if(currentRevision!==snapshot.revision)current.title=t("sheetStale");
            else if(result.state!=="valid")current.title=t(result.state==="unsupported"?"sheetUnsupported":result.state==="stale"?"sheetStale":"sheetCalcError");
          }
        }
        if(cell?.richText?.length&&cell.value.kind==='string'){
          // The grid keeps scalar text for calculation/clipboard, and
          // projects only trusted model runs, not source XML or HTML.
          const fragment=document.createDocumentFragment();
          for(const run of cell.richText){
            const span=make('span',run.text,'sheets-rich-run'),s=run.style;
            if(s?.fontFamily)span.style.fontFamily=s.fontFamily;
            if(s?.fontSize!==undefined)span.style.fontSize=Math.min(200,s.fontSize/9525)+'px';
            if(s?.bold!==undefined)span.style.fontWeight=s.bold?'700':'400';
            if(s?.italic!==undefined)span.style.fontStyle=s.italic?'italic':'normal';
            if(s?.underline!==undefined)span.style.textDecoration=s.underline?'underline':'none';
            if(s?.color)span.style.color=styleColor(s.color);
            fragment.append(span);
          }
          current.replaceChildren(fragment);
        }else{
          if(rendered.text.parentNode!==current)current.replaceChildren(rendered.text);
          if(rendered.text.data!==text)rendered.text.data=text;
        }
        current.classList.toggle("number",cell?.value.kind==="number"||result?.value?.kind==="number");
        current.classList.toggle("cell-error",cell?.value.kind==="error"||result?.state==="error");
        rendered.result=result;rendered.revision=currentRevision;
      }
    }
    for(const rendered of recycled)rendered.node.remove();
    // Row/column headers are recycled independently of cell contents.
    const rowSet=new Set(visibleRows),columnSet=new Set(visibleColumns),freeRows:HTMLDivElement[]=[],freeColumns:HTMLDivElement[]=[];
    for(const [row,header]of rowHeaders)if(!rowSet.has(row)){rowHeaders.delete(row);freeRows.push(header);}
    for(const [column,header]of columnHeaders)if(!columnSet.has(column)){columnHeaders.delete(column);freeColumns.push(header);}
    for(const row of visibleRows){
      let header=rowHeaders.get(row);
      if(!header){header=freeRows.pop();if(!header){header=make("div","","sheets-row-header");header.style.width=headerWidth+"px";header.setAttribute("role","rowheader");fragment.append(header);}
        header.textContent=String(row);header.style.height=rows.size(row)+"px";header.dataset.headerRow=String(row);rowHeaders.set(row,header);}
      const left="0px",top=rowPosition.get(row)!+"px";
      if(header.style.left!==left)header.style.left=left;if(header.style.top!==top)header.style.top=top;
    }
    for(const column of visibleColumns){
      let header=columnHeaders.get(column);
      if(!header){header=freeColumns.pop();if(!header){header=make("div","","sheets-column-header");header.style.height=headerHeight+"px";header.setAttribute("role","columnheader");fragment.append(header);}
        header.textContent=columnLabel(column);header.style.width=columns.size(column)+"px";header.dataset.headerColumn=String(column);columnHeaders.set(column,header);}
      const top="0px",left=columnPosition.get(column)!+"px";
      if(header.style.top!==top)header.style.top=top;if(header.style.left!==left)header.style.left=left;
    }
    for(const node of [...freeRows,...freeColumns])node.remove();
    corner.style.left="0px";corner.style.top="0px";
    layer.append(fragment);
    if(planned.has(state2.row+":"+state2.column)){if(scroll.getAttribute("aria-activedescendant")!==activeId)scroll.setAttribute("aria-activedescendant",activeId);}
    else if(scroll.hasAttribute("aria-activedescendant"))scroll.removeAttribute("aria-activedescendant");
    const displayedRange=rangeAddress();
    if(!addressDirty&&address.value!==displayedRange)address.value=displayedRange;
    if(!formulaDirty&&!editing){const source=cellSource();if(formula.value!==source)formula.value=source;}
    const selectionKey = state2.row + ":" + state2.column;
    if (controlsKey !== selectionKey) {
      controlsKey = selectionKey;
      const style = currentStyle(), cell = sheet.cells[selectionKey];
      for (const [id, key] of [["sheet-bold", "bold"], ["sheet-italic", "italic"], ["sheet-underline", "underline"]] as const) root2.querySelector("#" + id)?.setAttribute("aria-pressed", String(Boolean(style[key])));
      family.value = style.fontFamily ?? model.defaultTextStyle.fontFamily ?? "Arial";
      fontSize.value = String((style.fontSize ?? model.defaultTextStyle.fontSize ?? 139700) / 12700);
      const fg = root2.querySelector<HTMLInputElement>("#sheet-text-color"), bg = root2.querySelector<HTMLInputElement>("#sheet-fill-color");
      if (fg) fg.value = colorHex(style.color, "#222735");
      if (bg) bg.value = colorHex(style.fill, "#ffffff");
      namedStyle.value=cell?.namedStyle??'Normal';
      custom.value = cell?.numberFormat ?? "General";
      if (!Array.from(formats.options).some((o) => o.value === custom.value)) {
        const option = make("option", custom.value);
        option.value = custom.value;
        formats.append(option);
      }
      formats.value = custom.value;
      const rh = root2.querySelector<HTMLInputElement>("#sheet-row-height"), cw = root2.querySelector<HTMLInputElement>("#sheet-column-width");
      if (rh) rh.value = String(Math.round((sheet.rows[state2.row]?.height ?? sheet.defaultGeometry?.rowHeight ?? 24 * 9525) / 9525));
      if (cw) cw.value = String(Math.round((sheet.columns[state2.column]?.width ?? sheet.defaultGeometry?.columnWidth ?? 104 * 9525) / 9525));
    }
    const summarySignature = displayedRange + ":" + state2.calculatedRevision;
    if (summarySignature !== summaryKey) {
      const aggregate = (row:number, column:number, cell:CellSnapshot|undefined) => {
        if (!cell) return;
        const result = state2.calculation?.get(calculationKey(sheet.id, row, column)), value = cell.value.kind === "formula" ? state2.calculatedRevision === snapshot.revision && result?.state === "valid" ? result.value : void 0 : cell.value;
        if (value?.kind === "number") {
          count++;
          sum += value.value;
        }
      };
      if (range.rowCount * range.columnCount <= 1e4) {
        for (let row = range.row; row < range.row + range.rowCount; row++) for (let column = range.column; column < range.column + range.columnCount; column++) aggregate(row, column, sheet.cells[row + ":" + column]);
      } else for (const [key, cell] of Object.entries(sheet.cells)) {
        const [row, column] = key.split(":").map(Number);
        if (contains2(range, row, column)) aggregate(row, column, cell);
      }
      selectedSummary.textContent = t("sheetPosition", { address: displayedRange }) + " \xB7 " + t("sheetSelectionSummary", { count: new Intl.NumberFormat(host.locale()).format(count), sum: new Intl.NumberFormat(host.locale(), { maximumFractionDigits: 10 }).format(sum) });
      summaryKey = summarySignature;
    }
    const statusText=t(host.readOnly&&opaque.size?"sheetUnsupported":state2.calculating?"sheetCalculating":state2.calculationError?"sheetCalcError":"sheetCalculated");
    if(calcStatus.textContent!==statusText)calcStatus.textContent=statusText;
    const revision=String(state2.calculatedRevision??-1),calculating=String(Boolean(state2.calculating));
    if(root2.dataset.calculatedRevision!==revision)root2.dataset.calculatedRevision=revision;
    if(root2.dataset.calculating!==calculating)root2.dataset.calculating=calculating;
    if (editing) positionEditor();
  }
  function schedule() {
    if (frame) return;
    frame = requestAnimationFrame(() => {
      frame = 0;
      paint();
    });
  }
  scroll.onscroll = () => {
    state2.scrollTop = scroll.scrollTop;
    state2.scrollLeft = scroll.scrollLeft;
    schedule();
  };
  scroll.onpointerdown = (event) => {
    if (event.target === editor || host.composing()) return;
    if (!flush()) {
      event.preventDefault();
      return;
    }
    const node = (event.target instanceof Element ? event.target.closest<HTMLElement>("[data-row],[data-header-row],[data-header-column]") : null);
    if (!node) return;
    event.preventDefault();
    scroll.focus();
    dragging = true;
    if (node.dataset.headerRow) {
      select(Number(node.dataset.headerRow), 1, false, false);
      state2.column = limits.columns;
      paint();
      dragging = false;
    } else if (node.dataset.headerColumn) {
      select(1, Number(node.dataset.headerColumn), false, false);
      state2.row = limits.rows;
      paint();
      dragging = false;
    } else select(Number(node.dataset.row), Number(node.dataset.column), event.shiftKey, false);
  };
  scroll.onpointermove = (event) => {
    if (!dragging || !event.buttons) return;
    const node = (event.target instanceof Element ? event.target.closest<HTMLElement>("[data-row]") : null);
    if (node) select(Number(node.dataset.row), Number(node.dataset.column), true, false);
  };
  scroll.onpointerup = () => {
    dragging = false;
  };
  scroll.onpointerleave = () => {
    dragging = false;
  };
  scroll.ondblclick = (event) => {
    const node = (event.target instanceof Element ? event.target.closest<HTMLElement>("[data-row]") : null);
    if (node) beginEdit();
  };
  const move = (key:string, extend = false) => {
    const merge = activeMerge();
    let row = state2.row, column = state2.column;
    if (sheet.view?.rightToLeft && key === "ArrowLeft") key = "ArrowRight";
    else if (sheet.view?.rightToLeft && key === "ArrowRight") key = "ArrowLeft";
    if (key === "ArrowDown" || key === "Enter") row += merge?.rowCount ?? 1;
    else if (key === "ArrowUp") row--;
    else if (key === "ArrowRight" || key === "Tab") column += merge?.columnCount ?? 1;
    else if (key === "ArrowLeft") column--;
    else if (key === "Home") column = 1;
    else if (key === "PageDown") row += Math.max(1, Math.floor(scroll.clientHeight / 24) - 2);
    else if (key === "PageUp") row -= Math.max(1, Math.floor(scroll.clientHeight / 24) - 2);
    select(row, column, extend);
  };
  const clear = () => {
    changed(clearCells(host.snapshot(), sheet.id, selectionRange()));
  };
  scroll.onkeydown = (event) => {
    if (event.target === editor || host.composing()) return;
    if (event.ctrlKey || event.metaKey) {
      if (event.key.toLowerCase() === "a") {
        event.preventDefault();
        state2.anchorRow = state2.anchorColumn = 1;
        state2.row = limits.rows;
        state2.column = limits.columns;
        paint();
      }
      return;
    }
    if (["ArrowDown", "ArrowUp", "ArrowLeft", "ArrowRight", "Home", "PageDown", "PageUp", "Tab", "Enter"].includes(event.key)) {
      event.preventDefault();
      move(event.shiftKey && event.key === "Tab" ? "ArrowLeft" : event.key, event.shiftKey && event.key !== "Tab");
    } else if (event.key === "F2") {
      event.preventDefault();
      beginEdit();
    } else if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      guard(clear);
    } else if (event.key.length === 1 && !event.altKey) {
      event.preventDefault();
      beginEdit(event.key);
    }
  };
  for (const input of [editor, formula]) {
    input.addEventListener("compositionstart", () => host.setComposing(true));
    input.addEventListener("compositionend", () => {
      host.setComposing(false);
      if (input === formula) formulaDirty = true;
    });
    input.onkeydown = (event) => {
      if (event.isComposing || host.composing()) return;
      if (event.key === "Escape") {
        event.preventDefault();
        editing = false;
        formulaDirty = false;
        editor.hidden = true;
        formula.value = cellSource();
        scroll.focus();
      } else if (event.key === "Enter" || event.key === "Tab") {
        event.preventDefault();
        applyInput(input.value, input === editor ? event.key === "Tab" && event.shiftKey ? "BackTab" : event.key : void 0);
      }
    };
    input.onblur = (event) => {
      if (host.composing()) return;
      if (event.relatedTarget instanceof Node && root2.contains(event.relatedTarget)) return;
      if (input === editor && editing) flush();
      else if (input === formula && formulaDirty) flush();
    };
  }
  formula.oninput = () => {
    formulaDirty = true;
  };
  editor.oninput = () => {
    formula.value = editor.value;
  };
  // The address bar has independent draft input. Worker replies and resize/scroll
  // repaints must not replace it before the user confirms navigation.
  state2.navigate = (row, column, endRow, endColumn, focusGrid) => {
    if (!root2.isConnected || !host.isCurrent()) return;
    addressDirty = false;
    select(row, column, false);
    select(endRow, endColumn, true);
    if (focusGrid) scroll.focus();
  };
  function commitAddress(focusGrid:boolean):void {
    if (!addressDirty && address.value === rangeAddress()) { if(focusGrid) scroll.focus(); return; }
    // Capture the target before flush(), which can synchronously remount the grid.
    const target=address.value.trim();
    try {
      const match = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?$/.exec(target);
      if (!match) throw new OfficeError("invalid", t("sheetInputInvalid"));
      const col = (value:string) => [...value.toUpperCase()].reduce((n, character) => n * 26 + character.charCodeAt(0) - 64, 0);
      const row=Number(match[2]), column=col(match[1]), endRow=Number(match[4]??match[2]), endColumn=col(match[3]??match[1]);
      if (Math.max(row,endRow)>limits.rows || Math.max(column,endColumn)>limits.columns) throw new OfficeError("limit", "Selection exceeds workbook format limits");
      if (!flush()) return;
      addressDirty=false;
      // Dispatch to the live workspace, not a detached DOM captured before flush.
      state2.navigate?.(row,column,endRow,endColumn,focusGrid);
    } catch(error) { host.error(error); }
  }
  address.oninput = () => { addressDirty=true; };
  address.onchange = () => commitAddress(false);
  address.onkeydown = (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      commitAddress(true);
    } else if (event.key === "Escape") {
      event.preventDefault();
      addressDirty=false;
      address.value = rangeAddress();
      scroll.focus();
    }
  };
  scroll.oncopy = (event) => {
    if (editing || !event.clipboardData) return;
    guard(() => {
      const range = selectionRange();
      let clipboardModel:WorkbookSnapshot = model;
      if (host.readOnly && opaque.size) {
        const cached2:Record<string,CellSnapshot> = {};
        for (const key of opaque) {
          if (!key.startsWith(sheet.id + "!")) continue;
          const value = state2.calculation?.get(key)?.value;
          if (value) cached2[key.slice(sheet.id.length + 1)] = { value };
        }
        clipboardModel = { ...clipboardModel, sheets: clipboardModel.sheets.map((s) => s.id === sheet.id ? { ...s, cells: { ...s.cells, ...cached2 } } : s) };
      }
      const values = exportTsv(clipboardModel, sheet.id, range, { calculation: { cells: state2.calculation ?? /* @__PURE__ */ new Map(), computedCells: 0, steps: 0, get: (sid, r, c) => {
        const value = clipboardModel.sheets.find((x) => x.id === sid)?.cells[`${r}:${c}`]?.value;
        return value?.kind === "formula" ? state2.calculatedRevision === snapshot.revision ? state2.calculation?.get(calculationKey(sid, r, c)) ?? { state: "stale" } : { state: "stale" } : { state: "valid", value: value ?? { kind: "blank" } };
      } } });
      event.clipboardData!.setData("text/plain", values);
      const nonce = crypto.randomUUID();
      copied = { nonce, documentId: model.documentId, contentId: snapshot.contentId, sheetId: sheet.id, range };
      event.clipboardData!.setData("application/x-office-ex-sheets-range", nonce);
      event.preventDefault();
    });
  };
  scroll.onpaste = (event) => {
    if (editing || host.readOnly || !event.clipboardData) return;
    event.preventDefault();
    guard(() => {
      const marker = event.clipboardData!.getData("application/x-office-ex-sheets-range");
      if (copied && marker === copied.nonce && copied.documentId === model.documentId && copied.contentId === host.snapshot().contentId) changed(copyCells(host.snapshot(), copied.sheetId, copied.range, sheet.id, state2.row, state2.column));
      else changed(pasteTsv(host.snapshot(), sheet.id, state2.row, state2.column, event.clipboardData!.getData("text/plain"), options()));
    });
  };
  state2.onCalculation = () => {
    if (root2.isConnected && host.isCurrent()) paint();
  };
  const observer = new ResizeObserver(() => {
    if (!root2.isConnected) {
      observer.disconnect();
      return;
    }
    schedule();
  });
  observer.observe(scroll);
  queueMicrotask(() => {
    if (!root2.isConnected) return;
    scroll.scrollTop = state2.scrollTop;
    scroll.scrollLeft = state2.scrollLeft;
    recalculate(host, state2);
    paint();
  });
  const pick=(...ids:string[])=>ribbonNodes(root2,...ids);
  installRibbon(root2,state2,t,[
    {id:'home',label:'ribbonHome',groups:[
      {label:'ribbonFont',rows:[pick('sheet-font','sheet-font-size'),pick('sheet-bold','sheet-italic','sheet-underline','sheet-text-color','sheet-fill-color','sheet-border')]},
      {label:'ribbonAlignment',rows:[pick('sheet-align-start','sheet-align-center','sheet-align-end','sheet-wrap'),pick('sheet-merge','sheet-unmerge')]},
      {label:'ribbonNumbers',rows:[pick('sheet-number-format'),pick('sheet-custom-format')]},
      {label:'ribbonStyles',rows:[pick('sheet-named-style','sheet-conditional-open')]},
      {label:'sheetCharts',rows:[pick('sheet-chart-open')]},
      {label:'ribbonEditing',rows:[pick('sheet-clear'),pick('sheet-fill-down','sheet-fill-right')]}
    ]},
    {id:'cells',label:'ribbonCells',groups:[
      {label:'ribbonRows',rows:[pick('sheet-insert-rows','sheet-delete-rows'),pick('sheet-row-height')]},
      {label:'ribbonColumns',rows:[pick('sheet-insert-columns','sheet-delete-columns'),pick('sheet-column-width')]},
      {label:'ribbonWorksheets',rows:[pick('sheet-name','sheet-rename'),pick('sheet-delete','sheet-move-left','sheet-move-right')]}
    ]},
    {id:'formulas',label:'ribbonFormulas',groups:[
      {label:'ribbonCalculation',rows:[pick('sheet-formula-mode'),pick('sheet-recalculate','sheet-cancel-calculation')]},
      {label:'ribbonNumbers',rows:[pick('sheet-date-system')]}
    ]},
    {id:'view',label:'ribbonView',groups:[
      {label:'ribbonWindow',rows:[pick('sheet-freeze','sheet-unfreeze'),pick('sheet-grid-lines','sheet-rtl')]}
    ]}
  ],()=>!host.composing()&&flush());
  // Sheet creation remains beside the real worksheet tabs.
  decorateOfficeButton(add,'add',true);
  for(const container of [toolbar,second,management])if(!container.childElementCount)container.remove();
  if(!management.isConnected&&controls2.childElementCount===1)controls2.remove();
  else if(!management.childElementCount)controls2.remove();
  return root2;
}
