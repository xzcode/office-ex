/** Recovered M3 virtualized worksheet workspace. */
import { OfficeError, EMU_PER_CSS_PIXEL } from '@office-ex/contracts';
import { patchCommand, type DocumentSnapshot, type DocumentCommand, type Color, type CellStyle, type SheetView } from '@office-ex/core';
import { addSheet, renameSheet, deleteSheet, moveSheet, insertRows, deleteRows, insertColumns, deleteColumns, mergeCells, unmergeCells, clearCells, copyCells, formatCells, setDimensions, fillCells, pasteTsv, exportTsv, setCells, cellFromInput, formatCell, formatCellValue, calculationKey, localizeFormula, formulaSeparators, type CellRange, type WorkbookSnapshot, type CalculationCell, type CalculatedValue, type CellInputOptions } from '@office-ex/sheets';
import type { MessageKey } from '@office-ex/i18n';
import { GridAxis, columnLabel, cellLabel } from './sheets-geometry.js';

type CellSnapshot = WorkbookSnapshot['sheets'][number]['cells'][string];
export interface SheetsViewState {
  sheetId?:string; row:number; column:number; anchorRow:number; anchorColumn:number;
  scrollLeft:number; scrollTop:number; inputMode:'canonical'|'regional';
  calculation?:Map<string,CalculationCell>; calculatedRevision?:number;
  calculating?:boolean; calculationError?:boolean; onCalculation?:()=>void;
  cachedValues?:[string,CalculatedValue][]; cachedSources?:[string,string][];
  unsupportedFormulas?:string[]; flush?:()=>boolean;
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
interface RenderedCell {node:HTMLDivElement;result?:CalculationCell;revision?:number}
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
    input.value = String(Math.round((current?.[size] ?? defaultPx * 9525) / 9525));
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
  const rows = new GridAxis(limits.rows, 24, sheet.rows, "height"), columns = new GridAxis(limits.columns, 104, sheet.columns, "width");
  const headerWidth = 46, headerHeight = 26;
  const totalHeight = Math.min(3e7, rows.total + headerHeight), totalWidth = Math.min(3e7, columns.total + headerWidth);
  plane.style.height = totalHeight + "px";
  plane.style.width = totalWidth + "px";
  const displayRowScale = (rows.total + headerHeight) / totalHeight, displayColScale = (columns.total + headerWidth) / totalWidth;
  const layer = make("div", "", "sheets-cell-layer");
  plane.append(layer);
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
  let editing = false, formulaDirty = false, frame = 0, dragging = false, summaryKey = "", controlsKey = "";
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
  const visualStyle = (node:HTMLDivElement, cell:CellSnapshot|undefined) => {
    const style = { ...model.defaultTextStyle, ...cell?.styleId ? model.styles[cell.styleId]?.properties : {}, ...cell?.style };
    node.style.fontFamily = style.fontFamily ?? "sans-serif";
    node.style.fontSize = Math.min(200, (style.fontSize ?? 139700) / 9525) + "px";
    node.style.fontWeight = style.bold ? "700" : "400";
    node.style.fontStyle = style.italic ? "italic" : "normal";
    node.style.textDecoration = style.underline ? "underline" : "none";
    node.style.color = style.color ? styleColor(style.color) : "var(--text)";
    if (style.fill) node.style.backgroundColor = styleColor(style.fill);
    if (style.alignment) node.style.textAlign = style.alignment;
    if (style.wrapText) {
      node.style.whiteSpace = "pre-wrap";
      node.style.overflowWrap = "anywhere";
    }
    if (style.border) node.style.border = Math.max(0.5, style.border.width / 9525) + "px solid " + styleColor(style.border.color);
    node.style.alignContent = style.verticalAlignment === "top" ? "start" : style.verticalAlignment === "center" ? "center" : "end";
  };
  function paint() {
    if (!root2.isConnected) return;
    const scrollTop = scroll.scrollTop, scrollLeft = scroll.scrollLeft, height = scroll.clientHeight, width = scroll.clientWidth;
    const y = scrollTop * displayRowScale, x = scrollLeft * displayColScale;
    const frozenRows = sheet.view?.frozenRows ?? 0, frozenColumns = sheet.view?.frozenColumns ?? 0;
    const fixedRows = rows.visible(0, Math.min(height / 2, 200), 40).filter((r) => r <= frozenRows), fixedColumns = columns.visible(0, Math.min(width / 2, 400), 20).filter((c) => c <= frozenColumns);
    const visibleRows = [.../* @__PURE__ */ new Set([...fixedRows, ...rows.visible(y, height, 160)])].filter((r) => rows.size(r) > 0);
    const visibleColumns = [.../* @__PURE__ */ new Set([...fixedColumns, ...columns.visible(x, width, 60)])].filter((c) => columns.size(c) > 0);
    const fragment = document.createDocumentFragment(), range = selectionRange();
    let count = 0, sum = 0;
    const activeId = "sheet-cell-" + state2.row + "-" + state2.column;
    const painted = new Set<string>();
    const drawCell = (row:number, column:number, merge?:CellRange) => {
      const key = `${row}:${column}`;
      if (painted.has(key)) return;
      painted.add(key);
      const cell = sheet.cells[key], formulaKey = calculationKey(sheet.id, row, column), isOpaque = opaque.has(formulaKey), result = cell?.value.kind === "formula" || isOpaque ? state2.calculation?.get(formulaKey) : void 0;
      let rendered = renderedCells.get(key);
      const created = !rendered;
      if (!rendered) {
        const node = make("div", "", "sheets-cell");
        rendered = { node };
        renderedCells.set(key, rendered);
        node.id = "sheet-cell-" + row + "-" + column;
        node.dataset.row = String(row);
        node.dataset.column = String(column);
        node.dataset.address = cellLabel(row, column);
        node.setAttribute("role", "gridcell");
        node.setAttribute("aria-rowindex", String(row));
        node.setAttribute("aria-colindex", String(column));
        node.setAttribute("aria-label", cellLabel(row, column));
        node.dir = "auto";
        if (merge) {
          node.setAttribute("aria-rowspan", String(merge.rowCount));
          node.setAttribute("aria-colspan", String(merge.columnCount));
        }
        node.style.width = columns.start(column + (merge?.columnCount ?? 1)) - columns.start(column) + "px";
        node.style.height = rows.start(row + (merge?.rowCount ?? 1)) - rows.start(row) + "px";
        if (row <= frozenRows || column <= frozenColumns) node.classList.add("frozen");
        if (row <= frozenRows && column <= frozenColumns) node.classList.add("frozen-intersection");
        if (sheet.view?.showGridLines === false) node.classList.add("no-grid");
        visualStyle(node, cell);
        fragment.append(node);
      }
      const current = rendered.node, left = screenX(column, scrollLeft) + "px", top = screenY(row, scrollTop) + "px", selected2 = String(contains2(range, row, column));
      if (current.style.left !== left) current.style.left = left;
      if (current.style.top !== top) current.style.top = top;
      if (current.getAttribute("aria-selected") !== selected2) current.setAttribute("aria-selected", selected2);
      current.classList.toggle("active", row === state2.row && column === state2.column);
      if (created || rendered.result !== result || rendered.revision !== state2.calculatedRevision) {
        current.title = "";
        delete current.dataset.calculationState;
        if (isOpaque) {
          current.textContent = result?.value ? formatCellValue(result.value, "General", { locale: host.locale(), dateSystem: model.dateSystem }) : "\u2026";
          current.title = t("sheetUnsupported");
          current.dataset.calculationState = "unsupported";
        } else if (cell?.value.kind === "formula" && !result) {
          current.textContent = "\u2026";
          current.title = t("sheetFormulaPending");
          current.dataset.calculationState = "stale";
        } else {
          current.textContent = formatCell(cell, result, { locale: host.locale(), dateSystem: model.dateSystem });
          if (result) {
            current.dataset.calculationState = state2.calculatedRevision === snapshot.revision ? result.state : "stale";
            if (state2.calculatedRevision !== snapshot.revision) current.title = t("sheetStale");
            else if (result.state !== "valid") current.title = t(result.state === "unsupported" ? "sheetUnsupported" : result.state === "stale" ? "sheetStale" : "sheetCalcError");
          }
        }
        current.classList.toggle("number", cell?.value.kind === "number" || result?.value?.kind === "number");
        current.classList.toggle("cell-error", cell?.value.kind === "error" || result?.state === "error");
        rendered.result = result;
        rendered.revision = state2.calculatedRevision;
      }
    };
    for (const row of visibleRows) for (const column of visibleColumns) {
      const merge = sheet.merges.find((m) => contains2(m, row, column));
      if (merge) drawCell(merge.row, merge.column, merge);
      else drawCell(row, column);
    }
    for (const [key, rendered] of renderedCells) if (!painted.has(key)) {
      rendered.node.remove();
      renderedCells.delete(key);
    }
    for (const row of visibleRows) {
      let header = rowHeaders.get(row);
      if (!header) {
        header = make("div", String(row), "sheets-row-header");
        header.style.width = headerWidth + "px";
        header.style.height = rows.size(row) + "px";
        header.dataset.headerRow = String(row);
        header.setAttribute("role", "rowheader");
        rowHeaders.set(row, header);
        fragment.append(header);
      }
      header.style.left = scrollLeft + "px";
      header.style.top = screenY(row, scrollTop) + "px";
    }
    for (const column of visibleColumns) {
      let header = columnHeaders.get(column);
      if (!header) {
        header = make("div", columnLabel(column), "sheets-column-header");
        header.style.width = columns.size(column) + "px";
        header.style.height = headerHeight + "px";
        header.dataset.headerColumn = String(column);
        header.setAttribute("role", "columnheader");
        columnHeaders.set(column, header);
        fragment.append(header);
      }
      header.style.top = scrollTop + "px";
      header.style.left = screenX(column, scrollLeft) + "px";
    }
    for (const [row, header] of rowHeaders) if (!visibleRows.includes(row)) {
      header.remove();
      rowHeaders.delete(row);
    }
    for (const [column, header] of columnHeaders) if (!visibleColumns.includes(column)) {
      header.remove();
      columnHeaders.delete(column);
    }
    corner.style.left = scrollLeft + "px";
    corner.style.top = scrollTop + "px";
    layer.append(fragment);
    if (painted.has(state2.row + ":" + state2.column)) scroll.setAttribute("aria-activedescendant", activeId);
    else scroll.removeAttribute("aria-activedescendant");
    address.value = rangeAddress();
    if (!formulaDirty && !editing) formula.value = cellSource();
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
      custom.value = cell?.numberFormat ?? "General";
      if (!Array.from(formats.options).some((o) => o.value === custom.value)) {
        const option = make("option", custom.value);
        option.value = custom.value;
        formats.append(option);
      }
      formats.value = custom.value;
      const rh = root2.querySelector<HTMLInputElement>("#sheet-row-height"), cw = root2.querySelector<HTMLInputElement>("#sheet-column-width");
      if (rh) rh.value = String(Math.round((sheet.rows[state2.row]?.height ?? 24 * 9525) / 9525));
      if (cw) cw.value = String(Math.round((sheet.columns[state2.column]?.width ?? 104 * 9525) / 9525));
    }
    const summarySignature = rangeAddress() + ":" + state2.calculatedRevision;
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
      selectedSummary.textContent = t("sheetPosition", { address: rangeAddress() }) + " \xB7 " + t("sheetSelectionSummary", { count: new Intl.NumberFormat(host.locale()).format(count), sum: new Intl.NumberFormat(host.locale(), { maximumFractionDigits: 10 }).format(sum) });
      summaryKey = summarySignature;
    }
    calcStatus.textContent = t(host.readOnly && opaque.size ? "sheetUnsupported" : state2.calculating ? "sheetCalculating" : state2.calculationError ? "sheetCalcError" : "sheetCalculated");
    root2.dataset.calculatedRevision = String(state2.calculatedRevision ?? -1);
    root2.dataset.calculating = String(Boolean(state2.calculating));
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
  address.onchange = () => guard(() => {
    const match = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?$/.exec(address.value.trim());
    if (!match) throw new OfficeError("invalid", t("sheetInputInvalid"));
    const col = (s:string) => [...s.toUpperCase()].reduce((n, c3) => n * 26 + c3.charCodeAt(0) - 64, 0);
    const r = Number(match[2]), c = col(match[1]), r2 = Number(match[4] ?? match[2]), c2 = col(match[3] ?? match[1]);
    if (Math.max(r, r2) > limits.rows || Math.max(c, c2) > limits.columns) throw new OfficeError("limit", "Selection exceeds workbook format limits");
    select(r, c, false);
    select(r2, c2, true);
    scroll.focus();
  });
  address.onkeydown = (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      address.dispatchEvent(new Event("change"));
    } else if (event.key === "Escape") {
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
  return root2;
}
