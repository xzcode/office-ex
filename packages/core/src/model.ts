import { OfficeError, assert, assertCellAddress, assertId, assertUnicode, isRecord, type DocumentKind } from '@office-ex/contracts';

export type Color = { kind: 'rgba'; red: number; green: number; blue: number; alpha: number } | { kind: 'theme'; name: string; tint?: number };
export interface TextStyle {
  fontFamily?: string; fontSize?: number; bold?: boolean; italic?: boolean; underline?: boolean;
  color?: Color; alignment?: 'start' | 'center' | 'end' | 'justify';
  indent?: number; spaceBefore?: number; spaceAfter?: number; lineHeight?: number;
}
export interface NamedStyle { id: string; name: string; basedOn?: string; kind?: 'paragraph' | 'character'; toggles?: readonly ('bold' | 'italic')[]; properties: TextStyle }
export interface ResourceRef { id: string; sha256: string; mimeType: string; byteLength: number }
export interface DocumentBase {
  schemaVersion: 1; documentId: string; title: string;
  resources: Record<string, ResourceRef>; styles: Record<string, NamedStyle>; defaultTextStyle: TextStyle;
}
export interface TextRun { id: string; type: 'text'; text: string; styleId?: string; style?: TextStyle }
export type Inline = TextRun | { id: string; type: 'tab' | 'line-break' } | { id: string; type: 'image'; resourceId: string; width: number; height: number };
export interface Paragraph { id: string; type: 'paragraph'; styleId?: string; style?: TextStyle; runs: Inline[] }
export interface TableCell { id: string; blocks: WriterBlock[] }
export interface Table { id: string; type: 'table'; columnWidths: number[]; rows: { id: string; cells: TableCell[] }[]; cellPadding?: {top:number;right:number;bottom:number;left:number}; border?: {width:number;color:Color} }
export type WriterBlock = Paragraph | Table | { id: string; type: 'page-break' };
export interface WriterSection { id: string; page: { width: number; height: number; margins: { top: number; right: number; bottom: number; left: number } }; blocks: WriterBlock[] }
export interface WriterDocument extends DocumentBase { kind: 'writer'; sections: WriterSection[]; defaultParagraphStyleId?:string }

export type CellValue = { kind: 'blank' } | { kind: 'string'; value: string } | { kind: 'number'; value: number } | { kind: 'boolean'; value: boolean } | { kind: 'error'; value: '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' } | { kind: 'formula'; source: string };
export interface Cell { value: CellValue; styleId?: string; numberFormat?: string }
export interface Sheet {
  id: string; name: string; cells: Record<string, Cell>;
  rows: Record<string, { height?: number; hidden?: boolean }>;
  columns: Record<string, { width?: number; hidden?: boolean }>;
  merges: { row: number; column: number; rowCount: number; columnCount: number }[];
}
export interface SheetsDocument extends DocumentBase { kind: 'sheets'; dateSystem: '1900' | '1904'; sheets: Sheet[] }

export interface Geometry { x: number; y: number; width: number; height: number; rotation: number }
interface ShapeBase { id: string; geometry: Geometry; fill?: Color; stroke?: { color: Color; width: number } }
export type Shape = (ShapeBase & { type: 'text'; runs: TextRun[] }) | (ShapeBase & { type: 'rect' | 'ellipse' | 'line' }) | (ShapeBase & { type: 'image'; resourceId: string });
export interface Slide { id: string; background: Color; shapes: Shape[] }
export interface SlidesDocument extends DocumentBase { kind: 'slides'; size: { width: number; height: number }; slides: Slide[] }
export type OfficeDocument = WriterDocument | SheetsDocument | SlidesDocument;
export type Document = OfficeDocument;

/** Reject executable/prototyped data, cycles, excessive depth, and dangerous object keys. */
export function assertPlainData(value: unknown, maxNodes = 1_000_000, maxDepth = 128): void {
  const ancestors = new Set<object>();
  let nodes = 0;
  let textUnits = 0;
  const visit = (item: unknown, depth: number): void => {
    assert(++nodes <= maxNodes && depth <= maxDepth, 'limit', 'Model resource budget exceeded');
    if (item === null || typeof item === 'boolean') return;
    if (typeof item === 'number') { assert(Number.isFinite(item), 'invalid', 'Non-finite model number'); return; }
    if (typeof item === 'string') {
      textUnits += item.length;
      assert(textUnits <= 64 * 1024 * 1024, 'limit', 'Model text budget exceeded');
      assertUnicode(item); return;
    }
    assert(Array.isArray(item) || isRecord(item), 'invalid', 'Model must contain only plain serializable data');
    assert(!ancestors.has(item), 'invalid', 'Cyclic model');
    ancestors.add(item);
    if (Array.isArray(item)) {
      assert(item.length <= maxNodes, 'limit', 'Model array budget exceeded');
      assert(Object.keys(item).length === item.length && Object.getOwnPropertyNames(item).length === item.length + 1 && Object.getOwnPropertySymbols(item).length === 0, 'invalid', 'Sparse or decorated model array');
      for (let i = 0; i < item.length; i++) {
        const descriptor = Object.getOwnPropertyDescriptor(item, String(i));
        assert(descriptor && 'value' in descriptor, 'invalid', 'Model accessors are forbidden');
        visit(descriptor.value, depth + 1);
      }
    } else {
      assert(Object.getOwnPropertySymbols(item).length === 0, 'invalid', 'Model symbols are forbidden');
      for (const name of Object.getOwnPropertyNames(item)) {
        assert(!['__proto__', 'constructor', 'prototype'].includes(name), 'unsafe', 'Unsafe model key');
        const descriptor = Object.getOwnPropertyDescriptor(item, name);
        assert(descriptor && 'value' in descriptor && descriptor.enumerable, 'invalid', 'Model accessors and hidden properties are forbidden');
        visit(descriptor.value, depth + 1);
      }
    }
    ancestors.delete(item);
  };
  visit(value, 0);
}

function object(value: unknown, fields: readonly string[]): Record<string, unknown> {
  assert(isRecord(value), 'invalid', 'Expected model object');
  assert(Object.keys(value).every((key) => fields.includes(key)), 'invalid', 'Unknown model field');
  return value;
}
function list(value: unknown, min = 0): unknown[] { assert(Array.isArray(value) && value.length >= min, 'invalid', 'Invalid model list'); return value; }
function text(value: unknown, max = 1_000_000, empty = true): asserts value is string {
  assert(typeof value === 'string' && value.length <= max && (empty || value.length > 0), 'invalid', 'Invalid model text');
}
function integer(value: unknown, min = -Number.MAX_SAFE_INTEGER, max = Number.MAX_SAFE_INTEGER): asserts value is number {
  assert(typeof value === 'number' && Number.isSafeInteger(value) && value >= min && value <= max, 'invalid', 'Invalid integer model value');
}
function bool(value: unknown): void { assert(typeof value === 'boolean', 'invalid', 'Invalid boolean model value'); }
function optionalId(value: unknown, known: Record<string, unknown>): void { if (value !== undefined) { assertId(value); assert(Object.hasOwn(known, value), 'invalid', 'Missing model reference'); } }

export function validateColor(value: unknown): asserts value is Color {
  const item = object(value, ['kind', 'red', 'green', 'blue', 'alpha', 'name', 'tint']);
  if (item.kind === 'rgba') {
    assert(!('name' in item) && !('tint' in item), 'invalid', 'Unexpected RGBA fields');
    for (const key of ['red', 'green', 'blue', 'alpha']) integer(item[key], 0, 255);
  } else {
    assert(item.kind === 'theme' && !['red', 'green', 'blue', 'alpha'].some((key) => key in item), 'invalid', 'Invalid theme color');
    text(item.name, 128, false);
    if (item.tint !== undefined) assert(typeof item.tint === 'number' && Number.isFinite(item.tint) && item.tint >= -1 && item.tint <= 1, 'invalid', 'Invalid color tint');
  }
}

export function validateTextStyle(value: unknown): asserts value is TextStyle {
  const style = object(value, ['fontFamily', 'fontSize', 'bold', 'italic', 'underline', 'color', 'alignment', 'indent', 'spaceBefore', 'spaceAfter', 'lineHeight']);
  if (style.fontFamily !== undefined) { text(style.fontFamily, 512, false); assert(!/[\x00-\x1f\x7f]/.test(style.fontFamily), 'invalid', 'Invalid font family'); }
  if (style.fontSize !== undefined) integer(style.fontSize, 1, 12_700 * 1638);
  for (const key of ['bold', 'italic', 'underline']) if (style[key] !== undefined) bool(style[key]);
  if (style.color !== undefined) validateColor(style.color);
  if (style.alignment !== undefined) assert(['start', 'center', 'end', 'justify'].includes(String(style.alignment)), 'invalid', 'Invalid text alignment');
  if (style.indent !== undefined) integer(style.indent);
  for (const key of ['spaceBefore', 'spaceAfter', 'lineHeight']) if (style[key] !== undefined) integer(style[key], 0);
}

export function validateGeometry(value: unknown): asserts value is Geometry {
  const geometry = object(value, ['x', 'y', 'width', 'height', 'rotation']);
  integer(geometry.x); integer(geometry.y); integer(geometry.width, 0); integer(geometry.height, 0); integer(geometry.rotation);
  assert(Number.isSafeInteger(geometry.x + geometry.width) && Number.isSafeInteger(geometry.y + geometry.height), 'invalid', 'Geometry sum exceeds safe precision');
}

/** All model boundaries use this validator, including recovery and import adapters. */
export function validateDocument(value: unknown): OfficeDocument {
  assertPlainData(value);
  assert(isRecord(value), 'invalid', 'Expected document');
  if (value.schemaVersion !== 1) throw new OfficeError('unsupported', 'Unsupported document schema version');
  const fields = ['schemaVersion', 'documentId', 'title', 'resources', 'styles', 'defaultTextStyle', 'kind'];
  const model = object(value, [...fields, ...(value.kind === 'writer' ? ['sections','defaultParagraphStyleId'] : value.kind === 'sheets' ? ['dateSystem', 'sheets'] : value.kind === 'slides' ? ['size', 'slides'] : [])]);
  assertId(model.documentId); text(model.title, 1024, false);
  const ids = new Set<string>();
  const id = (candidate: unknown): void => { assertId(candidate); assert(!ids.has(candidate), 'invalid', 'Duplicate stable node identifier'); ids.add(candidate); };
  assert(isRecord(model.resources), 'invalid', 'Invalid resources');
  let resourceBytes = 0;
  for (const [key, item] of Object.entries(model.resources)) {
    assertId(key);
    const resource = object(item, ['id', 'sha256', 'mimeType', 'byteLength']);
    assert(resource.id === key, 'invalid', 'Resource identifier mismatch');
    assert(typeof resource.sha256 === 'string' && /^[a-f0-9]{64}$/.test(resource.sha256), 'invalid', 'Invalid resource digest');
    text(resource.mimeType, 255, false);
    assert(/^[a-z0-9!#$&^_.+-]+\/[a-z0-9!#$&^_.+-]+$/i.test(resource.mimeType), 'invalid', 'Invalid resource media type');
    integer(resource.byteLength, 0, 128 * 1024 * 1024);
    resourceBytes += resource.byteLength;
    assert(resourceBytes <= 128 * 1024 * 1024, 'limit', 'Document resource budget exceeded');
  }
  assert(isRecord(model.styles), 'invalid', 'Invalid styles');
  const styles = model.styles;
  for (const [key, item] of Object.entries(styles)) {
    assertId(key);
    const style = object(item, ['id', 'name', 'basedOn', 'kind', 'toggles', 'properties']);
    assert(style.id === key, 'invalid', 'Style identifier mismatch'); text(style.name, 512, false);
    optionalId(style.basedOn, styles); validateTextStyle(style.properties);
    if(style.kind!==undefined)assert(style.kind==='paragraph'||style.kind==='character','invalid','Invalid named style kind');
    if(style.toggles!==undefined){assert(Array.isArray(style.toggles)&&style.toggles.length<=2&&new Set(style.toggles).size===style.toggles.length&&style.toggles.every(k=>(k==='bold'||k==='italic')&&!Object.hasOwn(style.properties as object,k)),'invalid','Invalid named style toggle');}
    if(style.kind!==undefined&&style.basedOn!==undefined){const parent=styles[String(style.basedOn)];assert(isRecord(parent)&&parent.kind===style.kind,'invalid','Named style inheritance crosses style kinds');}
    const path = new Set<string>([key]);
    let parent = style.basedOn;
    while (parent !== undefined) {
      assertId(parent); assert(!path.has(parent), 'invalid', 'Style inheritance cycle'); path.add(parent);
      assert(path.size <= 128, 'limit', 'Style inheritance depth exceeded');
      const ancestor = styles[parent]; assert(isRecord(ancestor), 'invalid', 'Missing parent style'); parent = ancestor.basedOn;
    }
  }
  validateTextStyle(model.defaultTextStyle);
  const resources = model.resources;
  const run = (candidate: unknown, textOnly = false): void => {
    const inline = object(candidate, ['id', 'type', 'text', 'styleId', 'style', 'resourceId', 'width', 'height']); id(inline.id);
    if (inline.type === 'text') {
      assert(!['resourceId', 'width', 'height'].some((key) => key in inline), 'invalid', 'Unexpected text run fields');
      text(inline.text); optionalId(inline.styleId, styles); if(inline.styleId!==undefined){const ref=styles[String(inline.styleId)];assert(isRecord(ref)&&(ref.kind===undefined||ref.kind==='character'),'invalid','Text run requires a character style');} if (inline.style !== undefined) validateTextStyle(inline.style);
    } else if (inline.type === 'image' && !textOnly) {
      assert(!['text', 'style', 'styleId'].some((key) => key in inline), 'invalid', 'Unexpected image fields');
      assert(inline.resourceId !== undefined, 'invalid', 'Missing image resource'); optionalId(inline.resourceId, resources);
      integer(inline.width, 1); integer(inline.height, 1);
    } else {
      assert(!textOnly && (inline.type === 'tab' || inline.type === 'line-break') && Object.keys(inline).length === 2, 'invalid', 'Unsupported inline model');
    }
  };
  const block = (candidate: unknown, depth = 0): void => {
    assert(depth <= 64, 'limit', 'Table nesting budget exceeded');
    const value = object(candidate, ['id', 'type', 'styleId', 'style', 'runs', 'columnWidths', 'rows', 'cellPadding', 'border']); id(value.id);
    if (value.type === 'paragraph') {
      assert(!['rows','columnWidths','cellPadding','border'].some(k=>k in value), 'invalid', 'Unexpected paragraph fields');
      optionalId(value.styleId, styles); if(value.styleId!==undefined){const ref=styles[String(value.styleId)];assert(isRecord(ref)&&(ref.kind===undefined||ref.kind==='paragraph'),'invalid','Paragraph requires a paragraph style');} if (value.style !== undefined) validateTextStyle(value.style);
      list(value.runs).forEach((item) => run(item));
    } else if (value.type === 'table') {
      assert(!['runs', 'style', 'styleId'].some((key) => key in value), 'invalid', 'Unexpected table fields');
      const widths = list(value.columnWidths, 1); widths.forEach((width) => integer(width, 1));
      assert(widths.length<=64&&widths.reduce<number>((sum,w)=>sum+Number(w),0)<=Number.MAX_SAFE_INTEGER,'limit','Table width budget exceeded');
      if(value.cellPadding!==undefined){const pad=object(value.cellPadding,['top','right','bottom','left']);for(const key of ['top','right','bottom','left'])integer(pad[key],0,12_700*144);}
      if(value.border!==undefined){const border=object(value.border,['width','color']);integer(border.width,1,12_700*12);validateColor(border.color);}
      const rows=list(value.rows,1);assert(rows.length<=500&&rows.length*widths.length<=10000,'limit','Table cell budget exceeded');rows.forEach((candidate) => {
        const row = object(candidate, ['id', 'cells']); id(row.id);
        const cells = list(row.cells); assert(cells.length === widths.length, 'invalid', 'Non-rectangular table');
        cells.forEach((candidate) => { const cell = object(candidate, ['id', 'blocks']); id(cell.id); list(cell.blocks, 1).forEach((item) => block(item, depth + 1)); });
      });
    } else assert(value.type === 'page-break' && Object.keys(value).length === 2, 'invalid', 'Invalid writer block');
  };
  if (model.kind === 'writer') {
    optionalId(model.defaultParagraphStyleId,styles);
    if(model.defaultParagraphStyleId!==undefined){const ref=styles[String(model.defaultParagraphStyleId)];assert(isRecord(ref)&&ref.kind==='paragraph','invalid','Default paragraph style must be a paragraph style');}
    list(model.sections, 1).forEach((candidate) => {
      const section = object(candidate, ['id', 'page', 'blocks']); id(section.id);
      const page = object(section.page, ['width', 'height', 'margins']); integer(page.width, 1); integer(page.height, 1);
      const margins = object(page.margins, ['top', 'right', 'bottom', 'left']);
      integer(margins.top, 0); integer(margins.right, 0); integer(margins.bottom, 0); integer(margins.left, 0);
      assert(margins.left + margins.right < page.width && margins.top + margins.bottom < page.height, 'invalid', 'Margins leave no content area');
      list(section.blocks, 1).forEach((candidate) => block(candidate));
    });
  } else if (model.kind === 'sheets') {
    assert(model.dateSystem === '1900' || model.dateSystem === '1904', 'invalid', 'Invalid workbook date system');
    const names = new Set<string>();
    list(model.sheets, 1).forEach((candidate) => {
      const sheet = object(candidate, ['id', 'name', 'cells', 'rows', 'columns', 'merges']); id(sheet.id); text(sheet.name, 31, false);
      assert(!/[\[\]:*?/\\\x00-\x1f]/.test(sheet.name) && !sheet.name.startsWith("'") && !sheet.name.endsWith("'"), 'invalid', 'Invalid sheet name');
      const name = sheet.name.toLowerCase(); assert(!names.has(name), 'invalid', 'Duplicate sheet name'); names.add(name);
      assert(isRecord(sheet.cells), 'invalid', 'Invalid sparse cell map');
      for (const [key, item] of Object.entries(sheet.cells)) {
        const match = /^([1-9]\d*):([1-9]\d*)$/.exec(key); assert(match, 'invalid', 'Invalid cell key'); assertCellAddress(Number(match[1]), Number(match[2]));
        validateCell(item, styles);
      }
      for (const [key, max, size] of [['rows', 1_048_576, 'height'], ['columns', 16_384, 'width']] as const) {
        const dimensions = sheet[key]; assert(isRecord(dimensions), 'invalid', 'Invalid sheet dimensions');
        for (const [address, candidate] of Object.entries(dimensions)) {
          assert(/^[1-9]\d*$/.test(address), 'invalid', 'Invalid dimension address'); integer(Number(address), 1, max);
          const dimension = object(candidate, [size, 'hidden']);
          if (dimension[size] !== undefined) integer(dimension[size], 1);
          if (dimension.hidden !== undefined) bool(dimension.hidden);
        }
      }
      const merges = list(sheet.merges);
      assert(merges.length <= 10_000, 'limit', 'Merge budget exceeded');
      // Sweep by row; active ranges need comparison only while their rows overlap.
      const validated = merges.map((candidate) => {
        const merge = object(candidate, ['row', 'column', 'rowCount', 'columnCount']);
        integer(merge.row, 1); integer(merge.column, 1); integer(merge.rowCount, 1); integer(merge.columnCount, 1);
        assertCellAddress(merge.row, merge.column); assertCellAddress(merge.row + merge.rowCount - 1, merge.column + merge.columnCount - 1);
        return merge as unknown as Sheet['merges'][number];
      }).sort((a, b) => a.row - b.row);
      let active: Sheet['merges'] = [];
      for (const merge of validated) {
        active = active.filter((previous) => previous.row + previous.rowCount > merge.row);
        assert(!active.some((previous) => previous.column < merge.column + merge.columnCount && merge.column < previous.column + previous.columnCount), 'invalid', 'Overlapping merged ranges');
        active.push(merge);
      }
    });
  } else if (model.kind === 'slides') {
    const size = object(model.size, ['width', 'height']); integer(size.width, 1); integer(size.height, 1);
    list(model.slides).forEach((candidate) => {
      const slide = object(candidate, ['id', 'background', 'shapes']); id(slide.id); validateColor(slide.background);
      list(slide.shapes).forEach((candidate) => {
        const shape = object(candidate, ['id', 'type', 'geometry', 'fill', 'stroke', 'runs', 'resourceId']); id(shape.id); validateGeometry(shape.geometry);
        if (shape.fill !== undefined) validateColor(shape.fill);
        if (shape.stroke !== undefined) { const stroke = object(shape.stroke, ['color', 'width']); validateColor(stroke.color); integer(stroke.width, 0); }
        if (shape.type === 'text') { assert(!('resourceId' in shape), 'invalid', 'Unexpected text shape fields'); list(shape.runs).forEach((item) => run(item, true)); }
        else if (shape.type === 'image') { assert(!('runs' in shape) && shape.resourceId !== undefined, 'invalid', 'Invalid image shape'); optionalId(shape.resourceId, resources); }
        else assert(['rect', 'ellipse', 'line'].includes(String(shape.type)) && !('runs' in shape) && !('resourceId' in shape), 'invalid', 'Invalid shape type');
      });
    });
  } else throw new OfficeError('invalid', 'Invalid document kind');
  return value as unknown as OfficeDocument;
}

export function validateCell(value: unknown, styles: Record<string, unknown> = {}): asserts value is Cell {
  const cell = object(value, ['value', 'styleId', 'numberFormat']); optionalId(cell.styleId, styles);
  if (cell.numberFormat !== undefined) text(cell.numberFormat, 1024);
  const item = object(cell.value, ['kind', 'value', 'source']);
  if (item.kind === 'blank') assert(Object.keys(item).length === 1, 'invalid', 'Blank cell has a payload');
  else if (item.kind === 'formula') {
    assert(!('value' in item), 'invalid', 'Formula cache is not source content'); text(item.source, 8192, false); assert(item.source.startsWith('=') && item.source.length > 1, 'invalid', 'Invalid formula source');
  } else {
    assert(!('source' in item), 'invalid', 'Unexpected cell source');
    if (item.kind === 'string') text(item.value, 32767);
    else if (item.kind === 'number') assert(typeof item.value === 'number' && Number.isFinite(item.value), 'invalid', 'Invalid numeric cell');
    else if (item.kind === 'boolean') bool(item.value);
    else assert(item.kind === 'error' && ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'].includes(String(item.value)), 'invalid', 'Invalid cell value');
  }
}

/** New models are format independent. Format adapters own source/provenance metadata. */
export function createDocument(kind: 'writer', documentId: string, title?: string): WriterDocument;
export function createDocument(kind: 'sheets', documentId: string, title?: string): SheetsDocument;
export function createDocument(kind: 'slides', documentId: string, title?: string): SlidesDocument;
export function createDocument(kind: DocumentKind, documentId: string, title?: string): OfficeDocument;
export function createDocument(kind: DocumentKind, documentId: string, title = 'Untitled'): OfficeDocument {
  assert(kind === 'writer' || kind === 'sheets' || kind === 'slides', 'invalid', 'Invalid document kind');
  const base: DocumentBase = { schemaVersion: 1, documentId, title, resources: {}, styles: {}, defaultTextStyle: { fontFamily: 'sans-serif', fontSize: 152_400 } };
  const document: OfficeDocument = kind === 'writer'
    ? { ...base, kind, sections: [{ id: 'section-1', page: { width: 7_560_000, height: 10_692_000, margins: { top: 914_400, right: 914_400, bottom: 914_400, left: 914_400 } }, blocks: [{ id: 'paragraph-1', type: 'paragraph', runs: [{ id: 'run-1', type: 'text', text: '' }] }] }] }
    : kind === 'sheets' ? { ...base, kind, dateSystem: '1900', sheets: [{ id: 'sheet-1', name: 'Sheet1', cells: {}, rows: {}, columns: {}, merges: [] }] }
      : { ...base, kind, size: { width: 12_192_000, height: 6_858_000 }, slides: [{ id: 'slide-1', background: { kind: 'rgba', red: 255, green: 255, blue: 255, alpha: 255 }, shapes: [] }] };
  return validateDocument(document);
}

export function cloneData<T>(value: T): T {
  if (Array.isArray(value)) return value.map(cloneData) as T;
  if (isRecord(value)) return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, cloneData(item)])) as T;
  return value;
}
export function freezeData<T>(value: T): T {
  if (value !== null && typeof value === 'object' && !Object.isFrozen(value)) {
    for (const item of Object.values(value)) freezeData(item);
    Object.freeze(value);
  }
  return value;
}

export interface SchemaMigration { from: number; to: number; migrate: (copy: unknown) => unknown }
/** Migration only receives a detached copy; failure never mutates the original snapshot. */
export function migrateDocument(value: unknown, migrations: readonly SchemaMigration[] = []): OfficeDocument {
  assertPlainData(value); assert(isRecord(value), 'invalid', 'Expected document'); integer(value.schemaVersion, 0);
  assert(value.schemaVersion <= 1, 'unsupported', 'Newer document schema cannot be opened');
  let copy: unknown = cloneData(value);
  const visited = new Set<number>();
  while (isRecord(copy) && copy.schemaVersion !== 1) {
    const version = copy.schemaVersion; integer(version, 0); assert(!visited.has(version), 'invalid', 'Migration cycle'); visited.add(version);
    const candidates = migrations.filter((item) => item.from === version && item.to > version && item.to <= 1);
    assert(candidates.length === 1, 'unsupported', 'No unambiguous schema migration');
    const migration = candidates[0]!;
    copy = migration.migrate(copy); assertPlainData(copy);
    assert(isRecord(copy) && copy.schemaVersion === migration.to, 'invalid', 'Migration returned wrong schema version');
  }
  return validateDocument(copy);
}
