import { validateImageResource } from './images.js';
import { OfficeError, assert, assertCellAddress, assertId, assertUnicode, isRecord, type DocumentKind } from '@office-ex/contracts';

export type Color = { kind: 'rgba'; red: number; green: number; blue: number; alpha: number } | { kind: 'theme'; name: string; tint?: number; tintSpace?: 'linear' };
/** Presentation-only paragraph formatting. Values retain DrawingML units. */
export type SlideBullet = {kind:'none'} | ({kind:'char';character:string}|{kind:'number';scheme:'arabicPeriod'|'arabicParenR'|'arabicParenBoth'|'arabicPlain';startAt?:number}) & {font?:string;color?:Color;sizePercent?:number;sizeEmu?:number};
export interface SlideParagraphFormat {level?:number;marginLeft?:number;marginRight?:number;firstLineIndent?:number;beforePercent?:number;afterPercent?:number;bullet?:SlideBullet}
/** Writer paragraph bottom rule. Null explicitly suppresses an inherited rule. */
export interface WriterParagraphRule {color:Color;width:number;space:number}
/** Bounded WordprocessingML numbering. References are native positive numIds;
 * zero explicitly cancels inherited numbering. Definitions are immutable per id. */
export interface WriterListReference {numId:number;level:number}
export interface WriterListLevel {format:'decimal'|'bullet';start:number;pattern:string;indent:number;hanging:number;alignment:'start'|'center'|'end';neverRestart?:boolean;symbolBullet?:boolean}
export interface WriterNumbering {nextId:number;lists:Record<string,{levels:WriterListLevel[]}>}
export interface TextStyle {
  fontFamily?: string; fontAscii?:string; fontHighAnsi?:string; fontEastAsia?:string; fontComplex?:string; fontSize?: number; complexFontSize?:number; bold?: boolean; italic?: boolean; underline?: boolean;
  color?: Color; language?:string; eastAsiaLanguage?:string; bidiLanguage?:string; runDirection?:'ltr'|'rtl'; paragraphDirection?:'ltr'|'rtl'; alignment?: 'start' | 'center' | 'end' | 'justify';
  slideParagraph?:SlideParagraphFormat;
  characterSpacing?:number; kerningThreshold?:number;
  rightIndent?:number; firstLineIndent?:number; contextualSpacing?:boolean;
  paragraphBottomBorder?:WriterParagraphRule|null;
  writerList?:WriterListReference;
  indent?: number; spaceBefore?: number; spaceAfter?: number; lineHeight?: number; lineMultiple?:number; minLineHeight?:number; keepNext?:boolean; keepLines?:boolean; pageBreakBefore?:boolean; outlineLevel?:number; widowControl?:boolean;
}
export interface NamedStyle { id: string; name: string; basedOn?: string; kind?: 'paragraph' | 'character'; toggles?: readonly ('bold' | 'italic')[]; properties: TextStyle }
export interface ResourceRef { id: string; sha256: string; mimeType: string; byteLength: number; dataBase64?:string; pixelWidth?:number; pixelHeight?:number }
export interface DocumentBase {
  schemaVersion: 1; documentId: string; title: string;
  resources: Record<string, ResourceRef>; styles: Record<string, NamedStyle>; defaultTextStyle: TextStyle;
}
export interface TextRun { id: string; type: 'text'; text: string; styleId?: string; style?: TextStyle }
export type Inline = TextRun | { id: string; type: 'tab' | 'line-break' } | { id: string; type: 'image'; resourceId: string; width: number; height: number; drawingName?:string; pictureName?:string };
export interface Paragraph { id: string; type: 'paragraph'; styleId?: string; style?: TextStyle; runs: Inline[] }
/** Merged cells occupy a rectangular grid; only the top-left anchor owns
 * content. Covered slots reference that anchor and never hold hidden blocks. */
export interface TableCell { id: string; blocks: WriterBlock[]; merge?: {rows:number;columns:number}; coveredBy?:string }
export interface Table { id: string; type: 'table'; nativeTable?: {styleId?:'TableNormal'|'TableGrid'; autoWidth?:true; autoLayout?:true; directPadding?:true; directBorder?:true; look?:number}; columnWidths: number[]; rows: { id: string; cells: TableCell[] }[]; cellPadding?: {top:number;right:number;bottom:number;left:number}; border?: {width:number;color:Color} }
export type WriterBlock = Paragraph | Table | { id: string; type: 'page-break' };
/** A bounded single-paragraph running story. PAGE is a typed field, never
 * source markup or a script. Text before/after it is plain Unicode. */
export interface WriterPageStory {kind:'header'|'footer';text:string;pageNumber?:true;after?:string;styleId?:'Header'|'Footer';alignment?:'start'|'center'|'end';runStyle?:TextStyle}
export interface WriterSection { id: string; storyRefs?:{header?:string;footer?:string}; pageNumberStart?:number; pageExtras?: {headerDistance?:number;footerDistance?:number;columnSpace?:number;docGridLinePitch?:number}; page: { width: number; height: number; margins: { top: number; right: number; bottom: number; left: number } }; blocks: WriterBlock[] }
export interface WriterDocument extends DocumentBase { kind: 'writer'; sections: WriterSection[]; defaultParagraphStyleId?:string; numbering?:WriterNumbering; pageStories?:Record<string,WriterPageStory> }

export type CellValue = { kind: 'blank' } | { kind: 'string'; value: string } | { kind: 'number'; value: number } | { kind: 'boolean'; value: boolean } | { kind: 'error'; value: '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' } | { kind: 'formula'; source: string };
export interface CellStyle extends TextStyle { fill?: Color; border?: {color:Color;width:number}; verticalAlignment?: 'top'|'center'|'bottom'; wrapText?: boolean }
export interface SheetView { rightToLeft?:boolean; frozenRows?:number; frozenColumns?:number; showGridLines?:boolean }
export interface Cell { value: CellValue; styleId?: string; numberFormat?: string; style?: CellStyle }
export interface Sheet {
  id: string; name: string; cells: Record<string, Cell>;
  rows: Record<string, { height?: number; hidden?: boolean }>;
  columns: Record<string, { width?: number; hidden?: boolean }>;
  merges: { row: number; column: number; rowCount: number; columnCount: number }[];
  view?: SheetView;
}
export interface SheetsDocument extends DocumentBase { kind: 'sheets'; dateSystem: '1900' | '1904'; sheets: Sheet[] }

export interface Geometry { x: number; y: number; width: number; height: number; rotation: number }
export type SlidePreset = 'rect'|'ellipse'|'line'|'roundRect'|'triangle'|'diamond'|'parallelogram';
export type SlideDash = 'solid'|'dash'|'dot'|'dashDot'|'lgDash'|'lgDashDot'|'lgDashDotDot'|'sysDash'|'sysDot'|'sysDashDot'|'sysDashDotDot';
export interface SlideArrow {type:'triangle'|'arrow';width?:'sm'|'med'|'lg';length?:'sm'|'med'|'lg'}
export interface SlideStroke {color:Color;width:number;dash?:SlideDash;head?:SlideArrow;tail?:SlideArrow}
export interface SlideCrop {left:number;top:number;right:number;bottom:number}
export interface SlideChildSpace {x:number;y:number;width:number;height:number}
/** DrawingML rectangle connection-site order: top, left, bottom, right. */
export interface SlideConnection {shapeId:string;site:0|1|2|3}
export interface SlideConnections {start:SlideConnection;end:SlideConnection}
interface ShapeBase { id: string; geometry: Geometry; fill?: Color; stroke?: SlideStroke; flipH?:boolean; flipV?:boolean }
export interface SlideTextFrame { autofit?:{mode:'none'}|{mode:'shape'}|{mode:'normal';fontScale:number;lineReduction:number}; wrap?: boolean; verticalAlignment?: 'top'|'center'|'bottom'; margins?: {top:number;right:number;bottom:number;left:number}; direction?: 'ltr'|'rtl'|'auto' }
export interface SlideThemeFontDetails { eastAsia?:string; complex?:string; scripts?:Record<string,string> }
export interface SlideTheme { name:string; colors:Record<string,Extract<Color,{kind:'rgba'}>>; majorFont?:string; minorFont?:string; majorFontDetails?:SlideThemeFontDetails; minorFontDetails?:SlideThemeFontDetails }
export type Shape = (ShapeBase & { type: 'text'; runs: TextRun[]; textStyle?: TextStyle; textFrame?: SlideTextFrame }) | (ShapeBase & { type: SlidePreset; connector?:true; connections?:SlideConnections }) | (ShapeBase & { type: 'image'; resourceId: string; crop?:SlideCrop }) | (ShapeBase & {type:'group';childSpace:SlideChildSpace;children:Shape[]});
export interface Slide { id: string; background: Color; shapes: Shape[] }
export interface SlidesDocument extends DocumentBase { kind: 'slides'; size: { width: number; height: number }; slides: Slide[]; theme?: SlideTheme }
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
  const item = object(value, ['kind', 'red', 'green', 'blue', 'alpha', 'name', 'tint', 'tintSpace']);
  if (item.kind === 'rgba') {
    assert(!('name' in item) && !('tint' in item) && !('tintSpace' in item), 'invalid', 'Unexpected RGBA fields');
    for (const key of ['red', 'green', 'blue', 'alpha']) integer(item[key], 0, 255);
  } else {
    assert(item.kind === 'theme' && !['red', 'green', 'blue', 'alpha'].some((key) => key in item), 'invalid', 'Invalid theme color');
    text(item.name, 128, false);
    if (item.tint !== undefined) assert(typeof item.tint === 'number' && Number.isFinite(item.tint) && item.tint >= -1 && item.tint <= 1, 'invalid', 'Invalid color tint');
    if(item.tintSpace!==undefined)assert(item.tintSpace==='linear'&&item.tint!==undefined,'invalid','Invalid tint color space');
  }
}

export function isDocumentLanguageTag(value:unknown):value is string{
  if(typeof value!=='string'||!value.length||value.length>128)return false;
  try{return Intl.getCanonicalLocales(value).length===1;}catch{return false;}
}
export function validateSlideParagraph(value:unknown):asserts value is SlideParagraphFormat{
  const format=object(value,['level','marginLeft','marginRight','firstLineIndent','beforePercent','afterPercent','bullet']);
  if(format.level!==undefined)integer(format.level,0,8);
  for(const key of ['marginLeft','marginRight'])if(format[key]!==undefined)integer(format[key],0,51206400);
  if(format.firstLineIndent!==undefined)integer(format.firstLineIndent,-51206400,51206400);
  for(const key of ['beforePercent','afterPercent'])if(format[key]!==undefined)integer(format[key],0,1000000);
  if(format.bullet!==undefined){
    const bullet=object(format.bullet,['kind','character','scheme','startAt','font','color','sizePercent','sizeEmu']);
    assert(['none','char','number'].includes(String(bullet.kind)),'invalid','Invalid slide bullet');
    if(bullet.kind==='none'){assert(Object.keys(bullet).length===1,'invalid','Unexpected disabled bullet format');return;}
    if(bullet.kind==='char'){assert(typeof bullet.character==='string'&&Array.from(bullet.character).length===1&&!/[\u0000-\u0020\u007f]/u.test(bullet.character),'invalid','Invalid bullet character');assert(bullet.scheme===undefined&&bullet.startAt===undefined,'invalid','Unexpected bullet numbering');}
    else{assert(bullet.character===undefined&&['arabicPeriod','arabicParenR','arabicParenBoth','arabicPlain'].includes(String(bullet.scheme)),'invalid','Invalid slide numbering scheme');if(bullet.startAt!==undefined)integer(bullet.startAt,1,32767);}
    if(bullet.font!==undefined)validateTextStyle({fontFamily:bullet.font});
    if(bullet.color!==undefined)validateColor(bullet.color);
    if(bullet.sizePercent!==undefined)integer(bullet.sizePercent,25000,400000);
    if(bullet.sizeEmu!==undefined)integer(bullet.sizeEmu,12700,5080000);
    assert(!(bullet.sizePercent!==undefined&&bullet.sizeEmu!==undefined),'invalid','Conflicting bullet sizes');
  }
}
export function validateTextStyle(value: unknown): asserts value is TextStyle {
  const style = object(value, ['fontFamily','fontAscii','fontHighAnsi','fontEastAsia','fontComplex','fontSize','complexFontSize','bold','italic','underline','color','alignment','indent','spaceBefore','spaceAfter','lineHeight','lineMultiple','minLineHeight','keepNext','keepLines','pageBreakBefore','outlineLevel','widowControl','language','eastAsiaLanguage','bidiLanguage','runDirection','paragraphDirection','slideParagraph','characterSpacing','kerningThreshold','rightIndent','firstLineIndent','contextualSpacing','paragraphBottomBorder','writerList']);
  for(const key of ['fontFamily','fontAscii','fontHighAnsi','fontEastAsia','fontComplex'])if(style[key]!==undefined){text(style[key],512,false);assert(!/[\x00-\x1f\x7f]/.test(style[key] as string),'invalid','Invalid font family');}
  if (style.fontSize !== undefined) integer(style.fontSize, 1, 12_700 * 1638);
  if(style.complexFontSize!==undefined)integer(style.complexFontSize,1,12_700*1638);
  if(style.lineMultiple!==undefined)integer(style.lineMultiple,1,31680);
  if(style.minLineHeight!==undefined)integer(style.minLineHeight,0);
  assert(['lineHeight','lineMultiple','minLineHeight'].filter(k=>style[k]!==undefined).length<=1,'invalid','Conflicting line spacing modes');
  if(style.outlineLevel!==undefined)integer(style.outlineLevel,0,9);
  for(const key of ['keepNext','keepLines','pageBreakBefore','widowControl'])if(style[key]!==undefined)bool(style[key]);
  for (const key of ['bold', 'italic', 'underline']) if (style[key] !== undefined) bool(style[key]);
  if (style.color !== undefined) validateColor(style.color);
  for(const key of ['language','eastAsiaLanguage','bidiLanguage'])if(style[key]!==undefined)assert(isDocumentLanguageTag(style[key]),'invalid','Invalid document language tag');
  for(const key of ['runDirection','paragraphDirection'])if(style[key]!==undefined)assert(style[key]==='ltr'||style[key]==='rtl','invalid','Invalid document direction');
  if (style.alignment !== undefined) assert(['start', 'center', 'end', 'justify'].includes(String(style.alignment)), 'invalid', 'Invalid text alignment');
  if (style.indent !== undefined) integer(style.indent);
  if(style.characterSpacing!==undefined)integer(style.characterSpacing,-31680*635,31680*635);
  if(style.kerningThreshold!==undefined)integer(style.kerningThreshold,0,3276*6350);
  if(style.rightIndent!==undefined)integer(style.rightIndent,0,31680*635);
  if(style.firstLineIndent!==undefined)integer(style.firstLineIndent,-31680*635,31680*635);
  if(style.contextualSpacing!==undefined)bool(style.contextualSpacing);
  if(style.writerList!==undefined){const ref=object(style.writerList,['numId','level']);integer(ref.numId,0,1000000000);integer(ref.level,0,8);assert(ref.numId!==0||ref.level===0,'invalid','Disabled numbering must use level zero');}
  if(style.paragraphBottomBorder!==undefined&&style.paragraphBottomBorder!==null){
    const border=object(style.paragraphBottomBorder,['color','width','space']);validateColor(border.color);integer(border.width,2*1587,96*1588);integer(border.space,0,31*12700);
  }
  if(style.slideParagraph!==undefined){validateSlideParagraph(style.slideParagraph);const p=style.slideParagraph as SlideParagraphFormat;assert(!(p.beforePercent!==undefined&&style.spaceBefore!==undefined)&&!(p.afterPercent!==undefined&&style.spaceAfter!==undefined),'invalid','Conflicting paragraph spacing');}
  for (const key of ['spaceBefore', 'spaceAfter', 'lineHeight']) if (style[key] !== undefined) integer(style[key], 0);
}

/** Positive crop fractions of the original source image, in 1/1000 %. */
export function validateSlideCrop(value:unknown):asserts value is SlideCrop{
  const crop=object(value,['left','top','right','bottom']);
  for(const k of ['left','top','right','bottom'])integer(crop[k],0,99999);
  assert(Number(crop.left)+Number(crop.right)<100000&&Number(crop.top)+Number(crop.bottom)<100000,'invalid','Image crop removes the entire image');
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
  const model = object(value, [...fields, ...(value.kind === 'writer' ? ['sections','defaultParagraphStyleId','numbering','pageStories'] : value.kind === 'sheets' ? ['dateSystem', 'sheets'] : value.kind === 'slides' ? ['size', 'slides', 'theme'] : [])]);
  assertId(model.documentId); text(model.title, 1024, false);
  const ids = new Set<string>();
  const id = (candidate: unknown): void => { assertId(candidate); assert(!ids.has(candidate), 'invalid', 'Duplicate stable node identifier'); ids.add(candidate); };
  assert(isRecord(model.resources), 'invalid', 'Invalid resources');
  let resourceBytes = 0, packagedImageBytes=0;
  for (const [key, item] of Object.entries(model.resources)) {
    assertId(key);
    const resource = object(item, ['id', 'sha256', 'mimeType', 'byteLength', 'dataBase64', 'pixelWidth', 'pixelHeight']);
    assert(resource.id === key, 'invalid', 'Resource identifier mismatch');
    assert(typeof resource.sha256 === 'string' && /^[a-f0-9]{64}$/.test(resource.sha256), 'invalid', 'Invalid resource digest');
    text(resource.mimeType, 255, false);
    assert(/^[a-z0-9!#$&^_.+-]+\/[a-z0-9!#$&^_.+-]+$/i.test(resource.mimeType), 'invalid', 'Invalid resource media type');
    integer(resource.byteLength, 0, 128 * 1024 * 1024);
    if(resource.dataBase64!==undefined||resource.pixelWidth!==undefined||resource.pixelHeight!==undefined){
      validateImageResource(resource);packagedImageBytes+=resource.byteLength;
      assert(packagedImageBytes<=16*1024*1024,'limit','Packaged image budget exceeded');
    }
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
  const numbering=model.kind==='writer'&&model.numbering!==undefined?object(model.numbering,['nextId','lists']):undefined;
  if(numbering){
    integer(numbering.nextId,1,1000000001);assert(isRecord(numbering.lists),'invalid','Invalid numbering definitions');const entries=Object.entries(numbering.lists);assert(entries.length<=2048,'limit','Numbering definition budget exceeded');
    for(const [key,item] of entries){
      assert(/^[1-9]\d*$/.test(key)&&Number(key)<Number(numbering.nextId),'invalid','Invalid numbering identity');
      const def=object(item,['levels']),levels=list(def.levels,1);assert(levels.length<=9,'invalid','Numbering level budget exceeded');
      for(const [i,candidate]of levels.entries()){
        const level=object(candidate,['format','start','pattern','indent','hanging','alignment','neverRestart','symbolBullet']);
        assert(level.format==='decimal'||level.format==='bullet','unsupported','Unsupported list format');integer(level.start,1,32767);text(level.pattern,64,false);
        const pattern=String(level.pattern);
        assert(!/[\x00-\x1f\x7f]/.test(pattern),'invalid','List pattern controls');
        if(level.format==='bullet')assert(Array.from(pattern).length===1&&!/[\s%]/u.test(pattern),'invalid','Invalid list bullet');
        else assert(/^(?:[^%]|%[1-9])+$/.test(pattern)&&pattern.includes('%'+(i+1))&&[...pattern.matchAll(/%([1-9])/g)].every(m=>Number(m[1])<=i+1),'invalid','Invalid decimal list pattern');
        if(level.format==='decimal')assert([...pattern.matchAll(/%([1-9])/g)].every(m=>(levels[Number(m[1])-1] as WriterListLevel).format==='decimal'),'unsupported','Mixed composite numbering unsupported');
        integer(level.indent,635,31680*635);integer(level.hanging,635,Number(level.indent));assert(['start','center','end'].includes(String(level.alignment)),'invalid','Invalid list alignment');
        if(level.neverRestart!==undefined)bool(level.neverRestart);
        if(level.symbolBullet!==undefined)assert(level.format==='bullet'&&pattern==='•'&&level.symbolBullet===true,'invalid','Invalid Symbol bullet fallback');
      }
    }
  }
  function listReference(style:unknown,allowed:boolean):void{
    const ref=(style as TextStyle|undefined)?.writerList;if(!ref)return;
    assert(allowed,'invalid','Writer numbering is paragraph-only');
    if(ref.numId){const def=(numbering?.lists as WriterNumbering['lists']|undefined)?.[String(ref.numId)];assert(def&&def.levels[ref.level],'unsupported','Missing or unsupported numbering definition');}
  }
  listReference(model.defaultTextStyle,model.kind==='writer');
  for(const item of Object.values(styles)){const style=item as NamedStyle;listReference(style.properties,model.kind==='writer'&&style.kind==='paragraph');}
  const resources = model.resources;
  const run = (candidate: unknown, textOnly = false): void => {
    const inline = object(candidate, ['id', 'type', 'text', 'styleId', 'style', 'resourceId', 'width', 'height', 'drawingName', 'pictureName']); id(inline.id);
    if (inline.type === 'text') {
      assert(!['resourceId', 'width', 'height', 'drawingName', 'pictureName'].some((key) => key in inline), 'invalid', 'Unexpected text run fields');
      text(inline.text); optionalId(inline.styleId, styles); if(inline.styleId!==undefined){const ref=styles[String(inline.styleId)];assert(isRecord(ref)&&(ref.kind===undefined||ref.kind==='character'),'invalid','Text run requires a character style');} if (inline.style !== undefined){validateTextStyle(inline.style);listReference(inline.style,false);}
    } else if (inline.type === 'image' && !textOnly) {
      assert(!['text', 'style', 'styleId'].some((key) => key in inline), 'invalid', 'Unexpected image fields');
      assert(inline.resourceId !== undefined, 'invalid', 'Missing image resource'); optionalId(inline.resourceId, resources);
      integer(inline.width, 1); integer(inline.height, 1);if(inline.drawingName!==undefined)text(inline.drawingName,512);if(inline.pictureName!==undefined)text(inline.pictureName,512);
    } else {
      assert(!textOnly && (inline.type === 'tab' || inline.type === 'line-break') && Object.keys(inline).length === 2, 'invalid', 'Unsupported inline model');
    }
  };
  const block = (candidate: unknown, depth = 0): void => {
    assert(depth <= 64, 'limit', 'Table nesting budget exceeded');
    const value = object(candidate, ['id', 'type', 'styleId', 'style', 'runs', 'columnWidths', 'rows', 'cellPadding', 'border', 'nativeTable']); id(value.id);
    if (value.type === 'paragraph') {
      assert(!['rows','columnWidths','cellPadding','border','nativeTable'].some(k=>k in value), 'invalid', 'Unexpected paragraph fields');
      optionalId(value.styleId, styles); if(value.styleId!==undefined){const ref=styles[String(value.styleId)];assert(isRecord(ref)&&(ref.kind===undefined||ref.kind==='paragraph'),'invalid','Paragraph requires a paragraph style');} if (value.style !== undefined){validateTextStyle(value.style);listReference(value.style,true);}
      list(value.runs).forEach((item) => run(item));
    } else if (value.type === 'table') {
      assert(!['runs', 'style', 'styleId'].some((key) => key in value), 'invalid', 'Unexpected table fields');
      if(value.nativeTable!==undefined){const n=object(value.nativeTable,['styleId','autoWidth','autoLayout','directPadding','directBorder','look']);assert(Object.keys(n).length>0,'invalid','Empty native table metadata');if(n.styleId!==undefined)assert(n.styleId==='TableNormal'||n.styleId==='TableGrid','invalid','Unsupported native table style');for(const key of ['autoWidth','autoLayout','directPadding','directBorder'])if(n[key]!==undefined)assert(n[key]===true,'invalid','Invalid native table flag');if(n.look!==undefined){integer(n.look,0,0x7e0);assert((Number(n.look)&~0x7e0)===0,'invalid','Invalid table look mask');}}
      const widths = list(value.columnWidths, 1); widths.forEach((width) => integer(width, 1));
      assert(widths.length<=64&&widths.reduce<number>((sum,w)=>sum+Number(w),0)<=Number.MAX_SAFE_INTEGER,'limit','Table width budget exceeded');
      if(value.cellPadding!==undefined){const pad=object(value.cellPadding,['top','right','bottom','left']);for(const key of ['top','right','bottom','left'])integer(pad[key],0,12_700*144);}
      if(value.border!==undefined){const border=object(value.border,['width','color']);integer(border.width,1,12_700*12);validateColor(border.color);}
      const rows=list(value.rows,1);assert(rows.length<=500&&rows.length*widths.length<=10000,'limit','Table cell budget exceeded');
      const coverage:(string|undefined)[][]=rows.map(()=>Array<string|undefined>(widths.length).fill(undefined));
      rows.forEach((candidate,ri) => {
        const row = object(candidate, ['id', 'cells']); id(row.id);
        const cells = list(row.cells); assert(cells.length === widths.length, 'invalid', 'Non-rectangular table');
        cells.forEach((candidate,ci) => {
          const cell = object(candidate, ['id', 'blocks','merge','coveredBy']); id(cell.id);
          const expected=coverage[ri]![ci];
          if(expected!==undefined){
            assert(cell.coveredBy===expected&&cell.merge===undefined&&list(cell.blocks).length===0,'invalid','Invalid or content-bearing covered table slot');
            return;
          }
          assert(cell.coveredBy===undefined,'invalid','Dangling covered table slot');
          if(cell.merge!==undefined){
            const m=object(cell.merge,['rows','columns']);integer(m.rows,1,rows.length-ri);integer(m.columns,1,widths.length-ci);
            const h=Number(m.rows),w=Number(m.columns);assert(h*w>1,'invalid','Empty table merge');
            for(let r=ri;r<ri+h;r++)for(let c=ci;c<ci+w;c++){
              if(r===ri&&c===ci)continue;
              assert(coverage[r]![c]===undefined,'invalid','Overlapping table merges');coverage[r]![c]=String(cell.id);
            }
          }
          list(cell.blocks,1).forEach(item=>block(item,depth+1));
        });
      });
    } else assert(value.type === 'page-break' && Object.keys(value).length === 2, 'invalid', 'Invalid writer block');
  };
  if (model.kind === 'writer') {
    optionalId(model.defaultParagraphStyleId,styles);
    if(model.defaultParagraphStyleId!==undefined){const ref=styles[String(model.defaultParagraphStyleId)];assert(isRecord(ref)&&ref.kind==='paragraph','invalid','Default paragraph style must be a paragraph style');}
    const stories=model.pageStories;
    if(stories!==undefined){
      assert(isRecord(stories)&&Object.keys(stories).length>0&&Object.keys(stories).length<=256,'limit','Running story budget exceeded');
      for(const [key,item]of Object.entries(stories)){
        assertId(key);assert(/^[A-Za-z_][A-Za-z0-9_.-]{0,127}$/.test(key)&&!['__proto__','constructor','prototype'].includes(key),'invalid','Invalid story relationship identity');
        const s=object(item,['kind','text','pageNumber','after','styleId','alignment','runStyle']);
        assert(s.kind==='header'||s.kind==='footer','invalid','Invalid running story kind');
        text(s.text,4096);assert(!/[\x00-\x1f\x7f]/.test(String(s.text)),'invalid','Running story must be single-line text');
        if(s.pageNumber!==undefined)assert(s.pageNumber===true,'invalid','Invalid page field');
        if(s.after!==undefined){assert(s.pageNumber===true,'invalid','Text after missing PAGE field');text(s.after,4096,false);assert(!/[\x00-\x1f\x7f]/.test(String(s.after)),'invalid','Running story must be single-line text');}
        if(s.styleId!==undefined)assert(s.styleId===(s.kind==='header'?'Header':'Footer'),'invalid','Unexpected running story style');
        if(s.alignment!==undefined)assert(['start','center','end'].includes(String(s.alignment)),'invalid','Invalid story alignment');
        if(s.runStyle!==undefined){validateTextStyle(s.runStyle);assert(Object.keys(s.runStyle as object).every(k=>['fontFamily','fontAscii','fontHighAnsi','fontEastAsia','fontComplex','fontSize','complexFontSize','bold','italic','underline','color','language','eastAsiaLanguage','bidiLanguage','runDirection','characterSpacing','kerningThreshold'].includes(k)),'unsupported','Unsupported running story run format');}
      }
    }
    const sections=list(model.sections,1);assert(sections.length<=128,'limit','Writer section budget exceeded');sections.forEach((candidate) => {
      const section = object(candidate, ['id', 'page', 'blocks', 'pageExtras','storyRefs','pageNumberStart']); id(section.id);
      if(section.pageNumberStart!==undefined)integer(section.pageNumberStart,1,32767);
      if(section.storyRefs!==undefined){
        const refs=object(section.storyRefs,['header','footer']);assert(Object.keys(refs).length>0,'invalid','Empty running story references');
        for(const kind of ['header','footer'] as const)if(refs[kind]!==undefined){const key=refs[kind];assertId(key);assert(isRecord(stories)&&Object.hasOwn(stories,key)&&(stories[key] as WriterPageStory).kind===kind,'invalid','Missing or mismatched running story reference');}
      }
      if(section.pageExtras!==undefined){const extra=object(section.pageExtras,['headerDistance','footerDistance','columnSpace','docGridLinePitch']);for(const n of Object.values(extra))integer(n,0,31680*635);}
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
      const sheet = object(candidate, ['id', 'name', 'cells', 'rows', 'columns', 'merges', 'view']); id(sheet.id); text(sheet.name, 31, false);
      if(sheet.view!==undefined){const view=object(sheet.view,['rightToLeft','frozenRows','frozenColumns','showGridLines']);for(const field of ['rightToLeft','showGridLines'])if(view[field]!==undefined)bool(view[field]);if(view.frozenRows!==undefined)integer(view.frozenRows,0,1048575);if(view.frozenColumns!==undefined)integer(view.frozenColumns,0,16383);}
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
    if (model.theme !== undefined) {
      const theme = object(model.theme, ['name','colors','majorFont','minorFont','majorFontDetails','minorFontDetails']); text(theme.name,256,false);
      assert(isRecord(theme.colors) && Object.keys(theme.colors).length<=64,'invalid','Invalid slide theme colors');
      for(const [name,color] of Object.entries(theme.colors)){assert(/^[A-Za-z0-9_-]{1,128}$/.test(name),'invalid','Invalid theme color name');validateColor(color);assert(color.kind==='rgba','invalid','Theme colors cannot reference a theme');}
      for(const key of ['majorFont','minorFont'])if(theme[key]!==undefined)validateTextStyle({fontFamily:theme[key]});
      for(const key of ['majorFontDetails','minorFontDetails'])if(theme[key]!==undefined){const details=object(theme[key],['eastAsia','complex','scripts']);for(const name of ['eastAsia','complex'])if(details[name]!==undefined)validateTextStyle({fontFamily:details[name]});if(details.scripts!==undefined){assert(isRecord(details.scripts)&&Object.keys(details.scripts).length<=128,'invalid','Invalid slide script font map');for(const [script,font]of Object.entries(details.scripts)){assert(/^[A-Z][a-z]{3}$/.test(script),'invalid','Invalid theme font script');validateTextStyle({fontFamily:font});}}}
    }
    let shapeCount=0;
    list(model.slides).forEach((candidate) => {
      const slide = object(candidate, ['id', 'background', 'shapes']); id(slide.id); validateColor(slide.background);
      const validateShape=(candidate:unknown,depth:number):void=>{
        const shape = object(candidate, ['id', 'type', 'geometry', 'fill', 'stroke', 'runs', 'resourceId', 'textStyle', 'textFrame','flipH','flipV','connector','connections','crop','childSpace','children']); id(shape.id); validateGeometry(shape.geometry);
        assert(++shapeCount<=100000&&depth<=16,'limit','Presentation shape tree budget exceeded');
        if(shape.flipH!==undefined)bool(shape.flipH);if(shape.flipV!==undefined)bool(shape.flipV);
        if(shape.connector!==undefined)assert(shape.type==='line'&&shape.connector===true,'invalid','Invalid connector shape');
        if(shape.connections!==undefined){
          assert(depth===0&&shape.type==='line'&&shape.connector===true,'unsupported','Bound connectors must be top-level straight connectors');
          const refs=object(shape.connections,['start','end']),start=object(refs.start,['shapeId','site']),end=object(refs.end,['shapeId','site']);
          for(const ref of [start,end]){assertId(ref.shapeId);integer(ref.site,0,3);}
          assert(start.shapeId!==end.shapeId&&start.shapeId!==shape.id&&end.shapeId!==shape.id,'invalid','Invalid connector endpoints');
        }
        if(shape.type!=='group')assert(!('childSpace' in shape)&&!('children' in shape),'invalid','Unexpected group fields');
        if(shape.type!=='image')assert(!('crop' in shape),'invalid','Unexpected image crop');
        if(shape.type==='text')assert(!shape.flipH&&!shape.flipV,'unsupported','Flipped text is outside presentation profile');
        if (shape.fill !== undefined) validateColor(shape.fill);
        if (shape.stroke !== undefined) {
          const stroke=object(shape.stroke,['color','width','dash','head','tail']);validateColor(stroke.color);integer(stroke.width,0,12700*1000);
          if(stroke.dash!==undefined)assert(['solid','dash','dot','dashDot','lgDash','lgDashDot','lgDashDotDot','sysDash','sysDot','sysDashDot','sysDashDotDot'].includes(String(stroke.dash)),'invalid','Invalid presentation dash');
          for(const end of ['head','tail'])if(stroke[end]!==undefined){const arrow=object(stroke[end],['type','width','length']);assert(shape.type==='line'&&['triangle','arrow'].includes(String(arrow.type)),'invalid','Unsupported arrow shape');for(const key of ['width','length'])if(arrow[key]!==undefined)assert(['sm','med','lg'].includes(String(arrow[key])),'invalid','Invalid arrow size');}
        }
        if (shape.type === 'text') {
          assert(!('resourceId' in shape), 'invalid', 'Unexpected text shape fields'); list(shape.runs).forEach((item) => run(item, true));
          if(shape.textStyle!==undefined)validateTextStyle(shape.textStyle);
          if(shape.textFrame!==undefined){
            const frame=object(shape.textFrame,['wrap','verticalAlignment','margins','direction','autofit']);
            if(frame.autofit!==undefined){const fit=object(frame.autofit,['mode','fontScale','lineReduction']);if(fit.mode==='none'||fit.mode==='shape')assert(Object.keys(fit).length===1,'invalid','Unexpected text fit values');else{assert(fit.mode==='normal','invalid','Unsupported text fit mode');integer(fit.fontScale,1000,100000);integer(fit.lineReduction,0,99000);}}
            if((frame.autofit as {mode?:string}|undefined)?.mode==='shape'){
              const geometry=shape.geometry as Geometry;
              assert(depth===0&&geometry.rotation===0&&(!frame.verticalAlignment||frame.verticalAlignment==='top'),'unsupported','Shape-growing textbox must be top-level, unrotated and top-anchored');
            }
            if(frame.wrap!==undefined)bool(frame.wrap);
            if(frame.verticalAlignment!==undefined)assert(['top','center','bottom'].includes(String(frame.verticalAlignment)),'invalid','Invalid text frame alignment');
            if(frame.direction!==undefined)assert(['ltr','rtl','auto'].includes(String(frame.direction)),'invalid','Invalid text frame direction');
            if(frame.margins!==undefined){const margins=object(frame.margins,['top','right','bottom','left']);for(const key of ['top','right','bottom','left'])integer(margins[key],0);}
          }
        }
        else if (shape.type === 'image') { assert(!('textStyle' in shape)&&!('textFrame' in shape)&&!('runs' in shape) && shape.resourceId !== undefined, 'invalid', 'Invalid image shape'); optionalId(shape.resourceId, resources);if(shape.crop!==undefined)validateSlideCrop(shape.crop); }
        else if(shape.type==='group'){
          assert(!('textStyle' in shape)&&!('textFrame' in shape)&&!('runs' in shape)&&!('resourceId' in shape)&&!('fill' in shape)&&!('stroke' in shape),'invalid','Invalid group fields');
          const space=object(shape.childSpace,['x','y','width','height']);integer(space.x,-2147483648,2147483647);integer(space.y,-2147483648,2147483647);integer(space.width,1,2147483647);integer(space.height,1,2147483647);
          const geometry=shape.geometry as Geometry;assert(geometry.width>0&&geometry.height>0&&geometry.width/Number(space.width)<=1000&&geometry.height/Number(space.height)<=1000&&geometry.width/Number(space.width)>=.001&&geometry.height/Number(space.height)>=.001,'invalid','Group scale outside profile');
          const children=list(shape.children);assert(children.length>0&&children.length<=10000,'limit','Invalid group child count');children.forEach(child=>validateShape(child,depth+1));
        }
        else assert(['rect','ellipse','line','roundRect','triangle','diamond','parallelogram'].includes(String(shape.type)) && !('textStyle' in shape)&&!('textFrame' in shape)&& !('runs' in shape) && !('resourceId' in shape), 'invalid', 'Invalid shape type');
      };
      list(slide.shapes).forEach(candidate=>validateShape(candidate,0));
      const siblings=new Map((slide.shapes as Shape[]).map(shape=>[shape.id,shape]));
      for(const shape of slide.shapes as Shape[])if(shape.type==='line'&&shape.connections){
        for(const ref of [shape.connections.start,shape.connections.end]){const target=siblings.get(ref.shapeId);assert(target&&(target.type==='rect'||target.type==='text')&&target.geometry.width>0&&target.geometry.height>0,'unsupported','Connector endpoint requires a top-level rectangle or textbox');}
      }
    });
  } else throw new OfficeError('invalid', 'Invalid document kind');
  return value as unknown as OfficeDocument;
}

export function validateCell(value: unknown, styles: Record<string, unknown> = {}): asserts value is Cell {
  const cell = object(value, ['value', 'styleId', 'numberFormat', 'style']); optionalId(cell.styleId, styles);
  if(cell.style!==undefined){assert(isRecord(cell.style),'invalid','Invalid cell style');const {fill,border,verticalAlignment,wrapText,...font}=cell.style;validateTextStyle(font);if(fill!==undefined)validateColor(fill);if(border!==undefined){const edge=object(border,['color','width']);validateColor(edge.color);integer(edge.width,1,12700*12);}if(verticalAlignment!==undefined)assert(['top','center','bottom'].includes(String(verticalAlignment)),'invalid','Invalid vertical cell alignment');if(wrapText!==undefined)bool(wrapText);}
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
