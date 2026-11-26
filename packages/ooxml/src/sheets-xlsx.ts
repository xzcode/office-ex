// SPDX-License-Identifier: MIT
/** Closed Transitional XLSX profile. Unknown dependent parts/markup are not editable.
 * Source: ECMA-376 Part 1 SpreadsheetML, Part 2 OPC. No external Office engine.
 * Calculated caches are derived from the shared Office EX workbook model.
 */
import { XMLSerializer, type Element, type Document } from '@xmldom/xmldom';
import { assert, assertCellAddress, OfficeError } from '@office-ex/contracts';
import { createDocument, validateDocument, type SheetsDocument, type Sheet, type Cell, type CellStyle, type CellBorderSide,type CellBorderLine,type Color, type TextStyle } from '@office-ex/core';
import { columnName, columnNumber, parseFormula, formulaIsSupported, WorkbookCalculator, calculationKey, type CalculatedValue, type WorkbookSnapshot } from '@office-ex/sheets';
import { openOpc, type OpcPackage, type OpcRelationship } from './opc.js';
import { parseXml, childElements } from './xml.js';
import { writeZip } from './zip.js';
import type { ContainerOptions } from './limits.js';

const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG='http://schemas.openxmlformats.org/package/2006';
const CT='application/vnd.openxmlformats-officedocument.spreadsheetml.';
const DECL='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const XMLNS='http://www.w3.org/2000/xmlns/';
const XML='http://www.w3.org/XML/1998/namespace';
const A='http://schemas.openxmlformats.org/drawingml/2006/main',CP=PKG+'/metadata/core-properties',EP='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const CORE_REL=PKG+'/relationships/metadata/core-properties',EXT_REL=REL+'/extended-properties';
const VT='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
const encoder=new TextEncoder(),serializer=new XMLSerializer();
const xmlBytes=(doc:Document)=>encoder.encode(serializer.serializeToString(doc));
const canonical=(v:unknown):string=>v===undefined?'undefined':v===null||typeof v!=='object'?JSON.stringify(v):Array.isArray(v)?'['+v.map(canonical).join(',')+']':'{'+Object.entries(v).filter(([,v])=>v!==undefined).sort(([a],[b])=>a.localeCompare(b)).map(([k,v])=>JSON.stringify(k)+':'+canonical(v)).join(',')+'}';
const same=(a:unknown,b:unknown)=>canonical(a)===canonical(b);
const relationFile=(path:string)=>{const i=path.lastIndexOf('/');return path.slice(0,i+1)+'_rels/'+path.slice(i+1)+'.rels';};

const builtins:Readonly<Record<number,string>>={0:'General',1:'0',2:'0.00',3:'#,##0',4:'#,##0.00',9:'0%',10:'0.00%',11:'0.00E+00',14:'mm-dd-yy',20:'h:mm',21:'h:mm:ss',22:'m/d/yy h:mm',49:'@'};
const FONT_KEYS=['fontFamily','fontSize','bold','italic','underline','color'] as const;
const STYLE_KEYS=[...FONT_KEYS,'fill','border','borderEdges','alignment','verticalAlignment','wrapText'];
const ERROR=new Set(['#NULL!','#DIV/0!','#VALUE!','#REF!','#NAME?','#NUM!','#N/A']);
function requireProfile(condition:unknown,detail:string):asserts condition{if(!condition)throw new OfficeError('unsupported','XLSX profile: '+detail);}
function esc(s:string):string{requireProfile(!/[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]/.test(s),'XML text contains unsupported controls');return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;').replace(/\r/g,'&#13;');}
function encodeText(s:string):string{return esc(s.replace(/_x[0-9a-f]{4}_/gi,m=>'_x005F_'+m.slice(1)));}
function decodeText(s:string):string{return s.replace(/_x([0-9a-f]{4})_/gi,(_,hex:string)=>String.fromCharCode(parseInt(hex,16)));}
function attrs(el:Element,allowed:readonly string[]):void{
  for(let i=0;i<el.attributes.length;i++){const a=el.attributes.item(i)!;if(a.namespaceURI===XMLNS)continue;
    const key=a.namespaceURI===REL?'r:'+a.localName:a.namespaceURI===XML?'xml:'+a.localName:!a.namespaceURI?a.name:'?'+a.name;
    requireProfile(allowed.includes(key),'Unknown attribute '+key+' on '+el.localName);
  }
}
function nodes(el:Element,allowed:readonly string[]):Element[]{for(let n=el.firstChild;n;n=n.nextSibling)requireProfile(n.nodeType===1||n.nodeType===3&&!(n.nodeValue??'').trim(),'Unexpected XML content');const children=childElements(el);for(const e of children)requireProfile(e.namespaceURI===NS&&allowed.includes(e.localName??''),'Unknown element '+e.localName+' in '+el.localName);return children;}
function root(bytes:Uint8Array,name:string,options:ContainerOptions):Element{const el=parseXml(bytes,options).documentElement;requireProfile(el?.namespaceURI===NS&&el.localName===name,'Unexpected XML root');return el;}
function get(el:Element,name:string):string|undefined{return el.hasAttribute(name)?el.getAttribute(name)??undefined:undefined;}
function one(items:Element[],name:string,required=true):Element|undefined{const found=items.filter(x=>x.localName===name);requireProfile(found.length<2&&(!required||found.length===1),'Invalid '+name+' cardinality');return found[0];}
function integer(s:string|undefined,min:number,max:number):number{requireProfile(s!==undefined&&/^\d+$/.test(s),'Invalid integer');const v=Number(s);requireProfile(Number.isSafeInteger(v)&&v>=min&&v<=max,'Integer out of range');return v;}
function decimal(s:string|undefined,min=-Number.MAX_VALUE,max=Number.MAX_VALUE):number{requireProfile(s!==undefined&&/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(s),'Invalid numeric value');const n=Number(s);requireProfile(Number.isFinite(n)&&n>=min&&n<=max,'Numeric value out of range');return n;}
function flag(s:string|undefined,fallback=false):boolean{if(s===undefined)return fallback;requireProfile(['0','1','true','false'].includes(s),'Invalid boolean');return s==='1'||s==='true';}
function leaf(e:Element,allowed:readonly string[]=[]):string{attrs(e,allowed);for(let n=e.firstChild;n;n=n.nextSibling)requireProfile(n.nodeType===3||n.nodeType===4,'Unexpected nested markup');return e.textContent??'';}
function plain(e:Element,allowed:readonly string[]=[]):void{requireProfile(leaf(e,allowed)==='','Unexpected leaf text');}
function coordinate(s:string|undefined):[number,number]{const m=/^([A-Z]{1,3})([1-9]\d*)$/.exec(s??'');requireProfile(m,'Invalid cell address');const row=Number(m[2]),col=columnNumber(m[1]);assertCellAddress(row,col);return [row,col];}
function range(s:string|undefined){const [a,b,...rest]=(s??'').split(':');requireProfile(a&&b&&!rest.length,'Invalid range');const [r,c]=coordinate(a),[rr,cc]=coordinate(b);requireProfile(rr>=r&&cc>=c,'Invalid range order');return {row:r,column:c,rowCount:rr-r+1,columnCount:cc-c+1};}
function address(row:number,col:number){return columnName(col)+row;}
function rgb(c:Color):string{requireProfile(c.kind==='rgba'&&c.alpha===255,'Only opaque RGB cell colors supported');return 'FF'+[c.red,c.green,c.blue].map(n=>n.toString(16).padStart(2,'0')).join('').toUpperCase();}
type Rgba=Extract<Color,{kind:'rgba'}>;
interface ColorContext {palette?:Rgba[];indexed?:Rgba[];majorFont?:string;minorFont?:string}
function hexColor(value:string|undefined):Rgba{
 requireProfile(value&&/^[0-9A-F]{6}$/i.test(value),'Invalid RGB color');
 return {kind:'rgba',red:parseInt(value.slice(0,2),16),green:parseInt(value.slice(2,4),16),blue:parseInt(value.slice(4,6),16),alpha:255};
}
function readArgb(value:string|undefined):Rgba{
 // Spreadsheet RGB colors do not use an alpha channel for cell painting.
 // Accept normal FF and zero-prefixed RGB (common OpenPyXL indexed palette).
 requireProfile(value&&/^(?:FF|00)[0-9A-F]{6}$/i.test(value),'Unsupported transparent cell color');
 return hexColor(value.slice(2));
}
function tintColor(color:Rgba,tint:number):Rgba{
 if(!tint)return color;
 const r=color.red/255,g=color.green/255,b=color.blue/255,max=Math.max(r,g,b),min=Math.min(r,g,b),d=max-min,l=(max+min)/2;
 let h=0,s=0;
 if(d){s=d/(1-Math.abs(2*l-1));h=max===r?((g-b)/d+6)%6:max===g?(b-r)/d+2:(r-g)/d+4;h/=6;}
 const next=tint<0?l*(1+tint):l+(1-l)*tint;
 const c=(1-Math.abs(2*next-1))*s,x=c*(1-Math.abs((h*6)%2-1)),m=next-c/2,i=Math.floor(h*6);
 const rgb=([ [c,x,0],[x,c,0],[0,c,x],[0,x,c],[x,0,c],[c,0,x] ][i]??[c,x,0]).map(v=>Math.max(0,Math.min(255,Math.round((v+m)*255))));
 return {kind:'rgba',red:rgb[0]!,green:rgb[1]!,blue:rgb[2]!,alpha:255};
}
const standardIndexed=['000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','800000','008000','000080','808000','800080','008080','C0C0C0','808080','9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF','000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF','00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99','3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696','003366','339966','003300','333300','993300','993366','333399','333333'].map(hexColor);
function readRgb(e:Element,context:ColorContext={}):Color{
 plain(e,['rgb','theme','indexed','auto','tint']);
 const kinds=['rgb','theme','indexed','auto'].filter(k=>get(e,k)!==undefined);requireProfile(kinds.length===1,'Ambiguous color source');
 let color:Rgba;const kind=kinds[0]!;
 if(kind==='rgb')color=readArgb(get(e,kind));
 else if(kind==='theme'){const index=integer(get(e,kind),0,11);requireProfile(context.palette?.[index],'Missing workbook theme color');color=context.palette[index]!;}
 else if(kind==='indexed'){const index=integer(get(e,kind),0,65);const palette=context.indexed??standardIndexed;requireProfile(index<64?palette[index]:true,'Missing indexed color');color=index===64?hexColor('000000'):index===65?hexColor('FFFFFF'):palette[index]!;}
 else{requireProfile(flag(get(e,'auto')),'Unsupported automatic color');color=hexColor('000000');}
 return tintColor(color,get(e,'tint')===undefined?0:decimal(get(e,'tint'),-1,1));
}
function readWorkbookTheme(bytes:Uint8Array,options:ContainerOptions):ColorContext{
 const top=parseXml(bytes,options).documentElement!;
 requireProfile(top.namespaceURI===A&&top.localName==='theme','Unexpected theme root');
 const unique=(e:Element,local:string)=>{const found=childElements(e).filter(c=>c.namespaceURI===A&&c.localName===local);requireProfile(found.length===1,'Missing or duplicate theme '+local);return found[0]!;};
 const elements=unique(top,'themeElements'),scheme=unique(elements,'clrScheme'),fontScheme=unique(elements,'fontScheme');
 const palette=['lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'].map(key=>{
  const item=unique(scheme,key),nodes=childElements(item);requireProfile(nodes.length===1&&nodes[0]!.namespaceURI===A&&!childElements(nodes[0]!).length,'Unsupported theme color transforms');
  const c=nodes[0]!;requireProfile(c.localName==='srgbClr'||c.localName==='sysClr','Unsupported theme color type');
  return hexColor(get(c,c.localName==='sysClr'?'lastClr':'val'));
 });
 const font=(kind:string)=>{const value=get(unique(unique(fontScheme,kind),'latin'),'typeface');requireProfile(value,'Missing theme font');return value;};
 return {palette,majorFont:font('majorFont'),minorFont:font('minorFont')};
}
function styleCheck(s:TextStyle|CellStyle,base=false):void{requireProfile(Object.keys(s).every(k=>(base?FONT_KEYS:STYLE_KEYS).includes(k as never)),'Cell style includes unsupported properties');}
function fontXml(s:TextStyle):string{requireProfile(s.fontFamily&&s.fontSize,'Workbook font needs family and size');return '<font><sz val="'+s.fontSize/12700+'"/><name val="'+esc(s.fontFamily)+'"/>'+(s.bold?'<b/>':'')+(s.italic?'<i/>':'')+(s.underline?'<u/>':'')+(s.color?'<color rgb="'+rgb(s.color)+'"/>':'')+'</font>';}
function readFont(e:Element,context:ColorContext):TextStyle{attrs(e,[]);const c=nodes(e,['sz','name','b','i','u','color','family','scheme','charset']);const name=one(c,'name')!,sz=one(c,'sz')!;plain(name,['val']);plain(sz,['val']);const result:TextStyle={fontFamily:get(name,'val'),fontSize:Math.round(decimal(get(sz,'val'),.01,1638)*12700)};requireProfile(result.fontFamily,'Missing font name');
  for(const [node,key] of [['b','bold'],['i','italic'],['u','underline']] as const){const e=one(c,node,false);if(e){plain(e,['val']);const v=get(e,'val');result[key]=node==='u'&&v==='single'?true:node==='u'&&v==='none'?false:flag(v,true);}else result[key]=false;}
  const family=one(c,'family',false),charset=one(c,'charset',false),scheme=one(c,'scheme',false);
  if(family){plain(family,['val']);integer(get(family,'val'),0,14);}
  if(charset){plain(charset,['val']);integer(get(charset,'val'),0,255);}
  if(scheme){plain(scheme,['val']);const value=get(scheme,'val');requireProfile(['major','minor','none'].includes(value??''),'Unsupported font scheme');
    if(value!=='none'){const resolved=value==='major'?context.majorFont:context.minorFont;requireProfile(resolved,'Missing theme font scheme');result.fontFamily=resolved;}
  }
  const color=one(c,'color',false);if(color)result.color=readRgb(color,context);return result;}
const borderWidth:Readonly<Record<string,number>>={thin:9525,medium:25400,thick:38100};
const borderSides=['left','right','top','bottom'] as const;
const borderLines=new Set<CellBorderLine>(['hair','thin','medium','thick','dotted','dashed','mediumDashed','double','dashDot','mediumDashDot','dashDotDot','mediumDashDotDot','slantDashDot']);
type ImportedBorder=Pick<CellStyle,'border'|'borderEdges'>;
function borderXml(style:ImportedBorder):string{
 let edges=style.borderEdges;
 if(style.border){
  const line=Object.keys(borderWidth).find(k=>borderWidth[k]===style.border!.width);requireProfile(line,'Unsupported uniform border width');
  edges=Object.fromEntries(borderSides.map(side=>[side,{line:line as CellBorderLine,color:style.border!.color}]));
 }
 return '<border>'+borderSides.map(side=>{const e=edges?.[side];return e?'<'+side+' style="'+e.line+'"><color rgb="'+rgb(e.color)+'"/></'+side+'>':'<'+side+'/>';}).join('')+'<diagonal/></border>';
}
function parseBorder(e:Element,context:ColorContext):ImportedBorder|undefined{
 attrs(e,['diagonalUp','diagonalDown','outline']);requireProfile(!flag(get(e,'diagonalUp'))&&!flag(get(e,'diagonalDown'))&&flag(get(e,'outline'),true),'Unsupported diagonal or nonoutline border');
 const children=nodes(e,['left','right','top','bottom','diagonal']);
 const edges:NonNullable<CellStyle['borderEdges']>={};
 for(const side of borderSides){
  const el=one(children,side,false);if(!el)continue;
  attrs(el,['style']);const kind=get(el,'style'),c=nodes(el,['color']);
  if(!kind||kind==='none'){requireProfile(!c.length,'Unexpected border color without line');continue;}
  requireProfile(borderLines.has(kind as CellBorderLine),'Unsupported border line');
  const sourceColor=one(c,'color',false),color=sourceColor?readRgb(sourceColor,context):hexColor('000000');
  edges[side]={line:kind as CellBorderLine,color};
 }
 const diag=one(children,'diagonal',false);if(diag)plain(diag);
 const keys=Object.keys(edges) as CellBorderSide[];if(!keys.length)return undefined;
 const first=edges[keys[0]!]!;
 if(keys.length===4&&borderWidth[first.line]&&keys.every(side=>same(edges[side],first)))return {border:{width:borderWidth[first.line]!,color:first.color}};
 return {borderEdges:edges};
}
function styleWriter(document:WorkbookSnapshot){
  styleCheck(document.defaultTextStyle,true);requireProfile(Object.keys(document.resources).length===0&&Object.keys(document.styles).length===0,'Named styles or resources outside profile');
  const defaultFont={fontFamily:'Arial',fontSize:139700,...document.defaultTextStyle};const fonts=[fontXml(defaultFont)],fills=['<fill><patternFill patternType="none"/></fill>','<fill><patternFill patternType="gray125"/></fill>'],borders=['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const formats=new Map<string,number>(),xfs=['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const ids=new Map<string,number>();
  const dedup=(a:string[],v:string)=>{let i=a.indexOf(v);if(i<0){i=a.length;a.push(v);}return i;};
  const styleId=(cell:WorkbookSnapshot['sheets'][number]['cells'][string])=>{
    requireProfile(!cell.styleId,'Named cell style outside profile');const direct=cell.style??{};styleCheck(direct);
    const format=cell.numberFormat??'General';if(!Object.keys(direct).length&&format==='General')return 0;
    let numId=Number(Object.keys(builtins).find(k=>builtins[Number(k)]===format));
    if(!Number.isFinite(numId)){if(!formats.has(format))formats.set(format,164+formats.size);numId=formats.get(format)!;}
    const font:TextStyle={...defaultFont};for(const key of FONT_KEYS){const value=direct[key];if(value!==undefined)Object.assign(font,{[key]:value});}
    const fi=dedup(fonts,fontXml(font)),fill=direct.fill?dedup(fills,'<fill><patternFill patternType="solid"><fgColor rgb="'+rgb(direct.fill)+'"/><bgColor indexed="64"/></patternFill></fill>'):0;
    const border=direct.border||direct.borderEdges?dedup(borders,borderXml(direct)):0;
    const horizontal=direct.alignment?({start:'left',center:'center',end:'right',justify:'justify'} as const)[direct.alignment]:undefined;
    const vertical=direct.verticalAlignment;
    const alignment=(horizontal||vertical||direct.wrapText!==undefined)?'<alignment'+(horizontal?' horizontal="'+horizontal+'"':'')+(vertical?' vertical="'+vertical+'"':'')+(direct.wrapText!==undefined?' wrapText="'+Number(direct.wrapText)+'"':'')+'/>':'';
    const xf='<xf numFmtId="'+numId+'" fontId="'+fi+'" fillId="'+fill+'" borderId="'+border+'" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"'+(alignment?' applyAlignment="1">'+alignment+'</xf>':'/>');
    let id=ids.get(xf);if(id===undefined){id=xfs.length;xfs.push(xf);ids.set(xf,id);}requireProfile(xfs.length<=64000&&formats.size<=1000,'Style budget exceeded');return id;
  };
  const xml=()=>DECL+'<styleSheet xmlns="'+NS+'">'+(formats.size?'<numFmts count="'+formats.size+'">'+[...formats].map(([f,id])=>'<numFmt numFmtId="'+id+'" formatCode="'+esc(f)+'"/>').join('')+'</numFmts>':'')+
    '<fonts count="'+fonts.length+'">'+fonts.join('')+'</fonts><fills count="'+fills.length+'">'+fills.join('')+'</fills><borders count="'+borders.length+'">'+borders.join('')+'</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="'+xfs.length+'">'+xfs.join('')+'</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  return {styleId,xml};
}
interface ImportedStyle {style?:CellStyle;numberFormat?:string}
function readStyles(bytes:Uint8Array,options:ContainerOptions,theme:ColorContext={}):{font:TextStyle;styles:ImportedStyle[]}{
  const top=root(bytes,'styleSheet',options);attrs(top,[]);const items=nodes(top,['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs','tableStyles','colors']);
  const counted=(name:string,child:string,required=true)=>{const e=one(items,name,required);if(!e)return [];attrs(e,['count']);const list=nodes(e,[child]);requireProfile(integer(get(e,'count'),0,64000)===list.length,'Style count mismatch');return list;};
  const context:ColorContext={...theme};
  requireProfile(counted('dxfs','dxf',false).length===0,'Conditional/differential formatting outside profile');
  const tables=one(items,'tableStyles',false);if(tables){attrs(tables,['count','defaultTableStyle','defaultPivotStyle']);requireProfile(!nodes(tables,[]).length&&integer(get(tables,'count')??'0',0,0)===0,'Custom table style outside profile');}
  const colors=one(items,'colors',false);if(colors){
    attrs(colors,[]);const indexed=one(nodes(colors,['indexedColors']),'indexedColors',false);
    if(indexed){attrs(indexed,[]);const entries=nodes(indexed,['rgbColor']);requireProfile(entries.length<=64,'Indexed palette budget');context.indexed=entries.map(e=>{plain(e,['rgb']);return readArgb(get(e,'rgb'));});}
  }
  const formats=new Map(Object.entries(builtins).map(([i,f])=>[Number(i),f]));
  for(const e of counted('numFmts','numFmt',false)){plain(e,['numFmtId','formatCode']);const id=integer(get(e,'numFmtId'),164,65535);requireProfile(!formats.has(id)&&get(e,'formatCode')!==undefined,'Duplicate or missing custom format');formats.set(id,get(e,'formatCode')!);}
  const fonts=counted('fonts','font').map(e=>readFont(e,context));requireProfile(fonts.length>0,'Missing default font');
  const fills=counted('fills','fill').map(e=>{attrs(e,[]);const p=one(nodes(e,['patternFill']),'patternFill')!;attrs(p,['patternType']);const c=nodes(p,['fgColor','bgColor']),kind=get(p,'patternType')??'none';
    if(kind==='none'||kind==='gray125'){requireProfile(!c.length,'Unsupported fill colors');return undefined;}requireProfile(kind==='solid','Unsupported fill pattern');const bg=one(c,'bgColor',false);if(bg){attrs(bg,['indexed','rgb','theme','tint','auto']);requireProfile(!childElements(bg).length,'Invalid fill background');}return readRgb(one(c,'fgColor')!,context);});
  requireProfile(fills.length>=2&&!fills[0]&&!fills[1],'Missing default/reserved fills');
  const borders=counted('borders','border').map(e=>parseBorder(e,context));requireProfile(borders.length>0&&!borders[0],'Missing default border');
  const styleXfs=counted('cellStyleXfs','xf');requireProfile(styleXfs.length===1,'Named style inheritance outside profile');plain(styleXfs[0],['numFmtId','fontId','fillId','borderId']);for(const k of ['numFmtId','fontId','fillId','borderId'])requireProfile(get(styleXfs[0],k)==='0','Nondefault parent style');
  const named=counted('cellStyles','cellStyle');requireProfile(named.length===1,'Named styles outside profile');plain(named[0],['name','xfId','builtinId','hidden']);requireProfile(!flag(get(named[0],'hidden')),'Hidden Normal style');requireProfile(get(named[0],'name')==='Normal'&&get(named[0],'xfId')==='0'&&get(named[0],'builtinId')==='0','Unsupported named style');
  const styles=counted('cellXfs','xf').map((e,i)=>{attrs(e,['numFmtId','fontId','fillId','borderId','xfId','applyNumberFormat','applyFont','applyFill','applyBorder','applyAlignment','pivotButton','quotePrefix','applyProtection']);
    const numId=integer(get(e,'numFmtId'),0,65535),fi=integer(get(e,'fontId'),0,fonts.length-1),fill=integer(get(e,'fillId'),0,fills.length-1),border=integer(get(e,'borderId'),0,borders.length-1);
    requireProfile((get(e,'xfId')??'0')==='0'&&formats.has(numId),'Unsupported cell style inheritance/format');
    requireProfile(!flag(get(e,'pivotButton'))&&!flag(get(e,'quotePrefix'))&&!flag(get(e,'applyProtection')),'Unsupported pivot, quote or protection flag');
    for(const key of ['applyNumberFormat','applyFont','applyFill','applyBorder','applyAlignment'])if(get(e,key)!==undefined)requireProfile(flag(get(e,key)),'Disabled style property unsupported');
    const style:CellStyle={};if(fi!==0)Object.assign(style,fonts[fi]);if(fills[fill])style.fill=fills[fill];if(borders[border])Object.assign(style,borders[border]);
    const a=one(nodes(e,['alignment']),'alignment',false);if(a){plain(a,['horizontal','vertical','wrapText','textRotation','indent','relativeIndent','shrinkToFit','justifyLastLine','readingOrder']);requireProfile((get(a,'textRotation')??'0')==='0'&&(get(a,'indent')??'0')==='0'&&(get(a,'relativeIndent')??'0')==='0'&&!flag(get(a,'shrinkToFit'))&&!flag(get(a,'justifyLastLine'))&&(get(a,'readingOrder')??'0')==='0','Unsupported active alignment');const h=get(a,'horizontal'),v=get(a,'vertical');if(h&&h!=='general'){requireProfile(['left','center','right','justify'].includes(h),'Unsupported horizontal alignment');style.alignment=({left:'start',right:'end',center:'center',justify:'justify'} as const)[h as 'left'|'right'|'center'|'justify'];}if(v){requireProfile(['top','center','bottom'].includes(v),'Unsupported vertical alignment');style.verticalAlignment=v as CellStyle['verticalAlignment'];}if(get(a,'wrapText')!==undefined)style.wrapText=flag(get(a,'wrapText'));}
    if(i===0)requireProfile(fi===0&&fill===0&&border===0&&numId===0&&!a,'Nondefault base cell style');
    return {...(Object.keys(style).length?{style}:{}),...(numId?{numberFormat:formats.get(numId)!}:{})};});
  requireProfile(styles.length>0,'Missing cell styles');return {font:fonts[0],styles};
}
function scalarXml(v:CalculatedValue):{type:string;body:string}{switch(v.kind){case 'blank':return {type:'',body:''};case 'string':return {type:' t="inlineStr"',body:'<is><t xml:space="preserve">'+encodeText(v.value)+'</t></is>'};case 'number':return {type:'',body:'<v>'+v.value+'</v>'};case 'boolean':return {type:' t="b"',body:'<v>'+Number(v.value)+'</v>'};case 'error':return {type:' t="e"',body:'<v>'+esc(v.value)+'</v>'};}}
export function writeSheetsXlsx(input:WorkbookSnapshot):Uint8Array{
  const document=validateDocument(input);requireProfile(document.kind==='sheets','Expected workbook');requireProfile(document.sheets.length<=128,'Sheet budget exceeded');
  const style=styleWriter(document),calculation=new WorkbookCalculator().calculate(document,{maxSteps:5_000_000});
  const parts:Record<string,Uint8Array>={},encoder=new TextEncoder(),put=(path:string,value:string)=>{parts[path]=encoder.encode(value);};
  const sheets:string[]=[],rels:string[]=[],types:string[]=[];
  for(const [i,sheet] of document.sheets.entries()){
    const id=i+1,path='xl/worksheets/sheet'+id+'.xml';sheets.push('<sheet name="'+esc(sheet.name)+'" sheetId="'+id+'" r:id="rId'+id+'"/>');rels.push('<Relationship Id="rId'+id+'" Type="'+REL+'/worksheet" Target="worksheets/sheet'+id+'.xml"/>');types.push('<Override PartName="/'+path+'" ContentType="'+CT+'worksheet+xml"/>');
    const rows=new Map<number,{col:number;xml:string}[]>();
    for(const [key,cell] of Object.entries(sheet.cells)){
      const [r,c]=key.split(':').map(Number);const sid=style.styleId(cell),styleAttr=sid?' s="'+sid+'"':'';
      let v:{type:string;body:string};
      if(cell.value.kind==='formula'){const ast=parseFormula(cell.value.source);requireProfile(formulaIsSupported(ast),'Unknown formula prevents semantic save');const result=calculation.get(sheet.id,r,c);requireProfile(result.state==='valid'||result.state==='error','Formula not calculated');const cache=result.value;requireProfile(cache,'Missing formula result');const scalar=cache.kind==='string'?{type:' t="str"',body:'<v>'+encodeText(cache.value)+'</v>'}:scalarXml(cache);v={type:scalar.type,body:'<f>'+esc(cell.value.source.slice(1))+'</f>'+scalar.body};}
      else v=scalarXml(cell.value);
      const row=rows.get(r)??[];row.push({col:c,xml:'<c r="'+address(r,c)+'"'+styleAttr+v.type+'>'+v.body+'</c>'});rows.set(r,row);
    }
    for(const r of Object.keys(sheet.rows).map(Number))if(!rows.has(r))rows.set(r,[]);
    const rowXml=[...rows].sort((a,b)=>a[0]-b[0]).map(([r,c])=>{const d=sheet.rows[String(r)];return '<row r="'+r+'"'+(d?.height!==undefined?' ht="'+d.height/12700+'" customHeight="1"':'')+(d?.hidden?' hidden="1"':'')+'>'+c.sort((a,b)=>a.col-b.col).map(x=>x.xml).join('')+'</row>';}).join('');
    const colXml=Object.entries(sheet.columns).sort((a,b)=>Number(a[0])-Number(b[0])).map(([c,d])=>'<col min="'+c+'" max="'+c+'"'+(d.width!==undefined?' width="'+d.width/(9525*7)+'" customWidth="1"':'')+(d.hidden?' hidden="1"':'')+'/>').join('');
    const fr=sheet.view?.frozenRows??0,fc=sheet.view?.frozenColumns??0;
    const view='<sheetViews><sheetView workbookViewId="0"'+(sheet.view?.rightToLeft?' rightToLeft="1"':'')+(sheet.view?.showGridLines===false?' showGridLines="0"':'')+'>'+((fr||fc)?'<pane'+(fc?' xSplit="'+fc+'"':'')+(fr?' ySplit="'+fr+'"':'')+' topLeftCell="'+address(fr+1,fc+1)+'" activePane="'+(fr&&fc?'bottomRight':fr?'bottomLeft':'topRight')+'" state="frozen"/>':'')+'</sheetView></sheetViews>';
    const merge=sheet.merges.length?'<mergeCells count="'+sheet.merges.length+'">'+sheet.merges.map(m=>'<mergeCell ref="'+address(m.row,m.column)+':'+address(m.row+m.rowCount-1,m.column+m.columnCount-1)+'"/>').join('')+'</mergeCells>':'';
    const geometry=sheet.defaultGeometry,format=geometry?'<sheetFormatPr baseColWidth="8" defaultRowHeight="'+geometry.rowHeight/12700+'"'+(geometry.columnWidth!==undefined?' defaultColWidth="'+geometry.columnWidth/(9525*7)+'"':'')+' customHeight="1"/>':'';
    put(path,DECL+'<worksheet xmlns="'+NS+'">'+view+format+(colXml?'<cols>'+colXml+'</cols>':'')+'<sheetData>'+rowXml+'</sheetData>'+merge+'</worksheet>');
  }
  rels.push('<Relationship Id="rIdStyles" Type="'+REL+'/styles" Target="styles.xml"/>');
  put('xl/styles.xml',style.xml());put('xl/workbook.xml',DECL+'<workbook xmlns="'+NS+'" xmlns:r="'+REL+'"><workbookPr date1904="'+Number(document.dateSystem==='1904')+'"/><sheets>'+sheets.join('')+'</sheets><calcPr fullCalcOnLoad="1"/></workbook>');
  put('xl/_rels/workbook.xml.rels',DECL+'<Relationships xmlns="'+PKG+'/relationships">'+rels.join('')+'</Relationships>');
  put('_rels/.rels',DECL+'<Relationships xmlns="'+PKG+'/relationships"><Relationship Id="rId1" Type="'+REL+'/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  put('[Content_Types].xml',DECL+'<Types xmlns="'+PKG+'/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="'+CT+'sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="'+CT+'styles+xml"/>'+types.join('')+'</Types>');
  return writeZip(parts);
}
interface NativeSheet {id:string;path:string;relId:string;numericId:number}
export interface SheetsXlsxSource {document:SheetsDocument;cachedValues:[string,CalculatedValue][];native:{opc:OpcPackage;sheetParts:NativeSheet[];stylesPath?:string;sharedPath?:string;metadataPaths:string[];themePath?:string;conversionBlocked:boolean}}
/** Common Excel extended-properties sheet-name vector. Accept only one
 * worksheet category, exact counts and the actual workbook title sequence. */
function sheetTitleMetadata(element:Element,names?:readonly string[]):{count:Element;vector:Element;titles:Element[]}|undefined{
 const children=childElements(element),heading=one(children,'HeadingPairs',false),titles=one(children,'TitlesOfParts',false);
 if(!heading&&!titles)return undefined;
 requireProfile(heading&&titles,'Incomplete sheet-title metadata');
 const typed=(e:Element,local:string,allowed:readonly string[]=[])=>{
  requireProfile(e.namespaceURI===VT&&e.localName===local,'Unsupported sheet-title vector type');attrs(e,allowed);return e;
 };
 const single=(e:Element,local:string,allowed:readonly string[]=[])=>{
  const c=childElements(e);requireProfile(c.length===1,'Invalid sheet-title metadata cardinality');return typed(c[0]!,local,allowed);
 };
 attrs(heading,[]);attrs(titles,[]);
 const h=single(heading,'vector',['size','baseType']),v=single(titles,'vector',['size','baseType']);
 requireProfile(get(h,'size')==='2'&&get(h,'baseType')==='variant'&&['lpstr','lpwstr'].includes(get(v,'baseType')??''),'Unsupported sheet-title categories');
 const pair=childElements(h);requireProfile(pair.length===2,'Unsupported sheet-title category count');
 const label=single(typed(pair[0]!,'variant'),'lpstr'),count=single(typed(pair[1]!,'variant'),'i4');
 requireProfile(['Worksheets','工作表','Arbeitsblätter','Feuilles de calcul','Hojas de cálculo','ワークシート','워크시트'].includes(leaf(label)),'Unsupported sheet-title category');
 const items=childElements(v).map(e=>typed(e,get(v,'baseType')!));
 const n=integer(leaf(count),1,128);requireProfile(n===items.length&&integer(get(v,'size'),1,128)===n,'Sheet-title count mismatch');
 const values=items.map(e=>leaf(e));requireProfile(values.every(t=>t.length>0&&t.length<=31),'Invalid sheet-title name');
 if(names)requireProfile(same(values,names),'Sheet-title metadata does not match workbook sheets');
 return {count,vector:v,titles:items};
}
function admitMetadata(opc:OpcPackage,rr:OpcRelationship[],options:ContainerOptions):string[]{
 const seen=new Set<string>(),paths:string[]=[];
 for(const rel of rr){
  if(rel.type===REL+'/officeDocument'){requireProfile(rel.targetPart===opc.mainPart&&!seen.has(rel.type),'Invalid root office document relation');seen.add(rel.type);continue;}
  requireProfile([CORE_REL,EXT_REL].includes(rel.type)&&rel.mode==='Internal'&&rel.targetPart&&!rel.fragment&&!seen.has(rel.type),'Unknown root relationship');
  const path=rel.targetPart;seen.add(rel.type);requireProfile((opc.relationships.get(path)??[]).length===0&&!paths.includes(path),'Unexpected metadata dependencies');
  const element=parseXml(opc.parts.get(path)!,options).documentElement!;
  if(rel.type===CORE_REL)requireProfile(opc.contentTypes.get(path)==='application/vnd.openxmlformats-package.core-properties+xml'&&element.namespaceURI===CP&&element.localName==='coreProperties','Invalid core properties');
  else{
   requireProfile(opc.contentTypes.get(path)==='application/vnd.openxmlformats-officedocument.extended-properties+xml'&&element.namespaceURI===EP&&element.localName==='Properties','Invalid extended properties');
   sheetTitleMetadata(element);
   const keys=new Set<string>();
   for(const child of childElements(element)){const key=child.localName??'';requireProfile(child.namespaceURI===EP&&!keys.has(key)&&(['HeadingPairs','TitlesOfParts'].includes(key)||['Application','AppVersion','Company','Manager','DocSecurity','HyperlinksChanged','LinksUpToDate','ScaleCrop','SharedDoc','Template','TotalTime','PresentationFormat'].includes(key)&&childElements(child).length===0),'Sheet-dependent or unsupported extended property');keys.add(key);if(key==='DocSecurity')requireProfile((child.textContent??'').trim()==='0','Protected workbook metadata');if(key==='SharedDoc')requireProfile(['false','0'].includes((child.textContent??'').trim()),'Shared workbook metadata');}
  }
  paths.push(path);
 }
 requireProfile(seen.has(REL+'/officeDocument'),'Missing workbook relationship');return paths;
}
function verifySelection(e:Element):void{
 plain(e,['pane','activeCell','activeCellId','sqref']);const pane=get(e,'pane');requireProfile(!pane||['topLeft','topRight','bottomLeft','bottomRight'].includes(pane),'Unknown selection pane');
 if(get(e,'activeCell'))coordinate(get(e,'activeCell'));
 if(get(e,'activeCellId')!==undefined)integer(get(e,'activeCellId'),0,10000);
 const sqref=get(e,'sqref');if(sqref){requireProfile(sqref.length<=65536,'Selection range budget');for(const value of sqref.split(/\s+/)){if(value.includes(':'))range(value);else coordinate(value);}}
}
function verifyWorksheetAux(elements:Element[]):Sheet['defaultGeometry']|undefined{
 const dim=one(elements,'dimension',false);if(dim){plain(dim,['ref']);const ref=get(dim,'ref');if(ref?.includes(':'))range(ref);else coordinate(ref);}
 const properties=one(elements,'sheetPr',false);if(properties){attrs(properties,[]);for(const n of nodes(properties,['outlinePr','pageSetUpPr'])){if(n.localName==='outlinePr'){plain(n,['summaryBelow','summaryRight','showOutlineSymbols','applyStyles']);requireProfile(flag(get(n,'summaryBelow'),true)&&flag(get(n,'summaryRight'),true)&&flag(get(n,'showOutlineSymbols'),true)&&!flag(get(n,'applyStyles')),'Nonstandard worksheet outline');}else{plain(n,['autoPageBreaks','fitToPage']);requireProfile(flag(get(n,'autoPageBreaks'),true)&&!flag(get(n,'fitToPage')),'Unsupported page setup mode');}}}
 let geometry:Sheet['defaultGeometry'];
 const format=one(elements,'sheetFormatPr',false);if(format){
  plain(format,['baseColWidth','defaultRowHeight','defaultColWidth','customHeight','zeroHeight','thickTop','thickBottom','outlineLevelRow','outlineLevelCol']);
  requireProfile((get(format,'baseColWidth')??'8')==='8'&&!flag(get(format,'zeroHeight'))&&!flag(get(format,'thickTop'))&&!flag(get(format,'thickBottom'))&&(get(format,'outlineLevelRow')??'0')==='0'&&(get(format,'outlineLevelCol')??'0')==='0','Hidden or nonstandard default geometry/outline outside profile');
  if(get(format,'customHeight')!==undefined)flag(get(format,'customHeight'));
  const height=Math.round(decimal(get(format,'defaultRowHeight')??'15',.0001,409)*12700);
  const width=get(format,'defaultColWidth');
  geometry={rowHeight:height,...(width!==undefined?{columnWidth:Math.round(decimal(width,.0001,255)*9525*7)}:{})};
 }
 const margins=one(elements,'pageMargins',false);if(margins){plain(margins,['left','right','top','bottom','header','footer']);for(const key of ['left','right','top','bottom','header','footer'])decimal(get(margins,key),0,100);}
 return geometry;
}

export function openSheetsXlsx(bytes:Uint8Array,documentId:string,title='Workbook',options:ContainerOptions={}):SheetsXlsxSource{
  const opc=openOpc(bytes,options);requireProfile(opc.kind==='xlsx'&&opc.conformance==='transitional'&&!opc.hasMacros&&!opc.hasSignatures&&!opc.externalRelationships,'Protected or unsupported package');
  const main=opc.mainPart;requireProfile(main,'Missing workbook main part');
  const rr=opc.relationships.get('')??[],metadataPaths=admitMetadata(opc,rr,options);
  const wrels=opc.relationships.get(main)??[];requireProfile(wrels.every(r=>[REL+'/worksheet',REL+'/styles',REL+'/sharedStrings',REL+'/theme'].includes(r.type)&&r.mode==='Internal'&&r.targetPart&&!r.fragment),'Unknown workbook dependency');
  const targets=wrels.map(r=>r.targetPart!);requireProfile(new Set(targets).size===targets.length,'Aliased workbook dependency');
  for(const rel of wrels){const type=rel.type.slice(REL.length+1),mime=type==='theme'?'application/vnd.openxmlformats-officedocument.theme+xml':CT+type+'+xml';requireProfile(opc.contentTypes.get(rel.targetPart!)===mime,'Unexpected workbook part content type');}
  const expected=new Set(['[Content_Types].xml','_rels/.rels',main,relationFile(main),...metadataPaths,...targets]);
  requireProfile(opc.parts.size===expected.size&&[...opc.parts.keys()].every(p=>expected.has(p))&&[...opc.relationships.keys()].every(p=>p===''||p===main),'Unknown package parts/dependencies');
  const wb=root(opc.parts.get(main)!,'workbook',options);attrs(wb,[]);const children=nodes(wb,['fileVersion','workbookPr','workbookProtection','bookViews','sheets','definedNames','calcPr']);
  const fileVersion=one(children,'fileVersion',false);if(fileVersion)plain(fileVersion,['appName','lastEdited','lowestEdited','rupBuild']);
  const protection=one(children,'workbookProtection',false);if(protection)plain(protection);
  const defined=one(children,'definedNames',false);if(defined){attrs(defined,[]);nodes(defined,[]);}
  const bookViews=one(children,'bookViews',false);if(bookViews){attrs(bookViews,[]);const view=one(nodes(bookViews,['workbookView']),'workbookView')!;plain(view,['visibility','minimized','showHorizontalScroll','showVerticalScroll','showSheetTabs','xWindow','yWindow','windowWidth','windowHeight','tabRatio','firstSheet','activeTab','autoFilterDateGrouping']);requireProfile((get(view,'visibility')??'visible')==='visible'&&!flag(get(view,'minimized'))&&flag(get(view,'showHorizontalScroll'),true)&&flag(get(view,'showVerticalScroll'),true)&&flag(get(view,'showSheetTabs'),true),'Hidden/minimized workbook view');for(const key of ['firstSheet','activeTab'])if(get(view,key)!==undefined)integer(get(view,key),0,127);}

  const document=createDocument('sheets',documentId,title);const pr=one(children,'workbookPr',false);if(pr){plain(pr,['date1904','defaultThemeVersion','filterPrivacy','codeName']);requireProfile(!get(pr,'codeName'),'Macro workbook codeName');document.dateSystem=flag(get(pr,'date1904'))?'1904':'1900';}
  const calc=one(children,'calcPr',false);if(calc){plain(calc,['fullCalcOnLoad','calcId','calcMode','forceFullCalc','calcOnSave','concurrentCalc','concurrentManualCount']);requireProfile((get(calc,'calcMode')??'auto')==='auto','Manual or partial calculation mode');if(get(calc,'fullCalcOnLoad')!==undefined)flag(get(calc,'fullCalcOnLoad'));}
  const styleRel=wrels.filter(r=>r.type===REL+'/styles');requireProfile(styleRel.length<=1,'Duplicate styles relationship');
  const themeRel=wrels.filter(r=>r.type===REL+'/theme');requireProfile(themeRel.length<=1,'Duplicate theme relationship');const theme=themeRel.length?readWorkbookTheme(opc.parts.get(themeRel[0]!.targetPart!)!,options):{};
  const imported=styleRel.length?readStyles(opc.parts.get(styleRel[0].targetPart!)!,options,theme):{font:document.defaultTextStyle,styles:[{}] as ImportedStyle[]};
  document.defaultTextStyle=imported.font;
  const sst=wrels.filter(r=>r.type===REL+'/sharedStrings');requireProfile(sst.length<=1,'Duplicate shared strings relationship');let shared:string[]=[];
  if(sst.length){const e=root(opc.parts.get(sst[0].targetPart!)!,'sst',options);attrs(e,['count','uniqueCount']);shared=nodes(e,['si']).map(si=>{attrs(si,[]);return decodeText(leaf(one(nodes(si,['t']),'t')!,['xml:space']));});if(get(e,'uniqueCount')!==undefined)requireProfile(integer(get(e,'uniqueCount'),0,1000000)===shared.length,'Shared string count mismatch');}
  const sheets=one(children,'sheets')!;attrs(sheets,[]);const defs=nodes(sheets,['sheet']);requireProfile(defs.length>=1&&defs.length<=128,'Sheet count outside profile');document.sheets=[];
  const used=new Set<string>(),ids=new Set<number>(),cachedValues:[string,CalculatedValue][]=[],sheetParts:NativeSheet[]=[];
  let cellCount=0,worksheetAuxiliary=false;
  const scalar=(type:string|undefined,content:string|undefined):CalculatedValue=>{
    if(type==='str')return {kind:'string',value:decodeText(content??'')};
    if(content===undefined){requireProfile(!type||type==='n','Missing typed cell value');return {kind:'blank'};}
    if(!type||type==='n')return {kind:'number',value:decimal(content)};
    if(type==='b'){requireProfile(content==='0'||content==='1','Invalid boolean cell');return {kind:'boolean',value:content==='1'};}
    if(type==='e'){requireProfile(ERROR.has(content),'Unsupported error value');return {kind:'error',value:content as Extract<CalculatedValue,{kind:'error'}>['value']};}
    if(type==='s')return {kind:'string',value:shared[integer(content,0,shared.length-1)]};
    throw new OfficeError('unsupported','Unsupported XLSX cell type');
  };
  for(const [i,def] of defs.entries()){
    plain(def,['name','sheetId','r:id','state']);requireProfile((get(def,'state')??'visible')==='visible','Hidden worksheet outside profile');const name=get(def,'name');requireProfile(name,'Missing sheet name');
    const sheetNo=integer(get(def,'sheetId'),1,2147483647);requireProfile(!ids.has(sheetNo),'Duplicate sheet ID');ids.add(sheetNo);
    const id=def.getAttributeNS(REL,'id'),rel=wrels.find(r=>r.id===id&&r.type===REL+'/worksheet');requireProfile(rel?.targetPart&&!used.has(rel.targetPart),'Missing or duplicate sheet relationship');used.add(rel.targetPart);
    const sheet:Sheet={id:'sheet-'+(i+1),name,cells:{},rows:{},columns:{},merges:[]};
    sheetParts.push({id:sheet.id,path:rel.targetPart,relId:rel.id,numericId:sheetNo});
    const top=root(opc.parts.get(rel.targetPart)!,'worksheet',options);attrs(top,[]);const elements=nodes(top,['sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','mergeCells','pageMargins']);const defaultGeometry=verifyWorksheetAux(elements);if(defaultGeometry)sheet.defaultGeometry=defaultGeometry;worksheetAuxiliary||=elements.some(e=>['sheetPr','sheetFormatPr','pageMargins'].includes(e.localName??''));
    const views=one(elements,'sheetViews',false);if(views){attrs(views,[]);const view=one(nodes(views,['sheetView']),'sheetView')!;attrs(view,['workbookViewId','rightToLeft','showGridLines','tabSelected','zoomScale','zoomScaleNormal','zoomScaleSheetLayoutView','zoomScalePageLayoutView','topLeftCell','showZeros','showOutlineSymbols','showRowColHeaders','view']);requireProfile((get(view,'view')??'normal')==='normal'&&flag(get(view,'showZeros'),true)&&flag(get(view,'showOutlineSymbols'),true)&&flag(get(view,'showRowColHeaders'),true),'Unsupported worksheet view mode');if(get(view,'topLeftCell'))coordinate(get(view,'topLeftCell'));requireProfile(get(view,'workbookViewId')==='0','Unsupported workbook view');sheet.view={};if(get(view,'rightToLeft')!==undefined)sheet.view.rightToLeft=flag(get(view,'rightToLeft'));if(get(view,'showGridLines')!==undefined)sheet.view.showGridLines=flag(get(view,'showGridLines'));const viewNodes=nodes(view,['pane','selection']);for(const sel of viewNodes.filter(x=>x.localName==='selection'))verifySelection(sel);const pane=one(viewNodes,'pane',false);if(pane){plain(pane,['xSplit','ySplit','topLeftCell','activePane','state']);requireProfile(get(pane,'state')==='frozen','Unsupported split pane');const fc=integer(get(pane,'xSplit')??'0',0,16383),fr=integer(get(pane,'ySplit')??'0',0,1048575);requireProfile(fr||fc,'Empty frozen pane');requireProfile(get(pane,'topLeftCell')===address(fr+1,fc+1),'Noncanonical frozen pane origin');requireProfile(get(pane,'activePane')===(fr&&fc?'bottomRight':fr?'bottomLeft':'topRight'),'Unexpected active pane');sheet.view.frozenRows=fr;sheet.view.frozenColumns=fc;}if(!Object.keys(sheet.view).length)delete sheet.view;}
    const cols=one(elements,'cols',false);if(cols){attrs(cols,[]);for(const e of nodes(cols,['col'])){plain(e,['min','max','width','customWidth','hidden']);const start=integer(get(e,'min'),1,16384),end=integer(get(e,'max'),start,16384);const width=get(e,'width'),hidden=get(e,'hidden');for(let col=start;col<=end;col++){requireProfile(!sheet.columns[col],'Overlapping column metadata');sheet.columns[col]={...(width!==undefined?{width:Math.round(decimal(width,.00001,255)*9525*7)}:{}),...(hidden!==undefined?{hidden:flag(hidden)}:{})};}}}
    const data=one(elements,'sheetData')!;attrs(data,[]);const rowSeen=new Set<number>();for(const row of nodes(data,['row'])){attrs(row,['r','ht','customHeight','hidden','spans']);if(get(row,'spans'))requireProfile(/^\d+:\d+$/.test(get(row,'spans')!),'Invalid row spans');const r=integer(get(row,'r'),1,1048576);requireProfile(!rowSeen.has(r),'Duplicate row');rowSeen.add(r);const ht=get(row,'ht'),hidden=get(row,'hidden');if(ht!==undefined||hidden!==undefined)sheet.rows[r]={...(ht!==undefined?{height:Math.round(decimal(ht,.00001,409)*12700)}:{}),...(hidden!==undefined?{hidden:flag(hidden)}:{})};
      for(const el of nodes(row,['c'])){requireProfile(++cellCount<=100000,'Cell budget exceeded');attrs(el,['r','s','t']);const [rr,c]=coordinate(get(el,'r'));requireProfile(rr===r,'Cell row mismatch');const key=r+':'+c;requireProfile(!sheet.cells[key],'Duplicate cell');const st=imported.styles[integer(get(el,'s')??'0',0,imported.styles.length-1)];const type=get(el,'t'),ch=nodes(el,['f','v','is']),f=one(ch,'f',false),v=one(ch,'v',false),is=one(ch,'is',false);let value:Cell['value'];
        if(f){requireProfile(!is&&type!=='s'&&type!=='inlineStr','Invalid formula cell');const source='='+leaf(f);requireProfile(formulaIsSupported(parseFormula(source)),'Unknown formula outside editable profile');value={kind:'formula',source};if(v){const content=leaf(v);if(content!==''||type==='str')cachedValues.push([calculationKey(sheet.id,r,c),scalar(type,content)]);}}
        else if(type==='inlineStr'){requireProfile(is&&!v,'Invalid inline string');attrs(is,[]);value={kind:'string',value:decodeText(leaf(one(nodes(is,['t']),'t')!,['xml:space']))};}
        else{requireProfile(!is,'Unexpected inline string');value=scalar(type,v?leaf(v):undefined);}
        sheet.cells[key]={value,...st};
      }}
    const merges=one(elements,'mergeCells',false);if(merges){attrs(merges,['count']);sheet.merges=nodes(merges,['mergeCell']).map(e=>{plain(e,['ref']);return range(get(e,'ref'));});if(get(merges,'count')!==undefined)requireProfile(integer(get(merges,'count'),0,10000)===sheet.merges.length,'Merge count mismatch');}
    // Covered non-anchor values may not be discarded in a semantic editor.
    for(const [key,cell] of Object.entries(sheet.cells)){if(cell.value.kind==='blank')continue;const [r,c]=key.split(':').map(Number);requireProfile(!sheet.merges.some(m=>r>=m.row&&r<m.row+m.rowCount&&c>=m.column&&c<m.column+m.columnCount&&(r!==m.row||c!==m.column)),'Merged cell hides source content');}
    document.sheets.push(sheet);
  }
  requireProfile(used.size===wrels.filter(r=>r.type===REL+'/worksheet').length,'Unreferenced worksheet dependency');
  for(const path of metadataPaths)if(opc.contentTypes.get(path)==='application/vnd.openxmlformats-officedocument.extended-properties+xml')sheetTitleMetadata(parseXml(opc.parts.get(path)!,options).documentElement!,document.sheets.map(s=>s.name));
  return {document:validateDocument(document) as SheetsDocument,cachedValues,native:{opc,sheetParts,stylesPath:styleRel[0]?.targetPart??undefined,sharedPath:sst[0]?.targetPart??undefined,metadataPaths,themePath:themeRel[0]?.targetPart??undefined,conversionBlocked:!!(metadataPaths.length||themeRel.length||worksheetAuxiliary||one(children,'bookViews',false)||one(children,'fileVersion',false))}};
}
/** Append only newly required styles; original style ids never move. */
function mergeSourceStyles(source:SheetsXlsxSource,generated:SheetsXlsxSource):{mapping:number[];bytes?:Uint8Array}{
 const native=source.native,canonical=generated.native,generatedPath=canonical.stylesPath!;
 const fresh=canonical.opc.parts.get(generatedPath)!;
 const generatedStyles=readStyles(fresh,{});
 if(!native.stylesPath)return {mapping:generatedStyles.styles.map((_,i)=>i),bytes:fresh};
 const before=native.opc.parts.get(native.stylesPath)!;
 const theme=native.themePath?readWorkbookTheme(native.opc.parts.get(native.themePath)!,{}):{};
 const oldStyles=readStyles(before,{},theme);
 requireProfile(same(oldStyles.font,generatedStyles.font),'Changing the source workbook default font is outside this profile');
 const mapping=generatedStyles.styles.map(style=>oldStyles.styles.findIndex(old=>same(old,style)));
 if(mapping.every(i=>i>=0))return {mapping};
 const oldDoc=parseXml(before),newDoc=parseXml(fresh),old=oldDoc.documentElement!,next=newDoc.documentElement!;
 const element=(parent:Element,name:string)=>childElements(parent).find(e=>e.namespaceURI===NS&&e.localName===name);
 const oldXfs=element(old,'cellXfs')!,newXfs=childElements(element(next,'cellXfs')!);
 const oldNum=element(old,'numFmts');let numFmts=oldNum;
 const newNum=element(next,'numFmts'),numMap=new Map<number,number>(),existingNumbers=new Map<string,number>(),usedNumbers=new Set<number>();
 if(oldNum)for(const item of childElements(oldNum)){const id=integer(get(item,'numFmtId'),164,65535);existingNumbers.set(get(item,'formatCode')!,id);usedNumbers.add(id);}
 const formatId=(id:number):number=>{
  if(id<164)return id;
  if(numMap.has(id))return numMap.get(id)!;
  const freshFmt=childElements(newNum!).find(item=>Number(get(item,'numFmtId'))===id);requireProfile(freshFmt,'Missing generated custom format');
  const code=get(freshFmt,'formatCode')!;let target=existingNumbers.get(code);
  if(target===undefined){
   target=164;while(usedNumbers.has(target))target++;requireProfile(target<=65535,'Custom number format ids exhausted');
   usedNumbers.add(target);existingNumbers.set(code,target);
   if(!numFmts){numFmts=oldDoc.createElementNS(NS,'numFmts');old.insertBefore(numFmts,element(old,'fonts')!);}
   const node=oldDoc.importNode(freshFmt,true) as Element;node.setAttribute('numFmtId',String(target));numFmts.appendChild(node);numFmts.setAttribute('count',String(childElements(numFmts).length));
  }
  numMap.set(id,target);return target;
 };
 const components=new Map<string,Map<number,number>>();
 const component=(name:string,id:number):number=>{
  let cache=components.get(name);if(!cache){cache=new Map([[0,0]]);if(name==='fills')cache.set(1,1);components.set(name,cache);}
  if(cache.has(id))return cache.get(id)!;
  const original=element(old,name)!,node=childElements(element(next,name)!)[id];requireProfile(node,'Missing generated style component');
  const target=childElements(original).length;requireProfile(target<64000,'Style component budget exceeded');
  original.appendChild(oldDoc.importNode(node,true));original.setAttribute('count',String(target+1));cache.set(id,target);return target;
 };
 for(const [index,id]of mapping.entries()){
  if(id>=0)continue;
  const xf=oldDoc.importNode(newXfs[index]!,true) as Element;
  for(const [attribute,table]of [['fontId','fonts'],['fillId','fills'],['borderId','borders']] as const)xf.setAttribute(attribute,String(component(table,Number(get(xf,attribute)))));
  xf.setAttribute('numFmtId',String(formatId(Number(get(xf,'numFmtId')))));
  const target=childElements(oldXfs).length;requireProfile(target<64000,'Cell style budget exceeded');
  oldXfs.appendChild(xf);oldXfs.setAttribute('count',String(target+1));mapping[index]=target;
 }
 return {mapping,bytes:xmlBytes(oldDoc)};
}
function worksheetCells(top:Element):Map<string,Element>{
 const data=one(childElements(top),'sheetData')!,out=new Map<string,Element>();
 for(const row of childElements(data))for(const cell of childElements(row)){const [r,c]=coordinate(get(cell,'r'));out.set(r+':'+c,cell);}
 return out;
}
function effectiveCellStyle(cell:Cell|undefined):unknown{return {style:cell?.style??{},numberFormat:cell?.numberFormat??'General'};}
function dimensionRef(sheet:Sheet):string{
 let r1=1048576,c1=16384,r2=1,c2=1,any=false;
 const take=(r:number,c:number)=>{any=true;r1=Math.min(r1,r);c1=Math.min(c1,c);r2=Math.max(r2,r);c2=Math.max(c2,c);};
 for(const key of Object.keys(sheet.cells)){const [r,c]=key.split(':').map(Number);take(r!,c!);}
 for(const m of sheet.merges){take(m.row,m.column);take(m.row+m.rowCount-1,m.column+m.columnCount-1);}
 if(!any)return 'A1';
 const a=address(r1,c1),b=address(r2,c2);return a===b?a:a+':'+b;
}
function sheetOutput(source:SheetsXlsxSource,prior:NativeSheet|undefined,generated:SheetsXlsxSource,index:number,styles:number[]):Uint8Array{
 const fresh=generated.document.sheets[index]!,freshPath=generated.native.sheetParts[index]!.path;
 const nextDoc=parseXml(generated.native.opc.parts.get(freshPath)!),next=nextDoc.documentElement!;
 const old=prior?source.document.sheets.find(s=>s.id===prior.id):undefined;
 const oldDoc=prior?parseXml(source.native.opc.parts.get(prior.path)!):undefined,top=oldDoc?.documentElement??next;
 const doc=oldDoc??nextDoc,oldCells=oldDoc?worksheetCells(top):new Map<string,Element>();
 const sourceCache=new Map(source.cachedValues),freshCache=new Map(generated.cachedValues);
 const oldId=old?.id??'',freshId=fresh.id;
 const unchangedFormula=(key:string)=>same(sourceCache.get(oldId+'!'+key),freshCache.get(freshId+'!'+key));
 const equalSheet=old&&same({...old,id:fresh.id},fresh);
 const equalCaches=equalSheet&&Object.entries(fresh.cells).every(([k,c])=>c.value.kind!=='formula'||unchangedFormula(k));
 if(prior&&equalSheet&&equalCaches)return source.native.opc.parts.get(prior.path)!;
 const nextCells=worksheetCells(next),nextData=one(childElements(next),'sheetData')!;
 for(const [key,cell]of nextCells){
  const current=fresh.cells[key]!,previous=old?.cells[key],sourceCell=oldCells.get(key);
  // Retain unchanged cell XML including source shared-string index and theme
  // style id. Formula caches are updated whenever derived values differ.
  if(oldDoc&&sourceCell&&same(previous,current)&&(current.value.kind!=='formula'||unchangedFormula(key))){
   cell.parentNode!.replaceChild(nextDoc.importNode(sourceCell,true),cell);continue;
  }
  let mapped=styles[integer(get(cell,'s')??'0',0,styles.length-1)]!;
  if(sourceCell&&same(effectiveCellStyle(previous),effectiveCellStyle(current)))mapped=Number(get(sourceCell,'s')??'0');
  if(mapped)cell.setAttribute('s',String(mapped));else cell.removeAttribute('s');
 }
 if(!oldDoc)return xmlBytes(nextDoc);
 const order=['sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','mergeCells','pageMargins'];
 const replace=(name:string,value:Element|undefined)=>{
  const existing=one(childElements(top),name,false);
  if(!value){if(existing)top.removeChild(existing);return;}
  const copy=doc.importNode(value,true);
  if(existing)top.replaceChild(copy,existing);
  else{const later=childElements(top).find(e=>order.indexOf(e.localName!)>order.indexOf(name));top.insertBefore(copy,later??null);}
 };
 // Row metadata is entirely represented by the model except the source
 // optimization hint spans, which is intentionally regenerated/omitted.
 replace('sheetData',nextData);
 if(!same(old?.defaultGeometry,fresh.defaultGeometry))replace('sheetFormatPr',one(childElements(next),'sheetFormatPr',false));
 if(!same(old?.columns,fresh.columns))replace('cols',one(childElements(next),'cols',false));
 if(!same(old?.merges,fresh.merges))replace('mergeCells',one(childElements(next),'mergeCells',false));
 if(!same(old?.view,fresh.view)){
  const view=one(childElements(top),'sheetViews',false),newView=one(childElements(next),'sheetViews',false);
  if(view&&newView){
   const previous=one(childElements(view),'sheetView')!,replacement=one(childElements(newView),'sheetView')!;
   for(const key of ['zoomScale','zoomScaleNormal','zoomScaleSheetLayoutView','zoomScalePageLayoutView','showZeros','showOutlineSymbols','showRowColHeaders','view','tabSelected'])if(previous.hasAttribute(key))replacement.setAttribute(key,previous.getAttribute(key)!);
  }
  replace('sheetViews',newView);
 }
 const dimension=one(childElements(top),'dimension',false);if(dimension)dimension.setAttribute('ref',dimensionRef(fresh));
 return xmlBytes(doc);
}
/** Preserve original native workbook graph, inactive metadata and all
 * source-owned styles. New cells/styles/sheets are mapped to the original
 * identifiers; only dependencies whose model data changed are regenerated. */
export function saveSheetsXlsx(bytes:Uint8Array,input:WorkbookSnapshot):Uint8Array{
 const model=validateDocument(input);requireProfile(model.kind==='sheets','Expected workbook');
 const original=openSheetsXlsx(bytes,model.documentId,model.title),native=original.native,opc=native.opc;
 if(same(original.document,model))return bytes.slice();
 const generated=openSheetsXlsx(writeSheetsXlsx(model),model.documentId,model.title),stylePlan=mergeSourceStyles(original,generated);
 const parts=new Map(opc.parts),main=opc.mainPart!,mainDoc=parseXml(parts.get(main)!),top=mainDoc.documentElement!;
 const sheetList=one(childElements(top),'sheets')!,oldDefinitions=childElements(sheetList),byId=new Map(native.sheetParts.map((p,i)=>[p.id,{...p,element:oldDefinitions[i]!}]));
 const oldRels=opc.relationships.get(main)??[],relDoc=parseXml(parts.get(relationFile(main))!),relTop=relDoc.documentElement!;
 const usedRelIds=new Set(oldRels.map(r=>r.id)),usedSheetIds=new Set(native.sheetParts.map(s=>s.numericId)),usedPaths=new Set([...parts.keys()].map(x=>x.toLowerCase()));
 const allocate=(base:string,set:Set<string>)=>{let index=1,value=base+index;while(set.has(value))value=base+(++index);set.add(value);return value;};
 const nextPath=(prefix:string,suffix:string)=>{let i=1,p=prefix+i+suffix;while(usedPaths.has(p.toLowerCase()))p=prefix+(++i)+suffix;usedPaths.add(p.toLowerCase());return p;};
 const typesDoc=parseXml(parts.get('[Content_Types].xml')!),types=typesDoc.documentElement!;
 const typeNode=(path:string)=>childElements(types).find(e=>e.localName==='Override'&&get(e,'PartName')?.toLowerCase()===('/'+path).toLowerCase());
 const addType=(path:string,contentType:string)=>{const e=typesDoc.createElementNS(PKG+'/content-types','Override');e.setAttribute('PartName','/'+path);e.setAttribute('ContentType',contentType);types.appendChild(e);};
 const removeType=(path:string)=>{const node=typeNode(path);if(node)types.removeChild(node);};
 const addRel=(id:string,type:string,path:string)=>{const e=relDoc.createElementNS(PKG+'/relationships','Relationship');e.setAttribute('Id',id);e.setAttribute('Type',REL+'/'+type);e.setAttribute('Target','/'+path);relTop.appendChild(e);};
 let graphChanged=false,typesChanged=false,mainChanged=false;
 const currentIds=new Set(model.sheets.map(s=>s.id));
 for(const prior of native.sheetParts)if(!currentIds.has(prior.id)){
  parts.delete(prior.path);removeType(prior.path);
  const rel=childElements(relTop).find(e=>get(e,'Id')===prior.relId)!;relTop.removeChild(rel);graphChanged=typesChanged=mainChanged=true;
 }
 const newPaths:string[]=[],newDefs:Element[]=[];
 for(const [i,sheet]of model.sheets.entries()){
  const prior=byId.get(sheet.id);let path:string,def:Element;
  if(prior){
   path=prior.path;def=prior.element;
   if(get(def,'name')!==sheet.name){def.setAttribute('name',sheet.name);mainChanged=true;}
  }else{
   path=nextPath('xl/worksheets/office-ex-sheet-','.xml');
   const relId=allocate('rIdOexSheet',usedRelIds);let numericId=1;while(usedSheetIds.has(numericId))numericId++;usedSheetIds.add(numericId);
   def=mainDoc.createElementNS(NS,'sheet');def.setAttribute('name',sheet.name);def.setAttribute('sheetId',String(numericId));def.setAttributeNS(REL,'r:id',relId);
   addRel(relId,'worksheet',path);addType(path,CT+'worksheet+xml');graphChanged=typesChanged=mainChanged=true;
  }
  newPaths.push(path);newDefs.push(def);parts.set(path,sheetOutput(original,prior,generated,i,stylePlan.mapping));
 }
 if(!same(native.sheetParts.map(x=>x.id),model.sheets.map(x=>x.id))){
  while(sheetList.firstChild)sheetList.removeChild(sheetList.firstChild);
  for(const def of newDefs)sheetList.appendChild(def);mainChanged=true;
  const bookViews=one(childElements(top),'bookViews',false),view=bookViews&&one(childElements(bookViews),'workbookView',false);
  if(view)for(const key of ['activeTab','firstSheet']){
   const oldIndex=Number(get(view,key)??'0'),id=native.sheetParts[oldIndex]?.id,index=model.sheets.findIndex(s=>s.id===id);view.setAttribute(key,String(Math.max(0,index)));
  }
 }
 if(original.document.dateSystem!==model.dateSystem){
  let pr=one(childElements(top),'workbookPr',false);if(!pr){pr=mainDoc.createElementNS(NS,'workbookPr');top.insertBefore(pr,one(childElements(top),'workbookProtection',false)??one(childElements(top),'bookViews',false)??sheetList);}
  pr.setAttribute('date1904',model.dateSystem==='1904'?'1':'0');mainChanged=true;
 }
 if(stylePlan.bytes){
  let path=native.stylesPath;
  if(!path){path=nextPath('xl/office-ex-styles-','.xml');addRel(allocate('rIdOexStyles',usedRelIds),'styles',path);addType(path,CT+'styles+xml');graphChanged=typesChanged=true;}
  parts.set(path,stylePlan.bytes);
 }
 // Preserve the existing shared string table and indices. Count is optional
 // and derived; repair it only if the retained t=s references change.
 if(native.sharedPath){
  let references=0;for(const path of newPaths)for(const cell of worksheetCells(parseXml(parts.get(path)!).documentElement!).values())if(get(cell,'t')==='s')references++;
  const sharedDoc=parseXml(parts.get(native.sharedPath)!),shared=sharedDoc.documentElement!;
  if(get(shared,'count')!==undefined&&Number(get(shared,'count'))!==references){shared.removeAttribute('count');parts.set(native.sharedPath,xmlBytes(sharedDoc));}
 }
 // Cell-only edits preserve app.xml byte-for-byte. Structural edits update
 // only the already admitted sheet-title count and title vector.
 const oldNames=original.document.sheets.map(s=>s.name),newNames=model.sheets.map(s=>s.name);
 if(!same(oldNames,newNames))for(const path of native.metadataPaths){
  if(opc.contentTypes.get(path)!=='application/vnd.openxmlformats-officedocument.extended-properties+xml')continue;
  const doc=parseXml(parts.get(path)!),metadata=sheetTitleMetadata(doc.documentElement!,oldNames);if(!metadata)continue;
  metadata.count.textContent=String(newNames.length);metadata.vector.setAttribute('size',String(newNames.length));
  const exemplar=metadata.titles[0]!;
  for(const item of metadata.titles)metadata.vector.removeChild(item);
  for(const name of newNames){const item=exemplar.cloneNode(false) as Element;item.textContent=name;metadata.vector.appendChild(item);}
  parts.set(path,xmlBytes(doc));
 }
 if(mainChanged)parts.set(main,xmlBytes(mainDoc));if(graphChanged)parts.set(relationFile(main),xmlBytes(relDoc));if(typesChanged)parts.set('[Content_Types].xml',xmlBytes(typesDoc));
 const output=writeZip(parts),checked=openSheetsXlsx(output,model.documentId,model.title);
 requireProfile(same(checked.document,generated.document),'Saved workbook semantic verification failed');
 requireProfile(same(Object.fromEntries(checked.cachedValues),Object.fromEntries(generated.cachedValues)),'Saved formula cache verification failed');
 return output;
}
