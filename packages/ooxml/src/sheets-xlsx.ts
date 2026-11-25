// SPDX-License-Identifier: MIT
/** Closed Transitional XLSX profile. Unknown dependent parts/markup are not editable.
 * Source: ECMA-376 Part 1 SpreadsheetML, Part 2 OPC. No external Office engine.
 * Calculated caches are derived from the shared Office EX workbook model.
 */
import type { Element } from '@xmldom/xmldom';
import { assert, assertCellAddress, OfficeError } from '@office-ex/contracts';
import { createDocument, validateDocument, type SheetsDocument, type Sheet, type Cell, type CellStyle, type Color, type TextStyle } from '@office-ex/core';
import { columnName, columnNumber, parseFormula, formulaIsSupported, WorkbookCalculator, calculationKey, type CalculatedValue, type WorkbookSnapshot } from '@office-ex/sheets';
import { openOpc } from './opc.js';
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
const builtins:Readonly<Record<number,string>>={0:'General',1:'0',2:'0.00',3:'#,##0',4:'#,##0.00',9:'0%',10:'0.00%',11:'0.00E+00',14:'mm-dd-yy',20:'h:mm',21:'h:mm:ss',22:'m/d/yy h:mm',49:'@'};
const FONT_KEYS=['fontFamily','fontSize','bold','italic','underline','color'] as const;
const STYLE_KEYS=[...FONT_KEYS,'fill','border','alignment','verticalAlignment','wrapText'];
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
function readRgb(e:Element):Color{plain(e,['rgb']);const s=get(e,'rgb');requireProfile(s!==undefined&&/^FF[0-9A-F]{6}$/i.test(s),'Only opaque RGB color supported');return {kind:'rgba',red:parseInt(s.slice(2,4),16),green:parseInt(s.slice(4,6),16),blue:parseInt(s.slice(6,8),16),alpha:255};}
function styleCheck(s:TextStyle|CellStyle,base=false):void{requireProfile(Object.keys(s).every(k=>(base?FONT_KEYS:STYLE_KEYS).includes(k as never)),'Cell style includes unsupported properties');}
function fontXml(s:TextStyle):string{requireProfile(s.fontFamily&&s.fontSize,'Workbook font needs family and size');return '<font><sz val="'+s.fontSize/12700+'"/><name val="'+esc(s.fontFamily)+'"/>'+(s.bold?'<b/>':'')+(s.italic?'<i/>':'')+(s.underline?'<u/>':'')+(s.color?'<color rgb="'+rgb(s.color)+'"/>':'')+'</font>';}
function readFont(e:Element):TextStyle{attrs(e,[]);const c=nodes(e,['sz','name','b','i','u','color']);const name=one(c,'name')!,sz=one(c,'sz')!;plain(name,['val']);plain(sz,['val']);const result:TextStyle={fontFamily:get(name,'val'),fontSize:Math.round(decimal(get(sz,'val'),.01,1638)*12700)};requireProfile(result.fontFamily,'Missing font name');
  for(const [node,key] of [['b','bold'],['i','italic'],['u','underline']] as const){const e=one(c,node,false);if(e){plain(e,['val']);result[key]=flag(get(e,'val'),true);}else result[key]=false;}
  const color=one(c,'color',false);if(color)result.color=readRgb(color);return result;}
const borderWidth:Readonly<Record<string,number>>={thin:9525,medium:25400,thick:38100};
function borderXml(b:NonNullable<CellStyle['border']>):string{const kind=Object.keys(borderWidth).find(k=>borderWidth[k]===b.width);requireProfile(kind,'Unsupported border width');return '<border>'+['left','right','top','bottom'].map(x=>'<'+x+' style="'+kind+'"><color rgb="'+rgb(b.color)+'"/></'+x+'>').join('')+'<diagonal/></border>';}
function parseBorder(e:Element):CellStyle['border']{attrs(e,[]);const children=nodes(e,['left','right','top','bottom','diagonal']);let border:CellStyle['border'];let blank=0;
  for(const side of ['left','right','top','bottom']){const el=one(children,side)!;attrs(el,['style']);const kind=get(el,'style'),c=nodes(el,['color']);if(!kind){requireProfile(!c.length,'Unexpected border color');blank++;continue;}requireProfile(borderWidth[kind]!==undefined,'Unsupported border line');const color=readRgb(one(c,'color')!);const edge={color,width:borderWidth[kind]};requireProfile(!border||JSON.stringify(border)===JSON.stringify(edge),'Asymmetric borders outside profile');border=edge;}
  const diag=one(children,'diagonal',false);if(diag)plain(diag);requireProfile(!border||blank===0,'Partial borders outside profile');return border;}
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
    const border=direct.border?dedup(borders,borderXml(direct.border)):0;
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
function readStyles(bytes:Uint8Array,options:ContainerOptions):{font:TextStyle;styles:ImportedStyle[]}{
  const top=root(bytes,'styleSheet',options);attrs(top,[]);const items=nodes(top,['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles']);
  const counted=(name:string,child:string,required=true)=>{const e=one(items,name,required);if(!e)return [];attrs(e,['count']);const list=nodes(e,[child]);requireProfile(integer(get(e,'count'),0,64000)===list.length,'Style count mismatch');return list;};
  const formats=new Map(Object.entries(builtins).map(([i,f])=>[Number(i),f]));
  for(const e of counted('numFmts','numFmt',false)){plain(e,['numFmtId','formatCode']);const id=integer(get(e,'numFmtId'),164,65535);requireProfile(!formats.has(id)&&get(e,'formatCode')!==undefined,'Duplicate or missing custom format');formats.set(id,get(e,'formatCode')!);}
  const fonts=counted('fonts','font').map(readFont);requireProfile(fonts.length>0,'Missing default font');
  const fills=counted('fills','fill').map(e=>{attrs(e,[]);const p=one(nodes(e,['patternFill']),'patternFill')!;attrs(p,['patternType']);const c=nodes(p,['fgColor','bgColor']),kind=get(p,'patternType');
    if(kind==='none'||kind==='gray125'){requireProfile(!c.length,'Unsupported fill colors');return undefined;}requireProfile(kind==='solid','Unsupported fill pattern');const bg=one(c,'bgColor',false);if(bg){plain(bg,['indexed']);requireProfile(get(bg,'indexed')==='64','Unsupported fill background');}return readRgb(one(c,'fgColor')!);});
  const borders=counted('borders','border').map(parseBorder);
  const styleXfs=counted('cellStyleXfs','xf');requireProfile(styleXfs.length===1,'Named style inheritance outside profile');plain(styleXfs[0],['numFmtId','fontId','fillId','borderId']);for(const k of ['numFmtId','fontId','fillId','borderId'])requireProfile(get(styleXfs[0],k)==='0','Nondefault parent style');
  const named=counted('cellStyles','cellStyle');requireProfile(named.length===1,'Named styles outside profile');plain(named[0],['name','xfId','builtinId']);requireProfile(get(named[0],'name')==='Normal'&&get(named[0],'xfId')==='0'&&get(named[0],'builtinId')==='0','Unsupported named style');
  const styles=counted('cellXfs','xf').map((e,i)=>{attrs(e,['numFmtId','fontId','fillId','borderId','xfId','applyNumberFormat','applyFont','applyFill','applyBorder','applyAlignment']);
    const numId=integer(get(e,'numFmtId'),0,65535),fi=integer(get(e,'fontId'),0,fonts.length-1),fill=integer(get(e,'fillId'),0,fills.length-1),border=integer(get(e,'borderId'),0,borders.length-1);
    requireProfile((get(e,'xfId')??'0')==='0'&&formats.has(numId),'Unsupported cell style inheritance/format');
    for(const key of ['applyNumberFormat','applyFont','applyFill','applyBorder','applyAlignment'])if(get(e,key)!==undefined)requireProfile(flag(get(e,key)),'Disabled style property unsupported');
    const style:CellStyle={};if(fi!==0)Object.assign(style,fonts[fi]);if(fills[fill])style.fill=fills[fill];if(borders[border])style.border=borders[border];
    const a=one(nodes(e,['alignment']),'alignment',false);if(a){plain(a,['horizontal','vertical','wrapText']);const h=get(a,'horizontal'),v=get(a,'vertical');if(h){requireProfile(['left','center','right','justify'].includes(h),'Unsupported horizontal alignment');style.alignment=({left:'start',right:'end',center:'center',justify:'justify'} as const)[h as 'left'|'right'|'center'|'justify'];}if(v){requireProfile(['top','center','bottom'].includes(v),'Unsupported vertical alignment');style.verticalAlignment=v as CellStyle['verticalAlignment'];}if(get(a,'wrapText')!==undefined)style.wrapText=flag(get(a,'wrapText'));}
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
    put(path,DECL+'<worksheet xmlns="'+NS+'">'+view+(colXml?'<cols>'+colXml+'</cols>':'')+'<sheetData>'+rowXml+'</sheetData>'+merge+'</worksheet>');
  }
  rels.push('<Relationship Id="rIdStyles" Type="'+REL+'/styles" Target="styles.xml"/>');
  put('xl/styles.xml',style.xml());put('xl/workbook.xml',DECL+'<workbook xmlns="'+NS+'" xmlns:r="'+REL+'"><workbookPr date1904="'+Number(document.dateSystem==='1904')+'"/><sheets>'+sheets.join('')+'</sheets><calcPr fullCalcOnLoad="1"/></workbook>');
  put('xl/_rels/workbook.xml.rels',DECL+'<Relationships xmlns="'+PKG+'/relationships">'+rels.join('')+'</Relationships>');
  put('_rels/.rels',DECL+'<Relationships xmlns="'+PKG+'/relationships"><Relationship Id="rId1" Type="'+REL+'/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  put('[Content_Types].xml',DECL+'<Types xmlns="'+PKG+'/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="'+CT+'sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="'+CT+'styles+xml"/>'+types.join('')+'</Types>');
  return writeZip(parts);
}
export interface SheetsXlsxSource {document:SheetsDocument;cachedValues:[string,CalculatedValue][]}
export function openSheetsXlsx(bytes:Uint8Array,documentId:string,title='Workbook',options:ContainerOptions={}):SheetsXlsxSource{
  const opc=openOpc(bytes,options);requireProfile(opc.kind==='xlsx'&&opc.conformance==='transitional'&&!opc.hasMacros&&!opc.hasSignatures&&!opc.externalRelationships,'Protected or unsupported package');
  requireProfile(opc.mainPart==='xl/workbook.xml','Workbook location outside profile');
  const rr=opc.relationships.get('')??[];requireProfile(rr.length===1&&rr[0].targetPart==='xl/workbook.xml','Unknown root relationship');
  const wrels=opc.relationships.get('xl/workbook.xml')??[];requireProfile(wrels.every(r=>[REL+'/worksheet',REL+'/styles',REL+'/sharedStrings'].includes(r.type)&&r.targetPart&&!r.fragment),'Unknown workbook dependency');
  const expected=new Set(['[Content_Types].xml','_rels/.rels','xl/workbook.xml','xl/_rels/workbook.xml.rels',...wrels.map(r=>r.targetPart!)]);
  requireProfile(opc.parts.size===expected.size&&[...opc.parts.keys()].every(p=>expected.has(p))&&[...opc.relationships.keys()].every(p=>p===''||p==='xl/workbook.xml'),'Unknown package parts/dependencies');
  const wb=root(opc.parts.get('xl/workbook.xml')!,'workbook',options);attrs(wb,[]);const children=nodes(wb,['workbookPr','sheets','calcPr']);
  const document=createDocument('sheets',documentId,title);const pr=one(children,'workbookPr',false);if(pr){plain(pr,['date1904']);document.dateSystem=flag(get(pr,'date1904'))?'1904':'1900';}
  const calc=one(children,'calcPr',false);if(calc){plain(calc,['fullCalcOnLoad','calcId']);if(get(calc,'fullCalcOnLoad')!==undefined)flag(get(calc,'fullCalcOnLoad'));}
  const styleRel=wrels.filter(r=>r.type===REL+'/styles');requireProfile(styleRel.length<=1,'Duplicate styles relationship');
  const imported=styleRel.length?readStyles(opc.parts.get(styleRel[0].targetPart!)!,options):{font:document.defaultTextStyle,styles:[{}] as ImportedStyle[]};
  document.defaultTextStyle=imported.font;
  const sst=wrels.filter(r=>r.type===REL+'/sharedStrings');requireProfile(sst.length<=1,'Duplicate shared strings relationship');let shared:string[]=[];
  if(sst.length){const e=root(opc.parts.get(sst[0].targetPart!)!,'sst',options);attrs(e,['count','uniqueCount']);shared=nodes(e,['si']).map(si=>{attrs(si,[]);return decodeText(leaf(one(nodes(si,['t']),'t')!,['xml:space']));});if(get(e,'uniqueCount')!==undefined)requireProfile(integer(get(e,'uniqueCount'),0,1000000)===shared.length,'Shared string count mismatch');}
  const sheets=one(children,'sheets')!;attrs(sheets,[]);const defs=nodes(sheets,['sheet']);requireProfile(defs.length>=1&&defs.length<=128,'Sheet count outside profile');document.sheets=[];
  const used=new Set<string>(),ids=new Set<number>(),cachedValues:[string,CalculatedValue][]=[];
  let cellCount=0;
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
    plain(def,['name','sheetId','r:id']);const name=get(def,'name');requireProfile(name,'Missing sheet name');
    const sheetNo=integer(get(def,'sheetId'),1,2147483647);requireProfile(!ids.has(sheetNo),'Duplicate sheet ID');ids.add(sheetNo);
    const id=def.getAttributeNS(REL,'id'),rel=wrels.find(r=>r.id===id&&r.type===REL+'/worksheet');requireProfile(rel?.targetPart&&!used.has(rel.targetPart),'Missing or duplicate sheet relationship');used.add(rel.targetPart);
    const sheet:Sheet={id:'sheet-'+(i+1),name,cells:{},rows:{},columns:{},merges:[]};
    const top=root(opc.parts.get(rel.targetPart)!,'worksheet',options);attrs(top,[]);const elements=nodes(top,['sheetViews','cols','sheetData','mergeCells','dimension']);
    const dim=one(elements,'dimension',false);if(dim)plain(dim,['ref']);
    const views=one(elements,'sheetViews',false);if(views){attrs(views,[]);const view=one(nodes(views,['sheetView']),'sheetView')!;attrs(view,['workbookViewId','rightToLeft','showGridLines']);requireProfile(get(view,'workbookViewId')==='0','Unsupported workbook view');sheet.view={};if(get(view,'rightToLeft')!==undefined)sheet.view.rightToLeft=flag(get(view,'rightToLeft'));if(get(view,'showGridLines')!==undefined)sheet.view.showGridLines=flag(get(view,'showGridLines'));const pane=one(nodes(view,['pane']),'pane',false);if(pane){plain(pane,['xSplit','ySplit','topLeftCell','activePane','state']);requireProfile(get(pane,'state')==='frozen','Unsupported split pane');const fc=integer(get(pane,'xSplit')??'0',0,16383),fr=integer(get(pane,'ySplit')??'0',0,1048575);requireProfile(fr||fc,'Empty frozen pane');requireProfile(get(pane,'topLeftCell')===address(fr+1,fc+1),'Noncanonical frozen pane origin');requireProfile(get(pane,'activePane')===(fr&&fc?'bottomRight':fr?'bottomLeft':'topRight'),'Unexpected active pane');sheet.view.frozenRows=fr;sheet.view.frozenColumns=fc;}if(!Object.keys(sheet.view).length)delete sheet.view;}
    const cols=one(elements,'cols',false);if(cols){attrs(cols,[]);for(const e of nodes(cols,['col'])){plain(e,['min','max','width','customWidth','hidden']);const start=integer(get(e,'min'),1,16384),end=integer(get(e,'max'),start,16384);const width=get(e,'width'),hidden=get(e,'hidden');for(let col=start;col<=end;col++){requireProfile(!sheet.columns[col],'Overlapping column metadata');sheet.columns[col]={...(width!==undefined?{width:Math.round(decimal(width,.00001,255)*9525*7)}:{}),...(hidden!==undefined?{hidden:flag(hidden)}:{})};}}}
    const data=one(elements,'sheetData')!;attrs(data,[]);const rowSeen=new Set<number>();for(const row of nodes(data,['row'])){attrs(row,['r','ht','customHeight','hidden']);const r=integer(get(row,'r'),1,1048576);requireProfile(!rowSeen.has(r),'Duplicate row');rowSeen.add(r);const ht=get(row,'ht'),hidden=get(row,'hidden');if(ht!==undefined||hidden!==undefined)sheet.rows[r]={...(ht!==undefined?{height:Math.round(decimal(ht,.00001,409)*12700)}:{}),...(hidden!==undefined?{hidden:flag(hidden)}:{})};
      for(const el of nodes(row,['c'])){requireProfile(++cellCount<=100000,'Cell budget exceeded');attrs(el,['r','s','t']);const [rr,c]=coordinate(get(el,'r'));requireProfile(rr===r,'Cell row mismatch');const key=r+':'+c;requireProfile(!sheet.cells[key],'Duplicate cell');const st=imported.styles[integer(get(el,'s')??'0',0,imported.styles.length-1)];const type=get(el,'t'),ch=nodes(el,['f','v','is']),f=one(ch,'f',false),v=one(ch,'v',false),is=one(ch,'is',false);let value:Cell['value'];
        if(f){requireProfile(!is&&type!=='s'&&type!=='inlineStr','Invalid formula cell');const source='='+leaf(f);requireProfile(formulaIsSupported(parseFormula(source)),'Unknown formula outside editable profile');value={kind:'formula',source};if(v)cachedValues.push([calculationKey(sheet.id,r,c),scalar(type,leaf(v))]);}
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
  return {document:validateDocument(document) as SheetsDocument,cachedValues};
}
export function saveSheetsXlsx(source:Uint8Array,model:WorkbookSnapshot):Uint8Array{openSheetsXlsx(source,model.documentId,model.title);return writeSheetsXlsx(model);}
