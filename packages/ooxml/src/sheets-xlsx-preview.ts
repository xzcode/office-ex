// SPDX-License-Identifier: MIT
/** Non-editable value projection. This is deliberately NOT a semantic save
 * model: unknown layout/features stay source-owned and no formula executes. */
import {OfficeError} from '@office-ex/contracts';
import type {Element} from '@xmldom/xmldom';
import {columnNumber} from '@office-ex/sheets';
import {openOpc} from './opc.js';
import {parseXml,childElements} from './xml.js';
import {checkCancelled,type ContainerOptions} from './limits.js';
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',CT='application/vnd.openxmlformats-officedocument.spreadsheetml.';
const MAX_SHEETS=128,MAX_CELLS=100000,MAX_SHEET_CELLS=30000,MAX_TEXT=3*1024*1024,MAX_VALUE=32767;
export interface XlsxPreviewCell {value:string|null;kind:'value'|'formula-cache'|'formula-uncached'|'unsupported';formula?:string}
export interface XlsxPreviewSheet {name:string;cells:Record<string,XlsxPreviewCell>;firstRow:number;firstColumn:number;lastRow:number;lastColumn:number;truncated:boolean}
export interface XlsxValuePreview {sheets:XlsxPreviewSheet[];truncated:boolean}
function valid(value:unknown,message:string):asserts value{if(!value)throw new OfficeError('unsupported','XLSX value preview: '+message);}
function children(e:Element,name:string):Element[]{return childElements(e).filter(c=>c.namespaceURI===S&&c.localName===name);}
function one(e:Element,name:string,required=false):Element|undefined{const list=children(e,name);valid(list.length<2&&(!required||list.length===1),'Invalid '+name+' cardinality');return list[0];}
function number(s:string|null,min:number,max:number):number{valid(s!==null&&/^\d+$/.test(s),'Invalid numeric index');const n=Number(s);valid(Number.isSafeInteger(n)&&n>=min&&n<=max,'Index out of bounds');return n;}
function address(s:string|null):[number,number]{const m=/^([A-Z]{1,3})([1-9]\d*)$/.exec(s??'');valid(m,'Missing or invalid cell reference');return [number(m[2]!,1,1048576),columnNumber(m[1]!)];}
function root(bytes:Uint8Array|undefined,name:string,options:ContainerOptions):Element{valid(bytes,'Missing worksheet data');const e=parseXml(bytes,options).documentElement;valid(e?.namespaceURI===S&&e.localName===name,'Unexpected '+name+' root');return e;}
function decode(s:string):string{
 // OOXML escape sequences are interpreted once, never HTML or expressions.
 return s.replace(/_x([0-9a-f]{4})_/gi,(_,hex:string)=>String.fromCharCode(parseInt(hex,16))).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g,'�');
}
function text(e:Element):string{valid(!childElements(e).length,'Invalid scalar text');return e.textContent??'';}
function rich(e:Element):string{
 const direct=one(e,'t'),runs=children(e,'r');valid(!(direct&&runs.length),'Ambiguous string representation');
 if(direct)return decode(text(direct));
 return decode(runs.map(r=>text(one(r,'t',true)!)).join(''));
}
/** Reads only direct worksheet cells through validated OPC relationships.
 * Never follows worksheet dependencies, hyperlinks, external links, or any
 * executable content. The display is raw values, not Excel formatting. */
export function previewSheetsXlsx(bytes:Uint8Array,options:ContainerOptions={}):XlsxValuePreview{
 const opc=openOpc(bytes,options);
 valid(opc.kind==='xlsx'&&opc.conformance==='transitional'&&opc.mainPart&&!opc.hasMacros,'Unsupported preview container');
 const rels=opc.relationships.get(opc.mainPart)??[],wb=root(opc.parts.get(opc.mainPart),'workbook',options),defs=children(one(wb,'sheets',true)!,'sheet');
 valid(defs.length>0&&defs.length<=4096,'Worksheet count budget');
 const sharedRels=rels.filter(r=>r.type===R+'/sharedStrings');valid(sharedRels.length<=1,'Ambiguous shared strings');
 let shared:Element[]=[];
 if(sharedRels.length){
  const r=sharedRels[0]!;valid(r.mode==='Internal'&&r.targetPart&&!r.fragment&&opc.contentTypes.get(r.targetPart)===CT+'sharedStrings+xml','Invalid shared string relation');
  shared=children(root(opc.parts.get(r.targetPart),'sst',options),'si');
  valid(shared.length<=1000000,'Shared string budget');
 }
 const used=new Set<string>(),sheets:XlsxPreviewSheet[]=[];let total=0,totalText=0,truncated=false;
 for(const def of defs){
  checkCancelled(options);if((def.getAttribute('state')||'visible')!=='visible')continue;
  if(sheets.length>=MAX_SHEETS){truncated=true;break;}
  const name=def.getAttribute('name')??'',id=def.getAttributeNS(R,'id');
  valid(name.length>0&&name.length<=31,'Invalid sheet name');
  const matches=rels.filter(r=>r.id===id&&r.type===R+'/worksheet');if(matches.length!==1){truncated=true;continue;}
  const rel=matches[0]!;
  valid(rel.mode==='Internal'&&rel.targetPart&&!rel.fragment&&!used.has(rel.targetPart)&&opc.contentTypes.get(rel.targetPart)===CT+'worksheet+xml','Invalid worksheet relation');
  used.add(rel.targetPart);
  const ws=root(opc.parts.get(rel.targetPart),'worksheet',options),data=one(ws,'sheetData',true)!;
  const result:XlsxPreviewSheet={name,cells:Object.create(null),firstRow:1048576,firstColumn:16384,lastRow:1,lastColumn:1,truncated:false};
  const hiddenCols=new Uint8Array(16385),cols=one(ws,'cols');
  if(cols)for(const col of children(cols,'col'))if(['1','true'].includes(col.getAttribute('hidden')??'')){
   const min=number(col.getAttribute('min'),1,16384),max=number(col.getAttribute('max'),min,16384);hiddenCols.fill(1,min,max+1);
  }
  let count=0;
  rows:for(const row of children(data,'row')){
   checkCancelled(options);
   if(['1','true'].includes(row.getAttribute('hidden')??''))continue;
   const rn=number(row.getAttribute('r'),1,1048576);
   for(const c of children(row,'c')){
    const [r,col]=address(c.getAttribute('r'));valid(r===rn&&col<=16384,'Cell reference out of row bounds');
    if(hiddenCols[col])continue;
    const f=one(c,'f'),v=one(c,'v'),is=one(c,'is'),type=c.getAttribute('t')||'n';
    if(!f&&!v&&!is)continue;
    const key=r+':'+col;valid(!Object.hasOwn(result.cells,key),'Duplicate cell');
    let value:string|null=null,kind:XlsxPreviewCell['kind']=f?'formula-uncached':'value',formula:string|undefined;
    if(f){const expression=text(f);if(expression)formula='='+expression;}
    if(type==='inlineStr'&&is&&!f)value=rich(is);
    else if(v){const stored=text(v);if(stored!==''||type==='str'){
     if(type==='s'&&!f)value=rich(shared[number(stored,0,shared.length-1)]!);
     else if(type==='b'&&(stored==='0'||stored==='1'))value=stored==='1'?'TRUE':'FALSE';
     else if(['n','str','e','d'].includes(type))value=type==='str'?decode(stored):stored;
     else kind='unsupported';
    }}else if(!f)kind='unsupported';
    if(f&&value!==null)kind='formula-cache';
    if(value!==null&&value.length>MAX_VALUE){value=value.slice(0,MAX_VALUE);result.truncated=true;}
    if(formula&&formula.length>MAX_VALUE){formula=formula.slice(0,MAX_VALUE);result.truncated=true;}
    const length=(value?.length??0)+(formula?.length??0);
    if(total>=MAX_CELLS||count>=MAX_SHEET_CELLS||totalText+length>MAX_TEXT){result.truncated=true;break rows;}
    total++;count++;totalText+=length;
    result.cells[key]={value,kind,...(formula?{formula}:{})};
    result.firstRow=Math.min(result.firstRow,r);result.firstColumn=Math.min(result.firstColumn,col);
    result.lastRow=Math.max(result.lastRow,r);result.lastColumn=Math.max(result.lastColumn,col);
   }
  }
  if(!count)result.firstRow=result.firstColumn=1;
  truncated||=result.truncated;sheets.push(result);
 }
 valid(sheets.length,'No visible worksheet values can be previewed');
 return {sheets,truncated};
}
