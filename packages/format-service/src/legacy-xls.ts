// SPDX-License-Identifier: MIT
/** Original, closed BIFF8 workbook profile (MS-XLS BOF, BOUNDSHEET, SST,
 * XF, FORMULA/Ptg, WINDOW2, PANE, and cell records).
 * https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/
 * Exact reconstruction gates semantic edits; unknown records are never discarded.
 */
import {assert,assertUnicode,OfficeError} from '@office-ex/contracts';
import {createDocument,validateDocument,type SheetsDocument,type Sheet,type Cell,type CellStyle,type Color,type TextStyle} from '@office-ex/core';
import {readCfb,writeCfb,type CfbOptions} from '@office-ex/cfb';
import {parseFormula,serializeFormula,formulaIsSupported,WorkbookCalculator,calculationKey,type FormulaNode,type FormulaReference,type BinaryOperator,type FormulaError,type CalculatedValue,type WorkbookSnapshot} from '@office-ex/sheets';
const MAX=32*1024*1024, MAX_RECORD=8224;
const allowedFont=['fontFamily','fontSize','bold','italic','underline','color'];
const allowedStyle=[...allowedFont,'alignment','verticalAlignment','wrapText','fill','border'];
const errors:Record<FormulaError,number>={'#NULL!':0,'#DIV/0!':7,'#VALUE!':15,'#REF!':23,'#NAME?':29,'#NUM!':36,'#N/A':42};
const funcs:Record<string,number>={COUNT:0,IF:1,SUM:4,AVERAGE:5,MIN:6,MAX:7,ABS:24,ROUND:27,AND:36,OR:37,NOT:38,COUNTA:169};
const operators:Record<BinaryOperator,number>={'+':3,'-':4,'*':5,'/':6,'^':7,'&':8,'<':9,'<=':10,'=':11,'>=':12,'>':13,'<>':14};
const builtin:Record<number,string>={0:'General',1:'0',2:'0.00',3:'#,##0',4:'#,##0.00',9:'0%',10:'0.00%',11:'0.00E+00',14:'mm-dd-yy',20:'h:mm',21:'h:mm:ss',22:'m/d/yy h:mm',49:'@'};
function ok(v:unknown,m:string):asserts v{if(!v)throw new OfficeError('unsupported','XLS profile: '+m);}
function bounds(b:Uint8Array,p:number,n:number){assert(Number.isSafeInteger(p)&&Number.isSafeInteger(n)&&p>=0&&n>=0&&p+n<=b.length,'invalid','BIFF8 record bounds');}
function dv(b:Uint8Array){return new DataView(b.buffer,b.byteOffset,b.byteLength);}
function u16(b:Uint8Array,p:number){bounds(b,p,2);return dv(b).getUint16(p,true);}
function u32(b:Uint8Array,p:number){bounds(b,p,4);return dv(b).getUint32(p,true);}
function f64(b:Uint8Array,p:number){bounds(b,p,8);return dv(b).getFloat64(p,true);}
function w16(v:number){const b=new Uint8Array(2);dv(b).setUint16(0,v,true);return b;}
function w32(v:number){const b=new Uint8Array(4);dv(b).setUint32(0,v,true);return b;}
function w64(v:number){const b=new Uint8Array(8);dv(b).setFloat64(0,v,true);return b;}
function concat(a:readonly Uint8Array[]){const n=a.reduce((s,b)=>s+b.length,0);assert(n<=MAX,'limit','BIFF8 stream budget');const o=new Uint8Array(n);let p=0;for(const b of a){o.set(b,p);p+=b.length;}return o;}
function cat(...a:Uint8Array[]){return concat(a);}
function rec(t:number,b:Uint8Array=new Uint8Array()){ok(b.length<=MAX_RECORD,'Record exceeds BIFF8 limit');return cat(w16(t),w16(b.length),b);}
function bof(kind:number){return rec(0x0809,cat(w16(0x600),w16(kind),w16(0x0dbb),w16(0x07cc),w32(0x41),w32(6)));}
function text(s:string,short=false):Uint8Array{assertUnicode(s);ok(s.length<=(short?255:32767),'String exceeds profile');const wide=[...s].some(c=>c.charCodeAt(0)>255);const b=new Uint8Array(s.length*(wide?2:1));for(let i=0;i<s.length;i++){if(wide)dv(b).setUint16(i*2,s.charCodeAt(i),true);else b[i]=s.charCodeAt(i);}return cat(short?Uint8Array.of(s.length):w16(s.length),Uint8Array.of(Number(wide)),b);}
function readText(b:Uint8Array,p:number,short=false):[string,number]{bounds(b,p,short?2:3);const n=short?b[p]:u16(b,p);p+=short?1:2;const flag=b[p++];ok(flag===0||flag===1,'Rich/extended strings unsupported');const size=n*(flag?2:1);bounds(b,p,size);let s='';for(let i=0;i<n;i++)s+=String.fromCharCode(flag?u16(b,p+i*2):b[p+i]);assertUnicode(s);return [s,p+size];}
function eq(a:Uint8Array,b:Uint8Array){return a.length===b.length&&a.every((x,i)=>x===b[i]);}
function inverse<T extends string>(m:Record<T,number>,v:number):T{const k=Object.keys(m).find(k=>m[k as T]===v);ok(k!==undefined,'Unsupported formula/error token');return k as T;}
function coord(r:number,c:number){ok(Number.isInteger(r)&&r>=1&&r<=65536&&Number.isInteger(c)&&c>=1&&c<=256,'BIFF8 row/column limit');return cat(w16(r-1),w16(c-1));}
function refCol(r:FormulaReference){coord(r.row,r.column);return r.column-1|(r.columnAbsolute?0:0x4000)|(r.rowAbsolute?0:0x8000);}
function formula(source:string,names:readonly string[]):Uint8Array{
  const ast=parseFormula(source);ok(formulaIsSupported(ast),'Formula outside supported profile');
  const parts:Uint8Array[]=[];const push=(...p:Uint8Array[])=>parts.push(...p);
  const sheet=(name:string)=>{const i=names.findIndex(n=>n.toLowerCase()===name.toLowerCase());ok(i>=0,'Missing internal sheet');return w16(i);};
  function node(n:FormulaNode){switch(n.kind){
    case 'number':push(Number.isInteger(n.value)&&n.value>=0&&n.value<=65535?cat(Uint8Array.of(0x1e),w16(n.value)):cat(Uint8Array.of(0x1f),w64(n.value)));break;
    case 'string':push(Uint8Array.of(0x17),text(n.value,true));break;
    case 'boolean':push(Uint8Array.of(0x1d,Number(n.value)));break;
    case 'error':push(Uint8Array.of(0x1c,errors[n.value]));break;
    case 'missing':push(Uint8Array.of(0x16));break;
    case 'reference':push(Uint8Array.of(n.sheet?0x5a:0x44),...(n.sheet?[sheet(n.sheet)]:[]),w16(n.row-1),w16(refCol(n)));break;
    case 'range':ok(!n.axis&&n.start.sheet===n.end.sheet,'Whole-axis or three-dimensional ranges cannot be converted without semantic change');
      push(Uint8Array.of(n.start.sheet?0x3b:0x25),...(n.start.sheet?[sheet(n.start.sheet)]:[]),w16(n.start.row-1),w16(n.end.row-1),w16(refCol(n.start)),w16(refCol(n.end)));break;
    case 'unary':node(n.operand);push(Uint8Array.of(n.operator==='+'?0x12:n.operator==='-'?0x13:0x14));break;
    case 'binary':node(n.left);node(n.right);push(Uint8Array.of(operators[n.operator]));break;
    case 'call':ok(funcs[n.name]!==undefined&&n.args.length<=30,'Unsupported BIFF8 function/argument count');for(const a of n.args)node(a);push(Uint8Array.of(0x42,n.args.length),w16(funcs[n.name]));break;
  }}
  node(ast);const b=concat(parts);ok(b.length<=MAX_RECORD-22,'Formula token budget');return b;
}
function readFormula(b:Uint8Array,names:readonly string[]):string{
  const stack:FormulaNode[]=[];let p=0,tokens=0;
  const pop=()=>{const n=stack.pop();ok(n,'Malformed postfix formula');return n;};
  const ref=(r:number,c:number,name?:string):FormulaReference=>{ok((c&0x3fff)<256,'Unsupported BIFF8 column token');return {kind:'reference',...(name===undefined?{}:{sheet:name}),row:r+1,column:(c&255)+1,rowAbsolute:!(c&0x8000),columnAbsolute:!(c&0x4000)};};
  while(p<b.length){ok(++tokens<=4096,'Formula token budget');const t=b[p++];let n:FormulaNode;
    if(t===0x1e){n={kind:'number',value:u16(b,p)};p+=2;}
    else if(t===0x1f){n={kind:'number',value:f64(b,p)};p+=8;}
    else if(t===0x17){const [v,e]=readText(b,p,true);p=e;n={kind:'string',value:v};}
    else if(t===0x1d||t===0x1c){bounds(b,p,1);const v=b[p++];if(t===0x1d)ok(v<2,'Invalid boolean');n=t===0x1d?{kind:'boolean',value:!!v}:{kind:'error',value:inverse(errors,v)};}
    else if(t===0x16)n={kind:'missing'};
    else if(t===0x44||t===0x5a||t===0x25||t===0x3b){
      let name:string|undefined;if(t===0x5a||t===0x3b){name=names[u16(b,p)];p+=2;ok(name!==undefined,'External sheet index');}
      if(t===0x44||t===0x5a){n=ref(u16(b,p),u16(b,p+2),name);p+=4;}
      else{n={kind:'range',start:ref(u16(b,p),u16(b,p+4),name),end:ref(u16(b,p+2),u16(b,p+6),name)};p+=8;}
    }else if(t>=3&&t<=14){const right=pop(),left=pop();n={kind:'binary',operator:inverse(operators,t),left,right};}
    else if(t>=0x12&&t<=0x14)n={kind:'unary',operator:t===0x12?'+':t===0x13?'-':'%',operand:pop()};
    else if(t===0x42){bounds(b,p,3);const count=b[p++],name=inverse(funcs,u16(b,p));p+=2;ok(count<=30&&count<=stack.length,'Invalid function arguments');n={kind:'call',name,args:stack.splice(stack.length-count,count)};}
    else throw new OfficeError('unsupported','XLS profile: Unknown formula token');
    stack.push(n);
  }
  ok(stack.length===1,'Formula stack did not resolve');return serializeFormula(stack[0]);
}
type Style={style?:CellStyle;numberFormat?:string};
function styleTables(doc:WorkbookSnapshot){
  const palette:Color[]=[];
  const color=(c:Color|undefined,def:number)=>{if(c===undefined)return def;ok(c.kind==='rgba'&&c.alpha===255,'Only opaque RGB colors in XLS profile');let i=palette.findIndex(p=>p.kind==='rgba'&&p.red===c.red&&p.green===c.green&&p.blue===c.blue);if(i<0){i=palette.length;ok(i<56,'BIFF8 palette exhausted');palette.push(c);}return 8+i;};
  const fontRecords:Uint8Array[]=[],formats=new Map<string,number>(),xfRecords:Uint8Array[]=[];
  const font=(f:TextStyle)=>{const sz=Math.round((f.fontSize??139700)/635);ok(sz>0&&sz<=8180,'Font size outside BIFF8 profile');const name=f.fontFamily??'Arial';ok(name.length>0&&name.length<=31,'Font name length');const b=cat(w16(sz),w16((f.bold?1:0)|(f.italic?2:0)|(f.underline?4:0)),w16(color(f.color,0x7fff)),w16(f.bold?700:400),w16(0),Uint8Array.of(f.underline?1:0,0,0,0),text(name,true));let i=fontRecords.findIndex(x=>eq(x,b));if(i<0){i=fontRecords.length;ok(i<255,'Font budget');fontRecords.push(b);}return i>=4?i+1:i;};
  const base=doc.defaultTextStyle;ok(Object.keys(base).every(k=>allowedFont.includes(k)),'Unsupported workbook default style');font(base);
  function formatId(format:string){const k=Object.keys(builtin).find(i=>builtin[Number(i)]===format);if(k!==undefined)return Number(k);if(!formats.has(format)){ok(formats.size<1000,'Number format budget');formats.set(format,164+formats.size);}return formats.get(format)!;}
  function xf(c:Cell|undefined,style=false){
    const s=c?.style??{};ok(!c?.styleId&&Object.keys(s).every(k=>allowedStyle.includes(k)),'Unsupported cell style');
    const f:TextStyle={...base};for(const k of allowedFont)if(s[k as keyof TextStyle]!==undefined)Object.assign(f,{[k]:s[k as keyof TextStyle]});
    const h=s.alignment?({start:1,center:2,end:3,justify:5})[s.alignment]:0;
    const v=s.verticalAlignment?({top:0,center:1,bottom:2})[s.verticalAlignment]:2;
    let b1=0,b2=0,fill=0x20c0;
    if(s.border){const line=({9525:1,25400:2,38100:5} as Record<number,number>)[s.border.width];ok(line,'Only uniform thin, medium or thick border supported');const c=color(s.border.color,0x40);b1=line*0x1111|(c<<16)|(c<<23);b2=c|(c<<7);}
    if(s.fill){b2|=1<<26;fill=color(s.fill,0x40)|(0x41<<7);}
    return cat(w16(font(f)),w16(formatId(c?.numberFormat??'General')),w16(style?0xfff5:1),Uint8Array.of(h|(s.wrapText?8:0)|(v<<4),0,0,style?0xf4:0xf8),w32(b1),w32(b2),w16(fill));
  }
  for(let i=0;i<16;i++)xfRecords.push(xf(undefined,true));xfRecords.push(xf(undefined));
  function cellStyle(c:Cell){const b=xf(c);let i=xfRecords.findIndex((x,i)=>i>=16&&eq(x,b));if(i<0){i=xfRecords.length;ok(i<4000,'BIFF8 style budget');xfRecords.push(b);}return i;}
  function records(){const colors=new Uint8Array(56*4);for(let i=0;i<palette.length;i++){const c=palette[i] as Extract<Color,{kind:'rgba'}>;colors.set([c.red,c.green,c.blue,0],i*4);}
    return [...fontRecords.map(b=>rec(0x31,b)),...Array.from(formats,([s,i])=>rec(0x41e,cat(w16(i),text(s)))),...xfRecords.map(b=>rec(0xe0,b)),rec(0x293,cat(w16(0x8000),Uint8Array.of(0,0xff))),rec(0x92,cat(w16(56),colors))];
  }
  return {cellStyle,records};
}
function cache(v:CalculatedValue):{result:Uint8Array;extra:Uint8Array[]}{if(v.kind==='number')return {result:w64(v.value),extra:[]};const type=v.kind==='string'?0:v.kind==='boolean'?1:v.kind==='error'?2:3;return {result:Uint8Array.of(type,0,v.kind==='boolean'?Number(v.value):v.kind==='error'?errors[v.value]:0,0,0,0,255,255),extra:v.kind==='string'?[rec(0x207,text(v.value))]:[]};}
function workbook(input:WorkbookSnapshot):Uint8Array{
  const d=validateDocument(input);ok(d.kind==='sheets'&&d.sheets.length>0&&d.sheets.length<=128,'Workbook sheet limit');
  ok(!Object.keys(d.styles).length&&!Object.keys(d.resources).length&&!d.namedCellStyles,'Named styles/resources outside BIFF8 profile');
  const names=d.sheets.map(s=>s.name),styles=styleTables(d),calc=new WorkbookCalculator().calculate(d,{maxSteps:5_000_000});
  const strings:string[]=[],sstIndex=new Map<string,number>();let totalStrings=0,totalCells=0;
  const stringIndex=(s:string)=>{totalStrings++;let i=sstIndex.get(s);if(i===undefined){i=strings.length;strings.push(s);sstIndex.set(s,i);}return i;};
  const streams=d.sheets.map((s,index)=>{
    ok(!s.defaultGeometry,'Explicit worksheet default geometry outside BIFF8 profile');ok(!s.conditionalRules,'Conditional rules outside BIFF8 profile');ok(!s.charts,'Charts outside BIFF8 profile');
    const a:Uint8Array[]=[bof(0x10)];let maxR=0,maxC=0;
    const cells=Object.entries(s.cells).map(([k,c])=>{ok(!c.richText&&!c.namedStyle,'Rich text or named cell style outside BIFF8 profile');const [r,col]=k.split(':').map(Number);coord(r,col);maxR=Math.max(maxR,r);maxC=Math.max(maxC,col);return {r,col,c};}).sort((x,y)=>x.r-y.r||x.col-y.col);
    ok((totalCells+=cells.length)<=200000,'Cell budget');
    for(const [k,col] of Object.entries(s.columns).sort((a,b)=>Number(a[0])-Number(b[0]))){const c=Number(k);coord(1,c);const width=Math.round((col.width??64*9525)/(7*9525)*256);ok(width>=0&&width<=65535,'Column width limit');a.push(rec(0x7d,cat(w16(c-1),w16(c-1),w16(width),w16(16),w16(col.hidden?1:0),w16(0))));}
    for(const k of Object.keys(s.rows)){const r=Number(k);coord(r,1);maxR=Math.max(maxR,r);}
    for(const m of s.merges){coord(m.row,m.column);coord(m.row+m.rowCount-1,m.column+m.columnCount-1);maxR=Math.max(maxR,m.row+m.rowCount-1);maxC=Math.max(maxC,m.column+m.columnCount-1);}
    a.push(rec(0x200,cat(w32(0),w32(maxR),w16(0),w16(maxC),w16(0))));
    const rows=new Map<number,typeof cells>();for(const c of cells){const r=rows.get(c.r)??[];r.push(c);rows.set(c.r,r);}for(const k of Object.keys(s.rows))if(!rows.has(Number(k)))rows.set(Number(k),[]);
    for(const [r,cc] of [...rows].sort((a,b)=>a[0]-b[0])){const rr=s.rows[String(r)];const height=Math.round((rr?.height??255*635)/635);ok(height>0&&height<=8190,'Row height limit');a.push(rec(0x208,cat(w16(r-1),w16(0),w16(cc.length?cc.at(-1)!.col:0),w16(height),w32(0),w32(0x100|(rr?.height!==undefined?0x40:0)|(rr?.hidden?0x20:0)))));
      for(const {col,c} of cc){const prefix=cat(coord(r,col),w16(styles.cellStyle(c))),v=c.value;
        if(v.kind==='number')a.push(rec(0x203,cat(prefix,w64(v.value))));
        else if(v.kind==='string')a.push(rec(0xfd,cat(prefix,w32(stringIndex(v.value)))));
        else if(v.kind==='boolean'||v.kind==='error')a.push(rec(0x205,cat(prefix,Uint8Array.of(v.kind==='boolean'?Number(v.value):errors[v.value],v.kind==='error'?1:0))));
        else if(v.kind==='blank')a.push(rec(0x201,prefix));
        else{const tokens=formula(v.source,names),res=calc.get(s.id,r,col);ok((res.state==='valid'||res.state==='error')&&res.value,'Formula cache unavailable');const cached=cache(res.value);a.push(rec(6,cat(prefix,cached.result,w16(0),w32(0),w16(tokens.length),tokens)),...cached.extra);}
      }
    }
    for(let i=0;i<s.merges.length;i+=1000){const m=s.merges.slice(i,i+1000);a.push(rec(0xe5,cat(w16(m.length),...m.map(x=>cat(w16(x.row-1),w16(x.row+x.rowCount-2),w16(x.column-1),w16(x.column+x.columnCount-2))))));}
    const v=s.view??{},fr=v.frozenRows??0,fc=v.frozenColumns??0;ok(fr<65536&&fc<256,'Frozen pane limit');
    let flags=0xb4|(v.showGridLines!==false?2:0)|(v.rightToLeft?0x40:0)|(fr||fc?0x108:0)|(index===0?0x600:0);
    a.push(rec(0x23e,cat(w16(flags),w16(0),w16(0),w32(0x40),w16(0),w16(0),w32(0))));
    if(fr||fc)a.push(rec(0x41,cat(w16(fc),w16(fr),w16(fr),w16(fc),Uint8Array.of(fr&&fc?0:fr?2:1,0))));
    a.push(rec(0xa));return concat(a);
  });
  const globals=[bof(5),rec(0x42,w16(1200)),rec(0x3d,cat(w16(0),w16(0),w16(0x3fcf),w16(0x2a4e),w16(0x38),w16(0),w16(0),w16(1),w16(0x258))),rec(0x22,w16(d.dateSystem==='1904'?1:0)),...styles.records(),rec(0x1ae,cat(w16(names.length),w16(0x0401))),rec(0x17,cat(w16(names.length),...names.map((_,i)=>cat(w16(0),w16(i),w16(i)))))];
  // SST may continue between strings only. Never split a Unicode string mid-record.
  const sst:Uint8Array[]=[];let chunk:Uint8Array[]=[w32(totalStrings),w32(strings.length)],size=8;
  for(const s of strings){const b=text(s);ok(b.length<=MAX_RECORD,'Individual SST string exceeds the supported BIFF8 record budget');if(size+b.length>MAX_RECORD){sst.push(rec(sst.length?0x3c:0xfc,concat(chunk)));chunk=[];size=0;}chunk.push(b);size+=b.length;}
  sst.push(rec(sst.length?0x3c:0xfc,concat(chunk)));
  const bound=(offset:number,name:string)=>rec(0x85,cat(w32(offset),Uint8Array.of(0,0),text(name,true)));
  const gsize=globals.reduce((n,b)=>n+b.length,0)+names.reduce((n,name)=>n+bound(0,name).length,0)+sst.reduce((n,b)=>n+b.length,0)+4;
  let offset=gsize;const refs=names.map((name,i)=>{const b=bound(offset,name);offset+=streams[i].length;return b;});
  return concat([...globals,...refs,...sst,rec(0xa),...streams]);
}
interface RecordEntry{type:number;body:Uint8Array;next:number}
function record(b:Uint8Array,p:number):RecordEntry{const type=u16(b,p),len=u16(b,p+2);ok(len<=MAX_RECORD,'BIFF8 record limit');bounds(b,p+4,len);return{type,body:b.subarray(p+4,p+4+len),next:p+4+len};}
export interface LegacyXlsSource{document:SheetsDocument;cachedValues:[string,CalculatedValue][]}
export function openLegacyXls(bytes:Uint8Array,documentId:string,title='Workbook',options:CfbOptions={}):LegacyXlsSource{
  assert(bytes.length<=MAX,'limit','XLS input limit');const cfb=readCfb(bytes,options);ok(cfb.entries.length===1&&cfb.entries[0].type==='stream'&&cfb.entries[0].path.length===1&&cfb.entries[0].path[0].toLowerCase()==='workbook','Unknown CFB streams');
  const b=cfb.entries[0].data,d=createDocument('sheets',documentId,title);let p=0,count=0;
  const fonts:TextStyle[]=[],xf:Uint8Array[]=[],formats=new Map(Object.entries(builtin).map(([k,v])=>[Number(k),v])),palette:Color[]=[],rawFonts:Uint8Array[]=[],names:string[]=[],offsets:number[]=[],strings:string[]=[];
  let sstTotal=-1,sstCount=-1;
  while(p<b.length){ok(++count<=1000000,'BIFF record count');const r=record(b,p);p=r.next;const x=r.body;
    if(r.type===0xa)break;
    if(r.type===0x22){d.dateSystem=u16(x,0)?'1904':'1900';}
    else if(r.type===0x31)rawFonts.push(x);
    else if(r.type===0x41e){const [s,e]=readText(x,2);ok(e===x.length,'Format trailing bytes');formats.set(u16(x,0),s);}
    else if(r.type===0xe0)xf.push(x);
    else if(r.type===0x92){const n=u16(x,0);ok(n===56&&x.length===226,'Unsupported palette');for(let i=0;i<n;i++)palette.push({kind:'rgba',red:x[2+4*i],green:x[3+4*i],blue:x[4+4*i],alpha:255});}
    else if(r.type===0x85){ok(x[4]===0&&x[5]===0,'Hidden or non-worksheet sheet');offsets.push(u32(x,0));const [s,e]=readText(x,6,true);ok(e===x.length,'Bound sheet trailing bytes');names.push(s);}
    else if(r.type===0xfc||r.type===0x3c){let q=0;if(r.type===0xfc){ok(sstCount<0,'Duplicate SST');sstTotal=u32(x,0);sstCount=u32(x,4);q=8;}ok(sstCount>=0&&sstCount<=200000&&sstTotal>=sstCount,'SST limit');while(q<x.length){const [s,e]=readText(x,q);strings.push(s);q=e;ok(strings.length<=sstCount,'SST size mismatch');}}
    else ok([0x809,0x42,0x3d,0x293,0x1ae,0x17].includes(r.type),'Unknown workbook-global record '+r.type.toString(16));
  }
  ok(names.length>0&&names.length<=128&&offsets[0]===p&&strings.length===sstCount&&rawFonts.length>0&&xf.length>=17,'Workbook profile structure');
  const color=(i:number)=>{ok(i>=8&&i<8+palette.length,'Unsupported palette index');return palette[i-8];};
  for(const f of rawFonts){const [name,e]=readText(f,14,true);ok(e===f.length,'Font trailing bytes');const s:TextStyle={fontFamily:name,fontSize:u16(f,0)*635,bold:u16(f,6)===700,italic:!!(u16(f,2)&2),underline:f[10]===1};if(u16(f,4)!==0x7fff)s.color=color(u16(f,4));fonts.push(s);}
  d.defaultTextStyle=fonts[0];
  function style(i:number):Style{ok(i>=16&&i<xf.length,'Invalid cell XF');const x=xf[i];ok(x.length===20,'Invalid XF length');const fi=u16(x,0);ok(fi!==4&&fi-(fi>4?1:0)<fonts.length,'Invalid font index');const f=fonts[fi-(fi>4?1:0)],s:CellStyle={};
    if(fi!==0)Object.assign(s,f);
    const h=x[6]&7,v=(x[6]>>4)&7;ok([0,1,2,3,5].includes(h)&&v<=2,'Unsupported alignment');if(h)s.alignment=({1:'start',2:'center',3:'end',5:'justify'} as const)[h as 1|2|3|5];if(v!==2)s.verticalAlignment=v===0?'top':'center';if(x[6]&8)s.wrapText=true;
    const b1=u32(x,10),b2=u32(x,14);if(b1&0xffff){const w=({1:9525,2:25400,5:38100} as Record<number,number>)[b1&15];ok(w,'Unsupported border');s.border={width:w,color:color((b1>>>16)&127)};}if((b2>>>26)&63)s.fill=color(u16(x,18)&127);
    const format=formats.get(u16(x,2));ok(format!==undefined,'Unknown number format');return {...(Object.keys(s).length?{style:s}:{}),...(format!=='General'?{numberFormat:format}:{})};
  }
  const cachedValues:[string,CalculatedValue][]=[];d.sheets=[];let actualStrings=0;
  for(let i=0;i<names.length;i++){p=offsets[i];const end=offsets[i+1]??b.length;const s:Sheet={id:'sheet-'+(i+1),name:names[i],rows:{},columns:{},cells:{},merges:[]};let eof=false,seen=0;
    while(p<end){ok(++count<=1000000,'BIFF record count');const r=record(b,p);p=r.next;const x=r.body;ok(p<=end,'Overlapping worksheet records');if(r.type===0xa){eof=true;break;}
      if(r.type===0x7d){const c=u16(x,0)+1;ok(u16(x,2)===c-1,'Grouped column ranges outside profile');s.columns[String(c)]={width:Math.round(u16(x,4)/256*7*9525),...(u16(x,8)&1?{hidden:true}:{})};}
      else if(r.type===0x208){const flags=u32(x,12),row=u16(x,0)+1;if(flags&0x60)s.rows[String(row)]={...(flags&0x40?{height:u16(x,6)*635}:{}),...(flags&0x20?{hidden:true}:{})};}
      else if(r.type===0xe5){const n=u16(x,0);ok(n*8+2===x.length,'Merge record length');for(let j=0;j<n;j++){const q=2+8*j,row=u16(x,q)+1,column=u16(x,q+4)+1;s.merges.push({row,column,rowCount:u16(x,q+2)+2-row,columnCount:u16(x,q+6)+2-column});}}
      else if(r.type===0x23e){const f=u16(x,0);s.view={...(f&0x40?{rightToLeft:true}:{}),...(!(f&2)?{showGridLines:false}:{})};}
      else if(r.type===0x41){s.view={...s.view,frozenRows:u16(x,2),frozenColumns:u16(x,0)};}
      else if([0x203,0xfd,0x205,0x201,6].includes(r.type)){
        ok(++seen<=200000,'Worksheet cell budget');const row=u16(x,0)+1,col=u16(x,2)+1,k=row+':'+col;ok(!s.cells[k],'Duplicate cell');let value:Cell['value'];
        if(r.type===0x203)value={kind:'number',value:f64(x,6)};
        else if(r.type===0xfd){const str=strings[u32(x,6)];ok(str!==undefined,'SST index');actualStrings++;value={kind:'string',value:str};}
        else if(r.type===0x205){ok(x.length===8&&x[7]<=1,'Boolean/error record');value=x[7]?{kind:'error',value:inverse(errors,x[6])}:{kind:'boolean',value:!!x[6]};}
        else if(r.type===0x201)value={kind:'blank'};
        else{const length=u16(x,20);ok(length+22===x.length,'Formula record length');value={kind:'formula',source:readFormula(x.subarray(22),names)};let v:CalculatedValue;
          if(x[12]===255&&x[13]===255){const kind=x[6];if(kind===0){const str=record(b,p);ok(str.type===0x207&&str.next<=end,'Missing formula STRING cache');const [strValue,e]=readText(str.body,0);ok(e===str.body.length,'Formula STRING trailing bytes');p=str.next;v={kind:'string',value:strValue};}else if(kind===1)v={kind:'boolean',value:!!x[8]};else if(kind===2)v={kind:'error',value:inverse(errors,x[8])};else{ok(kind===3,'Unsupported formula cache');v={kind:'blank'};}}
          else v={kind:'number',value:f64(x,6)};cachedValues.push([calculationKey(s.id,row,col),v]);
        }
        s.cells[k]={value,...style(u16(x,4))};
      }else ok(r.type===0x809||r.type===0x200,'Unknown worksheet record '+r.type.toString(16));
    }
    ok(eof&&p===end,'Worksheet stream not closed');if(s.view&&!Object.keys(s.view).length)delete s.view;d.sheets.push(s);
  }
  ok(actualStrings===sstTotal,'SST occurrence count mismatch');const doc=validateDocument(d) as SheetsDocument;
  // Prove no unmodeled flags, records, cached results or dependencies would be lost.
  ok(eq(workbook(doc),b),'Source is not the lossless closed BIFF8 profile');return {document:doc,cachedValues};
}
export function writeLegacyXls(candidate:WorkbookSnapshot,options:CfbOptions={}):Uint8Array{return writeCfb({entries:[{type:'stream',path:['Workbook'],data:workbook(candidate)}]},options);}
export function saveLegacyXls(source:Uint8Array,model:WorkbookSnapshot):Uint8Array{const original=openLegacyXls(source,model.documentId,model.title);if(JSON.stringify(original.document)===JSON.stringify(model))return source.slice();return writeLegacyXls(model);}
