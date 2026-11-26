// SPDX-License-Identifier: MIT
/** Original Word 97 binary semantic adapter.
 * MS-DOC 2.4.1, 2.5, 2.8.35, 2.9.33, 2.9.175, 2.9.271.
 * https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/
 *
 * General unencrypted DOC: bounded main-story piece-table projection only.
 * Editable profile: a closed, canonical Unicode Word97 document with one
 * section, paragraphs, basic direct CHPX/PAPX, and no other streams/objects.
 * Prove source closure by reconstructing and comparing both semantic streams.
 * Unsupported source is never rewritten, converted or presented as editable.
 */
import {assert,assertUnicode,OfficeError,EMU_PER_TWIP} from '@office-ex/contracts';
import {createDocument,validateDocument,resolveTextStyle,freezeData,cloneData,type WriterDocument,type DeepReadonly,type TextStyle,type Paragraph,type TextRun,type WriterBlock} from '@office-ex/core';
import {readCfb,writeCfb,findCfbEntry,type CfbOptions,type ParsedCfbDocument} from '@office-ex/cfb';
const MAX_TEXT=1_000_000,MAX_PIECES=100000,MAX_DOC=32*1024*1024;
const TEXT_START=1024,PAIR_START=154,PAIR_COUNT=93;
const PAIRS={styles:1,sections:6,chpx:12,papx:13,fonts:15,dop:31,assoc:32,clx:33} as const;
function valid(v:unknown,message:string):asserts v{if(!v)throw new OfficeError('unsupported','DOC profile: '+message);}
function bounds(a:Uint8Array,start:number,length:number):void{assert(Number.isSafeInteger(start)&&Number.isSafeInteger(length)&&start>=0&&length>=0&&start<=a.length&&length<=a.length-start,'invalid','DOC offset exceeds stream');}
function u16(a:Uint8Array,p:number):number{bounds(a,p,2);return a[p]!+a[p+1]!*256;}
function i16(a:Uint8Array,p:number):number{const n=u16(a,p);return n&32768?n-65536:n;}
function u32(a:Uint8Array,p:number):number{bounds(a,p,4);return (a[p]!+a[p+1]!*256+a[p+2]!*65536+a[p+3]!*16777216)>>>0;}
function put16(a:Uint8Array,p:number,n:number):void{bounds(a,p,2);new DataView(a.buffer,a.byteOffset,a.byteLength).setUint16(p,n&65535,true);}
function put32(a:Uint8Array,p:number,n:number):void{bounds(a,p,4);new DataView(a.buffer,a.byteOffset,a.byteLength).setUint32(p,n>>>0,true);}
function w16(n:number):Uint8Array{const a=new Uint8Array(2);put16(a,0,n);return a;}
function w32(n:number):Uint8Array{const a=new Uint8Array(4);put32(a,0,n);return a;}
function cat(...parts:readonly Uint8Array[]):Uint8Array{const size=parts.reduce((n,a)=>n+a.length,0);assert(size<=MAX_DOC,'limit','DOC stream budget exceeded');const out=new Uint8Array(size);let at=0;for(const a of parts){out.set(a,at);at+=a.length;}return out;}
function utf16(text:string):Uint8Array{assertUnicode(text);const b=new Uint8Array(text.length*2);for(let i=0;i<text.length;i++)put16(b,i*2,text.charCodeAt(i));return b;}
function equal(a:Uint8Array,b:Uint8Array):boolean{return a.length===b.length&&a.every((v,i)=>v===b[i]);}
function stable(value:unknown):string{if(value===null||typeof value!=='object')return JSON.stringify(value);if(Array.isArray(value))return '['+value.map(stable).join(',')+']';return '{'+Object.entries(value).sort(([a],[b])=>a<b?-1:a>b?1:0).map(([k,v])=>JSON.stringify(k)+':'+stable(v)).join(',')+'}';}
type Pair={offset:number;length:number};
interface Fib {word:Uint8Array;table:Uint8Array;tableName:string;cfb:ParsedCfbDocument;flags:number;version:number;cp:number;otherStories:number;fibEnd:number;meaningful:number;pairs:Pair[]}
function fib(bytes:Uint8Array,options:CfbOptions={}):Fib{
  assert(bytes.byteLength<=MAX_DOC,'limit','DOC semantic input exceeds budget');const cfb=readCfb(bytes,options);
  const entry=findCfbEntry(cfb,['WordDocument']);assert(entry?.type==='stream','invalid','Missing WordDocument');const word=entry.data;
  assert(word.length>=34&&u16(word,0)===0xa5ec,'invalid','Invalid Word FIB');
  const base=u16(word,2),flags=u16(word,10),tableName=flags&0x200?'1Table':'0Table';
  valid(!(flags&0x8100),'Encrypted DOC is not supported');
  valid(!(flags&3),'Word templates/AutoText are not supported');
  const tableEntry=findCfbEntry(cfb,[tableName]);assert(tableEntry?.type==='stream','invalid','Missing Word table stream');const table=tableEntry.data;
  const csw=u16(word,32);assert(csw>=14&&csw<=256,'invalid','Invalid FIB short array');let at=34+csw*2;
  const cslw=u16(word,at);at+=2;assert(cslw>=22&&cslw<=256,'invalid','Invalid FIB long array');bounds(word,at,cslw*4);
  const meaningful=u32(word,at),cp=u32(word,at+12);assert(cp>0&&cp<=MAX_TEXT&&meaningful<=word.length,'limit','DOC main story exceeds budget');
  let otherStories=0;for(const i of [4,5,7,8,9,10]){const n=u32(word,at+i*4);assert(n<=MAX_TEXT,'limit','DOC substory exceeds budget');otherStories+=n;}
  at+=cslw*4;const count=u16(word,at);at+=2;assert(count>=PAIR_COUNT&&count<=512,'invalid','Invalid FIB pair count');bounds(word,at,count*8+2);
  const pairs:Pair[]=[];for(let i=0;i<count;i++)pairs.push({offset:u32(word,at+i*8),length:u32(word,at+i*8+4)});
  at+=count*8;const newWords=u16(word,at);at+=2;bounds(word,at,newWords*2);
  const version=newWords?u16(word,at):base;
  valid([0xc1,0xd9,0x101,0x10c,0x112].includes(version),'Unsupported Word binary version');
  return {word,table,tableName,cfb,flags,version,cp,otherStories,fibEnd:at+newWords*2,meaningful,pairs};
}
function part(f:Fib,index:number,required=true):Uint8Array{
  const p=f.pairs[index];assert(p&&(!required||p.length>0),'invalid','Missing DOC FIB structure');bounds(f.table,p.offset,p.length);return f.table.subarray(p.offset,p.offset+p.length);
}
interface Piece{start:number;end:number;fc:number;compressed:boolean;prm:number}
function pieces(f:Fib):Piece[]{
  const data=part(f,PAIRS.clx);let at=0,count=0;
  while(data[at]===1){assert(++count<=4096,'limit','CLX property list budget');const length=u16(data,at+1);bounds(data,at+3,length);at+=3+length;}
  assert(data[at]===2,'invalid','Missing CLX Pcdt');const size=u32(data,at+1);at+=5;
  assert(size===data.length-at&&size>=16&&(size-4)%12===0,'invalid','Invalid PlcPcd size');
  const n=(size-4)/12;assert(n<=MAX_PIECES,'limit','DOC piece budget exceeded');const endCP=u32(data,at+n*4);
  assert(u32(data,at)===0&&endCP>=f.cp&&endCP<=f.cp+f.otherStories+1,'invalid','Invalid DOC CP coverage');
  const out:Piece[]=[];
  for(let i=0;i<n;i++){
    const start=u32(data,at+i*4),end=u32(data,at+(i+1)*4),pos=at+(n+1)*4+i*8;
    assert(end>start,'invalid','Unsorted DOC CP table');
    const raw=u32(data,pos+2),compressed=(raw&0x40000000)!==0,fc=compressed?(raw&0x3fffffff)/2:raw&0x3fffffff;
    assert(!(raw&0x80000000)&&Number.isInteger(fc)&&fc>=f.fibEnd,'invalid','Invalid DOC piece offset');
    bounds(f.word,fc,(end-start)*(compressed?1:2));assert(fc+(end-start)*(compressed?1:2)<=f.meaningful,'invalid','DOC text outside meaningful stream');
    out.push({start,end,fc,compressed,prm:u16(data,pos+6)});
  }
  return out;
}
// MS-DOC compressed pieces use the specified 8-bit character mapping,
// independent of the host locale (MS-DOC 2.9.73 FcCompressed).
const DOC_COMPRESSED=[0x80,0x81,0x201a,0x192,0x201e,0x2026,0x2020,0x2021,0x2c6,0x2030,0x160,0x2039,0x152,0x8d,0x8e,0x8f,0x90,0x2018,0x2019,0x201c,0x201d,0x2022,0x2013,0x2014,0x2dc,0x2122,0x161,0x203a,0x153,0x9d,0x9e,0x178];
function mainText(f:Fib,plc:readonly Piece[]):string{
  const chunks:string[]=[];let units:number[]=[];
  for(const p of plc){
    if(p.start>=f.cp)break;const end=Math.min(p.end,f.cp);
    for(let i=p.start;i<end;i++){
      let value=p.compressed?f.word[p.fc+i-p.start]!:u16(f.word,p.fc+(i-p.start)*2);
      if(p.compressed&&value>=0x80&&value<=0x9f)value=DOC_COMPRESSED[value-0x80]!;
      units.push(value);if(units.length>=8192){chunks.push(String.fromCharCode(...units));units=[];}
    }
  }
  chunks.push(String.fromCharCode(...units));const text=chunks.join('');
  assert(text.length===f.cp&&text.endsWith('\r'),'invalid','Invalid DOC main story terminator');assertUnicode(text);return text;
}
export interface LegacyDocText {version:number;text:string;pieceCount:number;compressedPieces:number;otherStories:number}
export function inspectLegacyDocText(bytes:Uint8Array,options:CfbOptions={}):LegacyDocText{
  const f=fib(bytes,options),plc=pieces(f);return {version:f.version,text:mainText(f,plc),pieceCount:plc.length,compressedPieces:plc.filter(p=>p.compressed).length,otherStories:f.otherStories};
}

const LIDS:Readonly<Record<string,number>>={'en-US':0x409,'en-GB':0x809,'zh-CN':0x804,'zh-TW':0x404,'ja-JP':0x411,'ko-KR':0x412,'fr-FR':0x40c,'de-DE':0x407,'es-ES':0xc0a,'pt-BR':0x416,'ru-RU':0x419,'it-IT':0x410,'ar-SA':0x401,'he-IL':0x40d,'hi-IN':0x439,'th-TH':0x41e,'vi-VN':0x42a,'id-ID':0x421,'ms-MY':0x43e,'tr-TR':0x41f,'nl-NL':0x413,'pl-PL':0x415,'uk-UA':0x422,'und':0};
const CHAR=['fontFamily','fontSize','bold','italic','underline','color','language','eastAsiaLanguage'] as const;
const PARA=['alignment','indent','spaceBefore','spaceAfter','lineHeight','lineMultiple','minLineHeight'] as const;
function allowedStyle(style:DeepReadonly<TextStyle>,keys:readonly string[]):void{valid(Object.keys(style).every(k=>keys.includes(k)),'Style property not representable in the closed DOC profile');}
function twip(n:number,min=0,max=31680):number{const x=Math.round(n/EMU_PER_TWIP);valid(Number.isSafeInteger(x)&&x>=min&&x<=max,'Page or paragraph dimension out of range');return x;}
function prl(op:number,value:number,length:1|2|4):Uint8Array{return cat(w16(op),length===1?Uint8Array.of(value):length===2?w16(value):w32(value));}
function fontTable(names:readonly string[]):Uint8Array{
  const records=names.map(name=>{valid(name.length>0&&name.length<=64&&!/[\x00-\x1f\x7f]/.test(name),'Font name outside DOC profile');
    const chars=utf16(name+'\0'),record=new Uint8Array(40+chars.length);record[0]=record.length-1;record[1]=0;put16(record,2,400);record.set(chars,40);return record;});
  return cat(w16(names.length),w16(0),...records);
}
function styleSheet():Uint8Array{
  const info=new Uint8Array(18);put16(info,0,15);put16(info,2,10);put16(info,4,1);put16(info,6,108);put16(info,8,15);
  const base=new Uint8Array(10);put16(base,2,0xfff1);put16(base,4,2);
  const style=cat(base,w16(6),utf16('Normal\0'),w16(2),w16(0),w16(0));put16(style,6,style.length);
  return cat(w16(info.length),info,w16(style.length),style,new Uint8Array(28));
}
function dop():Uint8Array{const b=new Uint8Array(500);put16(b,2,4);put16(b,10,720);put16(b,12,1252);return b;}
function association():Uint8Array{return cat(w16(0xffff),w16(18),w16(0),new Uint8Array(36));}
function charProperties(style:DeepReadonly<TextStyle>,fonts:readonly string[]):Uint8Array{
  const font=fonts.indexOf(style.fontFamily??'Arial');valid(font>=0,'Missing DOC font');
  const size=style.fontSize??152400;valid(size%6350===0&&size>=6350&&size<=3276*6350,'DOC font size requires half-point precision');
  const values=[prl(0x4a4f,font,2),prl(0x4a50,font,2),prl(0x4a51,font,2),prl(0x4a43,size/6350,2),
    prl(0x0835,style.bold?1:0,1),prl(0x0836,style.italic?1:0,1),prl(0x2a3e,style.underline?1:0,1)];
  const color=style.color;
  if(color){valid(color.kind==='rgba'&&color.alpha===255,'DOC supports opaque RGB only');values.push(prl(0x6870,color.red+color.green*256+color.blue*65536,4));}
  for(const [key,op] of [['language',0x486d],['eastAsiaLanguage',0x486e]] as const){
    if(style[key]!==undefined){const lid=LIDS[style[key]!];valid(lid!==undefined,'Language tag outside closed DOC profile');values.push(prl(op,lid,2));}
  }
  return cat(...values);
}
function paraProperties(style:DeepReadonly<TextStyle>):Uint8Array{
  const parts=[w16(0)];
  if(style.alignment!==undefined)parts.push(prl(0x2403,({start:0,center:1,end:2,justify:3})[style.alignment],1));
  if(style.indent!==undefined)parts.push(prl(0x840f,twip(style.indent,-31680),2));
  if(style.spaceBefore!==undefined)parts.push(prl(0xa413,twip(style.spaceBefore),2));
  if(style.spaceAfter!==undefined)parts.push(prl(0xa414,twip(style.spaceAfter),2));
  if(style.lineHeight!==undefined||style.minLineHeight!==undefined||style.lineMultiple!==undefined){
    const size=style.lineMultiple??(style.lineHeight!==undefined?-twip(style.lineHeight,1):twip(style.minLineHeight!,1));
    parts.push(cat(w16(0x6412),w16(size),w16(style.lineMultiple!==undefined?1:0)));
  }
  return cat(...parts);
}
interface Span {start:number;end:number;data:Uint8Array}
function fkp(spans:readonly Span[],paragraph:boolean):Uint8Array|undefined{
  const n=spans.length,step=paragraph?13:1,header=4*(n+1)+step*n;
  if(!n||n>(paragraph?29:101)||header>=510)return undefined;
  const out=new Uint8Array(512);let tail=510;
  for(let i=0;i<n;i++){
    const span=spans[i]!;put32(out,i*4,span.start);if(i===n-1)put32(out,(i+1)*4,span.end);
    const payload=paragraph?(span.data.length%2?cat(Uint8Array.of((span.data.length+1)/2),span.data):cat(Uint8Array.of(0,span.data.length/2),span.data)):cat(Uint8Array.of(span.data.length),span.data);
    if(span.data.length>254)return undefined;
    const offset=(tail-payload.length)&~1;if(offset<header||offset<=0)return undefined;
    out.set(payload,offset);out[4*(n+1)+i*step]=offset/2;tail=offset;
  }
  out[511]=n;return out;
}
function packFkps(spans:Span[],paragraph:boolean,firstPage:number):{pages:Uint8Array[];plc:Uint8Array}{
  const pages:Uint8Array[]=[],starts:number[]=[];let i=0;
  while(i<spans.length){
    let last=i+1,best=fkp(spans.slice(i,last),paragraph);valid(best,'One FKP property record exceeds page capacity');
    while(last<spans.length){const next=fkp(spans.slice(i,last+1),paragraph);if(!next)break;best=next;last++;}
    starts.push(spans[i]!.start);pages.push(best);i=last;
  }
  starts.push(spans.at(-1)!.end);
  return {pages,plc:cat(...starts.map(w32),...pages.map((_,i)=>w32(firstPage+i)))};
}
interface Streams {word:Uint8Array;table:Uint8Array}
function writeStreams(candidate:unknown):Streams{
  const checked=validateDocument(candidate);valid(checked.kind==='writer','Expected Writer model');
  valid(checked.sections.length===1&&!Object.keys(checked.resources).length&&!Object.keys(checked.styles).length&&!checked.defaultParagraphStyleId,'DOC closed profile supports one section, direct styles and no resource objects');
  allowedStyle(checked.defaultTextStyle,[...CHAR,...PARA]);
  const section=checked.sections[0]!;valid(!checked.pageStories&&!section.storyRefs&&section.pageNumberStart===undefined,'Running stories and page numbering are unsupported in the closed DOC profile');valid(section.blocks.length<=10000,'DOC paragraph budget exceeded');valid(section.pageExtras===undefined||Object.keys(section.pageExtras).every(k=>['headerDistance','footerDistance','columnSpace'].includes(k)),'Unsupported DOC section extras');
  const texts:{text:string;style:TextStyle}[][]=[],paragraphStyles:TextStyle[]=[],fonts=new Set<string>(['Arial']);
  let total=0;
  for(const block of section.blocks){
    valid(block.type==='paragraph'||block.type==='page-break','Tables are outside the closed DOC profile');
    const list:{text:string;style:TextStyle}[]=[];
    if(block.type==='page-break'){list.push({text:'\f',style:{fontFamily:'Arial',fontSize:152400}});paragraphStyles.push({});}
    else{
      valid(!block.styleId,'Named paragraph styles require a later DOC adapter');allowedStyle(block.style??{},PARA);
      const base=resolveTextStyle(checked.defaultTextStyle,{},[],block.style??{}).style;
      const paragraph=Object.fromEntries(PARA.filter(k=>base[k]!==undefined).map(k=>[k,base[k]])) as TextStyle;
      paragraphStyles.push(paragraph);
      let emptyStyle:TextStyle|undefined;
      for(const run of block.runs){
        valid(run.type!=='image','Pictures are outside the closed DOC profile');
        const direct=run.type==='text'?run.style??{}:{};
        if(run.type==='text'){valid(!run.styleId,'Named character style unsupported');allowedStyle(direct,CHAR);}
        const effective=resolveTextStyle(base,{},[],direct).style,style=Object.fromEntries(CHAR.filter(k=>effective[k]!==undefined).map(k=>[k,effective[k]])) as TextStyle;
        const value=run.type==='text'?run.text:run.type==='tab'?'\t':'\n';
        valid(!/[\x00-\x08\x0b-\x1f\x7f\uFFFC]/.test(value),'Unsupported DOC control character');
        if(value){fonts.add(style.fontFamily??'Arial');list.push({text:value.replace(/\n/g,'\v'),style});}
        else if(run.type==='text')emptyStyle=style;
      }
      if(!list.length){const style=emptyStyle??Object.fromEntries(CHAR.filter(k=>base[k]!==undefined).map(k=>[k,base[k]])) as TextStyle;fonts.add(style.fontFamily??'Arial');list.push({text:'',style});}
    }
    total+=list.reduce((n,r)=>n+r.text.length,0)+1;valid(total<=MAX_TEXT,'DOC text budget exceeded');texts.push(list);
  }
  valid(texts.length>0,'Empty DOC section');
  const names=['Arial',...[...fonts].filter(x=>x!=='Arial').sort()];valid(names.length<=256,'DOC font count budget');
  let cp=0,logical='';const char:Span[]=[],para:Span[]=[];
  const append=(text:string,data:Uint8Array)=>{
    if(!text)return;const start=TEXT_START+cp*2;cp+=text.length;logical+=text;const end=TEXT_START+cp*2,previous=char.at(-1);
    if(previous&&equal(previous.data,data)&&previous.end===start)previous.end=end;else char.push({start,end,data});
  };
  for(let i=0;i<texts.length;i++){
    const list=texts[i]!,start=TEXT_START+cp*2;
    for(const r of list)append(r.text,charProperties(r.style,names));
    append('\r',charProperties(list.at(-1)!.style,names));
    para.push({start,end:TEXT_START+cp*2,data:paraProperties(paragraphStyles[i]!)});
  }
  valid(char.length<=10000,'DOC character format span budget exceeded');
  const text=utf16(logical),endText=TEXT_START+text.length,firstPage=Math.ceil(endText/512);
  const ch=packFkps(char,false,firstPage),pa=packFkps(para,true,firstPage+ch.pages.length);
  const sepOffset=(firstPage+ch.pages.length+pa.pages.length)*512;
  const p=section.page,extra=section.pageExtras;
  const sectionPr=cat(prl(0xb01f,twip(p.width,1),2),prl(0xb020,twip(p.height,1),2),prl(0xb021,twip(p.margins.left),2),prl(0xb022,twip(p.margins.right),2),prl(0x9023,twip(p.margins.top),2),prl(0x9024,twip(p.margins.bottom),2),prl(0xb025,0,2),prl(0xb017,twip(extra?.headerDistance??0),2),prl(0xb018,twip(extra?.footerDistance??0),2),prl(0x900c,twip(extra?.columnSpace??0),2),prl(0x301d,p.width>p.height?2:1,1));
  const word=new Uint8Array(Math.max(4096,sepOffset+2+sectionPr.length));word.set(text,TEXT_START);
  for(let i=0;i<ch.pages.length;i++)word.set(ch.pages[i]!,512*(firstPage+i));
  for(let i=0;i<pa.pages.length;i++)word.set(pa.pages[i]!,512*(firstPage+ch.pages.length+i));
  word.set(w16(sectionPr.length),sepOffset);word.set(sectionPr,sepOffset+2);
  // FibBase + FibRgW97 + FibRgLw97 + FibRgFcLcb97. The reserved old
  // fcMin/fcMac locations are also populated for independent old readers.
  put16(word,0,0xa5ec);put16(word,2,0xc1);put16(word,6,0x409);put16(word,10,0x1200);put16(word,12,0xbf);
  put32(word,24,TEXT_START);put32(word,28,endText);put16(word,32,14);put16(word,62,22);put32(word,64,word.length);put32(word,76,cp);put16(word,152,PAIR_COUNT);
  put16(word,PAIR_START+PAIR_COUNT*8,0);
  const pieces=cat(Uint8Array.of(2),w32(16),w32(0),w32(cp),w16(0),w32(TEXT_START),w16(0));
  const sed=cat(w32(0),w32(cp),w16(0),w32(sepOffset),w16(0),w32(0xffffffff));
  const parts:Uint8Array[]=[new Uint8Array(4)];let offset=4;
  const add=(index:number,data:Uint8Array)=>{put32(word,PAIR_START+index*8,offset);put32(word,PAIR_START+index*8+4,data.length);parts.push(data);offset+=data.length;};
  add(PAIRS.styles,styleSheet());add(PAIRS.sections,sed);add(PAIRS.chpx,ch.plc);add(PAIRS.papx,pa.plc);add(PAIRS.fonts,fontTable(names));add(PAIRS.dop,dop());add(PAIRS.assoc,association());add(PAIRS.clx,pieces);
  return {word,table:cat(...parts)};
}
function readFonts(data:Uint8Array):string[]{
  const count=u16(data,0);valid(count>=1&&count<=256&&u16(data,2)===0,'Font table outside profile');let at=4;const names:string[]=[];
  for(let i=0;i<count;i++){
    bounds(data,at,1);const size=data[at]!+1;bounds(data,at,size);valid(size>=42&&size%2===0,'Invalid font record');
    let name='';for(let p=at+40;p<at+size;p+=2){const c=u16(data,p);if(!c){valid(p===at+size-2,'Alternate font name unsupported');break;}name+=String.fromCharCode(c);}
    valid(name.length>0&&name.length<=64,'Invalid font name');names.push(name);at+=size;
  }valid(at===data.length,'Unexpected font table tail');return names;
}
function charStyle(data:Uint8Array,fonts:readonly string[]):TextStyle{
  const s:TextStyle={};let at=0;const seen=new Set<number>(),faces:number[]=[];
  while(at<data.length){
    const op=u16(data,at);at+=2;valid(!seen.has(op),'Duplicate CHPX property');seen.add(op);
    if([0x4a4f,0x4a50,0x4a51].includes(op)){faces.push(u16(data,at));at+=2;}
    else if(op===0x4a43){s.fontSize=u16(data,at)*6350;at+=2;}
    else if([0x0835,0x0836,0x2a3e].includes(op)){bounds(data,at,1);const v=data[at++]!;valid(v<=1,'Unsupported CHPX boolean/underline');s[op===0x0835?'bold':op===0x0836?'italic':'underline']=v===1;}
    else if(op===0x6870){const color=u32(data,at);at+=4;valid(color<=0xffffff,'Unsupported DOC automatic or theme color');s.color={kind:'rgba',red:color&255,green:color>>>8&255,blue:color>>>16&255,alpha:255};}
    else if(op===0x486d||op===0x486e){const lid=u16(data,at);at+=2;const tag=Object.entries(LIDS).find(([,v])=>v===lid)?.[0];valid(tag!==undefined,'Unsupported DOC LID');s[op===0x486d?'language':'eastAsiaLanguage']=tag;}
    else valid(false,'Unknown CHPX SPRM '+op.toString(16));
  }
  valid(faces.length===3&&faces.every(n=>n===faces[0])&&fonts[faces[0]!]&&s.fontSize!==undefined,'Unsupported font mapping');s.fontFamily=fonts[faces[0]!]!;return s;
}
function paragraphStyle(data:Uint8Array):TextStyle{
  valid(u16(data,0)===0,'Only Normal paragraph style supported');const s:TextStyle={};let at=2;const seen=new Set<number>();
  while(at<data.length){const op=u16(data,at);at+=2;valid(!seen.has(op),'Duplicate PAPX property');seen.add(op);
    if(op===0x2403){bounds(data,at,1);const n=data[at++]!;valid(n<=3,'Unsupported paragraph justification');s.alignment=(['start','center','end','justify'] as const)[n];}
    else if([0x840f,0xa413,0xa414].includes(op)){const n=op===0x840f?i16(data,at):u16(data,at);at+=2;s[op===0x840f?'indent':op===0xa413?'spaceBefore':'spaceAfter']=n*EMU_PER_TWIP;}
    else if(op===0x6412){const n=i16(data,at),multiple=u16(data,at+2);at+=4;valid(n!==0&&(multiple===0||multiple===1&&n>0),'Unsupported line spacing');if(multiple)s.lineMultiple=n;else if(n<0)s.lineHeight=-n*EMU_PER_TWIP;else s.minLineHeight=n*EMU_PER_TWIP;}
    else valid(false,'Unknown PAPX SPRM '+op.toString(16));
  }return s;
}
function readFkps(f:Fib,index:number,paragraph:boolean):Span[]{
  const data=part(f,index);valid(data.length>=12&&(data.length-4)%8===0,'Invalid FKP PLC size');const n=(data.length-4)/8,spans:Span[]=[];
  valid(n<=10000,'FKP page count limit');
  for(let i=0;i<n;i++){
    const from=u32(data,i*4),to=u32(data,(i+1)*4),pn=u32(data,(n+1)*4+i*4),offset=pn*512;
    bounds(f.word,offset,512);const b=f.word.subarray(offset,offset+512),count=b[511]!,step=paragraph?13:1,head=4*(count+1)+count*step;
    valid(count>0&&head<=510&&u32(b,0)===from&&u32(b,count*4)===to,'Invalid FKP range');
    for(let j=0;j<count;j++){
      const start=u32(b,j*4),end=u32(b,(j+1)*4),pointer=b[4*(count+1)+j*step]!*2;
      valid(start<end&&pointer>=head&&pointer<511,'Invalid FKP record pointer');
      let len=b[pointer]!,at=pointer+1;
      if(paragraph){if(len===0){len=b[at++]!*2;}else len=len*2-1;}
      valid(at+len<=511,'FKP property exceeds page');spans.push({start,end,data:b.slice(at,at+len)});
    }
  }
  valid(spans.every((s,i)=>!i||spans[i-1]!.end===s.start),'Noncontiguous FKP coverage');return spans;
}
function readSection(f:Fib):{page:WriterDocument['sections'][number]['page'];pageExtras?:WriterDocument['sections'][number]['pageExtras']}{
  const sed=part(f,PAIRS.sections);valid(sed.length===20&&u32(sed,0)===0&&u32(sed,4)===f.cp,'Only single DOC section is supported');
  const offset=u32(sed,10),length=u16(f.word,offset);bounds(f.word,offset+2,length);const data=f.word.subarray(offset+2,offset+2+length),values=new Map<number,number>();let at=0;
  const supported=[0xb01f,0xb020,0xb021,0xb022,0x9023,0x9024,0xb025,0xb017,0xb018,0x900c,0x301d];
  while(at<data.length){const op=u16(data,at);at+=2;valid(supported.includes(op)&&!values.has(op),'Unsupported section SPRM');
    if(op===0x301d){bounds(data,at,1);values.set(op,data[at++]!);}else{values.set(op,u16(data,at));at+=2;}}
  valid(supported.every(op=>values.has(op))&&values.get(0xb025)===0,'Missing section geometry or nonzero gutter');
  const get=(op:number)=>values.get(op)!*EMU_PER_TWIP;
  const page={width:get(0xb01f),height:get(0xb020),margins:{left:get(0xb021),right:get(0xb022),top:get(0x9023),bottom:get(0x9024)}};
  const extra={headerDistance:get(0xb017),footerDistance:get(0xb018),columnSpace:get(0x900c)};
  return {page,...(Object.values(extra).some(n=>n!==0)?{pageExtras:extra}:{})};
}
export interface LegacyDocSource {readonly document:DeepReadonly<WriterDocument>;readonly version:number}
const originals=new WeakMap<LegacyDocSource,{bytes:Uint8Array;word:Uint8Array;table:Uint8Array;cfb:ParsedCfbDocument}>();
export function openLegacyDoc(bytes:Uint8Array,documentId:string,title:string,options:CfbOptions={}):LegacyDocSource{
  const f=fib(bytes,options),plc=pieces(f);valid(f.version===0xc1&&f.otherStories===0&&f.flags===0x1200&&f.cfb.entries.length===2&&f.tableName==='1Table','Source outside closed editable DOC envelope');
  valid(plc.length===1&&!plc[0]!.compressed&&plc[0]!.fc===TEXT_START&&plc[0]!.prm===0&&plc[0]!.end===f.cp,'Unsupported editable DOC piece layout');
  const text=mainText(f,plc),fonts=readFonts(part(f,PAIRS.fonts)),ch=readFkps(f,PAIRS.chpx,false),pa=readFkps(f,PAIRS.papx,true);
  valid(ch[0]?.start===TEXT_START&&ch.at(-1)?.end===TEXT_START+f.cp*2&&pa[0]?.start===TEXT_START&&pa.at(-1)?.end===TEXT_START+f.cp*2,'Incomplete FKP coverage');
  const blocks:WriterBlock[]=[];let charIndex=0,cp=0;
  for(let i=0;i<pa.length;i++){
    const span=pa[i]!,start=(span.start-TEXT_START)/2,end=(span.end-TEXT_START)/2;
    valid(start===cp&&Number.isInteger(end)&&end>start&&text[end-1]==='\r'&&!text.slice(start,end-1).includes('\r'),'PAPX does not match paragraph boundaries');
    const direct=paragraphStyle(span.data),runs:TextRun[]=[];
    for(let at=start;at<end;){
      while(charIndex<ch.length&&ch[charIndex]!.end<=TEXT_START+at*2)charIndex++;
      valid(charIndex<ch.length,'CHPX coverage ended early');
      const c=ch[charIndex]!,next=Math.min(end,(c.end-TEXT_START)/2);
      valid(Number.isInteger(next)&&next>at,'Invalid CHPX CP boundary');
      const value=text.slice(at,Math.min(next,end-1)),style=charStyle(c.data,fonts);
      if(value||!runs.length)runs.push({id:'r-'+i+'-'+runs.length,type:'text',text:value.replace(/\v/g,'\n'),style});
      at=next;
    }
    if(text.slice(start,end-1)==='\f'){valid(!Object.keys(direct).length,'Styled page break unsupported');blocks.push({id:'break-'+i,type:'page-break'});}
    else{valid(!/[\x00-\x08\x0c-\x1f\x7f\uFFFC]/.test(runs.map(r=>r.text).join('')),'Unsupported DOC special character');blocks.push({id:'p-'+i,type:'paragraph',runs,...(Object.keys(direct).length?{style:direct}:{})});}
    cp=end;
  }
  const model=createDocument('writer',documentId,title);model.defaultTextStyle={};model.sections=[{id:'section-1',...readSection(f),blocks}];
  validateDocument(model);
  // Strong closed-profile check: no unknown FC/LCB, stream bytes or hidden
  // metadata may disappear during an edited save.
  const reconstructed=writeStreams(model);valid(equal(reconstructed.word,f.word)&&equal(reconstructed.table,f.table),'DOC source contains unmodeled binary structures');
  const source=Object.freeze({document:freezeData(model),version:f.version});originals.set(source,{bytes:bytes.slice(),word:f.word.slice(),table:f.table.slice(),cfb:f.cfb});return source;
}
export function writeLegacyDoc(candidate:unknown,options:CfbOptions={}):Uint8Array{
  const model=validateDocument(candidate);valid(model.kind==='writer','Expected Writer model');const streams=writeStreams(model);
  const result=writeCfb({majorVersion:3,entries:[{path:['WordDocument'],type:'stream',data:streams.word},{path:['1Table'],type:'stream',data:streams.table}]},options);
  openLegacyDoc(result,model.documentId,model.title,options);return result;
}
export function saveLegacyDoc(source:LegacyDocSource,candidate:unknown,options:CfbOptions={}):Uint8Array{
  const original=originals.get(source);assert(original,'invalid','Unknown DOC source capability');const model=validateDocument(candidate);valid(model.kind==='writer','Expected Writer model');
  assert(model.documentId===source.document.documentId,'conflict','DOC source identity mismatch');
  const streams=writeStreams(model);if(equal(streams.word,original.word)&&equal(streams.table,original.table))return original.bytes.slice();
  const bytes=writeCfb({...original.cfb,entries:original.cfb.entries.map(entry=>entry.type==='stream'?{...entry,data:entry.path[0]!.toUpperCase()==='WORDDOCUMENT'?streams.word:streams.table}:entry)},options);
  openLegacyDoc(bytes,model.documentId,model.title,options);return bytes;
}

