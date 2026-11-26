// SPDX-License-Identifier: MIT
/** Bounded read-only text outline for unsupported modern documents.
 * No active document, layout engine, relationship traversal to media, link
 * activation, field execution, revision reconstruction or semantic save. */
import type {Element} from '@xmldom/xmldom';
import {OfficeError} from '@office-ex/contracts';
import {openOpc} from './opc.js';
import {parseXml,childElements} from './xml.js';
import {checkCancelled,type ContainerOptions} from './limits.js';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const LIMIT=10000,TEXT_LIMIT=2*1024*1024;
export type OfficeTextBlock={kind:'paragraph';text:string}|{kind:'table';rows:string[][]};
export interface OfficeTextPreview {kind:'word'|'slides';pages:{title?:string;blocks:OfficeTextBlock[]}[];truncated:boolean}
function valid(c:unknown,m:string):asserts c{if(!c)throw new OfficeError('unsupported','Office text preview: '+m);}
const is=(e:Element,ns:string,name:string)=>e.namespaceURI===ns&&e.localName===name;
const select=(e:Element,ns:string,name:string)=>childElements(e).filter(c=>is(c,ns,name));
function one(e:Element,ns:string,name:string):Element|undefined{const list=select(e,ns,name);valid(list.length<=1,'Duplicate '+name);return list[0];}
function safeText(s:string):string{return s.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g,'�');}
function xmlRoot(bytes:Uint8Array|undefined,ns:string,name:string,options:ContainerOptions):Element{valid(bytes,'Missing text part');const e=parseXml(bytes,options).documentElement;valid(e&&is(e,ns,name),'Unexpected '+name+' root');return e;}
export function previewOfficeText(bytes:Uint8Array,options:ContainerOptions={}):OfficeTextPreview{
 const pkg=openOpc(bytes,options);
 valid(pkg.mainPart&&!pkg.hasMacros&&pkg.conformance==='transitional'&&(pkg.kind==='docx'||pkg.kind==='pptx'),'Unsupported text preview package');
 const preview:OfficeTextPreview={kind:pkg.kind==='docx'?'word':'slides',pages:[],truncated:false};
 let blocks=0,textLength=0;
 function admit(list:OfficeTextBlock[],block:OfficeTextBlock):boolean{
  const length=block.kind==='paragraph'?block.text.length:block.rows.reduce((a,r)=>a+r.reduce((b,s)=>b+s.length,0),0);
  const charge=block.kind==='paragraph'?1:1+block.rows.reduce((a,r)=>a+r.length,0);
  if(blocks+charge>LIMIT||textLength+length>TEXT_LIMIT){preview.truncated=true;return false;}
  blocks+=charge;textLength+=length;list.push(block);return true;
 }
 function wordParagraph(p:Element):string{
  const pieces:string[]=[];
  function walk(e:Element):void{
   if(e.namespaceURI!==W)return;
   const n=e.localName;
   if(n==='r'){
    const pr=one(e,W,'rPr');
    if(pr&&['vanish','webHidden'].some(k=>select(pr,W,k).some(x=>!['0','false','off'].includes(x.getAttributeNS(W,'val')??'1'))))return;
    for(const r of childElements(e)){
     if(is(r,W,'t'))pieces.push(r.textContent??'');
     else if(is(r,W,'tab'))pieces.push('\t');
     else if(is(r,W,'br')||is(r,W,'cr'))pieces.push('\n');
     // Ignore field instructions, drawing, objects, deleted/hidden runs.
    }
   }else if(['p','hyperlink','ins','smartTag','sdt','sdtContent','fldSimple','moveTo'].includes(n??''))for(const child of childElements(e))walk(child);
  }
  walk(p);return safeText(pieces.join(''));
 }
 function wordTable(table:Element):string[][]{
  const rows:string[][]=[];
  for(const tr of select(table,W,'tr').slice(0,500)){
   const cells=select(tr,W,'tc');if(cells.length>64)preview.truncated=true;
   rows.push(cells.slice(0,64).map(tc=>select(tc,W,'p').map(wordParagraph).join('\n')));
  }
  if(select(table,W,'tr').length>500)preview.truncated=true;
  return rows;
 }
 function wordBlocks(e:Element,list:OfficeTextBlock[]):void{
  for(const child of childElements(e)){
   checkCancelled(options);
   if(is(child,W,'p')){if(!admit(list,{kind:'paragraph',text:wordParagraph(child)}))return;}
   else if(is(child,W,'tbl')){if(!admit(list,{kind:'table',rows:wordTable(child)}))return;}
   else if(child.namespaceURI===W&&['sdt','sdtContent','ins','moveTo','customXml'].includes(child.localName??''))wordBlocks(child,list);
   if(blocks>=LIMIT||textLength>=TEXT_LIMIT){preview.truncated=true;return;}
  }
 }
 if(pkg.kind==='docx'){
  const main=xmlRoot(pkg.parts.get(pkg.mainPart),W,'document',options),body=one(main,W,'body');
  valid(body,'Missing document body');const list:OfficeTextBlock[]=[];wordBlocks(body,list);
  preview.pages.push({blocks:list});
 }else{
  const main=xmlRoot(pkg.parts.get(pkg.mainPart),P,'presentation',options),ids=one(main,P,'sldIdLst');valid(ids,'Missing slide list');
  const rels=pkg.relationships.get(pkg.mainPart)??[],used=new Set<string>();
  const paragraph=(p:Element):string=>safeText(childElements(p).flatMap(c=>is(c,A,'r')||is(c,A,'fld')?select(c,A,'t').map(t=>t.textContent??''):is(c,A,'br')?['\n']:[]).join(''));
  const bodyText=(e:Element|undefined)=>e?select(e,A,'p').map(paragraph).join('\n'):'';
  const slideIds=select(ids,P,'sldId');if(slideIds.length>500)preview.truncated=true;
  for(const s of slideIds.slice(0,500)){
   checkCancelled(options);
   const id=s.getAttributeNS(R,'id'),matches=rels.filter(r=>r.id===id&&r.type===R+'/slide');
   valid(matches.length===1,'Invalid slide relationship');
   const r=matches[0]!;valid(r.mode==='Internal'&&r.targetPart&&!r.fragment&&!used.has(r.targetPart)&&pkg.contentTypes.get(r.targetPart)==='application/vnd.openxmlformats-officedocument.presentationml.slide+xml','Invalid slide part');
   used.add(r.targetPart);const root=xmlRoot(pkg.parts.get(r.targetPart),P,'sld',options),cs=one(root,P,'cSld'),tree=cs&&one(cs,P,'spTree');valid(tree,'Missing slide content');
   const list:OfficeTextBlock[]=[];let title:string|undefined;
   function shapes(node:Element):void{
    for(const shape of childElements(node)){
     checkCancelled(options);
     if(is(shape,P,'grpSp'))shapes(shape);
     else if(is(shape,P,'sp')){
      const tx=bodyText(one(shape,P,'txBody'));
      if(!tx)continue;
      const nv=one(shape,P,'nvSpPr'),pr=nv&&one(nv,P,'nvPr'),ph=pr&&one(pr,P,'ph');
      if(ph&&['title','ctrTitle'].includes(ph.getAttribute('type')??'')&&!title)title=tx.slice(0,256);
      if(!admit(list,{kind:'paragraph',text:tx}))return;
     }else if(is(shape,P,'graphicFrame')){
      const graphic=one(shape,A,'graphic'),data=graphic&&one(graphic,A,'graphicData'),table=data&&one(data,A,'tbl');
      if(table){const trs=select(table,A,'tr');if(trs.length>500)preview.truncated=true;
       const rows=trs.slice(0,500).map(tr=>{const cells=select(tr,A,'tc');if(cells.length>64)preview.truncated=true;return cells.slice(0,64).map(tc=>bodyText(one(tc,A,'txBody')));});
       if(!admit(list,{kind:'table',rows}))return;
      }
     }
     if(blocks>=LIMIT||textLength>=TEXT_LIMIT){preview.truncated=true;return;}
    }
   }
   shapes(tree);preview.pages.push({...(title?{title}:{}),blocks:list});
   if(blocks>=LIMIT||textLength>=TEXT_LIMIT){preview.truncated=true;break;}
  }
 }
 valid(preview.pages.length&&preview.pages.some(p=>p.blocks.length),'No supported text content');
 return preview;
}
