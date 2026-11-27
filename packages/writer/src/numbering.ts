// SPDX-License-Identifier: MIT
import {assert,EMU_PER_POINT} from '@office-ex/contracts';
import {cloneData,patchCommand,writerEffectiveStyle,type DocumentSnapshot,type DocumentCommand,type PatchOperation,type WriterNumbering,type WriterListLevel,type WriterListReference,type WriterDocument,type DeepReadonly} from '@office-ex/core';
import {locateParagraph} from './paragraphs.js';
type ListKind='none'|'bullet'|'decimal';
const doc=(snapshot:DocumentSnapshot):DeepReadonly<WriterDocument>=>{assert(snapshot.document.kind==='writer','invalid','Expected Writer document');return snapshot.document;};
function definitions(kind:'bullet'|'decimal'):WriterListLevel[]{
 return Array.from({length:9},(_,i)=>({format:kind,start:1,pattern:kind==='bullet'?(i%3===0?'•':i%3===1?'○':'▪'):Array.from({length:i+1},(_,j)=>'%'+(j+1)).join('.')+'.',indent:(36+36*i)*EMU_PER_POINT,hanging:(36+18*i)*EMU_PER_POINT,alignment:'start'}));
}
function allocate(registry:WriterNumbering,levels:WriterListLevel[]):number{
 assert(registry.nextId<=1000000000&&Object.keys(registry.lists).length<2048,'limit','Numbering registry budget exhausted');
 const id=registry.nextId++;registry.lists[String(id)]={levels};return id;
}
export function setWriterList(snapshot:DocumentSnapshot,paragraphId:string,kind:ListKind,level=0):DocumentCommand{
 assert(['none','bullet','decimal'].includes(kind)&&Number.isInteger(level)&&level>=0&&level<=8,'invalid','Invalid list format or level');
 const model=doc(snapshot),loc=locateParagraph(snapshot,paragraphId),p=loc.paragraph,style=writerEffectiveStyle(model,p),old=style.writerList;
 const operations:PatchOperation[]=[];
 let ref:WriterListReference={numId:0,level:0};
 if(kind!=='none'){
  assert(style.paragraphDirection!=='rtl','unsupported','RTL numbering outside supported profile');
  const registry=cloneData(model.numbering??{nextId:1,lists:{}}) as WriterNumbering;
  const existing=old?.numId?registry.lists[String(old.numId)]:undefined;
  if(existing&&existing.levels[level]?.format===kind)ref={numId:old!.numId,level};
  else{
   const previous=loc.siblings[loc.block-1],previousRef=previous?.type==='paragraph'?writerEffectiveStyle(model,previous).writerList:undefined;
   if(previousRef?.numId&&registry.lists[String(previousRef.numId)]?.levels[level]?.format===kind)ref={numId:previousRef.numId,level};
   else{ref={numId:allocate(registry,definitions(kind)),level};operations.push({op:'set',path:['numbering'],value:registry});}
  }
 }
 const next={...p.style,writerList:ref};
 if(ref.numId){delete next.indent;delete next.firstLineIndent;} // use list defaults
 operations.push({op:'set',path:[...loc.parent,loc.block,'style'],value:next});
 return patchCommand(snapshot,'WRT-03',operations);
}
/** Restart the contiguous suffix of this instance as a new independent list. */
export function restartWriterList(snapshot:DocumentSnapshot,paragraphId:string,start=1):DocumentCommand{
 assert(Number.isInteger(start)&&start>=1&&start<=32767,'invalid','Invalid list starting value');
 const model=doc(snapshot),loc=locateParagraph(snapshot,paragraphId),ref=writerEffectiveStyle(model,loc.paragraph).writerList;
 assert(ref?.numId&&model.numbering,'invalid','Select a numbered paragraph');
 const registry=cloneData(model.numbering) as WriterNumbering,levels=cloneData(registry.lists[String(ref.numId)]!.levels);
 assert(levels[ref.level]!.format==='decimal','invalid','Only decimal numbering can restart');
 levels[ref.level]!.start=start;const id=allocate(registry,levels),ops:PatchOperation[]=[{op:'set',path:['numbering'],value:registry}];
 for(let i=loc.block;i<loc.siblings.length;i++){const p=loc.siblings[i];if(p?.type!=='paragraph')break;const current=writerEffectiveStyle(model,p).writerList;if(current?.numId!==ref.numId)break;ops.push({op:'set',path:[...loc.parent,i,'style'],value:{...p.style,writerList:{numId:id,level:current.level}}});}
 return patchCommand(snapshot,'WRT-03',ops);
}
