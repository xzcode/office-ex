// SPDX-License-Identifier: MIT
import {assert,assertId,EMU_PER_CSS_PIXEL} from '@office-ex/contracts';
import {patchCommand,validateImageResource,cloneData,assertParagraphRange,type DocumentSnapshot,type DocumentCommand,type DeepReadonly,type WriterDocument,type ImageResource,type Paragraph,type Inline,type PatchOperation} from '@office-ex/core';
import {locateParagraph,sliceWriterRuns,writerIdAllocator} from './paragraphs.js';
type ImageInline=Extract<Inline,{type:'image'}>;
function writer(s:DocumentSnapshot):DeepReadonly<WriterDocument>{assert(s.document.kind==='writer','invalid','Expected Writer document');return s.document;}
function locate(s:DocumentSnapshot,id:string){
  for(const section of writer(s).sections)for(const block of section.blocks){
    const paragraphs=block.type==='paragraph'?[block]:block.type==='table'?block.rows.flatMap(r=>r.cells.flatMap(c=>c.blocks.filter(b=>b.type==='paragraph'))):[];
    for(const p of paragraphs){
      if(p.type!=='paragraph')continue;
      const run=p.runs.findIndex(r=>r.id===id),image=p.runs[run];
      if(image?.type==='image')return {...locateParagraph(s,p.id),run,image};
    }
  }
  assert(false,'invalid','Inline image not found');
}
function geometry(maxWidth:number,maxHeight:number,width:number,height:number):void{
  assert(Number.isSafeInteger(width)&&Number.isSafeInteger(height)&&width>0&&height>0&&width<=maxWidth&&height<=maxHeight,'invalid','Image geometry exceeds usable paragraph area');
}
function prepare(s:DocumentSnapshot,resource:ImageResource,operationId:string,maxWidth:number,maxHeight:number){
  validateImageResource(resource);assertId(operationId);assert(operationId.length<=100,'invalid','Image operation ID too long');
  const scale=Math.min(1,maxWidth/(resource.pixelWidth*EMU_PER_CSS_PIXEL),maxHeight/(resource.pixelHeight*EMU_PER_CSS_PIXEL));
  const width=Math.max(1,Math.floor(resource.pixelWidth*EMU_PER_CSS_PIXEL*scale)),height=Math.max(1,Math.floor(resource.pixelHeight*EMU_PER_CSS_PIXEL*scale));
  geometry(maxWidth,maxHeight,width,height);
  const image:ImageInline={id:operationId+'-image',type:'image',resourceId:resource.id,width,height};
  const existing=writer(s).resources[resource.id];if(existing)validateImageResource(existing);
  const operations:PatchOperation[]=existing?[]:[{op:'set',path:['resources',resource.id],value:cloneData(resource)}];
  return {image,operations};
}
/** Existing image-paragraph command; intentionally separate from inline insert. */
export function insertImageAfter(s:DocumentSnapshot,paragraphId:string,resource:ImageResource,operationId:string):DocumentCommand{
  const p=locateParagraph(s,paragraphId),d=writer(s);
  const {image,operations}=prepare(s,resource,operationId,p.contentWidth,p.contentHeight);
  const block:Paragraph={id:operationId+'-p',type:'paragraph',...(d.defaultParagraphStyleId?{styleId:d.defaultParagraphStyleId}:{}),runs:[image]};
  return patchCommand(s,'WRT-03',[...operations,{op:'splice',path:p.parent,index:p.block+1,deleteCount:0,items:[block]}]);
}
/** Insert an atomic image at a caret or replace an explicitly selected range. */
export function insertImageAt(s:DocumentSnapshot,paragraphId:string,start:number,end:number,resource:ImageResource,operationId:string):DocumentCommand{
  const p=locateParagraph(s,paragraphId);assertParagraphRange(p.paragraph,start,end);
  const {image,operations}=prepare(s,resource,operationId,p.contentWidth,p.contentHeight);
  const allocate=writerIdAllocator(operationId+'-r'),used=new Set<string>();
  const left=sliceWriterRuns(p.runs,0,start,used,allocate),right=sliceWriterRuns(p.runs,end,p.text.length,used,allocate);
  return patchCommand(s,'WRT-03',[...operations,{op:'set',path:[...p.parent,p.block,'runs'],value:[...left,image,...right]}]);
}
export function resizeInlineImage(s:DocumentSnapshot,id:string,width:number):DocumentCommand{
  const loc=locate(s,id),r=writer(s).resources[loc.image.resourceId];validateImageResource(r);
  const height=Math.max(1,Math.round(width*r.pixelHeight/r.pixelWidth));geometry(loc.contentWidth,loc.contentHeight,width,height);
  return patchCommand(s,'WRT-03',[{op:'set',path:[...loc.parent,loc.block,'runs',loc.run],value:{...loc.image,width,height}}]);
}
export function removeInlineImage(s:DocumentSnapshot,id:string):DocumentCommand{
  const loc=locate(s,id);
  // Keep original resource bytes. Undo may still refer to them.
  if(loc.runs.length===1&&loc.siblings.length>1&&!loc.inCell)
    return patchCommand(s,'WRT-03',[{op:'splice',path:loc.parent,index:loc.block,deleteCount:1,items:[]}]);
  const runs=loc.runs.filter(r=>r.id!==id).map(r=>cloneData(r));
  if(!runs.length)runs.push({id:loc.image.id,type:'text',text:''});
  return patchCommand(s,'WRT-03',[{op:'set',path:[...loc.parent,loc.block,'runs'],value:runs}]);
}
