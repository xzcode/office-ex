// SPDX-License-Identifier: MIT
import {assert,assertId,EMU_PER_CSS_PIXEL} from '@office-ex/contracts';
import {patchCommand,makeImageResource,validateImageResource,cloneData,type DocumentSnapshot,type DocumentCommand,type DeepReadonly,type WriterDocument,type ImageResource,type Paragraph,type Inline,type ModelPath} from '@office-ex/core';
type ImageInline=Extract<Inline,{type:'image'}>;
function writer(s:DocumentSnapshot):DeepReadonly<WriterDocument>{assert(s.document.kind==='writer','invalid','Expected Writer document');return s.document;}
function locate(s:DocumentSnapshot,id:string){
  for(let section=0;section<writer(s).sections.length;section++){
    const sec=writer(s).sections[section]!;
    for(let block=0;block<sec.blocks.length;block++){
      const p=sec.blocks[block]!;
      if(p.type==='paragraph'&&p.runs.length===1&&p.runs[0]?.type==='image'&&p.runs[0].id===id)return {sec,p,image:p.runs[0],path:['sections',section,'blocks',block] as ModelPath,section,block};
    }
  }
  assert(false,'invalid','Image not found in a dedicated inline paragraph');
}
function geometry(sec:DeepReadonly<WriterDocument['sections'][number]>,width:number,height:number):void{
  assert(Number.isSafeInteger(width)&&Number.isSafeInteger(height)&&width>0&&height>0&&width<=sec.page.width-sec.page.margins.left-sec.page.margins.right&&height<=sec.page.height-sec.page.margins.top-sec.page.margins.bottom,'invalid','Image geometry exceeds usable page');
}
export function insertImageAfter(s:DocumentSnapshot,paragraphId:string,resource:ImageResource,operationId:string):DocumentCommand{
  validateImageResource(resource);assertId(operationId);assert(operationId.length<=100,'invalid','Image operation ID too long');
  const d=writer(s);
  for(let section=0;section<d.sections.length;section++){
    const sec=d.sections[section]!,block=sec.blocks.findIndex(p=>p.id===paragraphId);
    if(block<0)continue;
    assert(sec.blocks[block]?.type==='paragraph','invalid','Image insertion needs a top-level paragraph');
    const scale=Math.min(1,(sec.page.width-sec.page.margins.left-sec.page.margins.right)/(resource.pixelWidth*EMU_PER_CSS_PIXEL),(sec.page.height-sec.page.margins.top-sec.page.margins.bottom)/(resource.pixelHeight*EMU_PER_CSS_PIXEL));
    const width=Math.max(1,Math.floor(resource.pixelWidth*EMU_PER_CSS_PIXEL*scale)),height=Math.max(1,Math.floor(resource.pixelHeight*EMU_PER_CSS_PIXEL*scale));geometry(sec,width,height);
    const p:Paragraph={id:operationId+'-p',type:'paragraph',...(d.defaultParagraphStyleId?{styleId:d.defaultParagraphStyleId}:{}),runs:[{id:operationId+'-image',type:'image',resourceId:resource.id,width,height}]};
    const existing=d.resources[resource.id];if(existing)validateImageResource(existing);
    return patchCommand(s,'WRT-03',[...(!existing?[{op:'set' as const,path:['resources',resource.id],value:cloneData(resource)}]:[]),{op:'splice',path:['sections',section,'blocks'],index:block+1,deleteCount:0,items:[p]}]);
  }
  assert(false,'invalid','Image insertion paragraph not found');
}
export function resizeInlineImage(s:DocumentSnapshot,id:string,width:number):DocumentCommand{
  const loc=locate(s,id),r=writer(s).resources[loc.image.resourceId];validateImageResource(r);
  const height=Math.max(1,Math.round(width*r.pixelHeight/r.pixelWidth));geometry(loc.sec,width,height);
  return patchCommand(s,'WRT-03',[{op:'set',path:[...loc.path,'runs',0],value:{...loc.image,width,height}}]);
}
export function removeInlineImage(s:DocumentSnapshot,id:string):DocumentCommand{
  const loc=locate(s,id);
  // Retain image payloads as part of the undo/source closure, including unused
  // originals. No automatic resource GC is performed in this profile.
  if(loc.sec.blocks.length===1)return patchCommand(s,'WRT-03',[{op:'set',path:loc.path,value:{...loc.p,runs:[{id:loc.image.id,type:'text',text:''}]}}]);
  return patchCommand(s,'WRT-03',[{op:'splice',path:['sections',loc.section,'blocks'],index:loc.block,deleteCount:1,items:[]}]);
}
