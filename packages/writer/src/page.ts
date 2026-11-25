// SPDX-License-Identifier: MIT
import {assert,assertId,EMU_PER_TWIP} from '@office-ex/contracts';
import {patchCommand,assertPlainData,cloneData,type Paragraph,type WriterSection as MutableSection,type DocumentSnapshot,type DocumentCommand,type WriterSection} from '@office-ex/core';
export function setWriterPage(snapshot:DocumentSnapshot,sectionId:string,page:WriterSection['page']):DocumentCommand{
  assert(snapshot.document.kind==='writer','invalid','Expected Writer document');
  assertPlainData(page);assert(Object.keys(page).length===3&&Object.keys(page.margins).length===4,'invalid','Invalid page fields');
  const values=[page.width,page.height,page.margins.top,page.margins.right,page.margins.bottom,page.margins.left];
  assert(values.every(v=>Number.isSafeInteger(v)&&v>=0&&v%EMU_PER_TWIP===0&&v<=31680*EMU_PER_TWIP),'invalid','Page dimensions require whole twips');
  const width=page.width-page.margins.left-page.margins.right,height=page.height-page.margins.top-page.margins.bottom;
  assert(width>0&&height>0,'invalid','Margins leave no content area');
  const index=snapshot.document.sections.findIndex(s=>s.id===sectionId);assert(index>=0,'invalid','Section not found');
  for(const block of snapshot.document.sections[index]!.blocks){
    if(block.type==='table'){
      assert(block.columnWidths.reduce((a,b)=>a+b,0)<=width,'unsupported','Resize the table before reducing page width');
      const pad=block.cellPadding??{top:38100,bottom:38100,left:57150,right:57150};
      for(const row of block.rows)for(const cell of row.cells)for(const p of cell.blocks)if(p.type==='paragraph')for(const run of p.runs)if(run.type==='image')assert(run.height+pad.top+pad.bottom<=height,'unsupported','Resize the cell image before reducing page height');
    }
    if(block.type==='paragraph')for(const run of block.runs)if(run.type==='image')assert(run.width<=width&&run.height<=height,'unsupported','Resize the image before reducing page size');
  }
  return patchCommand(snapshot,'WRT-02',[{op:'set',path:['sections',index,'page'],value:page}]);
}

/** All supported section boundaries start a new page. A nonfinal Word section
 * ends in a paragraph, never a table or an explicit hard page break. */
export function splitWriterSectionAfter(snapshot:DocumentSnapshot,paragraphId:string,operationId:string):{command:DocumentCommand;sectionId:string;paragraphId:string}{
  assert(snapshot.document.kind==='writer','invalid','Expected Writer document');assertId(operationId);assert(operationId.length<=100,'invalid','Section operation ID too long');
  const document=snapshot.document;assert(document.sections.length<128,'limit','Section budget exceeded');
  for(let index=0;index<document.sections.length;index++){
    const original=document.sections[index]!,block=original.blocks.findIndex(b=>b.id===paragraphId);
    if(block<0)continue;assert(original.blocks[block]!.type==='paragraph','invalid','Section break needs a top-level paragraph');
    const first=cloneData(original) as MutableSection,second=cloneData(original) as MutableSection;
    first.blocks=first.blocks.slice(0,block+1);second.id=operationId+'-section';second.blocks=second.blocks.slice(block+1);
    const empty:Paragraph={id:operationId+'-p',type:'paragraph',...(document.defaultParagraphStyleId?{styleId:document.defaultParagraphStyleId}:{}),runs:[{id:operationId+'-r',type:'text',text:''}]};
    if(!second.blocks.length)second.blocks.push(empty);
    // Focus the first editable paragraph in the following section. Do not
    // introduce an extra paragraph before an existing table or image.
    const focus=second.blocks.find(b=>b.type==='paragraph')?.id??'';
    return {sectionId:second.id,paragraphId:focus,command:patchCommand(snapshot,'WRT-03',[{op:'splice',path:['sections'],index,deleteCount:1,items:[first,second]}])};
  }
  assert(false,'invalid','Section break paragraph not found');
}
