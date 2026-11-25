// SPDX-License-Identifier: MIT
import {assert,EMU_PER_TWIP} from '@office-ex/contracts';
import {patchCommand,assertPlainData,type DocumentSnapshot,type DocumentCommand,type WriterSection} from '@office-ex/core';
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
