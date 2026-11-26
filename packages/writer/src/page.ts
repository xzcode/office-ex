// SPDX-License-Identifier: MIT
import {assert,assertId,EMU_PER_TWIP} from '@office-ex/contracts';
import {patchCommand,assertPlainData,cloneData,writerStoryKey,type WriterStoryVariant,type WriterStoryParagraph,type PatchOperation,type Paragraph,type WriterSection as MutableSection,type DocumentSnapshot,type DocumentCommand,type WriterSection,type WriterPageStory,type ModelPath,type DeepReadonly,type WriterDocument} from '@office-ex/core';
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
    // New section inherits running stories and continues numbering.
    delete second.storyRefs;delete second.pageNumberStart;
    const empty:Paragraph={id:operationId+'-p',type:'paragraph',...(document.defaultParagraphStyleId?{styleId:document.defaultParagraphStyleId}:{}),runs:[{id:operationId+'-r',type:'text',text:''}]};
    if(!second.blocks.length)second.blocks.push(empty);
    // Focus the first editable paragraph in the following section. Do not
    // introduce an extra paragraph before an existing table or image.
    const focus=second.blocks.find(b=>b.type==='paragraph')?.id??'';
    return {sectionId:second.id,paragraphId:focus,command:patchCommand(snapshot,'WRT-03',[{op:'splice',path:['sections'],index,deleteCount:1,items:[first,second]}])};
  }
  assert(false,'invalid','Section break paragraph not found');
}

/** Return the effective (possibly inherited) default story reference. */
export function effectivePageStory(model:DeepReadonly<WriterDocument>,sectionId:string,kind:'header'|'footer',variant:WriterStoryVariant='default'):string|undefined{
  let ref:string|undefined;const slot=writerStoryKey(kind,variant);
  for(const s of model.sections){if(s.storyRefs?.[slot]!==undefined)ref=s.storyRefs[slot];if(s.id===sectionId)return ref;}
  assert(false,'invalid','Section not found');
}
/** Edit an explicit story; inherited stories are detached on edit so earlier
 * sections are not unexpectedly modified. Empty text creates a blank story. */
export function setWriterPageStory(snapshot:DocumentSnapshot,sectionId:string,kind:'header'|'footer',value:Pick<WriterPageStory,'text'|'after'|'pageNumber'|'alignment'>&{pageNumberStart?:number|null;distance?:number;variant?:WriterStoryVariant;differentFirstPage?:boolean;differentEvenOddPages?:boolean;paragraphs?:readonly DeepReadonly<WriterStoryParagraph>[]},operationId:string):DocumentCommand{
  assert(snapshot.document.kind==='writer','invalid','Expected Writer document');assertId(operationId);
  const doc=snapshot.document,index=doc.sections.findIndex(s=>s.id===sectionId);assert(index>=0,'invalid','Section not found');
  const section=doc.sections[index]!,refs={...section.storyRefs},stories=cloneData(doc.pageStories??{}) as Record<string,WriterPageStory>;
  const variant=value.variant??'default',slot=writerStoryKey(kind,variant);
  const effective=effectivePageStory(doc,sectionId,kind,variant),existing=effective?stories[effective]:undefined;
  // A distinct explicit part may still be shared by another section. Always
  // allocate for inherited or multiply referenced stories.
  const own=refs[slot],shared=own&&doc.sections.some((s,i)=>Object.entries(s.storyRefs??{}).some(([key,id])=>id===own&&(i!==index||key!==slot)));
  const id=own&&!shared?own:'rIdOexStory_'+operationId;
  assert(/^[A-Za-z_][A-Za-z0-9_.-]{0,127}$/.test(id),'invalid','Invalid story allocation');
  if(id!==own)assert(!Object.hasOwn(stories,id),'conflict','Duplicate story identity');
  const story:WriterPageStory={kind,text:value.text,...(value.pageNumber?{pageNumber:true,...(value.after?{after:value.after}:{})}:{}),...(value.alignment?{alignment:value.alignment}:{}),...(existing?.styleId?{styleId:existing.styleId}:{}),...(existing?.runStyle?{runStyle:existing.runStyle}:{})};
  if(value.paragraphs){
    assert(value.paragraphs.length>=1&&value.paragraphs.length<=16,'limit','Running story paragraph budget');
    const [first,...rest]=cloneData(value.paragraphs) as WriterStoryParagraph[];
    // Explicit paragraphs carry their own validated formatting and fields.
    stories[id]={kind,...first!,...(rest.length?{additionalParagraphs:rest}:{})};
  }else stories[id]={...story,...(existing?.additionalParagraphs?{additionalParagraphs:existing.additionalParagraphs}:{})};
  refs[slot]=id;
  const updated=cloneData(section) as MutableSection;updated.storyRefs=refs;
  if(value.differentFirstPage!==undefined){if(value.differentFirstPage)updated.differentFirstPage=true;else delete updated.differentFirstPage;}
  const key=kind==='header'?'headerDistance':'footerDistance',margin=section.page.margins[kind==='header'?'top':'bottom'];
  if(value.distance!==undefined){
    assert(Number.isSafeInteger(value.distance)&&value.distance>=0&&value.distance%EMU_PER_TWIP===0&&value.distance<=margin,'invalid','Invalid running story distance');
    updated.pageExtras={...updated.pageExtras,[key]:value.distance};
  }else if(!effective&&updated.pageExtras?.[key]===undefined)updated.pageExtras={...updated.pageExtras,[key]:Math.min(457200,Math.floor(margin/2/EMU_PER_TWIP)*EMU_PER_TWIP)};
  if(value.pageNumberStart!==undefined){
    assert(value.pageNumberStart===null||Number.isSafeInteger(value.pageNumberStart)&&value.pageNumberStart>=1&&value.pageNumberStart<=32767,'invalid','Invalid page number start');
    if(value.pageNumberStart===null)delete updated.pageNumberStart;else updated.pageNumberStart=value.pageNumberStart;
  }
  const operations:PatchOperation[]=[{op:'set',path:['pageStories'],value:stories},{op:'set',path:['sections',index],value:updated}];
  if(value.differentEvenOddPages!==undefined){if(value.differentEvenOddPages)operations.push({op:'set',path:['differentEvenOddPages'],value:true});else if(doc.differentEvenOddPages)operations.push({op:'delete',path:['differentEvenOddPages']});}
  return patchCommand(snapshot,'WRT-03',operations);
}
/** Undefined means continue numbering from the preceding section. */
export function setWriterPageNumberStart(snapshot:DocumentSnapshot,sectionId:string,start?:number):DocumentCommand{
  assert(snapshot.document.kind==='writer','invalid','Expected Writer document');
  const index=snapshot.document.sections.findIndex(s=>s.id===sectionId);assert(index>=0,'invalid','Section not found');
  assert(start===undefined||Number.isSafeInteger(start)&&start>=1&&start<=32767,'invalid','Invalid page number start');
  const s=cloneData(snapshot.document.sections[index]!) as MutableSection;
  if(start===undefined)delete s.pageNumberStart;else s.pageNumberStart=start;
  return patchCommand(snapshot,'WRT-03',[{op:'set',path:['sections',index],value:s}]);
}
