// SPDX-License-Identifier: MIT
/** M2 Writer transaction layer. Paragraph-wide offsets are UTF-16 at grapheme boundaries.
 * This first slice operates on top-level, text-only paragraphs; unsupported inline objects
 * and tables are never silently flattened. All changes go through DocumentSession.
 */
import { assert, assertId, assertUnicode, assertGraphemeRange } from '@office-ex/contracts';
import { patchCommand, validateTextStyle, assertPlainData, cloneData, type DocumentSnapshot, type DocumentCommand, type DeepReadonly, type TextStyle, type Paragraph, type TextRun, type WriterDocument } from '@office-ex/core';

type ReadRun = DeepReadonly<TextRun>;
type ReadParagraph = DeepReadonly<Paragraph>;
type Location = { section:number; block:number; paragraph:ReadParagraph; runs:readonly ReadRun[]; text:string };
function writer(snapshot:DocumentSnapshot):DeepReadonly<WriterDocument>{
  assert(snapshot.document.kind==='writer','invalid','Expected Writer document');
  return snapshot.document;
}
function locate(snapshot:DocumentSnapshot,id:string):Location{
  const model=writer(snapshot);
  for(let section=0;section<model.sections.length;section++){
    const block=model.sections[section]!.blocks.findIndex(b=>b.id===id);
    if(block<0)continue;
    const paragraph=model.sections[section]!.blocks[block]!;
    assert(paragraph.type==='paragraph','unsupported','Only paragraph editing is supported here');
    const runs:ReadRun[]=[];
    for(const run of paragraph.runs){
      assert(run.type==='text','unsupported','Paragraph contains an unsupported inline object');
      runs.push(run);
    }
    return {section,block,paragraph,runs,text:runs.map(r=>r.text).join('')};
  }
  assert(false,'invalid','Top-level paragraph not found');
}
export function paragraphText(snapshot:DocumentSnapshot,paragraphId:string):string{return locate(snapshot,paragraphId).text;}
function allocator(prefix:string):()=>string{
  assertId(prefix);assert(prefix.length<=110,'invalid','Operation ID is too long');let n=0;
  return ()=>prefix+'-'+(++n);
}
function asText(run:ReadRun,id:string,text:string):TextRun{
  return {...cloneData(run),id,text};
}
function templateAt(runs:readonly ReadRun[],offset:number):ReadRun|undefined{
  let cursor=0;
  for(const run of runs){const end=cursor+run.text.length;if(end>=offset&&run.text.length)return run;cursor=end;}
  return runs.at(-1);
}
function sliceRuns(runs:readonly ReadRun[],start:number,end:number,used:Set<string>,allocate:()=>string):TextRun[]{
  let cursor=0;const result:TextRun[]=[];
  for(const run of runs){
    const a=Math.max(start,cursor),b=Math.min(end,cursor+run.text.length);
    if(a<b){
      const id=used.has(run.id)?allocate():run.id;used.add(id);
      result.push(asText(run,id,run.text.slice(a-cursor,b-cursor)));
    }
    cursor+=run.text.length;
  }
  return result;
}
function inserted(template:ReadRun|undefined,text:string,allocate:()=>string):TextRun{
  return template?asText(template,allocate(),text):{id:allocate(),type:'text',text};
}
function nonEmpty(runs:TextRun[],template:ReadRun|undefined,allocate:()=>string):TextRun[]{
  return runs.length?runs:[inserted(template,'',allocate)];
}
function runPatch(snapshot:DocumentSnapshot,p:Location,runs:TextRun[]):DocumentCommand{
  return patchCommand(snapshot,'WRT-03',[{op:'set',path:['sections',p.section,'blocks',p.block,'runs'],value:runs}]);
}
/** Replace across runs; insertion inherits the preceding run (or first run at offset 0). */
export function replaceParagraphRange(snapshot:DocumentSnapshot,paragraphId:string,start:number,end:number,replacement:string,operationId:string):DocumentCommand{
  const p=locate(snapshot,paragraphId);
  assertGraphemeRange(p.text,start,end);assertUnicode(replacement);
  assert(replacement.length<=1_000_000,'limit','Replacement exceeds text budget');
  if(p.text.slice(start,end)===replacement)return patchCommand(snapshot,'WRT-03',[]);
  const allocate=allocator(operationId),used=new Set<string>(),template=templateAt(p.runs,start);
  const left=sliceRuns(p.runs,0,start,used,allocate);
  const right=sliceRuns(p.runs,end,p.text.length,used,allocate);
  return runPatch(snapshot,p,nonEmpty([...left,...(replacement?[inserted(template,replacement,allocate)]:[]),...right],template,allocate));
}
const CHARACTER_KEYS=['fontFamily','fontSize','bold','italic','underline','color'];
/** Split affected runs at grapheme-safe range endpoints and apply direct character style. */
export function formatParagraphRange(snapshot:DocumentSnapshot,paragraphId:string,start:number,end:number,style:TextStyle,operationId:string):DocumentCommand{
  const p=locate(snapshot,paragraphId);assertGraphemeRange(p.text,start,end);assertPlainData(style);validateTextStyle(style);
  assert(Object.keys(style).every(k=>CHARACTER_KEYS.includes(k)),'unsupported','Only direct character properties are supported');
  if(start===end||!Object.keys(style).length)return patchCommand(snapshot,'WRT-03',[]);
  const allocate=allocator(operationId),used=new Set<string>();
  const left=sliceRuns(p.runs,0,start,used,allocate);
  const middle=sliceRuns(p.runs,start,end,used,allocate).map(r=>({...r,style:{...r.style,...cloneData(style)}}));
  const right=sliceRuns(p.runs,end,p.text.length,used,allocate);
  return runPatch(snapshot,p,[...left,...middle,...right]);
}
/** One Enter operation creates a paragraph, retaining paragraph properties and run styles. */
export function splitParagraph(snapshot:DocumentSnapshot,paragraphId:string,offset:number,operationId:string):DocumentCommand{
  const p=locate(snapshot,paragraphId);assertGraphemeRange(p.text,offset,offset);
  const allocate=allocator(operationId),used=new Set<string>(),template=templateAt(p.runs,offset);
  const left=nonEmpty(sliceRuns(p.runs,0,offset,used,allocate),template,allocate);
  const right=nonEmpty(sliceRuns(p.runs,offset,p.text.length,used,allocate),template,allocate);
  const original=cloneData(p.paragraph);
  return patchCommand(snapshot,'WRT-03',[{op:'splice',path:['sections',p.section,'blocks'],index:p.block,deleteCount:1,items:[{...original,runs:left},{...original,id:allocate(),runs:right}]}]);
}
/** Join adjacent paragraphs in the same section; preceding paragraph properties win. */
export function joinParagraphs(snapshot:DocumentSnapshot,firstId:string,secondId:string):DocumentCommand{
  const first=locate(snapshot,firstId),second=locate(snapshot,secondId);
  assert(first.section===second.section&&first.block+1===second.block,'invalid','Paragraphs must be adjacent in the same section');
  const runs=[...first.runs,...second.runs].map(r=>cloneData(r));
  return patchCommand(snapshot,'WRT-03',[{op:'splice',path:['sections',first.section,'blocks'],index:first.block,deleteCount:2,items:[{...cloneData(first.paragraph),runs}]}]);
}
export function setParagraphAlignment(snapshot:DocumentSnapshot,paragraphId:string,alignment:NonNullable<TextStyle['alignment']>):DocumentCommand{
  const p=locate(snapshot,paragraphId);validateTextStyle({alignment});
  return patchCommand(snapshot,'WRT-03',[{op:'set',path:['sections',p.section,'blocks',p.block,'style'],value:{...p.paragraph.style,alignment}}]);
}
/** Page breaks are explicit blocks, not derived pagination. */
export function insertPageBreakAfter(snapshot:DocumentSnapshot,paragraphId:string,breakId:string):DocumentCommand{
  const p=locate(snapshot,paragraphId);assertId(breakId);
  return patchCommand(snapshot,'WRT-03',[{op:'splice',path:['sections',p.section,'blocks'],index:p.block+1,deleteCount:0,items:[{id:breakId,type:'page-break'}]}]);
}
