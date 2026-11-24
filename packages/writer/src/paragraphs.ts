// SPDX-License-Identifier: MIT
/** M2 Writer transaction layer. Paragraph-wide offsets are UTF-16 at grapheme boundaries.
 * Operates on top-level and simple table-cell text paragraphs. Unsupported inline objects
 * and nested tables are never silently flattened. All changes go through DocumentSession.
 */
import { assert, assertId, assertUnicode, assertGraphemeRange, graphemeBoundaries } from '@office-ex/contracts';
import { patchCommand, validateTextStyle, assertPlainData, cloneData, type DocumentSnapshot, type DocumentCommand, type ModelPath, type DeepReadonly, type TextStyle, type Paragraph, type TextRun, type WriterDocument } from '@office-ex/core';

type ReadRun = DeepReadonly<TextRun>;
type ReadParagraph = DeepReadonly<Paragraph>;
type Location = { section:number; block:number; parent:ModelPath; inCell:boolean; paragraph:ReadParagraph; runs:readonly ReadRun[]; text:string };
function writer(snapshot:DocumentSnapshot):DeepReadonly<WriterDocument>{
  assert(snapshot.document.kind==='writer','invalid','Expected Writer document');
  return snapshot.document;
}
function locate(snapshot:DocumentSnapshot,id:string):Location{
  const model=writer(snapshot);
  function find(blocks:DeepReadonly<WriterDocument['sections'][number]['blocks']>,section:number,parent:ModelPath,inCell:boolean):Location|undefined{
    for(let block=0;block<blocks.length;block++){
      const p=blocks[block]!;
      if(p.id===id){
        assert(p.type==='paragraph','unsupported','Only paragraph editing is supported here');
        const runs:ReadRun[]=[];
        for(const run of p.runs){assert(run.type==='text','unsupported','Paragraph contains an unsupported inline object');runs.push(run);}
        return {section,block,parent,inCell,paragraph:p,runs,text:runs.map(r=>r.text).join('')};
      }
      if(p.type==='table'&&!inCell)for(let row=0;row<p.rows.length;row++)for(let col=0;col<p.rows[row]!.cells.length;col++){
        const found=find(p.rows[row]!.cells[col]!.blocks,section,[...parent,block,'rows',row,'cells',col,'blocks'],true);
        if(found)return found;
      }
    }
    return undefined;
  }
  for(let section=0;section<model.sections.length;section++){
    const found=find(model.sections[section]!.blocks,section,['sections',section,'blocks'],false);if(found)return found;
  }
  assert(false,'invalid','Paragraph not found');
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
  return patchCommand(snapshot,'WRT-03',[{op:'set',path:[...p.parent,p.block,'runs'],value:runs}]);
}
/** Replace across runs; insertion inherits the preceding run (or first run at offset 0). */
export function replaceParagraphRange(snapshot:DocumentSnapshot,paragraphId:string,start:number,end:number,replacement:string,operationId:string):DocumentCommand{
  const p=locate(snapshot,paragraphId);
  assertGraphemeRange(p.text,start,end);assertUnicode(replacement);
  assert(replacement.length<=1_000_000,'limit','Replacement exceeds text budget');
  if(p.text.slice(start,end)===replacement)return patchCommand(snapshot,'WRT-03',[]);
  // Typical typing inside one run keeps its ID and style instead of fragmenting the model.
  let cursor=0;
  for(let index=0;index<p.runs.length;index++){
    const run=p.runs[index]!,next=cursor+run.text.length;
    if(start>=cursor&&end<=next){const value=run.text.slice(0,start-cursor)+replacement+run.text.slice(end-cursor);if(value.length<=1_000_000)return patchCommand(snapshot,'WRT-03',[{op:'set',path:[...p.parent,p.block,'runs',index,'text'],value}]);}
    cursor=next;
  }
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
  return splitParagraphRange(snapshot,paragraphId,offset,offset,operationId,operationId+'-paragraph');
}
/** Enter replaces a selected range and splits the paragraph in one transaction. */
export function splitParagraphRange(snapshot:DocumentSnapshot,paragraphId:string,start:number,end:number,operationId:string,newParagraphId:string):DocumentCommand{
  const p=locate(snapshot,paragraphId);assertGraphemeRange(p.text,start,end);assertId(newParagraphId);
  const allocate=allocator(operationId),used=new Set<string>(),template=templateAt(p.runs,start);
  const left=nonEmpty(sliceRuns(p.runs,0,start,used,allocate),template,allocate);
  const right=nonEmpty(sliceRuns(p.runs,end,p.text.length,used,allocate),template,allocate);
  const original=cloneData(p.paragraph);
  return patchCommand(snapshot,'WRT-03',[{op:'splice',path:p.parent,index:p.block,deleteCount:1,items:[{...original,runs:left},{...original,id:newParagraphId,runs:right}]}]);
}
/** A native input change becomes one grapheme-safe minimal replacement. */
export function paragraphEditDiff(before:string,after:string):{start:number;end:number;replacement:string}|undefined{
  if(before===after)return undefined;
  const a=graphemeBoundaries(before),b=graphemeBoundaries(after);
  let prefix=0,suffix=0;
  while(prefix<a.length-1&&prefix<b.length-1&&before.slice(a[prefix],a[prefix+1])===after.slice(b[prefix],b[prefix+1]))prefix++;
  while(suffix<a.length-1-prefix&&suffix<b.length-1-prefix&&before.slice(a[a.length-2-suffix],a[a.length-1-suffix])===after.slice(b[b.length-2-suffix],b[b.length-1-suffix]))suffix++;
  return {start:a[prefix]!,end:a[a.length-1-suffix]!,replacement:after.slice(b[prefix],b[b.length-1-suffix])};
}
/** Join adjacent paragraphs in the same section; preceding paragraph properties win. */
export function joinParagraphs(snapshot:DocumentSnapshot,firstId:string,secondId:string):DocumentCommand{
  const first=locate(snapshot,firstId),second=locate(snapshot,secondId);
  assert(JSON.stringify(first.parent)===JSON.stringify(second.parent)&&first.block+1===second.block,'invalid','Paragraphs must be adjacent in the same section');
  const runs=[...first.runs,...second.runs].map(r=>cloneData(r));
  return patchCommand(snapshot,'WRT-03',[{op:'splice',path:first.parent,index:first.block,deleteCount:2,items:[{...cloneData(first.paragraph),runs}]}]);
}
/** Apply an existing paragraph style without deleting direct formatting. */
export function setParagraphNamedStyle(snapshot:DocumentSnapshot,paragraphId:string,styleId:string):DocumentCommand{
  const p=locate(snapshot,paragraphId),model=writer(snapshot);
  assertId(styleId);const style=Object.hasOwn(model.styles,styleId)?model.styles[styleId]:undefined;
  assert(style?.kind==='paragraph','unsupported','Expected a supported paragraph style');
  return patchCommand(snapshot,'WRT-03',[{op:'set',path:[...p.parent,p.block,'styleId'],value:styleId}]);
}
/** Apply an existing character style to a grapheme-safe selection. */
export function setCharacterNamedStyle(snapshot:DocumentSnapshot,paragraphId:string,start:number,end:number,styleId:string,operationId:string):DocumentCommand{
  const p=locate(snapshot,paragraphId),model=writer(snapshot);
  assertGraphemeRange(p.text,start,end);assertId(styleId);
  assert(Object.hasOwn(model.styles,styleId)&&model.styles[styleId]?.kind==='character','unsupported','Expected a supported character style');
  if(start===end)return patchCommand(snapshot,'WRT-03',[]);
  const allocate=allocator(operationId),used=new Set<string>();
  const left=sliceRuns(p.runs,0,start,used,allocate);
  const middle=sliceRuns(p.runs,start,end,used,allocate).map(r=>({...r,styleId}));
  const right=sliceRuns(p.runs,end,p.text.length,used,allocate);
  return runPatch(snapshot,p,[...left,...middle,...right]);
}
export function setParagraphMetrics(snapshot:DocumentSnapshot,paragraphId:string,metrics:Pick<TextStyle,'indent'|'spaceBefore'|'spaceAfter'|'lineHeight'>):DocumentCommand{
  const p=locate(snapshot,paragraphId);assertPlainData(metrics);validateTextStyle(metrics);
  assert(Object.keys(metrics).every(k=>['indent','spaceBefore','spaceAfter','lineHeight'].includes(k)),'invalid','Unexpected paragraph property');
  const style={...p.paragraph.style,...cloneData(metrics)};
  if(style.lineHeight===0)delete style.lineHeight;
  return patchCommand(snapshot,'WRT-03',[{op:'set',path:[...p.parent,p.block,'style'],value:style}]);
}
export function setParagraphAlignment(snapshot:DocumentSnapshot,paragraphId:string,alignment:NonNullable<TextStyle['alignment']>):DocumentCommand{
  const p=locate(snapshot,paragraphId);validateTextStyle({alignment});
  return patchCommand(snapshot,'WRT-03',[{op:'set',path:[...p.parent,p.block,'style'],value:{...p.paragraph.style,alignment}}]);
}
/** Page breaks are explicit blocks, not derived pagination. */
export function insertPageBreakAfter(snapshot:DocumentSnapshot,paragraphId:string,breakId:string):DocumentCommand{
  const p=locate(snapshot,paragraphId);assertId(breakId);assert(!p.inCell,'unsupported','Page breaks inside cells are not supported');
  return patchCommand(snapshot,'WRT-03',[{op:'splice',path:p.parent,index:p.block+1,deleteCount:0,items:[{id:breakId,type:'page-break'}]}]);
}
