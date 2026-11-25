// SPDX-License-Identifier: MIT
/** Logical range operations across contiguous paragraphs in one section or one
 * cell. A range must not cross a table, page break, nested cell or section. */
import {assert} from '@office-ex/contracts';
import {assertParagraphRange,patchCommand,cloneData,projectParagraph,type DocumentSnapshot,type DocumentCommand,type DeepReadonly,type Paragraph,type PatchOperation,type TextStyle,type Inline} from '@office-ex/core';
import {locateParagraph,writerIdAllocator,sliceWriterRuns,templateAt,inserted,nonEmpty,validateWriterInsertion,formatParagraphRange,replaceParagraphRange} from './paragraphs.js';
export interface WriterPoint {paragraphId:string;offset:number}
export interface WriterRange {anchor:WriterPoint;focus:WriterPoint}
export interface WriterRangeSegment {paragraphId:string;start:number;end:number;paragraph:DeepReadonly<Paragraph>}
export function resolveWriterRange(snapshot:DocumentSnapshot,range:WriterRange){
  const a=locateParagraph(snapshot,range.anchor.paragraphId),b=locateParagraph(snapshot,range.focus.paragraphId);
  assert(JSON.stringify(a.parent)===JSON.stringify(b.parent),'unsupported','Selection crosses a section or table cell');
  assertParagraphRange(a.paragraph,range.anchor.offset,range.anchor.offset);assertParagraphRange(b.paragraph,range.focus.offset,range.focus.offset);
  const reverse=a.block>b.block||a.block===b.block&&range.anchor.offset>range.focus.offset;
  const first=reverse?b:a,last=reverse?a:b,start=reverse?range.focus.offset:range.anchor.offset,end=reverse?range.anchor.offset:range.focus.offset;
  const segments:WriterRangeSegment[]=[];
  for(let i=first.block;i<=last.block;i++){
    const p=first.siblings[i]!;assert(p.type==='paragraph','unsupported','Selection crosses a non-paragraph block');
    segments.push({paragraphId:p.id,start:i===first.block?start:0,end:i===last.block?end:projectParagraph(p).length,paragraph:p});
  }
  return {first,last,start,end,segments};
}
/** Copy plain text only. Mixed-image selection must use explicit image actions,
 * never silently discard pictures when cutting or copying. */
export function writerRangeText(snapshot:DocumentSnapshot,range:WriterRange):string{
  const resolved=resolveWriterRange(snapshot,range);
  return resolved.segments.map(s=>{
    const fragments=sliceWriterRuns(s.paragraph.runs,s.start,s.end,new Set(),writerIdAllocator('copy-temp'));
    assert(fragments.every(r=>r.type!=='image'),'unsupported','Plain text clipboard cannot represent selected images');
    return projectParagraph({...cloneData(s.paragraph),runs:fragments});
  }).join('\n');
}
export function replaceWriterRange(snapshot:DocumentSnapshot,range:WriterRange,replacement:string,operationId:string):{command:DocumentCommand;caret:WriterPoint}{
  const r=resolveWriterRange(snapshot,range);validateWriterInsertion(replacement);
  if(r.first.block===r.last.block&&!/[\r\n]/.test(replacement)){
    return {command:replaceParagraphRange(snapshot,r.first.paragraph.id,r.start,r.end,replacement,operationId),caret:{paragraphId:r.first.paragraph.id,offset:r.start+replacement.length}};
  }
  const lines=replacement.replace(/\r\n?/g,'\n').split('\n');
  assert(lines.length<=10000,'limit','Pasted paragraph count exceeds budget');
  const allocate=writerIdAllocator(operationId),used=new Set<string>(),template=templateAt(r.first.runs,r.start);
  const left=sliceWriterRuns(r.first.runs,0,r.start,used,allocate),right=sliceWriterRuns(r.last.runs,r.end,r.last.text.length,used,allocate);
  const paragraphs:Paragraph[]=[];
  for(let i=0;i<lines.length;i++){
    const tail=i===lines.length-1,part=lines[i]!;
    // Newly inserted paragraph properties come from the start paragraph.
    // A distinct existing end paragraph keeps its identity/properties when a
    // paragraph boundary remains between the unaffected prefix and suffix.
    const source=tail&&i>0&&r.first.block!==r.last.block?r.last.paragraph:r.first.paragraph;
    const id=i===0?r.first.paragraph.id:tail&&r.first.block!==r.last.block?r.last.paragraph.id:allocate();
    const runs:Inline[]=[...(i===0?left:[]),...(part?[inserted(template,part,allocate)]:[]),...(tail?right:[])];
    paragraphs.push({...cloneData(source),id,runs:nonEmpty(runs,template,allocate)});
  }
  const caret={paragraphId:paragraphs.at(-1)!.id,offset:(lines.length===1?r.start:0)+lines.at(-1)!.length};
  return {command:patchCommand(snapshot,'WRT-03',[{op:'splice',path:r.first.parent,index:r.first.block,deleteCount:r.last.block-r.first.block+1,items:paragraphs}]),caret};
}
export function formatWriterRange(snapshot:DocumentSnapshot,range:WriterRange,style:TextStyle,operationId:string):DocumentCommand{
  const r=resolveWriterRange(snapshot,range);const operations:PatchOperation[]=[];const allocate=writerIdAllocator(operationId);
  for(const s of r.segments){
    if(s.start===s.end)continue;
    const command=formatParagraphRange(snapshot,s.paragraphId,s.start,s.end,style,allocate());
    assert(command.type==='model.patch','internal','Expected range patch');operations.push(...command.payload.operations);
  }
  return patchCommand(snapshot,'WRT-03',operations);
}
/** Native continuous editor for one adjacent text-only paragraph group. Newline
 * separators are mapped to paragraph endpoints; existing soft breaks stay in
 * their original run. An image/table/section is a hard boundary of the group.
 */
export interface WriterTextGroup {text:string;paragraphs:readonly {id:string;start:number;end:number}[]}
export function writerTextGroup(snapshot:DocumentSnapshot,paragraphId:string):WriterTextGroup{
  const loc=locateParagraph(snapshot,paragraphId);
  const eligible=(b:unknown):b is DeepReadonly<Paragraph>=>!!b&&typeof b==='object'&&(b as DeepReadonly<Paragraph>).type==='paragraph'&&(b as DeepReadonly<Paragraph>).runs.every(r=>r.type!=='image')&&!projectParagraph(b as DeepReadonly<Paragraph>).includes('\r');
  assert(eligible(loc.paragraph),'unsupported','Continuous text editing excludes images');
  let first=loc.block,last=loc.block;
  while(first>0&&eligible(loc.siblings[first-1]))first--;
  while(last+1<loc.siblings.length&&eligible(loc.siblings[last+1]))last++;
  const paragraphs:{id:string;start:number;end:number}[]=[];const text:string[]=[];let offset=0;
  for(let i=first;i<=last;i++){
    const p=loc.siblings[i]!;assert(eligible(p),'internal','Invalid group block');
    const part=projectParagraph(p);paragraphs.push({id:p.id,start:offset,end:offset+part.length});text.push(part);offset+=part.length+1;
  }
  assert(offset<=1_000_001,'limit','Continuous text group exceeds edit budget');
  return {text:text.join('\n'),paragraphs};
}
export function textGroupPoint(group:WriterTextGroup,offset:number):WriterPoint{
  assert(Number.isSafeInteger(offset)&&offset>=0&&offset<=group.text.length,'invalid','Invalid text group offset');
  const p=group.paragraphs.find(p=>offset>=p.start&&offset<=p.end);assert(p,'invalid','Text group position unavailable');
  return {paragraphId:p.id,offset:offset-p.start};
}
export function textGroupOffset(group:WriterTextGroup,point:WriterPoint):number{
  const p=group.paragraphs.find(p=>p.id===point.paragraphId);assert(p&&point.offset>=0&&point.offset<=p.end-p.start,'invalid','Point outside text group');
  return p.start+point.offset;
}
