// SPDX-License-Identifier: MIT
/** A logical UTF-16 projection, independent of visual layout. Non-text inlines
 * occupy one atomic logical unit. Adjacent text runs share grapheme boundaries.
 */
import {assert,graphemeBoundaries} from '@office-ex/contracts';
import type {Inline,Paragraph} from './model.js';
import type {DeepReadonly} from './session.js';
export const OBJECT_REPLACEMENT='\uFFFC';
export function inlineText(inline:DeepReadonly<Inline>):string{
  return inline.type==='text'?inline.text:inline.type==='tab'?'\t':inline.type==='line-break'?'\n':OBJECT_REPLACEMENT;
}
export function projectParagraph(p:DeepReadonly<Paragraph>):string{return p.runs.map(inlineText).join('');}
export function paragraphBoundaries(p:DeepReadonly<Paragraph>):number[]{
  const result=[0];let at=0,plain='';
  const flush=()=>{for(const b of graphemeBoundaries(plain).slice(1))result.push(at+b);at+=plain.length;plain='';};
  for(const run of p.runs){
    if(run.type==='text')plain+=run.text;
    else{flush();at++;result.push(at);}
  }
  flush();return result;
}
export function assertParagraphRange(p:DeepReadonly<Paragraph>,start:number,end:number):void{
  assert(Number.isSafeInteger(start)&&Number.isSafeInteger(end)&&start>=0&&end>=start,'invalid','Invalid paragraph selection');
  const positions=new Set(paragraphBoundaries(p));assert(positions.has(start)&&positions.has(end),'invalid','Selection splits a grapheme or inline object');
}
