// SPDX-License-Identifier: MIT
import {EMU_PER_CSS_PIXEL} from '@office-ex/contracts';
import type {DeepReadonly} from '@office-ex/core';
import type {WriterLine} from '@office-ex/render';
/** A source paragraph rule is decoration, not editable text or a glyph. */
export function writerParagraphRule(line:DeepReadonly<WriterLine>):HTMLDivElement|undefined{
 const border=line.bottomBorder;if(!border)return undefined;
 const {rule}=border,c=rule.color,px=(n:number)=>n/EMU_PER_CSS_PIXEL+'px';
 const node=document.createElement('div');node.className='writer-paragraph-bottom-rule';node.dataset.paragraphId=line.paragraphId;node.setAttribute('aria-hidden','true');
 Object.assign(node.style,{position:'absolute',left:px(border.x),top:px(line.y+line.height-rule.width),width:px(border.width),height:px(rule.width),backgroundColor:c.kind==='rgba'?'rgba('+[c.red,c.green,c.blue,c.alpha/255].join(',')+')':'#20242c',pointerEvents:'none'});
 return node;
}
