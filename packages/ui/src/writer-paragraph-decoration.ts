// SPDX-License-Identifier: MIT
import {bundledFontFamily} from '@office-ex/pdf';
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

/** Numbering is derived decoration, never part of logical paragraph offsets. */
export function writerListMarker(line:DeepReadonly<WriterLine>):HTMLDivElement|undefined{
 const marker=line.listMarker;if(!marker)return undefined;
 const style=marker.style,c=style.color,px=(n:number)=>n/EMU_PER_CSS_PIXEL+'px',node=document.createElement('div');
 node.className='writer-list-marker';node.dataset.paragraphId=line.paragraphId;node.textContent=marker.text;
 Object.assign(node.style,{position:'absolute',left:px(marker.x),top:px(line.y),height:px(line.contentHeight??line.height),lineHeight:px(line.contentHeight??line.height),fontFamily:bundledFontFamily(style.fontFamily,style.language),fontSize:px(style.fontSize??152400),fontWeight:style.bold?'700':'400',fontStyle:style.italic?'italic':'normal',color:c?.kind==='rgba'?'rgba('+[c.red,c.green,c.blue,c.alpha/255].join(',')+')':'#20242c',whiteSpace:'pre',pointerEvents:'none'});
 return node;
}
