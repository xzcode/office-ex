// SPDX-License-Identifier: MIT
import {assert} from '@office-ex/contracts';
import {resolveTextStyle} from './resources.js';
import type {DeepReadonly} from './session.js';
import type {Paragraph,TextStyle,WriterDocument,WriterBlock,WriterListLevel} from './model.js';
export type WriterReadDocument=DeepReadonly<WriterDocument>;
export type WriterReadParagraph=DeepReadonly<Paragraph>;
// Table Grid's bounded source-owned paragraph defaults participate before
// named paragraph styles and direct formatting. Index only immutable snapshots.
const gridParagraphs=new WeakMap<WriterReadDocument,WeakSet<WriterReadParagraph>>();
function gridCells(model:WriterReadDocument):WeakSet<WriterReadParagraph>{
 const cached=gridParagraphs.get(model);if(cached)return cached;
 const set=new WeakSet<WriterReadParagraph>();
 for(const section of model.sections)for(const block of section.blocks)if(block.type==='table'&&block.nativeTable?.styleId==='TableGrid')
  for(const row of block.rows)for(const cell of row.cells)for(const p of cell.blocks)if(p.type==='paragraph')set.add(p);
 if(Object.isFrozen(model))gridParagraphs.set(model,set);return set;
}
export function writerEffectiveStyle(model:WriterReadDocument,p:WriterReadParagraph,grid?:boolean):DeepReadonly<TextStyle>{
 const defaults={...model.defaultTextStyle};
 if(grid??gridCells(model).has(p)){delete defaults.lineHeight;delete defaults.minLineHeight;defaults.spaceAfter=0;defaults.lineMultiple=240;}
 const id=p.styleId??model.defaultParagraphStyleId,style=resolveTextStyle(defaults,model.styles,id?[id]:[],p.style??{}).style,ref=style.writerList;
 if(!ref?.numId)return style;
 const level=model.numbering?.lists[String(ref.numId)]?.levels[ref.level];assert(level,'unsupported','Missing list level');
 // Numbering provides the default indentation; explicit paragraph indentation
 // wins. Text starts at the body column, not at the marker's hanging origin.
 return {...style,indent:p.style?.indent??level.indent,firstLineIndent:0};
}
export interface WriterNumberingMarker {text:string;level:DeepReadonly<WriterListLevel>;style:DeepReadonly<TextStyle>;indent:number;hanging:number}
export function buildWriterNumberingMarkers(model:WriterReadDocument):ReadonlyMap<string,WriterNumberingMarker>{
 const markers=new Map<string,WriterNumberingMarker>(),counters=new Map<number,(number|undefined)[]>();
 const visit=(blocks:readonly DeepReadonly<WriterBlock>[],grid=false)=>{
  for(const p of blocks){
   if(p.type==='table'){for(const row of p.rows)for(const cell of row.cells)visit(cell.blocks,p.nativeTable?.styleId==='TableGrid');continue;}
   if(p.type!=='paragraph')continue;
   const style=writerEffectiveStyle(model,p,grid),ref=style.writerList;if(!ref?.numId)continue;
   assert(style.paragraphDirection!=='rtl','unsupported','RTL list marker placement is not in the current profile');
   const def=model.numbering!.lists[String(ref.numId)]!,level=def.levels[ref.level]!;
   let values=counters.get(ref.numId);if(!values){values=[];counters.set(ref.numId,values);}
   values[ref.level]=(values[ref.level]??level.start-1)+1;
   for(let i=ref.level+1;i<def.levels.length;i++)if(!def.levels[i]!.neverRestart)values[i]=undefined;
   const text=level.format==='bullet'?level.pattern:level.pattern.replace(/%([1-9])/g,(_,n:string)=>{
    const i=Number(n)-1;return String(values![i]??def.levels[i]!.start);
   });
   const hanging=p.style?.firstLineIndent!==undefined?-p.style.firstLineIndent:level.hanging,indent=style.indent??level.indent;
   assert(hanging>0&&indent>=hanging,'unsupported','List marker requires an in-bounds hanging indent');
   markers.set(p.id,{text,level,indent,hanging,style:{...style,characterSpacing:0}});
  }
 };
 for(const section of model.sections)visit(section.blocks);
 return markers;
}
