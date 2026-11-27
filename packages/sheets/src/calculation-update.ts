// SPDX-License-Identifier: MIT
/** Revision-bound, value-only workbook transport for a dedicated local worker.
 * No calculation result or worker delta mutates the live document model. */
import {assert} from '@office-ex/contracts';
import type {CellValue} from '@office-ex/core';
import type {CalculationWorkbook,WorkbookSnapshot} from './calculation.js';
export interface CalculationVersion {readonly revision:number;readonly document:CalculationWorkbook}
export type CalculationUpdate=
 |{readonly kind:'full';readonly revision:number;readonly document:CalculationWorkbook}
 |{readonly kind:'delta';readonly revision:number;readonly baseRevision:number;readonly documentId:string;readonly sheets:readonly {readonly id:string;readonly changes:readonly (readonly [string,CellValue|null])[]}[]};

function sameValue(a:CellValue|undefined,b:CellValue):boolean{
 if(!a||a.kind!==b.kind)return false;
 if(a.kind==='blank')return true;
 if(a.kind==='formula'&&b.kind==='formula')return a.source===b.source;
 return 'value'in a&&'value'in b&&Object.is(a.value,b.value);
}
function full(document:WorkbookSnapshot,revision:number):CalculationUpdate{
 return {kind:'full',revision,document:{kind:'sheets',documentId:document.documentId,sheets:document.sheets.map(s=>({id:s.id,name:s.name,cells:Object.fromEntries(Object.entries(s.cells).map(([k,c])=>[k,{value:c.value}]))}))}};
}
export function createCalculationUpdate(previous:CalculationVersion|undefined,document:WorkbookSnapshot,revision:number):CalculationUpdate{
 const before=previous?.document;
 if(!previous||!before||before.documentId!==document.documentId||previous.revision>=revision||before.sheets.length!==document.sheets.length||before.sheets.some((s,i)=>s.id!==document.sheets[i]?.id||s.name!==document.sheets[i]?.name))return full(document,revision);
 const sheets:{id:string;changes:[string,CellValue|null][]}[]=[];
 for(const [i,sheet]of document.sheets.entries()){
  const old=before.sheets[i]!;if(old.cells===sheet.cells)continue;
  const changes:[string,CellValue|null][]=[];
  for(const [key,cell]of Object.entries(sheet.cells))if(!sameValue(old.cells[key]?.value,cell.value))changes.push([key,cell.value]);
  for(const key of Object.keys(old.cells))if(!Object.hasOwn(sheet.cells,key))changes.push([key,null]);
  if(changes.length)sheets.push({id:sheet.id,changes});
 }
 return {kind:'delta',revision,baseRevision:previous.revision,documentId:document.documentId,sheets};
}
/** Apply only to the exact acknowledged base. Retain unchanged sheet and
 * cell-map identities so the calculator can reuse its sparse indexes. */
export function applyCalculationUpdate(previous:CalculationVersion|undefined,update:CalculationUpdate):CalculationVersion{
 assert(Number.isSafeInteger(update.revision)&&update.revision>=0,'invalid','Invalid calculation revision');
 if(update.kind==='full')return {document:update.document,revision:update.revision};
 assert(previous&&previous.revision===update.baseRevision&&update.revision>update.baseRevision&&previous.document.documentId===update.documentId,'conflict','Calculation delta base mismatch');
 const doc=previous.document,changed=new Map(update.sheets.map(s=>[s.id,s.changes]));
 assert(changed.size===update.sheets.length&&[...changed.keys()].every(id=>doc.sheets.some(s=>s.id===id)),'invalid','Invalid calculation delta sheet');
 if(!changed.size)return {document:doc,revision:update.revision};
 const sheets=doc.sheets.map(sheet=>{
  const changes=changed.get(sheet.id);if(!changes?.length)return sheet;
  const cells={...sheet.cells},seen=new Set<string>();
  for(const [key,value]of changes){
   assert(/^[1-9]\d{0,6}:[1-9]\d{0,4}$/.test(key)&&!seen.has(key),'invalid','Invalid calculation delta cell');seen.add(key);
   const [r,c]=key.split(':').map(Number);assert(r!<=1048576&&c!<=16384,'invalid','Calculation cell outside workbook');
   if(value===null)delete cells[key];else cells[key]={value};
  }
  return {...sheet,cells};
 });
 return {document:{...doc,sheets},revision:update.revision};
}
