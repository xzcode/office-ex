// SPDX-License-Identifier: MIT
import {assert,assertId,EMU_PER_TWIP} from '@office-ex/contracts';
import {patchCommand,cloneData,assertPlainData,validateColor,type DocumentSnapshot,type DocumentCommand,type DeepReadonly,type Table,type TableCell,type WriterDocument,type ModelPath,type Color} from '@office-ex/core';
export const TABLE_LIMITS=Object.freeze({rows:500,columns:64,cells:10000,minWidth:240*EMU_PER_TWIP});
export const DEFAULT_TABLE_PADDING=Object.freeze({top:60*EMU_PER_TWIP,right:90*EMU_PER_TWIP,bottom:60*EMU_PER_TWIP,left:90*EMU_PER_TWIP});
type TableLocation={section:number;index:number;table:DeepReadonly<Table>;path:ModelPath;maxWidth:number};
function model(snapshot:DocumentSnapshot):DeepReadonly<WriterDocument>{
  assert(snapshot.document.kind==='writer','invalid','Expected Writer document');return snapshot.document;
}
function get(snapshot:DocumentSnapshot,id:string):TableLocation{
  const d=model(snapshot);
  for(let s=0;s<d.sections.length;s++){
    const section=d.sections[s]!,index=section.blocks.findIndex(b=>b.id===id);
    if(index<0)continue;const table=section.blocks[index]!;
    assert(table.type==='table','invalid','Expected table ID');
    // All mutating table commands are limited to the rectangular paragraph-only profile.
    assert(table.rows.every(r=>r.cells.every(c=>c.blocks.every(b=>b.type==='paragraph'))),'unsupported','Table contains unsupported blocks or inline content');
    return {section:s,index,table,path:['sections',s,'blocks',index],maxWidth:section.page.width-section.page.margins.left-section.page.margins.right};
  }
  assert(false,'invalid','Table not found');
}
function counts(rows:number,cols:number):void{
  assert(Number.isInteger(rows)&&Number.isInteger(cols)&&rows>=1&&cols>=1,'invalid','Invalid table dimensions');
  assert(rows<=TABLE_LIMITS.rows&&cols<=TABLE_LIMITS.columns&&rows*cols<=TABLE_LIMITS.cells,'limit','Table cell budget exceeded');
}
function index(n:number,max:number,allowEnd=false):void{assert(Number.isInteger(n)&&n>=0&&n<(allowEnd?max+1:max),'invalid','Invalid table index');}
function allocate(prefix:string):()=>string{assertId(prefix);assert(prefix.length<=105,'invalid','Table operation ID too long');let n=0;return ()=>prefix+'-'+(++n);}
function cell(next:()=>string,styleId?:string):TableCell{return {id:next(),blocks:[{id:next(),type:'paragraph',...(styleId?{styleId}:{}),runs:[{id:next(),type:'text',text:''}]}]};}
/** Structural row/column editing is deliberately unavailable until it can
 * adjust spans without losing source cells. Cell content and column widths
 * remain editable. */
export function hasMergedTableCells(table:DeepReadonly<Table>):boolean{return table.rows.some(row=>row.cells.some(c=>!!c.merge));}
function requireUnmerged(table:DeepReadonly<Table>):void{assert(!hasMergedTableCells(table),'unsupported','Split merged cells before inserting or deleting table rows or columns');}
function widthBounds(table:DeepReadonly<Table>,widths:readonly number[],max:number):void{
  const pad=table.cellPadding??DEFAULT_TABLE_PADDING;
  assert(widths.every(w=>Number.isSafeInteger(w)&&w%EMU_PER_TWIP===0&&w>=Math.max(TABLE_LIMITS.minWidth,pad.left+pad.right+EMU_PER_TWIP)),'invalid','Column width is too small or not a whole twip');
  assert(widths.reduce((a,b)=>a+b,0)<=max,'invalid','Table width exceeds page content width');
  for(const row of table.rows)for(let col=0;col<row.cells.length;col++){
    const c=row.cells[col]!;if(c.coveredBy)continue;
    const width=widths.slice(col,col+(c.merge?.columns??1)).reduce((a,b)=>a+b,0);
    for(const block of c.blocks)if(block.type==='paragraph')for(const run of block.runs)if(run.type==='image')assert(run.width<=width-pad.left-pad.right,'unsupported','Resize the image before narrowing its cell');
  }
}
function replace(snapshot:DocumentSnapshot,loc:TableLocation,table:Table):DocumentCommand{
  return patchCommand(snapshot,'WRT-03',[{op:'set',path:loc.path,value:table}]);
}
export function insertTableAfter(snapshot:DocumentSnapshot,paragraphId:string,rows:number,cols:number,operationId:string):DocumentCommand{
  counts(rows,cols);const doc=model(snapshot),next=allocate(operationId);
  for(let s=0;s<doc.sections.length;s++){
    const section=doc.sections[s]!,at=section.blocks.findIndex(b=>b.id===paragraphId);
    if(at<0)continue;assert(section.blocks[at]!.type==='paragraph','invalid','Table insertion requires a top-level paragraph');
    const twips=Math.floor((section.page.width-section.page.margins.left-section.page.margins.right)/EMU_PER_TWIP),unit=Math.floor(twips/cols);
    const widths=Array.from({length:cols},(_,i)=>(unit+(i<twips-unit*cols?1:0))*EMU_PER_TWIP);
    const table:Table={id:next(),type:'table',columnWidths:widths,rows:Array.from({length:rows},()=>({id:next(),cells:Array.from({length:cols},()=>cell(next,doc.defaultParagraphStyleId))})),cellPadding:{...DEFAULT_TABLE_PADDING},border:{width:6350,color:{kind:'rgba',red:0,green:0,blue:0,alpha:255}}};
    widthBounds(table,widths,section.page.width-section.page.margins.left-section.page.margins.right);
    return patchCommand(snapshot,'WRT-03',[{op:'splice',path:['sections',s,'blocks'],index:at+1,deleteCount:0,items:[table]}]);
  }
  assert(false,'invalid','Table insertion paragraph not found');
}
export function insertTableRow(snapshot:DocumentSnapshot,id:string,at:number,operationId:string):DocumentCommand{
  const loc=get(snapshot,id),table=cloneData(loc.table) as Table;requireUnmerged(table);index(at,table.rows.length,true);counts(table.rows.length+1,table.columnWidths.length);
  const next=allocate(operationId);
  table.rows.splice(at,0,{id:next(),cells:table.columnWidths.map(()=>cell(next,model(snapshot).defaultParagraphStyleId))});
  return replace(snapshot,loc,table);
}
export function deleteTableRow(snapshot:DocumentSnapshot,id:string,at:number):DocumentCommand{
  const loc=get(snapshot,id),table=cloneData(loc.table) as Table;requireUnmerged(table);index(at,table.rows.length);assert(table.rows.length>1,'invalid','Delete the table explicitly to remove its final row');
  table.rows.splice(at,1);return replace(snapshot,loc,table);
}
/** Split an adjacent column to insert one without expanding the table. */
export function insertTableColumn(snapshot:DocumentSnapshot,id:string,at:number,operationId:string):DocumentCommand{
  const loc=get(snapshot,id),table=cloneData(loc.table) as Table;requireUnmerged(table);index(at,table.columnWidths.length,true);counts(table.rows.length,table.columnWidths.length+1);
  const neighbor=at===table.columnWidths.length?at-1:at,original=table.columnWidths[neighbor]!,half=Math.floor(original/(2*EMU_PER_TWIP))*EMU_PER_TWIP;
  table.columnWidths[neighbor]=original-half;table.columnWidths.splice(at,0,half);
  const next=allocate(operationId);for(const row of table.rows)row.cells.splice(at,0,cell(next,model(snapshot).defaultParagraphStyleId));
  widthBounds(table,table.columnWidths,loc.maxWidth);
  return replace(snapshot,loc,table);
}
export function deleteTableColumn(snapshot:DocumentSnapshot,id:string,at:number):DocumentCommand{
  const loc=get(snapshot,id),table=cloneData(loc.table) as Table;requireUnmerged(table);index(at,table.columnWidths.length);assert(table.columnWidths.length>1,'invalid','Delete the table explicitly to remove its final column');
  const removed=table.columnWidths.splice(at,1)[0]!,neighbor=Math.min(at,table.columnWidths.length-1);
  table.columnWidths[neighbor]!+=removed;for(const row of table.rows)row.cells.splice(at,1);
  return replace(snapshot,loc,table);
}
export function resizeTableColumn(snapshot:DocumentSnapshot,id:string,at:number,width:number):DocumentCommand{
  const loc=get(snapshot,id),table=cloneData(loc.table) as Table;index(at,table.columnWidths.length);
  const delta=width-table.columnWidths[at]!;
  table.columnWidths[at]=width;
  if(table.columnWidths.length>1){const neighbor=at===table.columnWidths.length-1?at-1:at+1;table.columnWidths[neighbor]!-=delta;}
  widthBounds(table,table.columnWidths,loc.maxWidth);
  return replace(snapshot,loc,table);
}
export function setTableBorder(snapshot:DocumentSnapshot,id:string,border:{width:number;color:Color}|null):DocumentCommand{
  const loc=get(snapshot,id),table=cloneData(loc.table) as Table;
  if(border){assertPlainData(border);assert(Number.isSafeInteger(border.width)&&border.width>=1&&border.width<=152400,'invalid','Invalid border width');validateColor(border.color);table.border=cloneData(border);}
  else delete table.border;
  return replace(snapshot,loc,table);
}
export function deleteTable(snapshot:DocumentSnapshot,id:string):DocumentCommand{
  const loc=get(snapshot,id),section=model(snapshot).sections[loc.section]!;
  assert(section.blocks.length>1,'invalid','A section needs at least one block');
  return patchCommand(snapshot,'WRT-03',[{op:'splice',path:['sections',loc.section,'blocks'],index:loc.index,deleteCount:1,items:[]}]);
}
/** Merge an unmerged rectangle, retaining every original paragraph and
 * its ID in row-major order. Covered cells contain no duplicate content. */
export function mergeTableCells(snapshot:DocumentSnapshot,id:string,row:number,column:number,rows:number,columns:number):DocumentCommand{
  const loc=get(snapshot,id),table=cloneData(loc.table) as Table;
  index(row,table.rows.length);index(column,table.columnWidths.length);counts(rows,columns);
  assert(rows*columns>1&&row+rows<=table.rows.length&&column+columns<=table.columnWidths.length,'invalid','Merge exceeds table bounds');
  const anchor=table.rows[row]!.cells[column]!,blocks:TableCell['blocks']=[];
  for(let r=row;r<row+rows;r++)for(let c=column;c<column+columns;c++){
    const cell=table.rows[r]!.cells[c]!;assert(!cell.merge&&!cell.coveredBy,'unsupported','Merge overlaps an existing merge');
    blocks.push(...cell.blocks);
  }
  anchor.blocks=blocks;anchor.merge={rows,columns};
  for(let r=row;r<row+rows;r++)for(let c=column;c<column+columns;c++){
    if(r===row&&c===column)continue;
    const cell=table.rows[r]!.cells[c]!;cell.blocks=[];cell.coveredBy=anchor.id;
  }
  widthBounds(table,table.columnWidths,loc.maxWidth);
  return replace(snapshot,loc,table);
}
/** Split a merged rectangle. All existing content stays in its top-left
 * cell; the other cells receive new empty paragraphs. Undo restores the
 * original merge and exact stable content identities. */
export function splitTableCell(snapshot:DocumentSnapshot,id:string,row:number,column:number,operationId:string):DocumentCommand{
  const loc=get(snapshot,id),table=cloneData(loc.table) as Table;
  index(row,table.rows.length);index(column,table.columnWidths.length);
  const anchor=table.rows[row]!.cells[column]!,m=anchor.merge;
  assert(m&&!anchor.coveredBy,'invalid','Select the merged cell anchor');
  const next=allocate(operationId),styleId=model(snapshot).defaultParagraphStyleId;
  delete anchor.merge;
  for(let r=row;r<row+m.rows;r++)for(let c=column;c<column+m.columns;c++){
    if(r===row&&c===column)continue;
    const slot=table.rows[r]!.cells[c]!;delete slot.coveredBy;slot.blocks=[{id:next(),type:'paragraph',...(styleId?{styleId}:{}),runs:[{id:next(),type:'text',text:''}]}];
  }
  widthBounds(table,table.columnWidths,loc.maxWidth);
  return replace(snapshot,loc,table);
}
/** Locate a cell by paragraph ID; source coordinates stay stable until a structural edit. */
export function tableCellLocation(snapshot:DocumentSnapshot,paragraphId:string):{tableId:string;row:number;column:number;cellId:string}|undefined{
  for(const section of model(snapshot).sections)for(const b of section.blocks)if(b.type==='table')
    for(let row=0;row<b.rows.length;row++)for(let column=0;column<b.rows[row]!.cells.length;column++){
      const c=b.rows[row]!.cells[column]!;if(c.blocks.some(p=>p.id===paragraphId))return {tableId:b.id,row,column,cellId:c.id};
    }
  return undefined;
}
