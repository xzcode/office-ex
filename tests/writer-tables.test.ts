// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import {OfficeError,EMU_PER_TWIP,graphemeBoundaries} from '@office-ex/contracts';
import {createDocument,DocumentSession,cloneData,type WriterDocument} from '@office-ex/core';
import {insertTableAfter,insertTableRow,deleteTableRow,insertTableColumn,deleteTableColumn,resizeTableColumn,setTableBorder,deleteTable,tableCellLocation,replaceParagraphRange,paragraphText,joinParagraphs,insertPageBreakAfter} from '@office-ex/writer';
import {layoutWriter,hitTestWriterPage,writerCaretBox,type FontMetrics} from '@office-ex/render';
import {writeWriterDocx,openWriterDocx,saveWriterDocx,openOpc,writeZip} from '@office-ex/ooxml';
import {inspectForEditor,prepareWriterDocxSave,serializeDraft} from '@office-ex/format-service';
import {makeTableFixture,tableOf,writerModel,cellParagraph} from './writer-table-fixture.js';
import {styledFixture} from './writer-styled-fixture.js';
const fails=(code:string)=>(e:unknown)=>e instanceof OfficeError&&e.code===code;
const total=(widths:readonly number[])=>widths.reduce((a,b)=>a+b,0);
test('WRT-03 table row and column edits preserve unaffected stable cell IDs and undo exact state',()=>{
  const s=makeTableFixture(),before=s.snapshot().document,t=tableOf(writerModel(s)),id=t.id,initialWidth=total(t.columnWidths),firstCell=t.rows[0]!.cells[0]!.id;
  s.execute(insertTableRow(s.snapshot(),id,1,'row-insert'));
  s.execute(insertTableColumn(s.snapshot(),id,1,'column-insert'));
  const grown=tableOf(writerModel(s));assert.equal(grown.rows.length,3);assert.equal(grown.columnWidths.length,3);assert.equal(grown.rows[0]!.cells[0]!.id,firstCell);assert.equal(total(grown.columnWidths),initialWidth);
  assert.ok(grown.rows.every(r=>r.cells.length===3));
  s.execute(resizeTableColumn(s.snapshot(),id,0,grown.columnWidths[0]!-100*EMU_PER_TWIP));
  assert.equal(total(tableOf(writerModel(s)).columnWidths),initialWidth);
  s.execute(deleteTableRow(s.snapshot(),id,1));s.execute(deleteTableColumn(s.snapshot(),id,1));
  assert.equal(tableOf(writerModel(s)).rows[0]!.cells[0]!.id,firstCell);
  assert.equal(total(tableOf(writerModel(s)).columnWidths),initialWidth);
  for(let i=0;i<5;i++)s.undo();
  assert.deepEqual(s.snapshot().document,before);
});
test('WRT-03 cell paragraphs use existing grapheme editing and block cross-cell joins or page breaks',()=>{
  const s=makeTableFixture(),id=cellParagraph(s,0,0),other=cellParagraph(s,0,1);
  assert.equal(paragraphText(s.snapshot(),id),'中文');
  assert.deepEqual(tableCellLocation(s.snapshot(),id),{tableId:tableOf(writerModel(s)).id,row:0,column:0,cellId:tableOf(writerModel(s)).rows[0]!.cells[0]!.id});
  assert.throws(()=>joinParagraphs(s.snapshot(),id,other),fails('invalid'));
  assert.throws(()=>insertPageBreakAfter(s.snapshot(),id,'no-break'),fails('unsupported'));
  const second=cellParagraph(s,0,0,1);s.execute(joinParagraphs(s.snapshot(),id,second));
  assert.equal(paragraphText(s.snapshot(),id),'中文 😀');
  assert.throws(()=>replaceParagraphRange(s.snapshot(),id,4,5,'x','split-emoji'),fails('invalid'));
  s.undo();assert.equal(paragraphText(s.snapshot(),id),'中文');
});
test('WRT-03 width, size and structural guards reject invalid requests without mutating source',()=>{
  const s=new DocumentSession(createDocument('writer','guards'));
  assert.throws(()=>insertTableAfter(s.snapshot(),'paragraph-1',501,2,'bad'),fails('limit'));
  s.execute(insertTableAfter(s.snapshot(),'paragraph-1',1,1,'table'));
  const id=tableOf(writerModel(s)).id,old=s.snapshot().document;
  for(const op of [()=>deleteTableRow(s.snapshot(),id,0),()=>deleteTableColumn(s.snapshot(),id,0),()=>resizeTableColumn(s.snapshot(),id,0,1),()=>insertTableColumn(s.snapshot(),id,-1,'bad')])assert.throws(op,fails('invalid'));
  assert.deepEqual(s.snapshot().document,old);
  s.execute(setTableBorder(s.snapshot(),id,null));assert.equal(tableOf(writerModel(s)).border,undefined);
  s.execute(deleteTable(s.snapshot(),id));assert.equal(writerModel(s).sections[0]!.blocks.length,1);
  s.undo();s.undo();assert.deepEqual(s.snapshot().document,old);
});
const font:FontMetrics={environmentId:'table-test',measure(text,request){
  const b=graphemeBoundaries(text);return {advance:(b.length-1)*50000,ascent:80000,descent:20000,actualFamily:request.family,clusterPrecision:'exact',clusters:b.slice(0,-1).map((start,i)=>({start,end:b[i+1]!,left:i*50000,right:(i+1)*50000}))};
}};
function layoutFixture():DocumentSession{
  const d=createDocument('writer','table-layout');d.defaultTextStyle={fontSize:100000,fontFamily:'test'};
  d.sections[0]!.page={width:2000000,height:600000,margins:{left:50000,right:50000,top:50000,bottom:50000}};
  const s=new DocumentSession(d);s.execute(insertTableAfter(s.snapshot(),'paragraph-1',4,2,'table'));
  for(let row=0;row<4;row++)for(let col=0;col<2;col++){const p=cellParagraph(s,row,col);s.execute(replaceParagraphRange(s.snapshot(),p,0,0,'Cell '+row+','+col,'cell-'+row+'-'+col));}
  return s;
}
test('WRT-02 table layout paginates complete rows, bounds text in each cell and hits the correct cell',()=>{
  const s=layoutFixture(),before=s.snapshot(),layout=layoutWriter(before,font);
  assert.ok(layout.pages.length>1);assert.equal(layout.pages.flatMap(p=>p.tableCells).length,8);
  const rowPages=new Map<string,number>();
  for(let i=0;i<layout.pages.length;i++){const page=layout.pages[i]!;
    for(const cell of page.tableCells){
      assert.ok(cell.y>=page.margins.top&&cell.y+cell.height<=page.height-page.margins.bottom);
      assert.equal(rowPages.get(cell.rowId)??i,i);rowPages.set(cell.rowId,i);
      const hit=hitTestWriterPage(layout,i,{x:cell.x+cell.width/2,y:cell.y+cell.height/2});
      assert.ok(hit&&cell.paragraphIds.includes(hit.paragraphId));
      const caret=writerCaretBox(layout,{paragraphId:cell.paragraphIds[0]!,offset:0},i);assert.ok(caret&&caret.pageIndex===i&&caret.x>=cell.x);
    }
  }
  assert.deepEqual(s.snapshot(),before);
});
test('WRT-02 a row taller than a page fails explicitly without cutting off cell contents',()=>{
  const s=layoutFixture(),p=cellParagraph(s,0,0),text=paragraphText(s.snapshot(),p);
  s.execute(replaceParagraphRange(s.snapshot(),p,0,text.length,'line\n'.repeat(20),'tall'));
  const before=s.snapshot();assert.throws(()=>layoutWriter(before,font),fails('unsupported'));assert.deepEqual(s.snapshot(),before);
});
test('WRT-04 rectangular DOCX table roundtrip preserves widths, margins, border, all cell paragraphs and text styles',async()=>{
  const s=makeTableFixture(),model=writerModel(s),original=writeWriterDocx(model),opened=openWriterDocx(original,'roundtrip','Table');
  const a=tableOf(model),b=tableOf(opened.document);
  assert.deepEqual(b.columnWidths,a.columnWidths);assert.deepEqual(b.cellPadding,a.cellPadding);assert.deepEqual(b.border,a.border);
  assert.equal(b.rows.length,2);assert.equal(b.rows[0]!.cells[0]!.blocks.length,2);
  const p=b.rows[0]!.cells[0]!.blocks[0]!;if(p.type!=='paragraph'||p.runs[0]?.type!=='text')throw Error();
  assert.equal(p.runs[0].text,'中文');assert.equal(p.runs[0].style?.bold,true);
  assert.deepEqual(saveWriterDocx(opened,opened.document),original);
  const inspection=await inspectForEditor(original,'table.docx','desktop-table');assert.ok(inspection.editableWriter);
  const edit=new DocumentSession(inspection.editableWriter,{saved:true}),t=tableOf(writerModel(edit));edit.execute(insertTableRow(edit.snapshot(),t.id,2,'added'));
  const id=cellParagraph(edit,2,0);edit.execute(replaceParagraphRange(edit.snapshot(),id,0,0,'Added row','text'));
  const updated=prepareWriterDocxSave(serializeDraft(edit.snapshot().document),'desktop-table',original),reopened=openWriterDocx(updated,'desktop-table','Table');
  assert.equal(tableOf(reopened.document).rows.length,3);
  const before=openOpc(original),after=openOpc(updated);
  for(const [name,data] of before.parts)if(name!==before.mainPart)assert.deepEqual(after.parts.get(name),data);
});
test('WRT-04 table merge, nesting, unequal grid widths and unknown properties stay read-only',async()=>{
  const source=writeWriterDocx(writerModel(makeTableFixture())),opc=openOpc(source),dec=new TextDecoder(),enc=new TextEncoder();
  for(const mutate of [
    (x:string)=>x.replace('</w:tcPr>','<w:gridSpan w:val="2"/></w:tcPr>'),
    (x:string)=>x.replace('</w:tcPr>','<w:vMerge/></w:tcPr>'),
    (x:string)=>x.replace('</w:tcPr>','</w:tcPr><w:tbl/>'),
    (x:string)=>x.replace('<w:tcW ','<w:tcW w:extra="1" '),
    (x:string)=>x.replace(/<w:gridCol w:w="\d+"\/>/,'<w:gridCol w:w="1"/>'),
    (x:string)=>x.replace('<w:tblLayout w:type="fixed"/>','<w:tblLayout w:type="autofit"/>')
  ]){
    const parts=new Map(opc.parts);parts.set(opc.mainPart!,enc.encode(mutate(dec.decode(parts.get(opc.mainPart!)!))));
    const bytes=writeZip(parts);assert.throws(()=>openWriterDocx(bytes,'read','Blocked'),fails('unsupported'));
    assert.equal((await inspectForEditor(bytes,'unknown.docx','read')).report.readOnly,true);
  }
});
test('WRT-04 source named styles stay preserved and new table cells inherit the document default paragraph style',()=>{
  const bytes=styledFixture(),source=openWriterDocx(bytes,'styled-table','Styled'),s=new DocumentSession(source.document);
  const first=source.document.sections[0]!.blocks[0]!;
  s.execute(insertTableAfter(s.snapshot(),first.id,1,2,'new-table'));
  const table=tableOf(writerModel(s));const p=table.rows[0]!.cells[0]!.blocks[0]!;
  if(p.type!=='paragraph')throw Error();
  assert.equal(p.styleId,'Normal');assert.equal(writerModel(s).defaultParagraphStyleId,'Normal');
  const updated=saveWriterDocx(source,s.snapshot().document),next=openWriterDocx(updated,'styled-table','Styled');
  assert.equal(tableOf(next.document).rows.length,1);const a=openOpc(bytes),b=openOpc(updated);
  for(const [path,data] of a.parts)if(path!==a.mainPart)assert.deepEqual(b.parts.get(path),data);
});
