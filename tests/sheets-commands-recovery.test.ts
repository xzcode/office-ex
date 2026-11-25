import test from 'node:test';
import assert from 'node:assert/strict';
import {createDocument,DocumentSession} from '@office-ex/core';
import {OfficeError} from '@office-ex/contracts';
import {setCells,addSheet,renameSheet,insertRows,deleteRows,copyCells,mergeCells,unmergeCells,formatCells,pasteTsv,exportTsv,fillCells,parseCellInput,dateToSerial,serialToDate,formatCellValue,WorkbookCalculator,type WorkbookSnapshot} from '@office-ex/sheets';
function book(s:DocumentSession):WorkbookSnapshot{const d=s.snapshot().document;if(d.kind!=='sheets')throw Error('wrong kind');return d;}
test('sheet commands preserve cross-sheet formulas during rename, insertion, undo',()=>{
  const s=new DocumentSession(createDocument('sheets','commands'));s.execute(setCells(s.snapshot(),'sheet-1',[{row:1,column:1,cell:{value:{kind:'number',value:7}}}]));
  s.execute(addSheet(s.snapshot(),'Other','sheet-2'));s.execute(setCells(s.snapshot(),'sheet-2',[{row:1,column:1,cell:{value:{kind:'formula',source:'=Sheet1!A1*2'}}}]));
  s.execute(renameSheet(s.snapshot(),'sheet-1','Sales 2026'));s.execute(insertRows(s.snapshot(),'sheet-1',1));
  assert.equal(book(s).sheets[0].cells['2:1'].value.kind,'number');assert.deepEqual(new WorkbookCalculator().calculate(book(s)).get('sheet-2',1,1).value,{kind:'number',value:14});
  s.execute(deleteRows(s.snapshot(),'sheet-1',2));assert.deepEqual(new WorkbookCalculator().calculate(book(s)).get('sheet-2',1,1).value,{kind:'error',value:'#REF!'});s.undo();assert.deepEqual(new WorkbookCalculator().calculate(book(s)).get('sheet-2',1,1).value,{kind:'number',value:14});
});
test('TSV paste, copy relative formula, fill, format and merge are transactional',()=>{
  const s=new DocumentSession(createDocument('sheets','clipboard'));s.execute(pasteTsv(s.snapshot(),'sheet-1',1,1,'3\t= A1*2\n4\t=A2*2'));s.execute(copyCells(s.snapshot(),'sheet-1',{row:1,column:2,rowCount:1,columnCount:1},'sheet-1',3,2));assert.deepEqual(new WorkbookCalculator().calculate(book(s)).get('sheet-1',3,2).value,{kind:'number',value:0});
  s.execute(fillCells(s.snapshot(),'sheet-1',{row:1,column:1,rowCount:2,columnCount:1},'down'));assert.deepEqual(book(s).sheets[0].cells['2:1'].value,{kind:'number',value:3});
  s.execute(formatCells(s.snapshot(),'sheet-1',{row:1,column:1,rowCount:1,columnCount:2},{style:{bold:true},numberFormat:'0.00'}));assert.equal(book(s).sheets[0].cells['1:2'].style?.bold,true);
  const rg={row:10,column:1,rowCount:2,columnCount:2};s.execute(mergeCells(s.snapshot(),'sheet-1',rg));assert.throws(()=>setCells(s.snapshot(),'sheet-1',[{row:10,column:2,cell:{value:{kind:'number',value:4}}}]),OfficeError);s.execute(unmergeCells(s.snapshot(),'sheet-1',rg));assert.equal(book(s).sheets[0].merges.length,0);
  assert.match(exportTsv(book(s),'sheet-1',{row:1,column:1,rowCount:1,columnCount:2},{formulas:true}),/3.*=/);
});
test('regional input, percent and 1900 leap-year serial boundaries',()=>{
  assert.deepEqual(parseCellInput('1.234,5',{locale:'de-DE',type:'number',mode:'regional'}),{kind:'number',value:1234.5});
  assert.equal(dateToSerial(1900,2,28),59);assert.equal(dateToSerial(1900,3,1),61);assert.equal(dateToSerial(1904,1,1,'1904'),0);
  assert.match(formatCellValue({kind:'number',value:.125},'0.0%',{locale:'en-US'}),/12.5/);
});
