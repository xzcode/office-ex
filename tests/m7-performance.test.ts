// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import {createDocument,DocumentSession,cloneData,patchCommand, type SheetsDocument,type WriterDocument} from '@office-ex/core';
import {OfficeError,graphemeBoundaries} from '@office-ex/contracts';
import {WorkbookCalculator,setCells,formatCells,createCalculationUpdate,applyCalculationUpdate,renameSheet, type WorkbookSnapshot} from '@office-ex/sheets';
import {layoutWriter,WriterPaginationEngine,type FontMetrics} from '@office-ex/render';
import {setWriterPageStory,setWriterPageNumberStart} from '@office-ex/writer';
const fails=(code:string)=>(e:unknown)=>e instanceof OfficeError&&e.code===code;
const book=(s:DocumentSession)=>s.snapshot().document as WorkbookSnapshot;
const n=(value:number)=>({kind:'number' as const,value});

test('M7.4 cell edits retain unchanged cell identities and compact reversible history',()=>{
 const d=createDocument('sheets','perf-cells'),sh=d.sheets[0]!;
 for(let i=1;i<=5000;i++)sh.cells[i+':1']={value:n(i)};
 const s=new DocumentSession(d,{saved:true,historyBytes:100000}),initial=book(s),untouched=initial.sheets[0]!.cells['5000:1'];
 const result=s.execute(formatCells(s.snapshot(),sh.id,{row:1,column:1,rowCount:100,columnCount:1},{style:{bold:true}}));
 assert.equal(result.changedPaths.length,100);assert.equal(book(s).sheets[0]!.cells['5000:1'],untouched);
 assert.ok(s.historyUsage.bytes<30000);assert.equal(book(s).sheets[0]!.cells['1:1']!.style?.bold,true);
 s.undo();assert.deepEqual(book(s),initial);assert.equal(s.dirty,false);s.redo();
 const before=book(s),revision=s.revision;
 const bad=patchCommand(s.snapshot(),'SHT-02',[
  {op:'set',path:['sheets',0,'cells','1:1'],value:{value:n(44)}},
  {op:'set',path:['sheets',0,'cells','0:1'],value:{value:n(1)}}
 ]);
 assert.throws(()=>s.execute(bad));assert.equal(s.revision,revision);assert.equal(book(s),before);
});
test('M7.4 grouped sibling patches preserve sequential semantics and atomic inverse ordering',()=>{
 const d=createDocument('sheets','perf-patches');d.sheets[0]!.cells['1:1']={value:n(1)};
 const s=new DocumentSession(d),before=s.snapshot().document;
 s.execute(patchCommand(s.snapshot(),'SHT-02',[
  {op:'set',path:['sheets',0,'cells','1:1'],value:{value:n(2)}},
  {op:'delete',path:['sheets',0,'cells','1:1']},
  {op:'set',path:['sheets',0,'cells','1:1'],value:{value:n(3)}},
  {op:'set',path:['sheets',0,'cells','2:1'],value:{value:n(4)}}
 ]));
 assert.deepEqual(book(s).sheets[0]!.cells['1:1']!.value,n(3));s.undo();assert.deepEqual(s.snapshot().document,before);s.redo();
 assert.deepEqual(book(s).sheets[0]!.cells['2:1']!.value,n(4));
});
test('M7.4 calculation transport sends scalar-only initial state, exact revision deltas and no format payload',()=>{
 const d=createDocument('sheets','perf-worker'),a=d.sheets[0]!;
 a.cells['1:1']={value:n(5),style:{bold:true}};a.cells['1:2']={value:{kind:'formula',source:'=A1+1'}};
 a.cells['2:1']={value:{kind:'string',value:'Rich'},richText:[{text:'Rich',style:{bold:true}}]};
 d.sheets.push({id:'other',name:'Other',cells:{'1:1':{value:{kind:'formula',source:'=Sheet1!B1*2'}}},rows:{},columns:{},merges:[]});
 const s=new DocumentSession(d),calc=new WorkbookCalculator();
 let sent={document:book(s),revision:s.revision},update=createCalculationUpdate(undefined,sent.document,sent.revision);
 assert.equal(update.kind,'full');assert.doesNotMatch(JSON.stringify(update),/richText|bold|numberFormat/);
 let worker=applyCalculationUpdate(undefined,structuredClone(update));const first=calc.calculate(worker.document);
 assert.deepEqual(first.get('other',1,1).value,n(12));const other=worker.document.sheets[1];
 s.execute(formatCells(s.snapshot(),a.id,{row:1,column:1,rowCount:1,columnCount:1},{style:{italic:true}}));
 update=createCalculationUpdate(sent,book(s),s.revision);assert.deepEqual(update,{kind:'delta',baseRevision:0,revision:1,documentId:d.documentId,sheets:[]});
 const same=applyCalculationUpdate(worker,structuredClone(update));assert.equal(same.document,worker.document);assert.equal(calc.calculate(same.document).computedCells,0);
 assert.throws(()=>applyCalculationUpdate(same,update),fails('conflict'));worker=same;sent={document:book(s),revision:s.revision};
 s.execute(setCells(s.snapshot(),a.id,[{row:1,column:1,cell:{value:n(8)}}]));
 update=createCalculationUpdate(sent,book(s),s.revision);assert.equal(update.kind,'delta');if(update.kind==='delta')assert.deepEqual(update.sheets,[{id:a.id,changes:[['1:1',n(8)]]}]);
 worker=applyCalculationUpdate(worker,structuredClone(update));assert.equal(worker.document.sheets[1],other);
 assert.deepEqual(calc.calculate(worker.document).get('other',1,1).value,n(18));sent={document:book(s),revision:s.revision};
 s.undo();update=createCalculationUpdate(sent,book(s),s.revision);worker=applyCalculationUpdate(worker,structuredClone(update));assert.deepEqual(calc.calculate(worker.document).get('other',1,1).value,n(12));sent={document:book(s),revision:s.revision};
 s.execute(renameSheet(s.snapshot(),a.id,'Renamed'));update=createCalculationUpdate(sent,book(s),s.revision);assert.equal(update.kind,'full');
 worker=applyCalculationUpdate(worker,structuredClone(update));assert.deepEqual(calc.calculate(worker.document).get('other',1,1).value,n(12));
});
test('M7.4 single-column index skips unrelated rows, row scans remain budgeted',()=>{
 const d=createDocument('sheets','perf-sparse'),s=d.sheets[0]!;
 for(let i=1;i<=10000;i++)s.cells[i+':2']={value:n(i)};
 s.cells['1:1']={value:n(17)};
 for(let i=1;i<=100;i++)s.cells[i+':3']={value:{kind:'formula',source:'=SUM(A:A)'}};
 const calc=new WorkbookCalculator(),r=calc.calculate(d,{maxSteps:200000});
 assert.equal(r.computedCells,100);assert.deepEqual(r.get(s.id,100,3).value,n(17));assert.ok(r.steps<30000,r.steps.toString());
 // Multi-column traversal must not make invisible unbounded passes.
 const empty=createDocument('sheets','perf-budget'),sh=empty.sheets[0]!;
 for(let i=1;i<=5000;i++)sh.cells[i+':1']={value:n(i)};
 for(let i=1;i<=30;i++)sh.cells[i+':4']={value:{kind:'formula',source:'=SUM(B:C)'}};
 assert.throws(()=>new WorkbookCalculator().calculate(empty,{maxSteps:20000}),fails('limit'));
});
const font:FontMetrics={environmentId:'perf-layout',measure(text,request){const b=graphemeBoundaries(text);return {advance:(b.length-1)*50000,ascent:80000,descent:20000,actualFamily:request.family,clusterPrecision:'exact',clusters:b.slice(0,-1).map((start,i)=>({start,end:b[i+1]!,left:i*50000,right:(i+1)*50000}))};}};
test('M7.4 page number and header-only edits preserve body block identity and reuse cached body pages',()=>{
 const d=createDocument('writer','perf-pages');d.sections[0]!.blocks[0]!.type==='paragraph'&&d.sections[0]!.blocks[0]!.runs.push({id:'text-perf',type:'text',text:'Body'});
 const s=new DocumentSession(d);
 s.execute(setWriterPageStory(s.snapshot(),'section-1','header',{text:'Report',distance:127000},'header'));
 const engine=new WriterPaginationEngine(),a=engine.layout(s.snapshot(),font),blocks=(s.snapshot().document as WriterDocument).sections[0]!.blocks;
 s.execute(setWriterPageStory(s.snapshot(),'section-1','header',{text:'Updated',distance:127000},'header2'));
 assert.equal((s.snapshot().document as WriterDocument).sections[0]!.blocks,blocks);
 const b=engine.layout(s.snapshot(),font);assert.equal(engine.stats.sectionsReused,1);assert.equal(engine.stats.sectionsComputed,0);
 assert.equal(b.pages[0]!.lines,a.pages[0]!.lines);assert.equal(b.pages[0]!.runningStories![0]!.text,'Updated');
 s.execute(setWriterPageNumberStart(s.snapshot(),'section-1',10));const c=engine.layout(s.snapshot(),font);
 assert.equal(engine.stats.sectionsReused,1);assert.equal(c.pages[0]!.displayPageNumber,10);
 assert.deepEqual(c.pages,layoutWriter(s.snapshot(),font).pages);
});
