// SPDX-License-Identifier: MIT
// Cold formula calculation and warm format/undo; five measured repetitions after warmup.
import {createDocument,DocumentSession} from '@office-ex/core';
import {WorkbookCalculator,setCells,formatCells} from '@office-ex/sheets';
import {performance} from 'node:perf_hooks';
import {cpus} from 'node:os';
import assert from 'node:assert/strict';
const median=a=>a.sort((a,b)=>a-b)[Math.floor(a.length/2)];
const measure=(fn,n=5)=>{fn();return +median(Array.from({length:n},()=>{const now=performance.now();fn();return performance.now()-now})).toFixed(2);};
const d=createDocument('sheets','performance-batch'),s=d.sheets[0];
for(let r=1;r<=20000;r++)s.cells[r+':2']={value:{kind:'number',value:r}};
s.cells['1:1']={value:{kind:'number',value:17}};
for(let r=1;r<=200;r++)s.cells[r+':3']={value:{kind:'formula',source:'=SUM(A:A)'}};
const calcMs=measure(()=>{const result=new WorkbookCalculator().calculate(d,{maxSteps:2e6});assert.equal(result.get(s.id,1,3).value.value,17);},5);
const session=new DocumentSession(d),snapshot=()=>session.snapshot();
const formatMs=measure(()=>{session.execute(formatCells(snapshot(),s.id,{row:1,column:2,rowCount:100,columnCount:1},{style:{bold:true}}));session.undo();},5);
const result={node:process.version,cpu:cpus()[0]?.model,rows:20000,formulaCount:200,calcMs,formatAndUndoMs:formatMs,historyBytes:session.historyUsage.bytes,fullWorkbookJsonBytes:Buffer.byteLength(JSON.stringify(d))};
console.log(JSON.stringify(result));
