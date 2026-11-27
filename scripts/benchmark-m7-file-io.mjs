// SPDX-License-Identifier: MIT
// Targeted native-format adapter benchmark. Timings are diagnostic, not gates.
import {createDocument,DocumentSession} from '@office-ex/core';
import {writeSheetsXlsx,writeSlidesPptx} from '@office-ex/ooxml';
import {inspectForEditor,prepareSheetsSave,serializeDraft} from '@office-ex/format-service';
import {setCells} from '@office-ex/sheets';
import {performance} from 'node:perf_hooks';
import {cpus} from 'node:os';
import assert from 'node:assert/strict';
const median=a=>a.sort((a,b)=>a-b)[Math.floor(a.length/2)];
async function measure(fn){await fn();const samples=[];for(let i=0;i<3;i++){const start=performance.now();await fn();samples.push(performance.now()-start);}return +median(samples).toFixed(2);}
const d=createDocument('sheets','file-io-bench','Workbook'),s=d.sheets[0];
for(let r=1;r<=4000;r++){
 s.cells[r+':1']={value:{kind:'string',value:'Customer '+r}};
 s.cells[r+':2']={value:{kind:'number',value:r*7}};
 s.cells[r+':3']={value:{kind:'number',value:r%13}};
}
s.cells['1:4']={value:{kind:'formula',source:'=SUM(B:B)'}};
const bytes=writeSheetsXlsx(d);
const source=await inspectForEditor(bytes,'Workbook.xlsx',d.documentId);assert.ok(source.editableSheets);
const session=new DocumentSession(source.editableSheets);session.execute(setCells(session.snapshot(),s.id,[{row:1,column:2,cell:{value:{kind:'number',value:999}}}]));
const draft=serializeDraft(session.snapshot().document);
const p=createDocument('slides','file-io-deck','Deck');
p.slides=Array.from({length:30},(_,i)=>({id:'slide-'+(i+1),background:{kind:'rgba',red:255,green:255,blue:255,alpha:255},shapes:Array.from({length:10},(_,j)=>({id:'shape-'+i+'-'+j,type:'text',geometry:{x:127000+j*127000,y:127000+j*254000,width:4000000,height:220000,rotation:0},runs:[{id:'run-'+i+'-'+j,type:'text',text:'Report slide '+i+' line '+j+' — Office EX'}]}))}));
const deck=writeSlidesPptx(p);
const xlsxOpenMs=await measure(async()=>{const v=await inspectForEditor(bytes,'Workbook.xlsx',d.documentId);assert.equal(v.report.readOnly,false);});
const xlsxSaveMs=await measure(()=>{const result=prepareSheetsSave('xlsx',draft,d.documentId,bytes);assert.ok(result.length);});
const pptxOpenMs=await measure(async()=>{const v=await inspectForEditor(deck,'Deck.pptx',p.documentId);assert.equal(v.report.readOnly,false);});
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rows:4000,cells:12001,slides:30,shapes:300,xlsxBytes:bytes.length,pptxBytes:deck.length,xlsxOpenMs,xlsxSaveMs,pptxOpenMs}));
