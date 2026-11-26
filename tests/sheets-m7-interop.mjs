// SPDX-License-Identifier: MIT
import assert from 'node:assert/strict';
import {readFile,writeFile} from 'node:fs/promises';
import {join} from 'node:path';
import {DocumentSession} from '@office-ex/core';
import {openSheetsXlsx,saveSheetsXlsx,openOpc} from '@office-ex/ooxml';
import {inspectForEditor,prepareSheetsSave,serializeDraft} from '@office-ex/format-service';
import {setCells,cellFromInput,formatCells,applyNamedCellStyle,insertRows} from '@office-ex/sheets';
const dir=process.argv[2],r=(row,column)=>({row,column,rowCount:1,columnCount:1});
async function open(name){
 const bytes=await readFile(join(dir,name)),source=openSheetsXlsx(bytes,'m7-interoperability','Report');
 const inspection=await inspectForEditor(bytes,name,'m7-interoperability');
 assert.equal(inspection.report.readOnly,false,inspection.semanticFailure?.detail);
 assert.ok(Buffer.from(saveSheetsXlsx(bytes,source.document)).equals(bytes));
 return {bytes,source,session:new DocumentSession(source.document)};
}
const shared=await open('shared.xlsx'),sh=shared.session,old=shared.source.document.sheets[0].cells['1:1'];
assert.equal(old.richText.length,2);assert.equal(old.richText[0].style.bold,true);assert.equal(old.richText[1].style.italic,true);
sh.execute(setCells(sh.snapshot(),'sheet-1',[{row:1,column:1,cell:cellFromInput('Title New Tail',{type:'text'},old)}]));
const sharedEdited=prepareSheetsSave('xlsx',serializeDraft(sh.snapshot().document),'m7-interoperability',shared.bytes);
await writeFile(join(dir,'shared-edited.xlsx'),sharedEdited);
const inline=await open('inline.xlsx'),s=inline.session;
assert.equal(inline.source.document.namedCellStyles[0].name,'Report Format');
assert.equal(inline.source.document.sheets[0].cells['1:1'].richText.length,2);
s.execute(formatCells(s.snapshot(),'sheet-1',r(1,2),{style:{italic:true}}));
s.execute(applyNamedCellStyle(s.snapshot(),'sheet-1',r(1,3),'Report Format'));
s.execute(insertRows(s.snapshot(),'sheet-1',4));
const edited=prepareSheetsSave('xlsx',serializeDraft(s.snapshot().document),'m7-interoperability',inline.bytes);
await writeFile(join(dir,'inline-edited.xlsx'),edited);
console.log('Source-backed XLSX rich string, named style and chart range writes passed');
