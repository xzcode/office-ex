// SPDX-License-Identifier: MIT
import assert from 'node:assert/strict';
import {readFile,writeFile} from 'node:fs/promises';
import {join} from 'node:path';
import {openSheetsXlsx,saveSheetsXlsx,openOpc} from '@office-ex/ooxml';
import {inspectForEditor,prepareSheetsSave,serializeDraft} from '@office-ex/format-service';
const dir=process.argv[2],source=await readFile(join(dir,'external.xlsx')),original=openSheetsXlsx(source,'external','Workbook'),model=original.document;
assert.deepEqual(model.sheets.map(s=>s.name),['Sales','Other','Keep','Remove']);
assert.equal(original.native.conversionBlocked,true);
assert.equal(model.sheets[0].cells['1:1'].value.value,'中文 Shared');
assert.equal(model.sheets[0].cells['1:1'].style.bold,true);
assert.equal(model.sheets[0].cells['1:1'].style.fill.kind,'rgba');
assert.equal(model.defaultTextStyle.fontFamily,'Calibri');
assert.ok(Buffer.from(saveSheetsXlsx(source,model)).equals(source));
const inspected=await inspectForEditor(source,'external.xlsx',model.documentId);
assert.equal(inspected.report.readOnly,false);assert.equal(inspected.sheetsTemplateConversionBlocked,true);
const [sales,other,keep]=model.sheets;
sales.cells['1:2'].value={kind:'number',value:25};
sales.cells['1:1'].value={kind:'string',value:'Edited 中文'};
sales.cells['3:1']={value:{kind:'number',value:12.345},style:{bold:true,fill:{kind:'rgba',red:255,green:220,blue:120,alpha:255}},numberFormat:'0.0000'};
model.sheets=[keep,other,sales,{id:'added-sheet',name:'New Sheet',cells:{'1:1':{value:{kind:'string',value:'Added'}}},rows:{},columns:{},merges:[]}];
const output=prepareSheetsSave('xlsx',serializeDraft(model),model.documentId,source);
await writeFile(join(dir,'external-edited.xlsx'),output);
const read=openSheetsXlsx(output,model.documentId,'Workbook');
assert.deepEqual(read.document.sheets.map(s=>s.name),['Keep','Other','Sales','New Sheet']);
assert.equal(read.document.sheets[2].cells['1:1'].value.value,'Edited 中文');
assert.ok(read.cachedValues.some(([k,v])=>k==='sheet-3!2:2'&&v.kind==='number'&&v.value===32));
assert.deepEqual(read.native.sheetParts.slice(0,3).map(s=>s.path),[original.native.sheetParts[2].path,original.native.sheetParts[1].path,original.native.sheetParts[0].path]);
const before=openOpc(source).parts,after=openOpc(output).parts;
for(const path of [...original.native.metadataPaths,original.native.themePath,original.native.sheetParts[2].path,'_rels/.rels'])assert.deepEqual(after.get(path),before.get(path),path);
assert.equal(after.has(original.native.sheetParts[3].path),false);
// A second edit in the same session deliberately uses the original import
// baseline, not positional identities from the previous output.
sales.cells['1:2'].value={kind:'number',value:40};
const second=prepareSheetsSave('xlsx',serializeDraft(model),model.documentId,source);
await writeFile(join(dir,'external-second.xlsx'),second);
const secondRead=openSheetsXlsx(second,model.documentId);
assert.deepEqual(secondRead.native.sheetParts.map(s=>s.path),read.native.sheetParts.map(s=>s.path));
assert.ok(secondRead.cachedValues.some(([k,v])=>k==='sheet-3!2:2'&&v.kind==='number'&&v.value===47));
for(const file of ['external-protected.xlsx','external-conditional.xlsx','external-custom-name.xlsx']){
 const bytes=await readFile(join(dir,file));assert.equal((await inspectForEditor(bytes,file,'protected')).report.readOnly,true,file);
}
console.log('Independent XLSX template admitted; byte-exact no-edit, theme/style/metadata/part preservation, shared-string references, sheet add/delete/reorder, formula caches and second save verified; 3 unsupported sources protected.');
