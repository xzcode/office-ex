// SPDX-License-Identifier: MIT
import assert from 'node:assert/strict';
import {readFile,writeFile} from 'node:fs/promises';
import {join} from 'node:path';
import {cloneData} from '@office-ex/core';
import {inspectForEditor,serializeDraft,prepareWriterDocxSave,prepareSlidesSave,prepareSheetsSave,writeLegacyPpt} from '@office-ex/format-service';
import {openOpc} from '@office-ex/ooxml';
const folder=process.argv[2],read=name=>readFile(join(folder,name)),write=(name,bytes)=>writeFile(join(folder,name),bytes);
async function inspect(name){const bytes=await read(name);return {bytes,result:await inspectForEditor(bytes,name,'compat-corpus')};}
const excel=await inspect('standard.xlsx');assert.equal(excel.result.report.readOnly,false);const workbook=cloneData(excel.result.editableSheets);
workbook.sheets[0].cells['1:2'].value={kind:'number',value:300};
const editedXlsx=prepareSheetsSave('xlsx',serializeDraft(workbook),'compat-corpus',excel.bytes);
await write('standard-edited.xlsx',editedXlsx);
assert.deepEqual(openOpc(excel.bytes).parts.get('docProps/app.xml'),openOpc(editedXlsx).parts.get('docProps/app.xml'));
const complicated=await inspect('extended.xlsx');assert.equal(complicated.result.report.readOnly,true);assert.match(complicated.result.semanticFailure.detail,/Conditional\/differential/);
assert.deepEqual(complicated.result.xlsxValuePreview.sheets.map(s=>s.name),['业务数据','汇总']);
assert.equal(complicated.result.xlsxValuePreview.sheets[0].cells['1:1'].value,'业务数据');
assert.equal(complicated.result.xlsxValuePreview.sheets[0].cells['1:3'].value,'150');
assert.equal(complicated.result.xlsxValuePreview.sheets[0].cells['1:5'],undefined);assert.equal(complicated.result.xlsxValuePreview.sheets[0].cells['3:1'],undefined);
for(const name of ['ordinary-text.docx','ordinary.pptx'])assert.equal((await inspect(name)).result.report.readOnly,false,name);
const word=await inspect('ordinary-image.docx');assert.equal(word.result.report.readOnly,false);const doc=cloneData(word.result.editableWriter);
doc.sections[0].blocks[0].runs[0].width=457200;doc.sections[0].blocks[0].runs[0].height=320040;
await write('image-edited.docx',prepareWriterDocxSave(serializeDraft(doc),'compat-corpus',word.bytes));
const ppt=await inspect('ordinary-shapes.pptx');assert.equal(ppt.result.report.readOnly,false,ppt.result.semanticFailure?.detail);
const slides=cloneData(ppt.result.editableSlides);
assert.equal(slides.slides[0].shapes[0].textFrame.autofit.mode,'shape');
slides.slides[0].shapes[0].runs[0].text='编辑普通文本框';
await write('shapes-edited.pptx',prepareSlidesSave('pptx',serializeDraft(slides),'compat-corpus',ppt.bytes));
assert.throws(()=>writeLegacyPpt(slides));
for(const [name,kind,expected]of [['ordinary-table.docx','word','办公软件'],['advanced.docx','word','Read-only content'],['advanced.pptx','slides','文件']]){
 const {bytes,result}=await inspect(name);assert.equal(result.report.readOnly,true,name);assert.equal(result.editableWriter,undefined);assert.equal(result.editableSlides,undefined);
 assert.ok(result.semanticFailure?.detail,name);assert.equal(result.officeTextPreview.kind,kind);
 const paragraphs=result.officeTextPreview.pages.flatMap(p=>p.blocks.flatMap(b=>b.kind==='paragraph'?[b.text]:b.rows.flat()));
 assert.ok(paragraphs.includes(expected),name);
 if(kind==='slides')assert.equal(result.officeTextPreview.pages.length,2);
 assert.throws(()=>kind==='word'?prepareWriterDocxSave(serializeDraft(doc),'compat-corpus',bytes):prepareSlidesSave('pptx',serializeDraft(slides),'compat-corpus',bytes));
}
console.log('Independent Excel/Word/PPT ordinary edit and protected content preview corpus passed');
