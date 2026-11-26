// SPDX-License-Identifier: MIT
import assert from 'node:assert/strict';
import {readFile,writeFile} from 'node:fs/promises';
import {join} from 'node:path';
import {cloneData,DocumentSession,writerEffectiveStyle} from '@office-ex/core';
import {insertTableRow,replaceParagraphRange,setTableBorder,effectivePageStory,setWriterPageStory,setWriterPageNumberStart} from '@office-ex/writer';
import {inspectForEditor,serializeDraft,prepareWriterDocxSave,prepareSlidesSave,prepareSheetsSave,writeLegacyPpt} from '@office-ex/format-service';
import {openOpc,openWriterDocx,saveWriterDocx,writeZip,openSheetsXlsx,saveSheetsXlsx,writeSheetsXlsx} from '@office-ex/ooxml';
const folder=process.argv[2],read=name=>readFile(join(folder,name)),write=(name,bytes)=>writeFile(join(folder,name),bytes);
async function inspect(name){const bytes=await read(name);return {bytes,result:await inspectForEditor(bytes,name,'compat-corpus')};}
const excel=await inspect('standard.xlsx');assert.equal(excel.result.report.readOnly,false);const workbook=cloneData(excel.result.editableSheets);
workbook.sheets[0].cells['1:2'].value={kind:'number',value:300};
const editedXlsx=prepareSheetsSave('xlsx',serializeDraft(workbook),'compat-corpus',excel.bytes);
await write('standard-edited.xlsx',editedXlsx);
assert.deepEqual(openOpc(excel.bytes).parts.get('docProps/app.xml'),openOpc(editedXlsx).parts.get('docProps/app.xml'));
const externalSheet=await inspect('m6-sheet.xlsx');assert.equal(externalSheet.result.report.readOnly,false,externalSheet.result.semanticFailure?.detail);
const externalSource=openSheetsXlsx(externalSheet.bytes,'compat-corpus','Report');
assert.deepEqual(saveSheetsXlsx(externalSheet.bytes,externalSource.document),externalSheet.bytes);
assert.deepEqual(externalSource.document.sheets[0].defaultGeometry,{rowHeight:26*12700,columnWidth:18*9525*7});
assert.equal(externalSource.document.sheets[0].cells['1:1'].style.borderEdges.bottom.line,'double');
const externalModel=cloneData(externalSheet.result.editableSheets);
externalModel.sheets[0].cells['1:2'].value={kind:'number',value:21};
const externalEdited=prepareSheetsSave('xlsx',serializeDraft(externalModel),'compat-corpus',externalSheet.bytes);
await write('m6-sheet-edited.xlsx',externalEdited);
const externalBack=openSheetsXlsx(externalEdited,'compat-corpus','Report');
assert.deepEqual(externalBack.document.sheets[0].defaultGeometry,externalSource.document.sheets[0].defaultGeometry);
assert.deepEqual(externalBack.document.sheets[0].cells['1:1'].style.borderEdges,externalSource.document.sheets[0].cells['1:1'].style.borderEdges);
const lookup=await inspect('m6-lookups.xlsx');
assert.equal(lookup.result.report.readOnly,false,lookup.result.semanticFailure?.detail);
const lookupSource=openSheetsXlsx(lookup.bytes,'compat-corpus','Lookup report');
assert.deepEqual(saveSheetsXlsx(lookup.bytes,lookupSource.document),lookup.bytes);
const lookupModel=cloneData(lookup.result.editableSheets);
lookupModel.sheets[1].cells['1:3'].value={kind:'number',value:20};
const lookupEdited=prepareSheetsSave('xlsx',serializeDraft(lookupModel),'compat-corpus',lookup.bytes);
await write('m6-lookups-edited.xlsx',lookupEdited);
const lookupBack=openSheetsXlsx(lookupEdited,'compat-corpus','Lookup report');
const cache=new Map(lookupBack.cachedValues);
for(const [i,n]of [20,2,30,20,30,500,50,1048572,40,30,1].entries())assert.deepEqual(cache.get('sheet-1!'+(i+1)+':2'),{kind:'number',value:n});
const charts=await inspect('m6-charts.xlsx');assert.equal(charts.result.report.readOnly,false,charts.result.semanticFailure?.detail);
const chartSource=openSheetsXlsx(charts.bytes,'compat-corpus','Chart report');
assert.deepEqual(chartSource.document.sheets[0].charts.map(c=>c.kind),['column','line','pie']);
assert.ok(Buffer.from(saveSheetsXlsx(charts.bytes,chartSource.document)).equals(charts.bytes));
chartSource.document.sheets[0].cells['2:2'].value={kind:'number',value:17};
const editedChart=saveSheetsXlsx(charts.bytes,chartSource.document);await write('m6-charts-edited.xlsx',editedChart);
await write('m6-charts-new.xlsx',writeSheetsXlsx(chartSource.document));
chartSource.document.sheets[0].charts.splice(1,1);chartSource.document.sheets[0].name='Renamed chart';
await write('m6-charts-rebuilt.xlsx',saveSheetsXlsx(charts.bytes,chartSource.document));
const diagonal=await inspect('m6-unsupported-diagonal.xlsx');assert.equal(diagonal.result.report.readOnly,true);assert.match(diagonal.result.semanticFailure.detail,/diagonal/);
const complicated=await inspect('extended.xlsx');assert.equal(complicated.result.report.readOnly,true);assert.match(complicated.result.semanticFailure.detail,/Unknown element r in si|Conditional|Hidden worksheet/);
assert.deepEqual(complicated.result.xlsxValuePreview.sheets.map(s=>s.name),['业务数据','汇总']);
assert.equal(complicated.result.xlsxValuePreview.sheets[0].cells['1:1'].value,'业务数据');
assert.equal(complicated.result.xlsxValuePreview.sheets[0].cells['1:3'].value,'150');
assert.equal(complicated.result.xlsxValuePreview.sheets[0].cells['1:5'],undefined);assert.equal(complicated.result.xlsxValuePreview.sheets[0].cells['3:1'],undefined);
for(const name of ['ordinary-text.docx','ordinary.pptx'])assert.equal((await inspect(name)).result.report.readOnly,false,name);
const word=await inspect('ordinary-image.docx');assert.equal(word.result.report.readOnly,false);const doc=cloneData(word.result.editableWriter);
doc.sections[0].blocks[0].runs[0].width=457200;doc.sections[0].blocks[0].runs[0].height=320040;
await write('image-edited.docx',prepareWriterDocxSave(serializeDraft(doc),'compat-corpus',word.bytes));
const tableFile=await inspect('ordinary-table.docx');assert.equal(tableFile.result.report.readOnly,false,tableFile.result.semanticFailure?.detail);
const sourceTable=openWriterDocx(tableFile.bytes,'compat-corpus','Table');
assert.deepEqual(saveWriterDocx(sourceTable,sourceTable.document),tableFile.bytes);
const tableSession=new DocumentSession(sourceTable.document,{saved:true});
const tableOf=d=>d.sections[0].blocks.find(b=>b.type==='table');
const t=tableOf(tableSession.snapshot().document);
assert.deepEqual({...t.nativeTable},{styleId:'TableGrid',autoWidth:true,autoLayout:true,look:0x4a0});
assert.equal(t.border.width,6350);assert.deepEqual({...t.cellPadding},{top:0,right:68580,bottom:0,left:68580});
const cell=t.rows[1].cells[0].blocks[0];assert.equal(writerEffectiveStyle(tableSession.snapshot().document,cell).spaceAfter,0);assert.equal(writerEffectiveStyle(tableSession.snapshot().document,cell).lineMultiple,240);
tableSession.execute(replaceParagraphRange(tableSession.snapshot(),cell.id,0,'办公软件'.length,'办公软件已编辑','m6-cell'));
tableSession.execute(insertTableRow(tableSession.snapshot(),t.id,2,'m6-row'));
const added=tableOf(tableSession.snapshot().document).rows[2].cells[0].blocks[0];
tableSession.execute(replaceParagraphRange(tableSession.snapshot(),added.id,0,0,'新增行','m6-added'));
const edited=prepareWriterDocxSave(serializeDraft(tableSession.snapshot().document),'compat-corpus',tableFile.bytes);
await write('table-edited.docx',edited);
const reopenedTable=tableOf(openWriterDocx(edited,'compat-corpus','Table').document);
assert.deepEqual({...reopenedTable.nativeTable},{...t.nativeTable});assert.equal(reopenedTable.rows.length,3);
const partsBefore=openOpc(tableFile.bytes),partsAfter=openOpc(edited);
for(const [name,data] of partsBefore.parts)if(name!==partsBefore.mainPart)assert.deepEqual(partsAfter.parts.get(name),data,name);
tableSession.execute(setTableBorder(tableSession.snapshot(),t.id,null));
const borderless=saveWriterDocx(sourceTable,tableSession.snapshot().document);
assert.equal(tableOf(openWriterDocx(borderless,'compat-corpus','Table').document).border,undefined);
assert.match(new TextDecoder().decode(openOpc(borderless).parts.get(partsBefore.mainPart)),/w:val="nil"/);
tableSession.undo();assert.ok(tableOf(tableSession.snapshot().document).border);
// Style names alone are not a compatibility guarantee. A changed built-in
// style or conditional style must keep the source protected.
for(const mutate of [
 x=>x.replace('w:styleId="TableGrid"','w:styleId="TableGrid" w:customStyle="1"'),
 x=>x.replace(/(<w:style w:type="table" w:styleId="TableGrid">[\s\S]*?)(<\/w:style>)/,'$1<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>$2')
]){
 const parts=new Map(partsBefore.parts),styleName='word/styles.xml',xml=new TextDecoder().decode(parts.get(styleName)),changed=mutate(xml);
 assert.notEqual(changed,xml,'negative fixture must actually change');
 parts.set(styleName,new TextEncoder().encode(changed));const bytes=writeZip(parts);
 assert.equal((await inspectForEditor(bytes,'protected-style.docx','compat-corpus')).report.readOnly,true);
}
// Even an explicit RGB black border / equal-to-style padding is a direct
// override, not an inherited automatic-color border.
{
 const parts=new Map(partsBefore.parts),main=partsBefore.mainPart,xml=new TextDecoder().decode(parts.get(main));
 const border='<w:tblBorders>'+['top','left','bottom','right','insideH','insideV'].map(e=>'<w:'+e+' w:val="single" w:sz="4" w:space="0" w:color="000000"/>').join('')+'</w:tblBorders>';
 const pad='<w:tblCellMar>'+['top','left','bottom','right'].map(e=>'<w:'+e+' w:w="'+(['left','right'].includes(e)?108:0)+'" w:type="dxa"/>').join('')+'</w:tblCellMar>';
 const changed=xml.replace('<w:tblLook ',border+pad+'<w:tblLook ');assert.notEqual(changed,xml);
 parts.set(main,new TextEncoder().encode(changed));const bytes=writeZip(parts),source=openWriterDocx(bytes,'compat-corpus','Override'),model=cloneData(source.document),table=tableOf(model);
 assert.equal(table.nativeTable.directBorder,true);assert.equal(table.nativeTable.directPadding,true);
 table.rows[0].cells[0].blocks[0].runs[0].text='Direct override';
 const saved=saveWriterDocx(source,model),after=new TextDecoder().decode(openOpc(saved).parts.get(main));
 assert.match(after,/w:color="000000"/);assert.match(after,/<w:tblCellMar>/);
}
const normal=await inspect('normal-table.docx');assert.equal(normal.result.report.readOnly,false,normal.result.semanticFailure?.detail);
const normalDoc=cloneData(normal.result.editableWriter);tableOf(normalDoc).rows[0].cells[0].blocks[0].runs[0].text='Normal edited';
await write('normal-table-edited.docx',prepareWriterDocxSave(serializeDraft(normalDoc),'compat-corpus',normal.bytes));
for(const [name,edits]of [
 ['merged-table.docx',[[0,0,'Merged edited',1,2]]],
 ['complex-merged-table.docx',[[0,0,'矩形已编辑',2,2],[0,2,'纵向已编辑',3,1],[3,0,'横向已编辑',1,3]]]
]){
 const file=await inspect(name);assert.equal(file.result.report.readOnly,false,file.result.semanticFailure?.detail);
 const source=openWriterDocx(file.bytes,'compat-corpus','Merged'),session=new DocumentSession(source.document,{saved:true});
 assert.deepEqual(saveWriterDocx(source,source.document),file.bytes);
 for(const [row,col,text,rows,columns]of edits){
  const anchor=tableOf(session.snapshot().document).rows[row].cells[col];
  assert.deepEqual({...anchor.merge},{rows,columns});
  const p=anchor.blocks[0],old=p.runs.map(r=>r.text??'').join('');
  session.execute(replaceParagraphRange(session.snapshot(),p.id,0,old.length,text,'merge-'+row+'-'+col));
 }
 const edited=prepareWriterDocxSave(serializeDraft(session.snapshot().document),'compat-corpus',file.bytes);
 const reopened=tableOf(openWriterDocx(edited,'compat-corpus','Merged').document);
 for(const [row,col,text,rows,columns]of edits){
  const a=reopened.rows[row].cells[col];assert.deepEqual({...a.merge},{rows,columns});assert.equal(a.blocks[0].runs.map(r=>r.text).join(''),text);
  for(let r=row;r<row+rows;r++)for(let c=col;c<col+columns;c++)if(r!==row||c!==col){assert.equal(reopened.rows[r].cells[c].coveredBy,a.id);assert.equal(reopened.rows[r].cells[c].blocks.length,0);}
 }
 await write(name.replace('.docx','-edited.docx'),edited);
}
const invalidMerge=await inspect('invalid-merged-table.docx');assert.equal(invalidMerge.result.report.readOnly,true);assert.ok(invalidMerge.result.semanticFailure?.detail);assert.throws(()=>openWriterDocx(invalidMerge.bytes,'invalid','Invalid'));
const running=await inspect('running-stories.docx');
assert.equal(running.result.report.readOnly,false,running.result.semanticFailure?.detail);
const runningSource=openWriterDocx(running.bytes,'compat-corpus','Running stories');
assert.deepEqual(saveWriterDocx(runningSource,runningSource.document),running.bytes);
const runningSession=new DocumentSession(runningSource.document,{saved:true});
const first=runningSource.document.sections[0],second=runningSource.document.sections[1];
assert.equal(first.pageNumberStart,5);assert.equal(second.pageNumberStart,undefined);
const header=effectivePageStory(runningSource.document,first.id,'header'),footer=effectivePageStory(runningSource.document,first.id,'footer');
assert.equal(effectivePageStory(runningSource.document,second.id,'header'),header);
assert.equal(effectivePageStory(runningSource.document,second.id,'footer'),footer);
assert.equal(runningSource.document.pageStories[header].text,'Independent running header');
assert.equal(runningSource.document.pageStories[footer].pageNumber,true);
assert.equal(runningSource.document.pageStories[footer].text,'Page ');
assert.equal(runningSource.document.pageStories[footer].after,' / report');
runningSession.execute(setWriterPageStory(runningSession.snapshot(),first.id,'footer',{text:'Updated page ',pageNumber:true,after:' / report',alignment:'center'},'edit-footer'));
runningSession.execute(setWriterPageStory(runningSession.snapshot(),second.id,'header',{text:'Detached section header'},'detached-header'));
runningSession.execute(setWriterPageNumberStart(runningSession.snapshot(),second.id,20));
const runningEdited=prepareWriterDocxSave(serializeDraft(runningSession.snapshot().document),'compat-corpus',running.bytes);
const runningReopened=openWriterDocx(runningEdited,'compat-corpus','Running stories').document;
assert.equal(runningReopened.sections[1].pageNumberStart,20);
assert.notEqual(runningReopened.sections[1].storyRefs.header,header);
assert.equal(effectivePageStory(runningReopened,second.id,'footer'),footer);
assert.equal(runningReopened.pageStories[header].text,'Independent running header');
assert.equal(runningReopened.pageStories[runningReopened.sections[1].storyRefs.header].text,'Detached section header');
assert.equal(runningReopened.pageStories[footer].text,'Updated page ');
await write('running-stories-edited.docx',runningEdited);
const unsupportedStory=await inspect('unsupported-story-field.docx');assert.equal(unsupportedStory.result.report.readOnly,true);assert.match(unsupportedStory.result.semanticFailure.detail,/Unsupported running story field instruction/);
const ppt=await inspect('ordinary-shapes.pptx');assert.equal(ppt.result.report.readOnly,false,ppt.result.semanticFailure?.detail);
const slides=cloneData(ppt.result.editableSlides);
assert.equal(slides.slides[0].shapes[0].textFrame.autofit.mode,'shape');
slides.slides[0].shapes[0].runs[0].text='编辑普通文本框';
await write('shapes-edited.pptx',prepareSlidesSave('pptx',serializeDraft(slides),'compat-corpus',ppt.bytes));
assert.throws(()=>writeLegacyPpt(slides));
for(const [name,kind,expected]of [['advanced.docx','word','Read-only content'],['advanced.pptx','slides','文件']]){
 const {bytes,result}=await inspect(name);assert.equal(result.report.readOnly,true,name);assert.equal(result.editableWriter,undefined);assert.equal(result.editableSlides,undefined);
 assert.ok(result.semanticFailure?.detail,name);assert.equal(result.officeTextPreview.kind,kind);
 const paragraphs=result.officeTextPreview.pages.flatMap(p=>p.blocks.flatMap(b=>b.kind==='paragraph'?[b.text]:b.rows.flat()));
 assert.ok(paragraphs.includes(expected),name);
 if(kind==='slides')assert.equal(result.officeTextPreview.pages.length,2);
 assert.throws(()=>kind==='word'?prepareWriterDocxSave(serializeDraft(doc),'compat-corpus',bytes):prepareSlidesSave('pptx',serializeDraft(slides),'compat-corpus',bytes));
}
console.log('Independent Excel/Word/PPT ordinary edit and protected content preview corpus passed');
