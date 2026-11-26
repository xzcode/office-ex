import test from 'node:test';
import assert from 'node:assert/strict';
import {createDocument,DocumentSession,type SheetsDocument} from '@office-ex/core';
import {OfficeError} from '@office-ex/contracts';
import {writeSheetsXlsx,openSheetsXlsx,saveSheetsXlsx,openOpc,writeZip} from '@office-ex/ooxml';
import {writeLegacyXls,openLegacyXls,saveLegacyXls,inspectForEditor,prepareSheetsSave,serializeDraft} from '@office-ex/format-service';
import {WorkbookCalculator,parseFormula,serializeFormula} from '@office-ex/sheets';
import {readCfb,writeCfb} from '@office-ex/cfb';

import {fixture} from './sheets-format-fixture.js';

for(const [format,write,open] of [['xlsx',writeSheetsXlsx,openSheetsXlsx],['xls',writeLegacyXls,openLegacyXls]] as const){
  test(format+' typed cells, formula source/cache, styles, dimensions, merge and view roundtrip',()=>{
    const d=fixture(),b=write(d),r=open(b,d.documentId,d.title),s=r.document.sheets[0];
    assert.equal(r.document.dateSystem,'1904');assert.deepEqual(s.cells['1:2'].value,d.sheets[0].cells['1:2'].value);assert.equal(s.cells['4:1'].style?.bold,true);assert.equal(s.cells['4:1'].style?.italic,true);assert.equal(s.cells['4:1'].style?.underline,true);assert.deepEqual(s.cells['4:1'].style?.fill,d.sheets[0].cells['4:1'].style?.fill);assert.deepEqual(s.cells['4:1'].style?.border,d.sheets[0].cells['4:1'].style?.border);
    assert.deepEqual(s.merges,d.sheets[0].merges);assert.deepEqual(s.view,d.sheets[0].view);assert.equal(s.rows['1'].height,254000);assert.ok(Math.abs(s.columns['1'].width!-952500)<=635);
    const calc=new WorkbookCalculator().calculate(r.document);assert.deepEqual(calc.get('sheet-2',2,1).value,{kind:'number',value:27});assert.deepEqual(calc.get('sheet-1',8,1).value,{kind:'boolean',value:true});assert.ok(r.cachedValues.some(([,v])=>v.kind==='string'&&v.value==='你好'));
    for(const [i,sh] of d.sheets.entries())for(const [key,c] of Object.entries(sh.cells)){if(c.value.kind==='formula'){const actual=r.document.sheets[i].cells[key].value;assert.equal(actual.kind,'formula');assert.equal(serializeFormula(parseFormula(actual.source)),serializeFormula(parseFormula(c.value.source)));}}
    assert.doesNotThrow(()=>open(write(r.document),d.documentId,d.title));
  });
  test(format+' verifier authenticates document kind/session and rejects unsafe source',async()=>{
    const d=fixture(),b=prepareSheetsSave(format,serializeDraft(d),d.documentId);const inspect=await inspectForEditor(b,'test.'+format,d.documentId);assert.equal(inspect.report.readOnly,false);assert.equal(inspect.editableSheetsFormat,format);assert.throws(()=>prepareSheetsSave(format,serializeDraft(d),'wrong-id'),OfficeError);
    d.sheets[0].cells['1:1']={value:{kind:'number',value:30}};
    const saved=prepareSheetsSave(format,serializeDraft(d),d.documentId,b);assert.deepEqual(open(saved,d.documentId).document.sheets[0].cells['1:1'].value,{kind:'number',value:30});
  });
}
test('XLS conversion refuses out-of-range cells and whole-axis formulas',()=>{
  const d=fixture();d.sheets[0].cells['65537:1']={value:{kind:'number',value:3}};assert.throws(()=>writeLegacyXls(d),OfficeError);delete d.sheets[0].cells['65537:1'];d.sheets[0].cells['1:1']={value:{kind:'formula',source:'=SUM(C:C)'}};assert.throws(()=>writeLegacyXls(d),OfficeError);
});
test('unknown XLS CFB stream forces read-only and forbids semantic save',async()=>{
  const d=fixture(),cfb=readCfb(writeLegacyXls(d));const b=writeCfb({entries:[...cfb.entries,{type:'stream',path:['Unknown'],data:Uint8Array.of(1,2,3)}]});const view=await inspectForEditor(b,'unknown.xls',d.documentId);assert.equal(view.report.readOnly,true);assert.throws(()=>prepareSheetsSave('xls',serializeDraft(d),d.documentId,b),OfficeError);
});
test('unknown XLSX package part forces read-only and forbids semantic save',async()=>{
  const d=fixture(),parts=new Map(openOpc(writeSheetsXlsx(d)).parts);parts.set('custom.xml',new TextEncoder().encode('<custom/>'));const b=writeZip(parts);const view=await inspectForEditor(b,'unknown.xlsx',d.documentId);assert.equal(view.report.readOnly,true);assert.throws(()=>prepareSheetsSave('xlsx',serializeDraft(d),d.documentId,b),OfficeError);
});

test('XLS and XLSX no-edit native saves are exact source copies',()=>{
 const d=fixture(),xlsx=writeSheetsXlsx(d),xls=writeLegacyXls(d);
 assert.deepEqual(saveSheetsXlsx(xlsx,openSheetsXlsx(xlsx,d.documentId,d.title).document),xlsx);
 assert.deepEqual(saveLegacyXls(xls,openLegacyXls(xls,d.documentId,d.title).document),xls);
});

test('XLSX noncanonical main URI survives native save and structural sheet changes',()=>{
 const d=fixture(),original=openOpc(writeSheetsXlsx(d)),parts=new Map(original.parts),enc=(value:string)=>new TextEncoder().encode(value),dec=(value:Uint8Array)=>new TextDecoder().decode(value);
 parts.set('book/main.xml',parts.get('xl/workbook.xml')!);parts.delete('xl/workbook.xml');
 const oldRels=dec(parts.get('xl/_rels/workbook.xml.rels')!);
 parts.set('book/_rels/main.xml.rels',enc(oldRels.replace(/Target="(worksheets\/[^"]+|styles.xml)"/g,(_all,path)=>'Target="/xl/'+path+'"')));parts.delete('xl/_rels/workbook.xml.rels');
 parts.set('_rels/.rels',enc(dec(parts.get('_rels/.rels')!).replace('xl/workbook.xml','book/main.xml')));
 parts.set('[Content_Types].xml',enc(dec(parts.get('[Content_Types].xml')!).replace('/xl/workbook.xml','/book/main.xml')));
 const bytes=writeZip(parts),source=openSheetsXlsx(bytes,d.documentId,d.title);assert.equal(source.native.opc.mainPart,'book/main.xml');
 assert.deepEqual(saveSheetsXlsx(bytes,source.document),bytes);
 const model=source.document;model.sheets.reverse();model.sheets[1]!.cells['1:1']={value:{kind:'number',value:32}};
 const edited=saveSheetsXlsx(bytes,model),after=openSheetsXlsx(edited,d.documentId,d.title);
 assert.equal(after.native.opc.mainPart,'book/main.xml');assert.deepEqual(after.native.sheetParts.map(s=>s.path),['xl/worksheets/sheet2.xml','xl/worksheets/sheet1.xml']);
 assert.equal(after.document.sheets[1]!.cells['1:1']!.value.kind,'number');
 assert.ok(after.cachedValues.some(([key,value])=>key==='sheet-1!1:1'&&value.kind==='number'&&value.value===34));
 assert.deepEqual(after.native.opc.parts.get('_rels/.rels'),parts.get('_rels/.rels'));
});

test('M6.4 partial and mixed cell borders, explicit default geometry, and source-backed edit roundtrip',async()=>{
 const {formatCells}=await import('@office-ex/sheets');
 const d=createDocument('sheets','edge-profile');
 const sheet=d.sheets[0]!,black={kind:'rgba' as const,red:0,green:0,blue:0,alpha:255},red={...black,red:255};
 sheet.defaultGeometry={rowHeight:25*12700,columnWidth:18*9525*7};
 sheet.cells['1:1']={value:{kind:'string',value:'Report'},style:{borderEdges:{bottom:{line:'double',color:black},left:{line:'mediumDashed',color:red}}}};
 sheet.cells['1:2']={value:{kind:'number',value:5}};sheet.cells['2:2']={value:{kind:'formula',source:'=IFERROR(PRODUCT(B1,2),0)'}};
 const bytes=writeSheetsXlsx(d),original=openSheetsXlsx(bytes,d.documentId,d.title);
 assert.deepEqual(original.document.sheets[0]!.defaultGeometry,sheet.defaultGeometry);
 assert.deepEqual(original.document.sheets[0]!.cells['1:1']!.style?.borderEdges,sheet.cells['1:1']!.style?.borderEdges);
 assert.deepEqual(saveSheetsXlsx(bytes,original.document),bytes);
 original.document.sheets[0]!.cells['1:2']!.value={kind:'number',value:8};
 const edited=saveSheetsXlsx(bytes,original.document),reopened=openSheetsXlsx(edited,d.documentId,d.title);
 assert.deepEqual(reopened.document.sheets[0]!.defaultGeometry,sheet.defaultGeometry);
 assert.deepEqual(reopened.document.sheets[0]!.cells['1:1']!.style?.borderEdges,sheet.cells['1:1']!.style?.borderEdges);
 assert.deepEqual(openOpc(edited).parts.get('xl/styles.xml'),openOpc(bytes).parts.get('xl/styles.xml'));
 assert.ok(reopened.cachedValues.some(([k,v])=>k==='sheet-1!2:2'&&v.kind==='number'&&v.value===16));
 const session=new DocumentSession(reopened.document);
 session.execute(formatCells(session.snapshot(),'sheet-1',{row:1,column:1,rowCount:1,columnCount:1},{style:{border:{width:9525,color:black}}}));
 const updated=session.snapshot().document as SheetsDocument;
 assert.equal(updated.sheets[0]!.cells['1:1']!.style?.borderEdges,undefined);
 assert.deepEqual(updated.sheets[0]!.cells['1:1']!.style?.border,{width:9525,color:black});
 const back=openSheetsXlsx(saveSheetsXlsx(edited,updated),d.documentId,d.title);
 assert.deepEqual(back.document.sheets[0]!.cells['1:1']!.style?.border,{width:9525,color:black});
 assert.throws(()=>writeLegacyXls(d),OfficeError);
});
