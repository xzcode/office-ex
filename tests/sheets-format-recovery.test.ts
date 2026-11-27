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

test('M6.6 numeric conditional rules preserve source styles, derive paint, append styles and move ranges atomically',async()=>{
 const {addConditionalRule,removeConditionalRule,conditionalCellStyle,insertRows,deleteRows}=await import('@office-ex/sheets');
 const d=createDocument('sheets','conditional-rules'),sheet=d.sheets[0]!;
 const red={kind:'rgba' as const,red:255,green:20,blue:30,alpha:255},green={...red,red:0,green:180};
 sheet.cells['2:1']={value:{kind:'number',value:20}};
 sheet.cells['3:1']={value:{kind:'formula',source:'=A2*2'}};
 const s=new DocumentSession(d),r={row:2,column:1,rowCount:3,columnCount:1};
 s.execute(addConditionalRule(s.snapshot(),sheet.id,{ranges:[r],operator:'greaterThan',values:[10],style:{fill:red,bold:true}}));
 s.execute(addConditionalRule(s.snapshot(),sheet.id,{ranges:[r],operator:'between',values:[0,30],style:{fill:green,italic:true},stopIfTrue:true}));
 const m=()=>s.snapshot().document as SheetsDocument;
 assert.deepEqual(conditionalCellStyle(m().sheets[0]!.conditionalRules!,2,1,{kind:'number',value:20}),{fill:red,bold:true,italic:true});
 assert.equal(conditionalCellStyle(m().sheets[0]!.conditionalRules!,1,1,{kind:'number',value:20}),undefined);
 assert.equal(conditionalCellStyle(m().sheets[0]!.conditionalRules!,2,1,{kind:'string',value:'20'}),undefined);
 const before=writeSheetsXlsx(m()),original=openSheetsXlsx(before,d.documentId,d.title);
 assert.deepEqual(original.document.sheets[0]!.conditionalRules,m().sheets[0]!.conditionalRules);
 assert.deepEqual(saveSheetsXlsx(before,original.document),before);
 const oldStyles=openOpc(before).parts.get('xl/styles.xml');
 original.document.sheets[0]!.cells['2:1']!.value={kind:'number',value:50};
 const saved=saveSheetsXlsx(before,original.document);
 assert.deepEqual(openOpc(saved).parts.get('xl/styles.xml'),oldStyles);
 s.execute(insertRows(s.snapshot(),sheet.id,2,1));assert.equal(m().sheets[0]!.conditionalRules![0]!.ranges[0]!.row,3);
 s.undo();assert.equal(m().sheets[0]!.conditionalRules![0]!.ranges[0]!.row,2);
 s.execute(deleteRows(s.snapshot(),sheet.id,2,3));assert.equal(m().sheets[0]!.conditionalRules,undefined);s.undo();
 s.execute(removeConditionalRule(s.snapshot(),sheet.id,1));
 s.execute(addConditionalRule(s.snapshot(),sheet.id,{ranges:[r],operator:'lessThan',values:[0],style:{color:green}}));
 const next=saveSheetsXlsx(before,m()),back=openSheetsXlsx(next,d.documentId,d.title);
 assert.deepEqual(back.document.sheets[0]!.conditionalRules,m().sheets[0]!.conditionalRules);
 assert.throws(()=>writeLegacyXls(m()),OfficeError);
});

test('M6.6 native chart graph creation, no-edit preservation, source data updates, rename and removal',async()=>{
 const {addSheetChart,removeSheetChart,renameSheet,insertRows}=await import('@office-ex/sheets');
 const d=createDocument('sheets','charts-model'),sheet=d.sheets[0]!;
 sheet.cells['1:1']={value:{kind:'string',value:'Name'}};sheet.cells['1:2']={value:{kind:'string',value:'Value'}};
 for(let i=2;i<=4;i++){sheet.cells[i+':1']={value:{kind:'string',value:'Category '+i}};sheet.cells[i+':2']={value:{kind:'number',value:i*2}};}
 const session=new DocumentSession(d),r={row:1,column:1,rowCount:4,columnCount:2};
 for(const kind of ['column','line','pie'] as const)session.execute(addSheetChart(session.snapshot(),sheet.id,r,kind));
 const model=()=>session.snapshot().document as SheetsDocument;
 const bytes=writeSheetsXlsx(model()),source=openSheetsXlsx(bytes,d.documentId,d.title);
 assert.deepEqual(source.document.sheets[0]!.charts,model().sheets[0]!.charts);
 assert.deepEqual(saveSheetsXlsx(bytes,source.document),bytes);
 const initial=session.snapshot().document;
 session.execute(insertRows(session.snapshot(),sheet.id,2));
 assert.equal(model().sheets[0]!.charts![0]!.categories.row,3);
 session.undo();assert.deepEqual(session.snapshot().document,initial);
 assert.throws(()=>writeLegacyXls(model()),OfficeError);
 const before=openOpc(bytes);
 source.document.sheets[0]!.cells['2:2']!.value={kind:'number',value:19};
 const edited=saveSheetsXlsx(bytes,source.document),after=openOpc(edited);
 for(const [name,part]of before.parts)if(name.startsWith('xl/drawings/')||name.startsWith('xl/charts/'))assert.deepEqual(after.parts.get(name),part);
 session.execute(removeSheetChart(session.snapshot(),sheet.id,1));
 session.execute(renameSheet(session.snapshot(),sheet.id,'Changed & More'));
 const next=saveSheetsXlsx(bytes,model()),back=openSheetsXlsx(next,d.documentId,d.title);
 assert.deepEqual(back.document.sheets[0]!.charts,model().sheets[0]!.charts);
 assert.equal(back.document.sheets[0]!.name,'Changed & More');
 session.undo();session.undo();assert.equal(model().sheets[0]!.charts!.length,3);
});

test('M7.1 rich cell text edits preserve grapheme-aligned formatting, formulas and source runs',async()=>{
 const {cellFromInput,setCells,clearCells,formatCells,copyCells}=await import('@office-ex/sheets');
 const d=createDocument('sheets','rich-profile'),sh=d.sheets[0]!;
 const red={kind:'rgba' as const,red:180,green:0,blue:40,alpha:255};
 sh.cells['1:1']={value:{kind:'string',value:'甲😀乙尾'},richText:[{text:'甲😀',style:{bold:true,color:red}},{text:'乙尾',style:{italic:true}}]};
 sh.cells['1:2']={value:{kind:'formula',source:'=LEN(A1)'}};
 const b=writeSheetsXlsx(d),source=openSheetsXlsx(b,d.documentId,d.title);
 assert.deepEqual(source.document.sheets[0]!.cells['1:1']!.richText,sh.cells['1:1']!.richText);
 assert.deepEqual(saveSheetsXlsx(b,source.document),b);
 const s=new DocumentSession(source.document),model=()=>s.snapshot().document as SheetsDocument;
 const edited=cellFromInput('甲😀新乙尾',{type:'text'},model().sheets[0]!.cells['1:1']);
 assert.deepEqual(edited.richText,[{text:'甲😀新',style:{bold:true,color:red}},{text:'乙尾',style:{italic:true}}]);
 s.execute(setCells(s.snapshot(),sh.id,[{row:1,column:1,cell:edited}]));
 assert.deepEqual(new WorkbookCalculator().calculate(model()).get(sh.id,1,2).value,{kind:'number',value:5});
 s.execute(formatCells(s.snapshot(),sh.id,{row:1,column:1,rowCount:1,columnCount:1},{style:{underline:true}}));
 assert.ok(model().sheets[0]!.cells['1:1']!.richText!.every(r=>r.style?.underline));
 s.execute(copyCells(s.snapshot(),sh.id,{row:1,column:1,rowCount:1,columnCount:1},sh.id,2,1));
 const out=saveSheetsXlsx(b,model()),back=openSheetsXlsx(out,d.documentId,d.title);
 assert.deepEqual(back.document.sheets[0]!.cells['2:1']!.richText,model().sheets[0]!.cells['1:1']!.richText);
 s.execute(clearCells(s.snapshot(),sh.id,{row:1,column:1,rowCount:1,columnCount:1}));
 assert.equal(model().sheets[0]!.cells['1:1']?.richText,undefined);s.undo();
 assert.throws(()=>writeLegacyXls(model()),OfficeError);
 assert.equal(cellFromInput('42',{},edited).richText,undefined);
 const bad=createDocument('sheets','bad-rich');bad.sheets[0]!.cells['1:1']={value:{kind:'string',value:'A'},richText:[{text:'B'}]};assert.throws(()=>new DocumentSession(bad),OfficeError);
});

test('M7.1 source named styles remain linked across value and direct-format edits, copy and reset',async()=>{
 const {applyNamedCellStyle,formatCells,copyCells}=await import('@office-ex/sheets');
 const d=createDocument('sheets','named-profile'),sh=d.sheets[0]!;
 d.namedCellStyles=[{name:'Report',style:{fontFamily:'Arial',fontSize:152400,bold:true,fill:{kind:'rgba',red:255,green:230,blue:160,alpha:255}},numberFormat:'0.00'}];
 sh.cells['1:1']={value:{kind:'number',value:12}};
 const s=new DocumentSession(d);s.execute(applyNamedCellStyle(s.snapshot(),sh.id,{row:1,column:1,rowCount:1,columnCount:1},'Report'));
 const b=writeSheetsXlsx(s.snapshot().document as SheetsDocument),source=openSheetsXlsx(b,d.documentId,d.title);
 assert.equal(source.document.sheets[0]!.cells['1:1']!.namedStyle,'Report');
 assert.deepEqual(saveSheetsXlsx(b,source.document),b);
 const n=new DocumentSession(source.document),m=()=>n.snapshot().document as SheetsDocument;
 n.execute(formatCells(n.snapshot(),sh.id,{row:1,column:1,rowCount:1,columnCount:1},{style:{italic:true}}));
 n.execute(copyCells(n.snapshot(),sh.id,{row:1,column:1,rowCount:1,columnCount:1},sh.id,2,1));
 const out=saveSheetsXlsx(b,m()),back=openSheetsXlsx(out,d.documentId,d.title);
 assert.deepEqual(back.document.namedCellStyles,source.document.namedCellStyles);
 assert.equal(back.document.sheets[0]!.cells['2:1']!.namedStyle,'Report');
 assert.equal(back.document.sheets[0]!.cells['2:1']!.style?.italic,true);
 n.execute(applyNamedCellStyle(n.snapshot(),sh.id,{row:2,column:1,rowCount:1,columnCount:1},'Normal'));
 assert.equal(m().sheets[0]!.cells['2:1']!.namedStyle,undefined);n.undo();
 assert.throws(()=>writeLegacyXls(m()),OfficeError);
});

test('M7.1 chart anchors and source vectors rebase atomically during sheet row/column edits',async()=>{
 const {addSheetChart,insertRows,deleteRows,insertColumns,deleteColumns}=await import('@office-ex/sheets');
 const d=createDocument('sheets','chart-shift'),sh=d.sheets[0]!;
 sh.cells['1:1']={value:{kind:'string',value:'Category'}};sh.cells['1:2']={value:{kind:'string',value:'Values'}};
 for(let r=2;r<=5;r++)sh.cells[r+':2']={value:{kind:'number',value:r}};
 const s=new DocumentSession(d),model=()=>s.snapshot().document as SheetsDocument;
 for(const kind of ['column','line','pie'] as const)s.execute(addSheetChart(s.snapshot(),sh.id,{row:1,column:1,rowCount:5,columnCount:2},kind));
 const original=model(),charts=original.sheets[0]!.charts!;
 assert.ok(charts[0]!.anchor.row<charts[1]!.anchor.row&&charts[1]!.anchor.row<charts[2]!.anchor.row);
 const bytes=writeSheetsXlsx(original);
 s.execute(insertRows(s.snapshot(),sh.id,3,2));
 assert.equal(model().sheets[0]!.charts![0]!.categories.rowCount,6);
 s.execute(insertColumns(s.snapshot(),sh.id,1,2));
 const chart=model().sheets[0]!.charts![0]!;
 assert.equal(chart.categories.column,3);assert.equal(chart.series[0]!.values.column,4);assert.equal(chart.series[0]!.label!.column,4);
 const reopened=openSheetsXlsx(saveSheetsXlsx(bytes,model()),d.documentId,d.title);
 assert.deepEqual(reopened.document.sheets[0]!.charts,model().sheets[0]!.charts);
 assert.throws(()=>deleteColumns(s.snapshot(),sh.id,4),OfficeError);
 assert.throws(()=>deleteRows(s.snapshot(),sh.id,1),OfficeError);
 s.execute(deleteRows(s.snapshot(),sh.id,3,2));assert.equal(model().sheets[0]!.charts![0]!.categories.rowCount,4);
 s.undo();s.undo();s.undo();assert.deepEqual(model(),original);
});
