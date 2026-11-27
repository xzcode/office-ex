import test from 'node:test';
import assert from 'node:assert/strict';
import {createDocument,type SheetsDocument} from '@office-ex/core';
import {openOpc,openSheetsXlsx,writeSheetsXlsx,saveSheetsXlsx,previewSheetsXlsx,writeZip} from '@office-ex/ooxml';
import {inspectForEditor,prepareSheetsSave,serializeDraft} from '@office-ex/format-service';
const enc=(v:string)=>new TextEncoder().encode(v),dec=(b:Uint8Array)=>new TextDecoder().decode(b);
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',EP='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',VT='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
function book():SheetsDocument{
 const d=createDocument('sheets','compat','Fixture');d.sheets[0]!.name='Sales';d.sheets[0]!.cells={'1:1':{value:{kind:'string',value:'中文数据'}},'2:1':{value:{kind:'number',value:12}},'2:2':{value:{kind:'formula',source:'=A2*2'}}};
 d.sheets.push({id:'next-sheet',name:'Other',cells:{'1:1':{value:{kind:'boolean',value:true}}},rows:{},columns:{},merges:[]});return d;
}
function withMetadata(input:Uint8Array,names=['Sales','Other']):Uint8Array{
 const p=new Map(openOpc(input).parts);
 p.set('_rels/.rels',enc(dec(p.get('_rels/.rels')!).replace('</Relationships>','<Relationship Id="rIdProps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>')));
 p.set('[Content_Types].xml',enc(dec(p.get('[Content_Types].xml')!).replace('</Types>','<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>')));
 p.set('docProps/app.xml',enc('<Properties xmlns="'+EP+'" xmlns:vt="'+VT+'"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>'+names.length+'</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="'+names.length+'" baseType="lpstr">'+names.map(n=>'<vt:lpstr>'+n+'</vt:lpstr>').join('')+'</vt:vector></TitlesOfParts><AppVersion>16.0300</AppVersion></Properties>'));
 return writeZip(p);
}
test('XLSX ordinary worksheet-title metadata: edit native values and update only derived titles after structural changes',async()=>{
 const source=withMetadata(writeSheetsXlsx(book())),initial=openSheetsXlsx(source,'compat','Fixture');
 assert.equal((await inspectForEditor(source,'book.xlsx','compat')).report.readOnly,false);
 assert.deepEqual(saveSheetsXlsx(source,initial.document),source);
 initial.document.sheets[0]!.cells['2:1']!.value={kind:'number',value:21};
 const edited=saveSheetsXlsx(source,initial.document);
 assert.deepEqual(openOpc(edited).parts.get('docProps/app.xml'),openOpc(source).parts.get('docProps/app.xml'));
 assert.deepEqual(openSheetsXlsx(edited,'compat').cachedValues.find(([k])=>k==='sheet-1!2:2')?.[1],{kind:'number',value:42});
 initial.document.sheets.reverse();initial.document.sheets[1]!.name='Sales revised';
 initial.document.sheets.push({id:'third',name:'New & More',cells:{},rows:{},columns:{},merges:[]});
 const output=saveSheetsXlsx(source,initial.document);
 const xml=dec(openOpc(output).parts.get('docProps/app.xml')!);
 assert.match(xml,/<vt:i4>3<\/vt:i4>/);assert.match(xml,/<vt:lpstr>Other<\/vt:lpstr><vt:lpstr>Sales revised<\/vt:lpstr><vt:lpstr>New &amp; More<\/vt:lpstr>/);
 assert.deepEqual(openSheetsXlsx(output,'compat').document.sheets.map(s=>s.name),['Other','Sales revised','New & More']);
});
test('XLSX unknown or stale sheet-title metadata remains protected with readable values and actual diagnostic',async()=>{
 const bytes=withMetadata(writeSheetsXlsx(book()),['Wrong','Other']);
 const inspection=await inspectForEditor(bytes,'source.xlsx','compat');
 assert.equal(inspection.report.readOnly,true);assert.equal(inspection.editableSheets,undefined);
 assert.match(inspection.semanticFailure?.detail??'',/Sheet-title metadata does not match/);
 assert.equal(inspection.xlsxValuePreview?.sheets[0]?.cells['1:1']?.value,'中文数据');
 assert.ok(inspection.report.diagnostics.some(d=>d.code==='XLSX_VALUE_PREVIEW'));
});
test('XLSX protected fallback extracts cached formulas, rich text, hidden row/column/sheet exclusion without calculation',async()=>{
 const parts=new Map(openOpc(writeSheetsXlsx(book())).parts);
 const ws=dec(parts.get('xl/worksheets/sheet1.xml')!);
 const source=ws.replace(/<sheetData>[\s\S]*?<\/sheetData>/,'<cols><col min="4" max="4" hidden="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><t>中文</t></r><r><t>&lt;script&gt; literal</t></r></is></c><c r="B1"><f>WEBSERVICE("https://invalid.example")</f><v>57</v></c><c r="C1"><f t="shared" si="0"/></c><c r="D1"><v>12345</v></c></row><row r="2" hidden="1"><c r="A2"><v>999</v></c></row><row r="80"><c r="Z80" t="b"><v>1</v></c></row></sheetData>');
 parts.set('xl/worksheets/sheet1.xml',enc(source));
 const wb=dec(parts.get('xl/workbook.xml')!).replace('name="Other"','name="Other" state="hidden"');parts.set('xl/workbook.xml',enc(wb));
 const bytes=writeZip(parts),before=bytes.slice(),view=previewSheetsXlsx(bytes);
 assert.equal(view.sheets.length,1);const a=view.sheets[0]!;
 assert.equal(a.cells['1:1']?.value,'中文<script> literal');assert.deepEqual(a.cells['1:2'],{value:'57',kind:'formula-cache',formula:'=WEBSERVICE("https://invalid.example")'});
 assert.deepEqual(a.cells['1:3'],{value:null,kind:'formula-uncached'});assert.equal(a.cells['1:4'],undefined);assert.equal(a.cells['2:1'],undefined);
 assert.deepEqual(a.cells['80:26'],{value:'TRUE',kind:'value'});assert.equal(a.lastRow,80);assert.equal(a.lastColumn,26);
 const inspected=await inspectForEditor(bytes,'preview.xlsx','compat');
 assert.equal(inspected.report.readOnly,true);assert.equal(inspected.editableSheets,undefined);assert.ok(inspected.semanticFailure);assert.deepEqual(bytes,before);
 assert.throws(()=>prepareSheetsSave('xlsx',serializeDraft(book()),'compat',bytes));
});
