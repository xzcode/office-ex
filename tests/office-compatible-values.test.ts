// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import {createDocument} from '@office-ex/core';
import {OfficeError} from '@office-ex/contracts';
import {writeWriterDocx,writeSheetsXlsx,writeSlidesPptx,openOpc,writeZip,openWriterDocxCompatibility,openCompatibleOfficeValues,saveCompatibleOfficeValues,saveWriterDocxCompatibility} from '@office-ex/ooxml';
import {inspectForEditor,prepareCompatibleOfficeSave} from '@office-ex/format-service';
import {CompatibilitySession} from '../packages/ui/src/compatibility-workspace.js';
const enc=new TextEncoder(),dec=new TextDecoder();
function alter(bytes:Uint8Array,path:string,transform:(s:string)=>string){
 const p=openOpc(bytes),s=dec.decode(p.parts.get(path)!);const result=transform(s);assert.notEqual(result,s);
 p.parts.set(path,enc.encode(result));return writeZip(p.parts);
}
const request=(sourceSha256:string,edits:{id:number;text:string}[])=>({documentId:'fixture',sourceSha256,edits});
const error=(code:string)=>(e:unknown)=>e instanceof OfficeError&&e.code===code;
function sameParts(before:Uint8Array,after:Uint8Array,except:string[]){
 const a=openOpc(before),b=openOpc(after);assert.deepEqual([...a.parts.keys()].sort(),[...b.parts.keys()].sort());
 for(const [path,v]of a.parts)if(!except.includes(path))assert.deepEqual(b.parts.get(path),v,path);
}
function writer(){
 const d=createDocument('writer','fixture');const p=d.sections[0]!.blocks[0]!;if(p.type!=='paragraph')throw Error();
 p.runs=[{id:'r',type:'text',text:'Original text'}];return alter(writeWriterDocx(d),'word/document.xml',s=>s.replace('<w:rPr>','<w:rPr><w:strike/>').replace('<w:p>','<w:p><w:bookmarkStart w:id="7" w:name="keep"/>').replace('</w:p>','<w:bookmarkEnd w:id="7"/></w:p>'));
}
function sheets(){
 const d=createDocument('sheets','fixture'),s=d.sheets[0]!;
 s.cells['1:1']={value:{kind:'number',value:20}};s.cells['2:1']={value:{kind:'string',value:'Plain text'}};
 return alter(writeSheetsXlsx(d),'xl/worksheets/sheet1.xml',s=>s.replace('</worksheet>','<sheetProtection sheet="0"/></worksheet>'));
}
function slides(){
 const d=createDocument('slides','fixture');d.slides[0]!.shapes=[{id:'t',type:'text',geometry:{x:914400,y:914400,width:1828800,height:914400,rotation:0},runs:[{id:'r',type:'text',text:'Original slide'}]}];
 return writeSlidesPptx(d);
}
test('M7.6 Word fallback preserves strike and bookmark structure, supports empty and whitespace edits',async()=>{
 const src=writer(),i=await inspectForEditor(src,'sample.docx','fixture');
 assert.equal(i.report.readOnly,true);assert.ok(i.compatibilityWriter,i.semanticFailure?.detail);
 const v=i.compatibilityWriter;assert.equal(v.editableSegments,1);
 const s=new CompatibilitySession(v);s.set(0,'Changed text');assert.equal(s.dirty,true);s.undo();assert.equal(s.get(0),'Original text');assert.equal(s.dirty,false);s.redo();
 const saved=prepareCompatibleOfficeSave('docx',s.bytes('fixture'),'fixture',src);s.markSaved();assert.equal(s.dirty,false);
 const xml=dec.decode(openOpc(saved).parts.get('word/document.xml')!);assert.match(xml,/Changed text/);assert.match(xml,/bookmarkStart/);assert.match(xml,/w:strike/);
 sameParts(src,saved,['word/document.xml']);
 for(const text of ['', ' 前后空格 & < > ']){
  const out=saveWriterDocxCompatibility(src,'fixture',request(v.sourceSha256,[{id:0,text}]));
  const back=openWriterDocxCompatibility(out);assert.equal(back.blocks[0]?.kind,'paragraph');
  const p=back.blocks[0];if(p?.kind==='paragraph')assert.equal(p.text,text);
 }
 assert.deepEqual(saveWriterDocxCompatibility(src,'fixture',request(v.sourceSha256,[])),src);
 assert.throws(()=>saveWriterDocxCompatibility(src,'fixture',request('0'.repeat(64),[{id:0,text:'bad'}])),error('conflict'));
 assert.throws(()=>saveWriterDocxCompatibility(src,'fixture',request(v.sourceSha256,[{id:999,text:'bad'}])),error('invalid'));
 const protectedDoc=alter(src,'word/document.xml',s=>s.replace('<w:rPr>','<w:rPr><w:vanish/>'));
 assert.throws(()=>openWriterDocxCompatibility(protectedDoc),error('unsupported'));
});
test('M7.6 XLSX existing constants change without rebuilding styles, other cells or workbook',()=>{
 const d=createDocument('sheets','fixture'),s=d.sheets[0]!;
 s.cells['1:1']={value:{kind:'number',value:20}};s.cells['2:1']={value:{kind:'string',value:'Plain text'}};
 const source=writeSheetsXlsx(d),v=openCompatibleOfficeValues(source,'xlsx');assert.equal(v.editableSegments,2);
 const id=v.pages[0]!.fields.find(f=>f.label==='A1')!.id;
 const out=saveCompatibleOfficeValues(source,'xlsx','fixture',request(v.sourceSha256,[{id,text:'99.25'}]));
 assert.equal(openCompatibleOfficeValues(out,'xlsx').pages[0]!.fields.find(f=>f.label==='A1')!.text,'99.25');
 sameParts(source,out,['xl/worksheets/sheet1.xml']);
 assert.throws(()=>saveCompatibleOfficeValues(source,'xlsx','fixture',request(v.sourceSha256,[{id,text:'=SUM(A2)'}])),error('invalid'));
 assert.throws(()=>openCompatibleOfficeValues(sheets(),'xlsx'),error('unsupported'));
 s.cells['3:1']={value:{kind:'formula',source:'=A1+1'}};assert.throws(()=>openCompatibleOfficeValues(writeSheetsXlsx(d),'xlsx'),error('unsupported'));
});
test('M7.6 PPTX fallback patches only selected slide text, keeping all other source parts',()=>{
 const src=slides(),v=openCompatibleOfficeValues(src,'pptx');assert.equal(v.editableSegments,1);
 const out=saveCompatibleOfficeValues(src,'pptx','fixture',request(v.sourceSha256,[{id:0,text:'Edited slide'}]));
 assert.equal(openCompatibleOfficeValues(out,'pptx').pages[0]!.fields[0]!.text,'Edited slide');
 sameParts(src,out,['ppt/slides/slide1.xml']);
 assert.throws(()=>saveCompatibleOfficeValues(src,'pptx','fixture',request(v.sourceSha256,[{id:0,text:'a'},{id:0,text:'b'}])),error('invalid'));
});
