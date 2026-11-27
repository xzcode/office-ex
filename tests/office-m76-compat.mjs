// SPDX-License-Identifier: MIT
import assert from 'node:assert/strict';
import {readFile,writeFile} from 'node:fs/promises';
import {join} from 'node:path';
import {createDocument} from '@office-ex/core';
import {writeWriterDocx,writeSheetsXlsx,openOpc,writeZip,openWriterDocxCompatibility,openCompatibleOfficeValues} from '@office-ex/ooxml';
import {inspectForEditor,prepareCompatibleOfficeSave} from '@office-ex/format-service';
const dir=process.argv[2],enc=new TextEncoder(),dec=new TextDecoder();
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',PR='http://schemas.openxmlformats.org/package/2006/relationships';
const d=createDocument('writer','compat-docx');
d.sections[0].blocks[0].runs=[{id:'word-run',type:'text',text:'Original Word'}];
const p=openOpc(writeWriterDocx(d)),parts=new Map(p.parts),part='word/document.xml';
let xml=dec.decode(parts.get(part)).replace('<w:p>','<w:p><w:bookmarkStart w:id="1" w:name="Keep"/>').replace('<w:rPr>','<w:rPr><w:strike/>').replace('</w:p>','<w:bookmarkEnd w:id="1"/></w:p>');
// Genuine PNG preview and inert XLSX package. No OLE binary or executable.
const png=new Uint8Array(await readFile(join(dir,'preview.png')));
xml=xml.replace('<w:sectPr>','<w:p><w:r><w:object><v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="width:100pt;height:20pt"><v:imagedata xmlns:r="'+R+'" r:id="rPreview"/></v:shape><o:OLEObject xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="'+R+'" Type="Embed" ProgID="Excel.Sheet.12" DrawAspect="Icon" r:id="rEmbedded"/></w:object></w:r></w:p><w:sectPr>');
parts.set(part,enc.encode(xml));parts.set('word/media/preview.png',png);parts.set('word/embeddings/embedded.xlsx',writeSheetsXlsx(createDocument('sheets','embedded')));
parts.set('word/_rels/document.xml.rels',enc.encode('<Relationships xmlns="'+PR+'"><Relationship Id="rPreview" Type="'+R+'/image" Target="media/preview.png"/><Relationship Id="rEmbedded" Type="'+R+'/package" Target="embeddings/embedded.xlsx"/></Relationships>'));
parts.set('[Content_Types].xml',enc.encode(dec.decode(parts.get('[Content_Types].xml')).replace('</Types>','<Default Extension="png" ContentType="image/png"/><Override PartName="/word/embeddings/embedded.xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/></Types>')));
await writeFile(join(dir,'compat.docx'),writeZip(parts));
async function inspect(format){
 const bytes=await readFile(join(dir,'compat.'+format)),id='compat-'+format;
 const info=await inspectForEditor(bytes,'compat.'+format,id);
 const view=info.compatibilityWriter??info.compatibilityOffice;
 console.log(format,'full profile:',info.semanticFailure?.detail??'accepted','fallback:',view?.editableSegments??0);
 assert.equal(info.report.readOnly,true);assert.ok(view,format+' fallback not admitted');
 return {bytes,id,view};
}
async function save(format,source,edits){
 const envelope=enc.encode(JSON.stringify({documentId:source.id,sourceSha256:source.view.sourceSha256,edits}));
 const output=prepareCompatibleOfficeSave(format,envelope,source.id,source.bytes);
 await writeFile(join(dir,'compat-edited.'+format),output);
 const a=openOpc(source.bytes),b=openOpc(output);
 const except=format==='docx'?['word/document.xml']:format==='xlsx'?['xl/worksheets/sheet1.xml']:['ppt/slides/slide1.xml'];
 assert.deepEqual([...a.parts.keys()].sort(),[...b.parts.keys()].sort());for(const [name,data]of a.parts)if(!except.includes(name))assert.deepEqual(b.parts.get(name),data,name);
 return output;
}
const word=await inspect('docx');assert.equal(word.view.embeddedWorkbooks,1);assert.equal(word.view.images,1);assert.equal(Object.values(word.view.resources)[0].mimeType,'image/jpeg');
await save('docx',word,[{id:0,text:'Updated Word'}]);
const x=await inspect('xlsx'),fields=x.view.pages[0].fields;assert.equal(x.view.pages.length,1);assert.ok(!fields.some(f=>f.label==='A4'));
await save('xlsx',x,[{id:fields.find(f=>f.label==='A1').id,text:'99.25'},{id:fields.find(f=>f.label==='B1').id,text:'Updated cell'},{id:fields.find(f=>f.label==='A2').id,text:'0'}]);
const ppt=await inspect('pptx'),texts=ppt.view.pages[0].fields;
await save('pptx',ppt,[{id:texts.find(f=>f.text==='Original slide').id,text:'Updated slide'},{id:texts.find(f=>f.text==='Independent table cell').id,text:'Updated table cell'}]);
const formula=await readFile(join(dir,'formula-locked.xlsx'));assert.throws(()=>openCompatibleOfficeValues(formula,'xlsx'));
console.log('Native fallback edits and every unrelated uncompressed part verified');
