// SPDX-License-Identifier: MIT
// Reduced, synthetic reproductions of structural issues; no private source text.
import {createDocument} from '@office-ex/core';
import {writeWriterDocx,writeSheetsXlsx,openOpc,writeZip} from '@office-ex/ooxml';
import {inspectForEditor} from '@office-ex/format-service';
import {writeFile} from 'node:fs/promises';
import assert from 'node:assert/strict';
const model=createDocument('writer','compat-probe','Compatibility probe');
const base=writeWriterDocx(model),dec=new TextDecoder(),enc=new TextEncoder();
const pkg=openOpc(base),main=pkg.mainPart,root=dec.decode(pkg.parts.get(main));
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',P='http://schemas.openxmlformats.org/package/2006/relationships';
const relPath='word/_rels/document.xml.rels',styleType='application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',numberType='application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
function fixture(change,partsChanged={}){
 const parts=new Map(pkg.parts),rels=parts.has(relPath)?dec.decode(parts.get(relPath)):'<Relationships xmlns="'+P+'"></Relationships>';
 const relList=[];let types='';
 const add=(id,type,path,mime,data)=>{
  relList.push('<Relationship Id="'+id+'" Type="'+R+'/'+type+'" Target="'+path+'"/>');
  types+='<Override PartName="/word/'+path+'" ContentType="'+mime+'"/>';parts.set('word/'+path,typeof data==='string'?enc.encode(data):data);
 };
 const doc=change(root,add);
 parts.set(main,enc.encode(doc));if(relList.length){parts.set(relPath,enc.encode(rels.replace('</Relationships>',relList.join('')+'</Relationships>')));
  const ct=dec.decode(parts.get('[Content_Types].xml'));parts.set('[Content_Types].xml',enc.encode(ct.replace('</Types>',types+'</Types>')));}
 for(const [name,data]of Object.entries(partsChanged))parts.set(name,enc.encode(data));
 return writeZip(parts);
}
function paragraph(body,props=''){return '<w:p>'+(props?'<w:pPr>'+props+'</w:pPr>':'')+body+'</w:p>';}
function replaceFirstBody(s,p){return s.replace(/(<w:body>)<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/,'$1'+p);}
const text='<w:r><w:t>Sample text</w:t></w:r>',emptyStyle='<w:styles xmlns:w="'+W+'"/>';
const cases={
 control:base,
 embeddedWorkbook:fixture((s,add)=>{
  add('rIdSyntheticWorkbook','package','embeddings/probe.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',writeSheetsXlsx(createDocument('sheets','embedded')));
  return replaceFirstBody(s,paragraph('<w:r><w:object><o:OLEObject xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="'+R+'" Type="Embed" ProgID="Excel.Sheet.12" DrawAspect="Icon" r:id="rIdSyntheticWorkbook"/></w:object></w:r>'));
 }),
 missingMargins:fixture(s=>s.replace(/<w:pgMar\b[^>]*\/>/,'')),
 undefinedStyle:fixture((s,add)=>{add('rIdSyntheticStyles','styles','styles.xml',styleType,emptyStyle);return replaceFirstBody(s,paragraph(text,'<w:pStyle w:val="2"/><w:outlineLvl w:val="1"/>'));}),
 minimalNumbering:fixture((s,add)=>{add('rIdSyntheticNumbering','numbering','numbering.xml',numberType,'<w:numbering xmlns:w="'+W+'"><w:abstractNum w:abstractNumId="1"><w:lvl><w:numFmt w:val="bullet"/><w:suff w:val="tab"/><w:lvlText w:val="•"/><w:rPr><w:color w:val="3370ff"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>');return replaceFirstBody(s,paragraph(text,'<w:numPr><w:numId w:val="1"/></w:numPr><w:ind w:left="453"/>'));}),
 strike:fixture(s=>replaceFirstBody(s,paragraph('<w:r><w:rPr><w:strike w:val="true"/></w:rPr><w:t>Sample text</w:t></w:r>'))),
 shading:fixture(s=>replaceFirstBody(s,paragraph('<w:r><w:rPr><w:shd w:fill="EFF0F1"/></w:rPr><w:t>Sample text</w:t></w:r>'))),
 bookmark:fixture(s=>replaceFirstBody(s,paragraph('<w:bookmarkStart w:id="1" w:name="SampleBookmark"/>'+text+'<w:bookmarkEnd w:id="1"/>')))
};
const results=[];
for(const [name,bytes]of Object.entries(cases)){
 const i=await inspectForEditor(bytes,name+'.docx','probe-'+name);
 results.push({name,editable:!i.report.readOnly,reason:i.semanticFailure?.detail??null,preview:!!i.officeTextPreview});
 console.log(JSON.stringify(results.at(-1)));
 if(['control','strike','shading','missingMargins','undefinedStyle','minimalNumbering'].includes(name))assert.equal(i.report.readOnly,false,name+': '+i.semanticFailure?.detail);
}
if(process.argv[2])await writeFile(process.argv[2],JSON.stringify({source:'synthetic reduced reproductions',results},null,2)+'\n');
