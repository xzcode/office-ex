import {readFileSync,writeFileSync} from 'node:fs';
import assert from 'node:assert/strict';
import {createDocument,DocumentSession,projectParagraph} from '@office-ex/core';
import {replaceParagraphRange} from '@office-ex/writer';
import {writeLegacyDoc,openLegacyDoc,saveLegacyDoc,inspectForEditor} from '@office-ex/format-service';
import {writeWriterDocx,openWriterDocx,saveWriterDocx} from '@office-ex/ooxml';
const [mode,input,output,extra]=process.argv.slice(2);
if(mode==='ordinary'){
  const bytes=readFileSync(input),source=openWriterDocx(bytes,'ordinary','Ordinary');
  assert.deepEqual(Uint8Array.from(saveWriterDocx(source,source.document)),Uint8Array.from(bytes));
  const paragraphs=source.document.sections[0].blocks.filter(b=>b.type==='paragraph');
  assert.equal(projectParagraph(paragraphs[0]),'Independent 中文');
  const s=new DocumentSession(source.document,{saved:true});
  s.execute(replaceParagraphRange(s.snapshot(),paragraphs[0].id,0,11,'Office EX','ordinary-edit'));
  writeFileSync(output,saveWriterDocx(source,s.snapshot().document));
}else if(mode==='legacy'){
  const model=createDocument('writer','legacy-fixture','Legacy fixture');
  model.defaultTextStyle={fontFamily:'Arial',fontSize:152400,language:'zh-CN'};
  model.sections[0].blocks=[
    {id:'p1',type:'paragraph',style:{alignment:'center'},runs:[{id:'r1',type:'text',text:'Office EX 中文',style:{bold:true}},{id:'r2',type:'text',text:' direct style',style:{italic:true}}]},
    {id:'p2',type:'paragraph',runs:[{id:'r3',type:'text',text:'Second paragraph'}]}
  ];
  const original=writeLegacyDoc(model),source=openLegacyDoc(original,'legacy-fixture','Legacy fixture');
  assert.deepEqual(saveLegacyDoc(source,source.document),original);
  writeFileSync(input,original);
  const s=new DocumentSession(source.document,{saved:true});s.execute(replaceParagraphRange(s.snapshot(),'p-0',0,6,'Writer','edit'));
  writeFileSync(output,saveLegacyDoc(source,s.snapshot().document));
  const docx=writeWriterDocx(s.snapshot().document);writeFileSync(extra,docx);
  assert.equal(openWriterDocx(docx,'modern','Modern').document.sections[0].blocks.length,2);
}else throw Error('Unknown fixture mode');
