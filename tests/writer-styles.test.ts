// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import {OfficeError} from '@office-ex/contracts';
import {DocumentSession,resolveTextStyle,cloneData,type WriterDocument} from '@office-ex/core';
import {openWriterDocx,saveWriterDocx,openOpc} from '@office-ex/ooxml';
import {paragraphText,formatParagraphRange,setParagraphNamedStyle,setCharacterNamedStyle,replaceParagraphRange} from '@office-ex/writer';
import {inspectForEditor,prepareWriterDocxSave,serializeDraft} from '@office-ex/format-service';
import {styledFixture,alterPart} from './writer-styled-fixture.js';
const fails=(code:string)=>(e:unknown)=>e instanceof OfficeError&&e.code===code;
const rgb={kind:'rgba' as const,red:51,green:102,blue:153,alpha:255};
function paragraph(doc:WriterDocument|ReturnType<typeof openWriterDocx>['document'],index=0){
  const p=doc.sections[0]!.blocks[index]!;if(p.type!=='paragraph')throw Error('Expected paragraph');return p;
}
test('WRT-02/WRT-04 imports named styles, resolves theme colors/fonts and applies toggle inheritance in order',()=>{
  const bytes=styledFixture(),source=openWriterDocx(bytes,'style-test','Fixture'),doc=source.document;
  assert.equal(doc.defaultTextStyle.fontFamily,'Arial');assert.equal(doc.defaultTextStyle.fontSize,139700);
  assert.equal(doc.defaultTextStyle.spaceAfter,76200);
  assert.equal(paragraph(doc).styleId,'Heading');assert.equal(paragraph(doc,1).styleId,'Normal');
  const run=paragraph(doc).runs[0]!;if(run.type!=='text')throw Error();
  assert.equal(run.styleId,'Accent');assert.equal(doc.styles.Heading?.basedOn,'Normal');
  assert.deepEqual(doc.styles.Normal?.toggles,['bold']);assert.equal(doc.styles.UnusedList,undefined);assert.equal(doc.styles.UnusableChild,undefined);
  const para=resolveTextStyle(doc.defaultTextStyle,doc.styles,['Heading']).style;
  assert.equal(para.bold,false);assert.equal(para.alignment,'center');assert.equal(para.fontSize,228600);
  const effective=resolveTextStyle(para,doc.styles,['Accent']).style;
  assert.equal(effective.bold,true);assert.deepEqual(effective.color,rgb);
  assert.deepEqual(saveWriterDocx(source,doc),bytes);
});
test('WRT-04 edited styled DOCX preserves every non-main part byte-for-byte and reopens named references and direct overrides',async()=>{
  const bytes=styledFixture(),inspection=await inspectForEditor(bytes,'style.docx','style-test');
  assert.ok(inspection.editableWriter);
  const session=new DocumentSession(inspection.editableWriter,{saved:true});
  const id=paragraph(inspection.editableWriter).id;
  session.execute(replaceParagraphRange(session.snapshot(),id,0,6,'Edited','edit'));
  session.execute(formatParagraphRange(session.snapshot(),id,0,6,{fontFamily:'Times New Roman',fontSize:177800,color:{kind:'rgba',red:170,green:34,blue:17,alpha:255}},'fmt'));
  session.execute(setParagraphNamedStyle(session.snapshot(),id,'Normal'));
  const second=paragraph(inspection.editableWriter,1).id;
  session.execute(setCharacterNamedStyle(session.snapshot(),second,0,6,'Accent','character'));
  const changed=prepareWriterDocxSave(serializeDraft(session.snapshot().document),'style-test',bytes);
  const a=openOpc(bytes),b=openOpc(changed);
  for(const [path,data] of a.parts)if(path!==a.mainPart)assert.deepEqual(b.parts.get(path),data,path);
  const next=openWriterDocx(changed,'style-test','Edited').document;
  assert.equal(paragraph(next).styleId,'Normal');const run=paragraph(next).runs[0]!;
  if(run.type!=='text')throw Error();
  assert.equal(run.text,'Edited');assert.equal(run.styleId,'Accent');assert.equal(run.style?.fontFamily,'Times New Roman');assert.equal(run.style.fontSize,177800);
  assert.equal(paragraph(next,1).runs[0]!.type,'text');
  const r2=paragraph(next,1).runs[0]!;if(r2.type!=='text')throw Error();assert.equal(r2.styleId,'Accent');
  assert.match(new TextDecoder().decode(b.parts.get(b.mainPart!)!),/w:header="720" w:footer="720"/);
  assert.deepEqual(next.styles,openWriterDocx(bytes,'style-test','Fixture').document.styles);
  for(let i=0;i<4;i++)session.undo();
  assert.equal(session.dirty,false);assert.deepEqual(prepareWriterDocxSave(serializeDraft(session.snapshot().document),'style-test',bytes),bytes);
});
test('WRT-04 unsupported referenced/default style, unsupported theme color or inheritance cycle stays read-only',async()=>{
  const bytes=styledFixture();
  const variants=[
    alterPart(bytes,'word/document.xml',s=>s.replace('w:pStyle w:val="Heading"','w:pStyle w:val="UnusableChild"')),
    alterPart(bytes,'word/styles.xml',s=>s.replace('w:styleId="Normal"><w:name','w:styleId="Normal"><w:basedOn w:val="Heading"/><w:name')),
    alterPart(bytes,'word/styles.xml',s=>s.replace('w:themeColor="accent1"','w:themeColor="accent1" w:themeTint="80"')),
    alterPart(bytes,'word/styles.xml',s=>s.replace('w:asciiTheme="minorAscii"','w:asciiTheme="unknown"'))
  ];
  for(const input of variants){
    assert.throws(()=>openWriterDocx(input,'style-test','Fixture'),fails('unsupported'));
    const view=await inspectForEditor(input,'style.docx','style-test');assert.equal(view.editableWriter,undefined);assert.equal(view.report.readOnly,true);
  }
});
test('WRT-04 source style definitions/defaults are immutable and style application validates type and existence',()=>{
  const source=openWriterDocx(styledFixture(),'style-test','Fixture');
  const edited=cloneData(source.document) as WriterDocument;
  edited.styles.Normal!.properties.fontSize=190500;
  assert.throws(()=>saveWriterDocx(source,edited),fails('unsupported'));
  const changedDefaults=cloneData(source.document) as WriterDocument;changedDefaults.defaultTextStyle.fontSize=190500;
  assert.throws(()=>saveWriterDocx(source,changedDefaults),fails('unsupported'));
  const session=new DocumentSession(source.document,{saved:true}),id=paragraph(source.document).id;
  assert.throws(()=>setParagraphNamedStyle(session.snapshot(),id,'Accent'),fails('unsupported'));
  assert.throws(()=>setCharacterNamedStyle(session.snapshot(),id,0,1,'Heading','op'),fails('unsupported'));
  assert.throws(()=>setParagraphNamedStyle(session.snapshot(),id,'UnusedList'),fails('unsupported'));
  assert.equal(paragraphText(session.snapshot(),id),'Styled 中文 😀');assert.equal(session.dirty,false);
});
