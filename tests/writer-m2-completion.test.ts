// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import {createDocument,DocumentSession,makeImageResource,projectParagraph,type WriterDocument} from '@office-ex/core';
import {replaceParagraphRange,formatParagraphRange,insertImageAt,splitWriterSectionAfter,copyWriterFragment,pasteWriterFragment,locateParagraph} from '@office-ex/writer';
import {writeLegacyDoc,openLegacyDoc,saveLegacyDoc,inspectLegacyDocText,inspectForEditor,prepareWriterDocSave,serializeDraft} from '@office-ex/format-service';
import {readCfb,writeCfb} from '@office-ex/cfb';
import {writeWriterDocx,openWriterDocx} from '@office-ex/ooxml';
import {WriterPaginationEngine,layoutWriter,type FontMetrics} from '@office-ex/render';
import {graphemeBoundaries,OfficeError} from '@office-ex/contracts';
import {PNG} from './writer-image-fixture.js';
const fails=(code:string)=>(e:unknown)=>e instanceof OfficeError&&e.code===code;
const range=(id:string,start:number,end:number)=>({anchor:{paragraphId:id,offset:start},focus:{paragraphId:id,offset:end}});
const doc=(s:DocumentSession)=>s.snapshot().document as WriterDocument;
function fixture(){
  const d=createDocument('writer','legacy','Legacy');
  d.defaultTextStyle={fontFamily:'Arial',fontSize:152400,language:'zh-CN'};
  d.sections[0]!.blocks=[
    {type:'paragraph',id:'p1',style:{alignment:'center',spaceAfter:25400},runs:[{id:'r1',type:'text',text:'Office EX 中文 😀',style:{bold:true,color:{kind:'rgba',red:20,green:60,blue:100,alpha:255}}},{id:'r2',type:'text',text:' test',style:{italic:true}}]},
    {type:'paragraph',id:'p2',runs:[{id:'r3',type:'text',text:'Second paragraph'}]}
  ];
  return d;
}
test('LEG-02 DOC Unicode and direct CHPX/PAPX, exact copy and edited reopen',async()=>{
  const bytes=writeLegacyDoc(fixture()),source=openLegacyDoc(bytes,'legacy','Legacy');
  const first=source.document.sections[0]!.blocks[0]!;
  assert.equal(first.type,'paragraph');
  if(first.type!=='paragraph')throw Error();
  assert.equal(projectParagraph(first),'Office EX 中文 😀 test');
  assert.equal(first.style?.alignment,'center');
  const run=first.runs[0]!;assert.equal(run.type,'text');if(run.type!=='text')throw Error();
  assert.equal(run.style?.bold,true);assert.deepEqual(run.style?.color,{kind:'rgba',red:20,green:60,blue:100,alpha:255});
  assert.equal(inspectLegacyDocText(bytes).text,'Office EX 中文 😀 test\rSecond paragraph\r');
  assert.deepEqual(saveLegacyDoc(source,source.document),bytes);
  const opened=await inspectForEditor(bytes,'sample.doc','legacy');assert.equal(opened.editableWriterFormat,'doc');assert.equal(opened.report.readOnly,false);
  const s=new DocumentSession(source.document,{saved:true});s.execute(replaceParagraphRange(s.snapshot(),'p-0',0,6,'Writer','edit'));
  const edited=prepareWriterDocSave(serializeDraft(s.snapshot().document),'legacy',bytes);
  assert.match(inspectLegacyDocText(edited).text,/^Writer EX 中文/);
});
test('LEG-02 DOCX and DOC convert through one supported model, not renamed extensions',()=>{
  const first=writeLegacyDoc(fixture()),src=openLegacyDoc(first,'converted','Legacy');
  const dx=writeWriterDocx(src.document),modern=openWriterDocx(dx,'converted','Modern');
  const round=writeLegacyDoc(modern.document);assert.equal(inspectLegacyDocText(round).text,inspectLegacyDocText(first).text);
  assert.equal(round[0],0xd0);assert.equal(dx[0],0x50);
});
test('LEG-02 noncanonical streams and malformed text do not become editable',async()=>{
  const bytes=writeLegacyDoc(fixture()),cfb=readCfb(bytes);
  const opaque=writeCfb({...cfb,entries:[...cfb.entries,{path:['Opaque'],type:'stream',data:new Uint8Array([1,2,3])}]});
  assert.throws(()=>openLegacyDoc(opaque,'legacy','Legacy'),fails('unsupported'));
  const opened=await inspectForEditor(opaque,'opaque.doc','legacy');assert.equal(opened.report.readOnly,true);assert.ok(opened.legacyDocText?.text.includes('中文'));
  assert.throws(()=>prepareWriterDocSave(serializeDraft(fixture()),'legacy',opaque),fails('unsupported'));
  const d=fixture();d.sections.push({...d.sections[0]!,id:'other',blocks:[{id:'newp',type:'paragraph',runs:[{id:'newr',type:'text',text:'Other'}]}]});
  assert.throws(()=>writeLegacyDoc(d),fails('unsupported'));
});
test('WRT-03 rich clipboard copies style and picture bytes between documents and undoes atomically',()=>{
  const source=new DocumentSession(fixture());
  source.execute(insertImageAt(source.snapshot(),'p1',6,6,makeImageResource(PNG),'clip-img'));
  const p=locateParagraph(source.snapshot(),'p1').paragraph;
  const copied=copyWriterFragment(source.snapshot(),range('p1',0,projectParagraph(p).length));
  const target=new DocumentSession(createDocument('writer','paste-target'),{saved:true});
  const id=doc(target).sections[0]!.blocks[0]!.id,old=target.snapshot().document;
  const edit=pasteWriterFragment(target.snapshot(),range(id,0,0),copied.rich,'paste');
  target.execute(edit.command);
  const result=locateParagraph(target.snapshot(),id).paragraph;
  assert.equal(projectParagraph(result),projectParagraph(p));
  assert.equal(Object.keys(doc(target).resources).length,1);
  assert.ok(result.runs.some(r=>r.type==='text'&&r.style?.bold));
  target.undo();assert.deepEqual(target.snapshot().document,old);
  const corrupt=JSON.parse(copied.rich);corrupt.version=2;
  assert.throws(()=>pasteWriterFragment(target.snapshot(),range(id,0,0),JSON.stringify(corrupt),'bad'),fails('invalid'));
});
test('WRT-02 incremental pagination retains unchanged section pages and invalidates edited geometry',()=>{
  const s=new DocumentSession(fixture());
  s.execute(splitWriterSectionAfter(s.snapshot(),'p1','sec').command);
  const font:FontMetrics={environmentId:'m2-cache',measure(text,request){
    const b=graphemeBoundaries(text),clusters=b.slice(0,-1).map((start,i)=>({start,end:b[i+1]!,left:i*20000,right:(i+1)*20000}));
    return {advance:clusters.length*20000,ascent:100000,descent:40000,actualFamily:request.family,clusters,clusterPrecision:'exact'};
  }};
  const engine=new WriterPaginationEngine(),first=engine.layout(s.snapshot(),font);
  assert.equal(first.pages.length,2);
  s.execute(replaceParagraphRange(s.snapshot(),'p2',0,6,'Changed','update'));
  const next=engine.layout(s.snapshot(),font);
  assert.equal(engine.stats.sectionsReused,1);assert.equal(engine.stats.sectionsComputed,1);
  assert.equal(next.pages[0],first.pages[0]);assert.equal(next.revision,s.revision);
  assert.deepEqual(next.pages,layoutWriter(s.snapshot(),font).pages);
  assert.throws(()=>engine.layout(s.snapshot(),font,{maxPages:1}),fails('limit'));
});
