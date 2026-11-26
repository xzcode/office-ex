// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import {OfficeError,graphemeBoundaries} from '@office-ex/contracts';
import {createDocument,DocumentSession,cloneData,validateDocument,type WriterDocument} from '@office-ex/core';
import {setWriterPageStory,setWriterPageNumberStart,effectivePageStory,splitWriterSectionAfter,insertPageBreakAfter,replaceParagraphRange} from '@office-ex/writer';
import {layoutWriter,WriterPaginationEngine,type FontMetrics} from '@office-ex/render';
import {openWriterDocx,saveWriterDocx,writeWriterDocx,openOpc,writeZip} from '@office-ex/ooxml';
import {writeLegacyDoc,inspectForEditor} from '@office-ex/format-service';
const fails=(code:string)=>(e:unknown)=>e instanceof OfficeError&&e.code===code;
const font:FontMetrics={environmentId:'running-story-test',measure(text,request){
 const b=graphemeBoundaries(text);return {advance:(b.length-1)*50000,ascent:80000,descent:20000,actualFamily:request.family,clusterPrecision:'exact',clusters:b.slice(0,-1).map((start,i)=>({start,end:b[i+1]!,left:i*50000,right:(i+1)*50000}))};
}};
const model=(s:DocumentSession)=>s.snapshot().document as WriterDocument;
function fixture(){
 const doc=createDocument('writer','story-fixture');
 doc.defaultTextStyle={fontFamily:'sans-serif',fontSize:100000};
 const s=new DocumentSession(doc);
 s.execute(setWriterPageStory(s.snapshot(),'section-1','header',{text:'Report'},'header'));
 s.execute(setWriterPageStory(s.snapshot(),'section-1','footer',{text:'Page ',pageNumber:true,after:' / report',alignment:'center'},'footer'));
 s.execute(setWriterPageNumberStart(s.snapshot(),'section-1',5));
 return s;
}
test('M6.2 page stories are typed, validated, undoable and do not add body text',()=>{
 const s=fixture(),before=s.snapshot().document;
 const m=model(s),header=effectivePageStory(m,'section-1','header')!,footer=effectivePageStory(m,'section-1','footer')!;
 assert.equal(m.pageStories![header]!.text,'Report');assert.equal(m.pageStories![footer]!.pageNumber,true);
 assert.equal(m.sections[0]!.blocks.length,1);
 s.execute(setWriterPageStory(s.snapshot(),'section-1','footer',{text:'Changed',pageNumber:true,alignment:'end',pageNumberStart:12},'edit'));
 assert.equal(model(s).sections[0]!.pageNumberStart,12);s.undo();assert.deepEqual(s.snapshot().document,before);s.redo();
 const bad=cloneData(model(s));bad.sections[0]!.storyRefs={header:footer};assert.throws(()=>validateDocument(bad),fails('invalid'));
 const hidden=cloneData(model(s));hidden.pageStories![header]!.text='bad\nline';assert.throws(()=>validateDocument(hidden),fails('invalid'));
 assert.throws(()=>setWriterPageNumberStart(s.snapshot(),'section-1',0),fails('invalid'));
});
test('M6.2 section split inherits stories, detached edit leaves previous section untouched',()=>{
 const s=fixture(),h=effectivePageStory(model(s),'section-1','header');
 const split=splitWriterSectionAfter(s.snapshot(),'paragraph-1','sec');s.execute(split.command);
 assert.equal(model(s).sections[1]!.storyRefs,undefined);assert.equal(model(s).sections[1]!.pageNumberStart,undefined);
 assert.equal(effectivePageStory(model(s),split.sectionId,'header'),h);
 s.execute(setWriterPageStory(s.snapshot(),split.sectionId,'header',{text:'Second'},'second-h'));
 assert.notEqual(effectivePageStory(model(s),split.sectionId,'header'),h);
 assert.equal(model(s).pageStories![h!]!.text,'Report');
 s.undo();assert.equal(effectivePageStory(model(s),split.sectionId,'header'),h);
});
test('M6.2 layout computes actual PAGE fields across pages, inherited sections and restart',()=>{
 const s=fixture(),split=splitWriterSectionAfter(s.snapshot(),'paragraph-1','sec');s.execute(split.command);
 s.execute(setWriterPageNumberStart(s.snapshot(),split.sectionId,20));
 const initial=layoutWriter(s.snapshot(),font);
 assert.deepEqual(initial.pages.map(p=>p.displayPageNumber),[5,20]);
 assert.deepEqual(initial.pages.map(p=>p.runningStories!.find(b=>b.kind==='footer')!.text),['Page 5 / report','Page 20 / report']);
 for(const p of initial.pages){
  const h=p.runningStories!.find(b=>b.kind==='header')!,f=p.runningStories!.find(b=>b.kind==='footer')!;
  assert.ok(h.y+h.height<=p.margins.top);assert.ok(f.y>=p.height-p.margins.bottom);
  assert.ok(!p.lines.some(l=>l.paragraphId.startsWith('running-')));
 }
 const engine=new WriterPaginationEngine();engine.layout(s.snapshot(),font);
 s.execute(setWriterPageNumberStart(s.snapshot(),split.sectionId));
 const continuous=engine.layout(s.snapshot(),font);assert.deepEqual(continuous.pages.map(p=>p.displayPageNumber),[5,6]);
 // Body growth changes following linked pages without needing to rebuild their cached body.
 const firstParagraph=model(s).sections[0]!.blocks[0]!.id;
 // Insert a page-break before a terminal paragraph by adding another paragraph in the first section.
 const doc=cloneData(model(s));doc.sections[0]!.blocks.push({id:'extra-break',type:'page-break'},{id:'extra-p',type:'paragraph',runs:[{id:'extra-r',type:'text',text:'Extra page'}]});
 const next=new DocumentSession(doc);const shifted=layoutWriter(next.snapshot(),font);
 assert.deepEqual(shifted.pages.map(p=>p.displayPageNumber),[5,6,7]);
 assert.equal(shifted.pages[2]!.runningStories!.find(b=>b.kind==='footer')!.text,'Page 7 / report');
});
test('M6.2 new DOCX story relationships and decimal PAGE fields roundtrip safely',()=>{
 const s=fixture(),bytes=writeWriterDocx(model(s)),source=openWriterDocx(bytes,'story-fixture','Story');
 const first=source.document.sections[0]!,header=first.storyRefs!.header!,footer=first.storyRefs!.footer!;
 assert.equal(first.pageNumberStart,5);assert.equal(source.document.pageStories![footer]!.pageNumber,true);
 const opc=openOpc(bytes);const rels=opc.relationships.get(opc.mainPart!)!;
 assert.equal(rels.filter(r=>r.type.endsWith('/header')||r.type.endsWith('/footer')).length,2);
 assert.deepEqual(saveWriterDocx(source,source.document),bytes);
 const edit=new DocumentSession(source.document);
 edit.execute(setWriterPageStory(edit.snapshot(),first.id,'footer',{text:'Edited ',pageNumber:true},'footer-edit'));
 const out=saveWriterDocx(source,edit.snapshot().document),reopened=openWriterDocx(out,'story-fixture','Story');
 assert.equal(reopened.document.pageStories![footer]!.text,'Edited ');
 assert.equal(reopened.document.pageStories![header]!.text,'Report');
 const parts=openOpc(out);const headerPath=rels.find(r=>r.id===header)!.targetPart!;
 assert.deepEqual(parts.parts.get(headerPath),opc.parts.get(headerPath));
 assert.throws(()=>writeLegacyDoc(model(s)),fails('unsupported'));
});
test('M6.2 unsupported fields, first-page stories and malformed references remain protected',async()=>{
 const s=fixture(),bytes=writeWriterDocx(model(s)),opc=openOpc(bytes),dec=new TextDecoder(),enc=new TextEncoder();
 const footer=opc.relationships.get(opc.mainPart!)!.find(r=>r.type.endsWith('/footer'))!.targetPart!;
 for(const [path,mutate] of [[footer,(x:string)=>x.replace(' PAGE ',' DATE ')],[opc.mainPart!,(x:string)=>x.replace('w:type="default"','w:type="first"')],[opc.mainPart!,(x:string)=>x.replace('r:id="rIdOexStory_header"','r:id="missing"')]] as const){
  const parts=new Map(opc.parts),original=dec.decode(parts.get(path)!);const changed=mutate(original);assert.notEqual(changed,original);
  parts.set(path,enc.encode(changed));const blocked=writeZip(parts);
  assert.throws(()=>openWriterDocx(blocked,'protected','Protected'),fails('unsupported'));
  assert.equal((await inspectForEditor(blocked,'protected.docx','protected')).report.readOnly,true);
 }
});
