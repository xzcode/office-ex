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
test('M6.2 unsupported fields, invalid variants and malformed references remain protected',async()=>{
 const s=fixture(),bytes=writeWriterDocx(model(s)),opc=openOpc(bytes),dec=new TextDecoder(),enc=new TextEncoder();
 const footer=opc.relationships.get(opc.mainPart!)!.find(r=>r.type.endsWith('/footer'))!.targetPart!;
 for(const [path,mutate] of [[footer,(x:string)=>x.replace(' PAGE ',' DATE ')],[opc.mainPart!,(x:string)=>x.replace('w:type="default"','w:type="unknown"')],[opc.mainPart!,(x:string)=>x.replace('r:id="rIdOexStory_header"','r:id="missing"')]] as const){
  const parts=new Map(opc.parts),original=dec.decode(parts.get(path)!);const changed=mutate(original);assert.notEqual(changed,original);
  parts.set(path,enc.encode(changed));const blocked=writeZip(parts);
  assert.throws(()=>openWriterDocx(blocked,'protected','Protected'),fails('unsupported'));
  assert.equal((await inspectForEditor(blocked,'protected.docx','protected')).report.readOnly,true);
 }
});

function variantFixture(){
 const s=fixture();
 const text=(value:string,alignment:'start'|'center'|'end'='start')=>({text:value,alignment});
 const para=(value:string)=>({id:value+'-p',type:'paragraph' as const,runs:[{id:value+'-r',type:'text' as const,text:value}]});
 const doc=cloneData(model(s));doc.sections[0]!.blocks=[para('A'),{id:'break1',type:'page-break'},para('B'),{id:'break2',type:'page-break'},para('C')];
 doc.sections[0]!.page.margins.top=1270000;doc.sections[0]!.page.margins.bottom=1270000;doc.sections[0]!.pageNumberStart=4;
 const n=new DocumentSession(doc);
 n.execute(setWriterPageStory(n.snapshot(),'section-1','header',{text:'Odd title',variant:'default',paragraphs:[text('Odd title'),text('Odd second','end')],distance:127000},'odd'));
 n.execute(setWriterPageStory(n.snapshot(),'section-1','header',{text:'First title',variant:'first',differentFirstPage:true,paragraphs:[text('First title'),text('First subtitle','center')],distance:127000},'first'));
 n.execute(setWriterPageStory(n.snapshot(),'section-1','header',{text:'Even title',variant:'even',differentEvenOddPages:true,distance:127000},'even'));
 for(const variant of ['first','even'] as const)n.execute(setWriterPageStory(n.snapshot(),'section-1','footer',{text:variant+' ',variant,pageNumber:true,distance:127000},variant+'-f'));
 const m=cloneData(model(n));m.sections.push({...cloneData(m.sections[0]!),id:'linked',storyRefs:undefined,differentFirstPage:undefined,pageNumberStart:10,blocks:[para('D'),{id:'break3',type:'page-break'},para('E')]});
 // Optional properties must be absent, not explicit undefined in model data.
 delete m.sections[1]!.storyRefs;delete m.sections[1]!.differentFirstPage;
 return new DocumentSession(m);
}
test('M7.2 first page overrides numbered parity; linked sections inherit each variant independently',()=>{
 const s=variantFixture(),l=layoutWriter(s.snapshot(),font);
 assert.deepEqual(l.pages.map(p=>p.displayPageNumber),[4,5,6,10,11]);
 assert.deepEqual(l.pages.map(p=>p.runningStories!.filter(b=>b.kind==='header').map(b=>b.text)),[['First title','First subtitle'],['Odd title','Odd second'],['Even title'],['Even title'],['Odd title','Odd second']]);
 assert.deepEqual(l.pages.map(p=>p.runningStories!.filter(b=>b.kind==='footer').map(b=>b.text)),[['first 4'],['Page 5 / report'],['even 6'],['even 10'],['Page 11 / report']]);
 for(const p of l.pages)for(const kind of ['header','footer'] as const){
   const boxes=p.runningStories!.filter(b=>b.kind===kind);
   for(const [i,b]of boxes.entries()){assert.ok(b.y>=0&&b.y+b.height<=p.height);if(i)assert.ok(b.y>=boxes[i-1]!.y+boxes[i-1]!.height);assert.ok(kind==='header'?b.y+b.height<=p.margins.top:b.y>=p.height-p.margins.bottom);}
 }
 const before=s.snapshot().document,h=effectivePageStory(model(s),'section-1','header','even');
 s.execute(setWriterPageStory(s.snapshot(),'linked','header',{text:'Detached even',variant:'even'},'detached'));
 assert.notEqual(effectivePageStory(model(s),'linked','header','even'),h);assert.equal(effectivePageStory(model(s),'section-1','header','even'),h);
 assert.equal(layoutWriter(s.snapshot(),font).pages[3]!.runningStories!.find(b=>b.kind==='header')!.text,'Detached even');s.undo();assert.deepEqual(s.snapshot().document,before);
});
test('M7.2 native titlePg and evenAndOddHeaders flags survive new DOCX, source edit and toggle',()=>{
 const s=variantFixture(),bytes=writeWriterDocx(model(s)),src=openWriterDocx(bytes,'story-fixture','Story'),parts=openOpc(bytes);
 assert.equal(src.document.differentEvenOddPages,true);assert.equal(src.document.sections[0]!.differentFirstPage,true);assert.ok(src.document.sections[0]!.storyRefs!.firstHeader);
 const settings=parts.relationships.get(parts.mainPart!)!.find(r=>r.type.endsWith('/settings'))!.targetPart!;
 assert.match(new TextDecoder().decode(parts.parts.get(settings)!),/evenAndOddHeaders/);
 assert.deepEqual(saveWriterDocx(src,src.document),bytes);
 const n=new DocumentSession(src.document);
 n.execute(setWriterPageStory(n.snapshot(),src.document.sections[0]!.id,'header',{text:'First title',variant:'first',differentFirstPage:false,differentEvenOddPages:false},'toggle'));
 const saved=saveWriterDocx(src,n.snapshot().document),back=openWriterDocx(saved,'story-fixture','Story');
 assert.equal(back.document.differentEvenOddPages,undefined);assert.equal(back.document.sections[0]!.differentFirstPage,undefined);
 assert.ok(back.document.sections[0]!.storyRefs!.firstHeader);
 assert.deepEqual(back.document.pageStories![back.document.sections[0]!.storyRefs!.firstHeader!]!.additionalParagraphs?.map(p=>p.text),['First subtitle']);
 assert.equal(layoutWriter(new DocumentSession(back.document).snapshot(),font).pages[0]!.runningStories!.find(b=>b.kind==='header')!.text,'Odd title');
 assert.doesNotMatch(new TextDecoder().decode(openOpc(saved).parts.get(settings)!),/evenAndOddHeaders/);
 assert.throws(()=>writeLegacyDoc(model(s)),fails('unsupported'));
});
test('M7.2 absent variant stays blank; empty interior paragraphs occupy space; oversized stories fail closed',()=>{
 const s=variantFixture(),d=cloneData(model(s));
 delete d.sections[0]!.storyRefs!.firstFooter;
 const id=d.sections[0]!.storyRefs!.firstHeader!;
 d.pageStories![id]!.additionalParagraphs=[{text:''},{text:'After blank'}];
 const l=layoutWriter(new DocumentSession(d).snapshot(),font),boxes=l.pages[0]!.runningStories!;
 assert.equal(boxes.some(b=>b.kind==='footer'),false);
 assert.equal(boxes.length,2);assert.ok(boxes[1]!.y>boxes[0]!.y+boxes[0]!.height);
 d.sections[0]!.page.margins.top=250000;
 assert.throws(()=>layoutWriter(new DocumentSession(d).snapshot(),font),fails('unsupported'));
});
test('M7.2 enabling even/odd stories on a minimal source adds only an understood settings relationship',()=>{
 const s=fixture(),b=writeWriterDocx(model(s)),source=openWriterDocx(b,'story-fixture','Story'),n=new DocumentSession(source.document);
 n.execute(setWriterPageStory(n.snapshot(),source.document.sections[0]!.id,'header',{text:'Even',variant:'even',differentEvenOddPages:true},'add-even'));
 const out=saveWriterDocx(source,n.snapshot().document),back=openWriterDocx(out,'story-fixture','Story');
 assert.equal(back.document.differentEvenOddPages,true);assert.ok(back.document.sections[0]!.storyRefs!.evenHeader);
 assert.equal(openOpc(out).relationships.get(source.mainPart)!.filter(r=>r.type.endsWith('/settings')).length,1);
 const before=openOpc(b),after=openOpc(out);for(const [path,part]of before.parts)if(path.includes('header-oex')||path.includes('footer-oex'))assert.deepEqual(after.parts.get(path),part);
});
