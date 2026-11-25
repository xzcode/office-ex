// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import {createDocument,DocumentSession,buildWriterNumberingMarkers,validateDocument,cloneData,type WriterDocument,type Paragraph,type WriterNumbering} from '@office-ex/core';
import {setWriterList,restartWriterList,splitParagraph,replaceParagraphRange,copyWriterFragment,pasteWriterFragment} from '@office-ex/writer';
import {openWriterDocx,writeWriterDocx,saveWriterDocx,openOpc,writeZip} from '@office-ex/ooxml';
import {writeLegacyDoc,inspectForEditor} from '@office-ex/format-service';
import {WriterPaginationEngine,layoutWriter,type FontMetrics} from '@office-ex/render';
import {graphemeBoundaries,OfficeError} from '@office-ex/contracts';
const fails=(code:string)=>(e:unknown)=>e instanceof OfficeError&&e.code===code;
const doc=(s:DocumentSession)=>s.snapshot().document as WriterDocument;
const markers=(d:WriterDocument)=>[...buildWriterNumberingMarkers(d).values()].map(m=>m.text);
const para=(id:string,text:string):Paragraph=>({id,type:'paragraph',runs:[{id:id+'-run',type:'text',text}]});
const font:FontMetrics={environmentId:'numbering-test',measure(text,request){
 const b=graphemeBoundaries(text),clusters=b.slice(0,-1).map((start,i)=>({start,end:b[i+1]!,left:i*6000,right:(i+1)*6000}));return {advance:clusters.length*6000,ascent:100000,descent:30000,actualFamily:request.family,clusters,clusterPrecision:'exact'};
}};
test('M5 Word list creation, hierarchy, restart, removal and undo preserve logical paragraph text',()=>{
 const d=createDocument('writer','numbering');d.sections[0]!.blocks=[para('a','One'),para('b','Nested'),para('c','Nested again'),para('d','Next'),para('e','Last')];const s=new DocumentSession(d,{saved:true});
 for(const [id,level]of [['a',0],['b',1],['c',1],['d',0],['e',0]] as const)s.execute(setWriterList(s.snapshot(),id,'decimal',level));
 assert.deepEqual(markers(doc(s)),['1.','1.1.','1.2.','2.','3.']);assert.equal(Object.keys(doc(s).numbering!.lists).length,1);
 s.execute(restartWriterList(s.snapshot(),'d',7));assert.deepEqual(markers(doc(s)),['1.','1.1.','1.2.','7.','8.']);
 s.execute(setWriterList(s.snapshot(),'c','none'));assert.deepEqual(markers(doc(s)),['1.','1.1.','7.','8.']);s.undo();s.undo();
 assert.deepEqual(markers(doc(s)),['1.','1.1.','1.2.','2.','3.']);
 s.execute(splitParagraph(s.snapshot(),'a',1,'split-number'));assert.deepEqual(markers(doc(s)),['1.','2.','2.1.','2.2.','3.','4.']);
 assert.equal((doc(s).sections[0]!.blocks[0] as Paragraph).runs.map(r=>r.type==='text'?r.text:'').join(''),'O');
 s.undo();for(let i=0;i<5;i++)s.undo();assert.equal(s.dirty,false);assert.deepEqual(doc(s),d);
});
test('M5 Word numbering persists in new and edited DOCX, exact no-op, clipboard remaps and legacy rejects',async()=>{
 const d=createDocument('writer','numbering-format');d.defaultTextStyle={fontFamily:'Arial',fontSize:152400};d.sections[0]!.blocks=[para('a','First'),para('b','Second')];const s=new DocumentSession(d);
 s.execute(setWriterList(s.snapshot(),'a','decimal'));s.execute(setWriterList(s.snapshot(),'b','decimal',1));
 const bytes=writeWriterDocx(doc(s)),opened=openWriterDocx(bytes,d.documentId,d.title);
 assert.deepEqual(markers(opened.document as WriterDocument),['1.','1.1.']);assert.deepEqual(saveWriterDocx(opened,opened.document),bytes);
 assert.throws(()=>writeLegacyDoc(doc(s)),fails('unsupported'));
 const session=new DocumentSession(opened.document,{saved:true}),[a,b]=doc(session).sections[0]!.blocks as Paragraph[];
 session.execute(restartWriterList(session.snapshot(),b!.id,5));
 const edited=saveWriterDocx(opened,session.snapshot().document),reopened=openWriterDocx(edited,d.documentId,d.title);
 assert.deepEqual(markers(reopened.document as WriterDocument),['1.','1.5.']);
 // Verify source numbering stays immutable and only appended definitions change.
 const original=String(new TextDecoder().decode(openOpc(bytes).parts.get('word/numbering.xml')!));
 assert.match(original,/<w:num w:numId="1"/);
 session.undo();assert.deepEqual(saveWriterDocx(opened,session.snapshot().document),bytes);
 const fragment=copyWriterFragment(session.snapshot(),{anchor:{paragraphId:a!.id,offset:0},focus:{paragraphId:b!.id,offset:6}});
 assert.equal(fragment.plain,'First\nSecond');
 const target=createDocument('writer','target'),t=new DocumentSession(target),p=target.sections[0]!.blocks[0] as Paragraph;
 t.execute(setWriterList(t.snapshot(),p.id,'bullet'));
 const paste=pasteWriterFragment(t.snapshot(),{anchor:{paragraphId:p.id,offset:0},focus:{paragraphId:p.id,offset:0}},fragment.rich,'paste-number');
 t.execute(paste.command);assert.deepEqual(markers(doc(t)),['1.','1.1.']);assert.equal(Object.keys(doc(t).numbering!.lists).length,2);
 const bad=new Map(openOpc(bytes).parts);const xml=new TextDecoder().decode(bad.get('word/numbering.xml')!);bad.set('word/numbering.xml',new TextEncoder().encode(xml.replace('w:val="decimal"','w:val="upperLetter"')));
 assert.equal((await inspectForEditor(writeZip(bad),'unsupported.docx','protected')).report.readOnly,true);
});
test('M5 numbering crosses sections and invalidates cached later-section marker labels',()=>{
 const d=createDocument('writer','numbering-pages');d.sections[0]!.blocks=[para('p1','One'),para('p2','Two')];const s=new DocumentSession(d);
 s.execute(setWriterList(s.snapshot(),'p1','decimal'));s.execute(setWriterList(s.snapshot(),'p2','decimal'));
 const copy=cloneData(doc(s));const section=copy.sections[0]!;const b=section.blocks.pop()!;
 copy.sections.push({...cloneData(section),id:'section-two',blocks:[b]});
 const session=new DocumentSession(copy),engine=new WriterPaginationEngine();
 let layout=engine.layout(session.snapshot(),font);assert.deepEqual(layout.pages.flatMap(p=>p.lines.map(l=>l.listMarker?.text)),['1.','2.']);
 const old=layout.pages[1];session.execute(setWriterList(session.snapshot(),'p1','none'));layout=engine.layout(session.snapshot(),font);
 assert.equal(layout.pages[1]!.lines[0]!.listMarker?.text,'1.');assert.notEqual(layout.pages[1],old);
 assert.equal(layout.pages[1]!.lines[0]!.start,0);
});
