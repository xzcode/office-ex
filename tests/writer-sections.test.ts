// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import {OfficeError,EMU_PER_TWIP,graphemeBoundaries} from '@office-ex/contracts';
import {createDocument,DocumentSession,makeImageResource,projectParagraph,type WriterDocument} from '@office-ex/core';
import {replaceParagraphRange,splitParagraph,splitWriterSectionAfter,setWriterPage,insertImageAt,insertTableAfter,locateParagraph,formatParagraphRange,setParagraphDirection,resolveWriterRange} from '@office-ex/writer';
import {writeWriterDocx,openWriterDocx,saveWriterDocx,openOpc,writeZip} from '@office-ex/ooxml';
import {layoutWriter,type FontMetrics} from '@office-ex/render';
import {inspectForEditor,serializeDraft,parseDraft} from '@office-ex/format-service';
import {PNG,JPEG} from './writer-image-fixture.js';
const fails=(code:string)=>(e:unknown)=>e instanceof OfficeError&&e.code===code;
const page=(w:number,h:number)=>({width:w*EMU_PER_TWIP,height:h*EMU_PER_TWIP,margins:{top:720*EMU_PER_TWIP,right:720*EMU_PER_TWIP,bottom:720*EMU_PER_TWIP,left:720*EMU_PER_TWIP}});
const model=(s:DocumentSession)=>s.snapshot().document as WriterDocument;
function example(){
  const s=new DocumentSession(createDocument('writer','sections','Sections'),{saved:true});
  s.execute(setWriterPage(s.snapshot(),'section-1',page(11906,16838)));
  s.execute(replaceParagraphRange(s.snapshot(),'paragraph-1',0,0,'First 中文','first'));
  s.execute(splitParagraph(s.snapshot(),'paragraph-1',8,'split'));
  const id=model(s).sections[0]!.blocks[1]!.id;
  s.execute(replaceParagraphRange(s.snapshot(),id,0,0,'Second 日本語','second'));
  return {s,id};
}
const font:FontMetrics={environmentId:'sections-test',measure(text,request){
  const b=graphemeBoundaries(text);return {advance:(b.length-1)*50000,ascent:80000,descent:20000,actualFamily:request.family,clusterPrecision:'exact',clusters:b.slice(0,-1).map((start,i)=>({start,end:b[i+1]!,left:i*50000,right:(i+1)*50000}))};
}};
test('WRT-01 next-page sections retain IDs, independent page geometry and layout; cross-section selection is rejected',()=>{
  const {s,id}=example(),before=s.snapshot().document;
  const split=splitWriterSectionAfter(s.snapshot(),'paragraph-1','section');s.execute(split.command);
  assert.equal(split.paragraphId,id);assert.equal(model(s).sections[1]!.blocks[0]!.id,id);
  s.execute(setWriterPage(s.snapshot(),split.sectionId,page(15840,12240)));
  const l=layoutWriter(s.snapshot(),font);
  assert.equal(l.pages.length,2);assert.equal(l.pages[0]!.width,11906*EMU_PER_TWIP);assert.equal(l.pages[1]!.width,15840*EMU_PER_TWIP);
  assert.throws(()=>resolveWriterRange(s.snapshot(),{anchor:{paragraphId:'paragraph-1',offset:0},focus:{paragraphId:id,offset:1}}),fails('unsupported'));
  const original=writeWriterDocx(s.snapshot().document),read=openWriterDocx(original,'read','Read');
  assert.deepEqual(read.document.sections.map(s=>s.page),model(s).sections.map(s=>s.page));
  assert.equal(read.document.sections.length,2);assert.deepEqual(saveWriterDocx(read,read.document),original);
  s.undo();s.undo();assert.deepEqual(s.snapshot().document,before);
});
test('WRT-04 section offset and explicit one-column metadata preserved; continuous sections fail closed',async()=>{
  const {s}=example(),d=structuredClone(model(s));d.sections[0]!.pageExtras={headerDistance:720*EMU_PER_TWIP,footerDistance:360*EMU_PER_TWIP,columnSpace:720*EMU_PER_TWIP};
  const next=new DocumentSession(d);next.execute(splitWriterSectionAfter(next.snapshot(),'paragraph-1','split-section').command);
  const bytes=writeWriterDocx(next.snapshot().document),source=openWriterDocx(bytes,'read','Read');
  assert.deepEqual(source.document.sections.map(s=>s.pageExtras),model(next).sections.map(s=>s.pageExtras));
  const pkg=openOpc(bytes),parts=new Map(pkg.parts),main=pkg.mainPart!,text=new TextDecoder().decode(parts.get(main)!);
  parts.set(main,new TextEncoder().encode(text.replace('<w:sectPr>','<w:sectPr><w:type w:val="continuous"/>')));
  const invalid=writeZip(parts);assert.throws(()=>openWriterDocx(invalid,'bad','Read'),fails('unsupported'));
  assert.equal((await inspectForEditor(invalid,'bad.docx','bad')).report.readOnly,true);
});
test('I18N-04 document language and direct run/paragraph direction survive DOCX and OEX without locale normalization',()=>{
  const {s}=example();
  s.execute(formatParagraphRange(s.snapshot(),'paragraph-1',6,8,{language:'zh-TW',eastAsiaLanguage:'ja-JP',bidiLanguage:'ar',runDirection:'rtl'},'language'));
  s.execute(setParagraphDirection(s.snapshot(),'paragraph-1','rtl'));
  const bytes=writeWriterDocx(s.snapshot().document),source=openWriterDocx(bytes,'language','Read');
  const p=source.document.sections[0]!.blocks[0]!;assert.equal(p.type,'paragraph');if(p.type!=='paragraph')throw Error();
  assert.equal(projectParagraph(p),'First 中文');assert.equal(p.style?.paragraphDirection,'rtl');
  const run=p.runs.find(r=>r.type==='text'&&r.text==='中文')!;if(run.type!=='text')throw Error();
  assert.equal(run.style?.language,'zh-TW');assert.equal(run.style?.eastAsiaLanguage,'ja-JP');assert.equal(run.style?.bidiLanguage,'ar');assert.equal(run.style?.runDirection,'rtl');
  assert.deepEqual(parseDraft(serializeDraft(s.snapshot().document)),s.snapshot().document);
  assert.throws(()=>layoutWriter(s.snapshot(),font),fails('unsupported')); // No false bidi fidelity claims.
  assert.throws(()=>formatParagraphRange(s.snapshot(),'paragraph-1',0,1,{language:'bad locale'},'bad'),fails('invalid'));
  assert.deepEqual(saveWriterDocx(source,source.document),bytes);
});
