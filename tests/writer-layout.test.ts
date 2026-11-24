// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import { graphemeBoundaries, OfficeError, CancellationToken } from '@office-ex/contracts';
import { createDocument, DocumentSession, type WriterDocument } from '@office-ex/core';
import { layoutWriter, hitTestWriterPage, writerSelectionBoxes, writerCaretBox, type FontMetrics } from '@office-ex/render';
const font:FontMetrics={environmentId:'deterministic-test',measure(text,request){
  const boundaries=graphemeBoundaries(text);const clusters=boundaries.slice(0,-1).map((start,i)=>({start,end:boundaries[i+1]!,left:i*100,right:(i+1)*100}));
  return {advance:clusters.length*100,ascent:80,descent:20,actualFamily:request.family,clusters,clusterPrecision:'exact'};
}};
function document(text:string):WriterDocument{
  const d=createDocument('writer','layout');d.defaultTextStyle={fontSize:100,fontFamily:'test'};
  d.sections[0]!.page={width:500,height:500,margins:{top:50,bottom:50,left:50,right:50}};
  d.sections[0]!.blocks=[{id:'p',type:'paragraph',runs:[{id:'r',type:'text',text}]}];return d;
}
const fails=(code:string)=>(e:unknown)=>e instanceof OfficeError&&e.code===code;
test('WRT-02 measured line boxes wrap at spaces and graphemes, automatically paginate without mutating model',()=>{
  const d=document('ab cd ef gh ij kl mn op qr st');const s=new DocumentSession(d,{saved:true}),snap=s.snapshot(),result=layoutWriter(snap,font);
  assert.ok(result.pages.length>1);assert.equal(result.revision,snap.revision);assert.equal(result.contentId,snap.contentId);
  const lines=result.pages.flatMap(p=>p.lines);
  assert.equal(lines.map(l=>l.fragments.map(f=>f.text).join('')).join(''),'ab cd ef gh ij kl mn op qr st');
  for(const page of result.pages)for(const l of page.lines){assert.ok(l.width<=400);assert.ok(l.y+l.height<=450);assert.ok(l.y>=50);}
  assert.deepEqual(s.snapshot().document,snap.document);assert.equal(s.dirty,false);
});
test('WRT-02 explicit page breaks, empty paragraphs, trailing soft break and section geometry',()=>{
  const d=document('A\n');d.sections[0]!.blocks.push({id:'b',type:'page-break'},{id:'p2',type:'paragraph',runs:[]});
  const result=layoutWriter(new DocumentSession(d).snapshot(),font);
  assert.equal(result.pages.length,2);assert.equal(result.pages[0]!.lines.length,2);assert.equal(result.pages[0]!.lines[1]!.start,2);
  assert.equal(result.pages[1]!.lines[0]!.paragraphId,'p2');assert.equal(result.pages[1]!.lines[0]!.start,0);
});
test('WRT-02 caret and selection boxes use logical UTF16 grapheme positions over wrapped lines',()=>{
  const d=document('A😀B e\u0301C');const result=layoutWriter(new DocumentSession(d).snapshot(),font);
  const caret=writerCaretBox(result,{paragraphId:'p',offset:3});assert.ok(caret);
  assert.deepEqual(hitTestWriterPage(result,caret.pageIndex,{x:caret.x,y:caret.y+10}),{paragraphId:'p',offset:3});
  assert.equal(writerCaretBox(result,{paragraphId:'p',offset:2}),undefined);
  assert.throws(()=>writerSelectionBoxes(result,'p',2,3),fails('invalid'));
  assert.ok(writerSelectionBoxes(result,'p',1,7).length>=1);
});
test('WRT-02 direct formatting, alignment and actual font are represented in line fragments',()=>{
  const d=document('abc');const p=d.sections[0]!.blocks[0]!;if(p.type!=='paragraph')throw Error();
  p.style={alignment:'center',spaceBefore:10};p.runs=[{id:'r1',type:'text',text:'a',style:{bold:true}},{id:'r2',type:'text',text:'bc',style:{italic:true}}];
  const result=layoutWriter(new DocumentSession(d).snapshot(),font),line=result.pages[0]!.lines[0]!;
  assert.equal(line.x,100);assert.equal(line.y,60);
  assert.equal(line.fragments[0]!.request.bold,true);assert.equal(line.fragments[1]!.request.italic,true);
});
test('WRT-02 CJK line-break heuristic avoids starting a line with closing punctuation when a break fits',()=>{
  const d=document('甲乙丙丁，戊己');const result=layoutWriter(new DocumentSession(d).snapshot(),font),lines=result.pages.flatMap(p=>p.lines);
  assert.equal(lines[0]!.fragments.map(f=>f.text).join(''),'甲乙丙');assert.equal(lines[1]!.fragments.map(f=>f.text).join(''),'丁，戊己');
});
test('WRT-02 bounded layout cancels and rejects unsupported bidi, split graphemes and oversized content',()=>{
  const snap=new DocumentSession(document('abcdef')).snapshot();assert.throws(()=>layoutWriter(snap,font,{maxClusters:2}),fails('limit'));
  const c=new CancellationToken();c.cancel();assert.throws(()=>layoutWriter(snap,font,{cancellation:c}),fails('cancelled'));
  assert.throws(()=>layoutWriter(new DocumentSession(document('مرحبا')).snapshot(),font),fails('unsupported'));
  const d=document('');d.sections[0]!.blocks=[{id:'p',type:'paragraph',runs:[{id:'r1',type:'text',text:'e'},{id:'r2',type:'text',text:'\u0301'}]}];
  assert.throws(()=>layoutWriter(new DocumentSession(d).snapshot(),font),fails('unsupported'));
});
