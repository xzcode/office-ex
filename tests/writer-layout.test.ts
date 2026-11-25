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

test('M5 contextual spacing and paragraph bottom rules use adjacent style identity and exclude decoration from caret height',()=>{
 const d=document('');d.sections[0]!.page={width:2000000,height:5000000,margins:{top:100000,bottom:100000,left:100000,right:100000}};
 const rule={color:{kind:'rgba' as const,red:79,green:129,blue:189,alpha:255},width:12700,space:50800};
 d.styles={Title:{id:'Title',name:'Title',kind:'paragraph',properties:{fontSize:152400,spaceBefore:63500,spaceAfter:127000,contextualSpacing:true,paragraphBottomBorder:rule,indent:127000,rightIndent:127000,firstLineIndent:-63500}}};
 d.sections[0]!.blocks=[
 {id:'p1',type:'paragraph',styleId:'Title',runs:[{id:'r1',type:'text',text:'A'}]},
 {id:'p2',type:'paragraph',styleId:'Title',runs:[{id:'r2',type:'text',text:'B'}]},
 {id:'p3',type:'paragraph',runs:[{id:'r3',type:'text',text:'C'}]}
 ];
 const result=layoutWriter(new DocumentSession(d).snapshot(),font),[first,second,third]=result.pages[0]!.lines;
 assert.equal(first!.x,163500);assert.equal(first!.y,163500);assert.equal(first!.bottomBorder,undefined);
 assert.equal(second!.y,first!.y+first!.height);assert.equal(second!.bottomBorder!.x,227000);
 assert.equal(second!.bottomBorder!.width,1546000);assert.equal(second!.height-second!.contentHeight!,63500);
 assert.equal(third!.y,second!.y+second!.height+127000);
 const caret=writerCaretBox(result,{paragraphId:'p2',offset:1});assert.equal(caret?.height,second!.contentHeight);
 assert.equal(writerSelectionBoxes(result,'p2',0,1)[0]?.height,second!.contentHeight);
});

test('M5 right and first-line indent constrain wrapping and preserve hanging-line coordinates',()=>{
 const d=document('abcdef');const p=d.sections[0]!.blocks[0]!;if(p.type!=='paragraph')throw Error();
 p.style={indent:100,rightIndent:100,firstLineIndent:-50};
 const lines=layoutWriter(new DocumentSession(d).snapshot(),font).pages.flatMap(page=>page.lines);
 assert.equal(lines[0]!.x,100);assert.equal(lines[1]!.x,150);assert.equal(lines[0]!.fragments.map(f=>f.text).join(''),'ab');assert.ok(lines.every(line=>line.x+line.width<=350));
 p.style={indent:0,rightIndent:0,firstLineIndent:-100};assert.throws(()=>layoutWriter(new DocumentSession(d).snapshot(),font),fails('unsupported'));
});
