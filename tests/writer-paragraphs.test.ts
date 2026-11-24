// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import { OfficeError } from '@office-ex/contracts';
import { createDocument, DocumentSession, type WriterDocument } from '@office-ex/core';
import { paragraphText, replaceParagraphRange, formatParagraphRange, splitParagraph, joinParagraphs, setParagraphAlignment, insertPageBreakAfter } from '@office-ex/writer';
function session(){
  const model=createDocument('writer','writer-test');
  model.sections[0]!.blocks=[{id:'p',type:'paragraph',runs:[
    {id:'a',type:'text',text:'中文 ',style:{bold:true}},
    {id:'b',type:'text',text:'😀 e',style:{italic:true}},
    {id:'c',type:'text',text:'\u0301 τέλος',style:{underline:true}}
  ]}];
  return new DocumentSession(model,{saved:true});
}
const blocks=(s:DocumentSession)=>(s.snapshot().document as WriterDocument).sections[0]!.blocks;
const fails=(code:string)=>(e:unknown)=>e instanceof OfficeError&&e.code===code;
test('WRT-03 paragraph-wide edit retains unaffected run styles, Unicode and exact undo/redo',()=>{
  const s=session(),original=s.snapshot().document;
  s.execute(replaceParagraphRange(s.snapshot(),'p',3,5,'こんにちは','op'));
  assert.equal(paragraphText(s.snapshot(),'p'),'中文 こんにちは e\u0301 τέλος');
  const p=blocks(s)[0]!;assert.equal(p.type,'paragraph');if(p.type!=='paragraph')throw Error();
  assert.equal(p.runs[0]!.id,'a');assert.equal(p.runs.at(-1)!.id,'c');
  assert.ok(s.dirty);s.undo();assert.deepEqual(s.snapshot().document,original);assert.equal(s.dirty,false);
  s.redo();assert.equal(paragraphText(s.snapshot(),'p'),'中文 こんにちは e\u0301 τέλος');
});
test('WRT-03 rejects grapheme splits spanning adjacent runs and stale commands',()=>{
  const s=session(),original=s.snapshot().document;
  assert.throws(()=>replaceParagraphRange(s.snapshot(),'p',7,8,'x','op'),fails('invalid'));
  assert.throws(()=>splitParagraph(s.snapshot(),'p',4,'op'),fails('invalid'));
  const command=replaceParagraphRange(s.snapshot(),'p',0,2,'abc','op');
  s.execute(setParagraphAlignment(s.snapshot(),'p','center'));
  assert.throws(()=>s.execute(command),fails('conflict'));
  s.undo();assert.deepEqual(s.snapshot().document,original);
});
test('WRT-03 formatting a selected range splits runs without replacing other text',()=>{
  const s=session(),original=s.snapshot().document;
  s.execute(formatParagraphRange(s.snapshot(),'p',1,5,{underline:true},'fmt'));
  const p=blocks(s)[0]!;assert.equal(p.type,'paragraph');if(p.type!=='paragraph')throw Error();
  assert.equal(paragraphText(s.snapshot(),'p'),'中文 😀 e\u0301 τέλος');
  assert.ok(p.runs.some(r=>r.type==='text'&&r.style?.underline&&r.text.includes('😀')));
  assert.equal(new Set(p.runs.map(r=>r.id)).size,p.runs.length);
  s.undo();assert.deepEqual(s.snapshot().document,original);
});
test('WRT-03 paragraph split and join are atomic and preserve text and direct styles',()=>{
  const s=session(),original=s.snapshot().document;
  s.execute(splitParagraph(s.snapshot(),'p',3,'split'));
  assert.equal(blocks(s).length,2);
  const second=blocks(s)[1]!.id;
  assert.equal(paragraphText(s.snapshot(),'p'),'中文 ');
  assert.equal(paragraphText(s.snapshot(),second),'😀 e\u0301 τέλος');
  s.execute(joinParagraphs(s.snapshot(),'p',second));
  assert.equal(paragraphText(s.snapshot(),'p'),'中文 😀 e\u0301 τέλος');
  s.undo();s.undo();assert.deepEqual(s.snapshot().document,original);assert.equal(s.dirty,false);
});
test('WRT-03 empty paragraph remains editable; explicit break does not merge across blocks',()=>{
  const s=new DocumentSession(createDocument('writer','blank'));
  s.execute(splitParagraph(s.snapshot(),'paragraph-1',0,'split'));
  const second=blocks(s)[1]!.id;
  assert.equal(paragraphText(s.snapshot(),second),'');
  s.execute(replaceParagraphRange(s.snapshot(),second,0,0,'مرحبا','type'));
  s.execute(insertPageBreakAfter(s.snapshot(),'paragraph-1','break'));
  assert.equal(blocks(s)[1]!.type,'page-break');
  assert.throws(()=>joinParagraphs(s.snapshot(),'paragraph-1',second),fails('invalid'));
});
test('WRT-03 unsupported inline objects are blocked rather than flattened',()=>{
  const d=createDocument('writer','inline');
  d.sections[0]!.blocks=[{id:'p',type:'paragraph',runs:[{id:'t',type:'tab'},{id:'r',type:'text',text:'hello'}]}];
  const s=new DocumentSession(d,{saved:true});
  assert.throws(()=>replaceParagraphRange(s.snapshot(),'p',0,1,'x','op'),fails('unsupported'));
  assert.equal(s.dirty,false);assert.equal(s.revision,0);
});

test('WRT-03 direct formatting rejects accessor-backed data without invoking it',()=>{
  const s=session();
  let invoked=false;
  assert.throws(()=>formatParagraphRange(s.snapshot(),'p',0,1,{get bold():boolean{invoked=true;return true;}},'fmt'),fails('invalid'));
  assert.equal(invoked,false);assert.equal(s.revision,0);
});
