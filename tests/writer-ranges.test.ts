// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import {OfficeError,EMU_PER_POINT,graphemeBoundaries} from '@office-ex/contracts';
import {DocumentSession,createDocument,makeImageResource,projectParagraph,paragraphBoundaries,OBJECT_REPLACEMENT,type WriterDocument,type Paragraph} from '@office-ex/core';
import {replaceWriterRange,resolveWriterRange,writerRangeText,formatWriterRange,writerTextGroup,textGroupPoint,textGroupOffset,replaceParagraphRange,splitParagraph,insertPageBreakAfter,insertImageAt,resizeInlineImage,removeInlineImage,insertTableAfter,insertTableRow,resizeTableColumn,paragraphText,insertImageAfter,locateParagraph} from '@office-ex/writer';
import {writeWriterDocx,openWriterDocx,saveWriterDocx,openOpc} from '@office-ex/ooxml';
import {layoutWriter,hitTestWriterPage,type FontMetrics} from '@office-ex/render';
import {PNG,JPEG} from './writer-image-fixture.js';
const fails=(code:string)=>(e:unknown)=>e instanceof OfficeError&&e.code===code;
function doc(s:DocumentSession):WriterDocument{return s.snapshot().document as WriterDocument;}
function p(s:DocumentSession,id:string){return locateParagraph(s.snapshot(),id).paragraph;}
function range(a:string,from:number,b:string,to:number){return {anchor:{paragraphId:a,offset:from},focus:{paragraphId:b,offset:to}};}
function sample(){
  const d=createDocument('writer','cross-range');d.sections[0]!.blocks=[
    {type:'paragraph',id:'a',runs:[{type:'text',id:'a1',text:'甲😀 alpha',style:{italic:true}}]},
    {type:'paragraph',id:'b',style:{alignment:'center'},runs:[{type:'text',id:'b1',text:'middle',style:{bold:true}}]},
    {type:'paragraph',id:'c',style:{alignment:'end'},runs:[{type:'text',id:'c1',text:'e\u0301 omega',style:{underline:true}}]},
  ];return new DocumentSession(d,{saved:true});
}
test('WRT-03 reverse cross-paragraph selection copies logical text, replaces in one transaction and preserves unaffected formatting',()=>{
  const s=sample(),before=s.snapshot().document,r=range('c',2,'a',3);
  assert.equal(writerRangeText(s.snapshot(),r),' alpha\nmiddle\ne\u0301');
  assert.throws(()=>resolveWriterRange(s.snapshot(),range('a',2,'c',2)),fails('invalid'));
  const edit=replaceWriterRange(s.snapshot(),r,'新\n段','paste');
  s.execute(edit.command);
  assert.deepEqual(edit.caret,{paragraphId:'c',offset:1});
  assert.deepEqual(doc(s).sections[0]!.blocks.map(b=>b.id),['a','c']);
  assert.equal(projectParagraph(p(s,'a')),'甲😀新');assert.equal(projectParagraph(p(s,'c')),'段 omega');
  assert.equal(p(s,'c').style?.alignment,'end');const tail=p(s,'c').runs.at(-1)!;assert.equal(tail.type,'text');if(tail.type!=='text')throw Error();assert.equal(tail.style?.underline,true);assert.equal(tail.id,'c1');
  s.undo();assert.deepEqual(s.snapshot().document,before);assert.equal(s.dirty,false);s.redo();assert.equal(projectParagraph(p(s,'c')),'段 omega');
});
test('WRT-03 cross-paragraph direct formatting is one undo unit and structure boundaries cannot be crossed',()=>{
  const s=sample(),before=s.snapshot().document;
  s.execute(formatWriterRange(s.snapshot(),range('a',3,'c',2),{fontSize:177800},'fmt'));
  for(const id of ['a','b','c'])assert.ok(p(s,id).runs.some(r=>r.type==='text'&&r.style?.fontSize===177800));
  s.undo();assert.deepEqual(s.snapshot().document,before);
  s.execute(insertPageBreakAfter(s.snapshot(),'b','page'));
  assert.throws(()=>replaceWriterRange(s.snapshot(),range('a',0,'c',2),'x','blocked'),fails('unsupported'));
  const withBreak=s.snapshot();assert.throws(()=>writerTextGroup(s.snapshot(),'page'),fails('unsupported'));assert.deepEqual(s.snapshot(),withBreak);
});
test('WRT-03 continuous native group offsets round-trip at paragraph boundaries and multi-line paste retains soft breaks',()=>{
  const s=sample();s.execute(replaceParagraphRange(s.snapshot(),'b',6,6,'\nsoft','soft'));
  const g=writerTextGroup(s.snapshot(),'b');assert.equal(g.paragraphs.length,3);
  for(const id of ['a','b','c']){const loc=g.paragraphs.find(x=>x.id===id)!;for(const n of [loc.start,loc.end])assert.equal(textGroupOffset(g,textGroupPoint(g,n)),n);}
  const old=s.snapshot().document;const change=replaceWriterRange(s.snapshot(),range('b',0,'b',0),'line1\r\nline2\n','paste');s.execute(change.command);
  assert.equal(doc(s).sections[0]!.blocks.length,5);assert.equal(projectParagraph(p(s,change.caret.paragraphId)),'middle\nsoft');
  s.undo();assert.deepEqual(s.snapshot().document,old);
});
test('WRT-03 atomic inline objects preserve adjacent text, stable identity and image payload through edits and explicit removal',()=>{
  const s=sample(),r=makeImageResource(PNG),before=s.snapshot().document;
  s.execute(insertImageAt(s.snapshot(),'a',3,3,r,'pic'));
  assert.equal(projectParagraph(p(s,'a')),'甲😀'+OBJECT_REPLACEMENT+' alpha');
  assert.deepEqual(paragraphBoundaries(p(s,'a')).slice(0,4),[0,1,3,4]);
  assert.throws(()=>replaceParagraphRange(s.snapshot(),'a',0,0,OBJECT_REPLACEMENT,'fake'),fails('invalid'));
  assert.throws(()=>writerRangeText(s.snapshot(),range('a',3,'a',4)),fails('unsupported'));
  assert.throws(()=>writerTextGroup(s.snapshot(),'a'),fails('unsupported'));
  s.execute(replaceParagraphRange(s.snapshot(),'a',4,4,'New','text'));
  s.execute(resizeInlineImage(s.snapshot(),'pic-image',30*EMU_PER_POINT));
  const image=p(s,'a').runs.find(r=>r.type==='image')!;assert.equal(image.height,Math.round(30*EMU_PER_POINT*6/8));
  assert.ok(projectParagraph(p(s,'a')).includes(OBJECT_REPLACEMENT+'New'));
  s.execute(removeInlineImage(s.snapshot(),'pic-image'));
  assert.equal(projectParagraph(p(s,'a')),'甲😀New alpha');assert.ok(doc(s).resources[r.id]);
  for(let i=0;i<4;i++)s.undo();assert.deepEqual(s.snapshot().document,before);
});
const font:FontMetrics={environmentId:'mixed-inline-test',measure(text,request){
  const b=graphemeBoundaries(text);return {advance:(b.length-1)*50000,ascent:80000,descent:20000,actualFamily:request.family,clusterPrecision:'exact',clusters:b.slice(0,-1).map((start,i)=>({start,end:b[i+1]!,left:i*50000,right:(i+1)*50000}))};
}};
export function mixedFixture(){
  const s=new DocumentSession(createDocument('writer','mixed','Mixed images'),{saved:true});
  s.execute(replaceParagraphRange(s.snapshot(),'paragraph-1',0,0,'before after','text'));
  s.execute(insertImageAt(s.snapshot(),'paragraph-1',7,7,makeImageResource(PNG),'inline'));
  s.execute(insertTableAfter(s.snapshot(),'paragraph-1',2,2,'table'));
  const table=doc(s).sections[0]!.blocks[1]!;if(table.type!=='table')throw Error();
  const cell=table.rows[0]!.cells[0]!.blocks[0]!.id;
  s.execute(replaceParagraphRange(s.snapshot(),cell,0,0,'cell text','cell'));
  s.execute(insertImageAt(s.snapshot(),cell,5,5,makeImageResource(JPEG),'cell-pic'));
  return s;
}
test('WRT-02/WRT-04 mixed text and cell pictures share line boxes, caret hit test and DOCX resource round-trip',()=>{
  const s=mixedFixture(),original=s.snapshot(),l=layoutWriter(original,font);
  assert.equal(l.pages.flatMap(p=>p.images).length,2);
  for(let pageIndex=0;pageIndex<l.pages.length;pageIndex++)for(const image of l.pages[pageIndex]!.images){
    const page=l.pages[pageIndex]!,line=page.lines.find(line=>line.images.some(i=>i.imageId===image.imageId))!;
    assert.ok(image.y>=line.y&&image.y+image.height<=line.y+line.height);
    const hit=hitTestWriterPage(l,pageIndex,{x:image.x+1,y:image.y+1});assert.equal(hit?.paragraphId,image.paragraphId);
    if(image.imageId==='cell-pic-image'){
      const cell=page.tableCells.find(c=>c.paragraphIds.includes(image.paragraphId))!;
      assert.ok(image.x>=cell.x&&image.x+image.width<=cell.x+cell.width&&image.y+image.height<=cell.y+cell.height);
    }
  }
  assert.deepEqual(s.snapshot(),original);
  const bytes=writeWriterDocx(original.document),opened=openWriterDocx(bytes,'reopen','Mixed');
  const blocks=opened.document.sections[0]!.blocks;
  assert.equal(blocks[0]!.type,'paragraph');if(blocks[0]!.type!=='paragraph'||blocks[1]!.type!=='table')throw Error();
  assert.equal(projectParagraph(blocks[0]!),'before '+OBJECT_REPLACEMENT+'after');
  const cell=blocks[1]!.rows[0]!.cells[0]!.blocks[0]!;if(cell.type!=='paragraph')throw Error();
  assert.equal(projectParagraph(cell),'cell '+OBJECT_REPLACEMENT+'text');
  assert.deepEqual(opened.document.resources,doc(s).resources);assert.deepEqual(saveWriterDocx(opened,opened.document),bytes);
});
test('WRT-04 undoing an insertion after save retains original package media without blocking a later save',()=>{
  const s=new DocumentSession(createDocument('writer','undo-media'),{saved:true});
  s.execute(insertImageAfter(s.snapshot(),'paragraph-1',makeImageResource(PNG),'picture'));
  const saved=writeWriterDocx(s.snapshot().document),source=openWriterDocx(saved,'undo-media','Untitled');
  s.undo();assert.equal(Object.keys(doc(s).resources).length,0);
  const updated=saveWriterDocx(source,s.snapshot().document),a=openOpc(saved),b=openOpc(updated);
  for(const [name,data] of a.parts)if(name!==a.mainPart)assert.deepEqual(b.parts.get(name),data);
  assert.equal(openWriterDocx(updated,'undo-media','Untitled').document.sections[0]!.blocks.length,1);
});
test('WRT-03 table structural editing with embedded images preserves remaining cells and refuses narrowing below picture width',()=>{
  const s=mixedFixture(),table=doc(s).sections[0]!.blocks[1]!;if(table.type!=='table')throw Error();
  const original=s.snapshot().document;
  s.execute(insertTableRow(s.snapshot(),table.id,1,'row'));
  assert.equal((doc(s).sections[0]!.blocks[1] as typeof table).rows.length,3);
  s.undo();assert.deepEqual(s.snapshot().document,original);
  s.execute(resizeInlineImage(s.snapshot(),'cell-pic-image',120*EMU_PER_POINT));
  assert.throws(()=>resizeTableColumn(s.snapshot(),table.id,0,40*EMU_PER_POINT),fails('unsupported'));
});
