// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import {OfficeError,EMU_PER_TWIP,EMU_PER_POINT,graphemeBoundaries} from '@office-ex/contracts';
import {createDocument,DocumentSession,makeImageResource,validateImageResource,decodeImageBytes,inspectImage,cloneData,type WriterDocument} from '@office-ex/core';
import {insertImageAfter,resizeInlineImage,removeInlineImage,setWriterPage,setParagraphMetrics,insertPageBreakAfter} from '@office-ex/writer';
import {layoutWriter,type FontMetrics} from '@office-ex/render';
import {writeWriterDocx,openWriterDocx,saveWriterDocx,openOpc,writeZip} from '@office-ex/ooxml';
import {serializeDraft,parseDraft,inspectForEditor,prepareWriterDocxSave} from '@office-ex/format-service';
import {PNG,JPEG,imageFixture} from './writer-image-fixture.js';
import {styledFixture} from './writer-styled-fixture.js';
const fails=(code:string)=>(e:unknown)=>e instanceof OfficeError&&e.code===code;
function model(s:DocumentSession):WriterDocument{return s.snapshot().document as WriterDocument;}
test('WRT-03 packaged PNG/JPEG are content addressed with exact bytes, dimensions and bounded header checks',()=>{
  for(const [bytes,type] of [[PNG,'image/png'],[JPEG,'image/jpeg']] as const){
    const r=makeImageResource(bytes);assert.equal(r.mimeType,type);assert.equal(r.pixelWidth,8);assert.equal(r.pixelHeight,6);
    assert.deepEqual(decodeImageBytes(r.dataBase64),Uint8Array.from(bytes));validateImageResource(r);
    assert.throws(()=>validateImageResource({...r,pixelWidth:9}),fails('invalid'));
    assert.throws(()=>validateImageResource({...r,sha256:'0'.repeat(64)}),fails('invalid'));
    assert.throws(()=>inspectImage(bytes.subarray(0,bytes.length-1)),fails('invalid'));
  }
  const changed=Uint8Array.from(PNG);changed[24]^=1;assert.throws(()=>inspectImage(changed),fails('invalid'));
  assert.throws(()=>makeImageResource(new TextEncoder().encode('<svg onload="x"/>')),fails('unsupported'));
});
test('WRT-03 image insert, dedup, proportional resize, delete and undo retain payloads; OEX round-trip exact',()=>{
  const s=imageFixture(),original=s.snapshot().document;
  s.execute(insertImageAfter(s.snapshot(),'paragraph-1',makeImageResource(PNG),'duplicate'));
  assert.equal(Object.keys(model(s).resources).length,2);
  s.execute(resizeInlineImage(s.snapshot(),'duplicate-image',200*EMU_PER_POINT));
  const p=model(s).sections[0]!.blocks[1]!;
  if(p.type!=='paragraph'||p.runs[0]?.type!=='image')throw Error();
  assert.equal(p.runs[0].height,150*EMU_PER_POINT);
  s.execute(removeInlineImage(s.snapshot(),'duplicate-image'));assert.equal(Object.keys(model(s).resources).length,2);
  for(let i=0;i<3;i++)s.undo();assert.deepEqual(s.snapshot().document,original);
  assert.deepEqual(parseDraft(serializeDraft(s.snapshot().document)),s.snapshot().document);
});
test('WRT-04 new and edited DOCX embeds exact original PNG/JPEG, validates graph and preserves source media',async()=>{
  const s=imageFixture(),doc=model(s),original=writeWriterDocx(doc),opened=openWriterDocx(original,'opened','Pictures');
  assert.equal(Object.keys(opened.document.resources).length,2);
  assert.deepEqual(saveWriterDocx(opened,opened.document),original);
  const inspected=await inspectForEditor(original,'pictures.docx','pictures');
  assert.ok(inspected.editableWriter);
  const next=new DocumentSession(inspected.editableWriter,{saved:true});
  const p=next.snapshot().document as WriterDocument;
  const image=p.sections[0]!.blocks[1]!;
  if(image.type!=='paragraph'||image.runs[0]?.type!=='image')throw Error();
  next.execute(resizeInlineImage(next.snapshot(),image.runs[0].id,100*EMU_PER_POINT));
  const updated=prepareWriterDocxSave(serializeDraft(next.snapshot().document),'pictures',original);
  const before=openOpc(original),after=openOpc(updated);
  for(const [name,data] of before.parts)if(name!==before.mainPart)assert.deepEqual(after.parts.get(name),data,name);
  assert.deepEqual(openWriterDocx(updated,'pictures','Pictures').document.resources,opened.document.resources);
});
test('WRT-04 adding media to a source-owned styles profile changes only explicit graph parts and preserves source styles',()=>{
  const original=styledFixture(),source=openWriterDocx(original,'styled-image','Styled'),s=new DocumentSession(source.document);
  s.execute(insertImageAfter(s.snapshot(),source.document.sections[0]!.blocks[0]!.id,makeImageResource(PNG),'new-pic'));
  const updated=saveWriterDocx(source,s.snapshot().document),a=openOpc(original),b=openOpc(updated);
  for(const [name,data] of a.parts)if(![a.mainPart,'[Content_Types].xml','word/_rels/document.xml.rels'].includes(name))assert.deepEqual(b.parts.get(name),data,name);
  const reopened=openWriterDocx(updated,'styled-image','Styled');
  assert.equal(Object.keys(reopened.document.resources).length,1);assert.deepEqual(reopened.document.styles,source.document.styles);
  assert.deepEqual(saveWriterDocx(reopened,reopened.document),updated);
});
test('WRT-04 unsupported floating, cropped or mixed drawings stay read-only and preserve original',async()=>{
  const original=writeWriterDocx(model(imageFixture())),pkg=openOpc(original),dec=new TextDecoder(),enc=new TextEncoder();
  for(const mutate of [
    (s:string)=>s.replace('<wp:inline ','<wp:anchor ').replace('</wp:inline>','</wp:anchor>'),
    (s:string)=>s.replace('<a:stretch>','<a:srcRect l="1000"/><a:stretch>'),
    (s:string)=>s.replace('<w:drawing>','<w:t>Mixed text</w:t><w:drawing>')
  ]){
    const parts=new Map(pkg.parts);parts.set(pkg.mainPart!,enc.encode(mutate(dec.decode(parts.get(pkg.mainPart!)!))));
    const bytes=writeZip(parts);assert.throws(()=>openWriterDocx(bytes,'blocked','Blocked'),fails('unsupported'));
    assert.equal((await inspectForEditor(bytes,'blocked.docx','blocked')).report.readOnly,true);
  }
});
const font:FontMetrics={environmentId:'image-test',measure(text,request){
  const b=graphemeBoundaries(text);return {advance:(b.length-1)*50000,ascent:80000,descent:20000,actualFamily:request.family,clusterPrecision:'exact',clusters:b.slice(0,-1).map((start,i)=>({start,end:b[i+1]!,left:i*50000,right:(i+1)*50000}))};
}};
test('WRT-02 page settings, margins, paragraph metrics and image layout use one immutable revision',()=>{
  const s=imageFixture(),before=s.snapshot().document;
  s.execute(setWriterPage(s.snapshot(),'section-1',{width:15840*EMU_PER_TWIP,height:12240*EMU_PER_TWIP,margins:{top:720*EMU_PER_TWIP,right:720*EMU_PER_TWIP,bottom:720*EMU_PER_TWIP,left:720*EMU_PER_TWIP}}));
  s.execute(setParagraphMetrics(s.snapshot(),'paragraph-1',{indent:120*EMU_PER_TWIP,spaceAfter:120*EMU_PER_TWIP,lineHeight:360*EMU_PER_TWIP}));
  s.execute(insertPageBreakAfter(s.snapshot(),'paragraph-1','explicit'));
  const current=s.snapshot(),layout=layoutWriter(current,font);assert.ok(layout.pages.length>=2);
  const images=layout.pages.flatMap(p=>p.images);assert.equal(images.length,2);
  assert.ok(layout.pages.every(p=>p.images.every(i=>i.x>=p.margins.left&&i.y+i.height<=p.height-p.margins.bottom)));
  assert.deepEqual(s.snapshot(),current);
  const reopened=openWriterDocx(writeWriterDocx(current.document),'page','Page').document;
  assert.deepEqual(reopened.sections[0]!.page,model(s).sections[0]!.page);
  for(let i=0;i<3;i++)s.undo();assert.deepEqual(s.snapshot().document,before);
});
