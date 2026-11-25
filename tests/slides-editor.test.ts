// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import { OfficeError, EMU_PER_POINT, graphemeBoundaries } from '@office-ex/contracts';
import { createDocument, DocumentSession, makeImageResource, patchCommand, validateDocument, connectionPoint, resolveSlideConnector, type SlidesDocument, type Shape } from '@office-ex/core';
import { addSlide, duplicateSlide, deleteSlide, moveSlide, insertShapes, removeShapes, duplicateShapes, setShapeGeometries, setShapePaint, reorderShapes, alignShapes, insertSlideImage, aspectFitImage, replaceShapeText, formatShapeText, formatSlideParagraphs, shapeText, slideTextEditDiff, setShapeTextFrame, setSlideBackground, setSlideSize, setSlideTheme, DEFAULT_SLIDE_THEME, pasteShapes, groupShapes, ungroupShape, canUngroupShape, setSlideImageCrop, flipSlideShapes, locateShape, connectSlideShapes, setSlideConnections } from '@office-ex/slides';
import { buildSlideScene, resolveSlideColor, hitTest, shapeTransform, transform, inverse, groupChildTransform, slideShapeParentTransform } from '@office-ex/render';
import { PNG, JPEG } from './writer-image-fixture.js';
const fails=(code:string)=>(error:unknown)=>error instanceof OfficeError&&error.code===code;
const geometry={x:100000,y:200000,width:3000000,height:1000000,rotation:0};
const next=()=>{let n=0;return()=>`fresh-${++n}`;};
const model=(s:DocumentSession)=>s.snapshot().document as SlidesDocument;
function sample():DocumentSession {const doc=createDocument('slides','slides-edit');doc.slides[0]!.shapes=[{id:'rect',type:'rect',geometry,fill:{kind:'theme',name:'accent1'}},{id:'ellipse',type:'ellipse',geometry:{...geometry,x:600000,y:700000}},{id:'text',type:'text',geometry:{...geometry,y:1800000},runs:[{id:'run-a',type:'text',text:'A👨‍👩‍👧‍👦',style:{bold:true}},{id:'run-b',type:'text',text:'é中文\nالعربية עברית',style:{language:'ar',runDirection:'rtl',color:{kind:'rgba',red:1,green:2,blue:3,alpha:255}}}],textStyle:{fontSize:24*EMU_PER_POINT},textFrame:{direction:'auto',wrap:true}}];return new DocumentSession(doc,{saved:true});}

test('SLD-01/02 page insertion, duplication, reorder and deletion preserve stable object IDs and undo',()=>{
  const s=sample(),initial=s.snapshot();s.execute(addSlide(s.snapshot(),1,'second'));s.execute(duplicateSlide(s.snapshot(),'slide-1',next()));
  assert.deepEqual(model(s).slides.map(x=>x.id),['slide-1','fresh-1','second']);assert.notEqual(model(s).slides[1]!.shapes[2]!.id,'text');
  const copy=model(s).slides[1]!;assert.equal(new Set([copy.id,...copy.shapes.flatMap(x=>[x.id,...(x.type==='text'?x.runs.map(r=>r.id):[])])]).size,6);
  s.execute(moveSlide(s.snapshot(),'second',0));assert.deepEqual(model(s).slides.map(x=>x.id),['second','slide-1','fresh-1']);s.execute(deleteSlide(s.snapshot(),'slide-1'));assert.equal(model(s).slides.length,2);
  for(let i=0;i<4;i++)s.undo();assert.deepEqual(s.snapshot().document,initial.document);assert.equal(s.dirty,false);for(let i=0;i<4;i++)s.redo();assert.deepEqual(model(s).slides.map(x=>x.id),['second','fresh-1']);
});
test('SLD-02 duplicate IDs, last-slide deletion, invalid page dimensions and partial invalid gesture are atomic',()=>{
  const s=sample(),before=s.snapshot(),history=s.historyUsage;
  assert.throws(()=>s.execute(addSlide(s.snapshot(),1,'slide-1')),fails('invalid'));
  assert.throws(()=>duplicateSlide(s.snapshot(),'slide-1',()=> 'text'),fails('invalid'));
  assert.throws(()=>deleteSlide(s.snapshot(),'slide-1'),fails('invalid'));
  assert.throws(()=>setSlideSize(s.snapshot(),0,100),fails('invalid'));
  assert.throws(()=>s.execute(setShapeGeometries(s.snapshot(),'slide-1',{rect:{...geometry,x:500},ellipse:{...geometry,width:-1}})),fails('invalid'));
  assert.deepEqual(s.snapshot(),before);assert.deepEqual(s.historyUsage,history);assert.equal(s.dirty,false);
});
test('SLD-02 one multi-object geometry command is one transaction, preserves snapshot and rejects stale gesture',()=>{
  const s=sample(),before=s.snapshot(),command=setShapeGeometries(before,'slide-1',{rect:{...geometry,x:-400000,rotation:5400000},ellipse:{...geometry,x:7000000}});
  const result=s.execute(command);assert.equal(result.revision,1);assert.equal(s.historyUsage.items,1);assert.equal((before.document as SlidesDocument).slides[0]!.shapes[0]!.geometry.x,100000);
  assert.throws(()=>s.execute(command),fails('conflict'));s.undo();assert.deepEqual(s.snapshot().document,before.document);assert.equal(s.dirty,false);
});
test('SLD-02 layers and alignments operate on selected objects in paint order',()=>{
  const s=sample();s.execute(reorderShapes(s.snapshot(),'slide-1',['rect'],'front'));assert.deepEqual(model(s).slides[0]!.shapes.map(x=>x.id),['ellipse','text','rect']);
  s.execute(reorderShapes(s.snapshot(),'slide-1',['rect'],'backward'));assert.deepEqual(model(s).slides[0]!.shapes.map(x=>x.id),['ellipse','rect','text']);
  s.execute(alignShapes(s.snapshot(),'slide-1',['rect','ellipse'],'left'));const shapes=model(s).slides[0]!.shapes;assert.equal(shapes[0]!.geometry.x,shapes[1]!.geometry.x);
  s.execute(alignShapes(s.snapshot(),'slide-1',['text'],'bottom'));assert.equal(model(s).slides[0]!.shapes.find(x=>x.id==='text')!.geometry.y,model(s).size.height-geometry.height);
  s.execute(removeShapes(s.snapshot(),'slide-1',['rect','ellipse']));assert.deepEqual(model(s).slides[0]!.shapes.map(x=>x.id),['text']);s.undo();assert.equal(model(s).slides[0]!.shapes.length,3);
});
test('SLD-02 duplicate selection uses fresh node and run IDs with independent geometry',()=>{
  const s=sample();s.execute(duplicateShapes(s.snapshot(),'slide-1',['rect','text'],next()));assert.equal(model(s).slides[0]!.shapes.length,5);
  const copy=model(s).slides[0]!.shapes[4]!;assert.equal(copy.type,'text');assert.equal(copy.geometry.x,geometry.x+190500);assert.notEqual(copy.id,'text');validateDocument(model(s));
});
test('SLD-02/I18N-04 logical replacement preserves rich style outside complete grapheme boundaries',()=>{
  const s=sample(),original=model(s).slides[0]!.shapes[2]!,text=shapeText(original),end=graphemeBoundaries(text)[2]!;
  assert.throws(()=>replaceShapeText(s.snapshot(),'slide-1','text',2,3,'x',next()),fails('invalid'));
  s.execute(replaceShapeText(s.snapshot(),'slide-1','text',1,end,'🙂',next()));const edited=model(s).slides[0]!.shapes[2]!;assert.equal(shapeText(edited),'A🙂é中文\nالعربية עברית');assert.equal(edited.type,'text');if(edited.type==='text'){assert.equal(edited.runs[0]!.style!.bold,true);assert.equal(edited.runs.at(-1)!.style!.language,'ar');}
  s.undo();assert.deepEqual(model(s).slides[0]!.shapes[2],original);
});
test('SLD-02 empty text insertion and IME grapheme difference never duplicate stable run IDs',()=>{
  const s=sample();s.execute(insertShapes(s.snapshot(),'slide-1',3,[{id:'empty',type:'text',geometry,runs:[{id:'empty-run',type:'text',text:''}]}]));s.execute(replaceShapeText(s.snapshot(),'slide-1','empty',0,0,'你好',next()));assert.equal(shapeText(model(s).slides[0]!.shapes[3]!),'你好');validateDocument(model(s));
  for(const [before,after] of [['a','á'],['A🙂B','A😃B'],['👨‍👩‍👧‍👦test','👩‍👧test'],['ภาษาไทย','ภาษาไทย中文']]){const diff=slideTextEditDiff(before!,after!);assert.ok(graphemeBoundaries(before!).includes(diff.start));assert.ok(graphemeBoundaries(before!).includes(diff.end));assert.equal(before!.slice(0,diff.start)+diff.text+before!.slice(diff.end),after);}
});
test('SLD-02 rich text range formatting splits runs while preserving logical Unicode and paragraph metadata',()=>{
  const s=sample(),before=shapeText(model(s).slides[0]!.shapes[2]!);s.execute(formatShapeText(s.snapshot(),'slide-1','text',{italic:true,fontSize:28*EMU_PER_POINT},{start:0,end:1},next()));
  const shape=model(s).slides[0]!.shapes[2]!;assert.equal(shapeText(shape),before);if(shape.type==='text'){assert.equal(shape.runs[0]!.style!.italic,true);assert.equal(shape.runs[1]!.style!.italic,undefined);}
  s.execute(setShapeTextFrame(s.snapshot(),'slide-1','text',{direction:'rtl',wrap:false,verticalAlignment:'center',margins:{top:0,right:12700,bottom:0,left:12700}}));const scene=buildSlideScene(s.snapshot(),'slide-1'),node=scene.shapes[2]!;assert.equal(node.type,'text');if(node.type==='text'){assert.equal(node.direction,'rtl');assert.equal(node.wrap,false);assert.equal(node.paragraphs.length,2);assert.equal(node.paragraphs[1]!.runs.map(r=>r.text).join(''),'العربية עברית');}
});
test('SLD-02 images remain content addressed, aspect fitted, deduplicated and undoable as one transaction',()=>{
  for(const bytes of [PNG,JPEG]){const s=sample(),image=makeImageResource(bytes),fit=aspectFitImage(image,4000000,2000000);assert.ok(Math.abs(fit.width*image.pixelHeight-fit.height*image.pixelWidth)<=Math.max(image.pixelWidth,image.pixelHeight)/2,'Aspect fit rounds by at most half an EMU');
    s.execute(insertSlideImage(s.snapshot(),'slide-1',3,'picture',image,{...geometry,...fit}));assert.equal(s.historyUsage.items,1);assert.equal(model(s).resources[image.id]!.dataBase64,image.dataBase64);
    s.execute(insertSlideImage(s.snapshot(),'slide-1',4,'picture2',image,geometry));assert.equal(Object.keys(model(s).resources).length,1);s.undo();s.undo();assert.equal(Object.keys(model(s).resources).length,0);assert.equal(s.dirty,false);}
});
test('SLD-02 resource authority prevents replacement and cross-kind patches without changing history',()=>{
  const s=sample(),image=makeImageResource(PNG);s.execute(insertSlideImage(s.snapshot(),'slide-1',3,'picture',image,geometry));const before=s.snapshot(),history=s.historyUsage;
  assert.throws(()=>s.execute(patchCommand(s.snapshot(),'SLD-02',[{op:'set',path:['resources',image.id],value:image}])),fails('conflict'));
  assert.throws(()=>s.execute(patchCommand(s.snapshot(),'SLD-02',[{op:'delete',path:['resources',image.id]}])),fails('unsupported'));
  assert.throws(()=>s.execute(patchCommand(s.snapshot(),'SLD-02',[{op:'set',path:['defaultTextStyle'],value:{fontSize:100}}])),fails('unsupported'));
  assert.deepEqual(s.snapshot(),before);assert.deepEqual(s.historyUsage,history);
});
test('SLD-01/02 paste transfers immutable image payload and model objects across document histories',()=>{
  const s=sample(),image=makeImageResource(PNG);s.execute(insertSlideImage(s.snapshot(),'slide-1',3,'picture',image,geometry));const target=new DocumentSession(createDocument('slides','destination'),{saved:true});target.execute(pasteShapes(target.snapshot(),'slide-1',model(s).slides[0]!.shapes,{[image.id]:image},next()));
  assert.equal(model(target).slides[0]!.shapes.length,4);assert.equal(model(target).resources[image.id]!.sha256,image.sha256);target.undo();assert.equal(model(target).slides[0]!.shapes.length,0);assert.equal(target.dirty,false);assert.equal(model(s).slides[0]!.shapes.length,4);
});
test('SLD-01 scene binds revision, theme, actual dimensions, color tint, source language and center rotation',()=>{
  const s=sample();s.execute(setSlideTheme(s.snapshot(),DEFAULT_SLIDE_THEME));s.execute(setSlideBackground(s.snapshot(),'slide-1',{kind:'theme',name:'lt2'}));s.execute(setSlideSize(s.snapshot(),9144000,6858000));s.execute(setShapePaint(s.snapshot(),'slide-1',['rect'],{fill:{kind:'theme',name:'accent1',tint:.5}}));
  const snapshot=s.snapshot(),scene=buildSlideScene(snapshot,'slide-1');assert.equal(scene.revision,snapshot.revision);assert.equal(scene.contentId,snapshot.contentId);assert.equal(scene.width,9144000);assert.equal(scene.background.red,240);assert.equal(scene.shapes[0]!.fill!.red,161);assert.equal(Object.isFrozen(scene),true);assert.deepEqual(resolveSlideColor({kind:'theme',name:'text1'},DEFAULT_SLIDE_THEME),DEFAULT_SLIDE_THEME.colors.dk1);
  const rotated:Shape={id:'rotated',type:'rect',geometry:{...geometry,rotation:5400000}},point=transform(shapeTransform(rotated.geometry),{x:1500000,y:500000});assert.equal(hitTest([rotated],point),'rotated');
  const oldRevision=scene.revision;s.execute(setSlideSize(s.snapshot(),12192000,6858000));assert.equal(scene.revision,oldRevision);assert.equal(scene.width,9144000);
});
test('SLD-01 schema rejects invalid text-frame and cyclic theme colors without degrading source',()=>{
  const s=sample(),before=s.snapshot();assert.throws(()=>s.execute(setShapeTextFrame(s.snapshot(),'slide-1','text',{margins:{top:-1,right:0,bottom:0,left:0}})),fails('invalid'));
  assert.throws(()=>s.execute(setSlideTheme(s.snapshot(),{name:'invalid',colors:{accent1:{kind:'theme',name:'accent1'} as never}})),fails('invalid'));
  assert.deepEqual(s.snapshot(),before);assert.throws(()=>buildSlideScene(s.snapshot(),'missing'),fails('invalid'));assert.throws(()=>resolveSlideColor({kind:'theme',name:'unknown'}),fails('unsupported'));
});

test('Paragraph list commands affect only selected paragraphs, retain text and undo; numbering follows level',()=>{
 const d=createDocument('slides','list-transaction');d.slides[0]!.shapes=[{id:'text',type:'text',geometry,runs:[{id:'r',type:'text',text:'One\nTwo\nThree\nFour'}]}];
 const s=new DocumentSession(d,{saved:true}),ids=next();
 s.execute(formatSlideParagraphs(s.snapshot(),'slide-1','text',{start:0,end:18},{marginLeft:342900,firstLineIndent:-228600,bullet:{kind:'number',scheme:'arabicPeriod'}},ids));
 s.execute(formatSlideParagraphs(s.snapshot(),'slide-1','text',{start:4,end:7},{level:1,marginLeft:800100},ids));
 const scene=buildSlideScene(s.snapshot(),'slide-1'),shape=scene.shapes[0]!;
 assert.equal(shape.type,'text');if(shape.type!=='text')return;
 assert.deepEqual(shape.paragraphs.map(p=>p.marker?.text),['1.','1.','2.','3.']);
 assert.equal(shapeText(model(s).slides[0]!.shapes[0]!),'One\nTwo\nThree\nFour');
 s.undo();s.undo();assert.deepEqual(s.snapshot().document,d);assert.equal(s.dirty,false);
});

test('Presentation validation rejects conflicting bullet size and unbounded fit scale',()=>{
 const d=createDocument('slides','list-validation');d.slides[0]!.shapes=[{id:'text',type:'text',geometry,runs:[{id:'r',type:'text',text:'Test',style:{slideParagraph:{bullet:{kind:'char',character:'•',sizePercent:100000,sizeEmu:254000}}}}]}];
 assert.throws(()=>validateDocument(d),fails('invalid'));
 const shape=d.slides[0]!.shapes[0] as Extract<Shape,{type:'text'}>;shape.runs[0]!.style={};shape.textFrame={autofit:{mode:'normal',fontScale:0,lineReduction:0}};assert.throws(()=>validateDocument(d),fails('invalid'));
});

test('Nested graphic grouping, resource-aware duplication, exact rigid ungroup and undo',()=>{
 const d=createDocument('slides','graphic-group'),asset=makeImageResource(PNG);d.resources[asset.id]=asset;
 d.slides[0]!.shapes=[
  {id:'triangle-a',type:'triangle',geometry:{x:100000,y:200000,width:1000000,height:1200000,rotation:0}},
  {id:'image-a',type:'image',geometry:{x:1500000,y:200000,width:1000000,height:1200000,rotation:0},resourceId:asset.id,crop:{left:10000,right:10000,top:0,bottom:0}},
  {id:'last',type:'diamond',geometry:{x:4000000,y:1000000,width:1000000,height:1000000,rotation:0}}
 ];
 const s=new DocumentSession(d,{saved:true}),ids=next();
 assert.throws(()=>groupShapes(s.snapshot(),'slide-1',['triangle-a','last'],ids),fails('unsupported'));
 s.execute(groupShapes(s.snapshot(),'slide-1',['triangle-a','image-a'],ids));
 const group=model(s).slides[0]!.shapes[0]!;assert.equal(group.type,'group');if(group.type!=='group')return;
 assert.deepEqual(group.children,d.slides[0]!.shapes.slice(0,2));assert.equal(canUngroupShape(group),true);
 s.execute(duplicateShapes(s.snapshot(),'slide-1',[group.id],ids));
 const duplicate=model(s).slides[0]!.shapes.at(-1)!;assert.equal(duplicate.type,'group');if(duplicate.type!=='group')return;
 assert.notEqual(duplicate.children[0]!.id,group.children[0]!.id);
 const target=createDocument('slides','target-graphics'),dest=new DocumentSession(target);
 dest.execute(pasteShapes(dest.snapshot(),'slide-1',[group],{[asset.id]:asset},ids));assert.ok(model(dest).resources[asset.id]);assert.equal(model(dest).slides[0]!.shapes[0]!.type,'group');
 s.undo();s.execute(ungroupShape(s.snapshot(),'slide-1',group.id));
 assert.deepEqual(model(s).slides[0]!.shapes,d.slides[0]!.shapes);
 s.undo();s.undo();assert.deepEqual(model(s),d);
});

test('Crop, flip, shape hit geometry and unsupported scaled ungroup are transactional',()=>{
 const d=createDocument('slides','graphic-edit'),asset=makeImageResource(PNG);d.resources[asset.id]=asset;
 d.slides[0]!.shapes=[{id:'photo',type:'image',geometry,resourceId:asset.id},{id:'triangle',type:'triangle',geometry:{...geometry,rotation:0}}];
 const s=new DocumentSession(d,{saved:true});
 s.execute(setSlideImageCrop(s.snapshot(),'slide-1','photo',{left:12500,right:5000,top:2500,bottom:0}));
 assert.equal((model(s).slides[0]!.shapes[0] as Extract<Shape,{type:'image'}>).crop?.left,12500);
 s.execute(flipSlideShapes(s.snapshot(),'slide-1',['triangle'],'flipV'));
 const triangle=model(s).slides[0]!.shapes[1]!;
 assert.equal(hitTest([triangle],{x:geometry.x+geometry.width*.1,y:geometry.y+geometry.height*.05}),'triangle');
 assert.equal(hitTest([triangle],{x:geometry.x+geometry.width*.1,y:geometry.y+geometry.height*.95}),undefined);
 const scene=buildSlideScene(s.snapshot(),'slide-1');assert.equal(scene.shapes[1]!.flipV,true);
 assert.throws(()=>setSlideImageCrop(s.snapshot(),'slide-1','photo',{left:60000,right:40000,top:0,bottom:0}),fails('invalid'));
 s.execute(groupShapes(s.snapshot(),'slide-1',['photo','triangle'],next()));
 const group=model(s).slides[0]!.shapes[0]!;s.execute(setShapeGeometries(s.snapshot(),'slide-1',{[group.id]:{...group.geometry,width:group.geometry.width*2}}));
 assert.throws(()=>ungroupShape(s.snapshot(),'slide-1',group.id),fails('unsupported'));
 assert.equal(s.snapshot().document.resources[asset.id]!.sha256,asset.sha256);
});

test('Nested child editing preserves group transforms, sibling order and undo',()=>{
 const d=createDocument('slides','nested-editor');d.slides[0]!.shapes=[{id:'g',type:'group',geometry:{x:1000000,y:2000000,width:4000000,height:3000000,rotation:1800000},flipH:true,childSpace:{x:500000,y:600000,width:2000000,height:3000000},children:[
 {id:'a',type:'rect',geometry:{...geometry,x:550000,y:650000,width:600000,height:900000}},
 {id:'b',type:'text',geometry:{...geometry,x:1300000,y:1700000,width:700000,height:900000},runs:[{id:'r',type:'text',text:'Before'}]}
 ]}];
 const session=new DocumentSession(d,{saved:true}),ids=next(),original=session.snapshot(),g=d.slides[0]!.shapes[0]!;
 assert.equal(g.type,'group');if(g.type!=='group')return;
 const parent=slideShapeParentTransform(d.slides[0]!.shapes,'b')!;assert.deepEqual(parent,groupChildTransform(g));
 const world=transform(parent,{x:1300000,y:1700000}),local=transform(inverse(parent),world);assert.ok(Math.abs(local.x-1300000)<.001&&Math.abs(local.y-1700000)<.001);
 session.execute(replaceShapeText(session.snapshot(),'slide-1','b',0,6,'Edited',ids));
 session.execute(setShapeGeometries(session.snapshot(),'slide-1',{a:{...g.children[0]!.geometry,x:650000}}));
 session.execute(reorderShapes(session.snapshot(),'slide-1',['b'],'back'));
 const location=locateShape(session.snapshot(),'slide-1','b');assert.deepEqual(location.path,['slides',0,'shapes',0,'children',0]);assert.equal(shapeText(location.shape),'Edited');
 const after=model(session).slides[0]!.shapes[0]!;assert.equal(after.type,'group');if(after.type!=='group')return;
 assert.deepEqual(after.geometry,g.geometry);assert.deepEqual(after.childSpace,g.childSpace);assert.equal(after.children[1]!.geometry.x,650000);
 session.execute(duplicateShapes(session.snapshot(),'slide-1',['b'],ids));assert.equal(locateShape(session.snapshot(),'slide-1','b').siblings.length,3);
 session.execute(removeShapes(session.snapshot(),'slide-1',['b']));assert.equal((model(session).slides[0]!.shapes[0] as Extract<Shape,{type:'group'}>).children.length,2);
 for(let i=0;i<5;i++)session.undo();assert.deepEqual(session.snapshot().document,original.document);
});

test('Attached connectors follow moving/rotated targets, remap complete copies and detach on delete',()=>{
 const d=createDocument('slides','connector-editor');d.slides[0]!.shapes=[
 {id:'a',type:'rect',geometry:{x:1000000,y:1000000,width:1000000,height:1000000,rotation:0}},
 {id:'b',type:'rect',geometry:{x:5000000,y:1000000,width:1000000,height:1000000,rotation:0}}
 ];const s=new DocumentSession(d,{saved:true}),ids=next();
 s.execute(connectSlideShapes(s.snapshot(),'slide-1',['a','b'],ids));
 const line=model(s).slides[0]!.shapes[2]!;assert.equal(line.type,'line');if(line.type!=='line')return;
 assert.deepEqual(line.connections,{start:{shapeId:'a',site:3},end:{shapeId:'b',site:1}});
 assert.deepEqual(buildSlideScene(s.snapshot(),'slide-1').shapes[2]!.geometry,{x:2000000,y:1500000,width:3000000,height:0,rotation:0});
 const target=locateShape(s.snapshot(),'slide-1','b').shape;
 s.execute(setShapeGeometries(s.snapshot(),'slide-1',{b:{...target.geometry,y:2500000,rotation:5400000}}));
 const derived=buildSlideScene(s.snapshot(),'slide-1').shapes[2]!.geometry;
 assert.notDeepEqual(derived,line.geometry);
 assert.throws(()=>setShapeGeometries(s.snapshot(),'slide-1',{[line.id]:line.geometry}),fails('unsupported'));
 assert.throws(()=>groupShapes(s.snapshot(),'slide-1',['a','b'],ids),fails('unsupported'));
 s.execute(duplicateSlide(s.snapshot(),'slide-1',ids));
 const copy=model(s).slides[1]!,copied=copy.shapes[2]!;assert.equal(copied.type,'line');if(copied.type!=='line')return;
 assert.equal(copied.connections?.start.shapeId,copy.shapes[0]!.id);assert.equal(copied.connections?.end.shapeId,copy.shapes[1]!.id);
 s.undo();s.execute(duplicateShapes(s.snapshot(),'slide-1',[line.id],ids));const free=model(s).slides[0]!.shapes.at(-1)!;
 assert.equal(free.type,'line');if(free.type!=='line')return;assert.equal(free.connections,undefined);assert.equal(free.geometry.x,derived.x+190500);
 s.undo();s.execute(removeShapes(s.snapshot(),'slide-1',['b']));const detached=model(s).slides[0]!.shapes[1]!;assert.equal(detached.type,'line');if(detached.type!=='line')return;
 assert.equal(detached.connections,undefined);assert.deepEqual(detached.geometry,derived);s.undo();
 s.execute(setSlideConnections(s.snapshot(),'slide-1',line.id,null));const unbound=locateShape(s.snapshot(),'slide-1',line.id).shape;assert.equal(unbound.type,'line');if(unbound.type==='line')assert.equal(unbound.connections,undefined);
});
