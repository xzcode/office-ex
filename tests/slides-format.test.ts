// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import {createDocument,resolveSlideThemeColor,DEFAULT_SLIDE_THEME_COLORS,makeImageResource,walkSlideShapes,resolveSlideConnector,type SlidesDocument,type Shape} from '@office-ex/core';
import {OfficeError} from '@office-ex/contracts';
import {openOpc,writeZip,openSlidesPptx,writeSlidesPptx,saveSlidesPptx} from '@office-ex/ooxml';
import {openLegacyPpt,writeLegacyPpt,saveLegacyPpt,prepareSlidesSave,serializeDraft,inspectForEditor} from '@office-ex/format-service';
import {readCfb,writeCfb,type CfbEntry} from '@office-ex/cfb';
import {slidesFormatFixture} from './slides-format-fixture.js';
import {PNG} from './writer-image-fixture.js';

const text=(s:Shape)=>s.type==='text'?s.runs.map(r=>r.text).join(''):'';
const fail=(code?:string)=>(e:unknown)=>e instanceof OfficeError&&(!code||e.code===code);
const formats=[['pptx',writeSlidesPptx,openSlidesPptx,saveSlidesPptx,0],['ppt',writeLegacyPpt,openLegacyPpt,saveLegacyPpt,1588]] as const;
function scene(actual:SlidesDocument,expected:SlidesDocument,tolerance:number):void{
 assert.equal(actual.slides.length,expected.slides.length);
 for(const k of ['width','height'] as const)assert.ok(Math.abs(actual.size[k]-expected.size[k])<=tolerance);
 for(const [i,a]of actual.slides.entries()){
  const e=expected.slides[i]!;assert.deepEqual(resolveSlideThemeColor(a.background,actual.theme),resolveSlideThemeColor(e.background,expected.theme));
  assert.deepEqual(a.shapes.map(s=>s.type),e.shapes.map(s=>s.type));
  for(const [j,s]of a.shapes.entries()){
   const wanted=e.shapes[j]!;for(const k of ['x','y','width','height'] as const)assert.ok(Math.abs(s.geometry[k]-wanted.geometry[k])<=tolerance,`${i}/${j}/${k}`);
   assert.ok(Math.abs(s.geometry.rotation-wanted.geometry.rotation)<=1);assert.equal(text(s),text(wanted));
   if(wanted.fill)assert.deepEqual(s.fill&&resolveSlideThemeColor(s.fill,actual.theme),resolveSlideThemeColor(wanted.fill,expected.theme));
   if(wanted.stroke){assert.equal(s.stroke?.width,wanted.stroke.width);assert.deepEqual(s.stroke?.color&&resolveSlideThemeColor(s.stroke.color,actual.theme),resolveSlideThemeColor(wanted.stroke.color,expected.theme));}
   if(s.type==='text'){
    assert.equal(s.textFrame?.verticalAlignment,'center');assert.deepEqual(s.textFrame?.margins,{left:91440,right:91440,top:45720,bottom:45720});
    for(const run of s.runs){assert.equal(run.style?.fontSize,254000);assert.equal(run.style?.bold,true);assert.equal(run.style?.fontFamily,'Arial');assert.equal(run.style?.alignment,'center');assert.equal(run.style?.paragraphDirection,i===1?'rtl':'ltr');assert.equal(run.style?.language,i===1?'ar-SA':'zh-CN');}
   }
   if(s.type==='image'&&wanted.type==='image')assert.equal(actual.resources[s.resourceId]?.sha256,expected.resources[wanted.resourceId]?.sha256);
  }
 }
}
function cfbChange(bytes:Uint8Array,name:string,change:(b:Uint8Array)=>void):Uint8Array{
 const cfb=readCfb(bytes);return writeCfb({...cfb,entries:cfb.entries.map(e=>{if(e.type!=='stream'||e.path[0]!==name)return e;const data=e.data.slice();change(data);return {...e,data};})});
}
function records(b:Uint8Array,type:number,start=0,end=b.length):number[]{
 const v=new DataView(b.buffer,b.byteOffset,b.length),out:number[]=[];for(let at=start;at<end;){const len=v.getUint32(at+4,true),next=at+8+len;if(v.getUint16(at+2,true)===type)out.push(at);if((v.getUint16(at,true)&15)===15)out.push(...records(b,type,at+8,next));at=next;}return out;
}
for(const[format,write,open,save,tolerance]of formats){
 test(format+' independently editable three-slide scene and native reopen',()=>{
  const d=slidesFormatFixture(),b=write(d),r=open(b,d.documentId,d.title);scene(r.document,d,tolerance);assert.deepEqual(save(b,r.document),b);
  r.document.slides[0]!.shapes[0]!.geometry.x+=914400;const saved=save(b,r.document);scene(open(saved,d.documentId).document,r.document,tolerance);
  const again=open(saved,d.documentId).document;again.slides.reverse();assert.deepEqual(open(save(saved,again),d.documentId).document.slides.map(s=>s.background),again.slides.map(s=>s.background));
 });
 test(format+' main-owned preparation authenticates immutable session and kind',async()=>{
  const d=slidesFormatFixture(),bytes=prepareSlidesSave(format,serializeDraft(d),d.documentId);const view=await inspectForEditor(bytes,'deck.'+format,d.documentId);assert.equal(view.report.readOnly,false);assert.equal(view.editableSlidesFormat,format);assert.equal(view.editableSlides?.slides.length,3);
  assert.throws(()=>prepareSlidesSave(format,serializeDraft(d),'other'),fail('invalid'));assert.throws(()=>prepareSlidesSave(format,serializeDraft(createDocument('writer','writer')),'writer'),fail('invalid'));
  const abort=new AbortController();abort.abort();assert.throws(()=>open(bytes,d.documentId,'cancelled',{signal:abort.signal}),fail('cancelled'));
 });
 test(format+' creates an ordinary blank deck without image/text surrogate',()=>{
  const d=createDocument('slides','blank');const r=open(write(d),d.documentId).document;assert.equal(r.slides.length,1);assert.equal(r.slides[0]!.shapes.length,0);
 });
 test(format+' RTL logical alignment and theme colors match the shared scene',()=>{
  const d=createDocument('slides','rtl');d.slides[0]!.background={kind:'theme',name:'background1'};
  d.slides[0]!.shapes=[{id:'text',type:'text',geometry:{x:0,y:0,width:9144000,height:914400,rotation:0},textStyle:{alignment:'start',paragraphDirection:'rtl'},runs:[{id:'r',type:'text',text:'مرحبا שלום',style:{color:{kind:'theme',name:'accent2'},language:'ar'}}]}];
  const r=open(write(d),d.documentId).document,s=r.slides[0]!.shapes[0]!;assert.equal(s.type,'text');if(s.type!=='text')return;assert.equal(s.runs[0]!.style?.alignment,'start');assert.equal(s.runs[0]!.style?.paragraphDirection,'rtl');assert.deepEqual(resolveSlideThemeColor(s.runs[0]!.style!.color!,r.theme),resolveSlideThemeColor({kind:'theme',name:'accent2'}));
 });
}
test('PPTX ↔ genuine PPT conversion preserves three distinct editable scenes and images',()=>{
 const d=slidesFormatFixture(),pptx=openSlidesPptx(writeSlidesPptx(d),d.documentId).document,ppt=openLegacyPpt(writeLegacyPpt(pptx),d.documentId).document;scene(ppt,d,1588);scene(openSlidesPptx(writeSlidesPptx(ppt),d.documentId).document,d,1588);
 assert.deepEqual([...writeLegacyPpt(d).subarray(0,8)],[208,207,17,224,161,177,26,225]);assert.equal(String.fromCharCode(...writeSlidesPptx(d).subarray(0,2)),'PK');
});
test('PPT native save preserves old stream history and updates the genuine UserEdit chain',()=>{
 const b=writeLegacyPpt(slidesFormatFixture()),source=openLegacyPpt(b,'slides-fixture'),d=source.document;d.slides[0]!.shapes[0]!.geometry.x+=914400;const after=openLegacyPpt(saveLegacyPpt(b,d),d.documentId);
 assert.deepEqual(after.main.subarray(0,source.main.length),source.main);assert.deepEqual(after.pictures.subarray(0,source.pictures.length),source.pictures);assert.ok(after.currentEdit>source.currentEdit);
 assert.equal(new DataView(after.main.buffer,after.main.byteOffset).getUint32(after.currentEdit+16,true),source.currentEdit);
 const currentBefore=source.current.slice(),currentAfter=after.current.slice();currentBefore.fill(0,16,20);currentAfter.fill(0,16,20);assert.deepEqual(currentBefore,currentAfter);
});
test('PPTX native save changes only the edited slide part and keeps source style matrices',()=>{
 const parts=new Map(openOpc(writeSlidesPptx(slidesFormatFixture())).parts),theme='ppt/theme/theme1.xml';
 const xml=new TextDecoder().decode(parts.get(theme)!);parts.set(theme,new TextEncoder().encode(xml.replace('<a:effectLst/>','<a:effectLst><a:glow rad="12700"><a:srgbClr val="ABCDEF"/></a:glow></a:effectLst>')));
 const b=writeZip(parts),d=openSlidesPptx(b,'slides-fixture').document;d.slides[0]!.shapes[0]!.geometry.x+=914400;const after=openOpc(saveSlidesPptx(b,d));
 for(const[path,bytes]of parts)if(path!=='ppt/slides/slide1.xml')assert.deepEqual(after.parts.get(path),bytes,path);
});
test('PPTX parses nonstandard part URIs/prefixes and safely saves those source paths',()=>{
 const p=openOpc(writeSlidesPptx(slidesFormatFixture())),parts=new Map<string,Uint8Array>();for(const[path,b]of p.parts){let bytes=b;if(path.endsWith('.xml')||path.endsWith('.rels'))bytes=new TextEncoder().encode(new TextDecoder().decode(b).replaceAll('ppt/','deck/').replaceAll('xmlns:p=','xmlns:present=').replaceAll('<p:','<present:').replaceAll('</p:','</present:'));parts.set(path.replace(/^ppt\//,'deck/'),bytes);}
 const b=writeZip(parts),d=openSlidesPptx(b,'slides-fixture').document;d.slides[1]!.shapes[0]!.geometry.x+=914400;const saved=saveSlidesPptx(b,d),a=openOpc(saved);assert.equal(a.mainPart,'deck/presentation.xml');assert.ok(a.parts.has('deck/slides/slide2.xml'));assert.equal(openSlidesPptx(saved,d.documentId).document.slides[1]!.shapes[0]!.geometry.x,1828800);
});
test('unknown source features remain read-only and cannot authorize semantic saves',async()=>{
 const d=slidesFormatFixture(),p=new Map(openOpc(writeSlidesPptx(d)).parts);p.set('custom.xml',new TextEncoder().encode('<custom/>'));const x=writeZip(p),cfb=readCfb(writeLegacyPpt(d)),y=writeCfb({...cfb,entries:[...cfb.entries,{type:'stream',path:['Unknown'],data:Uint8Array.of(1)} as CfbEntry]});
 for(const[format,bytes]of [['pptx',x],['ppt',y]]as const){assert.equal((await inspectForEditor(bytes,'unknown.'+format,d.documentId)).report.readOnly,true);assert.throws(()=>prepareSlidesSave(format,serializeDraft(d),d.documentId,bytes),fail());}
});
test('PPTX refuses inherited objects, unresolved fill, external images and malformed colors',()=>{
 const original=openOpc(writeSlidesPptx(slidesFormatFixture()));const change=(path:string,from:string,to:string)=>{const p=new Map(original.parts);p.set(path,new TextEncoder().encode(new TextDecoder().decode(p.get(path)!).replace(from,to)));return writeZip(p);};
 assert.throws(()=>openSlidesPptx(change('ppt/slideLayouts/slideLayout1.xml','</p:spTree>','<p:sp/></p:spTree>'),'x'),fail());
 assert.throws(()=>openSlidesPptx(change('ppt/slides/slide1.xml','<a:noFill/>',''),'x'),fail());
 assert.throws(()=>openSlidesPptx(change('ppt/slides/slide1.xml','<a:srgbClr val="F4F7FB"></a:srgbClr>','<a:srgbClr val="F4F7FB"/><a:schemeClr val="dk1"/>'),'x'),fail());
 assert.throws(()=>openSlidesPptx(change('ppt/slides/_rels/slide1.xml.rels','Target="../media/','TargetMode="External" Target="https://example.invalid/'),'x'),fail());
});
test('PPT rejects unknown live records, encryption, invalid persist offsets and cyclic edits',()=>{
 const b=writeLegacyPpt(slidesFormatFixture()),source=openLegacyPpt(b,'deck');
 assert.throws(()=>openLegacyPpt(cfbChange(b,'Current User',x=>new DataView(x.buffer).setUint32(12,0xf3d1c4df,true)),'x'),fail());
 assert.throws(()=>openLegacyPpt(cfbChange(b,'PowerPoint Document',x=>new DataView(x.buffer).setUint32(source.currentEdit+16,source.currentEdit,true)),'x'),fail());
 assert.throws(()=>openLegacyPpt(cfbChange(b,'PowerPoint Document',x=>{const v=new DataView(x.buffer),at=v.getUint32(source.currentEdit+20,true);v.setUint32(at+12,x.length+10,true);}),'x'),fail());
 assert.throws(()=>openLegacyPpt(cfbChange(b,'PowerPoint Document',x=>{const at=records(x,3999)[0]!;new DataView(x.buffer).setUint16(at+2,0x7777,true);}),'x'),fail());
});
test('PPT locale markers cover all 22 product locales; unknown LCID cannot silently convert',()=>{
 const locales=['zh-CN','zh-TW','en-US','ja-JP','ko-KR','fr-FR','de-DE','es-ES','pt-BR','ru-RU','it-IT','ar','he-IL','hi-IN','th-TH','vi-VN','id-ID','ms-MY','tr-TR','nl-NL','pl-PL','uk-UA'];
 for(const language of locales){const d=createDocument('slides','locale');d.slides[0]!.shapes=[{id:'text',type:'text',geometry:{x:0,y:0,width:9144000,height:914400,rotation:0},runs:[{id:'r',type:'text',text:'Hello 你好 مرحبا',style:{language}}]}];const r=openLegacyPpt(writeLegacyPpt(d),d.documentId).document.slides[0]!.shapes[0]!;assert.equal(r.type,'text');if(r.type==='text')assert.equal(r.runs[0]!.style?.language,language==='ar'?'ar-SA':language);}
 const bytes=cfbChange(writeLegacyPpt(slidesFormatFixture()),'PowerPoint Document',b=>{const at=records(b,4010).at(-1)!;new DataView(b.buffer).setUint16(at+16,0x7fff,true);}),d=openLegacyPpt(bytes,'unknown-lcid').document;assert.throws(()=>writeSlidesPptx(d),fail('unsupported'));assert.doesNotThrow(()=>writeLegacyPpt(d));
});
test('target limitations are explicit and deleting an image does not strand the document',()=>{
 const d=slidesFormatFixture();const shape=d.slides[0]!.shapes.find(s=>s.type==='text')!;if(shape.type!=='text')return;shape.runs[0]!.style!.fontSize=260350;assert.throws(()=>writeLegacyPpt(d),fail('unsupported'));assert.doesNotThrow(()=>writeSlidesPptx(d));
 shape.runs[0]!.style!.fontSize=254000;for(const slide of d.slides)slide.shapes=slide.shapes.filter(s=>s.type!=='image');for(const[,write,open]of formats)assert.equal(open(write(d),d.documentId).document.slides[0]!.shapes.length,4);
});


test('PPTX retains standard dependency-free core properties across native saves',async()=>{
 const source=openOpc(writeSlidesPptx(slidesFormatFixture())),parts=new Map(source.parts),encode=(text:string)=>new TextEncoder().encode(text),decode=(bytes:Uint8Array)=>new TextDecoder().decode(bytes);
 const path='metadata/external-core.xml',core='<c:coreProperties xmlns:c="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:d="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>External deck 中文</dc:title><dc:creator>Independent author</dc:creator><c:revision>3</c:revision><d:modified xsi:type="d:W3CDTF">2026-11-25T00:00:00Z</d:modified></c:coreProperties>';
 const rootRel=decode(parts.get('_rels/.rels')!).replace('</Relationships>','<Relationship Id="externalMetadata" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="metadata/external-core.xml"/></Relationships>');
 parts.set('_rels/.rels',encode(rootRel));parts.set(path,encode(core));parts.set('[Content_Types].xml',encode(decode(parts.get('[Content_Types].xml')!).replace('</Types>','<Override PartName="/metadata/external-core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>')));
 const input=writeZip(parts),opened=openSlidesPptx(input,'external-core');assert.deepEqual(opened.native.metadataPaths,[path]);assert.deepEqual(saveSlidesPptx(input,opened.document),input);
 opened.document.slides[0]!.shapes[0]!.geometry.x+=914400;opened.document.slides.pop();const output=saveSlidesPptx(input,opened.document),saved=openOpc(output);assert.deepEqual(saved.parts.get(path),parts.get(path));assert.deepEqual(saved.parts.get('_rels/.rels'),parts.get('_rels/.rels'));assert.equal(openSlidesPptx(output,'external-core').document.slides.length,2);
 for(const invalid of [core.replace('</dc:title>','</dc:title><dc:title>Duplicate</dc:title>'),core.replace('Independent author','<c:category>nested</c:category>'),core.replace('xsi:type="d:W3CDTF"','xsi:type="c:W3CDTF"')]){const changed=new Map(parts);changed.set(path,encode(invalid));const bytes=writeZip(changed);assert.throws(()=>openSlidesPptx(bytes,'external-core'),fail());assert.equal((await inspectForEditor(bytes,'external.pptx','external-core')).report.readOnly,true);}
});


test('PPTX resolves solid master/layout backgrounds and keeps each slide layout during edit, reorder and insertion',async()=>{
 const source=openOpc(writeSlidesPptx(slidesFormatFixture())),parts=new Map(source.parts),encode=(text:string)=>new TextEncoder().encode(text),decode=(bytes:Uint8Array)=>new TextDecoder().decode(bytes);
 const master='ppt/slideMasters/slideMaster1.xml',masterRels='ppt/slideMasters/_rels/slideMaster1.xml.rels',blank='ppt/slideLayouts/slideLayout1.xml',blankRels='ppt/slideLayouts/_rels/slideLayout1.xml.rels';
 const bg=(hex:string)=>'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="'+hex+'"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>';
 const second='ppt/slideLayouts/slideLayout2.xml',inactive='ppt/slideLayouts/slideLayout3.xml';
 const blankXml=decode(parts.get(blank)!);parts.set(second,encode(blankXml.replace('<p:spTree>',bg('AABBCC')+'<p:spTree>')));
 // This unused source layout contains a genuine placeholder rather than an editable scene.
 const placeholder='<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>';
 parts.set(inactive,encode(blankXml.replace('type="blank"','type="title"').replace('</p:spTree>',placeholder+'</p:spTree>')));
 for(const path of ['ppt/slideLayouts/_rels/slideLayout2.xml.rels','ppt/slideLayouts/_rels/slideLayout3.xml.rels'])parts.set(path,parts.get(blankRels)!);
 parts.set(master,encode(decode(parts.get(master)!).replace('<p:spTree>',bg('112233')+'<p:spTree>').replace('</p:sldLayoutIdLst>','<p:sldLayoutId id="2147483650" r:id="rIdLayout2"/><p:sldLayoutId id="2147483651" r:id="rIdLayout3"/></p:sldLayoutIdLst>')));
 parts.set(masterRels,encode(decode(parts.get(masterRels)!).replace('</Relationships>','<Relationship Id="rIdLayout2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/><Relationship Id="rIdLayout3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout3.xml"/></Relationships>')));
 parts.set('[Content_Types].xml',encode(decode(parts.get('[Content_Types].xml')!).replace('</Types>',[second,inactive].map(path=>'<Override PartName="/'+path+'" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>').join('')+'</Types>')));
 const slide2Rels='ppt/slides/_rels/slide2.xml.rels';parts.set(slide2Rels,encode(decode(parts.get(slide2Rels)!).replace('slideLayout1.xml','slideLayout2.xml')));
 for(let i=1;i<=2;i++){const path='ppt/slides/slide'+i+'.xml';parts.set(path,encode(decode(parts.get(path)!).replace(/<p:bg>[\s\S]*?<\/p:bg>/,'')));}
 const input=writeZip(parts),opened=openSlidesPptx(input,'multi-layout');const inspection=await inspectForEditor(input,'multi-layout.pptx','multi-layout');assert.equal(inspection.report.readOnly,false);assert.equal(inspection.slidesTemplateConversionBlocked,true);assert.deepEqual(opened.native.layoutPaths,[blank,second,inactive]);assert.deepEqual(opened.native.editableLayoutPaths,[blank,second]);assert.deepEqual(opened.native.slideLayoutPaths,[blank,second,blank]);assert.deepEqual(opened.native.slideInheritedBackgrounds,[true,true,false]);
 assert.deepEqual(opened.document.slides.slice(0,2).map(s=>s.background),[{kind:'rgba',red:17,green:34,blue:51,alpha:255},{kind:'rgba',red:170,green:187,blue:204,alpha:255}]);
 opened.document.slides[1]!.shapes[0]!.geometry.x+=914400;opened.document.slides.reverse();const added={id:'new-slide',background:structuredClone(opened.document.slides[0]!.background),shapes:[]};opened.document.slides.push(added);
 const output=saveSlidesPptx(input,opened.document),saved=openOpc(output),reopened=openSlidesPptx(output,'multi-layout');assert.deepEqual(reopened.native.slideLayoutPaths,[blank,second,blank,blank]);assert.deepEqual(reopened.native.slideInheritedBackgrounds,[false,true,true,false]);assert.equal(reopened.document.slides[1]!.shapes[0]!.geometry.x,opened.document.slides[1]!.shapes[0]!.geometry.x);assert.deepEqual(reopened.document.slides.map(s=>s.background),opened.document.slides.map(s=>s.background));
 for(const path of [master,masterRels,blank,blankRels,second,inactive,'ppt/slideLayouts/_rels/slideLayout2.xml.rels','ppt/slideLayouts/_rels/slideLayout3.xml.rels'])assert.deepEqual(saved.parts.get(path),parts.get(path),path);
 const protectedParts=new Map(parts);protectedParts.set(slide2Rels,encode(decode(parts.get(slide2Rels)!).replace('slideLayout2.xml','slideLayout3.xml')));const protectedBytes=writeZip(protectedParts);assert.throws(()=>openSlidesPptx(protectedBytes,'protected-layout'),fail());assert.equal((await inspectForEditor(protectedBytes,'protected.pptx','protected-layout')).report.readOnly,true);
});


test('PPT and PPTX preserve soft line breaks and 150% paragraph spacing without turning them into paragraphs',()=>{
 const d=createDocument('slides','soft-break-spacing');d.slides[0]!.shapes=[{id:'text',type:'text',geometry:{x:0,y:0,width:9144000,height:1828800,rotation:0},runs:[{id:'r1',type:'text',text:'First\u2028Second\nThird',style:{fontSize:254000,lineMultiple:360}}]}];
 for(const[,write,open]of formats){const model=open(write(d),d.documentId).document;const shape=model.slides[0]!.shapes[0]!;assert.equal(shape.type,'text');if(shape.type==='text'){assert.equal(shape.runs.map(r=>r.text).join(''),'First\u2028Second\nThird');assert.ok(shape.runs.every(r=>r.style?.lineMultiple===360));}}
});

test('DrawingML theme tint/shade and script fonts remain source-semantic across PPTX round trips',()=>{
 const d=createDocument('slides','theme-maps');d.theme={name:'External theme',colors:{...DEFAULT_SLIDE_THEME_COLORS,accent1:{kind:'rgba',red:0,green:255,blue:0,alpha:255}},majorFont:'Calibri',minorFont:'Arial',majorFontDetails:{scripts:{Hans:'Songti SC',Arab:'Times New Roman'}},minorFontDetails:{scripts:{Hant:'MingLiU'}}};
 d.slides[0]!.shapes=[{id:'text',type:'text',geometry:{x:0,y:0,width:9144000,height:1828800,rotation:0},fill:{kind:'theme',name:'accent1',tint:.5,tintSpace:'linear'},runs:[{id:'r1',type:'text',text:'中文',style:{fontFamily:'+mj-lt',fontEastAsia:'+mj-ea',fontComplex:'+mj-cs',fontSize:254000,language:'zh-CN'}}]}];
 const after=openSlidesPptx(writeSlidesPptx(d),d.documentId).document;assert.deepEqual(after.theme?.majorFontDetails,d.theme.majorFontDetails);const shape=after.slides[0]!.shapes[0]!;assert.deepEqual(shape.fill,d.slides[0]!.shapes[0]!.fill);assert.deepEqual(resolveSlideThemeColor(shape.fill!,after.theme),{kind:'rgba',red:188,green:255,blue:188,alpha:255});assert.equal(shape.type,'text');if(shape.type==='text')assert.equal(shape.runs[0]!.style?.fontEastAsia,'+mj-ea');assert.throws(()=>writeLegacyPpt(after),fail('unsupported'));
});

test('PPTX multilevel bullet properties, decimal numbering, percentage spacing and cached fit round trip',()=>{
 const d=createDocument('slides','rich-lists');d.slides[0]!.shapes=[{id:'text',type:'text',geometry:{x:0,y:0,width:9144000,height:3000000,rotation:0},textFrame:{autofit:{mode:'normal',fontScale:85000,lineReduction:12000}},runs:[
 {id:'r1',type:'text',text:'First',style:{fontSize:254000,slideParagraph:{level:0,marginLeft:342900,firstLineIndent:-228600,beforePercent:20000,bullet:{kind:'char',character:'•',font:'Arial',sizePercent:90000,color:{kind:'rgba',red:50,green:60,blue:70,alpha:255}}}}},
 {id:'r2',type:'text',text:'\nSecond',style:{fontSize:254000,slideParagraph:{level:1,marginLeft:800100,firstLineIndent:-228600,bullet:{kind:'number',scheme:'arabicPeriod',startAt:5}}}},
 {id:'r3',type:'text',text:'\nThird',style:{fontSize:254000,slideParagraph:{level:1,marginLeft:800100,firstLineIndent:-228600,bullet:{kind:'number',scheme:'arabicPeriod'}}}}
 ]}];
 const source=writeSlidesPptx(d),after=openSlidesPptx(source,d.documentId).document,shape=after.slides[0]!.shapes[0]!;assert.equal(shape.type,'text');if(shape.type!=='text')return;
 assert.deepEqual(shape.textFrame?.autofit,{mode:'normal',fontScale:85000,lineReduction:12000});assert.equal(shape.runs.map(r=>r.text).join(''),'First\nSecond\nThird');
 const first=shape.runs.find(r=>r.text==='First')!.style!.slideParagraph!;assert.equal(first.beforePercent,20000);assert.deepEqual(first.bullet,{kind:'char',character:'•',color:{kind:'rgba',red:50,green:60,blue:70,alpha:255},sizePercent:90000,font:'Arial'});
 assert.equal(shape.runs.find(r=>r.text==='Second')!.style!.slideParagraph?.level,1);
 assert.throws(()=>writeLegacyPpt(after),fail('unsupported'));
 shape.runs.find(r=>r.text==='Third')!.text='Edited third';const reopened=openSlidesPptx(saveSlidesPptx(source,after),d.documentId).document;
 assert.equal((reopened.slides[0]!.shapes[0] as Extract<Shape,{type:'text'}>).runs.map(r=>r.text).join(''),'First\nSecond\nEdited third');
});

test('PPTX nested groups, crop, flips, default presets, styled connectors and media relationships round trip',async()=>{
 const d=createDocument('slides','graphics-interop'),asset=makeImageResource(PNG);d.resources[asset.id]=asset;
 const geo={x:914400,y:914400,width:1828800,height:914400,rotation:0},color={kind:'rgba' as const,red:20,green:40,blue:80,alpha:255};
 d.slides[0]!.shapes=[{id:'g',type:'group',geometry:{...geo,width:3657600,height:2743200,rotation:1800000},flipH:true,childSpace:{x:0,y:0,width:3657600,height:2743200},children:[
 {id:'crop-image',type:'image',geometry:geo,resourceId:asset.id,crop:{left:12500,right:5000,top:2500,bottom:3000},flipV:true},
 {id:'nested',type:'group',geometry:{...geo,x:0,y:0},childSpace:{x:0,y:0,width:1828800,height:914400},children:[
 {id:'preset-1',type:'roundRect',geometry:{...geo,x:0,y:0},fill:color},
 {id:'preset-2',type:'triangle',geometry:{...geo,x:0,y:0},stroke:{color,width:12700,dash:'dashDot'}},
 {id:'preset-3',type:'diamond',geometry:{...geo,x:0,y:0},fill:color},
 {id:'preset-4',type:'parallelogram',geometry:{...geo,x:0,y:0},fill:color},
 {id:'cxn',type:'line',connector:true,geometry:{...geo,x:0,y:0},flipV:true,stroke:{color,width:25400,dash:'dash',head:{type:'arrow',width:'sm'},tail:{type:'triangle',length:'lg'}}}
 ]}]}];
 const input=writeSlidesPptx(d),opened=openSlidesPptx(input,d.documentId).document;
 const source=Array.from(walkSlideShapes(d.slides[0]!.shapes)),actual=Array.from(walkSlideShapes(opened.slides[0]!.shapes));
 assert.equal(actual.length,source.length);for(const [i,s] of actual.entries()){const {id:ignored,...value}=s;const {id:ignored2,...expected}=source[i]!;if(s.type==='group'){assert.equal(s.type,expected.type);assert.deepEqual(s.geometry,expected.geometry);assert.deepEqual(s.childSpace,(source[i] as Extract<Shape,{type:'group'}>).childSpace);}else assert.deepEqual(value,expected);}
 assert.equal(Object.keys(opened.resources).length,1);assert.equal(opened.resources[asset.id]!.sha256,asset.sha256);assert.throws(()=>writeLegacyPpt(opened),fail('unsupported'));
 assert.deepEqual(saveSlidesPptx(input,opened),input);
 const picture=actual.find(s=>s.type==='image') as Extract<Shape,{type:'image'}>;picture.crop={left:10000,right:10000,top:0,bottom:0};
 const group=opened.slides[0]!.shapes[0]!;group.geometry.x+=914400;
 const saved=saveSlidesPptx(input,opened),reopened=openSlidesPptx(saved,d.documentId).document;
 assert.equal(reopened.slides[0]!.shapes[0]!.geometry.x,group.geometry.x);
 assert.deepEqual((Array.from(walkSlideShapes(reopened.slides[0]!.shapes)).find(s=>s.type==='image') as Extract<Shape,{type:'image'}>).crop,picture.crop);
 const opc=openOpc(saved),slide='ppt/slides/slide1.xml';assert.equal(opc.relationships.get(slide)!.filter(r=>r.type.endsWith('/image')).length,1);
 const parts=new Map(opc.parts),raw=new TextDecoder().decode(parts.get(slide)!);
 for(const changed of [raw.replace(/(<p:cNvPr id=")3"/,(_,prefix)=>prefix+'2"'),raw.replace(/(<a:srcRect l=")10000"/,(_,prefix)=>prefix+'99999"'),raw.replace('<p:cNvCxnSpPr/>','<p:cNvCxnSpPr><a:stCxn id="2" idx="0"/></p:cNvCxnSpPr>')]){
   assert.notEqual(changed,raw);const modified=new Map(parts);modified.set(slide,new TextEncoder().encode(changed));const bytes=writeZip(modified);assert.throws(()=>openSlidesPptx(bytes,d.documentId),fail());assert.equal((await inspectForEditor(bytes,'protected.pptx','protected')).report.readOnly,true);
 }
});

test('PPTX attached straight connector references survive shape movement and native reordering',async()=>{
 const d=createDocument('slides','connector-format'),g={x:914400,y:914400,width:914400,height:914400,rotation:0};
 d.slides[0]!.shapes=[{id:'a',type:'rect',geometry:g},{id:'b',type:'rect',geometry:{...g,x:4572000}},{id:'line',type:'line',connector:true,geometry:{x:1828800,y:1371600,width:2743200,height:0,rotation:0},connections:{start:{shapeId:'a',site:3},end:{shapeId:'b',site:1}},stroke:{width:25400,color:{kind:'rgba',red:0,green:0,blue:0,alpha:255},tail:{type:'triangle'}}}];
 const input=writeSlidesPptx(d),source=openSlidesPptx(input,d.documentId),shapes=source.document.slides[0]!.shapes;
 const [a,b,line]=shapes;assert.ok(a&&b&&line?.type==='line');assert.equal(line.connections?.start.shapeId,a.id);assert.equal(line.connections?.end.shapeId,b.id);
 b.geometry.y+=914400;source.document.slides[0]!.shapes=[line,b,a];
 const saved=saveSlidesPptx(input,source.document),reopened=openSlidesPptx(saved,d.documentId).document,after=reopened.slides[0]!.shapes;
 const current=after[0]!;assert.equal(current.type,'line');if(current.type!=='line')return;
 assert.deepEqual(current.connections,{start:{shapeId:after[2]!.id,site:3},end:{shapeId:after[1]!.id,site:1}});
 assert.equal(current.geometry.height,914400);
 const opc=openOpc(saved),path='ppt/slides/slide1.xml',raw=new TextDecoder().decode(opc.parts.get(path)!);
 assert.match(raw,/<a:stCxn id="4" idx="3"/);assert.match(raw,/<a:endCxn id="3" idx="1"/);
 const bad=raw.replace(/(<a:stCxn id=")4"/,(_,prefix)=>prefix+'999"');const parts=new Map(opc.parts);parts.set(path,new TextEncoder().encode(bad));
 assert.equal((await inspectForEditor(writeZip(parts),'bad.pptx','bad')).report.readOnly,true);
});

test('M6.7 native DrawingML table round trip preserves text, grid and flags; legacy conversion refuses',async()=>{
 const {DocumentSession}=await import('@office-ex/core');const {insertSlideTable,editSlideTable}=await import('@office-ex/slides');const s=new DocumentSession(createDocument('slides','m6-table'));
 s.execute(insertSlideTable(s.snapshot(),'slide-1',2,2,'grid'));s.execute(editSlideTable(s.snapshot(),'slide-1','grid',[['名称','Value'],['多行\n文本','42']]));const d=s.snapshot().document as SlidesDocument;const b=writeSlidesPptx(d);const opened=openSlidesPptx(b,d.documentId).document;const shape=opened.slides[0]!.shapes[0]!;assert.equal(shape.type,'table');if(shape.type!=='table')return;assert.deepEqual(shape.table.cells,[['名称','Value'],['多行\n文本','42']]);assert.equal(shape.table.firstRow,true);assert.equal(shape.table.bandRow,true);assert.deepEqual(saveSlidesPptx(b,opened),b);shape.table.cells[1]![1]='43';const out=openSlidesPptx(saveSlidesPptx(b,opened),d.documentId).document;const read=out.slides[0]!.shapes[0]!;assert.equal(read.type,'table');if(read.type==='table')assert.equal(read.table.cells[1]![1],'43');assert.throws(()=>writeLegacyPpt(d),fail('unsupported'));
});
