// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import {createDocument,resolveSlideThemeColor,type SlidesDocument,type Shape} from '@office-ex/core';
import {OfficeError} from '@office-ex/contracts';
import {openOpc,writeZip,openSlidesPptx,writeSlidesPptx,saveSlidesPptx} from '@office-ex/ooxml';
import {openLegacyPpt,writeLegacyPpt,saveLegacyPpt,prepareSlidesSave,serializeDraft,inspectForEditor} from '@office-ex/format-service';
import {readCfb,writeCfb,type CfbEntry} from '@office-ex/cfb';
import {slidesFormatFixture} from './slides-format-fixture.js';

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
