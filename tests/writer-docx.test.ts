// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import { OfficeError } from '@office-ex/contracts';
import { createDocument, DocumentSession, validateDocument, type WriterDocument } from '@office-ex/core';
import { openWriterDocx, writeWriterDocx, saveWriterDocx, openOpc, writeZip } from '@office-ex/ooxml';
import { replaceParagraphRange, paragraphText, formatParagraphRange, copyWriterFragment, pasteWriterFragment } from '@office-ex/writer';
const fails=(code:string)=>(e:unknown)=>e instanceof OfficeError&&e.code===code;
export function fixture():WriterDocument{
  const doc=createDocument('writer','fixture','Multilingual');
  doc.defaultTextStyle={fontFamily:'Arial',fontSize:152400};
  doc.sections[0]!.blocks=[
    {id:'p1',type:'paragraph',style:{alignment:'center',spaceAfter:25400},runs:[{id:'r1',type:'text',text:' 中文 & <Office> 😀 ',style:{bold:true,color:{kind:'rgba',red:32,green:80,blue:160,alpha:255}}},{id:'r2',type:'text',text:'日本語\n한국어\tالعربية',style:{italic:true,underline:true,fontSize:177800}}]},
    {id:'break',type:'page-break'},
    {id:'p2',type:'paragraph',runs:[{id:'r3',type:'text',text:'e\u0301 हिन्दी русский'}]}
  ];return doc;
}
test('WRT-04 semantic DOCX writer preserves multilingual runs, direct styles, paragraph alignment, breaks and page dimensions',()=>{
  const source=fixture(),bytes=writeWriterDocx(source),read=openWriterDocx(bytes,'fixture','Multilingual');
  const b=read.document.sections[0]!.blocks;
  assert.equal(b[0]!.type,'paragraph');const p=b[0]!;if(p.type!=='paragraph')throw Error();
  assert.equal(p.style?.alignment,'center');assert.equal(p.style?.spaceAfter,25400);
  assert.equal(p.runs[0]!.type,'text');if(p.runs[0]!.type!=='text'||p.runs[1]!.type!=='text')throw Error();
  assert.equal(p.runs[0].text,' 中文 & <Office> 😀 ');assert.equal(p.runs[0].style?.bold,true);
  assert.deepEqual(p.runs[0].style?.color,{kind:'rgba',red:32,green:80,blue:160,alpha:255});
  assert.equal(p.runs[1].text,'日本語\n한국어\tالعربية');assert.equal(p.runs[1].style?.fontSize,177800);
  assert.equal(p.runs[1].style?.italic,true);assert.equal(p.runs[1].style?.underline,true);
  assert.equal(b[1]!.type,'page-break');assert.equal(b[2]!.type,'paragraph');
  assert.ok(Math.abs(read.document.sections[0]!.page.width-source.sections[0]!.page.width)<=317.5);
});
test('WRT-04 original copy is byte exact; edited semantic save changes only main part and reopens',()=>{
  const bytes=writeWriterDocx(fixture()),source=openWriterDocx(bytes,'fixture','Multilingual');
  assert.deepEqual(saveWriterDocx(source,source.document),bytes);
  const session=new DocumentSession(source.document,{saved:true});const paragraph=source.document.sections[0]!.blocks[0]!.id;
  session.execute(replaceParagraphRange(session.snapshot(),paragraph,1,3,'繁體','edit'));
  const edited=saveWriterDocx(source,session.snapshot().document);
  const before=openOpc(bytes),after=openOpc(edited);
  for(const [path,data] of before.parts)if(path!==before.mainPart)assert.deepEqual(after.parts.get(path),data);
  const reopen=openWriterDocx(edited,'fixture','Multilingual'),reopened=new DocumentSession(reopen.document);
  assert.match(paragraphText(reopened.snapshot(),reopen.document.sections[0]!.blocks[0]!.id),/^ 繁體/);
  assert.deepEqual(saveWriterDocx(source,source.document),bytes);
  session.undo();assert.deepEqual(saveWriterDocx(source,session.snapshot().document),bytes);
});
test('WRT-04 recognizes nondefault part path and namespace prefix without hardcoding word/document.xml',()=>{
  const bytes=writeWriterDocx(fixture()),opc=openOpc(bytes),parts=new Map(opc.parts),enc=new TextEncoder(),dec=new TextDecoder();
  const old=opc.mainPart!,next='custom/story.xml';const xml=dec.decode(parts.get(old)!).replaceAll('xmlns:w=','xmlns:q=').replaceAll('<w:','<q:').replaceAll('</w:','</q:').replaceAll(' w:',' q:');
  parts.delete(old);parts.set(next,enc.encode(xml));
  for(const p of ['[Content_Types].xml','_rels/.rels'])parts.set(p,enc.encode(dec.decode(parts.get(p)!).replaceAll(old,next)));
  const doc=writeZip(parts),opened=openWriterDocx(doc,'fixture','Multilingual');assert.equal(opened.mainPart,next);
  assert.deepEqual(saveWriterDocx(opened,opened.document),doc);
});
test('WRT-04 refuses unsupported parts, elements and attributes; original bytes unchanged',()=>{
  const bytes=writeWriterDocx(fixture()),opc=openOpc(bytes),dec=new TextDecoder(),enc=new TextEncoder();
  const before=bytes.slice();
  for(const mutate of [
    (s:string)=>s.replace('<w:p>','<w:p w:unrecognizedFeature="00112233">'),
    (s:string)=>s.replace('<w:p>','<w:p><!--preserve this-->'),
    (s:string)=>s.replace('<w:t xml:space="preserve">','<w:fldSimple w:instr="DATE"/><w:t xml:space="preserve">'),
    (s:string)=>s.replace('<w:sz w:val="24"/>','<w:lang w:val="en-US" w:unknown="1"/><w:sz w:val="24"/>')
  ]){
    const parts=new Map(opc.parts);parts.set(opc.mainPart!,enc.encode(mutate(dec.decode(parts.get(opc.mainPart!)!))));
    assert.throws(()=>openWriterDocx(writeZip(parts),'fixture','Title'),fails('unsupported'));
  }
  const parts=new Map(opc.parts);parts.set('unknown.xml',enc.encode('<unknown/>'));
  parts.set('[Content_Types].xml',enc.encode(dec.decode(parts.get('[Content_Types].xml')!).replace('</Types>','<Default Extension="xml" ContentType="application/xml"/></Types>')));
  assert.throws(()=>openWriterDocx(writeZip(parts),'fixture','Title'),fails('unsupported'));
  assert.deepEqual(bytes,before);
});
test('WRT-04 export rejects named styles, tables, unknown identity and unrepresentable controls',()=>{
  const doc=fixture(),source=openWriterDocx(writeWriterDocx(doc),'fixture','Title');
  const foreign=fixture();foreign.documentId='other';assert.throws(()=>saveWriterDocx(source,foreign),fails('conflict'));
  const styled=fixture();styled.styles.Normal={id:'Normal',name:'Normal',properties:{bold:true}};assert.throws(()=>writeWriterDocx(styled),fails('unsupported'));
  const invalid=fixture();invalid.sections[0]!.blocks=[{id:'p',type:'paragraph',runs:[{id:'r',type:'text',text:'a\rb'}]}];
  assert.throws(()=>writeWriterDocx(invalid),fails('unsupported'));
});

test('WRT-04 control-only and empty runs round-trip without invented text',()=>{
  const doc=createDocument('writer','control-only');
  doc.sections[0]!.blocks=[{id:'p',type:'paragraph',runs:[
    {id:'tab',type:'tab'},{id:'line',type:'line-break'},{id:'empty',type:'text',text:''}
  ]}];
  const bytes=writeWriterDocx(doc),source=openWriterDocx(bytes,doc.documentId,doc.title);
  const block=source.document.sections[0]!.blocks[0]!;
  if(block.type!=='paragraph')throw Error();
  assert.deepEqual(block.runs.map(run=>run.type==='text'?run.text:run.type),['\t','\n','']);
  assert.deepEqual(saveWriterDocx(source,source.document),bytes);
});

test('M7.7 direct Writer strike and solid shading survive editing, clipboard, undo and native DOCX save',()=>{
  const d=createDocument('writer','decorations'),shade={kind:'rgba' as const,red:239,green:240,blue:241,alpha:255};
  d.sections[0]!.blocks=[{id:'p1',type:'paragraph',runs:[{id:'r1',type:'text',text:'Alpha',style:{writerStrike:true,writerShading:shade,underline:true}},{id:'r2',type:'text',text:' Beta',style:{writerStrike:false,writerShading:null}}]}];
  const bytes=writeWriterDocx(d),source=openWriterDocx(bytes,d.documentId,d.title);
  assert.deepEqual(saveWriterDocx(source,source.document),bytes);
  const p=source.document.sections[0]!.blocks[0]!;if(p.type!=='paragraph')throw Error();
  const first=p.runs[0]!,second=p.runs[1]!;if(first.type!=='text'||second.type!=='text')throw Error();
  assert.deepEqual(first.style?.writerShading,shade);assert.equal(first.style?.writerStrike,true);assert.equal(first.style?.underline,true);assert.equal(second.style?.writerStrike,false);assert.equal(second.style?.writerShading,null);
  const s=new DocumentSession(source.document,{saved:true});
  const range={anchor:{paragraphId:p.id,offset:0},focus:{paragraphId:p.id,offset:5}};
  const fragment=copyWriterFragment(s.snapshot(),range);
  s.execute(pasteWriterFragment(s.snapshot(),{anchor:{paragraphId:p.id,offset:10},focus:{paragraphId:p.id,offset:10}},fragment.rich,'decorate-paste').command);
  s.execute(formatParagraphRange(s.snapshot(),p.id,0,2,{writerStrike:false,writerShading:null},'decorate-clear'));
  const edited=saveWriterDocx(source,s.snapshot().document),reopened=openWriterDocx(edited,d.documentId,d.title);
  const out=reopened.document.sections[0]!.blocks[0]!;if(out.type!=='paragraph')throw Error();
  assert.equal(out.runs.map(r=>r.type==='text'?r.text:'').join(''),'Alpha BetaAlpha');
  const texts=out.runs.filter(r=>r.type==='text');
  assert.equal(texts[0]!.style?.writerStrike,false);assert.equal(texts[0]!.style?.writerShading,null);
  assert.equal(texts.at(-1)!.style?.writerStrike,true);assert.deepEqual(texts.at(-1)!.style?.writerShading,shade);
  const before=openOpc(bytes),after=openOpc(edited);
  for(const [path,data] of before.parts)if(path!==before.mainPart)assert.deepEqual(after.parts.get(path),data);
  s.undo();s.undo();assert.deepEqual(saveWriterDocx(source,s.snapshot().document),bytes);
});
test('M7.7 producer omitted shading pattern is admitted; patterned/theme shading and double strike stay guarded',()=>{
  const d=createDocument('writer','decorate-import');d.sections[0]!.blocks=[{id:'p',type:'paragraph',runs:[{id:'r',type:'text',text:'Marked',style:{bold:true}}]}];
  const pkg=openOpc(writeWriterDocx(d)),enc=new TextEncoder(),dec=new TextDecoder(),xml=dec.decode(pkg.parts.get(pkg.mainPart!)!);
  const make=(extra:string)=>{const parts=new Map(pkg.parts);parts.set(pkg.mainPart!,enc.encode(xml.replace('</w:rPr>',extra+'</w:rPr>')));return writeZip(parts);};
  const source=openWriterDocx(make('<w:strike w:val="true"/><w:shd w:fill="EFF0F1"/>'),d.documentId,d.title);
  const p=source.document.sections[0]!.blocks[0]!;if(p.type!=='paragraph'||p.runs[0]?.type!=='text')throw Error();
  assert.equal(p.runs[0].style?.writerStrike,true);assert.equal(p.runs[0].style?.writerShading?.kind,'rgba');
  for(const markup of ['<w:shd w:val="pct20" w:fill="EFF0F1"/>','<w:shd w:val="clear" w:themeFill="accent1" w:fill="EFF0F1"/>','<w:dstrike/>'])assert.throws(()=>openWriterDocx(make(markup),d.documentId,d.title),fails('unsupported'));
  const slide=createDocument('slides','no-writer-decorations');slide.defaultTextStyle.writerStrike=true;assert.throws(()=>validateDocument(slide),fails('unsupported'));
});
