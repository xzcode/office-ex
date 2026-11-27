// SPDX-License-Identifier: MIT
/** Internal rich clipboard, not an executable HTML import path.
 * Resolved source formatting is copied as direct styles. Target named style
 * IDs are not borrowed across documents. Image bytes remain content-addressed.
 */
import {assert,assertId} from '@office-ex/contracts';
import {createDocument,validateDocument,validateImageResource,resolveTextStyle,assertPlainData,patchCommand,cloneData,inlineText,type DocumentSnapshot,type DocumentCommand,type WriterDocument,type DeepReadonly,type Paragraph,type Inline,type ResourceRef,type WriterNumbering,type TextStyle,type PatchOperation} from '@office-ex/core';
import {resolveWriterRange,type WriterRange,type WriterPoint} from './ranges.js';
import {sliceWriterRuns,writerIdAllocator,templateAt,nonEmpty} from './paragraphs.js';
const MAX_CLIPBOARD=8*1024*1024;
const CHAR=['fontFamily','fontAscii','fontHighAnsi','fontEastAsia','fontComplex','fontSize','complexFontSize','bold','italic','underline','color','language','eastAsiaLanguage','bidiLanguage','runDirection','characterSpacing','kerningThreshold'] as const;
const PARA=['alignment','indent','spaceBefore','spaceAfter','lineHeight','lineMultiple','minLineHeight','keepNext','keepLines','pageBreakBefore','outlineLevel','widowControl','paragraphDirection','rightIndent','firstLineIndent','contextualSpacing','paragraphBottomBorder','writerList'] as const;
export const WRITER_CLIPBOARD_TYPE='application/x-office-ex-writer-fragment+json';
interface Fragment {magic:'office-ex-writer-fragment';version:1;paragraphs:Paragraph[];resources:Record<string,ResourceRef>;numbering?:WriterNumbering}
function pick(style:DeepReadonly<TextStyle>,keys:readonly (keyof TextStyle)[]):TextStyle{
  return Object.fromEntries(keys.filter(k=>style[k]!==undefined).map(k=>[k,cloneData(style[k])])) as TextStyle;
}
function parse(value:unknown):Fragment{
  assertPlainData(value);
  assert(value&&typeof value==='object'&&!Array.isArray(value),'invalid','Invalid Writer clipboard');
  const f=value as Fragment;
  assert(Object.keys(f).every(k=>['magic','version','paragraphs','resources','numbering'].includes(k))&&f.magic==='office-ex-writer-fragment'&&f.version===1&&Array.isArray(f.paragraphs)&&f.paragraphs.length>0&&f.paragraphs.length<=1000,'invalid','Invalid Writer clipboard envelope');
  assert(f.resources&&typeof f.resources==='object'&&!Array.isArray(f.resources),'invalid','Invalid clipboard resource map');
  const temp=createDocument('writer','clipboard-validation');temp.defaultTextStyle={};temp.sections[0]!.blocks=f.paragraphs;temp.resources=f.resources;if(f.numbering)temp.numbering=f.numbering;
  validateDocument(temp);
  const used=new Set<string>();
  for(const p of f.paragraphs){
    assert(p.type==='paragraph'&&p.styleId===undefined&&Object.keys(p.style??{}).every(k=>(PARA as readonly string[]).includes(k)),'invalid','Clipboard paragraph styles must be direct');
    for(const run of p.runs){
      if(run.type==='text')assert(run.styleId===undefined&&Object.keys(run.style??{}).every(k=>(CHAR as readonly string[]).includes(k)),'invalid','Clipboard run styles must be direct');
      else if(run.type==='image')used.add(run.resourceId);
    }
  }
  for(const [id,r] of Object.entries(f.resources)){validateImageResource(r);assert(used.has(id),'invalid','Unreferenced clipboard resource');}
  return f;
}
export function copyWriterFragment(snapshot:DocumentSnapshot,range:WriterRange):{plain:string;rich:string}{
  assert(snapshot.document.kind==='writer','invalid','Expected Writer document');const model=snapshot.document;
  const resolved=resolveWriterRange(snapshot,range),allocate=writerIdAllocator('clip'),paragraphs:Paragraph[]=[],resources:Record<string,ResourceRef>={};
  for(const segment of resolved.segments){
    const p=segment.paragraph,named=p.styleId??model.defaultParagraphStyleId,base=resolveTextStyle(model.defaultTextStyle,model.styles,named?[named]:[],p.style??{}).style;
    const runs=sliceWriterRuns(p.runs,segment.start,segment.end,new Set(),allocate).map(r=>{
      if(r.type==='text'){
        const effective=resolveTextStyle(base,model.styles,r.styleId?[r.styleId]:[],r.style??{}).style;
        return {id:allocate(),type:'text' as const,text:r.text,style:pick(effective,CHAR)};
      }
      if(r.type==='image'){const resource=model.resources[r.resourceId];validateImageResource(resource);resources[r.resourceId]=cloneData(resource);}
      return {...r,id:allocate()};
    });
    paragraphs.push({id:allocate(),type:'paragraph',style:pick(base,PARA),runs:runs.length?runs:[{id:allocate(),type:'text',text:''}]});
  }
  const lists:WriterNumbering['lists']={};
  for(const p of paragraphs){const id=p.style?.writerList?.numId;if(id){assert(model.numbering?.lists[String(id)],'invalid','Missing copied list definition');lists[String(id)]=cloneData(model.numbering.lists[String(id)]!) as WriterNumbering['lists'][string];}}
  const fragment:Fragment={magic:'office-ex-writer-fragment',version:1,paragraphs,resources,...(Object.keys(lists).length?{numbering:{nextId:Math.max(...Object.keys(lists).map(Number))+1,lists}}:{})};
  const rich=JSON.stringify(fragment);
  assert(new TextEncoder().encode(rich).length<=MAX_CLIPBOARD,'limit','Rich clipboard exceeds size budget');
  parse(fragment);
  // Explicitly include object replacement markers in the plain alternative,
  // never silently discard a copied picture. External plain text consumers
  // may see the marker; the Office EX plain paste path rejects it.
  return {plain:paragraphs.map(p=>p.runs.map(inlineText).join('')).join('\n'),rich};
}
export function pasteWriterFragment(snapshot:DocumentSnapshot,range:WriterRange,rich:string,operationId:string):{command:DocumentCommand;caret:WriterPoint}{
  assert(snapshot.document.kind==='writer','invalid','Expected Writer document');
  assert(typeof rich==='string'&&rich.length<=MAX_CLIPBOARD&&new TextEncoder().encode(rich).length<=MAX_CLIPBOARD,'limit','Rich clipboard exceeds size budget');
  let value:unknown;try{value=JSON.parse(rich);}catch{assert(false,'invalid','Invalid clipboard JSON');}
  const f=parse(value),r=resolveWriterRange(snapshot,range),model=snapshot.document,allocate=writerIdAllocator(operationId),used=new Set<string>();
  const left=sliceWriterRuns(r.first.runs,0,r.start,used,allocate),right=sliceWriterRuns(r.last.runs,r.end,r.last.text.length,used,allocate);
  const template=templateAt(r.first.runs,r.start),items:Paragraph[]=[],operations:PatchOperation[]=[];
  for(const [id,resource] of Object.entries(f.resources)){
    if(model.resources[id])validateImageResource(model.resources[id]);
    else operations.push({op:'set',path:['resources',id],value:resource});
  }
  const mapping=new Map<number,number>();
  if(f.numbering){
    const registry=cloneData(model.numbering??{nextId:1,lists:{}}) as WriterNumbering;
    for(const [id,def]of Object.entries(f.numbering.lists)){
      assert(registry.nextId<=1000000000&&Object.keys(registry.lists).length<2048,'limit','Clipboard numbering budget exceeded');
      const next=registry.nextId++;mapping.set(Number(id),next);registry.lists[String(next)]=cloneData(def);
    }
    operations.push({op:'set',path:['numbering'],value:registry});
  }
  let caret:WriterPoint|undefined;
  for(let i=0;i<f.paragraphs.length;i++){
    const src=cloneData(f.paragraphs[i]!),tail=i===f.paragraphs.length-1;
    const ref=src.style?.writerList;
    if(ref?.numId){const id=mapping.get(ref.numId);assert(id,'invalid','Missing clipboard numbering mapping');src.style={...src.style,writerList:{...ref,numId:id}};}
    if(i===0&&ref?.numId)assert(r.start===0&&r.end===r.last.text.length,'unsupported','Paste a numbered fragment into a whole paragraph or an empty paragraph');
    const pasted=src.runs.map(run=>{
      if(run.type==='image')assert(run.width<=r.first.contentWidth&&run.height<=r.first.contentHeight,'unsupported','Pasted image does not fit target');
      return {...cloneData(run),id:allocate()};
    });
    const id=i===0?r.first.paragraph.id:tail&&r.first.block!==r.last.block?r.last.paragraph.id:allocate();
    const base=i===0?{...cloneData(r.first.paragraph),...(ref?{style:{...r.first.paragraph.style,...src.style}}:{})}:{id,type:'paragraph' as const,...(model.defaultParagraphStyleId?{styleId:model.defaultParagraphStyleId}:{}),style:cloneData(src.style??{}),runs:[]};
    const before=i===0?left:[],after=tail?right:[];
    items.push({...base,id,runs:nonEmpty([...before,...pasted,...after],template,allocate)});
    if(tail)caret={paragraphId:id,offset:before.reduce((n,x)=>n+inlineText(x).length,0)+pasted.reduce((n,x)=>n+inlineText(x).length,0)};
  }
  operations.push({op:'splice',path:r.first.parent,index:r.first.block,deleteCount:r.last.block-r.first.block+1,items});
  assert(caret,'internal','Missing clipboard caret');return {command:patchCommand(snapshot,'WRT-03',operations),caret};
}

