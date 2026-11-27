// SPDX-License-Identifier: MIT
/** Limited source-backed editing for modern files outside the semantic
 * editor profile. Only admitted existing text/scalar slots can change. */
import type {Document,Element} from '@xmldom/xmldom';
import {OfficeError,assertUnicode,isRecord} from '@office-ex/contracts';
import {sha256} from '@office-ex/core';
import {openOpc,writeOpc,type OpcPackage} from './opc.js';
import {parseXml,serializeXml,childElements} from './xml.js';
import {bytesEqual,checkCancelled,type ContainerOptions} from './limits.js';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',XML='http://www.w3.org/XML/1998/namespace';
const LIMIT=20000,TEXT=2*1024*1024,SOURCE=32*1024*1024;
export type CompatibleValueKind='text'|'number'|'boolean';
export interface CompatibleValue {id:number;label:string;text:string;kind:CompatibleValueKind}
export interface CompatibleValuePage {name:string;fields:CompatibleValue[];locked:number}
export interface CompatibleOfficeView {format:'xlsx'|'pptx';sourceSha256:string;pages:CompatibleValuePage[];editableSegments:number;lockedBlocks:number}
interface Slot {part:string;node:Element;text:string;kind:CompatibleValueKind;shared:boolean}
interface Scan {pkg:OpcPackage;documents:Map<string,Document>;slots:Slot[];view:CompatibleOfficeView}
function need(v:unknown,m:string,code:'unsupported'|'invalid'|'limit'|'unsafe'|'conflict'='unsupported'):asserts v{if(!v)throw new OfficeError(code,'Compatible Office values: '+m);}
const is=(e:Element,ns:string,name:string)=>e.namespaceURI===ns&&e.localName===name;
const children=(e:Element,ns:string,name:string)=>childElements(e).filter(c=>is(c,ns,name));
function one(e:Element,ns:string,name:string,required=false){const list=children(e,ns,name);need(list.length<=1&&(!required||list.length===1),'Invalid '+name+' cardinality','invalid');return list[0];}
function scalar(e:Element){need(!childElements(e).length,'Nested scalar text','invalid');return e.textContent??'';}
function clean(t:string,kind:CompatibleValueKind):void{
 assertUnicode(t);need(t.length<=32767&&!/[\x00-\x1f\x7f\ufffe\uffff]/u.test(t),'Invalid replacement text','invalid');
 if(kind==='text')need(!/_x[0-9a-f]{4}_/i.test(t),'OOXML escaped strings require a dedicated adapter');
 if(kind==='number')need(/^-?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(t)&&Number.isFinite(Number(t)),'Expected finite numeric constant','invalid');
 if(kind==='boolean')need(t==='0'||t==='1','Expected boolean 0 or 1','invalid');
}
function scan(bytes:Uint8Array,format:'xlsx'|'pptx',options:ContainerOptions={}):Scan{
 need(bytes.length>0&&bytes.length<=SOURCE,'Source size budget','limit');
 const pkg=openOpc(bytes,options);need(pkg.kind===format&&pkg.mainPart&&pkg.conformance==='transitional','Wrong native format');
 need(!pkg.hasMacros&&!pkg.hasSignatures&&!pkg.externalRelationships,'Active or external document graph','unsafe');
 const docs=new Map<string,Document>(),slots:Slot[]=[],pages:CompatibleValuePage[]=[];let length=0,locked=0;
 const view:CompatibleOfficeView={format,sourceSha256:sha256(bytes),pages,editableSegments:0,lockedBlocks:0};
 const open=(part:string,ns:string,name:string):Element=>{
  let d=docs.get(part);if(!d){const data=pkg.parts.get(part);need(data,'Missing XML part','invalid');d=parseXml(data,options);docs.set(part,d);}
  const root=d.documentElement;need(root&&is(root,ns,name),'Wrong XML root','invalid');return root;
 };
 function slot(page:CompatibleValuePage,part:string,node:Element,label:string,kind:CompatibleValueKind,shared=false){
  const text=scalar(node);
  try{clean(text,kind);}catch{page.locked++;locked++;return;}
  need(slots.length<LIMIT&&(length+=text.length)<=TEXT,'Editable value budget','limit');
  const id=slots.length;slots.push({part,node,text,kind,shared});page.fields.push({id,label,text,kind});
 }
 function page(name:string){const p:CompatibleValuePage={name,fields:[],locked:0};pages.push(p);return p;}
 function lock(p:CompatibleValuePage){p.locked++;locked++;}
 const relationship=(owner:string,id:string,type:string)=>{
  const rs=(pkg.relationships.get(owner)??[]).filter(r=>r.id===id&&r.type===R+'/'+type);
  need(rs.length===1&&rs[0]!.mode==='Internal'&&rs[0]!.targetPart&&!rs[0]!.fragment,'Invalid '+type+' relationship','invalid');return rs[0]!.targetPart!;
 };
 const passive=new Set(format==='xlsx'?['officeDocument','extended-properties','worksheet','styles','sharedStrings','theme','printerSettings','image']:['officeDocument','extended-properties','slide','slideMaster','slideLayout','theme','themeOverride','printerSettings','presProps','viewProps','tableStyles','notesSlide','notesMaster','handoutMaster','image','chart','chartStyle','chartColorStyle','diagramData','diagramLayout','diagramColors','diagramQuickStyle','diagramDrawing','tags']);
 for(const rels of pkg.relationships.values())for(const r of rels){
  checkCancelled(options);
  const k=r.type.startsWith(R+'/')?r.type.slice(R.length+1):'';
  need(passive.has(k)||['http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties','http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail'].includes(r.type),'Unrecognized or dependent relationship: '+r.type,'unsupported');
 }
 if(format==='xlsx'){
  // A changing value can affect arbitrary formulas, charts and caches,
  // including hidden sheets. Until dependency-aware invalidation is
  // implemented, admit only workbooks with no calculation dependencies.
  const workbook=open(pkg.mainPart,S,'workbook');
  const names=one(workbook,S,'definedNames'),protection=one(workbook,S,'workbookProtection');
  // Common producers emit empty placeholders, which carry no semantics.
  need((!names||(!childElements(names).length&&!names.textContent?.trim()&&names.attributes.length===0))&&(!protection||(!childElements(protection).length&&!protection.textContent?.trim()&&protection.attributes.length===0))&&!['externalReferences','pivotCaches','extLst'].some(n=>one(workbook,S,n)),'Workbook contains dependencies or protection');
  const defs=children(one(workbook,S,'sheets',true)!,S,'sheet');need(defs.length>0&&defs.length<=128,'Worksheet budget','limit');
  const sharedRel=(pkg.relationships.get(pkg.mainPart)??[]).filter(r=>r.type===R+'/sharedStrings');need(sharedRel.length<=1,'Ambiguous shared strings','invalid');
  const sharedPart=sharedRel[0]?.targetPart,shared=sharedPart?children(open(sharedPart,S,'sst'),S,'si'):[];
  need(shared.length<=LIMIT,'Shared strings budget','limit');
  const counts=new Map<number,number>(),seen=new Set<string>();
  const worksheets=defs.map(def=>{
   const part=relationship(pkg.mainPart!,def.getAttributeNS(R,'id')??'','worksheet');need(!seen.has(part),'Duplicate worksheet part','invalid');seen.add(part);
   const root=open(part,S,'worksheet'),data=one(root,S,'sheetData',true)!;
   need(!['sheetProtection','protectedRanges','dataValidations','autoFilter','sortState','tableParts','drawing','legacyDrawing','pivotTableParts','extLst'].some(n=>one(root,S,n)),'Worksheet contains dependent or protected content');
   // Conditional formatting is preserved. Only literal cellIs thresholds
   // are allowed; expression rules can depend on arbitrary other cells.
   for(const c of children(root,S,'conditionalFormatting'))for(const rule of children(c,S,'cfRule')){
    need(rule.getAttribute('type')==='cellIs'&&children(rule,S,'formula').every(f=>/^-?\d+(?:\.\d+)?$/.test(scalar(f))),'Dependent conditional formatting');
   }
   const cells:{node:Element;address:string;hidden:boolean}[]=[],addresses=new Set<string>();
   const hiddenColumns=new Uint8Array(16385),cols=one(root,S,'cols');
   if(cols)for(const col of children(cols,S,'col'))if(['1','true'].includes(col.getAttribute('hidden')??'')){
    const min=Number(col.getAttribute('min')),max=Number(col.getAttribute('max'));need(Number.isInteger(min)&&Number.isInteger(max)&&min>=1&&max>=min&&max<=16384,'Invalid hidden column range','invalid');hiddenColumns.fill(1,min,max+1);
   }
   for(const row of children(data,S,'row'))for(const c of children(row,S,'c')){
    checkCancelled(options);const address=c.getAttribute('r')??'',match=/^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(address);
    need(match&&Number(match[2])<=1048576&&Number(match[2])===Number(row.getAttribute('r'))&&[...match[1]!].reduce((n,ch)=>n*26+ch.charCodeAt(0)-64,0)<=16384&&!addresses.has(address),'Invalid or duplicate cell reference','invalid');addresses.add(address);
    need(!one(c,S,'f'),'Formula recalculation is outside the fallback profile');
    if(c.getAttribute('t')==='s'){
     const v=one(c,S,'v',true)!,i=Number(scalar(v));need(Number.isSafeInteger(i)&&i>=0&&i<shared.length,'Invalid shared-string reference','invalid');counts.set(i,(counts.get(i)??0)+1);
    }
    cells.push({node:c,address,hidden:['1','true'].includes(row.getAttribute('hidden')??'')||!!hiddenColumns[[...match[1]!].reduce((n,ch)=>n*26+ch.charCodeAt(0)-64,0)]});
   }
   return {part,root,cells,visible:(def.getAttribute('state')||'visible')==='visible',name:def.getAttribute('name')??'Worksheet'};
  });
  need((pkg.relationships.get(pkg.mainPart)??[]).filter(r=>r.type===R+'/worksheet').length===seen.size,'Unreferenced worksheet relationship');
  for(const ws of worksheets){
   if(!ws.visible)continue;const p=page(ws.name);
   // Conservative for merged cells: retain all merge definitions but do
   // not edit any cell in a merged sheet until range ownership is mapped.
   const hasMerges=!!one(ws.root,S,'mergeCells');
   for(const {node:c,address,hidden}of ws.cells){
    if(hidden||hasMerges||c.hasAttribute('cm')||c.hasAttribute('vm')||childElements(c).some(e=>!['v','is'].some(n=>is(e,S,n)))){lock(p);continue;}
    const kind=c.getAttribute('t')||'n',v=one(c,S,'v'),inline=one(c,S,'is');
    if(kind==='n'&&v&&!inline){slot(p,ws.part,v,address,'number');continue;}
    if(kind==='b'&&v&&!inline){slot(p,ws.part,v,address,'boolean');continue;}
    if(kind==='inlineStr'&&inline&&!v&&childElements(inline).length===1){const t=one(inline,S,'t');if(t){slot(p,ws.part,t,address,'text');continue;}}
    if(kind==='s'&&v&&!inline&&sharedPart){const i=Number(scalar(v)),si=shared[i]!;if(counts.get(i)===1&&childElements(si).length===1){const t=one(si,S,'t');if(t){slot(p,sharedPart,t,address,'text',true);continue;}}}
    lock(p);
   }
  }
 }else{
  const main=open(pkg.mainPart,P,'presentation'),list=one(main,P,'sldIdLst',true)!;
  need(!one(main,P,'modifyVerifier'),'Presentation is protected');
  const ids=children(list,P,'sldId');need(ids.length>0&&ids.length<=500,'Slide budget','limit');const seen=new Set<string>();
  // Only ordinary text runs in ordinary top-level shapes and native table
  // cells. Charts, diagrams, images and complex group/layout content
  // remain opaque and never enter the write model.
  const textBody=(body:Element,part:string,p:CompatibleValuePage,label:string)=>{
   for(const [pi,para]of children(body,A,'p').entries()){
    if(childElements(para).some(e=>!['pPr','r','endParaRPr','br'].some(n=>is(e,A,n)))){lock(p);continue;}
    for(const [ri,run]of children(para,A,'r').entries()){
     const t=one(run,A,'t');if(!t||childElements(run).some(e=>!is(e,A,'rPr')&&!is(e,A,'t'))){lock(p);continue;}
     slot(p,part,t,label+' · '+(pi+1)+'.'+(ri+1),'text');
    }
   }
  };
  for(const [i,id]of ids.entries()){
   const part=relationship(pkg.mainPart,id.getAttributeNS(R,'id')??'','slide');need(!seen.has(part),'Duplicate slide','invalid');seen.add(part);
   const slide=open(part,P,'sld');if(slide.getAttribute('show')==='0'){locked++;continue;}
   const tree=one(one(slide,P,'cSld',true)!,P,'spTree',true)!,p=page('Slide '+(i+1));
   for(const [j,shape]of childElements(tree).entries()){
    if(is(shape,P,'sp')){
     const body=one(shape,P,'txBody');if(body)textBody(body,part,p,'Shape '+(j+1));else lock(p);
    }else if(is(shape,P,'graphicFrame')){
     const graphic=one(shape,A,'graphic'),data=graphic&&one(graphic,A,'graphicData');
     const tbl=data&&data.getAttribute('uri')==='http://schemas.openxmlformats.org/drawingml/2006/table'?one(data,A,'tbl'):undefined;
     if(!tbl){lock(p);continue;}
     for(const [r,row]of children(tbl,A,'tr').entries())for(const [c,cell]of children(row,A,'tc').entries()){
      if(['1','true'].includes(cell.getAttribute('hMerge')??'')||['1','true'].includes(cell.getAttribute('vMerge')??'')){lock(p);continue;}
      const body=one(cell,A,'txBody');if(body)textBody(body,part,p,'Table '+(j+1)+' R'+(r+1)+'C'+(c+1));else lock(p);
     }
    }else if(!is(shape,P,'nvGrpSpPr')&&!is(shape,P,'grpSpPr'))lock(p);
   }
  }
 }
 need(slots.length>0,'No supported editable values');
 view.editableSegments=slots.length;view.lockedBlocks=locked;return {pkg,documents:docs,slots,view};
}
export function openCompatibleOfficeValues(bytes:Uint8Array,format:'xlsx'|'pptx',options:ContainerOptions={}):CompatibleOfficeView{return scan(bytes,format,options).view;}
export function saveCompatibleOfficeValues(bytes:Uint8Array,format:'xlsx'|'pptx',documentId:string,value:unknown,options:ContainerOptions={}):Uint8Array{
 need(isRecord(value)&&Object.keys(value).length===3&&value.documentId===documentId&&typeof value.sourceSha256==='string'&&Array.isArray(value.edits)&&value.edits.length<=LIMIT,'Invalid edit envelope','invalid');
 const before=scan(bytes,format,options);need(value.sourceSha256===before.view.sourceSha256,'Source digest mismatch','conflict');
 const changed=new Map<number,string>(),seen=new Set<number>();let length=0;
 for(const item of value.edits){
  need(isRecord(item)&&Object.keys(item).length===2&&Number.isSafeInteger(item.id)&&typeof item.id==='number'&&!seen.has(item.id)&&typeof item.text==='string','Invalid edit','invalid');seen.add(item.id);
  const slot=before.slots[item.id];need(slot,'Unknown editable value','invalid');clean(item.text,slot.kind);need((length+=item.text.length)<=TEXT,'Replacement budget','limit');
  if(slot.shared){
   // Keep the shared-string uniqueness contract: an edit must not make
   // a different si entry equal to this one or change other cells.
   const root=before.documents.get(slot.part)!.documentElement!;
   need(!children(root,S,'si').some(si=>si!==slot.node.parentNode&&si.textContent===item.text),'Replacement duplicates another shared string');
  }
  if(slot.text!==item.text)changed.set(item.id,item.text);
 }
 if(!changed.size)return Uint8Array.from(bytes);
 const finalShared=new Map<Element,string>();for(const [id,text]of changed){const slot=before.slots[id]!;if(slot.shared)finalShared.set(slot.node.parentNode as Element,text);}
 for(const part of new Set([...changed.keys()].map(id=>before.slots[id]!).filter(s=>s.shared).map(s=>s.part))){const counts=new Map<string,number>();for(const si of children(before.documents.get(part)!.documentElement!,S,'si')){const text=finalShared.get(si)??si.textContent??'';counts.set(text,(counts.get(text)??0)+1);}for(const [si,text]of finalShared)if(si.ownerDocument===before.documents.get(part))need(counts.get(text)===1,'Replacement duplicates another shared string');}
 const original=new Map<string,Uint8Array>(),space=new Map<number,string|null>();
 for(const [i,text]of changed){
  const s=before.slots[i]!,doc=before.documents.get(s.part)!;if(!original.has(s.part))original.set(s.part,serializeXml(doc,options));
  space.set(i,s.node.hasAttributeNS(XML,'space')?s.node.getAttributeNS(XML,'space'):null);s.node.textContent=text;
  if(s.kind==='text'&&(/^\s|\s$/.test(text)))s.node.setAttributeNS(XML,'xml:space','preserve');
 }
 const replacements=new Map([...original.keys()].map(part=>[part,serializeXml(before.documents.get(part)!,options)]));
 const output=writeOpc(before.pkg,replacements,options),after=scan(output,format,options);
 need(after.slots.length===before.slots.length&&after.pkg.parts.size===before.pkg.parts.size,'Value edit changed structure','invalid');
 for(const [part,data]of before.pkg.parts)if(!replacements.has(part))need(bytesEqual(data,after.pkg.parts.get(part)??new Uint8Array()),'Unrelated part changed','invalid');
 for(const [i,s]of after.slots.entries()){
  const old=before.slots[i]!;need(s.part===old.part&&s.kind===old.kind&&s.text===(changed.get(i)??old.text),'Value did not round trip','invalid');
  if(changed.has(i)){s.node.textContent=old.text;const attr=space.get(i);if(attr==null)s.node.removeAttributeNS(XML,'space');else s.node.setAttributeNS(XML,'xml:space',attr);}
 }
 for(const [part,xml]of original)need(bytesEqual(serializeXml(after.documents.get(part)!,options),xml),'Unexpected XML change','invalid');
 return output;
}
