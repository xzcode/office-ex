// SPDX-License-Identifier: MIT
/** Source-preserving fallback, NOT the full Writer layout model.
 * The only writeable values are ordinary direct w:t nodes in ordinary body
 * paragraphs/table cells. Everything else remains inert source XML.
 * A fixed host-owned verifier rebuilds the output from authorized bytes.
 */
import type {Document,Element} from '@xmldom/xmldom';
import {OfficeError,assertUnicode,isRecord} from '@office-ex/contracts';
import {makeImageResource,sha256,type ImageResource} from '@office-ex/core';
import {openOpc,writeOpc,type OpcPackage} from './opc.js';
import {parseXml,serializeXml,childElements} from './xml.js';
import {bytesEqual,checkCancelled,type ContainerOptions} from './limits.js';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',V='urn:schemas-microsoft-com:vml',O='urn:schemas-microsoft-com:office:office',A='http://schemas.openxmlformats.org/drawingml/2006/main',WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',XML='http://www.w3.org/XML/1998/namespace';
const MAX_SOURCE=32*1024*1024,MAX_TEXT=2*1024*1024,MAX_SEGMENTS=20000,MAX_BLOCKS=10000,MAX_IMAGES=16*1024*1024,MAX_EMBEDS=32;
const XLSX='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
function need(value:unknown,message:string,code:'unsupported'|'invalid'|'unsafe'|'limit'|'conflict'='unsupported'):asserts value{if(!value)throw new OfficeError(code,'DOCX compatible text: '+message);}
const is=(e:Element,ns:string,name:string)=>e.namespaceURI===ns&&e.localName===name;
const w=(e:Element,name:string)=>is(e,W,name);
const get=(e:Element|undefined,name:string)=>e?.getAttributeNS(W,name)??undefined;
const find=(e:Element|undefined,ns:string,name:string)=>e?childElements(e).find(c=>is(c,ns,name)):undefined;
function descendants(e:Element,ns:string,name:string):Element[]{const result:Element[]=[];for(const child of childElements(e)){if(is(child,ns,name))result.push(child);result.push(...descendants(child,ns,name));}return result;}
function only(e:Element,ns:string,name:string):Element|undefined{const c=childElements(e).filter(x=>is(x,ns,name));need(c.length<=1,'Duplicate '+name,'invalid');return c[0];}
function enabled(e:Element|undefined):boolean{return !!e&&!['0','false','off'].includes(get(e,'val')??'1');}
function cleanText(value:string,max=32767):void{assertUnicode(value);need(value.length<=max&&!/[\x00-\x1f\x7f\ufffe\uffff]/u.test(value),'Unsupported text control or segment length','invalid');}
export interface CompatTextStyle{font?:string;size?:number;color?:string;fill?:string;bold?:boolean;italic?:boolean;strike?:boolean;underline?:boolean}
export interface CompatSegment{id:number;text:string;style:CompatTextStyle}
export type CompatItem={kind:'text';segment:CompatSegment}|{kind:'image';resourceId:string;embedded:boolean;width?:number}|{kind:'locked';label:'object'|'paragraph'};
export interface CompatParagraph{kind:'paragraph';items:CompatItem[];text:string;alignment?:'start'|'center'|'end';indent?:number;heading?:number;numbered?:boolean}
export type CompatBlock=CompatParagraph|{kind:'table';rows:CompatParagraph[][][]}|{kind:'locked';label:'object'|'paragraph'};
export interface CompatDocxView{sourceSha256:string;blocks:CompatBlock[];resources:Record<string,ImageResource>;editableSegments:number;paragraphCount:number;tables:number;images:number;embeddedWorkbooks:number;lockedBlocks:number}
interface Slot{element:Element;original:string}
interface Scan{pkg:OpcPackage;doc:Document;view:CompatDocxView;slots:Slot[]}
export interface CompatTextEdit{id:number;text:string}
export interface CompatDocxChanges{documentId:string;sourceSha256:string;edits:CompatTextEdit[]}
function checkActive(pkg:OpcPackage,options:ContainerOptions):void{
 need(pkg.kind==='docx'&&pkg.conformance==='transitional'&&pkg.mainPart==='word/document.xml','Expected standard Transitional DOCX');
 need(!pkg.hasMacros&&!pkg.hasSignatures&&!pkg.externalRelationships,'Macro, signature or external relationship','unsafe');
 // Closed source graph, including safe passive XML parts and embedded
 // modern spreadsheets. Do not open OLE binaries, links, ActiveX, or
 // alternate document chunks. Nested workbooks remain bytes, not objects.
 const main=pkg.mainPart,mainRels=pkg.relationships.get(main)??[];
 const rootAllowed=new Set([R+'/officeDocument',R+'/extended-properties','http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties','http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail']);
 need((pkg.relationships.get('')??[]).every(r=>rootAllowed.has(r.type)),'Unsupported root graph');
 const ordinary=new Set(['styles','numbering','settings','webSettings','fontTable','theme','header','footer','image']);
 let embedded=0;
 for(const [owner,rels]of pkg.relationships){
  if(!owner)continue;
  need(owner===main||mainRels.some(r=>r.targetPart===owner&&[R+'/header',R+'/footer'].includes(r.type)),'Unknown relationship owner');
  for(const rel of rels){
   checkCancelled(options);need(rel.mode==='Internal'&&rel.targetPart&&!rel.fragment,'External or fragment relationship','unsafe');
   const kind=rel.type.startsWith(R+'/')?rel.type.slice(R.length+1):rel.type==='http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects'?'styles':undefined;
   if(kind==='package'){
    need(owner===main&&pkg.contentTypes.get(rel.targetPart)===XLSX&&++embedded<=MAX_EMBEDS,'Only bounded embedded XLSX packages are supported');
    const nested=pkg.parts.get(rel.targetPart)!;need(nested.length<=8*1024*1024,'Embedded workbook budget','limit');
    const x=openOpc(nested,{...options,limits:{...options.limits,maxTotalBytes:Math.min(options.limits?.maxTotalBytes??32*1024*1024,32*1024*1024)}});
    need(x.kind==='xlsx'&&!x.hasMacros&&!x.hasSignatures&&!x.externalRelationships,'Unsafe embedded workbook','unsafe');
    // No nested embedded objects or active controls in a workbook.
    for(const rs of x.relationships.values())need(rs.every(r=>!/(?:\/(?:package|oleObject|control|activeXControl|externalLink|vbaProject|attachedTemplate))$/i.test(r.type)),'Nested active object','unsafe');
   }else need(kind&&ordinary.has(kind)&&(owner===main||kind==='image'),'Unsupported document relationship');
  }
 }
 // No protected editing or active field instructions. Reject tracked
 // changes/hidden text rather than expose their content in a fallback.
 for(const [path,bytes]of pkg.parts){
  if(!/\.xml$/i.test(path)||!path.startsWith('word/'))continue;
  checkCancelled(options);
  const root=parseXml(bytes,options).documentElement!;
  const blocked=['documentProtection','writeProtection','ins','del','moveFrom','moveTo','fldChar','fldSimple','instrText','altChunk','sdt','customXml','subDoc','vanish','webHidden','specVanish','pPrChange','rPrChange','tblPrChange','tcPrChange','trPrChange','sectPrChange'];
  const deny=new Set(blocked),stack=[root];while(stack.length){checkCancelled(options);const node=stack.pop()!;need(node.namespaceURI!==W||!deny.has(node.localName??''),'Protected, tracked, field or hidden text requires a separate adapter');stack.push(...childElements(node));}
 }
}
function scan(bytes:Uint8Array,options:ContainerOptions={},withImages=true):Scan{
 need(bytes.length>0&&bytes.length<=MAX_SOURCE,'Source size budget','limit');
 const pkg=openOpc(bytes,options);checkActive(pkg,options);
 const part=pkg.mainPart!,doc=parseXml(pkg.parts.get(part)!,options),root=doc.documentElement!,body=only(root,W,'body');
 need(w(root,'document')&&body,'Invalid main document','invalid');
 const slots:Slot[]=[],resources:Record<string,ImageResource>={},blocks:CompatBlock[]=[];
 let textLength=0,blockCount=0,imageBytes=0,pixels=0,paragraphCount=0,tables=0,images=0,embeddedWorkbooks=0,lockedBlocks=0;
 const rels=pkg.relationships.get(part)??[],relationships=new Map(rels.map(r=>[r.id,r]));
 function charge(text=''){checkCancelled(options);textLength+=text.length;need(textLength<=MAX_TEXT,'Text budget','limit');}
 function bump(){need(++blockCount<=MAX_BLOCKS,'Block budget','limit');checkCancelled(options);}
 function style(r:Element):CompatTextStyle{
  const pr=find(r,W,'rPr'),font=find(pr,W,'rFonts'),s:CompatTextStyle={};
  const face=get(font,'eastAsia')||get(font,'ascii')||get(font,'hAnsi');if(face&&face.length<=128&&!/[\x00-\x1f]/.test(face))s.font=face;
  const size=Number(get(find(pr,W,'sz'),'val'));if(Number.isFinite(size)&&size>=2&&size<=400)s.size=size/2;
  const color=get(find(pr,W,'color'),'val'),fill=get(find(pr,W,'shd'),'fill');if(color&&/^[0-9a-f]{6}$/i.test(color))s.color='#'+color;if(fill&&/^[0-9a-f]{6}$/i.test(fill))s.fill='#'+fill;
  s.bold=enabled(find(pr,W,'b'));s.italic=enabled(find(pr,W,'i'));s.strike=enabled(find(pr,W,'strike'));const u=get(find(pr,W,'u'),'val');s.underline=!!u&&u!=='none';return s;
 }
 function picture(element:Element,embedded:boolean):CompatItem{
  const refs=embedded?descendants(element,V,'imagedata'):descendants(element,A,'blip');
  const id=refs.length===1?refs[0]!.getAttributeNS(R,embedded?'id':'embed'):null,rel=id?relationships.get(id):undefined;
  need(rel?.type===R+'/image'&&rel.targetPart,'Missing passive preview image','invalid');
  const image=pkg.parts.get(rel.targetPart)!;
  need(['image/png','image/jpeg'].includes(pkg.contentTypes.get(rel.targetPart)??''),'Only passive raster images are displayed');
  if(embedded){
   const objects=descendants(element,O,'OLEObject');need(objects.length===1&&objects[0]!.getAttribute('Type')==='Embed'&&objects[0]!.getAttribute('ProgID')==='Excel.Sheet.12','Unsupported embedded object');
   const workbook=relationships.get(objects[0]!.getAttributeNS(R,'id')??'');
   need(workbook?.type===R+'/package'&&workbook.targetPart&&pkg.contentTypes.get(workbook.targetPart)===XLSX,'Invalid embedded workbook reference','invalid');embeddedWorkbooks++;
  }
  images++;need(images<=256,'Image count budget','limit');
  const key=rel.targetPart;
  let width:number|undefined;
  if(embedded){const shape=descendants(element,V,'shape')[0];const match=/(?:^|;)width:\s*(\d+(?:\.\d+)?)pt(?:;|$)/.exec(shape?.getAttribute('style')??'');if(match)width=Math.min(1200,Number(match[1])*4/3);}
  else{const extent=descendants(element,WP,'extent')[0],n=Number(extent?.getAttribute('cx'));if(Number.isFinite(n)&&n>0)width=Math.min(1200,n/9525);}
  if(withImages&&!Object.hasOwn(resources,key)){
   imageBytes+=image.length;need(imageBytes<=MAX_IMAGES,'Image byte budget','limit');
   const r=makeImageResource(image);
   // Some document producers store verified JPEG bytes in a PNG-named
   // part declared as image/png. Display the actual validated raster
   // MIME type; preserve the complete original package part and metadata.
   need(['image/png','image/jpeg'].includes(r.mimeType),'Unsupported actual raster format','invalid');
   pixels+=r.pixelWidth*r.pixelHeight;need(pixels<=80_000_000,'Image pixel budget','limit');resources[key]=r;
  }
  return {kind:'image',resourceId:key,embedded,...(width?{width}:{})};
 }
 function paragraph(p:Element):CompatParagraph{
  bump();paragraphCount++;
  const children=childElements(p),pr=only(p,W,'pPr'),items:CompatItem[]=[];
  const result:CompatParagraph={kind:'paragraph',items,text:''};
  const jc=get(find(pr,W,'jc'),'val');if(jc==='center')result.alignment='center';else if(jc==='right'||jc==='end')result.alignment='end';
  const ind=Number(get(find(pr,W,'ind'),'left'));if(Number.isFinite(ind)&&ind>0)result.indent=Math.min(160,ind/15);
  const level=get(find(pr,W,'outlineLvl'),'val');if(level!==undefined&&/^[0-8]$/.test(level))result.heading=Number(level)+1;
  result.numbered=!!find(pr,W,'numPr');
  // Only edit paragraphs whose visible text is direct and whose content
  // wrappers are inert. Never edit across a field, revision, or object.
  const plain=children.every(e=>['pPr','r','bookmarkStart','bookmarkEnd','proofErr'].some(n=>w(e,n)))&&children.filter(e=>w(e,'r')).every(r=>childElements(r).every(e=>w(e,'rPr')||w(e,'t')));
  for(const r of children){
   if(w(r,'pPr')||w(r,'bookmarkStart')||w(r,'bookmarkEnd')||w(r,'proofErr'))continue;
   if(!w(r,'r')){items.push({kind:'locked',label:'paragraph'});lockedBlocks++;continue;}
   for(const child of childElements(r)){
    if(w(child,'rPr'))continue;
    if(w(child,'object')){items.push(picture(child,true));continue;}
    if(w(child,'drawing')){items.push(picture(child,false));continue;}
    if(w(child,'t')){
     const text=child.textContent??'';charge(text);result.text+=text;
     if(!plain||childElements(child).length||/[\x00-\x1f\x7f]/.test(text)){continue;}
     cleanText(text);need(slots.length<MAX_SEGMENTS,'Text segment budget','limit');
     const id=slots.length;slots.push({element:child,original:text});items.push({kind:'text',segment:{id,text,style:style(r)}});
    }else if(w(child,'tab'))result.text+='\t';else if(w(child,'br')||w(child,'cr'))result.text+='\n';
    else{items.push({kind:'locked',label:'object'});lockedBlocks++;}
   }
  }
  if(!plain&&result.text){lockedBlocks++;/* visible but not an editable segment */}
  return result;
 }
 for(const e of childElements(body)){
  if(w(e,'p'))blocks.push(paragraph(e));
  else if(w(e,'tbl')){
   bump();tables++;const rows:CompatParagraph[][][]=[];
   const rs=childElements(e).filter(x=>w(x,'tr'));need(rs.length<=500,'Table row budget','limit');
   for(const row of rs){
    const cells=childElements(row).filter(x=>w(x,'tc'));need(cells.length<=64,'Table column budget','limit');
    rows.push(cells.map(cell=>{
     const ps=childElements(cell);need(ps.every(x=>w(x,'tcPr')||w(x,'p')),'Nested or complex table content requires a separate adapter');
     return ps.filter(x=>w(x,'p')).map(paragraph);
    }));
   }blocks.push({kind:'table',rows});
  }else if(!w(e,'sectPr')){bump();lockedBlocks++;blocks.push({kind:'locked',label:'paragraph'});}
 }
 need(slots.length>0,'No safely editable body text');
 return {pkg,doc,slots,view:{sourceSha256:sha256(bytes),blocks,resources,editableSegments:slots.length,paragraphCount,tables,images,embeddedWorkbooks,lockedBlocks}};
}
export function openWriterDocxCompatibility(bytes:Uint8Array,options:ContainerOptions={}):CompatDocxView{return scan(bytes,options).view;}
function validatedChanges(value:unknown,documentId:string):CompatDocxChanges{
 need(isRecord(value)&&Object.keys(value).every(k=>['documentId','sourceSha256','edits'].includes(k))&&value.documentId===documentId&&typeof value.sourceSha256==='string'&&/^[0-9a-f]{64}$/.test(value.sourceSha256)&&Array.isArray(value.edits)&&value.edits.length<=MAX_SEGMENTS,'Invalid text edit envelope','invalid');
 const seen=new Set<number>();let total=0;
 const edits=value.edits.map(item=>{
  need(isRecord(item)&&Object.keys(item).length===2&&typeof item.id==='number'&&Number.isSafeInteger(item.id)&&item.id>=0&&item.id<MAX_SEGMENTS&&typeof item.text==='string'&&!seen.has(item.id),'Invalid or repeated text segment','invalid');
  cleanText(item.text);seen.add(item.id);total+=item.text.length;need(total<=MAX_TEXT,'Replacement text budget','limit');return {id:item.id,text:item.text};
 });
 return {documentId,sourceSha256:value.sourceSha256,edits};
}
export function saveWriterDocxCompatibility(source:Uint8Array,documentId:string,changes:unknown,options:ContainerOptions={}):Uint8Array{
 const request=validatedChanges(changes,documentId),before=scan(source,options,false);
 need(request.sourceSha256===before.view.sourceSha256,'Original source digest mismatch','conflict');
 const changed=new Map<number,string>();
 for(const edit of request.edits){const slot=before.slots[edit.id];need(slot,'Unknown text segment','invalid');if(slot.original!==edit.text)changed.set(edit.id,edit.text);}
 if(!changed.size)return Uint8Array.from(source);
 const originalXml=serializeXml(before.doc,options),spaces=new Map<number,string|null>();
 for(const [id,value]of changed){
  const e=before.slots[id]!.element;spaces.set(id,e.hasAttributeNS(XML,'space')?e.getAttributeNS(XML,'space'):null);
  e.textContent=value;if(/^\s|\s$/.test(value))e.setAttributeNS(XML,'xml:space','preserve');
 }
 const main=before.pkg.mainPart!,part=serializeXml(before.doc,options);
 const output=writeOpc(before.pkg,new Map([[main,part]]),options);
 const after=scan(output,options,false);
 need(after.slots.length===before.slots.length&&after.pkg.parts.size===before.pkg.parts.size,'Text edit changed document topology','invalid');
 for(const [path,data]of before.pkg.parts)if(path!==main)need(bytesEqual(data,after.pkg.parts.get(path)??new Uint8Array()),'Unrelated package part changed','invalid');
 for(const [id,slot]of after.slots.entries()){
  need(slot.original===(changed.get(id)??before.slots[id]!.original),'Text segment did not round trip','invalid');
  if(changed.has(id)){slot.element.textContent=before.slots[id]!.original;const space=spaces.get(id);if(space==null)slot.element.removeAttributeNS(XML,'space');else slot.element.setAttributeNS(XML,'xml:space',space);}
 }
 // Prove that undoing just the requested text values yields the same
 // parsed XML, including bookmarks, styles, table properties and OLE/VML.
 need(bytesEqual(serializeXml(after.doc,options),originalXml),'Unexpected structural XML change','invalid');
 return output;
}
