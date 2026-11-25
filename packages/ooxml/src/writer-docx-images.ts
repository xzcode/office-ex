// SPDX-License-Identifier: MIT
/** Closed DrawingML profile: one embedded inline raster per image paragraph.
 * Images are content-addressed; new media/relationship/type entries are explicit.
 * Unchanged source media and unrelated parts are preserved.
 */
import type {Element} from '@xmldom/xmldom';
import {OfficeError,assert} from '@office-ex/contracts';
import {makeImageResource,validateImageResource,decodeImageBytes,type ImageResource,type DeepReadonly,type ResourceRef,type Inline} from '@office-ex/core';
import {openOpc,writeOpc,type OpcPackage} from './opc.js';
import {parseXml,serializeXml} from './xml.js';
import {writeZip} from './zip.js';
import type {ContainerOptions} from './limits.js';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',A='http://schemas.openxmlformats.org/drawingml/2006/main',PIC='http://schemas.openxmlformats.org/drawingml/2006/picture';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',PKG='http://schemas.openxmlformats.org/package/2006/relationships',CT='http://schemas.openxmlformats.org/package/2006/content-types',XMLNS='http://www.w3.org/2000/xmlns/';
export const IMAGE_REL=R+'/image';
type ImageInline=Extract<Inline,{type:'image'}>;
export interface MediaGraph{resources:Record<string,ImageResource>;byRelationship:Map<string,string>;byResource:Map<string,string>;usedDrawingIds:Set<number>}
export interface MediaPlan{relationships:ReadonlyMap<string,string>;additions:readonly {id:string;path:string;resource:DeepReadonly<ImageResource>}[]}
function valid(v:unknown,reason:string):asserts v{if(!v)throw new OfficeError('unsupported','DOCX image profile: '+reason);}
function element(el:Element,ns:string,n:string):void{valid(el.namespaceURI===ns&&el.localName===n,'Expected '+n);}
function attrs(el:Element,allowed:readonly string[]=[]):void{
  for(let i=0;i<el.attributes.length;i++){const a=el.attributes.item(i)!;if(a.namespaceURI===XMLNS)continue;valid(allowed.includes(a.namespaceURI===R?'r:'+a.localName:a.namespaceURI?'?'+a.name:a.name),'Unexpected '+el.localName+' attribute');}
}
function children(el:Element):Element[]{const result:Element[]=[];for(let n=el.firstChild;n;n=n.nextSibling){if(n.nodeType===1)result.push(n as Element);else valid(n.nodeType===3&&!(n.nodeValue??'').trim(),'Unexpected image content');}return result;}
function seq(el:Element,names:readonly (readonly [string,string])[]):Element[]{
  const c=children(el);valid(c.length===names.length,'Unsupported drawing children');
  for(let i=0;i<c.length;i++)element(c[i]!,names[i]![0],names[i]![1]);return c;
}
function empty(el:Element):void{valid(!children(el).length,'Unexpected nested image content');}
function integer(el:Element,key:string,min=0,max=100_000_000):number{const s=el.getAttribute(key);valid(s!==null&&/^\d+$/.test(s),'Invalid drawing number');const v=Number(s);valid(Number.isSafeInteger(v)&&v>=min&&v<=max,'Drawing number out of bounds');return v;}
function zeros(el:Element,keys:readonly string[]):void{attrs(el,keys);empty(el);for(const k of keys)if(el.hasAttribute(k))valid(integer(el,k)===0,'Nonzero drawing offset');}
function positiveSize(el:Element):{width:number;height:number}{attrs(el,['cx','cy']);empty(el);return {width:integer(el,'cx',1),height:integer(el,'cy',1)};}
function readNonvisual(el:Element,positive:boolean):{id:number;name:string}{attrs(el,['id','name','descr','title']);empty(el);const id=integer(el,'id',positive?1:0,0x7fffffff);valid(!el.getAttribute('descr')&&!el.getAttribute('title'),'Image alt metadata not supported by this profile');const name=el.getAttribute('name');valid(name!==null&&name.length<=512&&!/[\x00-\x1f\x7f]/.test(name),'Invalid image name');return {id,name};}
export function loadMedia(pkg:OpcPackage,main:string):MediaGraph{
  const g:MediaGraph={resources:Object.create(null),byRelationship:new Map(),byResource:new Map(),usedDrawingIds:new Set()};
  let size=0;
  for(const r of pkg.relationships.get(main)??[]){
    if(r.type!==IMAGE_REL)continue;
    valid(r.mode==='Internal'&&r.targetPart&&r.fragment===null,'External or invalid image target');
    const data=pkg.parts.get(r.targetPart),type=pkg.contentTypes.get(r.targetPart);
    valid(data&&(type==='image/png'||type==='image/jpeg'),'Unsupported image content type');
    const resource=makeImageResource(data);valid(resource.mimeType===type,'Image type mismatch');
    valid(!Object.hasOwn(g.resources,resource.id),'Duplicate media content uses different relationships');
    size+=resource.byteLength;valid(size<=16*1024*1024,'Document image payload budget exceeded');
    g.resources[resource.id]=resource;g.byRelationship.set(r.id,resource.id);g.byResource.set(resource.id,r.id);
  }
  return g;
}
/** Parse only the fully represented inline picture geometry, no wrapping,
 * crops, transforms, effects, linking, arbitrary graphics or external targets.
 */
export function readInlinePicture(drawing:Element,graph:MediaGraph,id:string):ImageInline{
  element(drawing,W,'drawing');attrs(drawing);
  const [inline]=seq(drawing,[[WP,'inline']]);attrs(inline!,['distT','distB','distL','distR']);for(const k of ['distT','distB','distL','distR'])if(inline!.hasAttribute(k))valid(integer(inline!,k)===0,'Inline distances must be zero');
  const c=children(inline!);let i=0;const extent=c[i++]!;valid(extent,'Missing drawing extent');element(extent,WP,'extent');const size=positiveSize(extent);
  if(c[i]?.namespaceURI===WP&&c[i]?.localName==='effectExtent')zeros(c[i++]!,['l','t','r','b']);
  const docPr=c[i++]!;valid(docPr,'Missing drawing properties');element(docPr,WP,'docPr');const drawingProps=readNonvisual(docPr,true),drawingId=drawingProps.id;valid(!graph.usedDrawingIds.has(drawingId),'Duplicate drawing ID');graph.usedDrawingIds.add(drawingId);
  if(c[i]?.namespaceURI===WP&&c[i]?.localName==='cNvGraphicFramePr'){
    const p=c[i++]!;attrs(p);const x=children(p);valid(x.length<=1,'Unsupported frame properties');
    if(x.length){element(x[0]!,A,'graphicFrameLocks');attrs(x[0]!,['noChangeAspect']);empty(x[0]!);valid(x[0]!.getAttribute('noChangeAspect')==='1','Unsupported graphic locks');}
  }
  const graphic=c[i++]!;valid(graphic&&i===c.length,'Unsupported inline drawing');element(graphic,A,'graphic');attrs(graphic);
  const [data]=seq(graphic,[[A,'graphicData']]);attrs(data!,['uri']);valid(data!.getAttribute('uri')===PIC,'Expected picture graphic');
  const [pic]=seq(data!,[[PIC,'pic']]);attrs(pic!);const [nv,fill,sp]=seq(pic!,[[PIC,'nvPicPr'],[PIC,'blipFill'],[PIC,'spPr']]);
  attrs(nv!);const [cnv,props]=seq(nv!,[[PIC,'cNvPr'],[PIC,'cNvPicPr']]);const pictureProps=readNonvisual(cnv!,false);attrs(props!);empty(props!);
  attrs(fill!);const [blip,stretch]=seq(fill!,[[A,'blip'],[A,'stretch']]);attrs(blip!,['r:embed']);empty(blip!);
  const rel=blip!.getAttributeNS(R,'embed')??'',resourceId=graph.byRelationship.get(rel);valid(resourceId,'Missing embedded image relationship');
  attrs(stretch!);const [rect]=seq(stretch!,[[A,'fillRect']]);attrs(rect!);empty(rect!);
  attrs(sp!);const [transform,geometry]=seq(sp!,[[A,'xfrm'],[A,'prstGeom']]);attrs(transform!);
  const [off,ext]=seq(transform!,[[A,'off'],[A,'ext']]);zeros(off!,['x','y']);const second=positiveSize(ext!);
  valid(second.width===size.width&&second.height===size.height,'Picture transform and extent disagree');
  attrs(geometry!,['prst']);valid(geometry!.getAttribute('prst')==='rect','Only rectangular picture geometry');const [av]=seq(geometry!,[[A,'avLst']]);attrs(av!);empty(av!);
  return {id,type:'image',resourceId,...size,drawingName:drawingProps.name,pictureName:pictureProps.name};
}
function relationshipPath(main:string):string{const i=main.lastIndexOf('/');return main.slice(0,i+1)+'_rels/'+main.slice(i+1)+'.rels';}
export function planMedia(resources:Readonly<Record<string,DeepReadonly<ResourceRef>>>,main:string,source?:OpcPackage,baseline?:MediaGraph):MediaPlan{
  const relationships=new Map<string,string>(),additions:{id:string;path:string;resource:DeepReadonly<ImageResource>}[]=[];
  for(const [id,resource] of Object.entries(baseline?.resources??{})){const actual=resources[id];valid(!actual||actual.sha256===resource.sha256&&actual.byteLength===resource.byteLength,'Source image resource mismatch');}
  const used=new Set((source?.relationships.get(main)??[]).map(r=>r.id));let total=0;
  for(const resource of Object.values({...baseline?.resources,...resources}).sort((a,b)=>a.id<b.id?-1:a.id>b.id?1:0)){
    validateImageResource(resource);total+=resource.byteLength;valid(total<=16*1024*1024,'Document image budget exceeded');
    const previous=baseline?.byResource.get(resource.id);
    if(previous){relationships.set(resource.id,previous);continue;}
    const id='rIdOexImage_'+resource.sha256,path=main.slice(0,main.lastIndexOf('/')+1)+'media/oex-'+resource.sha256+(resource.mimeType==='image/png'?'.png':'.jpeg');
    valid(!used.has(id)&&!source?.parts.has(path),'Image relationship or part collision');used.add(id);
    relationships.set(resource.id,id);additions.push({id,path,resource});
  }
  return {relationships,additions};
}
export function writeInlinePicture(image:DeepReadonly<ImageInline>,plan:MediaPlan,serial:number):string{
  const rel=plan.relationships.get(image.resourceId);valid(rel,'Missing image resource in save plan');
  const {width,height}=image;valid(Number.isSafeInteger(width)&&Number.isSafeInteger(height)&&width>0&&height>0&&width<=100000000&&height<=100000000,'Image geometry exceeds DrawingML limits');
  const escape=(v:string)=>{valid(v.length<=512&&!/[\x00-\x1f\x7f]/.test(v),'Invalid picture name');return v.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');};
  const drawingName=escape(image.drawingName??'Picture '+serial),pictureName=escape(image.pictureName??'Picture '+serial);
  return '<w:drawing><wp:inline xmlns:wp="'+WP+'" xmlns:a="'+A+'" xmlns:pic="'+PIC+'" xmlns:r="'+R+'" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="'+width+'" cy="'+height+'"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="'+serial+'" name="'+drawingName+'"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="'+PIC+'"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="'+pictureName+'"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="'+rel+'"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'+width+'" cy="'+height+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';
}
export function addMediaParts(parts:Map<string,Uint8Array>,main:string,plan:MediaPlan,options:ContainerOptions):void{
  if(!plan.additions.length)return;
  const types=parseXml(parts.get('[Content_Types].xml')!,options),root=types.documentElement!;
  const expected=relationshipPath(main),actual=[...parts.keys()].find(p=>p.toLowerCase()===expected.toLowerCase())??expected;
  const rels=parseXml(parts.get(actual)??'<Relationships xmlns="'+PKG+'"/>',options),relRoot=rels.documentElement!;
  for(const item of plan.additions){
    const override=types.createElementNS(CT,'Override');override.setAttribute('PartName','/'+item.path);override.setAttribute('ContentType',item.resource.mimeType);root.appendChild(override);
    const rel=rels.createElementNS(PKG,'Relationship');rel.setAttribute('Id',item.id);rel.setAttribute('Type',IMAGE_REL);rel.setAttribute('Target','/'+item.path);relRoot.appendChild(rel);
    parts.set(item.path,decodeImageBytes(item.resource.dataBase64));
  }
  parts.set('[Content_Types].xml',serializeXml(types,options));parts.set(actual,serializeXml(rels,options));
}
/** Source was already admitted by the closed semantic profile. Verify its
 * immutable original map via writeOpc before explicitly adding media parts.
 */
export function saveMediaPackage(pkg:OpcPackage,main:string,xml:string,plan:MediaPlan,options:ContainerOptions):Uint8Array{
  const replacement=new TextEncoder().encode(xml);
  if(!plan.additions.length)return writeOpc(pkg,new Map([[main,replacement]]),options);
  writeOpc(pkg,new Map(),options);
  const parts=new Map(pkg.parts);parts.set(main,replacement);addMediaParts(parts,main,plan,options);
  const result=writeZip(parts,options);openOpc(result,options);return result;
}
