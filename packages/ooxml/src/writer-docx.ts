// SPDX-License-Identifier: MIT
/** M2 narrow WordprocessingML semantic profile, using the existing bounded OPC layer.
 * No named styles, table/image content, tracked changes, external links or implicit
 * section defaults may be rewritten. Unsupported documents stay on the inspection path.
 */
import type { Element, Document } from '@xmldom/xmldom';
import { OfficeError, assert, EMU_PER_TWIP } from '@office-ex/contracts';
import { createDocument, validateDocument, freezeData, type DeepReadonly, type WriterDocument, type WriterBlock, type Paragraph, type TextRun, type TextStyle, type WriterSection } from '@office-ex/core';
import { openOpc, writeOpc, type OpcPackage } from './opc.js';
import { parseXml } from './xml.js';
import { writeZip } from './zip.js';
import type { ContainerOptions } from './limits.js';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML='http://www.w3.org/XML/1998/namespace';
const XMLNS='http://www.w3.org/2000/xmlns/';
const ROOT_REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const MAIN_TYPE='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const DECL='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const CHAR_KEYS=['fontFamily','fontSize','bold','italic','underline','color'];
const PARA_KEYS=['alignment','indent','spaceBefore','spaceAfter','lineHeight'];
const MAX_BLOCKS=10000;
const MAX_RUNS=50000;
const MAX_TEXT=4*1024*1024;
function unsupported(detail:string):never{throw new OfficeError('unsupported','DOCX semantic profile: '+detail);}
function valid(value:unknown,detail:string):asserts value{if(!value)unsupported(detail);}
function isW(el:Element,name:string):boolean{return el.namespaceURI===W&&el.localName===name;}
function attrs(el:Element,allowed:readonly string[]=[]):void{
  for(let i=0;i<el.attributes.length;i++){
    const a=el.attributes.item(i)!;
    if(a.namespaceURI===XMLNS)continue;
    const name=a.namespaceURI===W?'w:'+a.localName:a.namespaceURI===XML?'xml:'+a.localName:'?'+a.name;
    valid(allowed.includes(name),'Unrecognized attribute on '+el.localName);
  }
}
function elements(el:Element):Element[]{
  const output:Element[]=[];
  for(let node=el.firstChild;node;node=node.nextSibling){
    if(node.nodeType===1)output.push(node as Element);
    else valid(node.nodeType===3&&!(node.nodeValue??'').trim(),'Non-element content in '+el.localName);
  }
  return output;
}
function requireElement(el:Element,name:string):void{valid(isW(el,name),'Expected '+name);}
function prop(el:Element,name:string):string|undefined{return el.hasAttributeNS(W,name)?el.getAttributeNS(W,name)??undefined:undefined;}
function numberProp(el:Element,name:string,min:number,max:number):number{
  const raw=prop(el,name);valid(raw!==undefined&&/^-?\d+$/.test(raw),'Invalid numeric '+name);
  const n=Number(raw);valid(Number.isSafeInteger(n)&&n>=min&&n<=max,'Out-of-range '+name);return n;
}
function empty(el:Element):void{valid(elements(el).length===0,'Unsupported nested element in '+el.localName);}
function toggled(el:Element):boolean{
  attrs(el,['w:val']);empty(el);const raw=prop(el,'val')??'1';
  valid(['0','1','true','false','on','off'].includes(raw),'Unknown boolean property');return ['1','true','on'].includes(raw);
}
function readRunStyle(el:Element):TextStyle{
  requireElement(el,'rPr');attrs(el);const style:TextStyle={},seen=new Set<string>();
  for(const child of elements(el)){
    valid(child.namespaceURI===W&&child.localName!==null&&!seen.has(child.localName),'Duplicate or foreign run property');seen.add(child.localName);
    switch(child.localName){
      case 'b':style.bold=toggled(child);break;
      case 'i':style.italic=toggled(child);break;
      case 'u':{attrs(child,['w:val']);empty(child);const value=prop(child,'val');valid(value==='single'||value==='none','Unsupported underline');style.underline=value==='single';break;}
      case 'sz':attrs(child,['w:val']);empty(child);style.fontSize=numberProp(child,'val',1,3276)*6350;break;
      case 'color':{attrs(child,['w:val']);empty(child);const color=prop(child,'val')??'';valid(/^[0-9a-fA-F]{6}$/.test(color),'Unsupported color');style.color={kind:'rgba',red:parseInt(color.slice(0,2),16),green:parseInt(color.slice(2,4),16),blue:parseInt(color.slice(4,6),16),alpha:255};break;}
      case 'rFonts':{attrs(child,['w:ascii','w:hAnsi','w:eastAsia','w:cs']);empty(child);const font=prop(child,'ascii');valid(font&&['hAnsi','eastAsia','cs'].every(n=>prop(child,n)===font),'Mixed or inherited font mapping');style.fontFamily=font;break;}
      default:unsupported('Unsupported run property '+child.localName);
    }
  }
  return style;
}
function readParagraphStyle(el:Element):TextStyle{
  requireElement(el,'pPr');attrs(el);const style:TextStyle={},seen=new Set<string>();
  for(const child of elements(el)){
    valid(child.namespaceURI===W&&child.localName!==null&&!seen.has(child.localName),'Duplicate or foreign paragraph property');seen.add(child.localName);
    switch(child.localName){
      case 'jc':{attrs(child,['w:val']);empty(child);const v=prop(child,'val');valid(v==='left'||v==='center'||v==='right'||v==='both','Unsupported alignment');style.alignment=v==='left'?'start':v==='right'?'end':v==='both'?'justify':v;break;}
      case 'ind':attrs(child,['w:left']);empty(child);style.indent=numberProp(child,'left',-31680,31680)*EMU_PER_TWIP;break;
      case 'spacing':{
        attrs(child,['w:before','w:after','w:line','w:lineRule']);empty(child);
        if(prop(child,'before')!==undefined)style.spaceBefore=numberProp(child,'before',0,31680)*EMU_PER_TWIP;
        if(prop(child,'after')!==undefined)style.spaceAfter=numberProp(child,'after',0,31680)*EMU_PER_TWIP;
        if(prop(child,'line')!==undefined){valid(prop(child,'lineRule')==='exact','Unsupported line-spacing rule');style.lineHeight=numberProp(child,'line',1,31680)*EMU_PER_TWIP;}
        else valid(prop(child,'lineRule')===undefined,'Line rule without line value');
        valid(Object.keys(style).some(k=>['spaceBefore','spaceAfter','lineHeight'].includes(k)),'Empty spacing property');
        break;
      }
      default:unsupported('Unsupported paragraph property '+child.localName);
    }
  }
  return style;
}
function readPage(el:Element):WriterSection['page']{
  requireElement(el,'sectPr');attrs(el);
  const items=elements(el);valid(items.length===2&&isW(items[0]!,'pgSz')&&isW(items[1]!,'pgMar'),'Explicit single-section page size and margins required');
  const size=items[0]!,margins=items[1]!;
  attrs(size,['w:w','w:h','w:orient']);empty(size);attrs(margins,['w:top','w:right','w:bottom','w:left','w:header','w:footer','w:gutter']);empty(margins);
  const width=numberProp(size,'w',1,31680),height=numberProp(size,'h',1,31680),orientation=prop(size,'orient');
  valid(orientation===undefined||orientation===(width>height?'landscape':'portrait'),'Inconsistent page orientation');
  for(const n of ['header','footer','gutter'])valid(prop(margins,n)===undefined||prop(margins,n)==='0','Unsupported header/footer/gutter offset');
  return {width:width*EMU_PER_TWIP,height:height*EMU_PER_TWIP,margins:{top:numberProp(margins,'top',0,31680)*EMU_PER_TWIP,right:numberProp(margins,'right',0,31680)*EMU_PER_TWIP,bottom:numberProp(margins,'bottom',0,31680)*EMU_PER_TWIP,left:numberProp(margins,'left',0,31680)*EMU_PER_TWIP}};
}
function simpleText(el:Element):string{
  attrs(el,['xml:space']);let text='';
  for(let n=el.firstChild;n;n=n.nextSibling){valid(n.nodeType===3,'Unsupported child inside w:t');text+=n.nodeValue??'';}
  const space=el.getAttributeNS(XML,'space');valid(!space||space==='preserve'||space==='default','Unknown whitespace mode');
  valid(!/^[ \t\r\n]|[ \t\r\n]$/.test(text)||space==='preserve','Ambiguous unpreserved whitespace');
  valid(!/[\r\n\t]/.test(text),'Raw line/tab control inside w:t');
  return text;
}
function readMain(xml:Document,documentId:string,title:string):WriterDocument{
  const root=xml.documentElement!;requireElement(root,'document');attrs(root);
  for(let n=xml.firstChild;n;n=n.nextSibling)if(n!==root)valid(n.nodeType===3&&!(n.nodeValue??'').trim()||n.nodeType===7&&n.nodeName==='xml','Unsupported document-level node');
  const contents=elements(root);valid(contents.length===1&&isW(contents[0]!,'body'),'Single body required');
  const body=contents[0]!;attrs(body);const children=elements(body);
  valid(children.length>=2&&children.length<=MAX_BLOCKS+1,'Missing paragraphs or block limit exceeded');
  const page=readPage(children.at(-1)!);const blocks:WriterBlock[]=[];let runCount=0,textCount=0;
  for(let bi=0;bi<children.length-1;bi++){
    const p=children[bi]!;requireElement(p,'p');attrs(p);const nodes=elements(p);
    let style:TextStyle|undefined;let i=0;
    if(nodes[0]&&isW(nodes[0],'pPr')){style=readParagraphStyle(nodes[0]);i++;}
    const runs:TextRun[]=[];let pageBreak=false;
    for(;i<nodes.length;i++){
      const r=nodes[i]!;requireElement(r,'r');attrs(r);valid(++runCount<=MAX_RUNS,'Run limit exceeded');
      const items=elements(r);let j=0;let direct:TextStyle|undefined;
      if(items[0]&&isW(items[0],'rPr')){direct=readRunStyle(items[0]);j++;}
      let text='';
      for(;j<items.length;j++){
        const child=items[j]!;
        if(isW(child,'t'))text+=simpleText(child);
        else if(isW(child,'tab')){attrs(child);empty(child);text+='\t';}
        else if(isW(child,'br')){
          attrs(child,['w:type']);empty(child);const type=prop(child,'type');
          if(type==='page'){
            valid(!style&&!direct&&nodes.length===1&&items.length===1,'Page break must occupy its own unformatted paragraph');pageBreak=true;
          }else{valid(type===undefined||type==='textWrapping','Unsupported break type');text+='\n';}
        }else unsupported('Unsupported run content '+child.localName);
      }
      textCount+=text.length;valid(textCount<=MAX_TEXT,'Text budget exceeded');
      if(!pageBreak)runs.push({id:'r-'+bi+'-'+runs.length,type:'text',text,...(direct&&Object.keys(direct).length?{style:direct}:{})});
    }
    if(pageBreak)blocks.push({id:'break-'+bi,type:'page-break'});
    else blocks.push({id:'p-'+bi,type:'paragraph',runs:runs.length?runs:[{id:'r-'+bi+'-0',type:'text',text:''}],...(style&&Object.keys(style).length?{style}:{})});
  }
  const result=createDocument('writer',documentId,title);result.defaultTextStyle={};result.sections=[{id:'section-1',page,blocks}];
  return validateDocument(result) as WriterDocument;
}
function escapeXml(value:string):string{
  for(const char of value){const c=char.codePointAt(0)!;valid(c===9||c===10||c===13||c>=32&&c<=0xD7FF||c>=0xE000&&c<=0xFFFD||c>=0x10000&&c<=0x10FFFF,'Illegal XML character');}
  return value.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;').replace(/\r/g,'&#13;').replace(/\n/g,'&#10;').replace(/\t/g,'&#9;');
}
function tag(name:string,properties:Record<string,string|number>):string{
  return '<w:'+name+Object.entries(properties).map(([k,v])=>' w:'+k+'="'+escapeXml(String(v))+'"').join('')+'/>';
}
function twip(value:number,min:number,max:number):number{
  const n=Math.round(value/EMU_PER_TWIP);valid(Number.isSafeInteger(n)&&n>=min&&n<=max,'Geometry outside DOCX limits');return n;
}
function writeRunStyle(style:DeepReadonly<TextStyle>):string{
  valid(Object.keys(style).every(k=>CHAR_KEYS.includes(k)),'Unsupported character style');
  let r='';
  if(style.fontFamily!==undefined)r+=tag('rFonts',{ascii:style.fontFamily,hAnsi:style.fontFamily,eastAsia:style.fontFamily,cs:style.fontFamily});
  for(const [key,xml] of [['bold','b'],['italic','i']] as const)if(style[key]!==undefined)r+=tag(xml,{val:style[key]?1:0});
  if(style.color!==undefined){
    valid(style.color.kind==='rgba'&&style.color.alpha===255,'Theme or translucent text color unsupported');
    const v=[style.color.red,style.color.green,style.color.blue].map(n=>n.toString(16).padStart(2,'0')).join('').toUpperCase();r+=tag('color',{val:v});
  }
  if(style.fontSize!==undefined){const n=Math.round(style.fontSize/6350);valid(n>=1&&n<=3276,'Font size outside DOCX limits');r+=tag('sz',{val:n});}
  if(style.underline!==undefined)r+=tag('u',{val:style.underline?'single':'none'});
  return r?'<w:rPr>'+r+'</w:rPr>':'';
}
function writeParagraphStyle(style:DeepReadonly<TextStyle>):string{
  valid(Object.keys(style).every(k=>PARA_KEYS.includes(k)),'Unsupported paragraph style');
  let p='';const spacing:Record<string,string|number>={};
  if(style.spaceBefore!==undefined)spacing.before=twip(style.spaceBefore,0,31680);
  if(style.spaceAfter!==undefined)spacing.after=twip(style.spaceAfter,0,31680);
  if(style.lineHeight!==undefined){spacing.line=twip(style.lineHeight,1,31680);spacing.lineRule='exact';}
  if(Object.keys(spacing).length)p+=tag('spacing',spacing);
  if(style.indent!==undefined)p+=tag('ind',{left:twip(style.indent,-31680,31680)});
  if(style.alignment!==undefined)p+=tag('jc',{val:({start:'left',center:'center',end:'right',justify:'both'})[style.alignment]});
  return p?'<w:pPr>'+p+'</w:pPr>':'';
}
function writeText(text:string):string{
  valid(!text.includes('\r'),'CR must be represented explicitly by a supported line break');
  if(!text)return '<w:t xml:space="preserve"></w:t>';
  return text.split(/(\n|\t)/).filter(value=>value!=='').map(value=>value==='\n'?'<w:br/>':value==='\t'?'<w:tab/>':'<w:t xml:space="preserve">'+escapeXml(value)+'</w:t>').join('');
}
function mainXml(candidate:unknown):{model:WriterDocument;xml:string}{
  const checked=validateDocument(candidate);valid(checked.kind==='writer','Expected Writer model');
  valid(checked.sections.length===1&&!Object.keys(checked.styles).length&&!Object.keys(checked.resources).length,'Only one section with direct formatting and no named styles or assets');
  valid(Object.keys(checked.defaultTextStyle).every(k=>CHAR_KEYS.includes(k)),'Unsupported default text style');
  const section=checked.sections[0]!;valid(section.blocks.length<=MAX_BLOCKS,'Block budget exceeded');
  let totalRuns=0,totalText=0;let body='';
  for(const block of section.blocks){
    if(block.type==='page-break'){body+='<w:p><w:r><w:br w:type="page"/></w:r></w:p>';continue;}
    valid(block.type==='paragraph'&&!block.styleId,'Tables and named paragraph styles unsupported');
    body+='<w:p>'+writeParagraphStyle(block.style??{});
    for(const inline of block.runs.length?block.runs:[{id:'empty',type:'text' as const,text:''}]){
      valid(++totalRuns<=MAX_RUNS,'Run budget exceeded');
      valid(inline.type!=='image','Images are outside this profile');
      if(inline.type==='text'){
        valid(!inline.styleId,'Named run style unsupported');totalText+=inline.text.length;valid(totalText<=MAX_TEXT,'Text budget exceeded');
        body+='<w:r>'+writeRunStyle({...checked.defaultTextStyle,...inline.style})+writeText(inline.text)+'</w:r>';
      }else body+='<w:r>'+writeRunStyle(checked.defaultTextStyle)+(inline.type==='tab'?'<w:tab/>':'<w:br/>')+'</w:r>';
    }
    body+='</w:p>';
  }
  const page=section.page,width=twip(page.width,1,31680),height=twip(page.height,1,31680),m=page.margins;
  body+='<w:sectPr>'+tag('pgSz',{w:width,h:height,...(width>height?{orient:'landscape'}:{})})+tag('pgMar',{top:twip(m.top,0,31680),right:twip(m.right,0,31680),bottom:twip(m.bottom,0,31680),left:twip(m.left,0,31680),header:0,footer:0,gutter:0})+'</w:sectPr>';
  return {model:checked,xml:DECL+'<w:document xmlns:w="'+W+'"><w:body>'+body+'</w:body></w:document>'};
}
function requireMinimal(pkg:OpcPackage):string{
  valid(pkg.kind==='docx'&&pkg.conformance==='transitional'&&pkg.mainPart,'Expected Transitional DOCX');
  valid(!pkg.hasMacros&&!pkg.hasSignatures&&!pkg.externalRelationships,'Protected or linked document');
  const main=pkg.mainPart;
  // Deliberately fail closed: any additional part or relationship can affect semantics.
  valid(pkg.parts.size===3&&pkg.parts.has('[Content_Types].xml')&&pkg.parts.has('_rels/.rels')&&pkg.parts.has(main),'Additional package parts require a later adapter');
  const root=pkg.relationships.get('')??[];
  valid(pkg.relationships.size===1&&root.length===1&&root[0]!.type===ROOT_REL&&root[0]!.targetPart===main&&root[0]!.fragment===null,'Unsupported relationship graph');
  return main;
}
export interface WriterDocxSource{readonly document:DeepReadonly<WriterDocument>;readonly mainPart:string}
const sources=new WeakMap<WriterDocxSource,{pkg:OpcPackage;canonical:string;documentId:string}>();
export function openWriterDocx(bytes:Uint8Array,documentId:string,title:string,options:ContainerOptions={}):WriterDocxSource{
  const pkg=openOpc(bytes,options),part=requireMinimal(pkg),document=readMain(parseXml(pkg.parts.get(part)!,options),documentId,title);
  const source=Object.freeze({document:freezeData(document),mainPart:part});
  sources.set(source,{pkg,canonical:mainXml(document).xml,documentId});
  return source;
}
const encoder=new TextEncoder();
export function writeWriterDocx(candidate:unknown,options:ContainerOptions={}):Uint8Array{
  const {model,xml}=mainXml(candidate),part='word/document.xml';
  const bytes=writeZip(new Map([
    ['[Content_Types].xml',encoder.encode(DECL+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/'+part+'" ContentType="'+MAIN_TYPE+'"/></Types>')],
    ['_rels/.rels',encoder.encode(DECL+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="'+ROOT_REL+'" Target="'+part+'"/></Relationships>')],
    [part,encoder.encode(xml)]
  ]),options);
  const checked=openWriterDocx(bytes,model.documentId,model.title,options);
  valid(mainXml(checked.document).xml===xml,'Semantic output verification failed');
  return bytes;
}
/** No-edit returns original archive bytes. Edited copy changes only the supported main part.
 * No IO occurs here: desktop save planning/atomic file commit belongs to a later M2 slice.
 */
export function saveWriterDocx(source:WriterDocxSource,candidate:unknown,options:ContainerOptions={}):Uint8Array{
  const saved=sources.get(source);assert(saved,'invalid','Unknown DOCX source capability');
  const {model,xml}=mainXml(candidate);assert(model.documentId===saved.documentId,'conflict','Document identity mismatch');
  if(xml===saved.canonical)return writeOpc(saved.pkg,new Map(),options);
  const bytes=writeOpc(saved.pkg,new Map([[source.mainPart,encoder.encode(xml)]]),options);
  const checked=openWriterDocx(bytes,model.documentId,model.title,options);
  valid(mainXml(checked.document).xml===xml,'Edited DOCX verification failed');
  return bytes;
}
