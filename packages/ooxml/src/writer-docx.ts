// SPDX-License-Identifier: MIT
/** M2 narrow WordprocessingML semantic profile, using the existing bounded OPC layer.
 * Only supported named styles and preserved themes; nested/merged tables, images, tracked changes, external links and unsupported
 * section defaults may be rewritten. Unsupported documents stay on the inspection path.
 */
import type { Element, Document } from '@xmldom/xmldom';
import { OfficeError, assert, EMU_PER_TWIP } from '@office-ex/contracts';
import { createDocument, validateDocument, buildWriterNumberingMarkers, freezeData, isDocumentLanguageTag, type DeepReadonly, type WriterDocument, type WriterBlock, type Paragraph, type TextRun, type Inline, type TextStyle, type WriterSection, type WriterPageStory, type NamedStyle, type Table } from '@office-ex/core';
import { openOpc, writeOpc, type OpcPackage } from './opc.js';
import { parseXml, serializeXml } from './xml.js';
import { writeZip } from './zip.js';
import type { ContainerOptions } from './limits.js';
import { loadMedia, planMedia, readInlinePicture, writeInlinePicture, addMediaParts, saveMediaPackage, type MediaGraph, type MediaPlan } from './writer-docx-images.js';
import { readStyleProfile, type DocxStyleProfile, type DocxStorySource, type ThemeProfile, requireStyleReference } from './writer-docx-styles.js';

import {readNumberingReference,numberingChanged,addNumberingParts} from './writer-docx-numbering.js';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML='http://www.w3.org/XML/1998/namespace';
const XMLNS='http://www.w3.org/2000/xmlns/';
const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const ROOT_REL=REL+'/officeDocument';
const CT='http://schemas.openxmlformats.org/package/2006/content-types',PKG='http://schemas.openxmlformats.org/package/2006/relationships';
const MAIN_TYPE='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const DECL='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const CHAR_KEYS=['fontFamily','fontAscii','fontHighAnsi','fontEastAsia','fontComplex','fontSize','complexFontSize','bold','italic','underline','color','language','eastAsiaLanguage','bidiLanguage','runDirection','characterSpacing','kerningThreshold'];
const PARA_KEYS=['alignment','indent','spaceBefore','spaceAfter','lineHeight','lineMultiple','minLineHeight','keepNext','keepLines','outlineLevel','widowControl','paragraphDirection','rightIndent','firstLineIndent','contextualSpacing','paragraphBottomBorder','writerList'];
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
    if(a.namespaceURI==='http://schemas.openxmlformats.org/markup-compatibility/2006'&&a.localName==='Ignorable'&&['hdr','ftr'].includes(el.localName??'')&&a.value.split(/\s+/).every(p=>p&&el.lookupNamespaceURI(p)?.startsWith('http://schemas.microsoft.com/office/')))continue;
    // rsid identifies an editing session, not tracked-change content. No
    // revision markup is admitted. Ignore only the enumerated metadata attrs.
    if(a.namespaceURI===W&&['rsidR','rsidRPr','rsidRDefault','rsidP','rsidDel','rsidSect'].includes(a.localName??'')&&['p','r','sectPr'].includes(el.localName??'')&&/^[0-9A-Fa-f]{8}$/.test(a.value))continue;
    if(a.namespaceURI==='http://schemas.openxmlformats.org/markup-compatibility/2006'&&a.localName==='Ignorable'&&el.localName==='document'&&a.value.split(/\s+/).every(prefix=>prefix&&el.lookupNamespaceURI(prefix)?.startsWith('http://schemas.microsoft.com/office/')))continue;
    const name=a.namespaceURI===W?'w:'+a.localName:a.namespaceURI===XML?'xml:'+a.localName:a.namespaceURI===REL?'r:'+a.localName:'?'+a.name;
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
export function readRunStyle(el:Element,reference?:(id:string)=>void,theme?:ThemeProfile,named=false,toggles?:('bold'|'italic')[]):TextStyle{
  requireElement(el,'rPr');attrs(el);const style:TextStyle={},seen=new Set<string>();
  for(const child of elements(el)){
    valid(child.namespaceURI===W&&child.localName!==null&&!seen.has(child.localName),'Duplicate or foreign run property');seen.add(child.localName);
    switch(child.localName){
      case 'lang':{attrs(child,['w:val','w:eastAsia','w:bidi']);empty(child);let count=0;for(const [key,property] of [['val','language'],['eastAsia','eastAsiaLanguage'],['bidi','bidiLanguage']] as const){const lang=prop(child,key);if(lang!==undefined){valid(isDocumentLanguageTag(lang),'Unsupported document language tag');style[property]=lang;count++;}}valid(count>0,'Empty language property');break;}
      case 'rtl':{valid(!named,'Named RTL toggle is outside this profile');style.runDirection=toggled(child)?'rtl':'ltr';break;}
      case 'rStyle':{valid(reference,'Unexpected character style reference');attrs(child,['w:val']);empty(child);const id=prop(child,'val');valid(id,'Missing character style reference');reference(id);break;}
      case 'b':{const value=toggled(child);if(named){if(value)toggles?.push('bold');}else style.bold=value;break;}
      case 'i':{const value=toggled(child);if(named){if(value)toggles?.push('italic');}else style.italic=value;break;}
      case 'u':{attrs(child,['w:val']);empty(child);const value=prop(child,'val');valid(value==='single'||value==='none','Unsupported underline');style.underline=value==='single';break;}
      case 'spacing':attrs(child,['w:val']);empty(child);style.characterSpacing=numberProp(child,'val',-31680,31680)*EMU_PER_TWIP;break;
      case 'kern':attrs(child,['w:val']);empty(child);style.kerningThreshold=numberProp(child,'val',0,3276)*6350;break;
      case 'sz':attrs(child,['w:val']);empty(child);style.fontSize=numberProp(child,'val',1,3276)*6350;break;
      case 'szCs':attrs(child,['w:val']);empty(child);style.complexFontSize=numberProp(child,'val',1,3276)*6350;break;
      case 'bCs':case 'iCs':{valid(named||theme,'Independent complex-script toggle requires preserved source style');const value=toggled(child);const corresponding=child.localName==='bCs'?'bold':'italic';const direct=elements(el).find(e=>isW(e,corresponding==='bold'?'b':'i'));valid(direct&&toggled(direct)===value,'Independent complex-script toggle needs a later adapter');break;}
      case 'color':{attrs(child,theme?['w:val','w:themeColor','w:themeShade','w:themeTint']:['w:val']);empty(child);const key=prop(child,'themeColor'),shade=prop(child,'themeShade'),tint=prop(child,'themeTint');valid(!(shade&&tint)&&(!shade||/^[0-9a-fA-F]{2}$/.test(shade))&&(!tint||/^[0-9a-fA-F]{2}$/.test(tint))&&(!(shade||tint)||key&&theme?.colors[key]),'Invalid theme transform');const color=shade||tint?prop(child,'val')??'':key!==undefined?(theme?.colors[key]??''):prop(child,'val')??'';valid(/^[0-9a-fA-F]{6}$/.test(color),'Unsupported color');style.color={kind:'rgba',red:parseInt(color.slice(0,2),16),green:parseInt(color.slice(2,4),16),blue:parseInt(color.slice(4,6),16),alpha:255};break;}
      case 'rFonts':{attrs(child,theme?['w:ascii','w:hAnsi','w:eastAsia','w:cs','w:asciiTheme','w:hAnsiTheme','w:eastAsiaTheme','w:cstheme']:['w:ascii','w:hAnsi','w:eastAsia','w:cs']);empty(child);const fontKeys=[['ascii','asciiTheme'],['hAnsi','hAnsiTheme'],['eastAsia','eastAsiaTheme'],['cs','cstheme']] as const;const fonts=fontKeys.map(([literal,themed])=>{const ref=prop(child,themed);return ref!==undefined?theme?.fonts[ref]:prop(child,literal);});const font=fonts[0]??fonts[1];valid(font,'Missing Latin font mapping');
        for(const [literal,themed] of fontKeys){const ref=prop(child,themed);if(ref!==undefined)valid(theme&&Object.hasOwn(theme.fonts,ref),'Unknown theme font');}
        if(fonts.every(f=>f===font))style.fontFamily=font;
        else {style.fontFamily=font;for(let i=0;i<fonts.length;i++)if(fonts[i])style[(['fontAscii','fontHighAnsi','fontEastAsia','fontComplex'] as const)[i]!]=fonts[i];}
        break;}
      default:unsupported('Unsupported run property '+child.localName);
    }
  }
  return style;
}
export function readParagraphStyle(el:Element,reference?:(id:string)=>void,sectionEnd?:Element,theme?:ThemeProfile):TextStyle{
  requireElement(el,'pPr');attrs(el);const style:TextStyle={},seen=new Set<string>();
  for(const child of elements(el)){
    valid(child.namespaceURI===W&&child.localName!==null&&!seen.has(child.localName),'Duplicate or foreign paragraph property');seen.add(child.localName);
    switch(child.localName){
      case 'numPr':style.writerList=readNumberingReference(child);break;
      case 'bidi':style.paragraphDirection=toggled(child)?'rtl':'ltr';break;
      case 'keepNext':style.keepNext=toggled(child);break;
      case 'keepLines':style.keepLines=toggled(child);break;
      case 'widowControl':style.widowControl=toggled(child);break;
      case 'outlineLvl':attrs(child,['w:val']);empty(child);style.outlineLevel=numberProp(child,'val',0,9);break;
      case 'sectPr':{valid(child===sectionEnd,'Unexpected section properties inside paragraph');break;}
      case 'pStyle':{valid(reference,'Unexpected paragraph style reference');attrs(child,['w:val']);empty(child);const id=prop(child,'val');valid(id,'Missing paragraph style reference');reference(id);break;}
      case 'jc':{attrs(child,['w:val']);empty(child);const v=prop(child,'val');valid(v==='left'||v==='center'||v==='right'||v==='both','Unsupported alignment');style.alignment=v==='left'?'start':v==='right'?'end':v==='both'?'justify':v;break;}
      case 'ind':{
        attrs(child,['w:left','w:right','w:firstLine','w:hanging']);empty(child);
        valid(!(prop(child,'firstLine')!==undefined&&prop(child,'hanging')!==undefined),'Conflicting first-line and hanging indent');
        if(prop(child,'left')!==undefined)style.indent=numberProp(child,'left',-31680,31680)*EMU_PER_TWIP;
        if(prop(child,'right')!==undefined)style.rightIndent=numberProp(child,'right',0,31680)*EMU_PER_TWIP;
        if(prop(child,'firstLine')!==undefined)style.firstLineIndent=numberProp(child,'firstLine',0,31680)*EMU_PER_TWIP;
        if(prop(child,'hanging')!==undefined)style.firstLineIndent=-numberProp(child,'hanging',0,31680)*EMU_PER_TWIP;
        valid((style.indent!==undefined||style.rightIndent!==undefined||style.firstLineIndent!==undefined),'Empty indentation property');
        break;
      }
      case 'contextualSpacing':style.contextualSpacing=toggled(child);break;
      case 'pBdr':{
        attrs(child);const edges=elements(child);valid(edges.length===1&&isW(edges[0]!,'bottom'),'Only a single bottom paragraph border is supported');
        const border=edges[0]!;attrs(border,theme?['w:val','w:sz','w:space','w:color','w:themeColor','w:themeTint','w:themeShade']:['w:val','w:sz','w:space','w:color']);empty(border);
        const kind=prop(border,'val');
        if(kind==='none'||kind==='nil'){valid(['sz','space','color','themeColor','themeTint','themeShade'].every(k=>prop(border,k)===undefined),'Unexpected disabled border attributes');style.paragraphBottomBorder=null;break;}
        valid(kind==='single','Unsupported paragraph border style');
        const key=prop(border,'themeColor'),tint=prop(border,'themeTint'),shade=prop(border,'themeShade');
        valid(!(shade&&tint)&&(!shade||/^[0-9a-fA-F]{2}$/.test(shade))&&(!tint||/^[0-9a-fA-F]{2}$/.test(tint))&&(!(shade||tint)||key&&theme?.colors[key]),'Invalid border color transform');
        const hex=shade||tint?prop(border,'color'):key?theme?.colors[key]:prop(border,'color');
        valid(hex&&/^[0-9a-f]{6}$/i.test(hex),'Unsupported paragraph border color');
        const color={kind:'rgba' as const,red:parseInt(hex.slice(0,2),16),green:parseInt(hex.slice(2,4),16),blue:parseInt(hex.slice(4,6),16),alpha:255};
        style.paragraphBottomBorder={color,width:Math.round(numberProp(border,'sz',2,96)*12700/8),space:prop(border,'space')===undefined?0:numberProp(border,'space',0,31)*12700};break;
      }
      case 'spacing':{
        attrs(child,['w:before','w:after','w:line','w:lineRule']);empty(child);
        if(prop(child,'before')!==undefined)style.spaceBefore=numberProp(child,'before',0,31680)*EMU_PER_TWIP;
        if(prop(child,'after')!==undefined)style.spaceAfter=numberProp(child,'after',0,31680)*EMU_PER_TWIP;
        if(prop(child,'line')!==undefined){const rule=prop(child,'lineRule')??'auto',n=numberProp(child,'line',1,31680);if(rule==='exact')style.lineHeight=n*EMU_PER_TWIP;else if(rule==='auto')style.lineMultiple=n;else if(rule==='atLeast')style.minLineHeight=n*EMU_PER_TWIP;else unsupported('Unsupported line spacing rule');}
        else valid(prop(child,'lineRule')===undefined,'Line rule without line value');
        valid(Object.keys(style).some(k=>['spaceBefore','spaceAfter','lineHeight','lineMultiple','minLineHeight'].includes(k)),'Empty spacing property');
        break;
      }
      default:unsupported('Unsupported paragraph property '+child.localName);
    }
  }
  return style;
}
function readPage(el:Element,stories:Record<string,WriterPageStory>):Pick<WriterSection,'page'|'pageExtras'|'storyRefs'|'pageNumberStart'>{
  requireElement(el,'sectPr');attrs(el);const items=elements(el);let i=0;
  const refs:NonNullable<WriterSection['storyRefs']>={};
  while(items[i]&&(isW(items[i]!,'headerReference')||isW(items[i]!,'footerReference'))){
    const ref=items[i++]!,kind=isW(ref,'headerReference')?'header':'footer';
    attrs(ref,['w:type','r:id']);empty(ref);const id=ref.getAttributeNS(REL,'id');
    valid(prop(ref,'type')==='default'&&id&&!refs[kind]&&Object.hasOwn(stories,id)&&stories[id]!.kind===kind,'Unsupported first/even or invalid running story reference');refs[kind]=id;
  }
  if(items[i]&&isW(items[i]!,'type')){const type=items[i++]!;attrs(type,['w:val']);empty(type);valid(prop(type,'val')==='nextPage','Only next-page sections are supported');}
  const size=items[i++]!,margins=items[i++]!;
  valid(size&&margins&&isW(size,'pgSz')&&isW(margins,'pgMar'),'Explicit section page size and margins required');
  attrs(size,['w:w','w:h','w:orient']);empty(size);attrs(margins,['w:top','w:right','w:bottom','w:left','w:header','w:footer','w:gutter']);empty(margins);
  const width=numberProp(size,'w',1,31680),height=numberProp(size,'h',1,31680),orientation=prop(size,'orient');
  valid(orientation===undefined||orientation===(width>height?'landscape':'portrait'),'Inconsistent page orientation');
  valid(prop(margins,'gutter')===undefined||prop(margins,'gutter')==='0','Unsupported gutter');
  const extras:NonNullable<WriterSection['pageExtras']>={};
  for(const [name,key] of [['header','headerDistance'],['footer','footerDistance']] as const){
    if(prop(margins,name)!==undefined){const amount=numberProp(margins,name,0,31680)*EMU_PER_TWIP;if(amount||refs[name as 'header'|'footer'])extras[key]=amount;}
  }
  let pageNumberStart:number|undefined;
  if(items[i]&&isW(items[i]!,'pgNumType')){
    const num=items[i++]!;attrs(num,['w:start','w:fmt']);empty(num);valid(prop(num,'fmt')===undefined||prop(num,'fmt')==='decimal','Only decimal page numbering is supported');
    if(prop(num,'start')!==undefined)pageNumberStart=numberProp(num,'start',1,32767);
  }
  if(items[i]&&isW(items[i]!,'cols')){
    const cols=items[i++]!;attrs(cols,['w:num','w:space','w:equalWidth','w:sep']);empty(cols);
    valid((prop(cols,'num')===undefined||prop(cols,'num')==='1')&&(prop(cols,'equalWidth')===undefined||['1','true','on'].includes(prop(cols,'equalWidth')!))&&(prop(cols,'sep')===undefined||['0','false','off'].includes(prop(cols,'sep')!)),'Only a single column is supported');
    if(prop(cols,'space')!==undefined)extras.columnSpace=numberProp(cols,'space',0,31680)*EMU_PER_TWIP;
  }
  if(items[i]&&isW(items[i]!,'docGrid')){const grid=items[i++]!;attrs(grid,['w:type','w:linePitch']);empty(grid);valid(prop(grid,'type')===undefined||prop(grid,'type')==='default','Active document grid unsupported');if(prop(grid,'linePitch')!==undefined)extras.docGridLinePitch=numberProp(grid,'linePitch',0,31680)*EMU_PER_TWIP;}
  valid(i===items.length,'Unsupported section properties');
  return {page:{width:width*EMU_PER_TWIP,height:height*EMU_PER_TWIP,margins:{top:numberProp(margins,'top',0,31680)*EMU_PER_TWIP,right:numberProp(margins,'right',0,31680)*EMU_PER_TWIP,bottom:numberProp(margins,'bottom',0,31680)*EMU_PER_TWIP,left:numberProp(margins,'left',0,31680)*EMU_PER_TWIP}},...(Object.keys(extras).length?{pageExtras:extras}:{}),...(Object.keys(refs).length?{storyRefs:refs}:{}),...(pageNumberStart!==undefined?{pageNumberStart}:{})};
}
function simpleText(el:Element):string{
  attrs(el,['xml:space']);let text='';
  for(let n=el.firstChild;n;n=n.nextSibling){valid(n.nodeType===3,'Unsupported child inside w:t');text+=n.nodeValue??'';}
  const space=el.getAttributeNS(XML,'space');valid(!space||space==='preserve'||space==='default','Unknown whitespace mode');
  valid(!/^[ \t\r\n]|[ \t\r\n]$/.test(text)||space==='preserve','Ambiguous unpreserved whitespace');
  valid(!/[\r\n\t]/.test(text),'Raw line/tab control inside w:t');
  return text;
}
/** Only a single, uniformly formatted, tab-free paragraph is admitted.
 * PAGE is the only supported field; its cached result is never executed.
 * No other active markup, drawing, alternate story or nested relationship. */
function readPageStory(bytes:Uint8Array,kind:'header'|'footer',profile:DocxStyleProfile|undefined,options:ContainerOptions):WriterPageStory{
  const root=parseXml(bytes,options).documentElement!;requireElement(root,kind==='header'?'hdr':'ftr');attrs(root);
  const nodes=elements(root);valid(nodes.length===1&&isW(nodes[0]!,'p'),'Running story must contain one paragraph');
  const p=nodes[0]!;attrs(p);const items=elements(p);let i=0;
  const story:WriterPageStory={kind,text:''};
  if(items[i]&&isW(items[i]!,'pPr')){
    const pr=items[i++]!;attrs(pr);const seen=new Set<string>();
    for(const e of elements(pr)){
      const n=e.localName;valid(e.namespaceURI===W&&n&&!seen.has(n),'Duplicate or foreign story paragraph property');seen.add(n);
      if(n==='pStyle'){
        attrs(e,['w:val']);empty(e);const id=prop(e,'val'),expected=kind==='header'?'Header':'Footer';
        valid(id,'Missing running story style identifier');if(id!==profile?.defaultParagraph){valid(id===expected&&profile?.builtinStoryStyles?.has(expected),'Unsupported running story style');story.styleId=expected;}
      }else if(n==='jc'){
        attrs(e,['w:val']);empty(e);const v=prop(e,'val');valid(v==='left'||v==='center'||v==='right','Unsupported running story alignment');story.alignment=v==='left'?'start':v==='right'?'end':'center';
      }else unsupported('Unsupported running story paragraph property '+n);
    }
  }
  let uniform:TextStyle|undefined,haveStyle=false;
  const recordStyle=(style:TextStyle)=>{if(!haveStyle){uniform=style;haveStyle=true;}else valid(canonicalData(uniform??{})===canonicalData(style),'Mixed running story run formatting');};
  const readRun=(r:Element):{style:TextStyle;items:Element[]}=>{
    requireElement(r,'r');attrs(r);const nodes=elements(r);let style:TextStyle={};
    if(nodes[0]&&isW(nodes[0]!,'rPr'))style=readRunStyle(nodes.shift()!);
    return {style,items:nodes};
  };
  const plain=(nodes:Element[])=>nodes.map(n=>{requireElement(n,'t');return simpleText(n);}).join('');
  const field=(r:Element,type:string)=>{const data=readRun(r);valid(!Object.keys(data.style).length&&data.items.length===1&&isW(data.items[0]!,'fldChar'),'Unsupported field marker run');const e=data.items[0]!;attrs(e,['w:fldCharType']);empty(e);valid(prop(e,'fldCharType')===type,'Invalid PAGE field chain');};
  const isMarker=(r:Element|undefined,type:string)=>!!r&&isW(r,'r')&&elements(r).some(e=>isW(e,'fldChar')&&prop(e,'fldCharType')===type);
  const setField=(result:string)=>{valid(!story.pageNumber&&/^\d{0,10}$/.test(result),'Only one decimal PAGE field per story');story.pageNumber=true;};
  const append=(s:string)=>{if(story.pageNumber)story.after=(story.after??'')+s;else story.text+=s;};
  while(i<items.length){
    const e=items[i++]!;
    if(isW(e,'fldSimple')){
      attrs(e,['w:instr']);valid((prop(e,'instr')??'').trim().toUpperCase()==='PAGE','Unsupported running story field');
      let value='';for(const r of elements(e)){const d=readRun(r),t=plain(d.items);if(t)recordStyle(d.style);value+=t;}setField(value);continue;
    }
    if(isMarker(e,'begin')){
      field(e,'begin');let instruction='';
      while(i<items.length&&!isMarker(items[i],'separate')&&!isMarker(items[i],'end')){
        const d=readRun(items[i++]!);valid(!Object.keys(d.style).length&&d.items.length===1&&isW(d.items[0]!,'instrText'),'Unsupported PAGE instruction');
        const ins=d.items[0]!;attrs(ins,['xml:space']);let raw='';for(let n=ins.firstChild;n;n=n.nextSibling){valid(n.nodeType===3,'Nested instruction');raw+=n.nodeValue??'';}instruction+=raw;
      }
      valid(instruction.trim().toUpperCase()==='PAGE'&&instruction.length<=64,'Unsupported running story field instruction');
      let value='';
      if(isMarker(items[i],'separate')){field(items[i++]!,'separate');while(i<items.length&&!isMarker(items[i],'end')){const d=readRun(items[i++]!),t=plain(d.items);if(t)recordStyle(d.style);value+=t;}}
      valid(i<items.length&&isMarker(items[i],'end'),'Unterminated PAGE field');field(items[i++]!,'end');setField(value);continue;
    }
    const d=readRun(e),t=plain(d.items);if(t)recordStyle(d.style);append(t);
  }
  if(story.after==='')delete story.after;
  if(uniform&&Object.keys(uniform).length)story.runStyle=uniform;
  return story;
}
/** Rectangular table profile, including verified default-template table styles.
 * The editor projects the stored grid; native autofit is preserved, not simulated.
 */
const TABLE_EDGES=['top','left','bottom','right','insideH','insideV'] as const;
// Transitional CT_TblCellMar sequence: top, left, bottom, right.
const CELL_EDGES=['top','left','bottom','right'] as const;
const TABLE_DEFAULT_PADDING={top:38100,right:57150,bottom:38100,left:57150};
const BUILTIN_TABLE_PADDING={top:0,right:108*EMU_PER_TWIP,bottom:0,left:108*EMU_PER_TWIP};
const GRID_BORDER={width:6350,color:{kind:'rgba' as const,red:0,green:0,blue:0,alpha:255}};
const LOOK_BITS=[['firstRow',0x20],['lastRow',0x40],['firstColumn',0x80],['lastColumn',0x100],['noHBand',0x200],['noVBand',0x400]] as const;
function readLook(el:Element):number{
 attrs(el,['w:val',...LOOK_BITS.map(([name])=>'w:'+name)]);empty(el);
 const raw=prop(el,'val');valid(raw&&/^[a-fA-F0-9]{4}$/.test(raw),'Invalid table look');const mask=parseInt(raw,16);valid((mask&~0x7e0)===0,'Unknown table look flags');
 for(const [name,bit] of LOOK_BITS){const flag=prop(el,name);if(flag!==undefined)valid((flag==='0'||flag==='1')&&((flag==='1')===Boolean(mask&bit)),'Conflicting table look flags');}
 return mask;
}
function lookTag(mask:number):string{return tag('tblLook',{val:mask.toString(16).padStart(4,'0').toUpperCase(),...Object.fromEntries(LOOK_BITS.map(([name,bit])=>[name,mask&bit?1:0]))});}
function dxa(el:Element,min=0):number{
  attrs(el,['w:w','w:type']);empty(el);valid(prop(el,'type')==='dxa','Expected explicit twip table dimension');
  return numberProp(el,'w',min,31680)*EMU_PER_TWIP;
}
function readTable(el:Element,id:string,maxWidth:number,read:(p:Element)=>WriterBlock,charge:(n:number)=>void,profile?:DocxStyleProfile):Table{
  requireElement(el,'tbl');attrs(el);
  const children=elements(el);valid(children.length>=3&&isW(children[0]!,'tblPr')&&isW(children[1]!,'tblGrid'),'Explicit table properties and grid required');
  const props=children[0]!,grid=children[1]!;attrs(props);attrs(grid);
  const builtin=profile?.builtinTableStyles?.has('TableNormal')===true,native:NonNullable<Table['nativeTable']>={};
  let width:number|undefined,padding:Table['cellPadding'],border:Table['border'],fixed=false,borderSeen=false;
  const seen=new Set<string>();
  for(const item of elements(props)){
    const n=item.localName;valid(item.namespaceURI===W&&n!==null&&!seen.has(n),'Duplicate or unknown table property');seen.add(n);
    if(n==='tblStyle'){attrs(item,['w:val']);empty(item);const style=prop(item,'val');valid(builtin&&(style==='TableNormal'||style==='TableGrid')&&profile?.builtinTableStyles?.has(style),'Unsupported table style '+style);native.styleId=style;}
    else if(n==='tblW'){if(prop(item,'type')==='auto'){attrs(item,['w:w','w:type']);empty(item);valid(builtin&&prop(item,'w')==='0','Unsupported auto table width');native.autoWidth=true;}else width=dxa(item,1);}
    else if(n==='tblLayout'){attrs(item,['w:type']);empty(item);const kind=prop(item,'type');valid(kind==='fixed'||(builtin&&kind==='autofit'),'Unsupported table layout');fixed=kind==='fixed';}
    else if(n==='tblLook'){valid(builtin,'Table look requires a verified table style');native.look=readLook(item);}
    else if(n==='tblInd'){valid(builtin&&dxa(item)===0,'Unsupported table indentation');}
    else if(n==='tblCellMar'){
      attrs(item);const margins=elements(item);valid(margins.length===4,'Explicit cell margins required');
      const pad:Record<string,number>={};
      for(const child of margins){const edge=child.localName;valid(child.namespaceURI===W&&edge!==null&&(CELL_EDGES as readonly string[]).includes(edge)&&!Object.hasOwn(pad,edge),'Unsupported cell margin');pad[edge]=dxa(child);}
      padding={top:pad.top!,right:pad.right!,bottom:pad.bottom!,left:pad.left!};
    }else if(n==='tblBorders'){
      borderSeen=true;attrs(item);const edges=elements(item);valid(edges.length===6,'Uniform six-edge table border required');
      const used=new Set<string>();let spec:string|undefined;
      for(const edge of edges){
        const key=edge.localName;valid(edge.namespaceURI===W&&key!==null&&(TABLE_EDGES as readonly string[]).includes(key)&&!used.has(key),'Unsupported table border edge');used.add(key);
        attrs(edge,['w:val','w:sz','w:space','w:color']);empty(edge);
        const value=prop(edge,'val');
        if(value==='nil'){valid(builtin&&['sz','space','color'].every(k=>prop(edge,k)===undefined)&&(spec===undefined||spec==='nil'),'Mixed or unsupported nil border');spec='nil';continue;}
        valid(value==='single'&&spec!=='nil','Only uniform single-line borders supported');
        const size=numberProp(edge,'sz',2,96),color=prop(edge,'color')??'';
        valid(/^[0-9a-fA-F]{6}$/.test(color)&&(prop(edge,'space')===undefined||prop(edge,'space')==='0'),'Unsupported table border color or spacing');
        const code=size+':'+color.toUpperCase();valid(spec===undefined||spec===code,'Mixed table border styles unsupported');spec=code;
        border={width:Math.round(size*12700/8),color:{kind:'rgba',red:parseInt(color.slice(0,2),16),green:parseInt(color.slice(2,4),16),blue:parseInt(color.slice(4,6),16),alpha:255}};
      }
    }else unsupported('Unsupported table property '+n);
  }
  valid((fixed||builtin)&&(width!==undefined||native.autoWidth)&& (padding||builtin),'Explicit or verified table geometry required');
  if(!fixed)native.autoLayout=true;
  if(builtin&&padding)native.directPadding=true;
  if(builtin&&borderSeen)native.directBorder=true;
  if(!padding)padding={...BUILTIN_TABLE_PADDING};
  if(!borderSeen&&native.styleId==='TableGrid')border={...GRID_BORDER,color:{...GRID_BORDER.color}};
  const columns=elements(grid);valid(columns.length>=1&&columns.length<=64,'Table column limit exceeded');
  const columnWidths=columns.map(col=>{requireElement(col,'gridCol');attrs(col,['w:w']);empty(col);return numberProp(col,'w',1,31680)*EMU_PER_TWIP;});
  const gridWidth=columnWidths.reduce((a,b)=>a+b,0);
  valid((native.autoWidth||gridWidth===width)&&gridWidth<=maxWidth&&columnWidths.every(w=>w>padding!.left+padding!.right),'Invalid table grid or content width');
  const rows=children.slice(2);valid(rows.length>=1&&rows.length<=500&&rows.length*columns.length<=10000,'Table cell budget exceeded');charge(rows.length*columns.length);
  type Cell=Table['rows'][number]['cells'][number];
  type Vertical={anchor:Cell;span:number};
  let previous=new Map<number,Vertical>();
  const started:Cell[]=[];
  const resultRows:Table['rows']=rows.map((row,ri)=>{
    requireElement(row,'tr');attrs(row);const cells=elements(row);valid(cells.length>0&&cells.length<=columns.length,'Invalid table row cell count');
    const slots:Cell[]=[],active=new Map<number,Vertical>();let ci=0;
    for(const cell of cells){
      requireElement(cell,'tc');attrs(cell);const body=elements(cell);
      valid(body.length>=2&&isW(body[0]!,'tcPr'),'Explicit cell properties and paragraph required');
      const tcPr=body[0]!;attrs(tcPr);let cellWidth:number|undefined,span=1,vertical:'restart'|'continue'|undefined;
      const seen=new Set<string>();
      for(const item of elements(tcPr)){
        const n=item.localName;valid(item.namespaceURI===W&&n!==null&&!seen.has(n),'Duplicate or foreign cell property');seen.add(n);
        if(n==='tcW')cellWidth=dxa(item,1);
        else if(n==='gridSpan'){attrs(item,['w:val']);empty(item);span=numberProp(item,'val',1,64);}
        else if(n==='vMerge'){attrs(item,['w:val']);empty(item);const v=prop(item,'val')??'continue';valid(v==='restart'||v==='continue','Unsupported vertical merge mode');vertical=v;}
        else unsupported('Unsupported table cell property '+n);
      }
      valid(ci+span<=columns.length,'Cell span exceeds table grid');
      const width=columnWidths.slice(ci,ci+span).reduce((a,b)=>a+b,0);
      valid(cellWidth===width,'Cell width differs from its grid span');
      const slotId=(c:number)=>id+'-cell-'+ri+'-'+c;
      if(vertical==='continue'){
        const chain=previous.get(ci);valid(chain&&chain.span===span&&chain.anchor.merge,'Vertical merge has no matching preceding anchor');
        // A continuation may not hide text, drawings, paragraph styling, or
        // nested content. Only Word's empty continuation paragraph is admitted.
        valid(body.length===2&&isW(body[1]!,'p'),'Non-empty vertical merge continuation');attrs(body[1]!);empty(body[1]!);
        chain.anchor.merge.rows++;active.set(ci,chain);
        for(let c=ci;c<ci+span;c++)slots.push({id:slotId(c),blocks:[],coveredBy:chain.anchor.id});
      }else{
        const blocks=body.slice(1).map(p=>{requireElement(p,'p');const block=read(p);valid(block.type==='paragraph','Page break in table cell');
          for(const run of block.runs)if(run.type==='image')valid(run.width<=width-padding!.left-padding!.right,'Image exceeds cell width');return block;});
        const anchor:Cell={id:slotId(ci),blocks,...(span>1||vertical==='restart'?{merge:{rows:1,columns:span}}:{})};
        slots.push(anchor);for(let c=ci+1;c<ci+span;c++)slots.push({id:slotId(c),blocks:[],coveredBy:anchor.id});
        if(vertical==='restart'){active.set(ci,{anchor,span});started.push(anchor);}
      }
      ci+=span;
    }
    valid(ci===columns.length,'Non-rectangular table row');previous=active;
    return {id:id+'-row-'+ri,cells:slots};
  });
  valid(started.every(c=>c.merge!.rows>1),'Vertical merge restart without continuation');
  return {id,type:'table',columnWidths,cellPadding:padding,...(border?{border}:{}),...(Object.keys(native).length?{nativeTable:native}:{}),rows:resultRows};
}
function writeTable(table:DeepReadonly<Table>,maxWidth:number,write:(p:DeepReadonly<Paragraph>)=>string,charge:(n:number)=>void,profile?:DocxStyleProfile):string{
  const pad=table.cellPadding??TABLE_DEFAULT_PADDING;
  const exact=(n:number,min=0)=>{valid(n%EMU_PER_TWIP===0,'Table dimensions require whole twips');return twip(n,min,31680);};
  const widths=table.columnWidths.map(w=>exact(w,1)),sum=widths.reduce((a,b)=>a+b,0);
  valid(sum<=31680&&sum*EMU_PER_TWIP<=maxWidth&&table.columnWidths.every(w=>w>pad.left+pad.right),'Table exceeds page or cell content width');
  valid(table.rows.length<=500&&widths.length<=64&&table.rows.length*widths.length<=10000,'Table cell budget exceeded');charge(table.rows.length*widths.length);
  const native=table.nativeTable,builtin=profile?.builtinTableStyles?.has('TableNormal')===true;
  valid(!native||builtin&&(!native.styleId||profile?.builtinTableStyles?.has(native.styleId)),'Native table metadata requires its verified source styles');
  let properties=(native?.styleId?tag('tblStyle',{val:native.styleId}):'')+tag('tblW',native?.autoWidth?{w:0,type:'auto'}:{w:sum,type:'dxa'});
  const inherited=native?.styleId==='TableGrid'?GRID_BORDER:undefined;
  if(table.border&&(native?.directBorder||canonicalData(table.border)!==canonicalData(inherited??null))){
    const c=table.border.color;valid(c.kind==='rgba'&&c.alpha===255,'Only opaque RGB table borders supported');
    const size=Math.round(table.border.width*8/12700);valid(size>=2&&size<=96&&Math.round(size*12700/8)===table.border.width,'Table border size is not representable');
    const color=[c.red,c.green,c.blue].map(v=>v.toString(16).padStart(2,'0')).join('').toUpperCase();
    properties+='<w:tblBorders>'+TABLE_EDGES.map(edge=>tag(edge,{val:'single',sz:size,space:0,color})).join('')+'</w:tblBorders>';
  }else if(!table.border&&(inherited||native?.directBorder))properties+='<w:tblBorders>'+TABLE_EDGES.map(edge=>tag(edge,{val:'nil'})).join('')+'</w:tblBorders>';
  properties+=tag('tblLayout',{type:native?.autoLayout?'autofit':'fixed'});
  if(!builtin||native?.directPadding||canonicalData(pad)!==canonicalData(BUILTIN_TABLE_PADDING))properties+='<w:tblCellMar>'+CELL_EDGES.map(edge=>tag(edge,{w:exact(pad[edge]),type:'dxa'})).join('')+'</w:tblCellMar>';
  if(native?.look!==undefined)properties+=lookTag(native.look);
  const anchors=new Map<string,{row:number;column:number;cell:DeepReadonly<Table['rows'][number]['cells'][number]>}>();
  for(const [ri,row] of table.rows.entries())for(const [ci,cell] of row.cells.entries())if(cell.merge)anchors.set(cell.id,{row:ri,column:ci,cell});
  function cellProperties(i:number,span:number,vertical?:'restart'|'continue'):string{
    const width=widths.slice(i,i+span).reduce((a,b)=>a+b,0);
    return '<w:tcPr>'+tag('tcW',{w:width,type:'dxa'})+(span>1?tag('gridSpan',{val:span}):'')+(vertical?tag('vMerge',vertical==='restart'?{val:'restart'}:{}):'')+'</w:tcPr>';
  }
  return '<w:tbl><w:tblPr>'+properties+'</w:tblPr><w:tblGrid>'+widths.map(w=>tag('gridCol',{w})).join('')+'</w:tblGrid>'+table.rows.map((row,ri)=>'<w:tr>'+row.cells.map((cell,i)=>{
    if(cell.coveredBy){
      const anchor=anchors.get(cell.coveredBy);valid(anchor&&anchor.cell.merge,'Invalid covered table slot');
      if(i!==anchor.column||ri===anchor.row)return '';
      return '<w:tc>'+cellProperties(i,anchor.cell.merge.columns,'continue')+'<w:p/></w:tc>';
    }
    valid(cell.blocks.length>0&&cell.blocks.every(b=>b.type==='paragraph'),'Nested tables or non-paragraph cell blocks unsupported');
    const span=cell.merge?.columns??1,width=table.columnWidths.slice(i,i+span).reduce((a,b)=>a+b,0);
    for(const block of cell.blocks)if(block.type==='paragraph')for(const run of block.runs)if(run.type==='image')valid(run.width<=width-pad.left-pad.right,'Image exceeds cell width');
    return '<w:tc>'+cellProperties(i,span,cell.merge&&cell.merge.rows>1?'restart':undefined)+cell.blocks.map(b=>write(b as DeepReadonly<Paragraph>)).join('')+'</w:tc>';
  }).join('')+'</w:tr>').join('')+'</w:tbl>';
}

function readMain(xml:Document,documentId:string,title:string,profile:DocxStyleProfile|undefined,media:MediaGraph,stories:Record<string,WriterPageStory>):WriterDocument{
  const root=xml.documentElement!;requireElement(root,'document');attrs(root);
  for(let n=xml.firstChild;n;n=n.nextSibling)if(n!==root)valid(n.nodeType===3&&!(n.nodeValue??'').trim()||n.nodeType===7&&n.nodeName==='xml','Unsupported document-level node');
  const contents=elements(root);valid(contents.length===1&&isW(contents[0]!,'body'),'Single body required');
  const body=contents[0]!;attrs(body);const children=elements(body);
  valid(children.length>=2&&children.length<=MAX_BLOCKS+1,'Missing paragraphs or block limit exceeded');
  const sections:WriterSection[]=[];let runCount=0,textCount=0;
  let page:WriterSection['page'];
  let paragraphIndex=0,cellCount=0;
  const readParagraph=(p:Element,inCell=false,sectionEnd?:Element):WriterBlock=>{
    const bi=paragraphIndex++;valid(paragraphIndex<=MAX_BLOCKS,'Paragraph budget exceeded');requireElement(p,'p');attrs(p);const nodes=elements(p);
    let style:TextStyle|undefined,styleId=profile?.defaultParagraph;let i=0;
    if(nodes[0]&&isW(nodes[0],'pPr')){style=readParagraphStyle(nodes[0],profile?id=>{styleId=requireStyleReference(profile,id,'paragraph');}:undefined,sectionEnd);i++;}
    const runs:Inline[]=[];let pageBreak=false;
    for(;i<nodes.length;i++){
      const r=nodes[i]!;requireElement(r,'r');attrs(r);valid(++runCount<=MAX_RUNS,'Run limit exceeded');
      const items=elements(r);let j=0;let direct:TextStyle|undefined,runStyleId:string|undefined;
      if(items[0]&&isW(items[0],'rPr')){direct=readRunStyle(items[0],profile?id=>{runStyleId=requireStyleReference(profile,id,'character');}:undefined);j++;}
      if(items[j]&&isW(items[j]!,'drawing')){
        valid(items.length===1&&!direct&&!runStyleId,'An inline drawing must be in its own unformatted run');
        const picture=readInlinePicture(items[j]!,media,'r-'+bi+'-image-'+runs.length);
        valid(picture.width<=page.width-page.margins.left-page.margins.right&&picture.height<=page.height-page.margins.top-page.margins.bottom,'Image exceeds usable page');runs.push(picture);continue;
      }
      let text='';
      for(;j<items.length;j++){
        const child=items[j]!;
        if(isW(child,'t'))text+=simpleText(child);
        else if(isW(child,'tab')){attrs(child);empty(child);text+='\t';}
        else if(isW(child,'br')){
          attrs(child,['w:type']);empty(child);const type=prop(child,'type');
          if(type==='page'){
            valid(!inCell&&!sectionEnd&&!style&&!direct&&!runStyleId&&(!profile||!styleId||styleId===profile.defaultParagraph)&&nodes.length===1&&items.length===1,'Page break must occupy its own unformatted paragraph');pageBreak=true;
          }else{valid(type===undefined||type==='textWrapping','Unsupported break type');text+='\n';}
        }else unsupported('Unsupported run content '+child.localName);
      }
      textCount+=text.length;valid(textCount<=MAX_TEXT,'Text budget exceeded');
      if(!pageBreak)runs.push({id:'r-'+bi+'-'+runs.length,type:'text',text,...(runStyleId?{styleId:runStyleId}:{}),...(direct&&Object.keys(direct).length?{style:direct}:{})});
    }
    if(pageBreak)return {id:'break-'+bi,type:'page-break'};
    return {id:'p-'+bi,type:'paragraph',runs:runs.length?runs:[{id:'r-'+bi+'-0',type:'text',text:''}],...(styleId?{styleId}:{}),...(style&&Object.keys(style).length?{style}:{})};
  };
  // Section properties belong to the paragraph ending the preceding
  // section. Resolve geometry before importing its paragraphs and tables.
  const ends:{index:number;properties:Element;onParagraph:boolean}[]=[];
  for(let i=0;i<children.length-1;i++){
    const p=children[i]!;if(!isW(p,'p'))continue;
    const pr=elements(p)[0];if(!pr||!isW(pr,'pPr'))continue;
    const props=elements(pr),nested=props.filter(e=>isW(e,'sectPr'));
    if(nested.length){valid(nested.length===1&&props.at(-1)===nested[0],'Section properties must end paragraph properties');ends.push({index:i,properties:nested[0]!,onParagraph:true});}
  }
  ends.push({index:children.length-2,properties:children.at(-1)!,onParagraph:false});
  valid(ends.length<=128,'Section budget exceeded');
  let start=0;
  for(const end of ends){
    valid(end.index>=start,'Empty section is not supported');
    const geometry=readPage(end.properties,stories);page=geometry.page;const blocks:WriterBlock[]=[];
    for(let bi=start;bi<=end.index;bi++){
      const block=children[bi]!;
      if(isW(block,'tbl')){
        valid(!profile||profile.tableDefaultsSafe,'Default table style requires a later adapter');
        blocks.push(readTable(block,'table-'+bi,page.width-page.margins.left-page.margins.right,
          p=>readParagraph(p,true),n=>{cellCount+=n;valid(cellCount<=10000,'Document table-cell budget exceeded');},profile));
      }else blocks.push(readParagraph(block,false,end.onParagraph&&bi===end.index?end.properties:undefined));
    }
    sections.push({id:'section-'+(sections.length+1),...geometry,blocks});start=end.index+1;
  }
  const result=createDocument('writer',documentId,title);result.defaultTextStyle=profile?.defaults??{};result.styles=profile?.styles??{};result.resources=media.resources;if(profile?.defaultParagraph)result.defaultParagraphStyleId=profile.defaultParagraph;result.sections=sections;if(Object.keys(stories).length)result.pageStories=stories;if(profile?.numbering)result.numbering=profile.numbering.numbering;
  const checked=validateDocument(result) as WriterDocument;buildWriterNumberingMarkers(checked);return checked;
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
function writeRunStyle(style:DeepReadonly<TextStyle>,styleId?:string):string{
  valid(Object.keys(style).every(k=>CHAR_KEYS.includes(k)),'Unsupported character style');
  let r=styleId?tag('rStyle',{val:styleId}):'';
  if(style.fontFamily!==undefined||style.fontAscii!==undefined||style.fontHighAnsi!==undefined||style.fontEastAsia!==undefined||style.fontComplex!==undefined){
    const mapped=style.fontAscii!==undefined||style.fontHighAnsi!==undefined||style.fontEastAsia!==undefined||style.fontComplex!==undefined;const fonts:Record<string,string>={};for(const [key,value] of [['ascii',mapped?style.fontAscii:style.fontFamily],['hAnsi',mapped?style.fontHighAnsi:style.fontFamily],['eastAsia',mapped?style.fontEastAsia:style.fontFamily],['cs',mapped?style.fontComplex:style.fontFamily]] as const)if(value!==undefined)fonts[key]=value;r+=tag('rFonts',fonts);
  }
  for(const [key,xml] of [['bold','b'],['italic','i']] as const)if(style[key]!==undefined)r+=tag(xml,{val:style[key]?1:0});
  if(style.color!==undefined){
    valid(style.color.kind==='rgba'&&style.color.alpha===255,'Theme or translucent text color unsupported');
    const v=[style.color.red,style.color.green,style.color.blue].map(n=>n.toString(16).padStart(2,'0')).join('').toUpperCase();r+=tag('color',{val:v});
  }
  if(style.characterSpacing!==undefined)r+=tag('spacing',{val:twip(style.characterSpacing,-31680,31680)});
  if(style.kerningThreshold!==undefined){valid(style.kerningThreshold%6350===0,'Kerning threshold must be a half-point value');r+=tag('kern',{val:style.kerningThreshold/6350});}
  if(style.fontSize!==undefined){const n=Math.round(style.fontSize/6350);valid(n>=1&&n<=3276,'Font size outside DOCX limits');r+=tag('sz',{val:n});}
  if(style.complexFontSize!==undefined){const n=Math.round(style.complexFontSize/6350);valid(n>=1&&n<=3276,'Complex font size outside DOCX limits');r+=tag('szCs',{val:n});}
  if(style.underline!==undefined)r+=tag('u',{val:style.underline?'single':'none'});
  if(style.runDirection!==undefined)r+=tag('rtl',{val:style.runDirection==='rtl'?1:0});
  const langs:Record<string,string>={};for(const [key,property] of [['val','language'],['eastAsia','eastAsiaLanguage'],['bidi','bidiLanguage']] as const)if(style[property]!==undefined)langs[key]=style[property]!;
  if(Object.keys(langs).length)r+=tag('lang',langs);
  return r?'<w:rPr>'+r+'</w:rPr>':'';
}
function writeParagraphStyle(style:DeepReadonly<TextStyle>,styleId?:string,sectionEnd=''):string{
  valid(Object.keys(style).every(k=>PARA_KEYS.includes(k)),'Unsupported paragraph style');
  let p=styleId?tag('pStyle',{val:styleId}):'';for(const key of ['keepNext','keepLines','widowControl'] as const)if(style[key]!==undefined)p+=tag(key,{val:style[key]?1:0});if(style.paragraphDirection!==undefined)p+=tag('bidi',{val:style.paragraphDirection==='rtl'?1:0});const spacing:Record<string,string|number>={};
  if(style.spaceBefore!==undefined)spacing.before=twip(style.spaceBefore,0,31680);
  if(style.spaceAfter!==undefined)spacing.after=twip(style.spaceAfter,0,31680);
  if(style.lineHeight!==undefined){spacing.line=twip(style.lineHeight,1,31680);spacing.lineRule='exact';}
  if(style.lineMultiple!==undefined){spacing.line=style.lineMultiple;spacing.lineRule='auto';}
  if(style.minLineHeight!==undefined){spacing.line=twip(style.minLineHeight,1,31680);spacing.lineRule='atLeast';}
  if(style.writerList)p+='<w:numPr>'+tag('ilvl',{val:style.writerList.level})+tag('numId',{val:style.writerList.numId})+'</w:numPr>';
  if(style.paragraphBottomBorder!==undefined){
    const b=style.paragraphBottomBorder;
    if(b===null)p+='<w:pBdr>'+tag('bottom',{val:'nil'})+'</w:pBdr>';
    else{valid(b.color.kind==='rgba'&&b.color.alpha===255,'Paragraph border requires opaque RGB');const size=Math.round(b.width*8/12700);valid(size>=2&&size<=96&&Math.round(size*12700/8)===b.width&&b.space%12700===0,'Paragraph border precision');const rgb=[b.color.red,b.color.green,b.color.blue].map(n=>n.toString(16).padStart(2,'0')).join('').toUpperCase();p+='<w:pBdr>'+tag('bottom',{val:'single',sz:size,space:b.space/12700,color:rgb})+'</w:pBdr>';}
  }
  if(Object.keys(spacing).length)p+=tag('spacing',spacing);
  const indentation:Record<string,number>={};
  if(style.indent!==undefined)indentation.left=twip(style.indent,-31680,31680);
  if(style.rightIndent!==undefined)indentation.right=twip(style.rightIndent,0,31680);
  if(style.firstLineIndent!==undefined)indentation[style.firstLineIndent<0?'hanging':'firstLine']=twip(Math.abs(style.firstLineIndent),0,31680);
  if(Object.keys(indentation).length)p+=tag('ind',indentation);
  if(style.contextualSpacing!==undefined)p+=tag('contextualSpacing',{val:style.contextualSpacing?1:0});
  if(style.alignment!==undefined)p+=tag('jc',{val:({start:'left',center:'center',end:'right',justify:'both'})[style.alignment]});
  if(style.outlineLevel!==undefined)p+=tag('outlineLvl',{val:style.outlineLevel});
  return p||sectionEnd?'<w:pPr>'+p+sectionEnd+'</w:pPr>':'';
}
function writeText(text:string):string{
  valid(!text.includes('\r'),'CR must be represented explicitly by a supported line break');
  if(!text)return '<w:t xml:space="preserve"></w:t>';
  return text.split(/(\n|\t)/).filter(value=>value!=='').map(value=>value==='\n'?'<w:br/>':value==='\t'?'<w:tab/>':'<w:t xml:space="preserve">'+escapeXml(value)+'</w:t>').join('');
}
function canonicalData(value:unknown):string{
  if(value===null||typeof value!=='object')return JSON.stringify(value);
  if(Array.isArray(value))return '['+value.map(canonicalData).join(',')+']';
  return '{'+Object.entries(value).sort(([a],[b])=>a<b?-1:a>b?1:0).map(([k,v])=>JSON.stringify(k)+':'+canonicalData(v)).join(',')+'}';
}
function mainXml(candidate:unknown,profile:DocxStyleProfile|undefined,media:MediaPlan):{model:WriterDocument;xml:string}{
  const checked=validateDocument(candidate);valid(checked.kind==='writer','Expected Writer model');buildWriterNumberingMarkers(checked);
  valid(checked.sections.length<=128,'Section budget exceeded');
  if(profile){valid(checked.defaultParagraphStyleId===profile.defaultParagraph&&canonicalData(checked.styles)===canonicalData(profile.styles)&&canonicalData(checked.defaultTextStyle)===canonicalData(profile.defaults),'Source style definitions and defaults are immutable');}
  else{valid(!Object.keys(checked.styles).length,'New DOCX named style definitions are not supported');valid(Object.keys(checked.defaultTextStyle).every(k=>CHAR_KEYS.includes(k)),'Unsupported default text style');}
  valid(checked.sections.reduce((n,s)=>n+s.blocks.length,0)<=MAX_BLOCKS,'Block budget exceeded');
  let section=checked.sections[0]!;
  let totalRuns=0,totalText=0,paragraphCount=0,cellCount=0,imageCount=0;let body='';
  const writeParagraph=(block:DeepReadonly<Paragraph>,inCell=false,sectionEnd=''):string=>{
    valid(++paragraphCount<=MAX_BLOCKS,'Paragraph budget exceeded');let body='';
    const paragraphStyle=block.styleId??profile?.defaultParagraph;
    if(paragraphStyle)requireStyleReference(profile,paragraphStyle,'paragraph');
    body+='<w:p>'+writeParagraphStyle(block.style??{},paragraphStyle,sectionEnd);
    for(const inline of block.runs.length?block.runs:[{id:'empty',type:'text' as const,text:''}]){
      valid(++totalRuns<=MAX_RUNS,'Run budget exceeded');
      if(inline.type==='image'){
        valid(inline.width<=section.page.width-section.page.margins.left-section.page.margins.right&&inline.height<=section.page.height-section.page.margins.top-section.page.margins.bottom,'Inline image exceeds page bounds');
        body+='<w:r>'+writeInlinePicture(inline,media,++imageCount)+'</w:r>';
      }else if(inline.type==='text'){
        if(inline.styleId)requireStyleReference(profile,inline.styleId,'character');totalText+=inline.text.length;valid(totalText<=MAX_TEXT,'Text budget exceeded');
        body+='<w:r>'+writeRunStyle(profile?inline.style??{}:{...checked.defaultTextStyle,...inline.style},inline.styleId)+writeText(inline.text)+'</w:r>';
      }else body+='<w:r>'+writeRunStyle(profile?{}:checked.defaultTextStyle)+(inline.type==='tab'?'<w:tab/>':'<w:br/>')+'</w:r>';
    }
    return body+'</w:p>';
  };
  const sectionXml=(s:DeepReadonly<WriterSection>):string=>{
    const p=s.page,width=twip(p.width,1,31680),height=twip(p.height,1,31680),m=p.margins,extra=s.pageExtras;
    const refs=(['header','footer'] as const).map(kind=>s.storyRefs?.[kind]?'<w:'+kind+'Reference w:type="default" r:id="'+escapeXml(s.storyRefs[kind]!)+'"/>':'').join('');
    return '<w:sectPr>'+refs+tag('pgSz',{w:width,h:height,...(width>height?{orient:'landscape'}:{})})+
      tag('pgMar',{top:twip(m.top,0,31680),right:twip(m.right,0,31680),bottom:twip(m.bottom,0,31680),left:twip(m.left,0,31680),header:twip(extra?.headerDistance??0,0,31680),footer:twip(extra?.footerDistance??0,0,31680),gutter:0})+
      (s.pageNumberStart===undefined?'':tag('pgNumType',{start:s.pageNumberStart,fmt:'decimal'}))+
      (extra?.columnSpace===undefined?'':tag('cols',{num:1,space:twip(extra.columnSpace,0,31680)}))+(extra?.docGridLinePitch===undefined?'':tag('docGrid',{type:'default',linePitch:twip(extra.docGridLinePitch,0,31680)}))+'</w:sectPr>';
  };
  for(let si=0;si<checked.sections.length;si++){
    section=checked.sections[si]!;const final=si===checked.sections.length-1;
    valid(final||section.blocks.at(-1)?.type==='paragraph','Nonfinal section must end in a paragraph');
    const maxWidth=section.page.width-section.page.margins.left-section.page.margins.right;
    const end=sectionXml(section);
    for(let bi=0;bi<section.blocks.length;bi++){
      const block=section.blocks[bi]!;
      if(block.type==='page-break')body+='<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
      else if(block.type==='table'){valid(!profile||profile.tableDefaultsSafe,'Default table style requires a later adapter');body+=writeTable(block,maxWidth,p=>writeParagraph(p,true),n=>{cellCount+=n;valid(cellCount<=10000,'Document table-cell budget exceeded');},profile);}
      else body+=writeParagraph(block,false,!final&&bi===section.blocks.length-1?end:'');
    }
    if(final)body+=end;
  }
  return {model:checked,xml:DECL+'<w:document xmlns:w="'+W+'"'+(checked.sections.some(s=>s.storyRefs)?' xmlns:r="'+REL+'"':'')+'><w:body>'+body+'</w:body></w:document>'};
}
function storyXml(story:DeepReadonly<WriterPageStory>,profile:DocxStyleProfile|undefined,defaults:DeepReadonly<TextStyle>):string{
  if(story.styleId)valid(profile?.builtinStoryStyles?.has(story.styleId),'Running story requires verified source style');
  const p=(story.styleId?tag('pStyle',{val:story.styleId}):'')+(story.alignment?tag('jc',{val:({start:'left',center:'center',end:'right'})[story.alignment]}):'');
  const style=writeRunStyle(profile?story.runStyle??{}:{...defaults,...story.runStyle});
  const run=(text:string)=>'<w:r>'+style+writeText(text)+'</w:r>';
  const root=story.kind==='header'?'hdr':'ftr';
  return DECL+'<w:'+root+' xmlns:w="'+W+'"><w:p>'+(p?'<w:pPr>'+p+'</w:pPr>':'')+run(story.text)+(story.pageNumber?'<w:fldSimple w:instr=" PAGE ">'+run('1')+'</w:fldSimple>'+(story.after?run(story.after):''):'')+'</w:p></w:'+root+'>';
}
type StoryPlan={changes:Map<string,Uint8Array>;additions:{id:string;path:string;kind:'header'|'footer';bytes:Uint8Array}[];canonical:Record<string,string>};
function planStories(model:DeepReadonly<WriterDocument>,main:string,profile:DocxStyleProfile|undefined,source?:OpcPackage,baseline?:ReadonlyMap<string,DocxStorySource>,prior?:Readonly<Record<string,string>>):StoryPlan{
  const result:StoryPlan={changes:new Map(),additions:[],canonical:Object.create(null)};
  // Previously admitted relationships remain source-owned. Undo after a save
  // may omit a newly created, now unreferenced story: preserve the part.
  for(const [id,story]of Object.entries(model.pageStories??{})){
    const xml=storyXml(story,profile,model.defaultTextStyle);result.canonical[id]=xml;
    const old=baseline?.get(id);
    if(old){valid(old.kind===story.kind,'Cannot change source running story kind');if(prior?.[id]!==xml)result.changes.set(old.path,encoder.encode(xml));continue;}
    valid(!source?.relationships.get(main)?.some(r=>r.id===id),'Running story relationship collision');
    const path=main.slice(0,main.lastIndexOf('/')+1)+story.kind+'-oex-'+id+'.xml';
    valid(![...source?.parts.keys()??[]].some(p=>p.toLowerCase()===path.toLowerCase())&&!result.additions.some(p=>p.path.toLowerCase()===path.toLowerCase()),'Running story part collision');
    result.additions.push({id,path,kind:story.kind,bytes:encoder.encode(xml)});
  }
  // Keep orphaned source stories byte-exact (relationships may be inherited
  // or restored by undo); don't silently remove parts from existing files.
  for(const id of baseline?.keys()??[]){if(!Object.hasOwn(result.canonical,id))result.canonical[id]=prior![id]!;}
  return result;
}
function addStoryParts(parts:Map<string,Uint8Array>,main:string,plan:StoryPlan,options:ContainerOptions):void{
  for(const [path,bytes]of plan.changes)parts.set(path,bytes);
  if(!plan.additions.length)return;
  const i=main.lastIndexOf('/'),expected=main.slice(0,i+1)+'_rels/'+main.slice(i+1)+'.rels',actual=[...parts.keys()].find(p=>p.toLowerCase()===expected.toLowerCase())??expected;
  const rels=parseXml(parts.get(actual)??'<Relationships xmlns="'+PKG+'"/>',options),types=parseXml(parts.get('[Content_Types].xml')!,options);
  const relRoot=rels.documentElement!,typeRoot=types.documentElement!,ids=new Set(elements(relRoot).map(e=>e.getAttribute('Id')));
  for(const item of plan.additions){
    valid(!ids.has(item.id)&&!parts.has(item.path),'Running story part or relationship collision');ids.add(item.id);
    const rel=rels.createElementNS(PKG,'Relationship');rel.setAttribute('Id',item.id);rel.setAttribute('Type',REL+'/'+item.kind);rel.setAttribute('Target','/'+item.path);relRoot.appendChild(rel);
    const override=types.createElementNS(CT,'Override');override.setAttribute('PartName','/'+item.path);override.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.wordprocessingml.'+item.kind+'+xml');typeRoot.appendChild(override);
    parts.set(item.path,item.bytes);
  }
  parts.set(actual,serializeXml(rels,options));parts.set('[Content_Types].xml',serializeXml(types,options));
}
export interface WriterDocxSource{readonly document:DeepReadonly<WriterDocument>;readonly mainPart:string}
const sources=new WeakMap<WriterDocxSource,{pkg:OpcPackage;canonical:string;documentId:string;profile?:DocxStyleProfile;media:MediaGraph;storySources:ReadonlyMap<string,DocxStorySource>;storyCanonical:Readonly<Record<string,string>>}>();
export function openWriterDocx(bytes:Uint8Array,documentId:string,title:string,options:ContainerOptions={}):WriterDocxSource{
  const pkg=openOpc(bytes,options),{part,profile,stories:storySources}=readStyleProfile(pkg,options,readRunStyle,readParagraphStyle),media=loadMedia(pkg,part);
  const stories:Record<string,WriterPageStory>=Object.create(null);
  for(const [id,item]of storySources)stories[id]=readPageStory(pkg.parts.get(item.path)!,item.kind,profile,options);
  const document=readMain(parseXml(pkg.parts.get(part)!,options),documentId,title,profile,media,stories);
  const source=Object.freeze({document:freezeData(document),mainPart:part}),storyCanonical=Object.fromEntries(Object.entries(stories).map(([id,s])=>[id,storyXml(s,profile,document.defaultTextStyle)]));
  sources.set(source,{pkg,canonical:mainXml(document,profile,planMedia(document.resources,part,pkg,media)).xml,documentId,profile,media,storySources,storyCanonical});
  return source;
}
const encoder=new TextEncoder();
export function writeWriterDocx(candidate:unknown,options:ContainerOptions={}):Uint8Array{
  const model=validateDocument(candidate);valid(model.kind==='writer','Expected Writer model');const part='word/document.xml',media=planMedia(model.resources,part),{xml}=mainXml(model,undefined,media);
  const parts=new Map([
    ['[Content_Types].xml',encoder.encode(DECL+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/'+part+'" ContentType="'+MAIN_TYPE+'"/></Types>')],
    ['_rels/.rels',encoder.encode(DECL+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="'+ROOT_REL+'" Target="'+part+'"/></Relationships>')],
    [part,encoder.encode(xml)]
  ]);
  const stories=planStories(model,part,undefined);addStoryParts(parts,part,stories,options);
  addMediaParts(parts,part,media,options);addNumberingParts(parts,part,undefined,model.numbering,options);const bytes=writeZip(parts,options);
  const checked=openWriterDocx(bytes,model.documentId,model.title,options),saved=sources.get(checked)!;
  valid(canonicalData(checked.document.numbering??null)===canonicalData(model.numbering??null)&&saved.canonical===xml&&canonicalData(saved.storyCanonical)===canonicalData(stories.canonical)&&canonicalData(checked.document.resources)===canonicalData(model.resources),'Semantic output verification failed');
  return bytes;
}
/** No-edit returns original archive bytes. Edited copy changes only the supported main part.
 * No IO occurs here: desktop save planning/atomic file commit belongs to a later M2 slice.
 */
export function saveWriterDocx(source:WriterDocxSource,candidate:unknown,options:ContainerOptions={}):Uint8Array{
  const saved=sources.get(source);assert(saved,'invalid','Unknown DOCX source capability');
  const model=validateDocument(candidate);valid(model.kind==='writer','Expected Writer model');assert(model.documentId===saved.documentId,'conflict','Document identity mismatch');
  const media=planMedia(model.resources,source.mainPart,saved.pkg,saved.media),{xml}=mainXml(model,saved.profile,media);
  const changed=numberingChanged(saved.profile?.numbering,model.numbering),stories=planStories(model,source.mainPart,saved.profile,saved.pkg,saved.storySources,saved.storyCanonical);
  if(xml===saved.canonical&&!media.additions.length&&!changed&&!stories.changes.size&&!stories.additions.length)return writeOpc(saved.pkg,new Map(),options);
  let bytes:Uint8Array;
  if(changed||stories.changes.size||stories.additions.length){writeOpc(saved.pkg,new Map(),options);const parts=new Map(saved.pkg.parts);parts.set(source.mainPart,encoder.encode(xml));addStoryParts(parts,source.mainPart,stories,options);addMediaParts(parts,source.mainPart,media,options);addNumberingParts(parts,source.mainPart,saved.profile?.numbering,model.numbering,options);bytes=writeZip(parts,options);}
  else bytes=saveMediaPackage(saved.pkg,source.mainPart,xml,media,options);
  const checked=openWriterDocx(bytes,model.documentId,model.title,options);
  valid(canonicalData(checked.document.numbering??null)===canonicalData(model.numbering??null)&&sources.get(checked)?.canonical===xml&&canonicalData(sources.get(checked)?.storyCanonical)===canonicalData(stories.canonical)&&canonicalData(checked.document.resources)===canonicalData({...saved.media.resources,...model.resources})&&canonicalData(checked.document.styles)===canonicalData(model.styles)&&canonicalData(checked.document.defaultTextStyle)===canonicalData(saved.profile?.defaults??{}),'Edited DOCX verification failed');
  return bytes;
}
