// SPDX-License-Identifier: MIT
/** M2 narrow WordprocessingML semantic profile, using the existing bounded OPC layer.
 * Only supported named styles and preserved themes; nested/merged tables, images, tracked changes, external links and unsupported
 * section defaults may be rewritten. Unsupported documents stay on the inspection path.
 */
import type { Element, Document } from '@xmldom/xmldom';
import { OfficeError, assert, EMU_PER_TWIP } from '@office-ex/contracts';
import { createDocument, validateDocument, freezeData, isDocumentLanguageTag, type DeepReadonly, type WriterDocument, type WriterBlock, type Paragraph, type TextRun, type Inline, type TextStyle, type WriterSection, type NamedStyle, type Table } from '@office-ex/core';
import { openOpc, writeOpc, type OpcPackage } from './opc.js';
import { parseXml } from './xml.js';
import { writeZip } from './zip.js';
import type { ContainerOptions } from './limits.js';
import { loadMedia, planMedia, readInlinePicture, writeInlinePicture, addMediaParts, saveMediaPackage, type MediaGraph, type MediaPlan } from './writer-docx-images.js';
import { readStyleProfile, type DocxStyleProfile, type ThemeProfile, requireStyleReference } from './writer-docx-styles.js';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML='http://www.w3.org/XML/1998/namespace';
const XMLNS='http://www.w3.org/2000/xmlns/';
const ROOT_REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const MAIN_TYPE='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const DECL='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const CHAR_KEYS=['fontFamily','fontSize','bold','italic','underline','color','language','eastAsiaLanguage','bidiLanguage','runDirection'];
const PARA_KEYS=['alignment','indent','spaceBefore','spaceAfter','lineHeight','paragraphDirection'];
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
      case 'sz':attrs(child,['w:val']);empty(child);style.fontSize=numberProp(child,'val',1,3276)*6350;break;
      case 'color':{attrs(child,theme?['w:val','w:themeColor']:['w:val']);empty(child);const key=prop(child,'themeColor');const color=key!==undefined?(theme?.colors[key]??''):prop(child,'val')??'';valid(/^[0-9a-fA-F]{6}$/.test(color),'Unsupported color');style.color={kind:'rgba',red:parseInt(color.slice(0,2),16),green:parseInt(color.slice(2,4),16),blue:parseInt(color.slice(4,6),16),alpha:255};break;}
      case 'rFonts':{attrs(child,theme?['w:ascii','w:hAnsi','w:eastAsia','w:cs','w:asciiTheme','w:hAnsiTheme','w:eastAsiaTheme','w:cstheme']:['w:ascii','w:hAnsi','w:eastAsia','w:cs']);empty(child);const fontKeys=[['ascii','asciiTheme'],['hAnsi','hAnsiTheme'],['eastAsia','eastAsiaTheme'],['cs','cstheme']] as const;const fonts=fontKeys.map(([literal,themed])=>{const ref=prop(child,themed);return ref!==undefined?theme?.fonts[ref]:prop(child,literal);});const font=fonts[0];valid(font&&fonts.every(f=>f===font),'Mixed or inherited font mapping');style.fontFamily=font;break;}
      default:unsupported('Unsupported run property '+child.localName);
    }
  }
  return style;
}
export function readParagraphStyle(el:Element,reference?:(id:string)=>void,sectionEnd?:Element):TextStyle{
  requireElement(el,'pPr');attrs(el);const style:TextStyle={},seen=new Set<string>();
  for(const child of elements(el)){
    valid(child.namespaceURI===W&&child.localName!==null&&!seen.has(child.localName),'Duplicate or foreign paragraph property');seen.add(child.localName);
    switch(child.localName){
      case 'bidi':style.paragraphDirection=toggled(child)?'rtl':'ltr';break;
      case 'sectPr':{valid(child===sectionEnd,'Unexpected section properties inside paragraph');break;}
      case 'pStyle':{valid(reference,'Unexpected paragraph style reference');attrs(child,['w:val']);empty(child);const id=prop(child,'val');valid(id,'Missing paragraph style reference');reference(id);break;}
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
function readPage(el:Element):Pick<WriterSection,'page'|'pageExtras'>{
  requireElement(el,'sectPr');attrs(el);const items=elements(el);let i=0;
  if(items[i]&&isW(items[i]!,'type')){const type=items[i++]!;attrs(type,['w:val']);empty(type);valid(prop(type,'val')==='nextPage','Only next-page sections are supported');}
  const size=items[i++]!,margins=items[i++]!;
  valid(size&&margins&&isW(size,'pgSz')&&isW(margins,'pgMar'),'Explicit section page size and margins required');
  attrs(size,['w:w','w:h','w:orient']);empty(size);attrs(margins,['w:top','w:right','w:bottom','w:left','w:header','w:footer','w:gutter']);empty(margins);
  const width=numberProp(size,'w',1,31680),height=numberProp(size,'h',1,31680),orientation=prop(size,'orient');
  valid(orientation===undefined||orientation===(width>height?'landscape':'portrait'),'Inconsistent page orientation');
  valid(prop(margins,'gutter')===undefined||prop(margins,'gutter')==='0','Unsupported gutter');
  const extras:NonNullable<WriterSection['pageExtras']>={};
  for(const [name,key] of [['header','headerDistance'],['footer','footerDistance']] as const){
    if(prop(margins,name)!==undefined){const amount=numberProp(margins,name,0,31680)*EMU_PER_TWIP;if(amount)extras[key]=amount;}
  }
  if(items[i]&&isW(items[i]!,'cols')){
    const cols=items[i++]!;attrs(cols,['w:num','w:space','w:equalWidth','w:sep']);empty(cols);
    valid((prop(cols,'num')===undefined||prop(cols,'num')==='1')&&(prop(cols,'equalWidth')===undefined||['1','true','on'].includes(prop(cols,'equalWidth')!))&&(prop(cols,'sep')===undefined||['0','false','off'].includes(prop(cols,'sep')!)),'Only a single column is supported');
    if(prop(cols,'space')!==undefined)extras.columnSpace=numberProp(cols,'space',0,31680)*EMU_PER_TWIP;
  }
  valid(i===items.length,'Unsupported section properties');
  return {page:{width:width*EMU_PER_TWIP,height:height*EMU_PER_TWIP,margins:{top:numberProp(margins,'top',0,31680)*EMU_PER_TWIP,right:numberProp(margins,'right',0,31680)*EMU_PER_TWIP,bottom:numberProp(margins,'bottom',0,31680)*EMU_PER_TWIP,left:numberProp(margins,'left',0,31680)*EMU_PER_TWIP}},...(Object.keys(extras).length?{pageExtras:extras}:{})};
}
function simpleText(el:Element):string{
  attrs(el,['xml:space']);let text='';
  for(let n=el.firstChild;n;n=n.nextSibling){valid(n.nodeType===3,'Unsupported child inside w:t');text+=n.nodeValue??'';}
  const space=el.getAttributeNS(XML,'space');valid(!space||space==='preserve'||space==='default','Unknown whitespace mode');
  valid(!/^[ \t\r\n]|[ \t\r\n]$/.test(text)||space==='preserve','Ambiguous unpreserved whitespace');
  valid(!/[\r\n\t]/.test(text),'Raw line/tab control inside w:t');
  return text;
}
/** Fixed-width rectangular table profile. Each cell contains only paragraphs.
 * Require explicit geometry and margins rather than invent Word defaults.
 */
const TABLE_EDGES=['top','left','bottom','right','insideH','insideV'] as const;
// Transitional CT_TblCellMar sequence: top, left, bottom, right.
const CELL_EDGES=['top','left','bottom','right'] as const;
const TABLE_DEFAULT_PADDING={top:38100,right:57150,bottom:38100,left:57150};
function dxa(el:Element,min=0):number{
  attrs(el,['w:w','w:type']);empty(el);valid(prop(el,'type')==='dxa','Expected explicit twip table dimension');
  return numberProp(el,'w',min,31680)*EMU_PER_TWIP;
}
function readTable(el:Element,id:string,maxWidth:number,read:(p:Element)=>WriterBlock,charge:(n:number)=>void):Table{
  requireElement(el,'tbl');attrs(el);
  const children=elements(el);valid(children.length>=3&&isW(children[0]!,'tblPr')&&isW(children[1]!,'tblGrid'),'Explicit table properties and grid required');
  const props=children[0]!,grid=children[1]!;attrs(props);attrs(grid);
  let width:number|undefined,padding:Table['cellPadding'],border:Table['border'],fixed=false;
  const seen=new Set<string>();
  for(const item of elements(props)){
    const n=item.localName;valid(item.namespaceURI===W&&n!==null&&!seen.has(n),'Duplicate or unknown table property');seen.add(n);
    if(n==='tblW')width=dxa(item,1);
    else if(n==='tblLayout'){attrs(item,['w:type']);empty(item);valid(prop(item,'type')==='fixed','Only fixed-width tables supported');fixed=true;}
    else if(n==='tblCellMar'){
      attrs(item);const margins=elements(item);valid(margins.length===4,'Explicit cell margins required');
      const pad:Record<string,number>={};
      for(const child of margins){const edge=child.localName;valid(child.namespaceURI===W&&edge!==null&&(CELL_EDGES as readonly string[]).includes(edge)&&!Object.hasOwn(pad,edge),'Unsupported cell margin');pad[edge]=dxa(child);}
      padding={top:pad.top!,right:pad.right!,bottom:pad.bottom!,left:pad.left!};
    }else if(n==='tblBorders'){
      attrs(item);const edges=elements(item);valid(edges.length===6,'Uniform six-edge table border required');
      const used=new Set<string>();let spec:string|undefined;
      for(const edge of edges){
        const key=edge.localName;valid(edge.namespaceURI===W&&key!==null&&(TABLE_EDGES as readonly string[]).includes(key)&&!used.has(key),'Unsupported table border edge');used.add(key);
        attrs(edge,['w:val','w:sz','w:space','w:color']);empty(edge);
        const value=prop(edge,'val');valid(value==='single','Only uniform single-line borders supported');
        const size=numberProp(edge,'sz',2,96),color=prop(edge,'color')??'';
        valid(/^[0-9a-fA-F]{6}$/.test(color)&&(prop(edge,'space')===undefined||prop(edge,'space')==='0'),'Unsupported table border color or spacing');
        const code=size+':'+color.toUpperCase();valid(spec===undefined||spec===code,'Mixed table border styles unsupported');spec=code;
        border={width:Math.round(size*12700/8),color:{kind:'rgba',red:parseInt(color.slice(0,2),16),green:parseInt(color.slice(2,4),16),blue:parseInt(color.slice(4,6),16),alpha:255}};
      }
    }else unsupported('Unsupported table property '+n);
  }
  valid(fixed&&width!==undefined&&padding,'Fixed table geometry and cell margins required');
  const columns=elements(grid);valid(columns.length>=1&&columns.length<=64,'Table column limit exceeded');
  const columnWidths=columns.map(col=>{requireElement(col,'gridCol');attrs(col,['w:w']);empty(col);return numberProp(col,'w',1,31680)*EMU_PER_TWIP;});
  valid(columnWidths.reduce((a,b)=>a+b,0)===width&&width<=maxWidth&&columnWidths.every(w=>w>padding!.left+padding!.right),'Invalid table grid or content width');
  const rows=children.slice(2);valid(rows.length>=1&&rows.length<=500&&rows.length*columns.length<=10000,'Table cell budget exceeded');charge(rows.length*columns.length);
  return {id,type:'table',columnWidths,cellPadding:padding,...(border?{border}:{}),rows:rows.map((row,ri)=>{
    requireElement(row,'tr');attrs(row);const cells=elements(row);valid(cells.length===columns.length,'Non-rectangular or merged table');
    return {id:id+'-row-'+ri,cells:cells.map((cell,ci)=>{
      requireElement(cell,'tc');attrs(cell);const body=elements(cell);
      valid(body.length>=2&&isW(body[0]!,'tcPr'),'Explicit cell properties and paragraph required');
      const tcPr=body[0]!;attrs(tcPr);const items=elements(tcPr);
      valid(items.length===1&&isW(items[0]!,'tcW')&&dxa(items[0]!,1)===columnWidths[ci],'Cell width differs from grid or unsupported cell properties');
      const blocks=body.slice(1).map(p=>{requireElement(p,'p');const block=read(p);valid(block.type==='paragraph','Page break in table cell');
        for(const run of block.runs)if(run.type==='image')valid(run.width<=columnWidths[ci]!-padding!.left-padding!.right,'Image exceeds cell width');return block;});
      return {id:id+'-cell-'+ri+'-'+ci,blocks};
    })};
  })};
}
function writeTable(table:DeepReadonly<Table>,maxWidth:number,write:(p:DeepReadonly<Paragraph>)=>string,charge:(n:number)=>void):string{
  const pad=table.cellPadding??TABLE_DEFAULT_PADDING;
  const exact=(n:number,min=0)=>{valid(n%EMU_PER_TWIP===0,'Table dimensions require whole twips');return twip(n,min,31680);};
  const widths=table.columnWidths.map(w=>exact(w,1)),sum=widths.reduce((a,b)=>a+b,0);
  valid(sum<=31680&&sum*EMU_PER_TWIP<=maxWidth&&table.columnWidths.every(w=>w>pad.left+pad.right),'Table exceeds page or cell content width');
  valid(table.rows.length<=500&&widths.length<=64&&table.rows.length*widths.length<=10000,'Table cell budget exceeded');charge(table.rows.length*widths.length);
  let properties=tag('tblW',{w:sum,type:'dxa'});
  if(table.border){
    const c=table.border.color;valid(c.kind==='rgba'&&c.alpha===255,'Only opaque RGB table borders supported');
    const size=Math.round(table.border.width*8/12700);valid(size>=2&&size<=96&&Math.round(size*12700/8)===table.border.width,'Table border size is not representable');
    const color=[c.red,c.green,c.blue].map(v=>v.toString(16).padStart(2,'0')).join('').toUpperCase();
    properties+='<w:tblBorders>'+TABLE_EDGES.map(edge=>tag(edge,{val:'single',sz:size,space:0,color})).join('')+'</w:tblBorders>';
  }
  properties+=tag('tblLayout',{type:'fixed'})+'<w:tblCellMar>'+CELL_EDGES.map(edge=>tag(edge,{w:exact(pad[edge]),type:'dxa'})).join('')+'</w:tblCellMar>';
  return '<w:tbl><w:tblPr>'+properties+'</w:tblPr><w:tblGrid>'+widths.map(w=>tag('gridCol',{w})).join('')+'</w:tblGrid>'+table.rows.map(row=>'<w:tr>'+row.cells.map((cell,i)=>{
    valid(cell.blocks.length>0&&cell.blocks.every(b=>b.type==='paragraph'),'Nested tables or non-paragraph cell blocks unsupported');
    for(const block of cell.blocks)if(block.type==='paragraph')for(const run of block.runs)if(run.type==='image')valid(run.width<=table.columnWidths[i]!-pad.left-pad.right,'Image exceeds cell width');
    return '<w:tc><w:tcPr>'+tag('tcW',{w:widths[i]!,type:'dxa'})+'</w:tcPr>'+cell.blocks.map(b=>write(b as DeepReadonly<Paragraph>)).join('')+'</w:tc>';
  }).join('')+'</w:tr>').join('')+'</w:tbl>';
}

function readMain(xml:Document,documentId:string,title:string,profile:DocxStyleProfile|undefined,media:MediaGraph):WriterDocument{
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
    const geometry=readPage(end.properties);page=geometry.page;const blocks:WriterBlock[]=[];
    for(let bi=start;bi<=end.index;bi++){
      const block=children[bi]!;
      if(isW(block,'tbl')){
        valid(!profile||profile.tableDefaultsSafe,'Default table style requires a later adapter');
        blocks.push(readTable(block,'table-'+bi,page.width-page.margins.left-page.margins.right,
          p=>readParagraph(p,true),n=>{cellCount+=n;valid(cellCount<=10000,'Document table-cell budget exceeded');}));
      }else blocks.push(readParagraph(block,false,end.onParagraph&&bi===end.index?end.properties:undefined));
    }
    sections.push({id:'section-'+(sections.length+1),...geometry,blocks});start=end.index+1;
  }
  const result=createDocument('writer',documentId,title);result.defaultTextStyle=profile?.defaults??{};result.styles=profile?.styles??{};result.resources=media.resources;if(profile?.defaultParagraph)result.defaultParagraphStyleId=profile.defaultParagraph;result.sections=sections;
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
function writeRunStyle(style:DeepReadonly<TextStyle>,styleId?:string):string{
  valid(Object.keys(style).every(k=>CHAR_KEYS.includes(k)),'Unsupported character style');
  let r=styleId?tag('rStyle',{val:styleId}):'';
  if(style.fontFamily!==undefined)r+=tag('rFonts',{ascii:style.fontFamily,hAnsi:style.fontFamily,eastAsia:style.fontFamily,cs:style.fontFamily});
  for(const [key,xml] of [['bold','b'],['italic','i']] as const)if(style[key]!==undefined)r+=tag(xml,{val:style[key]?1:0});
  if(style.color!==undefined){
    valid(style.color.kind==='rgba'&&style.color.alpha===255,'Theme or translucent text color unsupported');
    const v=[style.color.red,style.color.green,style.color.blue].map(n=>n.toString(16).padStart(2,'0')).join('').toUpperCase();r+=tag('color',{val:v});
  }
  if(style.fontSize!==undefined){const n=Math.round(style.fontSize/6350);valid(n>=1&&n<=3276,'Font size outside DOCX limits');r+=tag('sz',{val:n});}
  if(style.underline!==undefined)r+=tag('u',{val:style.underline?'single':'none'});
  if(style.runDirection!==undefined)r+=tag('rtl',{val:style.runDirection==='rtl'?1:0});
  const langs:Record<string,string>={};for(const [key,property] of [['val','language'],['eastAsia','eastAsiaLanguage'],['bidi','bidiLanguage']] as const)if(style[property]!==undefined)langs[key]=style[property]!;
  if(Object.keys(langs).length)r+=tag('lang',langs);
  return r?'<w:rPr>'+r+'</w:rPr>':'';
}
function writeParagraphStyle(style:DeepReadonly<TextStyle>,styleId?:string,sectionEnd=''):string{
  valid(Object.keys(style).every(k=>PARA_KEYS.includes(k)),'Unsupported paragraph style');
  let p=styleId?tag('pStyle',{val:styleId}):'';if(style.paragraphDirection!==undefined)p+=tag('bidi',{val:style.paragraphDirection==='rtl'?1:0});const spacing:Record<string,string|number>={};
  if(style.spaceBefore!==undefined)spacing.before=twip(style.spaceBefore,0,31680);
  if(style.spaceAfter!==undefined)spacing.after=twip(style.spaceAfter,0,31680);
  if(style.lineHeight!==undefined){spacing.line=twip(style.lineHeight,1,31680);spacing.lineRule='exact';}
  if(Object.keys(spacing).length)p+=tag('spacing',spacing);
  if(style.indent!==undefined)p+=tag('ind',{left:twip(style.indent,-31680,31680)});
  if(style.alignment!==undefined)p+=tag('jc',{val:({start:'left',center:'center',end:'right',justify:'both'})[style.alignment]});
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
  const checked=validateDocument(candidate);valid(checked.kind==='writer','Expected Writer model');
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
    return '<w:sectPr>'+tag('pgSz',{w:width,h:height,...(width>height?{orient:'landscape'}:{})})+
      tag('pgMar',{top:twip(m.top,0,31680),right:twip(m.right,0,31680),bottom:twip(m.bottom,0,31680),left:twip(m.left,0,31680),header:twip(extra?.headerDistance??0,0,31680),footer:twip(extra?.footerDistance??0,0,31680),gutter:0})+
      (extra?.columnSpace===undefined?'':tag('cols',{num:1,space:twip(extra.columnSpace,0,31680)}))+'</w:sectPr>';
  };
  for(let si=0;si<checked.sections.length;si++){
    section=checked.sections[si]!;const final=si===checked.sections.length-1;
    valid(final||section.blocks.at(-1)?.type==='paragraph','Nonfinal section must end in a paragraph');
    const maxWidth=section.page.width-section.page.margins.left-section.page.margins.right;
    const end=sectionXml(section);
    for(let bi=0;bi<section.blocks.length;bi++){
      const block=section.blocks[bi]!;
      if(block.type==='page-break')body+='<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
      else if(block.type==='table'){valid(!profile||profile.tableDefaultsSafe,'Default table style requires a later adapter');body+=writeTable(block,maxWidth,p=>writeParagraph(p,true),n=>{cellCount+=n;valid(cellCount<=10000,'Document table-cell budget exceeded');});}
      else body+=writeParagraph(block,false,!final&&bi===section.blocks.length-1?end:'');
    }
    if(final)body+=end;
  }
  return {model:checked,xml:DECL+'<w:document xmlns:w="'+W+'"><w:body>'+body+'</w:body></w:document>'};
}
export interface WriterDocxSource{readonly document:DeepReadonly<WriterDocument>;readonly mainPart:string}
const sources=new WeakMap<WriterDocxSource,{pkg:OpcPackage;canonical:string;documentId:string;profile?:DocxStyleProfile;media:MediaGraph}>();
export function openWriterDocx(bytes:Uint8Array,documentId:string,title:string,options:ContainerOptions={}):WriterDocxSource{
  const pkg=openOpc(bytes,options),{part,profile}=readStyleProfile(pkg,options,readRunStyle,readParagraphStyle),media=loadMedia(pkg,part),document=readMain(parseXml(pkg.parts.get(part)!,options),documentId,title,profile,media);
  const source=Object.freeze({document:freezeData(document),mainPart:part});
  sources.set(source,{pkg,canonical:mainXml(document,profile,planMedia(document.resources,part,pkg,media)).xml,documentId,profile,media});
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
  addMediaParts(parts,part,media,options);const bytes=writeZip(parts,options);
  const checked=openWriterDocx(bytes,model.documentId,model.title,options),saved=sources.get(checked)!;
  valid(saved.canonical===xml&&canonicalData(checked.document.resources)===canonicalData(model.resources),'Semantic output verification failed');
  return bytes;
}
/** No-edit returns original archive bytes. Edited copy changes only the supported main part.
 * No IO occurs here: desktop save planning/atomic file commit belongs to a later M2 slice.
 */
export function saveWriterDocx(source:WriterDocxSource,candidate:unknown,options:ContainerOptions={}):Uint8Array{
  const saved=sources.get(source);assert(saved,'invalid','Unknown DOCX source capability');
  const model=validateDocument(candidate);valid(model.kind==='writer','Expected Writer model');assert(model.documentId===saved.documentId,'conflict','Document identity mismatch');
  const media=planMedia(model.resources,source.mainPart,saved.pkg,saved.media),{xml}=mainXml(model,saved.profile,media);
  if(xml===saved.canonical&&!media.additions.length)return writeOpc(saved.pkg,new Map(),options);
  const bytes=saveMediaPackage(saved.pkg,source.mainPart,xml,media,options);
  const checked=openWriterDocx(bytes,model.documentId,model.title,options);
  valid(sources.get(checked)?.canonical===xml&&canonicalData(checked.document.resources)===canonicalData({...saved.media.resources,...model.resources})&&canonicalData(checked.document.styles)===canonicalData(model.styles)&&canonicalData(checked.document.defaultTextStyle)===canonicalData(saved.profile?.defaults??{}),'Edited DOCX verification failed');
  return bytes;
}
