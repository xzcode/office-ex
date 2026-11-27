// SPDX-License-Identifier: MIT
/** Uses Chromium's HarfBuzz/ICU/CSS text layout and Range geometry, not a
 * secondary office editor. The source of truth remains the Writer model.
 * The displayed paragraph is made by the exact same projection as measurement.
 */
import { bundledFontFamily } from '@office-ex/pdf';
import {assert,EMU_PER_CSS_PIXEL,OfficeError} from '@office-ex/contracts';
import {paragraphBoundaries,inlineText,resolveTextStyle,writerEffectiveStyle,type DeepReadonly,type Paragraph,type WriterDocument,type TextStyle,type Inline} from '@office-ex/core';
import type {WriterParagraphShaper,NativeParagraphShape,WriterLine,WriterCaret,WriterImageBox,WriterVisualCluster} from '@office-ex/render';
import {writerParagraphGeometry} from '@office-ex/render';
const PX=EMU_PER_CSS_PIXEL;
const emu=(value:number)=>Math.round(value*PX);
const px=(value:number)=>value/PX+'px';
const generic=/^(sans-serif|serif|monospace|system-ui)$/;
function fontFamily(value:string|undefined):string{return !value?'sans-serif':generic.test(value)?value:JSON.stringify(value)+', sans-serif';}
function styleText(el:HTMLElement,style:DeepReadonly<TextStyle>):void{
  el.style.fontFamily=[...new Set([style.fontAscii??style.fontFamily,style.fontHighAnsi,style.fontEastAsia,style.fontComplex].filter((v):v is string=>!!v))].map(v=>bundledFontFamily(v,style.language)).concat(bundledFontFamily(undefined,style.language)).join(', ');el.style.fontSize=px(style.fontSize??152400);
  el.style.fontWeight=style.bold?'700':'400';el.style.fontStyle=style.italic?'italic':'normal';
  el.style.textDecoration=[style.underline?'underline':'',style.writerStrike?'line-through':''].filter(Boolean).join(' ')||'none';
  const shade=style.writerShading;if(shade?.kind==='rgba')el.style.backgroundColor='rgb('+[shade.red,shade.green,shade.blue].join(',')+')';
  if(style.characterSpacing!==undefined)el.style.letterSpacing=px(style.characterSpacing);
  if(style.kerningThreshold!==undefined)el.style.fontKerning=(style.fontSize??152400)>=style.kerningThreshold?'normal':'none';
  const c=style.color;el.style.color=c?.kind==='rgba'?'rgba('+[c.red,c.green,c.blue,c.alpha/255].join(',')+')':'#20242c';
  if(style.language)el.lang=style.language;
  if(style.runDirection){el.dir=style.runDirection;el.style.unicodeBidi='embed';}
}
type Piece={start:number;end:number;node:Text|HTMLImageElement;image?:DeepReadonly<Extract<Inline,{type:'image'}>>};
interface Projection {element:HTMLDivElement;pieces:Piece[];pitch:number;indent:number;boxWidth:number;style:DeepReadonly<TextStyle>;warnings:string[];length:number}
function projection(model:DeepReadonly<WriterDocument>,p:DeepReadonly<Paragraph>,width:number):Projection{
  const warnings=new Set<string>(),named=p.styleId??model.defaultParagraphStyleId;
  const para=resolveTextStyle(model.defaultTextStyle,model.styles,named?[named]:[],p.style??{});
  for(const d of para.diagnostics)warnings.add(d.code);
  const style=writerEffectiveStyle(model,p);
  const geometry=writerParagraphGeometry(style,width),indent=geometry.origin;if(style.indent!==undefined&&style.indent<0)warnings.add('writer.negative-indent-limited');
  const resolved=p.runs.map(run=>{
    if(run.type!=='text')return style;
    const result=resolveTextStyle(style,model.styles,run.styleId?[run.styleId]:[],run.style??{});
    for(const d of result.diagnostics)warnings.add(d.code);return result.style;
  });
  // A conservative uniform line box avoids splitting one native glyph or
  // clipping a tall inline picture when moving a page slice. Word's varying
  // line metrics and exact-height clipping are still fidelity limitations.
  let natural=Math.ceil((style.fontSize??152400)*1.55);
  for(let i=0;i<p.runs.length;i++){
    const run=p.runs[i]!;
    natural=Math.max(natural,run.type==='image'?run.height+Math.ceil((style.fontSize??152400)*0.4):Math.ceil((resolved[i]!.fontSize??152400)*1.55));
  }
  const pitch=Math.max(1,Math.ceil(Math.max(natural*(style.lineMultiple??240)/240,style.lineHeight??0,style.minLineHeight??0,natural)/PX*64)/64);
  warnings.add('writer.native-browser-line-metrics');
  if(style.lineHeight!==undefined&&style.lineHeight<natural)warnings.add('writer.exact-lineheight-limited');
  const element=document.createElement('div');element.className='writer-native-paragraph';
  styleText(element,style);element.style.boxSizing='border-box';element.style.width=px(geometry.boxWidth);
  element.style.lineHeight=pitch+'px';element.style.margin='0';element.style.padding='0';element.style.paddingLeft=px(geometry.padLeft);element.style.paddingRight=px(geometry.padRight);element.style.textIndent=px(geometry.first);element.style.border='0';
  element.style.whiteSpace='pre-wrap';element.style.overflowWrap='anywhere';element.style.wordBreak='normal';element.style.lineBreak='strict';
  element.style.tabSize='4';element.style.unicodeBidi='normal';element.dir=style.paragraphDirection??(style.firstLineIndent!==undefined?'ltr':'auto');
  element.style.textAlign=style.alignment??'start';
  const pieces:Piece[]=[];let offset=0;
  for(let i=0;i<p.runs.length;i++){
    const run=p.runs[i]!,content=inlineText(run);
    if(run.type==='image'){
      const resource=model.resources[run.resourceId];assert(resource?.dataBase64&&(resource.mimeType==='image/png'||resource.mimeType==='image/jpeg'),'unsupported','Missing image bytes');
      assert(run.width<=geometry.bodyWidth-Math.max(0,geometry.first),'unsupported','Image exceeds paragraph width');
      const img=document.createElement('img');img.alt='';img.draggable=false;img.style.width=px(run.width);img.style.height=px(run.height);img.style.verticalAlign='top';img.style.objectFit='fill';img.style.display='inline-block';img.src='data:'+resource.mimeType+';base64,'+resource.dataBase64;
      element.append(img);pieces.push({start:offset,end:offset+1,node:img,image:run});offset++;
    }else if(content.length){
      assert(!content.includes('\r'),'unsupported','Raw carriage return cannot be shaped without normalization');
      const span=document.createElement('span');styleText(span,resolved[i]!);span.style.lineHeight='inherit';span.style.verticalAlign='top';const text=document.createTextNode(content);span.append(text);element.append(span);
      pieces.push({start:offset,end:offset+content.length,node:text});offset+=content.length;
    }
  }
  if(!offset||pieces.at(-1)?.node.textContent?.endsWith('\n')){
    // A visual zero-width sentinel represents the empty final line. It is not
    // part of the text projection, source model, clipboard or DOCX.
    const sentinel=document.createElement('span');sentinel.textContent='\u200b';sentinel.dataset.writerSentinel='true';element.append(sentinel);
  }
  return {element,pieces,pitch,indent,boxWidth:geometry.boxWidth,style,warnings:[...warnings],length:offset};
}
function endpoint(piece:Piece,offset:number):{node:Node;offset:number}{
  if(piece.node.nodeType===Node.TEXT_NODE)return {node:piece.node,offset:offset-piece.start};
  const parent=piece.node.parentNode!;const index=Array.prototype.indexOf.call(parent.childNodes,piece.node) as number;return {node:parent,offset:index+(offset===piece.end?1:0)};
}
function union(rects:DOMRect[]):{left:number;right:number;top:number;bottom:number}{
  return {left:Math.min(...rects.map(r=>r.left)),right:Math.max(...rects.map(r=>r.right)),top:Math.min(...rects.map(r=>r.top)),bottom:Math.max(...rects.map(r=>r.bottom))};
}
interface CacheEntry{width:number;defaultParagraphStyleId?:string;defaults:unknown;styles:unknown;resources:unknown;numbering:unknown;result:NativeParagraphShape}
/** Session lifetime cache. Immutable paragraph identity + style/resources and
 * width invalidate precisely. Font loading changes the environment generation. */
export class BrowserParagraphShaper implements WriterParagraphShaper{
  private cache=new WeakMap<object,CacheEntry>();
  private generation=0;
  private signature='';
  get environmentId():string{
    const current=navigator.userAgent+':'+devicePixelRatio+':'+document.fonts.status;
    if(current!==this.signature){this.signature=current;this.generation++;this.cache=new WeakMap();}
    return 'chromium-range:'+current+':'+this.generation;
  }
  measurements=0;hits=0;
  shape(model:DeepReadonly<WriterDocument>,p:DeepReadonly<Paragraph>,width:number):NativeParagraphShape{
    void this.environmentId;
    const cached=this.cache.get(p);
    if(cached&&cached.width===width&&cached.defaultParagraphStyleId===model.defaultParagraphStyleId&&cached.defaults===model.defaultTextStyle&&cached.styles===model.styles&&cached.resources===model.resources&&cached.numbering===model.numbering){this.hits++;return cached.result;}
    const boundaries=paragraphBoundaries(p);assert(boundaries.length<=20001,'limit','Native paragraph shaping budget exceeded');
    const q=projection(model,p,width);
    const host=document.createElement('div');host.className='writer-native-measure';host.style.position='fixed';host.style.left='-100000px';host.style.top='0';host.style.visibility='hidden';host.style.pointerEvents='none';host.style.zIndex='-1';host.style.margin='0';host.append(q.element);document.body.append(host);
    try{
      const box=q.element.getBoundingClientRect(),height=box.height,pitch=q.pitch;
      assert(Number.isFinite(height)&&height>0&&height<1000000,'limit','Native paragraph height out of range');
      const rows=Math.max(1,Math.round(height/pitch));
      assert(rows<=10000&&Math.abs(rows*pitch-height)<Math.max(0.5,rows*0.02),'unsupported','Native line height requires a variable line-box adapter');
      const rowOf=(y:number)=>Math.max(0,Math.min(rows-1,Math.floor((y-box.top)/pitch)));
      const lines:WriterLine[]=Array.from({length:rows},(_,i)=>({paragraphId:p.id,start:0,end:0,x:q.indent,y:0,width:q.boxWidth,height:emu(pitch),baseline:emu(pitch*0.8),carets:[],fragments:[],images:[],visualClusters:[],native:{originX:q.indent,boxWidth:q.boxWidth,sourceTop:emu(i*pitch),paragraphWidth:width}}));
      let pieceIndex=0,lastOffset=0,lastRow=0;
      const range=document.createRange();
      const collapse=(pos:{node:Node;offset:number}):DOMRect|undefined=>{range.setStart(pos.node,pos.offset);range.collapse(true);return Array.from(range.getClientRects()).find(r=>r.height>0);};
      const addCaret=(line:WriterLine,offset:number,x:number)=>{if(!line.carets.some(c=>c.offset===offset&&Math.abs(c.x-x)<=1))line.carets.push({offset,x});};
      for(let i=0;i<boundaries.length-1;i++){
        const start=boundaries[i]!,end=boundaries[i+1]!;
        while(pieceIndex<q.pieces.length-1&&start>=q.pieces[pieceIndex]!.end)pieceIndex++;
        const first=q.pieces[pieceIndex]!;assert(first&&start>=first.start,'invalid','Native position missing');
        let lastIndex=pieceIndex;while(end>q.pieces[lastIndex]!.end)lastIndex++;
        const last=q.pieces[lastIndex]!,a=endpoint(first,start),b=endpoint(last,end);
        range.setStart(a.node,a.offset);range.setEnd(b.node,b.offset);
        let rects=first.image?[first.node instanceof HTMLImageElement?first.node.getBoundingClientRect():range.getBoundingClientRect()]:Array.from(range.getClientRects());
        rects=rects.filter(r=>r.height>0||r.width>0);
        const ac=collapse(a),bc=collapse(b);
        let row=rects.length?rowOf((rects[0]!.top+rects[0]!.bottom)/2):ac?rowOf((ac.top+ac.bottom)/2):lastRow;
        assert(row>=lastRow,'unsupported','Unexpected native vertical order');
        // Newline ranges may contain a rect at the following line. Assign
        // the break to the preceding logical line and keep the next caret.
        const chars=first.node.nodeType===Node.TEXT_NODE&&first===last?(first.node.textContent??'').slice(start-first.start,end-first.start):'';
        if(chars==='\n')row=ac?rowOf((ac.top+ac.bottom)/2):lastRow;
        const line=lines[row]!,onLine=rects.filter(r=>rowOf((r.top+r.bottom)/2)===row);
        const rect=onLine.length?union(onLine):{left:ac?.left??box.left,right:ac?.left??box.left,top:box.top+row*pitch,bottom:box.top+(row+1)*pitch};
        const left=emu(rect.left-box.left)+q.indent,right=emu(rect.right-box.left)+q.indent;
        const startX=ac&&rowOf((ac.top+ac.bottom)/2)===row?emu(ac.left-box.left)+q.indent:undefined;
        const endX=bc&&rowOf((bc.top+bc.bottom)/2)===row?emu(bc.left-box.left)+q.indent:undefined;
        const near=(x:number|undefined):number|undefined=>x!==undefined&&(Math.abs(x-left)<=2*PX||Math.abs(x-right)<=2*PX)?(Math.abs(x-left)<Math.abs(x-right)?left:right):undefined;
        let l=near(startX),r=near(endX);
        if(l===undefined&&r!==undefined)l=r===left?right:left;
        if(r===undefined&&l!==undefined)r=l===left?right:left;
        if(l===undefined||r===undefined||l===r&&right>left){const rtl=/[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}]/u.test(chars);l=rtl?right:left;r=rtl?left:right;}
        if(!line.carets.length)line.start=start;
        line.end=end;addCaret(line,start,l);addCaret(line,end,r);
        line.visualClusters!.push({start,end,x:left,width:Math.max(0,right-left)});
        if(first.image){
          const img=first.image;line.images.push({paragraphId:p.id,imageId:img.id,resourceId:img.resourceId,x:left,y:emu(rect.top-box.top-row*pitch),width:img.width,height:img.height});
        }
        lastOffset=end;lastRow=row;
      }
      // Empty paragraphs and trailing hard lines still have caret positions.
      for(let i=0;i<lines.length;i++){
        const line=lines[i]!;
        if(!line.carets.length){line.start=line.end=i===0?0:lines[i-1]!.end;const sentinel=q.element.querySelector('[data-writer-sentinel]');const rect=sentinel?.getBoundingClientRect();line.carets.push({offset:line.end,x:rect?emu(rect.left-box.left)+q.indent:q.indent});}
        assert(line.end>=line.start&&line.carets.every(c=>Number.isFinite(c.x)),'invalid','Invalid native caret');
      }
      const result:NativeParagraphShape={lines,warnings:q.warnings,spaceBefore:q.style.spaceBefore??0,spaceAfter:q.style.spaceAfter??0,clusters:boundaries.length-1};
      this.cache.set(p,{width,defaultParagraphStyleId:model.defaultParagraphStyleId,defaults:model.defaultTextStyle,styles:model.styles,resources:model.resources,numbering:model.numbering,result});this.measurements++;return result;
    }finally{host.remove();}
  }
  render(model:DeepReadonly<WriterDocument>,p:DeepReadonly<Paragraph>,width:number):HTMLDivElement{
    return projection(model,p,width).element;
  }
}

