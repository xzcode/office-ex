// SPDX-License-Identifier: MIT
/** Bounded, revision-scoped Writer layout. First profile: horizontal LTR text paragraphs,
 * direct/named text styles, explicit line/page breaks. Not full Word shaping or pagination.
 * Document coordinates are EMU. No derived geometry is persisted in the source model.
 */
import { assert, graphemeBoundaries, type Cancellation } from '@office-ex/contracts';
import { resolveTextStyle, freezeData, paragraphBoundaries, OBJECT_REPLACEMENT, type DocumentSnapshot, type DeepReadonly, type TextStyle, type WriterSection, type Paragraph, type WriterDocument, type Table, type Inline } from '@office-ex/core';
import { measureText, type FontMetrics, type FontRequest, type Point } from './index.js';

export interface WriterPosition {paragraphId:string;offset:number}
export interface WriterCaret {offset:number;x:number}
export interface WriterFragment {text:string;x:number;width:number;request:FontRequest;style:DeepReadonly<TextStyle>}
export interface WriterVisualCluster {start:number;end:number;x:number;width:number}
export interface WriterNativeSlice {originX:number;boxWidth:number;sourceTop:number;paragraphWidth:number}
export interface WriterLine {paragraphId:string;start:number;end:number;x:number;y:number;width:number;height:number;baseline:number;carets:WriterCaret[];fragments:WriterFragment[];images:WriterImageBox[];visualClusters?:WriterVisualCluster[];native?:WriterNativeSlice}
export interface WriterTableCellBox {tableId:string;rowId:string;cellId:string;paragraphIds:string[];x:number;y:number;width:number;height:number;border?:DeepReadonly<NonNullable<Table['border']>>}
export interface WriterImageBox {paragraphId:string;imageId:string;resourceId:string;x:number;y:number;width:number;height:number}
export interface WriterPage {sectionId:string;width:number;height:number;margins:DeepReadonly<WriterSection['page']['margins']>;lines:WriterLine[];tableCells:WriterTableCellBox[];images:WriterImageBox[]}
export interface WriterLayout {revision:number;contentId:string;environmentId:string;pages:WriterPage[];warnings:string[]}
/** Optional native shaping backend. All geometry is relative to the paragraph
 * content box in EMU. The display backend uses the identical native paragraph
 * projection clipped to the page's line ranges. */
export interface NativeParagraphShape {lines:readonly WriterLine[];warnings:readonly string[];spaceBefore:number;spaceAfter:number;clusters:number}
export interface WriterParagraphShaper {readonly environmentId:string;shape(model:DeepReadonly<WriterDocument>,paragraph:DeepReadonly<Paragraph>,width:number):NativeParagraphShape}
export interface WriterLayoutOptions {maxPages?:number;maxClusters?:number;cancellation?:Cancellation;paragraphShaper?:WriterParagraphShaper}
type ImageRun=DeepReadonly<Extract<Inline,{type:'image'}>>;
type Glyph={start:number;end:number;text:string;advance:number;ascent:number;descent:number;style:DeepReadonly<TextStyle>;request:FontRequest;kind:'text'|'tab'|'break'|'image';image?:ImageRun};
const COMPLEX=/[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Gujarati}\p{Script=Gurmukhi}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Thai}\p{Script=Lao}\p{Script=Khmer}\p{Script=Myanmar}\u202a-\u202e\u2066-\u2069]/u;
const CJK=/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;
const NO_START=/^[、。，．？！：；）］｝〉》」』】〕〗〙〛ぁぃぅぇぉっゃゅょァィゥェォッャュョー]/u;
const NO_END=/[（［｛〈《「『【〔〖〘〚]$/u;
function breakBetween(a:Glyph,b:Glyph|undefined):boolean{
  if(!b||a.kind==='break'||b.kind==='break'||a.kind==='image'||b.kind==='image')return true;
  if(NO_END.test(a.text)||NO_START.test(b.text))return false;
  return a.kind==='tab'||/[\u0020\u2000-\u200A\u3000\-]$/u.test(a.text)||CJK.test(a.text)||CJK.test(b.text);
}
function budget(value:number|undefined,fallback:number,maximum:number):number{
  const n=value??fallback;assert(Number.isSafeInteger(n)&&n>0&&n<=maximum,'invalid','Invalid Writer layout budget');return n;
}
function requestFor(provider:FontMetrics,text:string,style:DeepReadonly<TextStyle>,warnings:Set<string>){
  const measured=measureText(provider,text,style,'ltr');
  for(const d of measured.diagnostics)warnings.add(d.code);
  // Actual family is retained for the draw path; never modify the original font name.
  return {metrics:measured.metrics,request:{...measured.request,family:measured.metrics.actualFamily}};
}
function glyphsFor(model:DeepReadonly<WriterDocument>,p:DeepReadonly<Paragraph>,provider:FontMetrics,warnings:Set<string>,charge:(n:number)=>void,check:()=>void):{glyphs:Glyph[];style:DeepReadonly<TextStyle>;minimum:number}{
  const paragraphStyle=p.styleId??model.defaultParagraphStyleId;
  const para=resolveTextStyle(model.defaultTextStyle,model.styles,paragraphStyle?[paragraphStyle]:[],p.style??{});
  for(const d of para.diagnostics)warnings.add(d.code);
  const style=para.style;assert(style.paragraphDirection!=='rtl','unsupported','RTL paragraph needs a bidi layout adapter');const pieces:{text:string;style:DeepReadonly<TextStyle>;image?:ImageRun}[]=[];
  for(const run of p.runs){
    check();
    if(run.type==='image'){
      assert(model.resources[run.resourceId]?.dataBase64,'unsupported','Missing packaged image resource');pieces.push({text:OBJECT_REPLACEMENT,style,image:run});
    }else if(run.type==='text'){
      const resolved=resolveTextStyle(style,model.styles,run.styleId?[run.styleId]:[],run.style??{});
      for(const d of resolved.diagnostics)warnings.add(d.code);
      assert(resolved.style.runDirection!=='rtl','unsupported','RTL run needs a bidi layout adapter');pieces.push({text:run.text,style:resolved.style});
    }else pieces.push({text:run.type==='tab'?'\t':'\n',style});
  }
  const whole=pieces.map(x=>x.text).join('');
  assert(whole.length<=100000,'limit','Paragraph exceeds Writer layout text budget');
  assert(!COMPLEX.test(whole),'unsupported','Complex-script and bidi shaping need a later layout adapter');
  const boundary=paragraphBoundaries(p),valid=new Set(boundary);
  charge(boundary.length-1);
  const glyphs:Glyph[]=[];let offset=0;
  for(const piece of pieces){
    check();const end=offset+piece.text.length;
    // Don't measure halves of a combined character using different fonts.
    assert(valid.has(offset)&&valid.has(end),'unsupported','Grapheme crosses a formatting boundary');
    if(piece.image){const image=piece.image,m=requestFor(provider,' ',piece.style,warnings);glyphs.push({start:offset,end:offset+1,text:OBJECT_REPLACEMENT,advance:image.width,ascent:image.height,descent:0,style:piece.style,request:m.request,kind:'image',image});offset=end;continue;}
    const bounds=graphemeBoundaries(piece.text);let index=0;
    while(index<bounds.length-1){
      check();
      const begin=bounds[index]!,next=bounds[index+1]!,cluster=piece.text.slice(begin,next);
      if(cluster==='\t'||cluster==='\r'||cluster==='\n'||cluster==='\r\n'){
        const m=requestFor(provider,' ',piece.style,warnings);
        glyphs.push({start:offset+begin,end:offset+next,text:cluster,advance:cluster==='\t'?Math.max(1,Math.round(m.metrics.advance*4)):0,ascent:Math.ceil(m.metrics.ascent),descent:Math.ceil(m.metrics.descent),style:piece.style,request:m.request,kind:cluster==='\t'?'tab':'break'});
        index++;continue;
      }
      // Bounded shaping chunks; never split a grapheme. The cluster widths may
      // change when re-shaped at a line boundary, so mark this profile estimated.
      let last=index+1;
      while(last<bounds.length-1&&last-index<128&&bounds[last+1]!-begin<=512){
        if(/[\r\n\t]/.test(piece.text.slice(bounds[last]!,bounds[last+1]!)))break;
        last++;
      }
      const sample=piece.text.slice(begin,bounds[last]!);
      const m=requestFor(provider,sample,piece.style,warnings);
      for(const c of m.metrics.clusters)glyphs.push({start:offset+begin+c.start,end:offset+begin+c.end,text:sample.slice(c.start,c.end),advance:Math.max(0,Math.round(c.right-c.left)),ascent:Math.ceil(m.metrics.ascent),descent:Math.ceil(m.metrics.descent),style:piece.style,request:m.request,kind:'text'});
      index=last;
    }
    offset=end;
  }
  warnings.add('writer.estimated-wrap');
  return {glyphs,style,minimum:Math.max(1,Math.ceil((style.fontSize??152400)*1.2))};
}
/** Returns line boxes and logical caret coordinates. Unsupported content fails
 * explicitly; callers keep the source intact and show a localized preview warning.
 */
export function layoutWriter(snapshot:DocumentSnapshot,provider:FontMetrics,options:WriterLayoutOptions={}):DeepReadonly<WriterLayout>{
  assert(snapshot.document.kind==='writer','invalid','Expected Writer snapshot');
  const model=snapshot.document;
  const maxPages=budget(options.maxPages,100,500),maxClusters=budget(options.maxClusters,30000,100000);
  const warnings=new Set<string>(),pages:WriterPage[]=[];let clusterCount=0,blocks=0;
  const check=()=>options.cancellation?.throwIfCancelled();
  const charge=(n:number)=>{clusterCount+=n;assert(clusterCount<=maxClusters,'limit','Writer layout cluster budget exceeded');};
  for(const section of model.sections){
    check();const spec=section.page,m=spec.margins;
    const usableWidth=spec.width-m.left-m.right,usableHeight=spec.height-m.top-m.bottom;
    assert(usableWidth>0&&usableHeight>0,'invalid','Page has no usable area');
    const makePage=():WriterPage=>{assert(pages.length<maxPages,'limit','Writer layout page budget exceeded');const p:WriterPage={sectionId:section.id,width:spec.width,height:spec.height,margins:m,lines:[],tableCells:[],images:[]};pages.push(p);return p;};
    let page=makePage(),y=m.top;
    const advancePage=()=>{page=makePage();y=m.top;};
    const space=(n:number)=>{if(n<=0)return;if(y+n>spec.height-m.bottom&&(page.lines.length||page.images.length))advancePage();y=Math.min(spec.height-m.bottom,y+n);};

    const layParagraph=(block:DeepReadonly<Paragraph>,originX:number,boxWidth:number,addSpace:(n:number)=>void,place:(line:WriterLine)=>void):void=>{
      check();assert(++blocks<=10000,'limit','Writer layout block budget exceeded');
      if(options.paragraphShaper){
        const shaped=options.paragraphShaper.shape(model,block,boxWidth);
        assert(Number.isSafeInteger(shaped.clusters)&&shaped.clusters>=0,'invalid','Invalid native cluster count');
        charge(shaped.clusters);for(const warning of shaped.warnings)warnings.add(warning);
        addSpace(shaped.spaceBefore);
        for(const line of shaped.lines){
          check();assert(line.paragraphId===block.id&&line.height>0&&line.height<=usableHeight&&line.native,'unsupported','Native line exceeds usable page height');
          const adjusted:WriterLine={...line,x:line.x+originX,carets:line.carets.map(c=>({...c,x:c.x+originX})),
            fragments:line.fragments.map(f=>({...f,x:f.x+originX})),images:line.images.map(i=>({...i,x:i.x+originX})),
            visualClusters:line.visualClusters?.map(c=>({...c,x:c.x+originX})),native:{...line.native,originX:line.native.originX+originX}};
          place(adjusted);
        }
        addSpace(shaped.spaceAfter);return;
      }
      const {glyphs,style,minimum}=glyphsFor(model,block,provider,warnings,charge,check);
      const indent=Math.max(0,style.indent??0);
      if((style.indent??0)<0)warnings.add('writer.negative-indent-limited');
      const available=boxWidth-indent;
      assert(available>0,'unsupported','Paragraph indent consumes page width');
      if(style.alignment==='justify')warnings.add('writer.justification-pending');
      addSpace(style.spaceBefore??0);
      const emit=(tokens:Glyph[],start:number,end:number,forced:boolean)=>{
        check();
        const width=tokens.reduce((sum,g)=>sum+g.advance,0);
        const picturesOnly=tokens.length>0&&tokens.every(g=>g.kind==='image');
        let ascent=picturesOnly?0:Math.ceil((style.fontSize??152400)*0.8),descent=picturesOnly?0:Math.ceil((style.fontSize??152400)*0.2);
        for(const g of tokens){if(g.image){assert(g.advance<=available,'unsupported','Inline image exceeds paragraph width');ascent=Math.max(ascent,g.ascent);}else{ascent=Math.max(ascent,g.ascent,Math.ceil(g.request.size*0.8));descent=Math.max(descent,g.descent,Math.ceil(g.request.size*0.2));}}
        const natural=Math.max(picturesOnly?0:minimum,ascent+descent);
        if(style.keepNext||style.keepLines||style.widowControl)warnings.add('writer.keep-pagination-limited');
        if(style.lineHeight!==undefined&&style.lineHeight<natural)warnings.add('writer.exact-lineheight-limited');
        const height=Math.max(natural,Math.ceil(natural*(style.lineMultiple??240)/240),style.lineHeight??0,style.minLineHeight??0);
        assert(height<=usableHeight,'unsupported','Line exceeds usable page height');
        const x=originX+indent+(style.alignment==='center'?Math.max(0,Math.round((available-width)/2)):style.alignment==='end'?Math.max(0,available-width):0);
        if(width>available)warnings.add('writer.line-overflow');
        if(forced)warnings.add('writer.grapheme-wrap');
        const baseline=Math.floor((height-ascent-descent)/2)+ascent;
        const carets:WriterCaret[]=[{offset:start,x}],fragments:WriterFragment[]=[],images:WriterImageBox[]=[];let px=x;
        for(const g of tokens){
          if(g.image){images.push({paragraphId:block.id,imageId:g.image.id,resourceId:g.image.resourceId,x:px,y:baseline-g.image.height,width:g.image.width,height:g.image.height});}
          else if(g.kind==='text'){
            const prev=fragments.at(-1);
            if(prev&&prev.x+prev.width===px&&prev.style===g.style&&JSON.stringify(prev.request)===JSON.stringify(g.request)){prev.text+=g.text;prev.width+=g.advance;}
            else fragments.push({text:g.text,x:px,width:g.advance,request:g.request,style:g.style});
          }
          px+=g.advance;carets.push({offset:g.end,x:px});
        }
        if(carets.at(-1)!.offset!==end)carets.push({offset:end,x:px});
        place({paragraphId:block.id,start,end,x,y:0,width,height,baseline,carets,fragments,images});
      };
      let begin=0;let finalHard=false;
      while(begin<glyphs.length){
        check();let i=begin,width=0,lastBreak=-1;
        for(;i<glyphs.length;i++){
          const g=glyphs[i]!;
          if(g.kind==='break')break;
          if(width+g.advance>available&&i>begin)break;
          width+=g.advance;
          if(breakBetween(g,glyphs[i+1]))lastBreak=i+1;
        }
        const hard=i<glyphs.length&&glyphs[i]!.kind==='break';
        const reachedEnd=i===glyphs.length;
        if(hard){const tokens=glyphs.slice(begin,i+1);emit(tokens,glyphs[begin]!.start,glyphs[i]!.end,false);begin=i+1;finalHard=begin===glyphs.length;}
        else{const end=reachedEnd?i:lastBreak>begin?lastBreak:i;assert(end>begin,'internal','Writer layout made no progress');emit(glyphs.slice(begin,end),glyphs[begin]!.start,glyphs[end-1]!.end,!reachedEnd&&lastBreak<=begin);begin=end;finalHard=false;}
      }
      if(!glyphs.length||finalHard){const end=glyphs.at(-1)?.end??0;emit([],end,end,false);}
      addSpace(style.spaceAfter??0);
    };
    const placeLine=(line:WriterLine,top:number):void=>{
      const images=line.images.map(image=>({...image,y:image.y+top}));
      page.lines.push({...line,y:top,baseline:line.baseline+top,images});page.images.push(...images);
    };
    const placeTop=(line:WriterLine):void=>{
      if(y+line.height>spec.height-m.bottom)advancePage();
      placeLine(line,y);y+=line.height;
    };
    const layTable=(table:DeepReadonly<Table>):void=>{
      const pad=table.cellPadding??{top:38100,right:57150,bottom:38100,left:57150};
      const totalWidth=table.columnWidths.reduce((a,b)=>a+b,0);
      assert(totalWidth<=usableWidth,'unsupported','Table exceeds page content width');
      assert(table.columnWidths.every(w=>w>pad.left+pad.right),'unsupported','Table cell has no content width');
      warnings.add('writer.table-row-pagination');
      for(const row of table.rows){
        check();assert(++blocks<=10000,'limit','Writer table row budget exceeded');
        const rowCells:{box:WriterTableCellBox;lines:WriterLine[];height:number}[]=[];
        let x=m.left,rowHeight=0;
        for(let col=0;col<row.cells.length;col++){
          const cell=row.cells[col]!,width=table.columnWidths[col]!,lines:WriterLine[]=[];let localY=pad.top;
          const addSpace=(n:number)=>{localY+=n;assert(localY+pad.bottom<=usableHeight,'unsupported','Table row exceeds page content height');};
          const place=(line:WriterLine)=>{
            assert(line.x+line.width<=x+width-pad.right,'unsupported','Table cell line exceeds content width');
            lines.push({...line,y:localY,baseline:localY+line.baseline,images:line.images.map(image=>({...image,y:image.y+localY}))});addSpace(line.height);
          };
          for(const child of cell.blocks){
            assert(child.type==='paragraph','unsupported','Nested tables and page breaks within cells unsupported');
            layParagraph(child,x+pad.left,width-pad.left-pad.right,addSpace,place);
          }
          const h=localY+pad.bottom;
          assert(h>0&&h<=usableHeight,'unsupported','Table row exceeds page content height');
          rowHeight=Math.max(rowHeight,h);
          const box:WriterTableCellBox={tableId:table.id,rowId:row.id,cellId:cell.id,paragraphIds:cell.blocks.map(b=>b.id),x,y:0,width,height:0,...(table.border?{border:table.border}:{})};
          rowCells.push({box,lines,height:h});x+=width;
        }
        // Keep every row together. Oversized rows are rejected rather than clipped.
        if(y+rowHeight>spec.height-m.bottom)advancePage();
        for(const item of rowCells){
          page.tableCells.push({...item.box,y,height:rowHeight});
          for(const line of item.lines){const images=line.images.map(image=>({...image,y:image.y+y}));page.lines.push({...line,y:y+line.y,baseline:y+line.baseline,images});page.images.push(...images);}
        }
        y+=rowHeight;
      }
    };
    for(const block of section.blocks){
      check();
      if(block.type==='page-break'){assert(++blocks<=10000,'limit','Writer layout block budget exceeded');advancePage();continue;}
      if(block.type==='table'){layTable(block);continue;}
      layParagraph(block,m.left,usableWidth,space,placeTop);
    }
  }
  return freezeData({revision:snapshot.revision,contentId:snapshot.contentId,environmentId:options.paragraphShaper?.environmentId??provider.environmentId,pages,warnings:[...warnings].sort()});
}
type ReadLayout=DeepReadonly<WriterLayout>;
export function hitTestWriterPage(layout:ReadLayout,pageIndex:number,point:Point):WriterPosition|undefined{
  assert(Number.isFinite(point.x)&&Number.isFinite(point.y),'invalid','Invalid Writer hit point');
  const page=layout.pages[pageIndex];if(!page?.lines.length)return undefined;
  const cell=page.tableCells.find(c=>point.x>=c.x&&point.x<c.x+c.width&&point.y>=c.y&&point.y<c.y+c.height);
  const candidates=cell?page.lines.filter(l=>cell.paragraphIds.includes(l.paragraphId)):page.lines;
  if(!candidates.length)return undefined;
  const distance=(line:DeepReadonly<WriterLine>)=>Math.max(line.y-point.y,point.y-line.y-line.height,0);
  const horizontal=(line:DeepReadonly<WriterLine>)=>Math.max(line.x-point.x,point.x-line.x-line.width,0);
  let nearest=candidates[0]!;
  for(const line of candidates)if(distance(line)<distance(nearest)||distance(line)===distance(nearest)&&horizontal(line)<horizontal(nearest))nearest=line;
  let caret=nearest.carets[0]!;
  for(const c of nearest.carets)if(Math.abs(c.x-point.x)<Math.abs(caret.x-point.x))caret=c;
  return {paragraphId:nearest.paragraphId,offset:caret.offset};
}
export interface WriterSelectionBox {pageIndex:number;x:number;y:number;width:number;height:number}
export function writerSelectionBoxes(layout:ReadLayout,paragraphId:string,start:number,end:number):WriterSelectionBox[]{
  assert(Number.isSafeInteger(start)&&Number.isSafeInteger(end)&&start>=0&&end>=start,'invalid','Invalid Writer selection');
  const output:WriterSelectionBox[]=[];
  for(let pageIndex=0;pageIndex<layout.pages.length;pageIndex++)for(const line of layout.pages[pageIndex]!.lines){
    if(line.paragraphId!==paragraphId||start===end||end<=line.start||start>=line.end)continue;
    const a=Math.max(start,line.start),b=Math.min(end,line.end);
    const first=line.carets.find(c=>c.offset===a),last=line.carets.find(c=>c.offset===b);
    assert(first&&last,'invalid','Selection splits a grapheme');
    if(line.visualClusters){
      const selected=line.visualClusters.filter(c=>c.start>=a&&c.end<=b).sort((a,b)=>a.x-b.x);
      for(const c of selected){const prev=output.at(-1);if(prev&&prev.pageIndex===pageIndex&&prev.y===line.y&&prev.x+prev.width>=c.x-100){prev.width=Math.max(prev.width,c.x+c.width-prev.x);}else output.push({pageIndex,x:c.x,y:line.y,width:Math.max(1,c.width),height:line.height});}
    }else output.push({pageIndex,x:Math.min(first.x,last.x),y:line.y,width:Math.abs(last.x-first.x),height:line.height});
  }
  return output;
}
export function writerCaretBox(layout:ReadLayout,position:WriterPosition,preferredPage?:number,preferredX?:number):WriterSelectionBox|undefined{
  const found:WriterSelectionBox[]=[];
  for(let pageIndex=0;pageIndex<layout.pages.length;pageIndex++)for(const line of layout.pages[pageIndex]!.lines){
    if(line.paragraphId!==position.paragraphId)continue;
    const caret=line.carets.find(c=>c.offset===position.offset);
    if(caret)found.push({pageIndex,x:caret.x,y:line.y,width:0,height:line.height});
  }
  const candidates=found.filter(b=>b.pageIndex===preferredPage);const pool=candidates.length?candidates:found;
  if(preferredX!==undefined&&pool.length)return pool.reduce((a,b)=>Math.abs(a.x-preferredX)<=Math.abs(b.x-preferredX)?a:b);
  return pool.at(-1);
}
