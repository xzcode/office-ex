// SPDX-License-Identifier: MIT
/** Bounded, revision-scoped Writer layout. First profile: horizontal LTR text paragraphs,
 * direct/named text styles, explicit line/page breaks. Not full Word shaping or pagination.
 * Document coordinates are EMU. No derived geometry is persisted in the source model.
 */
import { assert, graphemeBoundaries, type Cancellation } from '@office-ex/contracts';
import { resolveTextStyle, freezeData, type DocumentSnapshot, type DeepReadonly, type TextStyle, type WriterSection, type Paragraph, type WriterDocument } from '@office-ex/core';
import { measureText, type FontMetrics, type FontRequest, type Point } from './index.js';

export interface WriterPosition {paragraphId:string;offset:number}
export interface WriterCaret {offset:number;x:number}
export interface WriterFragment {text:string;x:number;width:number;request:FontRequest;style:DeepReadonly<TextStyle>}
export interface WriterLine {paragraphId:string;start:number;end:number;x:number;y:number;width:number;height:number;baseline:number;carets:WriterCaret[];fragments:WriterFragment[]}
export interface WriterPage {sectionId:string;width:number;height:number;margins:DeepReadonly<WriterSection['page']['margins']>;lines:WriterLine[]}
export interface WriterLayout {revision:number;contentId:string;environmentId:string;pages:WriterPage[];warnings:string[]}
export interface WriterLayoutOptions {maxPages?:number;maxClusters?:number;cancellation?:Cancellation}
type Glyph={start:number;end:number;text:string;advance:number;ascent:number;descent:number;style:DeepReadonly<TextStyle>;request:FontRequest;kind:'text'|'tab'|'break'};
const COMPLEX=/[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Gujarati}\p{Script=Gurmukhi}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Thai}\p{Script=Lao}\p{Script=Khmer}\p{Script=Myanmar}\u202a-\u202e\u2066-\u2069]/u;
const CJK=/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;
const NO_START=/^[、。，．？！：；）］｝〉》」』】〕〗〙〛ぁぃぅぇぉっゃゅょァィゥェォッャュョー]/u;
const NO_END=/[（［｛〈《「『【〔〖〘〚]$/u;
function breakBetween(a:Glyph,b:Glyph|undefined):boolean{
  if(!b||a.kind==='break'||b.kind==='break')return true;
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
  const para=resolveTextStyle(model.defaultTextStyle,model.styles,p.styleId?[p.styleId]:[],p.style??{});
  for(const d of para.diagnostics)warnings.add(d.code);
  const style=para.style;const pieces:{text:string;style:DeepReadonly<TextStyle>}[]=[];
  for(const run of p.runs){
    check();
    assert(run.type!=='image','unsupported','Inline images need a later layout adapter');
    if(run.type==='text'){
      const resolved=resolveTextStyle(style,model.styles,run.styleId?[run.styleId]:[],run.style??{});
      for(const d of resolved.diagnostics)warnings.add(d.code);
      pieces.push({text:run.text,style:resolved.style});
    }else pieces.push({text:run.type==='tab'?'\t':'\n',style});
  }
  const whole=pieces.map(x=>x.text).join('');
  assert(whole.length<=100000,'limit','Paragraph exceeds Writer layout text budget');
  assert(!COMPLEX.test(whole),'unsupported','Complex-script and bidi shaping need a later layout adapter');
  const boundary=graphemeBoundaries(whole),valid=new Set(boundary);
  charge(boundary.length-1);
  const glyphs:Glyph[]=[];let offset=0;
  for(const piece of pieces){
    check();const end=offset+piece.text.length;
    // Don't measure halves of a combined character using different fonts.
    assert(valid.has(offset)&&valid.has(end),'unsupported','Grapheme crosses a formatting boundary');
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
    const makePage=():WriterPage=>{assert(pages.length<maxPages,'limit','Writer layout page budget exceeded');const p:WriterPage={sectionId:section.id,width:spec.width,height:spec.height,margins:m,lines:[]};pages.push(p);return p;};
    let page=makePage(),y=m.top;
    const advancePage=()=>{page=makePage();y=m.top;};
    const space=(n:number)=>{if(n<=0)return;if(y+n>spec.height-m.bottom&&page.lines.length)advancePage();y=Math.min(spec.height-m.bottom,y+n);};
    for(const block of section.blocks){
      check();assert(++blocks<=10000,'limit','Writer layout block budget exceeded');
      if(block.type==='page-break'){advancePage();continue;}
      assert(block.type==='paragraph','unsupported','Tables need a later layout adapter');
      const {glyphs,style,minimum}=glyphsFor(model,block,provider,warnings,charge,check);
      const indent=Math.max(0,style.indent??0);
      if((style.indent??0)<0)warnings.add('writer.negative-indent-limited');
      const available=usableWidth-indent;
      assert(available>0,'unsupported','Paragraph indent consumes page width');
      if(style.alignment==='justify')warnings.add('writer.justification-pending');
      space(style.spaceBefore??0);
      const emit=(tokens:Glyph[],start:number,end:number,forced:boolean)=>{
        check();
        const width=tokens.reduce((sum,g)=>sum+g.advance,0);
        let ascent=Math.ceil((style.fontSize??152400)*0.8),descent=Math.ceil((style.fontSize??152400)*0.2);
        for(const g of tokens){ascent=Math.max(ascent,g.ascent,Math.ceil(g.request.size*0.8));descent=Math.max(descent,g.descent,Math.ceil(g.request.size*0.2));}
        const natural=Math.max(minimum,ascent+descent);
        if(style.lineHeight!==undefined&&style.lineHeight<natural)warnings.add('writer.exact-lineheight-limited');
        const height=Math.max(natural,style.lineHeight??0);
        assert(height<=usableHeight,'unsupported','Line exceeds usable page height');
        if(y+height>spec.height-m.bottom)advancePage();
        const x=m.left+indent+(style.alignment==='center'?Math.max(0,Math.round((available-width)/2)):style.alignment==='end'?Math.max(0,available-width):0);
        if(width>available)warnings.add('writer.line-overflow');
        if(forced)warnings.add('writer.grapheme-wrap');
        const carets:WriterCaret[]=[{offset:start,x}],fragments:WriterFragment[]=[];let px=x;
        for(const g of tokens){
          if(g.kind==='text'){
            const prev=fragments.at(-1);
            if(prev&&prev.x+prev.width===px&&prev.style===g.style&&JSON.stringify(prev.request)===JSON.stringify(g.request)){prev.text+=g.text;prev.width+=g.advance;}
            else fragments.push({text:g.text,x:px,width:g.advance,request:g.request,style:g.style});
          }
          px+=g.advance;carets.push({offset:g.end,x:px});
        }
        if(carets.at(-1)!.offset!==end)carets.push({offset:end,x:px});
        page.lines.push({paragraphId:block.id,start,end,x,y,width,height,baseline:y+Math.floor((height-ascent-descent)/2)+ascent,carets,fragments});y+=height;
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
      space(style.spaceAfter??0);
    }
  }
  return freezeData({revision:snapshot.revision,contentId:snapshot.contentId,environmentId:provider.environmentId,pages,warnings:[...warnings].sort()});
}
type ReadLayout=DeepReadonly<WriterLayout>;
export function hitTestWriterPage(layout:ReadLayout,pageIndex:number,point:Point):WriterPosition|undefined{
  assert(Number.isFinite(point.x)&&Number.isFinite(point.y),'invalid','Invalid Writer hit point');
  const page=layout.pages[pageIndex];if(!page?.lines.length)return undefined;
  const distance=(line:DeepReadonly<WriterLine>)=>Math.max(line.y-point.y,point.y-line.y-line.height,0);
  let nearest=page.lines[0]!;
  for(const line of page.lines)if(distance(line)<distance(nearest))nearest=line;
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
    output.push({pageIndex,x:first.x,y:line.y,width:Math.max(0,last.x-first.x),height:line.height});
  }
  return output;
}
export function writerCaretBox(layout:ReadLayout,position:WriterPosition,preferredPage?:number):WriterSelectionBox|undefined{
  const found:WriterSelectionBox[]=[];
  for(let pageIndex=0;pageIndex<layout.pages.length;pageIndex++)for(const line of layout.pages[pageIndex]!.lines){
    if(line.paragraphId!==position.paragraphId)continue;
    const caret=line.carets.find(c=>c.offset===position.offset);
    if(caret)found.push({pageIndex,x:caret.x,y:line.y,width:0,height:line.height});
  }
  return found.find(b=>b.pageIndex===preferredPage)??found.at(-1);
}
