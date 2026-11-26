// SPDX-License-Identifier: MIT
/** Bounded, revision-scoped Writer layout. First profile: horizontal LTR text paragraphs,
 * direct/named text styles, explicit line/page breaks. Not full Word shaping or pagination.
 * Document coordinates are EMU. No derived geometry is persisted in the source model.
 */
import { assert, graphemeBoundaries, type Cancellation } from '@office-ex/contracts';
import { resolveTextStyle, writerEffectiveStyle, buildWriterNumberingMarkers, type WriterNumberingMarker, freezeData, paragraphBoundaries, OBJECT_REPLACEMENT, type DocumentSnapshot, type DeepReadonly, type TextStyle, type WriterParagraphRule, type WriterSection, type WriterPageStory, type Paragraph, type WriterDocument, type Table, type Inline } from '@office-ex/core';
import { measureText, type FontMetrics, type FontRequest, type Point } from './index.js';

export interface WriterPosition {paragraphId:string;offset:number}
export interface WriterCaret {offset:number;x:number}
export interface WriterFragment {text:string;x:number;width:number;request:FontRequest;style:DeepReadonly<TextStyle>}
export interface WriterVisualCluster {start:number;end:number;x:number;width:number}
export interface WriterNativeSlice {originX:number;boxWidth:number;sourceTop:number;paragraphWidth:number}
export interface WriterParagraphDecoration {x:number;width:number;rule:DeepReadonly<WriterParagraphRule>}
export interface WriterListMarkerBox {text:string;x:number;width:number;style:DeepReadonly<TextStyle>;request:FontRequest}
export interface WriterLine {paragraphId:string;start:number;end:number;x:number;y:number;width:number;height:number;baseline:number;carets:WriterCaret[];fragments:WriterFragment[];images:WriterImageBox[];visualClusters?:WriterVisualCluster[];native?:WriterNativeSlice;contentHeight?:number;bottomBorder?:WriterParagraphDecoration;listMarker?:WriterListMarkerBox}
export interface WriterTableCellBox {tableId:string;rowId:string;cellId:string;paragraphIds:string[];x:number;y:number;width:number;height:number;border?:DeepReadonly<NonNullable<Table['border']>>}
export interface WriterImageBox {paragraphId:string;imageId:string;resourceId:string;x:number;y:number;width:number;height:number}
export interface WriterRunningStoryBox {kind:'header'|'footer';storyId:string;paragraph:Paragraph;paragraphWidth:number;x:number;y:number;width:number;height:number;text:string;fragment?:WriterFragment;baseline?:number}
export interface WriterPage {displayPageNumber?:number;runningStories?:WriterRunningStoryBox[];sectionId:string;width:number;height:number;margins:DeepReadonly<WriterSection['page']['margins']>;lines:WriterLine[];tableCells:WriterTableCellBox[];images:WriterImageBox[]}
export interface WriterLayout {revision:number;contentId:string;environmentId:string;pages:WriterPage[];warnings:string[]}
/** Optional native shaping backend. All geometry is relative to the paragraph
 * content box in EMU. The display backend uses the identical native paragraph
 * projection clipped to the page's line ranges. */
export interface NativeParagraphShape {lines:readonly WriterLine[];warnings:readonly string[];spaceBefore:number;spaceAfter:number;clusters:number}
export interface WriterParagraphShaper {readonly environmentId:string;shape(model:DeepReadonly<WriterDocument>,paragraph:DeepReadonly<Paragraph>,width:number):NativeParagraphShape}
export interface WriterLayoutOptions {maxPages?:number;maxClusters?:number;cancellation?:Cancellation;paragraphShaper?:WriterParagraphShaper;numberingMarkers?:ReadonlyMap<string,WriterNumberingMarker>;skipRunningStories?:boolean}
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
  const style=writerEffectiveStyle(model,p);assert(style.paragraphDirection!=='rtl','unsupported','RTL paragraph needs a bidi layout adapter');const pieces:{text:string;style:DeepReadonly<TextStyle>;image?:ImageRun}[]=[];
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
/** The physical paragraph bounds stay independent of the first-line offset.
 * A hanging indent extends the first line into the paragraph's leading margin.
 * Explicit RTL uses the right edge; unspecified Word paragraph direction is LTR
 * for this geometry calculation. */
export function writerParagraphGeometry(style:DeepReadonly<TextStyle>,width:number){
 const left=Math.max(0,style.indent??0),right=style.rightIndent??0,first=style.firstLineIndent??0,rtl=style.paragraphDirection==='rtl';
 const hanging=Math.min(0,first),origin=left+(rtl?0:hanging),end=width-right-(rtl?hanging:0),bodyWidth=width-left-right;
 assert(origin>=0&&end<=width&&bodyWidth>0&&bodyWidth-Math.max(0,first)>0,'unsupported','Paragraph indentation exceeds its content box');
 return {left,right,first,rtl,origin,boxWidth:end-origin,bodyWidth,padLeft:left-origin,padRight:end-(width-right)};
}
/** Returns line boxes and logical caret coordinates. Unsupported content fails
 * explicitly; callers keep the source intact and show a localized preview warning.
 */
export function layoutWriter(snapshot:DocumentSnapshot,provider:FontMetrics,options:WriterLayoutOptions={}):DeepReadonly<WriterLayout>{
  assert(snapshot.document.kind==='writer','invalid','Expected Writer snapshot');
  const model=snapshot.document;
  const maxPages=budget(options.maxPages,100,500),maxClusters=budget(options.maxClusters,30000,100000);
  const warnings=new Set<string>(),pages:WriterPage[]=[];let clusterCount=0,blocks=0;const numberingMarkers=options.numberingMarkers??buildWriterNumberingMarkers(model);
  const check=()=>options.cancellation?.throwIfCancelled();
  const charge=(n:number)=>{clusterCount+=n;assert(clusterCount<=maxClusters,'limit','Writer layout cluster budget exceeded');};
  for(const section of model.sections){
    check();const spec=section.page,m=spec.margins;
    const usableWidth=spec.width-m.left-m.right,usableHeight=spec.height-m.top-m.bottom;
    assert(usableWidth>0&&usableHeight>0,'invalid','Page has no usable area');
    const makePage=():WriterPage=>{assert(pages.length<maxPages,'limit','Writer layout page budget exceeded');const p:WriterPage={sectionId:section.id,width:spec.width,height:spec.height,margins:m,lines:[],tableCells:[],images:[]};pages.push(p);return p;};
    let page=makePage(),y=m.top;
    const advancePage=()=>{page=makePage();y=m.top;};
    const bottom=spec.height-m.bottom,hasContent=()=>!!(page.lines.length||page.tableCells.length||page.images.length);

    type ReadBlock=DeepReadonly<WriterSection['blocks'][number]>;
    const styleCache=new WeakMap<DeepReadonly<Paragraph>,DeepReadonly<TextStyle>>();
    const effective=(p:DeepReadonly<Paragraph>)=>{
      let style=styleCache.get(p);
      if(!style){const id=p.styleId??model.defaultParagraphStyleId;style=writerEffectiveStyle(model,p);styleCache.set(p,style);}
      return style;
    };
    const layParagraph=(block:DeepReadonly<Paragraph>,originX:number,boxWidth:number,addSpace:(n:number)=>void,place:(line:WriterLine)=>void,previous?:ReadBlock,next?:ReadBlock):void=>{
      check();assert(++blocks<=10000,'limit','Writer layout block budget exceeded');
      const paraStyle=effective(block),g=writerParagraphGeometry(paraStyle,boxWidth),styleId=block.styleId??model.defaultParagraphStyleId;
      const sameStyle=(other:ReadBlock|undefined)=>other?.type==='paragraph'&&(other.styleId??model.defaultParagraphStyleId)===styleId;
      const before=paraStyle.contextualSpacing&&sameStyle(previous)?0:paraStyle.spaceBefore??0;
      const after=paraStyle.contextualSpacing&&sameStyle(next)?0:paraStyle.spaceAfter??0;
      const border=paraStyle.paragraphBottomBorder;
      const following=next?.type==='paragraph'?effective(next):undefined;
      const continues=border&&following?.paragraphBottomBorder&&JSON.stringify(border)===JSON.stringify(following.paragraphBottomBorder)&&(paraStyle.indent??0)===(following.indent??0)&&(paraStyle.rightIndent??0)===(following.rightIndent??0);
      let pending:WriterLine|undefined,firstOutput=true;
      const emitLine=(line:WriterLine)=>{
        if(firstOutput){firstOutput=false;const marker=numberingMarkers.get(block.id);
          if(marker){
            const measured=measureText(provider,marker.text,marker.style,paraStyle.paragraphDirection??'ltr'),advance=Math.ceil(measured.metrics.advance);
            const hanging=Math.min(marker.indent,Math.max(marker.hanging,advance+12700));if(hanging>marker.hanging)warnings.add('writer.list-marker-width-expanded');
            assert(advance<=hanging,'unsupported','List marker exceeds available leading margin');
            const left=paraStyle.paragraphDirection==='rtl'?originX+boxWidth-(paraStyle.rightIndent??0):originX+marker.indent-hanging;
            // For the bounded profile, RTL numbered paragraphs remain protected.
            assert(paraStyle.paragraphDirection!=='rtl','unsupported','RTL list marker placement is not in the current profile');
            const shift=marker.level.alignment==='end'?hanging-advance:marker.level.alignment==='center'?Math.round((hanging-advance)/2):0;
            line={...line,listMarker:{text:marker.text,x:left+shift,width:advance,style:marker.style,request:{...measured.request,family:measured.metrics.actualFamily}}};
          }
        }
        if(pending)place(pending);pending=line;
      };
      const finish=()=>{
        if(pending){
          if(border&&!continues){
            const height=pending.height+border.space+border.width;
            assert(height<=usableHeight,'unsupported','Paragraph rule exceeds page content height');
            pending={...pending,contentHeight:pending.height,height,bottomBorder:{x:originX+g.left,width:g.bodyWidth,rule:border}};
          }
          place(pending);pending=undefined;
        }
        addSpace(after);
      };
      if(options.paragraphShaper){
        const shaped=options.paragraphShaper.shape(model,block,boxWidth);
        assert(Number.isSafeInteger(shaped.clusters)&&shaped.clusters>=0,'invalid','Invalid native cluster count');
        charge(shaped.clusters);for(const warning of shaped.warnings)warnings.add(warning);
        addSpace(before);
        for(const line of shaped.lines){
          check();assert(line.paragraphId===block.id&&line.height>0&&line.height<=usableHeight&&line.native,'unsupported','Native line exceeds usable page height');
          const adjusted:WriterLine={...line,x:line.x+originX,carets:line.carets.map(c=>({...c,x:c.x+originX})),
            fragments:line.fragments.map(f=>({...f,x:f.x+originX})),images:line.images.map(i=>({...i,x:i.x+originX})),
            visualClusters:line.visualClusters?.map(c=>({...c,x:c.x+originX})),native:{...line.native,originX:line.native.originX+originX}};
          emitLine(adjusted);
        }
        finish();return;
      }
      const {glyphs,style,minimum}=glyphsFor(model,block,provider,warnings,charge,check);
      const indent=g.left;
      if((style.indent??0)<0)warnings.add('writer.negative-indent-limited');
      const available=g.bodyWidth;
      assert(available>0,'unsupported','Paragraph indent consumes page width');
      if(style.alignment==='justify')warnings.add('writer.justification-pending');
      addSpace(before);
      let firstLine=true;
      const lineWidth=()=>available-(firstLine?g.first:0);
      const emit=(tokens:Glyph[],start:number,end:number,forced:boolean)=>{
        const usable=lineWidth(),lineIndent=indent+(firstLine?g.first:0);
        check();
        const width=tokens.reduce((sum,g)=>sum+g.advance,0);
        const picturesOnly=tokens.length>0&&tokens.every(g=>g.kind==='image');
        let ascent=picturesOnly?0:Math.ceil((style.fontSize??152400)*0.8),descent=picturesOnly?0:Math.ceil((style.fontSize??152400)*0.2);
        for(const g of tokens){if(g.image){assert(g.advance<=usable,'unsupported','Inline image exceeds paragraph width');ascent=Math.max(ascent,g.ascent);}else{ascent=Math.max(ascent,g.ascent,Math.ceil(g.request.size*0.8));descent=Math.max(descent,g.descent,Math.ceil(g.request.size*0.2));}}
        const natural=Math.max(picturesOnly?0:minimum,ascent+descent);
        if(style.lineHeight!==undefined&&style.lineHeight<natural)warnings.add('writer.exact-lineheight-limited');
        const height=Math.max(natural,Math.ceil(natural*(style.lineMultiple??240)/240),style.lineHeight??0,style.minLineHeight??0);
        assert(height<=usableHeight,'unsupported','Line exceeds usable page height');
        const x=originX+lineIndent+(style.alignment==='center'?Math.max(0,Math.round((usable-width)/2)):style.alignment==='end'?Math.max(0,usable-width):0);
        if(width>usable)warnings.add('writer.line-overflow');
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
        emitLine({paragraphId:block.id,start,end,x,y:0,width,height,baseline,carets,fragments,images});firstLine=false;
      };
      let begin=0;let finalHard=false;
      while(begin<glyphs.length){
        check();let i=begin,width=0,lastBreak=-1;
        for(;i<glyphs.length;i++){
          const g=glyphs[i]!;
          if(g.kind==='break')break;
          if(width+g.advance>lineWidth()&&i>begin)break;
          width+=g.advance;
          if(breakBetween(g,glyphs[i+1]))lastBreak=i+1;
        }
        const hard=i<glyphs.length&&glyphs[i]!.kind==='break';
        const reachedEnd=i===glyphs.length;
        if(hard){const tokens=glyphs.slice(begin,i+1);emit(tokens,glyphs[begin]!.start,glyphs[i]!.end,false);begin=i+1;finalHard=begin===glyphs.length;}
        else{const end=reachedEnd?i:lastBreak>begin?lastBreak:i;assert(end>begin,'internal','Writer layout made no progress');emit(glyphs.slice(begin,end),glyphs[begin]!.start,glyphs[end-1]!.end,!reachedEnd&&lastBreak<=begin);begin=end;finalHard=false;}
      }
      if(!glyphs.length||finalHard){const end=glyphs.at(-1)?.end??0;emit([],end,end,false);}
      finish();
    };
    const placeLine=(line:WriterLine,top:number):void=>{
      const images=line.images.map(image=>({...image,y:image.y+top}));
      page.lines.push({...line,y:top,baseline:line.baseline+top,images});page.images.push(...images);
    };
    type ParagraphPlan={lines:WriterLine[];before:number;after:number;style:DeepReadonly<TextStyle>;height:number};
    function planParagraph(block:DeepReadonly<Paragraph>,previous?:ReadBlock,next?:ReadBlock):ParagraphPlan{
      const spaces:number[]=[],lines:WriterLine[]=[];
      layParagraph(block,m.left,usableWidth,n=>spaces.push(n),l=>lines.push(l),previous,next);
      assert(spaces.length===2&&lines.length>0,'internal','Incomplete paragraph plan');
      return {lines,before:spaces[0]!,after:spaces[1]!,style:effective(block),height:lines.reduce((n,l)=>n+l.height,0)};
    }
    // Count only complete line boxes. This also works for variable line
    // heights and inline images, not just fixed-height paragraphs.
    function firstChunkHeight(plan:ParagraphPlan):number{
      if(plan.style.keepLines&&plan.height<=usableHeight)return plan.height;
      if(plan.style.widowControl!==false&&plan.lines.length>1){
        // For a three-line paragraph there is no legal 2+2 split.
        if(plan.lines.length===3&&plan.height<=usableHeight)return plan.height;
        const firstTwo=plan.lines[0]!.height+plan.lines[1]!.height;
        if(firstTwo<=usableHeight)return firstTwo;
      }
      return plan.lines[0]!.height;
    }
    function reserve(height:number):boolean{
      if(height>usableHeight)return false;
      if(y+height>bottom&&hasContent())advancePage();
      return true;
    }
    function flowParagraph(plan:ParagraphPlan):void{
      const {lines,style}=plan;
      const minStart=firstChunkHeight(plan);
      if(y+plan.before+minStart>bottom&&hasContent())advancePage();
      // Large before/after spacing cannot generate empty overflow pages or
      // push the first visible line outside the usable area.
      const before=Math.min(plan.before,Math.max(0,bottom-y-minStart));
      if(before!==plan.before)warnings.add('writer.paragraph-spacing-limited');
      y+=before;
      if(style.keepLines&&plan.height>usableHeight)warnings.add('writer.keep-lines-limited');
      let at=0;
      while(at<lines.length){
        check();
        let fit=0,height=0;
        for(let i=at;i<lines.length&&y+height+lines[i]!.height<=bottom;i++){height+=lines[i]!.height;fit++;}
        if(!fit){
          assert(hasContent(),'unsupported','Line cannot fit on empty page');
          advancePage();continue;
        }
        const remaining=lines.length-at;
        if(fit<remaining&&style.widowControl!==false){
          // At least two lines on either side of a break when possible.
          // A fresh page is tried at most once for an impossible pair.
          if(fit===1){
            if(hasContent()&&lines[at]!.height+lines[at+1]!.height<=usableHeight){advancePage();continue;}
            warnings.add('writer.widow-control-limited');
          }else if(remaining-fit===1){
            if(fit>=3){fit--;}
            else if(hasContent()&&lines.slice(at).reduce((n,l)=>n+l.height,0)<=usableHeight){advancePage();continue;}
            else warnings.add('writer.widow-control-limited');
          }
        }
        for(let i=0;i<fit;i++){const line=lines[at++]!;placeLine(line,y);y+=line.height;}
        if(at<lines.length)advancePage();
      }
      // Do not create a page for trailing spacing alone.
      y=Math.min(bottom,y+plan.after);
    }
    type TablePlan={firstGroupHeight:number;place():void};
    function flowParagraphRun(plans:ParagraphPlan[],table?:TablePlan):void{
      for(let i=0;i<plans.length;){
        const start=plans[i]!;
        // A forced paragraph break takes precedence over keep-with-next.
        // A break at the very beginning of a page does not add a blank page.
        if(start.style.pageBreakBefore&&hasContent())advancePage();
        let end=i;
        while(end+1<plans.length&&plans[end]!.style.keepNext&&!plans[end+1]!.style.pageBreakBefore)end++;
        const withTable=end===plans.length-1&&!!table&&!!plans[end]!.style.keepNext;
        if(end===i&&!withTable){flowParagraph(start);i++;continue;}
        let height=0;
        for(let j=i;j<end;j++){const p=plans[j]!;height+=p.before+p.height+p.after;}
        const last=plans[end]!;
        height+=last.before+(withTable?last.height+last.after+table!.firstGroupHeight:firstChunkHeight(last));
        if(reserve(height)){
          for(let j=i;j<=end;j++)flowParagraph(plans[j]!);
        }else{
          warnings.add('writer.keep-next-limited');
          for(let j=i;j<=end;j++){
            const p=plans[j]!,following=plans[j+1];
            if(j<end&&following)reserve(p.before+p.height+p.after+following.before+firstChunkHeight(following));
            else if(withTable)reserve(p.before+p.height+p.after+table!.firstGroupHeight);
            flowParagraph(p);
          }
        }
        i=end+1;
      }
    }
    const planTable=(table:DeepReadonly<Table>):TablePlan=>{
      const pad=table.cellPadding??{top:38100,right:57150,bottom:38100,left:57150};
      const totalWidth=table.columnWidths.reduce((a,b)=>a+b,0);
      assert(totalWidth<=usableWidth,'unsupported','Table exceeds page content width');
      assert(table.columnWidths.every(w=>w>pad.left+pad.right),'unsupported','Table cell has no content width');
      warnings.add('writer.table-row-pagination');
      // Lay out only owning cells. A row-spanning anchor is a single
      // rectangle; covered slots have neither text nor interior borders.
      const xPositions=[m.left];
      for(const width of table.columnWidths)xPositions.push(xPositions.at(-1)!+width);
      type LaidCell={box:WriterTableCellBox;lines:WriterLine[];height:number;rowSpan:number};
      const laidRows:LaidCell[][]=table.rows.map(()=>[]);
      const heights=table.rows.map(()=>Math.max(1,pad.top+pad.bottom));
      for(const [ri,row] of table.rows.entries()){
        check();assert(++blocks<=10000,'limit','Writer table row budget exceeded');
        for(let col=0;col<row.cells.length;col++){
          const cell=row.cells[col]!;if(cell.coveredBy)continue;
          const rowSpan=cell.merge?.rows??1,colSpan=cell.merge?.columns??1;
          const x=xPositions[col]!,width=xPositions[col+colSpan]!-x,lines:WriterLine[]=[];let localY=pad.top;
          const addSpace=(n:number)=>{localY+=n;assert(localY+pad.bottom<=usableHeight,'unsupported','Table cell exceeds page content height');};
          const place=(line:WriterLine)=>{
            assert(line.x+line.width<=x+width-pad.right,'unsupported','Table cell line exceeds content width');
            lines.push({...line,y:localY,baseline:localY+line.baseline,images:line.images.map(image=>({...image,y:image.y+localY}))});addSpace(line.height);
          };
          for(const [childIndex,child] of cell.blocks.entries()){
            assert(child.type==='paragraph','unsupported','Nested tables and page breaks within cells unsupported');
            if(effective(child).pageBreakBefore)warnings.add('writer.cell-page-break-before-limited');
            layParagraph(child,x+pad.left,width-pad.left-pad.right,addSpace,place,cell.blocks[childIndex-1],cell.blocks[childIndex+1]);
          }
          const h=localY+pad.bottom;
          assert(h>0&&h<=usableHeight,'unsupported','Table cell exceeds page content height');
          if(rowSpan===1)heights[ri]=Math.max(heights[ri]!,h);
          const box:WriterTableCellBox={tableId:table.id,rowId:row.id,cellId:cell.id,paragraphIds:cell.blocks.map(b=>b.id),x,y:0,width,height:0,...(table.border?{border:table.border}:{})};
          laidRows[ri]!.push({box,lines,height:h,rowSpan});
        }
      }
      // Satisfy all spanning height constraints without shrinking rows.
      // Share deficits evenly; subsequent spans can only increase heights.
      for(const [ri,cells] of laidRows.entries())for(const item of cells){
        if(item.rowSpan===1)continue;
        const current=heights.slice(ri,ri+item.rowSpan).reduce((a,b)=>a+b,0),extra=Math.max(0,item.height-current);
        if(extra)for(let k=ri;k<ri+item.rowSpan;k++)heights[k]!+=extra/item.rowSpan;
      }
      const offsets=[0];for(const h of heights)offsets.push(offsets.at(-1)!+h);
      // A vertical merge connects rows transitively. Keep each connected
      // row group on a single page; fail explicitly instead of clipping
      // or splitting a merged cell across pages.
      const groups:{start:number;end:number;height:number}[]=[];
      for(let start=0;start<table.rows.length;){
        check();let end=start+1;
        for(let r=start;r<end;r++)for(const item of laidRows[r]!)end=Math.max(end,r+item.rowSpan);
        const height=offsets[end]!-offsets[start]!;
        assert(height<=usableHeight,'unsupported','Merged table row group exceeds page content height');
        groups.push({start,end,height});start=end;
      }
      return {firstGroupHeight:groups[0]!.height,place(){
      for(const {start,end,height:groupHeight}of groups){
        check();if(y+groupHeight>spec.height-m.bottom)advancePage();
        for(let r=start;r<end;r++){
          const top=y+offsets[r]!-offsets[start]!;
          for(const item of laidRows[r]!){
            page.tableCells.push({...item.box,y:top,height:offsets[r+item.rowSpan]!-offsets[r]!});
            for(const line of item.lines){const images=line.images.map(image=>({...image,y:image.y+top}));page.lines.push({...line,y:top+line.y,baseline:top+line.baseline,images});page.images.push(...images);}
          }
        }
        y+=groupHeight;
      }
      }};
    };
    for(let blockIndex=0;blockIndex<section.blocks.length;){
      check();const block=section.blocks[blockIndex]!;
      if(block.type==='page-break'){assert(++blocks<=10000,'limit','Writer layout block budget exceeded');advancePage();blockIndex++;continue;}
      if(block.type==='table'){planTable(block).place();blockIndex++;continue;}
      const plans:ParagraphPlan[]=[];
      while(blockIndex<section.blocks.length){
        const p=section.blocks[blockIndex]!;if(p.type!=='paragraph')break;
        plans.push(planParagraph(p,section.blocks[blockIndex-1],section.blocks[blockIndex+1]));blockIndex++;
      }
      const next=section.blocks[blockIndex];
      const table=next?.type==='table'&&plans.at(-1)?.style.keepNext?planTable(next):undefined;
      flowParagraphRun(plans,table);
      if(table){table.place();blockIndex++;}
    }
  }
  const decorated=options.skipRunningStories?pages:layoutRunningStories(model,pages,provider,options,warnings);
  return freezeData({revision:snapshot.revision,contentId:snapshot.contentId,environmentId:options.paragraphShaper?.environmentId??provider.environmentId,pages:decorated,warnings:[...warnings].sort()});
}
/** Recompute running decorations only after the complete page sequence is
 * assembled. A reused section body may have a different absolute page
 * number when an earlier section grows. Stories never enter body caret
 * or clipboard geometry. */
export function layoutRunningStories(model:DeepReadonly<WriterDocument>,pages:readonly DeepReadonly<WriterPage>[],provider:FontMetrics,options:WriterLayoutOptions,warnings:Set<string>):readonly DeepReadonly<WriterPage>[] {
  if(!model.pageStories&&!model.sections.some(s=>s.pageNumberStart!==undefined))return pages;
  const refs=new Map<string,{header?:string;footer?:string}>();
  let inherited:{header?:string;footer?:string}={};
  const sections=new Map(model.sections.map(s=>[s.id,s]));
  for(const section of model.sections){inherited={...inherited,...section.storyRefs};refs.set(section.id,inherited);}
  let previous='',number=0,clusters=0;
  return pages.map((page,index)=>{
    options.cancellation?.throwIfCancelled();
    const section=sections.get(page.sectionId);assert(section,'invalid','Running story section not found');
    number=previous!==page.sectionId&&section.pageNumberStart!==undefined?section.pageNumberStart:number+1;previous=page.sectionId;
    const decorations:WriterRunningStoryBox[]=[];
    for(const kind of ['header','footer'] as const){
      const id=refs.get(section.id)?.[kind];if(!id)continue;
      const story=model.pageStories?.[id];assert(story&&story.kind===kind,'invalid','Missing running story');
      const text=story.text+(story.pageNumber?String(number)+(story.after??''):'');if(!text)continue;
      clusters+=graphemeBoundaries(text).length-1;assert(clusters<=100000,'limit','Running story layout cluster budget exceeded');
      const p:Paragraph={id:'running-'+index+'-'+kind,type:'paragraph',runs:[{id:'running-'+index+'-'+kind+'-run',type:'text',text,...(story.runStyle?{style:story.runStyle}:{})}],style:{...(story.styleId?{spaceAfter:0,lineMultiple:240}:{}),...(story.alignment?{alignment:story.alignment}:{})}};
      const style=writerEffectiveStyle(model,p);
      assert(!style.writerList&&!style.paragraphBottomBorder,'unsupported','Numbered or decorated running story paragraph is unsupported');
      const width=page.width-page.margins.left-page.margins.right,geometry=writerParagraphGeometry(style,width);
      const before=style.spaceBefore??0,after=style.spaceAfter??0;
      let height:number,x=page.margins.left+geometry.origin,boxWidth=geometry.boxWidth,fragment:WriterFragment|undefined,baseline:number|undefined;
      if(options.paragraphShaper){
        const shaped=options.paragraphShaper.shape(model,p,width);
        assert(shaped.lines.length===1&&shaped.lines[0]?.native,'unsupported','Running story must fit one line');
        const l=shaped.lines[0]!,native=l.native!;height=l.height;x=page.margins.left+native.originX;boxWidth=native.boxWidth;
        for(const warning of shaped.warnings)warnings.add(warning);
      }else{
        const run=resolveTextStyle(style,model.styles,[],story.runStyle??{}).style;
        const measured=requestFor(provider,text,run,warnings),advance=measured.metrics.advance;
        assert(advance<=geometry.bodyWidth-Math.max(0,geometry.first),'unsupported','Running story must fit one line');
        height=Math.ceil(Math.max(measured.metrics.ascent+measured.metrics.descent,(run.fontSize??152400)*1.55,style.lineHeight??0,style.minLineHeight??0));
        const available=geometry.bodyWidth-Math.max(0,geometry.first),offset=style.alignment==='center'?(available-advance)/2:style.alignment==='end'?available-advance:0;
        fragment={text,x:x+geometry.padLeft+Math.max(0,geometry.first)+offset,width:advance,request:measured.request,style:run};
        baseline=height/2+(measured.metrics.ascent-measured.metrics.descent)/2;
      }
      const distance=section.pageExtras?.[kind==='header'?'headerDistance':'footerDistance']??0,total=before+height+after;
      const margin=page.margins[kind==='header'?'top':'bottom'];assert(distance+total<=margin,'unsupported','Running story exceeds its page margin');
      const y=kind==='header'?distance+before:page.height-distance-after-height;
      decorations.push({kind,storyId:id,paragraph:p,paragraphWidth:width,x,y,width:boxWidth,height,text,...(fragment?{fragment,baseline:y+baseline!}:{})});
    }
    return {...page,displayPageNumber:number,...(decorations.length?{runningStories:decorations}:{})};
  });
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
      for(const c of selected){const prev=output.at(-1);if(prev&&prev.pageIndex===pageIndex&&prev.y===line.y&&prev.x+prev.width>=c.x-100){prev.width=Math.max(prev.width,c.x+c.width-prev.x);}else output.push({pageIndex,x:c.x,y:line.y,width:Math.max(1,c.width),height:line.contentHeight??line.height});}
    }else output.push({pageIndex,x:Math.min(first.x,last.x),y:line.y,width:Math.abs(last.x-first.x),height:line.contentHeight??line.height});
  }
  return output;
}
export function writerCaretBox(layout:ReadLayout,position:WriterPosition,preferredPage?:number,preferredX?:number):WriterSelectionBox|undefined{
  const found:WriterSelectionBox[]=[];
  for(let pageIndex=0;pageIndex<layout.pages.length;pageIndex++)for(const line of layout.pages[pageIndex]!.lines){
    if(line.paragraphId!==position.paragraphId)continue;
    const caret=line.carets.find(c=>c.offset===position.offset);
    if(caret)found.push({pageIndex,x:caret.x,y:line.y,width:0,height:line.contentHeight??line.height});
  }
  const candidates=found.filter(b=>b.pageIndex===preferredPage);const pool=candidates.length?candidates:found;
  if(preferredX!==undefined&&pool.length)return pool.reduce((a,b)=>Math.abs(a.x-preferredX)<=Math.abs(b.x-preferredX)?a:b);
  return pool.at(-1);
}
