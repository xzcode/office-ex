// SPDX-License-Identifier: MIT
/** Page editing is a projection of the immutable model and page layout.
 * The small native textarea is only an IME/clipboard input bridge at the
 * actual caret; neither DOM mutations nor the rendered page are source data.
 */
import {bundledFontFamily} from '@office-ex/pdf';
import {assert,OfficeError,EMU_PER_CSS_PIXEL,EMU_PER_POINT} from '@office-ex/contracts';
import {paragraphBoundaries,inlineText,resolveTextStyle,type DeepReadonly,type WriterDocument,type Paragraph,type TextStyle,type DocumentSnapshot} from '@office-ex/core';
import {locateParagraph,resolveWriterRange,replaceWriterRange,replaceParagraphRange,formatWriterRange,setParagraphAlignment,copyWriterFragment,pasteWriterFragment,WRITER_CLIPBOARD_TYPE,type WriterRange,type WriterPoint} from '@office-ex/writer';
import {WriterPaginationEngine,hitTestWriterPage,writerCaretBox,writerSelectionBoxes,type WriterLayout,type WriterLine} from '@office-ex/render';
import {CanvasFontMetrics} from './font-metrics.js';
import {BrowserParagraphShaper} from './native-writer-shaper.js';
import type {WriterHost,WriterViewState} from './writer-workspace.js';

const PX=EMU_PER_CSS_PIXEL,css=(n:number)=>n/PX+'px';
export interface NativePageCache{shaper:BrowserParagraphShaper;engine:WriterPaginationEngine;metrics:CanvasFontMetrics}
const node=<K extends keyof HTMLElementTagNameMap>(tag:K,text='',className?:string):HTMLElementTagNameMap[K]=>{const n=document.createElement(tag);n.textContent=text;if(className)n.className=className;return n;};
function doc(snapshot:DocumentSnapshot):DeepReadonly<WriterDocument>{assert(snapshot.document.kind==='writer','invalid','Expected Writer');return snapshot.document;}
function firstPoint(model:DeepReadonly<WriterDocument>):WriterPoint|undefined{
  for(const section of model.sections)for(const block of section.blocks){
    if(block.type==='paragraph')return {paragraphId:block.id,offset:0};
    if(block.type==='table')for(const row of block.rows)for(const cell of row.cells)for(const p of cell.blocks)if(p.type==='paragraph')return {paragraphId:p.id,offset:0};
  }
  return undefined;
}
function collapsed(range:WriterRange):boolean{return range.anchor.paragraphId===range.focus.paragraphId&&range.anchor.offset===range.focus.offset;}
function color(style:DeepReadonly<TextStyle>):string{const c=style.color;return c?.kind==='rgba'?'rgba('+[c.red,c.green,c.blue,c.alpha/255].join(',')+')':'#20242c';}
function at(element:HTMLElement,x:number,y:number,width?:number,height?:number):void{element.style.position='absolute';element.style.left=css(x);element.style.top=css(y);if(width!==undefined)element.style.width=css(width);if(height!==undefined)element.style.height=css(height);}
export function nativeWriterPage(host:WriterHost,state:WriterViewState):HTMLElement{
  const root=node('section','','writer-native-workspace'),toolbar=node('div','','writer-controls'),pagebar=node('div','','writer-pagebar'),viewport=node('div','','writer-native-viewport'),paper=node('div','','writer-native-paper'),info=node('p','','muted');
  paper.id='writer-native-page';paper.setAttribute('aria-label',host.t('writerPagePreview'));paper.dir='ltr';
  const input=node('textarea','','writer-native-input');input.id='writer-page-input';input.setAttribute('aria-label',host.t('writerPageEdit'));input.autocomplete='off';input.spellcheck=false;input.rows=1;input.dir='auto';
  const overlays=node('div','','writer-native-overlays');overlays.setAttribute('aria-hidden','true');
  const cache=state.nativeCache??={shaper:new BrowserParagraphShaper(),engine:new WriterPaginationEngine(),metrics:new CanvasFontMetrics()};
  let layout:DeepReadonly<WriterLayout>|undefined,range=state.pageRange,preferredX:number|undefined,composition:{range:WriterRange;revision:number}|undefined,dragging=false,refreshing=false;
  let revision=-1;
  const t=host.t;
  const button=(key:Parameters<typeof t>[0],id:string,fn:()=>void)=>{const b=node('button',t(key));b.id=id;b.type='button';b.onmousedown=e=>e.preventDefault();b.onclick=()=>safe(fn);return b;};
  const safe=(fn:()=>void)=>{try{fn();}catch(error){host.error(error);}};
  function normalize():void{
    try{if(range){resolveWriterRange(host.snapshot(),range);return;}}catch{/* stale selection after undo */}
    const first=firstPoint(doc(host.snapshot()));range=first?{anchor:first,focus:first}:undefined;
  }
  function store():void{
    state.pageRange=range;
    if(range){state.paragraphId=range.focus.paragraphId;state.start=state.end=range.focus.offset;}
  }
  function move(point:WriterPoint,extend=false,follow=true,x?:number):void{
    const candidate={anchor:extend&&range?range.anchor:point,focus:point};
    resolveWriterRange(host.snapshot(),candidate);range=candidate;preferredX=x;store();
    if(follow&&layout){const c=writerCaretBox(layout,point,state.pageIndex,x);if(c&&c.pageIndex!==state.pageIndex){state.pageIndex=c.pageIndex;draw();return;}}
    highlight();
  }
  function caretStyle():DeepReadonly<TextStyle>{
    const d=doc(host.snapshot());if(!range)return d.defaultTextStyle;
    const p=locateParagraph(host.snapshot(),range.focus.paragraphId).paragraph,named=p.styleId??d.defaultParagraphStyleId,base=resolveTextStyle(d.defaultTextStyle,d.styles,named?[named]:[],p.style??{}).style;let cursor=0;
    for(const run of p.runs){const end=cursor+inlineText(run).length;if(run.type==='text'&&range.focus.offset<=end)return resolveTextStyle(base,d.styles,run.styleId?[run.styleId]:[],run.style??{}).style;cursor=end;}
    return base;
  }
  function highlight():void{
    if(!layout||layout.revision!==host.snapshot().revision||!paper.isConnected)return;
    overlays.replaceChildren();
    if(!range)return;
    const caret=writerCaretBox(layout,range.focus,state.pageIndex,preferredX);
    if(!collapsed(range)){
      for(const seg of resolveWriterRange(host.snapshot(),range).segments)for(const rect of writerSelectionBoxes(layout,seg.paragraphId,seg.start,seg.end))if(rect.pageIndex===state.pageIndex){
        const box=node('div','','writer-native-selection');at(box,rect.x,rect.y,Math.max(rect.width,PX),rect.height);overlays.append(box);
      }
    }else if(caret?.pageIndex===state.pageIndex){
      const mark=node('div','','writer-native-caret');at(mark,caret.x,caret.y,Math.round(1.5*PX),caret.height);overlays.append(mark);
    }
    if(caret?.pageIndex===state.pageIndex){
      const style=caretStyle();at(input,caret.x,caret.y,Math.max(1,style.fontSize??152400),caret.height);
      input.style.fontFamily=bundledFontFamily(style.fontFamily,style.language);input.style.fontSize=css(style.fontSize??152400);input.style.color=color(style);input.dir=style.paragraphDirection??'auto';
      input.hidden=false;
    }else input.hidden=true;
    if(!composition){
      const style=caretStyle();
      if(document.activeElement!==size)size.value=String((style.fontSize??152400)/EMU_PER_POINT);
      if(document.activeElement!==family)family.value=style.fontFamily??'sans-serif';
    }
  }
  const exit=button('writerExitPageEdit','writer-exit-page-edit',()=>{
    if(host.composing())return;state.pageMode=false;state.focus=true;store();host.render();
  });
  const previous=button('writerPreviousPage','writer-native-previous',()=>{if(!composition){state.pageIndex=Math.max(0,state.pageIndex-1);draw();}});
  const next=button('writerNextPage','writer-native-next',()=>{if(!composition&&layout){state.pageIndex=Math.min(layout.pages.length-1,state.pageIndex+1);draw();}});
  const pageNumber=node('span');
  pagebar.append(previous,pageNumber,next);
  toolbar.append(exit);
  function selectedFormat(style:TextStyle):void{
    if(!range||collapsed(range)||host.composing())return;
    if(host.execute(formatWriterRange(host.snapshot(),range,style,crypto.randomUUID())))refresh(true);
  }
  for(const [key,label] of [['bold','writerBold'],['italic','writerItalic'],['underline','writerUnderline']] as const){
    toolbar.append(button(label,'writer-page-'+key,()=>{
      if(!range||collapsed(range))return;
      const model=doc(host.snapshot());let all=true,count=0;
      for(const seg of resolveWriterRange(host.snapshot(),range).segments){
        const p=seg.paragraph,named=p.styleId??model.defaultParagraphStyleId,base=resolveTextStyle(model.defaultTextStyle,model.styles,named?[named]:[],p.style??{}).style;
        let at=0;for(const run of p.runs){const end=at+inlineText(run).length;if(run.type==='text'&&at<seg.end&&end>seg.start){count++;all=all&&resolveTextStyle(base,model.styles,run.styleId?[run.styleId]:[],run.style??{}).style[key]===true;}at=end;}
      }
      if(count)selectedFormat({[key]:!all});
    }));
  }
  const field=(key:Parameters<typeof t>[0],id:string,type:string)=>{const label=node('label','','writer-format-field'),input=node('input');input.type=type;input.id=id;label.append(node('span',t(key)),input);toolbar.append(label);return input;};
  const family=field('writerFontFamily','writer-page-font','text');family.maxLength=512;family.onchange=()=>safe(()=>selectedFormat({fontFamily:family.value.trim()}));
  const size=field('writerFontSize','writer-page-size','number');size.min='1';size.max='1638';size.step='0.5';size.onchange=()=>safe(()=>{const n=size.valueAsNumber;assert(Number.isFinite(n)&&n>=1&&n<=1638&&Number.isInteger(n*2),'invalid','Invalid font size');selectedFormat({fontSize:n*EMU_PER_POINT});});
  for(const [align,label] of [['start','writerAlignLeft'],['center','writerAlignCenter'],['end','writerAlignRight']] as const)toolbar.append(button(label,'writer-page-align-'+align,()=>{if(range&&!host.composing()&&host.execute(setParagraphAlignment(host.snapshot(),range.focus.paragraphId,align)))refresh(true);}));
  function draw():void{
    if(!layout||!root.isConnected||layout.revision!==host.snapshot().revision)return;
    const model=doc(host.snapshot());state.pageIndex=Math.max(0,Math.min(layout.pages.length-1,state.pageIndex));const page=layout.pages[state.pageIndex]!;
    paper.style.width=css(page.width);paper.style.height=css(page.height);
    const content:HTMLElement[]=[];
    for(const cell of page.tableCells){
      const border=node('div','','writer-native-cell');at(border,cell.x,cell.y,cell.width,cell.height);
      if(cell.border){border.style.border=css(cell.border.width)+' solid '+color({color:cell.border.color});}content.push(border);
    }
    // Recreate full native paragraph projection and clip exactly at the
    // measured line boundaries. A slice never reshapes only the selected
    // substring: paragraph context, joining and bidi order remain identical.
    const slices=new Map<string,DeepReadonly<WriterLine>[]>();
    for(const line of page.lines){if(!line.native)continue;const list=slices.get(line.paragraphId)??[];list.push(line);slices.set(line.paragraphId,list);}
    for(const [id,lines] of slices){
      lines.sort((a,b)=>a.y-b.y);const p=locateParagraph(host.snapshot(),id).paragraph;
      let i=0;
      while(i<lines.length){
        const a=lines[i]!,native=a.native!;let j=i+1;
        while(j<lines.length&&lines[j]!.native!.originX===native.originX&&Math.abs(lines[j]!.y-(lines[j-1]!.y+lines[j-1]!.height))<=1&&Math.abs(lines[j]!.native!.sourceTop-(lines[j-1]!.native!.sourceTop+lines[j-1]!.height))<=1)j++;
        const last=lines[j-1]!,clip=node('div','','writer-native-slice');at(clip,native.originX,a.y,native.boxWidth,last.y+last.height-a.y);
        clip.style.overflow='hidden';const projection=cache.shaper.render(model,p,native.paragraphWidth);projection.style.position='absolute';projection.style.top=css(-native.sourceTop);projection.style.left='0';clip.append(projection);content.push(clip);i=j;
      }
    }
    paper.replaceChildren(...content,overlays,input);
    pageNumber.textContent=t('writerPageNumber',{current:host.number(state.pageIndex+1),total:host.number(layout.pages.length)});
    previous.disabled=state.pageIndex===0;next.disabled=state.pageIndex===layout.pages.length-1;
    paper.hidden=false;root.dataset.pages=String(layout.pages.length);root.dataset.measured=String(cache.shaper.measurements);root.dataset.reused=String(cache.engine.stats.pagesReused);root.dataset.revision=String(revision);
    info.textContent=t('writerNativePageNote');if(layout.warnings.some(w=>w!=='writer.native-browser-line-metrics'&&w!=='writer.table-row-pagination'))info.textContent+=' '+t('writerLayoutLimits');
    highlight();
  }
  function refresh(focus=false):void{
    if(refreshing||!root.isConnected||host.composing())return;refreshing=true;
    try{
      const snapshot=host.snapshot();normalize();store();
      layout=cache.engine.layout(snapshot,cache.metrics,{paragraphShaper:cache.shaper,maxPages:100,maxClusters:20000});revision=snapshot.revision;
      if(range){const caret=writerCaretBox(layout,range.focus,state.pageIndex,preferredX);if(caret)state.pageIndex=caret.pageIndex;}
      draw();if(focus&&!input.hidden)input.focus({preventScroll:true});
    }catch(error){layout=undefined;paper.hidden=true;info.textContent=t('writerPreviewUnavailable');host.error(error);}
    finally{refreshing=false;}
  }
  function replace(text:string,soft=false):void{
    if(!range)return;
    const snapshot=host.snapshot();
    const resolved=resolveWriterRange(snapshot,range);
    if(soft){
      assert(resolved.first.block===resolved.last.block,'unsupported','Soft break requires one paragraph');
      if(host.execute(replaceParagraphRange(snapshot,resolved.first.paragraph.id,resolved.start,resolved.end,'\n',crypto.randomUUID()))){const point={paragraphId:resolved.first.paragraph.id,offset:resolved.start+1};range={anchor:point,focus:point};preferredX=undefined;refresh(true);}return;
    }
    const result=replaceWriterRange(snapshot,range,text,crypto.randomUUID());
    if(host.execute(result.command)){range={anchor:result.caret,focus:result.caret};preferredX=undefined;refresh(true);}
  }
  function erase(direction:'left'|'right'):void{
    if(!range)return;if(!collapsed(range)){replace('');return;}
    const loc=locateParagraph(host.snapshot(),range.focus.paragraphId),offset=range.focus.offset,bounds=paragraphBoundaries(loc.paragraph),i=bounds.indexOf(offset);
    assert(i>=0,'invalid','Invalid caret');const adjacent=direction==='left'?bounds[i-1]:bounds[i+1];
    let other:WriterPoint|undefined;
    if(adjacent!==undefined)other={paragraphId:loc.paragraph.id,offset:adjacent};
    else{
      const block=loc.siblings[loc.block+(direction==='left'?-1:1)];
      if(block?.type==='paragraph')other={paragraphId:block.id,offset:direction==='left'?paragraphBoundaries(block).at(-1)!:0};
    }
    if(other){range={anchor:range.focus,focus:other};replace('');}
  }
  function moveVisual(key:string,extend:boolean):void{
    if(!range||!layout)return;
    const position=range.focus,box=writerCaretBox(layout,position,state.pageIndex,preferredX);
    if(!box)return;
    const page=layout.pages[box.pageIndex]!;
    const current=page.lines.find(l=>l.paragraphId===position.paragraphId&&l.y===box.y&&l.carets.some(c=>c.offset===position.offset));
    if(!current)return;
    let point:WriterPoint|undefined,x:number|undefined;
    if(key==='ArrowLeft'||key==='ArrowRight'||key==='Home'||key==='End'){
      const sorted=[...current.carets].sort((a,b)=>a.x-b.x||a.offset-b.offset);
      const here=sorted.reduce((best,c,i)=>c.offset===position.offset&&Math.abs(c.x-box.x)<Math.abs((sorted[best]?.x??Infinity)-box.x)?i:best,sorted.findIndex(c=>c.offset===position.offset));
      let target=key==='Home'?sorted[0]:key==='End'?sorted.at(-1):sorted[here+(key==='ArrowLeft'?-1:1)];
      if(target){point={paragraphId:current.paragraphId,offset:target.offset};x=target.x;}
      else{
        const loc=locateParagraph(host.snapshot(),position.paragraphId),bounds=paragraphBoundaries(loc.paragraph),index=bounds.indexOf(position.offset),n=bounds[index+(key==='ArrowLeft'?-1:1)];
        if(n!==undefined)point={paragraphId:position.paragraphId,offset:n};
        else{const block=loc.siblings[loc.block+(key==='ArrowLeft'?-1:1)];if(block?.type==='paragraph')point={paragraphId:block.id,offset:key==='ArrowLeft'?paragraphBoundaries(block).at(-1)!:0};}
      }
    }else{
      const all=layout.pages.flatMap((p,pageIndex)=>p.lines.map(line=>({line,pageIndex})));
      const index=all.findIndex(v=>v.pageIndex===box.pageIndex&&v.line===current),step=key==='ArrowUp'?-1:1;
      for(let i=index+step;i>=0&&i<all.length;i+=step){
        const target=all[i]!;if(target.pageIndex===box.pageIndex&&target.line.y===current.y)continue;
        const c=target.line.carets.reduce((best,c)=>Math.abs(c.x-box.x)<Math.abs(best.x-box.x)?c:best,target.line.carets[0]!);
        point={paragraphId:target.line.paragraphId,offset:c.offset};x=c.x;break;
      }
    }
    if(point)move(point,extend,true,x);
  }
  function pointAt(event:PointerEvent):{point:WriterPoint;x:number}|undefined{
    if(!layout||layout.revision!==host.snapshot().revision)return;
    const p=layout.pages[state.pageIndex]!,rect=paper.getBoundingClientRect(),x=(event.clientX-rect.left)/rect.width*p.width,y=(event.clientY-rect.top)/rect.height*p.height;
    const point=hitTestWriterPage(layout,state.pageIndex,{x,y});return point?{point,x}:undefined;
  }
  paper.onpointerdown=event=>{
    if(event.button!==0||host.composing()||event.target===input)return;
    event.preventDefault();safe(()=>{const hit=pointAt(event);if(!hit)return;move(hit.point,event.shiftKey,false,hit.x);dragging=true;paper.setPointerCapture(event.pointerId);if(!input.hidden)input.focus({preventScroll:true});});
  };
  paper.onpointermove=event=>{if(!dragging||host.composing())return;const hit=pointAt(event);if(hit)try{move(hit.point,true,false,hit.x);}catch{/* Structural boundary: retain the last valid selection. */}};
  paper.onpointerup=()=>{dragging=false;};paper.onpointercancel=()=>{dragging=false;};
  input.addEventListener('compositionstart',()=>{if(range){composition={range,revision:host.snapshot().revision};host.setComposing(true);input.classList.add('composing');}});
  input.addEventListener('compositionend',event=>{
    const saved=composition;composition=undefined;host.setComposing(false);input.classList.remove('composing');const text=input.value||event.data;input.value='';
    if(!saved)return;safe(()=>{assert(host.snapshot().revision===saved.revision,'conflict','Document changed during IME input');range=saved.range;if(text)replace(text);});
  });
  input.oninput=event=>{if(host.composing()||(event as InputEvent).isComposing)return;const text=input.value;input.value='';if(text)safe(()=>replace(text));};
  input.onkeydown=event=>{
    if(host.composing()||event.isComposing||event.key==='Process')return;
    if(event.ctrlKey||event.metaKey){
      if(event.key.toLowerCase()==='a'){event.preventDefault();safe(()=>{
        if(!range)return;const loc=locateParagraph(host.snapshot(),range.focus.paragraphId);
        let first=loc.block,last=loc.block;
        while(first>0&&loc.siblings[first-1]!.type==='paragraph')first--;
        while(last+1<loc.siblings.length&&loc.siblings[last+1]!.type==='paragraph')last++;
        const a=loc.siblings[first]!,b=loc.siblings[last]!;assert(a.type==='paragraph'&&b.type==='paragraph','internal','Invalid paragraph range');
        range={anchor:{paragraphId:a.id,offset:0},focus:{paragraphId:b.id,offset:paragraphBoundaries(b).at(-1)!}};store();highlight();
      });}
      return;
    }
    if(['ArrowLeft','ArrowRight','ArrowUp','ArrowDown','Home','End'].includes(event.key)){event.preventDefault();safe(()=>moveVisual(event.key,event.shiftKey));}
    else if(event.key==='Backspace'||event.key==='Delete'){event.preventDefault();safe(()=>erase(event.key==='Backspace'?'left':'right'));}
    else if(event.key==='Enter'){event.preventDefault();safe(()=>replace('\n',event.shiftKey));}
    else if(event.key==='Tab'){event.preventDefault();safe(()=>replace('\t'));}
  };
  for(const kind of ['copy','cut'] as const)input.addEventListener(kind,event=>{
    if(host.composing()||!range||collapsed(range))return;event.preventDefault();
    safe(()=>{const content=copyWriterFragment(host.snapshot(),range!);assert(event.clipboardData,'unsupported','Clipboard unavailable');event.clipboardData.setData('text/plain',content.plain);event.clipboardData.setData(WRITER_CLIPBOARD_TYPE,content.rich);if(kind==='cut')replace('');});
  });
  input.addEventListener('paste',event=>{
    if(host.composing()||!range)return;event.preventDefault();safe(()=>{
      const data=event.clipboardData;if(!data)return;
      const rich=data.getData(WRITER_CLIPBOARD_TYPE);
      if(rich){const r=pasteWriterFragment(host.snapshot(),range!,rich,crypto.randomUUID());if(host.execute(r.command)){range={anchor:r.caret,focus:r.caret};preferredX=undefined;refresh(true);}}
      else replace(data.getData('text/plain'));
    });
  });
  viewport.append(paper);root.append(toolbar,pagebar,viewport,info);
  requestAnimationFrame(()=>refresh(true));return root;
}

