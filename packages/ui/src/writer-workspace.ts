// SPDX-License-Identifier: MIT
import { EMU_PER_CSS_PIXEL, EMU_PER_POINT, graphemeBoundaries, OfficeError } from '@office-ex/contracts';
import { resolveTextStyle, type DocumentSnapshot, type DocumentCommand, type DeepReadonly, type TextStyle, type WriterDocument } from '@office-ex/core';
import { paragraphText, replaceParagraphRange, splitParagraphRange, joinParagraphs, formatParagraphRange, setParagraphAlignment, setParagraphNamedStyle, setCharacterNamedStyle, insertPageBreakAfter, paragraphEditDiff } from '@office-ex/writer';
import { layoutWriter, hitTestWriterPage, writerCaretBox, writerSelectionBoxes, type WriterLayout } from '@office-ex/render';
import type { MessageKey } from '@office-ex/i18n';
import { CanvasFontMetrics } from './font-metrics.js';

export interface WriterViewState{paragraphId?:string;start:number;end:number;pageIndex:number;focus?:boolean}
export interface WriterHost{
  snapshot():DocumentSnapshot;
  execute(command:DocumentCommand):boolean;
  render():void;
  error(error:unknown):void;
  composing():boolean;
  setComposing(value:boolean):void;
  t(key:MessageKey,args?:Readonly<Record<string,string|number>>):string;
  number(value:number):string;
}
function node<K extends keyof HTMLElementTagNameMap>(tag:K,value='',className?:string):HTMLElementTagNameMap[K]{
  const n=document.createElement(tag);n.textContent=value;if(className)n.className=className;return n;
}
function modelOf(snapshot:DocumentSnapshot):DeepReadonly<WriterDocument>{
  if(snapshot.document.kind!=='writer')throw new Error('Expected Writer snapshot');return snapshot.document;
}
function inputId(index:number):string{return index===0?'draft-text':'writer-paragraph-'+index;}
function button(label:string,fn:()=>void,id?:string):HTMLButtonElement{
  const b=node('button',label);b.type='button';if(id)b.id=id;b.onclick=fn;return b;
}
function nearestBoundary(value:string,offset:number):number{
  const b=graphemeBoundaries(value);let best=0;
  for(const n of b){if(Math.abs(n-offset)<Math.abs(best-offset))best=n;}
  return best;
}
function safeColor(style:DeepReadonly<TextStyle>):string{
  const c=style.color;return c?.kind==='rgba'?'rgba('+c.red+','+c.green+','+c.blue+','+c.alpha/255+')':'#20242c';
}
export function writerWorkspace(host:WriterHost,state:WriterViewState):HTMLElement{
  const root=node('div','','writer-workspace');
  const controls=node('div','','writer-controls'),panes=node('div','','writer-panes');
  const edit=node('section','','writer-edit-pane'),preview=node('section','','writer-preview-pane');
  const inputs=new Map<string,HTMLTextAreaElement>();
  const t=host.t;
  controls.setAttribute('role','toolbar');controls.setAttribute('aria-label',t('writerFormatting'));
  const note=node('p',t('writerEditNote'),'muted');root.append(controls,note,panes);panes.append(edit,preview);
  edit.append(node('h2',t('textLabel')));preview.append(node('h2',t('writerPagePreview')));
  const pageBar=node('div','','writer-pagebar'),pageText=node('span','');
  const canvas=node('canvas','','writer-page-canvas');canvas.id='writer-page-preview';canvas.setAttribute('role','img');canvas.setAttribute('aria-label',t('writerPagePreview'));
  const layoutNote=node('p',t('writerEstimated'),'muted');
  const previous=button(t('writerPreviousPage'),()=>{state.pageIndex=Math.max(0,state.pageIndex-1);paint();},'writer-previous-page');
  const next=button(t('writerNextPage'),()=>{if(layout)state.pageIndex=Math.min(layout.pages.length-1,state.pageIndex+1);paint();},'writer-next-page');
  previous.disabled=true;next.disabled=true;pageBar.append(previous,pageText,next);preview.append(pageBar,canvas,layoutNote);
  let layout:DeepReadonly<WriterLayout>|undefined, timer:ReturnType<typeof setTimeout>|undefined;
  const provider=new CanvasFontMetrics();
  function activeParagraph(){return state.paragraphId?inputs.get(state.paragraphId):undefined;}
  function select(id:string,start:number,end=start,focus=false):void{
    const input=inputs.get(id);if(!input)return;
    const a=nearestBoundary(input.value,Math.max(0,Math.min(start,input.value.length))),b=nearestBoundary(input.value,Math.max(0,Math.min(end,input.value.length)));
    state.paragraphId=id;state.start=Math.min(a,b);state.end=Math.max(a,b);
    input.setSelectionRange(Math.min(a,b),Math.max(a,b));if(focus)input.focus({preventScroll:true});state.paragraphId=id;state.start=Math.min(a,b);state.end=Math.max(a,b);
    syncControls();follow();paint();
  }
  function readSelection(input:HTMLTextAreaElement,id:string):void{
    if(host.composing())return;
    state.paragraphId=id;state.start=input.selectionStart;state.end=input.selectionEnd;syncControls();follow();paint();
  }
  function follow():void{
    if(layout&&state.paragraphId){const caret=writerCaretBox(layout,{paragraphId:state.paragraphId,offset:state.end},state.pageIndex);if(caret)state.pageIndex=caret.pageIndex;}
  }
  function paint():void{
    if(!root.isConnected||!layout||layout.revision!==host.snapshot().revision)return;
    state.pageIndex=Math.max(0,Math.min(state.pageIndex,layout.pages.length-1));const page=layout.pages[state.pageIndex]!;
    const scale=Math.min(1,1200/(page.width/EMU_PER_CSS_PIXEL),1600/(page.height/EMU_PER_CSS_PIXEL));
    canvas.width=Math.max(1,Math.round(page.width/EMU_PER_CSS_PIXEL*scale));canvas.height=Math.max(1,Math.round(page.height/EMU_PER_CSS_PIXEL*scale));
    const ctx=canvas.getContext('2d');if(!ctx)return;
    ctx.fillStyle='#fff';ctx.fillRect(0,0,canvas.width,canvas.height);
    const unit=EMU_PER_CSS_PIXEL,px=(value:number)=>value/unit;
    ctx.setTransform(canvas.width/px(page.width),0,0,canvas.height/px(page.height),0,0);
    ctx.beginPath();ctx.rect(px(page.margins.left),px(page.margins.top),px(page.width-page.margins.left-page.margins.right),px(page.height-page.margins.top-page.margins.bottom));ctx.clip();
    if(state.paragraphId){
      try{
        ctx.fillStyle='#dce8ff';
        for(const rect of writerSelectionBoxes(layout,state.paragraphId,state.start,state.end))if(rect.pageIndex===state.pageIndex)ctx.fillRect(px(rect.x),px(rect.y),px(rect.width),px(rect.height));
      }catch{/* Native textarea may temporarily select inside a grapheme. Never mutate source. */}
    }
    ctx.textBaseline='alphabetic';ctx.textAlign='left';ctx.direction='ltr';
    for(const line of page.lines)for(const fragment of line.fragments){
      const r=fragment.request,family=/^(sans-serif|serif|monospace|system-ui)$/.test(r.family)?r.family:JSON.stringify(r.family);
      ctx.font=(r.italic?'italic ':'')+(r.bold?'bold ':'')+px(r.size)+'px '+family;
      ctx.fillStyle=safeColor(fragment.style);ctx.fillText(fragment.text,px(fragment.x),px(line.baseline));
      if(fragment.style.underline){ctx.fillRect(px(fragment.x),px(line.baseline+r.size*0.06),px(fragment.width),Math.max(0.5,px(r.size*0.05)));}
    }
    if(state.paragraphId&&state.start===state.end){
      const caret=writerCaretBox(layout,{paragraphId:state.paragraphId,offset:state.end},state.pageIndex);
      if(caret?.pageIndex===state.pageIndex){ctx.fillStyle='#245bca';ctx.fillRect(px(caret.x),px(caret.y),Math.max(1,px(page.width/canvas.width*1.5)),px(caret.height));}
    }
    pageText.textContent=t('writerPageNumber',{current:host.number(state.pageIndex+1),total:host.number(layout.pages.length)});
    previous.disabled=state.pageIndex===0;next.disabled=state.pageIndex===layout.pages.length-1;
    canvas.hidden=false;layoutNote.textContent=t('writerEstimated');
    if(layout.warnings.some(w=>w!=='writer.estimated-wrap'&&w!=='font.cluster-positions-estimated'))layoutNote.textContent+=' '+t('writerLayoutLimits');
  }
  function updateLayout():void{
    if(!root.isConnected||host.composing())return;
    try{
      const snapshot=host.snapshot();const result=layoutWriter(snapshot,provider,{maxPages:100,maxClusters:20000});
      if(host.snapshot().revision!==result.revision)return;
      layout=result;follow();paint();
    }catch{
      layout=undefined;canvas.hidden=true;previous.disabled=true;next.disabled=true;pageText.textContent='';layoutNote.textContent=t('writerPreviewUnavailable');
    }
  }
  function scheduleLayout():void{
    if(timer!==undefined)clearTimeout(timer);
    // Derive layout from current revision after typing settles; never alter model state.
    timer=setTimeout(()=>{timer=undefined;updateLayout();},100);
  }
  canvas.onclick=e=>{
    if(!layout||host.composing()||layout.revision!==host.snapshot().revision)return;
    const page=layout.pages[state.pageIndex],box=canvas.getBoundingClientRect();if(!page||box.width<=0||box.height<=0)return;
    const hit=hitTestWriterPage(layout,state.pageIndex,{x:(e.clientX-box.left)/box.width*page.width,y:(e.clientY-box.top)/box.height*page.height});
    if(hit)select(hit.paragraphId,hit.offset,hit.offset,true);
  };
  function currentSelection(){const input=activeParagraph();if(!input||!state.paragraphId)return undefined;
    const start=nearestBoundary(input.value,input.selectionStart),end=nearestBoundary(input.value,input.selectionEnd);
    return {id:state.paragraphId,input,start:Math.min(start,end),end:Math.max(start,end)};
  }
  function safe(fn:()=>void):void{try{fn();}catch(e){host.error(e);}}
  const actions:{button:HTMLButtonElement;range?:boolean;enabled?:()=>boolean}[]=[];
  const formatFields:{element:HTMLInputElement|HTMLSelectElement;range:boolean}[]=[];
  function tool(key:MessageKey,id:string,fn:()=>void,range=false):void{
    const b=button(t(key),()=>safe(fn),id);b.onmousedown=e=>e.preventDefault();actions.push({button:b,range});controls.append(b);
  }
  function selectedStyle(key:'bold'|'italic'|'underline'):boolean{
    const selection=currentSelection();if(!selection||selection.start===selection.end)return false;
    const doc=modelOf(host.snapshot());
    const p=doc.sections.flatMap(s=>s.blocks).find(b=>b.id===selection.id);
    if(!p||p.type!=='paragraph')return false;
    const base=resolveTextStyle(doc.defaultTextStyle,doc.styles,p.styleId?[p.styleId]:[],p.style??{}).style;
    let cursor=0,selected=0,all=true;
    for(const r of p.runs){if(r.type!=='text')return false;const end=cursor+r.text.length;
      if(cursor<selection.end&&end>selection.start){selected++;all=all&&resolveTextStyle(base,doc.styles,r.styleId?[r.styleId]:[],r.style??{}).style[key]===true;}cursor=end;}
    return selected>0&&all;
  }
  function selectedParagraph(){
    const selection=currentSelection();if(!selection)return undefined;
    const doc=modelOf(host.snapshot());const p=doc.sections.flatMap(s=>s.blocks).find(b=>b.id===selection.id);
    return p?.type==='paragraph'?p:undefined;
  }
  function effectiveSelectedStyle():DeepReadonly<TextStyle>|undefined{
    const selection=currentSelection(),p=selectedParagraph();if(!selection||!p)return undefined;
    const doc=modelOf(host.snapshot()),base=resolveTextStyle(doc.defaultTextStyle,doc.styles,p.styleId?[p.styleId]:[],p.style??{}).style;
    let cursor=0;
    for(const r of p.runs){if(r.type!=='text')return base;const end=cursor+r.text.length;
      if(selection.start<=end&&(selection.start<end||selection.start===selection.end))return resolveTextStyle(base,doc.styles,r.styleId?[r.styleId]:[],r.style??{}).style;cursor=end;
    }
    return base;
  }
  function syncControls():void{
    const selection=currentSelection();
    for(const action of actions)action.button.disabled=!selection||host.composing()||!!action.range&&selection.start===selection.end;
    for(const field of formatFields)field.element.disabled=!selection||host.composing()||field.range&&selection.start===selection.end;
    const style=effectiveSelectedStyle(),p=selectedParagraph();
    // The first selected run supplies the displayed value; mixed selections are
    // not represented as uniform. A field change applies only that one property.
    if(style){
      const family=root.querySelector<HTMLInputElement>('#writer-font-family');
      const size=root.querySelector<HTMLInputElement>('#writer-font-size');
      const color=root.querySelector<HTMLInputElement>('#writer-text-color');
      if(family&&document.activeElement!==family)family.value=style.fontFamily??'sans-serif';
      if(size&&document.activeElement!==size)size.value=String((style.fontSize??152400)/EMU_PER_POINT);
      if(color&&document.activeElement!==color){
        const c=style.color;color.value=c?.kind==='rgba'?'#'+[c.red,c.green,c.blue].map(n=>n.toString(16).padStart(2,'0')).join(''):'#20242c';
      }
      const named=root.querySelector<HTMLSelectElement>('#writer-paragraph-style');
      if(named&&document.activeElement!==named)named.value=p?.styleId??'';
      const character=root.querySelector<HTMLSelectElement>('#writer-character-style');
      if(character&&document.activeElement!==character){
        const sel=currentSelection();let offset=0,current='';
        if(sel&&p)for(const r of p.runs){if(r.type!=='text')break;const end=offset+r.text.length;if(sel.start<end||sel.start===end&&sel.start===sel.end){current=r.styleId??'';break;}offset=end;}
        character.value=current;
      }
    }
    for(const key of ['bold','italic','underline'] as const){
      const b=root.querySelector<HTMLButtonElement>('#writer-'+key);if(b)b.setAttribute('aria-pressed',String(selectedStyle(key)));
    }
  }
  function applyDirectStyle(style:TextStyle):void{
    const sel=currentSelection();if(!sel||sel.start===sel.end||host.composing())return;
    if(host.execute(formatParagraphRange(host.snapshot(),sel.id,sel.start,sel.end,style,crypto.randomUUID()))){
      select(sel.id,sel.start,sel.end,true);scheduleLayout();
    }
  }
  function format(key:'bold'|'italic'|'underline'):void{applyDirectStyle({[key]:!selectedStyle(key)});}
  const properties=node('div','','writer-format-fields');
  function field<K extends HTMLInputElement|HTMLSelectElement>(key:MessageKey,element:K,range=true):K{
    const label=node('label','','writer-format-field');label.append(node('span',t(key)),element);element.setAttribute('aria-label',t(key));
    properties.append(label);formatFields.push({element,range});return element;
  }
  const family=field('writerFontFamily',node('input'));family.id='writer-font-family';family.type='text';family.maxLength=512;family.autocomplete='off';
  const fontList=node('datalist');fontList.id='writer-font-options';
  for(const face of ['sans-serif','serif','monospace','Arial','Calibri','Aptos','Times New Roman','Helvetica','Noto Sans','Noto Serif','Microsoft YaHei','SimSun','PingFang SC','Hiragino Sans','Malgun Gothic']){
    const option=node('option');option.value=face;fontList.append(option);
  }
  family.setAttribute('list',fontList.id);properties.append(fontList);
  family.onchange=()=>safe(()=>{const value=family.value.trim();if(!value||/[\x00-\x1f\x7f]/.test(value))throw new OfficeError('invalid','Invalid font family');applyDirectStyle({fontFamily:value});});
  const size=field('writerFontSize',node('input'));size.id='writer-font-size';size.type='number';size.min='1';size.max='1638';size.step='0.5';
  size.onchange=()=>safe(()=>{const n=size.valueAsNumber;if(!Number.isFinite(n)||n<1||n>1638||!Number.isInteger(n*2))throw new OfficeError('invalid','Font size must use half-point increments');applyDirectStyle({fontSize:n*EMU_PER_POINT});});
  const color=field('writerTextColor',node('input'));color.id='writer-text-color';color.type='color';
  color.onchange=()=>safe(()=>{const v=color.value;if(!/^#[a-f0-9]{6}$/i.test(v))throw new OfficeError('invalid','Invalid text color');applyDirectStyle({color:{kind:'rgba',red:parseInt(v.slice(1,3),16),green:parseInt(v.slice(3,5),16),blue:parseInt(v.slice(5,7),16),alpha:255}});});
  const doc=modelOf(host.snapshot());
  for(const kind of ['paragraph','character'] as const){
    const list=Object.values(doc.styles).filter(s=>s.kind===kind);
    if(!list.length)continue;
    const selectStyle=field(kind==='paragraph'?'writerParagraphStyle':'writerCharacterStyle',node('select'),kind==='character');selectStyle.id='writer-'+kind+'-style';
    const blank=node('option',t('writerChooseStyle'));blank.value='';blank.disabled=true;selectStyle.append(blank);
    for(const style of list){const item=node('option',style.name);item.value=style.id;selectStyle.append(item);}
    selectStyle.onchange=()=>safe(()=>{
      const sel=currentSelection();if(!sel||!selectStyle.value||host.composing())return;
      const command=kind==='paragraph'?setParagraphNamedStyle(host.snapshot(),sel.id,selectStyle.value):setCharacterNamedStyle(host.snapshot(),sel.id,sel.start,sel.end,selectStyle.value,crypto.randomUUID());
      if(host.execute(command)){select(sel.id,sel.start,sel.end,true);scheduleLayout();}
    });
  }
  root.insertBefore(properties,note);
  for(const [key,label] of [['bold','writerBold'],['italic','writerItalic'],['underline','writerUnderline']] as const)tool(label,'writer-'+key,()=>format(key),true);
  for(const [alignment,label] of [['start','writerAlignLeft'],['center','writerAlignCenter'],['end','writerAlignRight']] as const)
    tool(label,'writer-align-'+alignment,()=>{const sel=currentSelection();if(sel&&!host.composing()&&host.execute(setParagraphAlignment(host.snapshot(),sel.id,alignment)))scheduleLayout();});
  function split():void{
    const sel=currentSelection();if(!sel||host.composing())return;
    const newId=crypto.randomUUID(),prefix=crypto.randomUUID();
    if(host.execute(splitParagraphRange(host.snapshot(),sel.id,sel.start,sel.end,prefix,newId))){
      state.paragraphId=newId;state.start=0;state.end=0;state.focus=true;host.render();
    }
  }
  tool('writerSplit','writer-split',split);
  tool('writerPageBreak','writer-page-break',()=>{const sel=currentSelection();if(sel&&!host.composing()&&host.execute(insertPageBreakAfter(host.snapshot(),sel.id,crypto.randomUUID())))host.render();});
  function join(id:string,direction:'previous'|'next'):void{
    const d=modelOf(host.snapshot());
    for(const section of d.sections){
      const index=section.blocks.findIndex(b=>b.id===id);if(index<0)continue;
      const a=direction==='previous'?section.blocks[index-1]:section.blocks[index],b=direction==='previous'?section.blocks[index]:section.blocks[index+1];
      if(a?.type!=='paragraph'||b?.type!=='paragraph')return;
      const offset=paragraphText(host.snapshot(),a.id).length;
      if(host.execute(joinParagraphs(host.snapshot(),a.id,b.id))){state.paragraphId=a.id;state.start=offset;state.end=offset;state.focus=true;host.render();}return;
    }
  }
  let index=0;
  for(const section of modelOf(host.snapshot()).sections)for(const p of section.blocks){
    if(p.type!=='paragraph'||p.runs.some(r=>r.type!=='text')){
      const indicator=node('p',t(p.type==='page-break'?'writerPageBreak':'writerUnsupportedBlock'),'writer-block-note');edit.append(indicator);continue;
    }
    const input=node('textarea','','writer-paragraph draft-text');input.id=inputId(index++);input.dataset.paragraphId=p.id;
    input.dir='auto';input.rows=3;input.setAttribute('aria-label',t('textLabel')+' '+host.number(index));input.value=paragraphText(host.snapshot(),p.id);
    inputs.set(p.id,input);edit.append(input);
    const commit=()=>{
      try{
        const s=host.snapshot(),old=paragraphText(s,p.id),change=paragraphEditDiff(old,input.value);
        if(change&&!host.execute(replaceParagraphRange(s,p.id,change.start,change.end,change.replacement,crypto.randomUUID())))input.value=paragraphText(host.snapshot(),p.id);
      }catch(error){input.value=paragraphText(host.snapshot(),p.id);host.error(error);}
      readSelection(input,p.id);scheduleLayout();
    };
    input.addEventListener('compositionstart',()=>{host.setComposing(true);syncControls();});
    input.addEventListener('compositionend',()=>{host.setComposing(false);commit();syncControls();});
    input.oninput=event=>{if(!host.composing()&&!(event as InputEvent).isComposing)commit();};
    input.onfocus=()=>readSelection(input,p.id);input.onselect=()=>readSelection(input,p.id);input.onkeyup=()=>readSelection(input,p.id);
    input.onpointerup=()=>readSelection(input,p.id);
    input.onkeydown=e=>{
      if(host.composing()||e.isComposing||e.key==='Process'||e.ctrlKey||e.metaKey||e.altKey)return;
      if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();readSelection(input,p.id);safe(split);}
      else if(e.key==='Backspace'&&input.selectionStart===0&&input.selectionEnd===0){e.preventDefault();safe(()=>join(p.id,'previous'));}
      else if(e.key==='Delete'&&input.selectionStart===input.value.length&&input.selectionEnd===input.value.length){e.preventDefault();safe(()=>join(p.id,'next'));}
    };
  }
  if(!state.paragraphId||!inputs.has(state.paragraphId)){state.paragraphId=inputs.keys().next().value;state.start=0;state.end=0;}
  const initial=activeParagraph();if(initial){initial.setSelectionRange(state.start,state.end);if(state.focus){state.focus=false;requestAnimationFrame(()=>{if(initial.isConnected)initial.focus({preventScroll:true});});}}
  syncControls();requestAnimationFrame(updateLayout);return root;
}
