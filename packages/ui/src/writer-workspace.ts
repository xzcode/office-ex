// SPDX-License-Identifier: MIT
import { EMU_PER_CSS_PIXEL, EMU_PER_POINT, EMU_PER_TWIP, graphemeBoundaries, OfficeError } from '@office-ex/contracts';
import { resolveTextStyle, inlineText, paragraphBoundaries, OBJECT_REPLACEMENT, type DocumentSnapshot, type DocumentCommand, type DeepReadonly, type TextStyle, type WriterDocument, type WriterBlock, type Paragraph, type Table, type ImageResource } from '@office-ex/core';
import { paragraphText, replaceParagraphRange, splitParagraphRange, joinParagraphs, formatParagraphRange, setParagraphAlignment, setParagraphNamedStyle, setCharacterNamedStyle, insertPageBreakAfter, paragraphEditDiff, insertTableAfter, insertTableRow, deleteTableRow, insertTableColumn, deleteTableColumn, resizeTableColumn, setTableBorder, deleteTable, tableCellLocation, insertImageAfter, insertImageAt, resizeInlineImage, removeInlineImage, setWriterPage, splitWriterSectionAfter, setParagraphMetrics, setParagraphDirection, locateParagraph, writerTextGroup, textGroupPoint, textGroupOffset, replaceWriterRange, resolveWriterRange, formatWriterRange, type WriterRange, type WriterTextGroup } from '@office-ex/writer';
import { layoutWriter, hitTestWriterPage, writerCaretBox, writerSelectionBoxes, type WriterLayout } from '@office-ex/render';
import type { MessageKey } from '@office-ex/i18n';
import { CanvasFontMetrics } from './font-metrics.js';

export interface WriterViewState{paragraphId?:string;start:number;end:number;pageIndex:number;focus?:boolean;continuous?:boolean;groupRange?:WriterRange;pageSectionId?:string}
export interface WriterHost{
  snapshot():DocumentSnapshot;
  execute(command:DocumentCommand):boolean;
  render():void;
  error(error:unknown):void;
  composing():boolean;
  setComposing(value:boolean):void;
  t(key:MessageKey,args?:Readonly<Record<string,string|number>>):string;
  number(value:number):string;
  prepareImage(bytes:Uint8Array):Promise<ImageResource>;
  isCurrent():boolean;
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
  let continuousRange:WriterRange|undefined,continuousPick:((id:string,offset:number)=>boolean)|undefined;
  const decodedImages=new Map<string,HTMLImageElement>();
  function raster(resourceId:string):HTMLImageElement|undefined{
    const prior=decodedImages.get(resourceId);if(prior)return prior;
    const r=modelOf(host.snapshot()).resources[resourceId];if(!r?.dataBase64||!(r.mimeType==='image/png'||r.mimeType==='image/jpeg'))return undefined;
    const image=new Image();decodedImages.set(resourceId,image);image.onload=()=>paint();image.onerror=()=>{if(root.isConnected)layoutNote.textContent=t('writerPreviewUnavailable');};image.src='data:'+r.mimeType+';base64,'+r.dataBase64;return image;
  }
  function activeParagraph(){return state.paragraphId?inputs.get(state.paragraphId):undefined;}
  function paragraphBoundary(id:string,offset:number):number{
    const p=locateParagraph(host.snapshot(),id).paragraph;let best=0;
    for(const n of paragraphBoundaries(p))if(Math.abs(n-offset)<Math.abs(best-offset))best=n;
    return best;
  }
  function select(id:string,start:number,end=start,focus=false):void{
    const input=inputs.get(id);if(!input)return;
    const a=paragraphBoundary(id,start),b=paragraphBoundary(id,end);
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
        const segments=continuousRange?resolveWriterRange(host.snapshot(),continuousRange).segments:[{paragraphId:state.paragraphId,start:state.start,end:state.end}];
        for(const seg of segments)for(const rect of writerSelectionBoxes(layout,seg.paragraphId,seg.start,seg.end))if(rect.pageIndex===state.pageIndex)ctx.fillRect(px(rect.x),px(rect.y),px(rect.width),px(rect.height));
      }catch{/* Native textarea may temporarily select inside a grapheme. Never mutate source. */}
    }
    for(const cell of page.tableCells){
      const border=cell.border;
      if(border){ctx.strokeStyle=safeColor({color:border.color});ctx.lineWidth=Math.max(0.35,px(border.width));ctx.strokeRect(px(cell.x),px(cell.y),px(cell.width),px(cell.height));}
    }
    for(const item of page.images){const img=raster(item.resourceId);if(img?.complete&&img.naturalWidth)ctx.drawImage(img,px(item.x),px(item.y),px(item.width),px(item.height));else{ctx.fillStyle='#eef1f5';ctx.fillRect(px(item.x),px(item.y),px(item.width),px(item.height));}}
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
    if(hit){if(continuousPick){if(!continuousPick(hit.paragraphId,hit.offset))return;}else if(inputs.has(hit.paragraphId))select(hit.paragraphId,hit.offset,hit.offset,true);}
  };
  function currentSelection(){const input=activeParagraph();if(!input||!state.paragraphId)return undefined;
    const start=paragraphBoundary(state.paragraphId,input.selectionStart),end=paragraphBoundary(state.paragraphId,input.selectionEnd);
    return {id:state.paragraphId,input,start:Math.min(start,end),end:Math.max(start,end)};
  }
  function safe(fn:()=>void):void{try{fn();}catch(e){host.error(e);}}
  const actions:{button:HTMLButtonElement;range?:boolean;enabled?:()=>boolean}[]=[];
  const tableInputs:HTMLInputElement[]=[];
  const metricFields:HTMLInputElement[]=[];
  function contextTable(){const id=state.paragraphId;if(!id)return undefined;const location=tableCellLocation(host.snapshot(),id);if(!location)return undefined;const table=modelOf(host.snapshot()).sections.flatMap(s=>s.blocks).find(b=>b.id===location.tableId);return table?.type==='table'?{...location,table}:undefined;}
  const formatFields:{element:HTMLInputElement|HTMLSelectElement;range:boolean}[]=[];
  function tool(key:MessageKey,id:string,fn:()=>void,range=false,enabled?:()=>boolean,parent:HTMLElement=controls):void{
    const b=button(t(key),()=>safe(fn),id);b.onmousedown=e=>e.preventDefault();actions.push({button:b,range,enabled});parent.append(b);
  }
  function selectedStyle(key:'bold'|'italic'|'underline'):boolean{
    const selection=currentSelection();if(!selection||selection.start===selection.end)return false;
    const doc=modelOf(host.snapshot());
    const p=selectedParagraph();
    if(!p||p.type!=='paragraph')return false;
    const base=resolveTextStyle(doc.defaultTextStyle,doc.styles,(p.styleId??doc.defaultParagraphStyleId)?[p.styleId??doc.defaultParagraphStyleId!]:[],p.style??{}).style;
    let cursor=0,selected=0,all=true;
    for(const r of p.runs){const end=cursor+inlineText(r).length;
      if(r.type==='text'&&cursor<selection.end&&end>selection.start){selected++;all=all&&resolveTextStyle(base,doc.styles,r.styleId?[r.styleId]:[],r.style??{}).style[key]===true;}cursor=end;}
    return selected>0&&all;
  }
  function selectedParagraph(){
    const selection=currentSelection();if(!selection)return undefined;
    const doc=modelOf(host.snapshot());
    for(const section of doc.sections)for(const block of section.blocks){
      if(block.type==='paragraph'&&block.id===selection.id)return block;
      if(block.type==='table')for(const row of block.rows)for(const cell of row.cells)for(const p of cell.blocks)if(p.type==='paragraph'&&p.id===selection.id)return p;
    }
    return undefined;
  }
  function effectiveSelectedStyle():DeepReadonly<TextStyle>|undefined{
    const selection=currentSelection(),p=selectedParagraph();if(!selection||!p)return undefined;
    const doc=modelOf(host.snapshot()),base=resolveTextStyle(doc.defaultTextStyle,doc.styles,(p.styleId??doc.defaultParagraphStyleId)?[p.styleId??doc.defaultParagraphStyleId!]:[],p.style??{}).style;
    let cursor=0;
    for(const r of p.runs){const end=cursor+inlineText(r).length;
      if(r.type==='text'&&selection.start<=end&&(selection.start<end||selection.start===selection.end))return resolveTextStyle(base,doc.styles,r.styleId?[r.styleId]:[],r.style??{}).style;cursor=end;
    }
    return base;
  }
  function syncControls():void{
    const selection=currentSelection();
    for(const action of actions)action.button.disabled=!selection||host.composing()||!!action.range&&selection.start===selection.end||!!action.enabled&&!action.enabled();
    const tableContext=contextTable();
    for(const input of tableInputs)input.disabled=host.composing()||!selection||(input.id==='writer-table-width'?!tableContext:!!tableContext);
    const width=root.querySelector<HTMLInputElement>('#writer-table-width');if(width&&document.activeElement!==width)width.value=tableContext?String(tableContext.table.columnWidths[tableContext.column]!/EMU_PER_POINT):'';
    const border=root.querySelector<HTMLButtonElement>('#writer-table-border');if(border)border.setAttribute('aria-pressed',String(!!tableContext?.table.border));
    for(const field of formatFields)field.element.disabled=!selection||host.composing()||field.range&&selection.start===selection.end;
    const style=effectiveSelectedStyle(),p=selectedParagraph();
    for(const field of metricFields){field.disabled=!selection||host.composing();const key=field.dataset.metric as 'indent'|'spaceBefore'|'spaceAfter'|'lineHeight';if(document.activeElement!==field)field.value=String((style?.[key]??0)/EMU_PER_POINT);}
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
      if(named&&document.activeElement!==named)named.value=p?.styleId??modelOf(host.snapshot()).defaultParagraphStyleId??'';
      const lang=root.querySelector<HTMLInputElement>('#writer-text-language');if(lang&&document.activeElement!==lang)lang.value=style.language??'';
      const dir=root.querySelector<HTMLSelectElement>('#writer-paragraph-direction');if(dir&&document.activeElement!==dir)dir.value=style.paragraphDirection??'';
      const character=root.querySelector<HTMLSelectElement>('#writer-character-style');
      if(character&&document.activeElement!==character){
        const sel=currentSelection();let offset=0,current='';
        if(sel&&p)for(const r of p.runs){const end=offset+inlineText(r).length;if(r.type==='text'&&(sel.start<end||sel.start===end&&sel.start===sel.end)){current=r.styleId??'';break;}offset=end;}
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
  const lang=field('writerTextLanguage',node('input'));lang.id='writer-text-language';lang.type='text';lang.maxLength=128;lang.placeholder='en-US';
  lang.onchange=()=>safe(()=>{const value=lang.value.trim();if(!value)throw new OfficeError('invalid','Document language tag is required');applyDirectStyle({language:value});});
  const dir=field('writerParagraphDirection',node('select'),false);dir.id='writer-paragraph-direction';
  for(const [value,key] of [['','writerChooseStyle'],['ltr','writerLtr'],['rtl','writerRtl']] as const){const item=node('option',t(key));item.value=value;if(!value)item.disabled=true;dir.append(item);}
  dir.onchange=()=>safe(()=>{const sel=currentSelection();if(sel&&!host.composing()&&(dir.value==='ltr'||dir.value==='rtl')&&host.execute(setParagraphDirection(host.snapshot(),sel.id,dir.value))){state.focus=true;host.render();}});
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
  tool('writerPageBreak','writer-page-break',()=>{const sel=currentSelection();if(sel&&!host.composing()&&host.execute(insertPageBreakAfter(host.snapshot(),sel.id,crypto.randomUUID())))host.render();},false,()=>!contextTable());
  const tableBar=node('div','','writer-table-controls');
  tableBar.setAttribute('role','toolbar');tableBar.setAttribute('aria-label',t('writerTable'));
  function tableField(key:MessageKey,id:string,value:string,min:number,max:number,step:string):HTMLInputElement{
    const label=node('label','','writer-format-field'),input=node('input');input.id=id;input.type='number';input.min=String(min);input.max=String(max);input.step=step;input.value=value;input.setAttribute('aria-label',t(key));
    label.append(node('span',t(key)),input);tableBar.append(label);tableInputs.push(input);return input;
  }
  const rowsField=tableField('writerTableRows','writer-table-rows','2',1,50,'1');
  const colsField=tableField('writerTableColumns','writer-table-columns','2',1,16,'1');
  function focusTable(tableId:string,row=0,column=0):void{
    const table=modelOf(host.snapshot()).sections.flatMap(s=>s.blocks).find(b=>b.id===tableId);
    if(table?.type!=='table')return;
    const cell=table.rows[Math.min(row,table.rows.length-1)]!.cells[Math.min(column,table.columnWidths.length-1)]!;
    const paragraph=cell.blocks.find(b=>b.type==='paragraph');
    if(paragraph){state.paragraphId=paragraph.id;state.start=0;state.end=0;state.focus=true;}
  }
  function changeTable(make:(c:NonNullable<ReturnType<typeof contextTable>>)=>DocumentCommand,rowDelta=0,colDelta=0):void{
    const c=contextTable();if(!c||host.composing())return;
    if(host.execute(make(c))){focusTable(c.tableId,c.row+rowDelta,c.column+colDelta);host.render();}
  }
  tool('writerInsertTable','writer-insert-table',()=>{
    const sel=currentSelection();if(!sel||host.composing()||contextTable())return;
    const prefix=crypto.randomUUID();
    if(host.execute(insertTableAfter(host.snapshot(),sel.id,rowsField.valueAsNumber,colsField.valueAsNumber,prefix))){
      focusTable(prefix+'-1');host.render();
    }
  },false,()=>!contextTable(),tableBar);
  tool('writerAddRow','writer-table-add-row',()=>changeTable(c=>insertTableRow(host.snapshot(),c.tableId,c.row+1,crypto.randomUUID()),1),false,()=>!!contextTable(),tableBar);
  tool('writerDeleteRow','writer-table-delete-row',()=>changeTable(c=>deleteTableRow(host.snapshot(),c.tableId,c.row)),false,()=>!!contextTable()&&contextTable()!.table.rows.length>1,tableBar);
  tool('writerAddColumn','writer-table-add-column',()=>changeTable(c=>insertTableColumn(host.snapshot(),c.tableId,c.column+1,crypto.randomUUID()),0,1),false,()=>!!contextTable(),tableBar);
  tool('writerDeleteColumn','writer-table-delete-column',()=>changeTable(c=>deleteTableColumn(host.snapshot(),c.tableId,c.column)),false,()=>!!contextTable()&&contextTable()!.table.columnWidths.length>1,tableBar);
  const widthField=tableField('writerColumnWidth','writer-table-width','',12,1584,'0.05');
  widthField.onchange=()=>safe(()=>{const n=widthField.valueAsNumber;if(!Number.isFinite(n)||n<12)throw new OfficeError('invalid','Invalid column width');changeTable(c=>resizeTableColumn(host.snapshot(),c.tableId,c.column,Math.round(n*20)*EMU_PER_TWIP));});
  tool('writerTableBorder','writer-table-border',()=>changeTable(c=>setTableBorder(host.snapshot(),c.tableId,c.table.border?null:{width:6350,color:{kind:'rgba',red:0,green:0,blue:0,alpha:255}})),false,()=>!!contextTable(),tableBar);
  tool('writerDeleteTable','writer-table-delete',()=>{
    const c=contextTable();if(!c||host.composing())return;
    if(host.execute(deleteTable(host.snapshot(),c.tableId))){state.paragraphId=undefined;state.start=0;state.end=0;host.render();}
  },false,()=>!!contextTable(),tableBar);
  root.insertBefore(tableBar,note);
  const pictureBar=node('div','','writer-table-controls');
  let pictureMode:'after'|'inline'='after';
  const imagePicker=node('input');imagePicker.type='file';imagePicker.accept='image/png,image/jpeg,.png,.jpg,.jpeg';imagePicker.id='writer-image-file';imagePicker.hidden=true;
  pictureBar.append(imagePicker);
  tool('writerInsertImage','writer-insert-image',()=>{pictureMode='after';imagePicker.click();},false,undefined,pictureBar);
  tool('writerInsertInlineImage','writer-insert-inline-image',()=>{pictureMode='inline';imagePicker.click();},false,undefined,pictureBar);
  imagePicker.onchange=()=>{const file=imagePicker.files?.[0],sel=currentSelection(),revision=host.snapshot().revision,mode=pictureMode;imagePicker.value='';
    if(!file||!sel||host.composing())return;
    void(async()=>{
      if(file.size>8*1024*1024)throw new OfficeError('limit','Image exceeds byte budget');
      const resource=await host.prepareImage(new Uint8Array(await file.arrayBuffer()));
      if(!root.isConnected||!host.isCurrent()||host.composing()||host.snapshot().revision!==revision)throw new OfficeError('conflict','Document changed during image selection');
      const id=crypto.randomUUID(),command=mode==='inline'?insertImageAt(host.snapshot(),sel.id,sel.start,sel.end,resource,id):insertImageAfter(host.snapshot(),sel.id,resource,id);
      if(host.execute(command)){if(mode==='inline'){state.start=state.end=sel.start+1;state.focus=true;}host.render();}
    })().catch(host.error);
  };
  root.insertBefore(pictureBar,note);
  const pageBarControls=node('div','','writer-table-controls');
  const sections=modelOf(host.snapshot()).sections;
  let docSection=sections.find(s=>s.id===state.pageSectionId)??sections[0]!;state.pageSectionId=docSection.id;const pg=docSection.page;
  const sectionChoice=node('select');sectionChoice.id='writer-section';sectionChoice.setAttribute('aria-label',t('writerSection'));
  for(let i=0;i<sections.length;i++){const item=node('option',t('writerSectionNumber',{number:host.number(i+1)}));item.value=sections[i]!.id;sectionChoice.append(item);}sectionChoice.value=docSection.id;pageBarControls.append(sectionChoice);
  function setting(key:MessageKey,id:string,value:number):HTMLInputElement{
    const label=node('label','','writer-format-field'),input=node('input');input.type='number';input.id=id;input.step='0.05';input.min='0';input.max='1584';input.value=String(Math.round(value/EMU_PER_TWIP)/20);input.setAttribute('aria-label',t(key));label.append(node('span',t(key)),input);pageBarControls.append(label);return input;
  }
  const paper= node('select');paper.id='writer-paper';paper.setAttribute('aria-label',t('writerPaper'));
  for(const [value,label] of [['custom','writerCustomPaper'],['a4','writerA4'],['letter','writerLetter'],['a3','writerA3']] as const){const option=node('option',t(label));option.value=value;paper.append(option);}
  const orientation=node('select');orientation.id='writer-orientation';orientation.setAttribute('aria-label',t('writerOrientation'));
  for(const [value,key] of [['portrait','writerPortrait'],['landscape','writerLandscape']] as const){const option=node('option',t(key));option.value=value;orientation.append(option);}orientation.value=pg.width>pg.height?'landscape':'portrait';
  pageBarControls.append(paper,orientation);
  const pw=setting('writerPageWidth','writer-page-width',pg.width),ph=setting('writerPageHeight','writer-page-height',pg.height);
  const top=setting('writerMarginTop','writer-margin-top',pg.margins.top),right=setting('writerMarginRight','writer-margin-right',pg.margins.right),bottom=setting('writerMarginBottom','writer-margin-bottom',pg.margins.bottom),left=setting('writerMarginLeft','writer-margin-left',pg.margins.left);
  sectionChoice.onchange=()=>{if(host.composing()){sectionChoice.value=docSection.id;return;}docSection=sections.find(s=>s.id===sectionChoice.value)??sections[0]!;state.pageSectionId=docSection.id;const p=docSection.page;
    paper.value='custom';orientation.value=p.width>p.height?'landscape':'portrait';for(const [input,value] of [[pw,p.width],[ph,p.height],[top,p.margins.top],[right,p.margins.right],[bottom,p.margins.bottom],[left,p.margins.left]] as const)input.value=String(Math.round(value/EMU_PER_TWIP)/20);
  };
  tool('writerSectionBreak','writer-section-break',()=>{const sel=currentSelection();if(!sel||host.composing()||contextTable())return;const split=splitWriterSectionAfter(host.snapshot(),sel.id,crypto.randomUUID());if(host.execute(split.command)){state.pageSectionId=split.sectionId;state.paragraphId=split.paragraphId||undefined;state.start=state.end=0;state.focus=true;host.render();}},false,()=>!contextTable(),pageBarControls);
  const presets:Readonly<Record<string,readonly [number,number]>>={a4:[11906,16838],letter:[12240,15840],a3:[16838,23811]};
  paper.onchange=()=>{const pair=presets[paper.value];if(pair){const [w,h]=orientation.value==='landscape'?[pair[1],pair[0]]:pair;pw.value=String(w/20);ph.value=String(h/20);}};
  orientation.onchange=()=>{if((pw.valueAsNumber>ph.valueAsNumber)!==(orientation.value==='landscape')){const v=pw.value;pw.value=ph.value;ph.value=v;}};
  const emu=(input:HTMLInputElement)=>Math.round(input.valueAsNumber*20)*EMU_PER_TWIP;
  const apply=button(t('writerApplyPage'),()=>safe(()=>{
    if(host.composing())return;
    const values={width:emu(pw),height:emu(ph),margins:{top:emu(top),right:emu(right),bottom:emu(bottom),left:emu(left)}};
    if(host.execute(setWriterPage(host.snapshot(),docSection.id,values)))host.render();
  }),'writer-apply-page');pageBarControls.append(apply);
  root.insertBefore(pageBarControls,note);
  const metricsBar=node('div','','writer-table-controls');
  for(const [key,label] of [['indent','writerIndent'],['spaceBefore','writerSpaceBefore'],['spaceAfter','writerSpaceAfter'],['lineHeight','writerLineHeight']] as const){
    const wrapper=node('label','','writer-format-field'),input=node('input');input.type='number';input.min=key==='indent'?'-1584':'0';input.max='1584';input.step='0.05';input.dataset.metric=key;input.id='writer-'+key;input.setAttribute('aria-label',t(label));
    wrapper.append(node('span',t(label)),input);metricsBar.append(wrapper);metricFields.push(input);
    input.onchange=()=>safe(()=>{const sel=currentSelection();if(sel&&!host.composing()&&host.execute(setParagraphMetrics(host.snapshot(),sel.id,{[key]:emu(input)}))){select(sel.id,sel.start,sel.end,true);scheduleLayout();}});
  }
  root.insertBefore(metricsBar,note);
  function join(id:string,direction:'previous'|'next'):void{
    const d=modelOf(host.snapshot());
    const groups:DeepReadonly<WriterBlock[]>[]=[];
    for(const section of d.sections){
      groups.push(section.blocks);
      for(const b of section.blocks)if(b.type==='table')for(const r of b.rows)for(const c of r.cells)groups.push(c.blocks);
    }
    for(const blocks of groups){
      const index=blocks.findIndex(b=>b.id===id);if(index<0)continue;
      const a=direction==='previous'?blocks[index-1]:blocks[index],b=direction==='previous'?blocks[index]:blocks[index+1];
      if(a?.type!=='paragraph'||b?.type!=='paragraph')return;
      const offset=paragraphText(host.snapshot(),a.id).length;
      if(host.execute(joinParagraphs(host.snapshot(),a.id,b.id))){state.paragraphId=a.id;state.start=offset;state.end=offset;state.focus=true;host.render();}return;
    }
  }
  let index=0;
  function renderImageCard(image:DeepReadonly<Extract<Paragraph['runs'][number],{type:'image'}>>,parent:HTMLElement):void{
      const r=modelOf(host.snapshot()).resources[image.resourceId],card=node('div','','writer-image-card');card.dataset.imageId=image.id;
      if(r?.dataBase64&&(r.mimeType==='image/png'||r.mimeType==='image/jpeg')){const preview=node('img');preview.alt=t('writerImage');preview.src='data:'+r.mimeType+';base64,'+r.dataBase64;preview.width=Math.max(1,Math.min(480,Math.round(image.width/EMU_PER_CSS_PIXEL)));card.append(preview);}
      const label=node('label','','writer-format-field'),width=node('input');width.type='number';width.min='1';width.max='1584';width.step='0.05';width.value=String(Math.round(image.width/EMU_PER_TWIP)/20);width.setAttribute('aria-label',t('writerImageWidth'));width.className='writer-image-width';
      label.append(node('span',t('writerImageWidth')),width);
      width.onchange=()=>safe(()=>{if(host.composing())return;const n=Math.round(width.valueAsNumber*20)*EMU_PER_TWIP;if(host.execute(resizeInlineImage(host.snapshot(),image.id,n)))host.render();});
      const remove=button(t('writerRemoveImage'),()=>safe(()=>{if(!host.composing()&&host.execute(removeInlineImage(host.snapshot(),image.id)))host.render();}),'writer-remove-image');
      const actions=node('div','','writer-table-controls');actions.append(label,remove);card.append(actions);parent.append(card);
  }
  function renderParagraph(p:DeepReadonly<WriterBlock>,parent:HTMLElement,cell?:{row:number;column:number}):void{
    if(p.type==='paragraph'&&!cell&&p.runs.length===1&&p.runs[0]?.type==='image'){renderImageCard(p.runs[0],parent);return;}
    if(p.type!=='paragraph'){
      const indicator=node('p',t(p.type==='page-break'?'writerPageBreak':'writerUnsupportedBlock'),'writer-block-note');parent.append(indicator);return;
    }
    const input=node('textarea','','writer-paragraph draft-text');input.id=inputId(index++);input.dataset.paragraphId=p.id;
    const d=modelOf(host.snapshot()),named=p.styleId??d.defaultParagraphStyleId,style=resolveTextStyle(d.defaultTextStyle,d.styles,named?[named]:[],p.style??{}).style;
    input.dir=style.paragraphDirection??'auto';if(style.language)input.lang=style.language;input.rows=cell?2:3;input.setAttribute('aria-label',cell?t('writerCell',{row:host.number(cell.row+1),column:host.number(cell.column+1)})+' · '+host.number(index):t('textLabel')+' '+host.number(index));input.value=paragraphText(host.snapshot(),p.id);
    if(cell){input.classList.add('writer-cell-input');input.dataset.row=String(cell.row);input.dataset.column=String(cell.column);}
    inputs.set(p.id,input);parent.append(input);
    const pictures=p.runs.filter(r=>r.type==='image');
    if(pictures.length){parent.append(node('p',t('writerInlineImageNote'),'muted'));for(const picture of pictures)renderImageCard(picture,parent);}
    const commit=()=>{
      try{
        const s=host.snapshot(),old=paragraphText(s,p.id),change=paragraphEditDiff(old,input.value);
        if(change){if(old.slice(change.start,change.end).includes(OBJECT_REPLACEMENT)||change.replacement.includes(OBJECT_REPLACEMENT))throw new OfficeError('unsupported','Use the explicit image actions to change inline objects');
          if(!host.execute(replaceParagraphRange(s,p.id,change.start,change.end,change.replacement,crypto.randomUUID())))input.value=paragraphText(host.snapshot(),p.id);}
      }catch(error){input.value=paragraphText(host.snapshot(),p.id);host.error(error);}
      readSelection(input,p.id);scheduleLayout();
    };
    input.addEventListener('compositionstart',()=>{host.setComposing(true);syncControls();});
    input.addEventListener('compositionend',()=>{host.setComposing(false);commit();syncControls();});
    input.oninput=event=>{if(!host.composing()&&!(event as InputEvent).isComposing)commit();};
    input.onfocus=()=>readSelection(input,p.id);input.onselect=()=>readSelection(input,p.id);input.onkeyup=()=>readSelection(input,p.id);
    input.onpointerup=()=>readSelection(input,p.id);
    for(const event of ['copy','cut'] as const)input.addEventListener(event,e=>{
      if(input.value.slice(input.selectionStart,input.selectionEnd).includes(OBJECT_REPLACEMENT)){
        e.preventDefault();host.error(new OfficeError('unsupported','Plain text clipboard cannot represent images'));
      }
    });
    input.onpaste=e=>{
      if(host.composing())return;
      e.preventDefault();safe(()=>{
        const raw=e.clipboardData?.getData('text/plain');if(raw===undefined)return;
        const start=paragraphBoundary(p.id,input.selectionStart),end=paragraphBoundary(p.id,input.selectionEnd);
        if(input.value.slice(start,end).includes(OBJECT_REPLACEMENT)||raw.includes(OBJECT_REPLACEMENT))throw new OfficeError('unsupported','Plain text paste cannot replace images');
        if(!/[\r\n]/.test(raw)){
          if(host.execute(replaceParagraphRange(host.snapshot(),p.id,start,end,raw,crypto.randomUUID()))){
            input.value=paragraphText(host.snapshot(),p.id);select(p.id,start+raw.length,start+raw.length,true);scheduleLayout();
          }
        }else{
          const change=replaceWriterRange(host.snapshot(),{anchor:{paragraphId:p.id,offset:start},focus:{paragraphId:p.id,offset:end}},raw,crypto.randomUUID());
          if(host.execute(change.command)){state.paragraphId=change.caret.paragraphId;state.start=state.end=change.caret.offset;state.focus=true;host.render();}
        }
      });
    };
    input.onkeydown=e=>{
      if(host.composing()||e.isComposing||e.key==='Process'||e.ctrlKey||e.metaKey||e.altKey)return;
      if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();readSelection(input,p.id);safe(split);}
      else if(e.key==='Backspace'&&input.selectionStart===0&&input.selectionEnd===0){e.preventDefault();safe(()=>join(p.id,'previous'));}
      else if(e.key==='Delete'&&input.selectionStart===input.value.length&&input.selectionEnd===input.value.length){e.preventDefault();safe(()=>join(p.id,'next'));}
    };
  }
  function renderTable(table:DeepReadonly<Table>):void{
    const wrap=node('div','','writer-table-wrap');
    const grid=node('table','','writer-edit-table');grid.dataset.tableId=table.id;grid.setAttribute('aria-label',t('writerTable'));
    const columns=node('colgroup');const sum=table.columnWidths.reduce((a,b)=>a+b,0);
    for(const width of table.columnWidths){const col=node('col');col.style.width=String(width/sum*100)+'%';columns.append(col);}
    const body=node('tbody');grid.append(columns,body);
    for(let row=0;row<table.rows.length;row++){
      const tr=node('tr');body.append(tr);
      for(let column=0;column<table.columnWidths.length;column++){
        const cell=table.rows[row]!.cells[column]!,td=node('td');td.dataset.cellId=cell.id;tr.append(td);
        for(const p of cell.blocks)renderParagraph(p,td,{row,column});
      }
    }
    wrap.append(grid);edit.append(wrap);
  }
  for(const [i,section] of modelOf(host.snapshot()).sections.entries()){
    if(i>0)edit.append(node('p',t('writerSectionNumber',{number:host.number(i+1)}),'writer-block-note'));
    for(const p of section.blocks)if(p.type==='table')renderTable(p);else renderParagraph(p,edit);
  }
  if(!state.paragraphId||!inputs.has(state.paragraphId)){state.paragraphId=inputs.keys().next().value;state.start=0;state.end=0;}
  const initial=activeParagraph();if(initial){initial.setSelectionRange(state.start,state.end);if(state.focus){state.focus=false;requestAnimationFrame(()=>{if(initial.isConnected)initial.focus({preventScroll:true});});}}
  const continuousToggle=button(t('writerContinuous'),()=>safe(()=>{
    if(host.composing()||!state.paragraphId)return;
    if(state.continuous){state.continuous=false;state.groupRange=undefined;state.focus=true;host.render();return;}
    writerTextGroup(host.snapshot(),state.paragraphId);state.continuous=true;
    state.groupRange={anchor:{paragraphId:state.paragraphId,offset:state.start},focus:{paragraphId:state.paragraphId,offset:state.end}};host.render();
  }),'writer-continuous');
  continuousToggle.disabled=!state.paragraphId||host.composing();
  root.insertBefore(continuousToggle,controls);
  if(state.continuous&&state.paragraphId){
    try{
      let group=writerTextGroup(host.snapshot(),state.paragraphId);
      if(group.text.includes('\r'))throw new OfficeError('unsupported','Continuous native input cannot preserve raw CR');
      root.dataset.continuous='true';continuousToggle.textContent=t('writerExitContinuous');
      const panel=node('section','','writer-continuous-panel'),toolbar=node('div','','writer-controls'),input=node('textarea','','writer-continuous-input draft-text');
      input.id='writer-continuous-text';input.rows=14;input.dir='auto';input.setAttribute('aria-label',t('writerContinuous'));input.value=group.text;
      const help=node('p',t('writerContinuousNote'),'muted');panel.append(toolbar,help,input);edit.append(panel);
      function currentRange():WriterRange{
        const start=nearestBoundary(group.text,input.selectionStart),end=nearestBoundary(group.text,input.selectionEnd);
        return {anchor:textGroupPoint(group,input.selectionDirection==='backward'?end:start),focus:textGroupPoint(group,input.selectionDirection==='backward'?start:end)};
      }
      function storeRange():void{
        if(host.composing())return;
        const r=currentRange();continuousRange=r;state.groupRange=r;state.paragraphId=r.focus.paragraphId;state.start=state.end=r.focus.offset;
        follow();paint();
      }
      function setGroupRange(r:WriterRange):void{
        const a=textGroupOffset(group,r.anchor),b=textGroupOffset(group,r.focus);input.setSelectionRange(Math.min(a,b),Math.max(a,b),a>b?'backward':'forward');
        storeRange();
      }
      try{if(state.groupRange)setGroupRange(state.groupRange);else{input.setSelectionRange(0,0);storeRange();}}catch{input.setSelectionRange(0,0);storeRange();}
      continuousPick=(id,offset)=>{try{setGroupRange({anchor:{paragraphId:id,offset},focus:{paragraphId:id,offset}});input.focus({preventScroll:true});return true;}catch{return false;}};
      let captured:{start:number;end:number;revision:number}|undefined;
      input.addEventListener('beforeinput',()=>{if(!host.composing())captured={start:input.selectionStart,end:input.selectionEnd,revision:host.snapshot().revision};});
      const commit=()=>{
        try{
          const old=group.text,next=input.value;
          let diff=paragraphEditDiff(old,next);
          // Use a captured native selection when its replacement exactly
          // explains the input; this avoids ambiguity across identical runs.
          if(diff&&captured&&captured.revision===host.snapshot().revision){
            const {start,end}=captured,suffix=old.slice(end),prefix=old.slice(0,start);
            if(start<=end&&next.length>=prefix.length+suffix.length&&next.startsWith(prefix)&&next.endsWith(suffix)&&nearestBoundary(old,start)===start&&nearestBoundary(old,end)===end)
              diff={start,end,replacement:next.slice(start,next.length-suffix.length)};
          }
          captured=undefined;
          if(diff){
            const range={anchor:textGroupPoint(group,diff.start),focus:textGroupPoint(group,diff.end)};
            const change=replaceWriterRange(host.snapshot(),range,diff.replacement,crypto.randomUUID());
            if(!host.execute(change.command)){input.value=group.text;return;}
            group=writerTextGroup(host.snapshot(),change.caret.paragraphId);input.value=group.text;setGroupRange({anchor:change.caret,focus:change.caret});
          }
          storeRange();scheduleLayout();
        }catch(error){input.value=group.text;captured=undefined;host.error(error);}
      };
      input.addEventListener('compositionstart',()=>{host.setComposing(true);});
      input.addEventListener('compositionend',()=>{host.setComposing(false);commit();});
      input.oninput=e=>{if(!host.composing()&&!(e as InputEvent).isComposing)commit();};
      input.onselect=storeRange;input.onkeyup=storeRange;input.onpointerup=storeRange;
      input.addEventListener('copy',e=>{if(input.value.slice(input.selectionStart,input.selectionEnd).includes(OBJECT_REPLACEMENT)){e.preventDefault();host.error(new OfficeError('unsupported','Object marker cannot be copied as text'));}});
      for(const [key,label] of [['bold','writerBold'],['italic','writerItalic'],['underline','writerUnderline']] as const){
        const b=button(t(label),()=>safe(()=>{
          if(host.composing())return;const range=currentRange(),segments=resolveWriterRange(host.snapshot(),range).segments;
          const doc=modelOf(host.snapshot());let all=true,count=0;
          for(const seg of segments){
            const p=seg.paragraph,named=p.styleId??doc.defaultParagraphStyleId,base=resolveTextStyle(doc.defaultTextStyle,doc.styles,named?[named]:[],p.style??{}).style;let offset=0;
            for(const r of p.runs){const end=offset+inlineText(r).length;if(r.type==='text'&&offset<seg.end&&end>seg.start){count++;all=all&&resolveTextStyle(base,doc.styles,r.styleId?[r.styleId]:[],r.style??{}).style[key]===true;}offset=end;}
          }
          if(count&&host.execute(formatWriterRange(host.snapshot(),range,{[key]:!all},crypto.randomUUID()))){setGroupRange(range);input.focus({preventScroll:true});scheduleLayout();}
        }),'writer-continuous-'+key);b.onmousedown=e=>e.preventDefault();toolbar.append(b);
      }
      requestAnimationFrame(()=>{if(input.isConnected){input.focus({preventScroll:true});storeRange();}});
    }catch(error){state.continuous=false;state.groupRange=undefined;host.error(error);}
  }
  syncControls();requestAnimationFrame(updateLayout);return root;
}
