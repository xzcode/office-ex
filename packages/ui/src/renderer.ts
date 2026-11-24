import { OfficeError, asFailure, EMU_PER_CSS_PIXEL, type DocumentKind, type Result } from '@office-ex/contracts';
import { DocumentSession, createDocument, validateDocument, type DocumentCommand, type SaveTicket, type OfficeDocument } from '@office-ex/core';
import { writerWorkspace, type WriterViewState } from './writer-workspace.js';
import { setCells } from '@office-ex/sheets';
import { insertShapes } from '@office-ex/slides';
import { hitTest, shapeTransform, measureText } from '@office-ex/render';
import type { Inspection } from '@office-ex/format-service';
import type { DesktopBridge, OpenFileResult, PublicSettings, RecoveryItem, MenuAction } from './desktop-api.js';
import { availableLocales, TARGET_LOCALES, REGIONAL_LOCALES, resolveLocale, resolveRegion, localeDirection, localeLabel, isMessageKey, message as translate, diagnosticKey, errorKey, formatNumber, formatDate, type MessageKey } from '@office-ex/i18n';
import { CanvasFontMetrics } from './font-metrics.js';

type WorkerType='inspect'|'serialize'|'recover';
interface Tab {id:string;name:string;token?:string;session?:DocumentSession;inspection?:Inspection;busy:boolean;error?:{code:string;message:string};lastCheckpoint?:string;selectedShape?:string;writerState?:WriterViewState;format?:'docx'|'oex'}
const $=<T extends HTMLElement>(selector:string):T=>document.querySelector<T>(selector)!;
const bridge:DesktopBridge=window.office;
const tabs=new Map<string,Tab>();
const tasks=new Map<string,{worker:Worker;reject:(error:unknown)=>void;timer:ReturnType<typeof setTimeout>}>();
let active:string|null=null, taskCounter=0, recoveryItems:RecoveryItem[]=[], corruptRecovery=0, closing=false, composing=false;
let settings:PublicSettings={version:2,theme:'system',language:'system',region:'system',documentLanguage:'system',recent:[]};
let systemLocales:readonly string[]=['en-US'];
const uiLocale=()=>resolveLocale(settings.language,systemLocales);
const regionLocale=()=>resolveRegion(settings.region,systemLocales);
const t=(key:MessageKey,args?:Readonly<Record<string,string|number>>)=>translate(uiLocale(),key,args);
const number=(value:number,options?:Intl.NumberFormatOptions)=>formatNumber(regionLocale(),value,options);
function text<K extends keyof HTMLElementTagNameMap>(tag:K,value:string,className?:string):HTMLElementTagNameMap[K]{const el=document.createElement(tag);el.textContent=value;if(className)el.className=className;return el;}
function button(label:string,action:()=>void,className?:string):HTMLButtonElement{const el=text('button',label,className);el.type='button';el.onclick=action;return el;}
function unwrap<T>(result:Result<T>):T{if(!result.ok)throw new OfficeError(result.error.code,result.error.message);return result.value;}
function errorMessage(error:unknown):void{const failure=asFailure(error);notify(errorKey(failure.code),failure.code!=='cancelled',failure.code==='cancelled'?undefined:failure.message);}
let notification:{key:MessageKey;error:boolean;detail?:string}|undefined;
function message(value:string,error=false):void{notification=undefined;showMessage(value,error);}
function showMessage(value:string,error=false,detail?:string):void{const node=$('#message');node.hidden=!value;node.textContent=value;node.classList.toggle('error',error);node.title=detail??'';}
function notify(key:MessageKey,error=false,detail?:string):void{notification={key,error,detail};showMessage(t(key),error,detail);}
function run(work:()=>Promise<unknown>):void{void work().catch(errorMessage);}
function isSaving(session:DocumentSession|undefined):boolean{return session?.saveState==='saving';}
function current():Tab|undefined{return active?tabs.get(active):undefined;}
function displayTitle(tab:Tab):string{return tab.format==='docx'?tab.name:tab.session?.snapshot().document.title??tab.name;}
function formatLabel(tab:Tab):string{return t(tab.format==='docx'?'docxDocument':'draftLabel');}
function cancelTask(id:string):void{const task=tasks.get(id);if(task){task.worker.terminate();clearTimeout(task.timer);tasks.delete(id);task.reject(new OfficeError('cancelled','Task cancelled'));}}
function task<T>(id:string,type:WorkerType,payload:{bytes?:Uint8Array;filename?:string;documentId?:string;document?:unknown}):Promise<T>{
  cancelTask(id);
  if(tasks.size>=4)return Promise.reject(new OfficeError('limit','Worker concurrency budget exceeded'));
  return new Promise((resolve,reject)=>{
    const worker=new Worker(new URL('./format-worker.js',import.meta.url),{type:'module'}),request=++taskCounter;
    const finish=():void=>{const current=tasks.get(id);if(current?.worker===worker){clearTimeout(current.timer);tasks.delete(id);}worker.terminate();};
    const timer=setTimeout(()=>{finish();reject(new OfficeError('limit','Worker exceeded the 15 second task budget'));},15_000);
    tasks.set(id,{worker,reject,timer});
    worker.onmessage=(event:MessageEvent<Result<T>&{id:number}>)=>{if(event.data.id!==request)return;finish();try{resolve(unwrap(event.data));}catch(e){reject(e);}};
    worker.onerror=()=>{finish();reject(new OfficeError('internal','Isolated task failed'));};
    worker.postMessage({id:request,type,...payload});
  });
}
async function syncTabs():Promise<void>{unwrap(await bridge.updateTabs([...tabs.values()].map(tab=>({id:tab.id,title:displayTitle(tab),dirty:tab.session?.dirty??false,revision:tab.session?.revision??0}))));}
function applyTheme():void{const dark=settings.theme==='dark'||settings.theme==='system'&&matchMedia('(prefers-color-scheme:dark)').matches;document.documentElement.dataset.theme=dark?'dark':'light';document.documentElement.lang=uiLocale();document.documentElement.dir=localeDirection(uiLocale());}
function applyLanguage():void{
  for(const node of document.querySelectorAll<HTMLElement>('[data-i18n]')){const key=node.dataset.i18n;if(isMessageKey(key))node.textContent=t(key);}
  for(const node of document.querySelectorAll<HTMLElement>('[data-i18n-aria]')){const key=node.dataset.i18nAria;if(isMessageKey(key))node.setAttribute('aria-label',t(key));}
  for(const id of ['language','region','document-language']){const select=$<HTMLSelectElement>('#'+id);for(const option of Array.from(select.options)){if(option.value!=='system'&&option.value!=='und')option.textContent=localeLabel(option.value,uiLocale())+' ('+option.value+')';}}
  applyTheme();
  if(notification)showMessage(t(notification.key),notification.error,notification.detail);
  $('#region-sample').textContent=number(12345.67,{minimumFractionDigits:2,maximumFractionDigits:2})+' · '+formatDate(regionLocale(),new Date(2026,10,24,12),{dateStyle:'medium'});
}
function fillLocaleSelects():void{
  for(const [id,values] of [['language',availableLocales()],['region',REGIONAL_LOCALES],['document-language',TARGET_LOCALES.map(x=>x[0])]] as const){
    const select=$<HTMLSelectElement>('#'+id);select.replaceChildren();
    const system=text('option',t('system'));system.value='system';system.dataset.i18n='system';select.append(system);
    for(const locale of values){const item=text('option',localeLabel(locale,uiLocale())+' ('+locale+')');item.value=locale;select.append(item);}
    if(id==='document-language'){const item=text('option','und');item.value='und';select.append(item);}
  }
}
function renderLocalized():void{
  const content=$('#content'),scroll=content.scrollTop,left=content.scrollLeft;
  render();content.scrollTop=scroll;content.scrollLeft=left;
}
function recent():void{const list=$('#recent-list');list.replaceChildren();if(!settings.recent.length){list.append(text('div',t('noRecent'),'recent-item'));return;}for(const item of settings.recent){const b=button(item.name,()=>run(async()=>openFile(unwrap(await bridge.openRecent(item.id)))),'recent-item');b.title=item.name;list.append(b);}}
function chrome():void{
  const list=$('#tablist');list.replaceChildren();
  for(const tab of tabs.values()){
    const element=document.createElement('div');element.className='tab';element.setAttribute('role','tab');element.setAttribute('aria-selected',String(tab.id===active));element.tabIndex=0;
    const title=displayTitle(tab);
    element.append(text('span',`${tab.session?.dirty?'• ':''}${title}`,'tab-title'));
    element.onclick=()=>{if(composing){notify('finishComposition');return;}active=tab.id;render();};element.onkeydown=e=>{if(composing){notify('finishComposition');return;}if(e.key==='Enter'||e.key===' '){e.preventDefault();active=tab.id;render();}else if(e.key==='ArrowRight'||e.key==='ArrowLeft'){e.preventDefault();const ids=[...tabs.keys()],i=ids.indexOf(tab.id);active=ids[(i+((e.key==='ArrowRight')===(localeDirection(uiLocale())==='ltr')?1:ids.length-1))%ids.length]!;render();const selected=document.querySelector<HTMLElement>('.tab[aria-selected=true]');selected?.focus();}};
    const close=button('×',()=>run(async()=>{await closeTab(tab.id);}),'tab-close');close.setAttribute('aria-label',`${t('close')} ${title}`);close.onclick=e=>{e.stopPropagation();run(async()=>{await closeTab(tab.id);});};element.append(close);list.append(element);
  }
  const tab=current(),session=tab?.session;
  $('#breadcrumb').textContent=tab?(session?`${t(session.snapshot().document.kind)} / ${formatLabel(tab)}`:t('diagnostics')):t('workspace');
  $<HTMLButtonElement>('#undo').disabled=!session?.canUndo||Boolean(tab?.busy)||session.saveState==='saving';$<HTMLButtonElement>('#redo').disabled=!session?.canRedo||Boolean(tab?.busy)||session.saveState==='saving';
  $<HTMLButtonElement>('#save').disabled=!tab||tab.busy||session?.saveState==='saving'||(!session&&!tab.inspection);$<HTMLButtonElement>('#save').textContent=tab&&!session?t('copy'):t('save');
  $<HTMLButtonElement>('#save-as').disabled=!tab||tab.busy||session?.saveState==='saving'||(!session&&!tab.inspection);
  $('#status').textContent=tab?.busy?t('working'):t('ready');$('#document-status').textContent=session?`${t('revision')} ${session.revision} · ${session.dirty?t('unsaved'):t('saved')}`:tab?.inspection?t('readOnly'):'';
}
function home():void{
  const content=$('#content');content.replaceChildren(text('div',t('eyebrow'),'eyebrow'),text('h1',t('welcome')),text('p',t('welcomeSub'),'intro'));
  if(recoveryItems.length){const heading=document.createElement('div');heading.className='section-heading';heading.append(text('h2',t('recoveries')));content.append(heading,text('p',t('recoveryNote'),'muted'));
    for(const entry of recoveryItems.slice(0,12)){const row=document.createElement('div');row.className='recovery-item';row.append(text('span',`${t('snapshot')} · ${formatDate(regionLocale(),entry.createdAt)} · ${number(Math.ceil(entry.byteLength/1024))} KB`),button(t('restore'),()=>run(()=>restore(entry))),button(t('discard'),()=>run(async()=>{unwrap(await bridge.discardRecovery(entry.id));recoveryItems=recoveryItems.filter(e=>e.id!==entry.id);render();})));content.append(row);}}
  if(corruptRecovery)content.append(text('p',t('recoveryCorrupt'),'diagnostic error'));
  const heading=document.createElement('div');heading.className='section-heading';heading.append(text('h2',t('newHeading')),text('span',t('draftHint')));content.append(heading);
  const cards=document.createElement('div');cards.className='cards';
  for(const [kind,mark,desc] of [['writer','W','writerSub'],['sheets','S','sheetsSub'],['slides','P','slidesSub']] as const){const b=button('',()=>run(()=>newDraft(kind)),'draft-card');b.dataset.kind=kind;b.append(text('span',mark,'card-mark'),text('strong',t(kind)),text('small',t(desc)));cards.append(b);}content.append(cards);
  const docxActions=document.createElement('div');docxActions.className='docx-new-row';
  const createDocx=button(t('docxNew'),()=>run(()=>newDraft('writer',undefined,'docx')),'solid');createDocx.id='new-docx';
  docxActions.append(createDocx,text('span',t('docxNote'),'muted'));content.append(docxActions);
  const drop=document.createElement('div');drop.className='drop-zone';const labels=document.createElement('div');labels.append(text('strong',t('dropTitle')),text('p',t('dropHint')));drop.append(text('span','⇧','drop-icon'),labels,button(t('choose'),()=>run(openDialog)));content.append(drop);
  const note=document.createElement('div');note.className='boundary-note';note.append(text('span','ⓘ'),text('span',t('boundary')));content.append(note);
  const formats=document.createElement('div');formats.className='section-heading';formats.append(text('h2',t('formats')),text('span',t('stage')));content.append(formats);
  const grid=document.createElement('div');grid.className='format-grid';for(const [label,value] of [['modern','DOCX · XLSX · PPTX'],['legacy','DOC · XLS · PPT'],['draft','OEX']] as const){const row=document.createElement('div');row.className='format-item';row.append(text('strong',value),text('small',t(label)));grid.append(row);}content.append(grid);
}
function execute(tab:Tab,command:DocumentCommand,rerender=true):boolean{try{tab.session!.execute(command);message('');if(rerender)render();else chrome();run(syncTabs);return true;}catch(error){errorMessage(error);return false;}}
function documentView(tab:Tab):void{
  const session=tab.session!,snapshot=session.snapshot(),model=snapshot.document,content=$('#content');content.replaceChildren();
  const header=document.createElement('div');header.className='document-header';const title=document.createElement('input');title.id='draft-title';title.dir='auto';title.value=displayTitle(tab);title.maxLength=1024;title.readOnly=tab.format==='docx';title.setAttribute('aria-label',t('draftTitle'));
  title.onchange=()=>{if(tab.format==='docx')return;execute(tab,{documentId:session.documentId,expectedRevision:session.revision,type:'document.rename',payload:{title:title.value}},false);};
  header.append(title,text('span',formatLabel(tab),'status-chip'));content.append(header);
  const card=document.createElement('div');card.className='workspace-card';
  if(model.kind==='writer'){
    tab.writerState??={start:0,end:0,pageIndex:0};
    card.classList.add('writer-workspace-card');
    card.append(writerWorkspace({snapshot:()=>session.snapshot(),execute:command=>execute(tab,command,false),render,error:errorMessage,composing:()=>composing,setComposing:value=>{composing=value;},t,number},tab.writerState));
  }else if(model.kind==='sheets'){
    card.append(text('label',t('cellsLabel')));const grid=document.createElement('div');grid.className='mini-grid';for(const h of ['', 'A','B','C'])grid.append(text('span',h));
    const sheet=model.sheets[0]!;
    for(let row=1;row<=3;row++){grid.append(text('span',String(row)));for(let column=1;column<=3;column++){
      const value=sheet.cells[`${row}:${column}`]?.value;const input=document.createElement('input');input.dir='auto';input.setAttribute('aria-label',`${String.fromCharCode(64+column)}${row}`);input.value=value?.kind==='formula'?value.source:value&&'value'in value?String(value.value):'';
      input.onchange=()=>{const source=input.value;const cell=source===''?null:{value:source.startsWith('=')?{kind:'formula' as const,source}:/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(source)&&Number.isFinite(Number(source))?{kind:'number' as const,value:Number(source)}:{kind:'string' as const,value:source}};execute(tab,setCells(session.snapshot(),sheet.id,[{row,column,cell}]),false);};grid.append(input);
    }}card.append(grid,text('p',t('cellsNote'),'muted'));
  }else{
    card.append(text('label',t('slideLabel')));const preview=document.createElement('div');preview.className='slide-preview';const canvas=document.createElement('canvas');canvas.width=960;canvas.height=540;canvas.setAttribute('aria-label',t('slideLabel'));canvas.setAttribute('role','img');preview.append(canvas);card.append(preview);
    const ctx=canvas.getContext('2d')!,scale=canvas.width/model.size.width,slide=model.slides[0]!;ctx.fillStyle='#ffffff';ctx.fillRect(0,0,canvas.width,canvas.height);
    for(const shape of slide.shapes){const matrix=shapeTransform(shape.geometry);ctx.save();ctx.setTransform(matrix[0]*scale,matrix[1]*scale,matrix[2]*scale,matrix[3]*scale,matrix[4]*scale,matrix[5]*scale);ctx.fillStyle=shape.fill?.kind==='rgba'?`rgba(${shape.fill.red},${shape.fill.green},${shape.fill.blue},${shape.fill.alpha/255})`:'#426fca';ctx.strokeStyle=tab.selectedShape===shape.id?'#ec903f':'#3159a5';ctx.lineWidth=(tab.selectedShape===shape.id?3:1)/scale;if(shape.type==='ellipse'){ctx.beginPath();ctx.ellipse(shape.geometry.width/2,shape.geometry.height/2,shape.geometry.width/2,shape.geometry.height/2,0,0,2*Math.PI);ctx.fill();ctx.stroke();}else if(shape.type==='rect'){ctx.fillRect(0,0,shape.geometry.width,shape.geometry.height);ctx.strokeRect(0,0,shape.geometry.width,shape.geometry.height);}ctx.restore();}
    canvas.onclick=event=>{const box=canvas.getBoundingClientRect();tab.selectedShape=hitTest(slide.shapes,{x:(event.clientX-box.left)/box.width*model.size.width,y:(event.clientY-box.top)/box.height*model.size.height});render();};
    card.append(text('p',t('shapeNote'),'muted'),button(t('addShape'),()=>{const s=session.snapshot();if(s.document.kind!=='slides')return;const list=s.document.slides[0]!;execute(tab,insertShapes(s,list.id,list.shapes.length,[{id:crypto.randomUUID(),type:'rect',geometry:{x:1_000_000+(list.shapes.length%5)*600_000,y:1_000_000+(list.shapes.length%5)*400_000,width:3_000_000,height:1_500_000,rotation:0},fill:{kind:'rgba',red:67,green:111,blue:202,alpha:255}}]));}));
  }
  content.append(card,text('p',t(tab.format==='docx'?'docxNote':'draftNote'),'boundary-note'));
  const details=document.createElement('div');details.className='info-row';details.append(text('span',`${t('revision')}: ${session.revision}`),text('span',`${t('history')}: ${number(session.historyUsage.items)} · ${number(Math.ceil(session.historyUsage.bytes/1024))} KB`));
  // A real browser font measurement verifies the native provider without claiming pagination fidelity.
  const measured=measureText(new CanvasFontMetrics(),'Office EX 中文',{fontFamily:'sans-serif',fontSize:152400});details.append(text('span',`${t('canvas')}: ${number(measured.metrics.advance/EMU_PER_CSS_PIXEL,{maximumFractionDigits:1,minimumFractionDigits:1})} px`));content.append(details);
}
function inspectionView(tab:Tab):void{
  const content=$('#content'),view=tab.inspection!;content.replaceChildren(text('div',t('readOnly'),'eyebrow'),text('h1',tab.name));
  const row=document.createElement('div');row.className='info-row';row.append(text('span',`${t('format')}: ${view.detection.format.toUpperCase()}`),text('span',`${t('version')}: ${view.detection.version}`));content.append(row);
  const diags=document.createElement('div');diags.className='diagnostics';for(const diagnostic of view.report.diagnostics){const label=t(diagnosticKey(diagnostic.messageKey));diags.append(text('div',`${label}  [${diagnostic.featureId}]`,`diagnostic ${diagnostic.severity==='error'?'error':''}`));}content.append(diags);
  const heading=document.createElement('div');heading.className='section-heading';heading.append(text('h2',`${t('parts')} (${view.entries.length})`),text('span',`${view.sha256.slice(0,16)}…`));content.append(heading);
  if(!view.entries.length){content.append(text('p',t('noEntries'),'muted'));return;}
  const table=document.createElement('table');table.className='parts-table';const thead=document.createElement('thead'),tr=document.createElement('tr');for(const value of [t('part'),t('size'),t('digest')])tr.append(text('th',value));thead.append(tr);table.append(thead);const body=document.createElement('tbody');
  for(const entry of view.entries.slice(0,500)){const tr=document.createElement('tr');tr.append(text('td',entry.name),text('td',`${number(entry.size)} B`),text('td',entry.sha256));body.append(tr);}table.append(body);content.append(table);if(view.entries.length>500)content.append(text('p',t('entriesLimited',{count:number(view.entries.length)}),'muted'));
}
function render():void{if(composing)return;applyLanguage();recent();chrome();const tab=current();if(!tab){home();return;}if(tab.busy){const content=$('#content');const loading=document.createElement('div');loading.className='loading';loading.append(text('span','','spinner'),text('span',t('reading')),button(t('cancel'),()=>{cancelTask(tab.id);tab.busy=false;render();}));content.replaceChildren(loading);return;}if(tab.error){$('#content').replaceChildren(text('h1',tab.name),text('p',t(errorKey(tab.error.code)),'diagnostic error'));return;}if(tab.session)documentView(tab);else if(tab.inspection)inspectionView(tab);}
async function refreshSettings():Promise<void>{const init=unwrap(await bridge.initialize());settings=init.settings;systemLocales=init.systemLocales;applyLanguage();recent();}
async function newDraft(kind:DocumentKind,model?:OfficeDocument,format:'oex'|'docx'='oex'):Promise<void>{if(composing){notify('finishComposition');return;}if(tabs.size>=32)throw new OfficeError('limit',t('limitTabs'));const id=crypto.randomUUID();const document=model?validateDocument({...model,documentId:id,title:`${model.title} (${t('restore')})`}):createDocument(kind,id,`${t('untitled')} ${t(kind)}`);if(format==='docx'&&document.kind!=='writer')throw new OfficeError('invalid','DOCX requires a Writer document');const tab:Tab={id,name:document.title,session:new DocumentSession(document),busy:false,format};tabs.set(id,tab);active=id;await syncTabs();message('');render();}
async function openDialog():Promise<void>{if(composing){notify('finishComposition');return;}const file=unwrap(await bridge.openFile());if(file)await openFile(file);}
async function openFile(file:OpenFileResult):Promise<void>{
  if(composing){if(!file.duplicate&&!tabs.has(file.documentId))unwrap(await bridge.releaseDocument(file.documentId,false));notify('finishComposition');return;}
  if(file.duplicate&&tabs.has(file.documentId)){active=file.documentId;render();return;}
  if(tabs.size>=32){unwrap(await bridge.releaseDocument(file.documentId,false));throw new OfficeError('limit',t('limitTabs'));}
  const tab:Tab={id:file.documentId,name:file.name,token:file.token,busy:true};tabs.set(tab.id,tab);active=tab.id;await syncTabs();message('');render();
  try{const inspection=await task<Inspection>(tab.id,'inspect',{bytes:file.bytes,filename:file.name,documentId:file.documentId});if(!tabs.has(tab.id))return;tab.inspection=inspection;
    if(inspection.editableWriter){const model=validateDocument(inspection.editableWriter);if(model.kind!=='writer'||model.documentId!==tab.id)throw new OfficeError('invalid','Invalid DOCX model identity');tab.session=new DocumentSession(model,{saved:true});tab.format='docx';}
    else if(inspection.draft){const model=validateDocument(inspection.draft);tab.session=new DocumentSession(model,{saved:true});tab.format='oex';}
  }catch(error){if(!tabs.has(tab.id))return;const f=asFailure(error);tab.error=f;}
  finally{tab.busy=false;if(tabs.has(tab.id)){await syncTabs();await refreshSettings();render();}}
}
function flushInputs():void{const focused=document.activeElement;if(focused instanceof HTMLInputElement)focused.dispatchEvent(new Event('change'));}
async function save(tab:Tab|undefined,saveAs=false):Promise<boolean>{
  if(!tab||tab.busy||composing||tab.session?.saveState==='saving')return false;flushInputs();await syncTabs();
  if(!tab.session){const result=unwrap(await bridge.saveCopy({documentId:tab.id,sourceToken:tab.token!}));if(!result)return false;notify(result.warnings?.length?'savedMaintenanceWarning':result.durability==='synced'?'savedNote':'partialDirectorySync');return true;}
  const session=tab.session;let ticket:SaveTicket|undefined;
  try{ticket=session.beginSave();chrome();const bytes=await task<Uint8Array>(`${tab.id}-save`,'serialize',{document:ticket.document});
    const request={documentId:tab.id,sourceToken:tab.token,bytes,revision:ticket.revision,contentId:ticket.contentId,saveAs,title:displayTitle(tab)};
    const result=unwrap(await (tab.format==='docx'?bridge.saveDocx(request):bridge.saveDraft(request)));
    if(!result){session.cancelSave(ticket);return false;}tab.token=result.token;tab.name=result.name;session.completeSave(ticket);tab.lastCheckpoint=ticket.contentId;
    if(tab.format==='docx'&&current()===tab){const title=$<HTMLInputElement>('#draft-title');if(title)title.value=tab.name;}
    await syncTabs();await refreshSettings();chrome();notify(result.warnings?.length?'savedMaintenanceWarning':result.durability==='synced'?'savedNote':'partialDirectorySync');return true;
  }catch(error){if(ticket&&session.saveState==='saving')session.cancelSave(ticket);throw error;}finally{chrome();}
}
async function closeTab(id:string):Promise<boolean>{if(composing){notify('finishComposition');return false;}const tab=tabs.get(id);if(!tab)return true;if(tab.session?.saveState==='saving'){notify('working');return false;}flushInputs();await syncTabs();const decision=unwrap(await bridge.confirmClose(id));if(decision==='cancel')return false;if(decision==='save'&&!await save(tab))return false;
  if(tab.session?.dirty&&decision==='save')return false; // A new edit during save must still be reviewed.
  cancelTask(id);cancelTask(`${id}-save`);cancelTask(`${id}-recovery`);unwrap(await bridge.releaseDocument(id,decision==='discard'));tab.session?.close();tabs.delete(id);if(active===id)active=[...tabs.keys()].at(-1)??null;await syncTabs();render();return true;}
async function closeWindow():Promise<void>{if(closing)return;closing=true;try{for(const id of [...tabs.keys()])if(!await closeTab(id))return;unwrap(await bridge.finishClose());}finally{closing=false;}}
async function checkpoint(tab:Tab):Promise<void>{const session=tab.session;if(!session||!session.dirty||session.contentId===tab.lastCheckpoint||session.saveState==='saving'||tab.busy)return;const snapshot=session.snapshot(),savedMarker=session.lastSavedContentId;const bytes=await task<Uint8Array>(`${tab.id}-recovery`,'serialize',{document:snapshot.document});if(!tabs.has(tab.id)||!session.dirty||isSaving(session)||session.lastSavedContentId!==savedMarker)return;unwrap(await bridge.checkpoint({documentId:tab.id,contentId:snapshot.contentId,revision:snapshot.revision,bytes}));tab.lastCheckpoint=snapshot.contentId;}
async function restore(entry:RecoveryItem):Promise<void>{const result=unwrap(await bridge.recover(entry.id));const model=await task<OfficeDocument>('recover','recover',{bytes:result.bytes});await newDraft(model.kind,model);notify('recovered');}
async function action(value:MenuAction):Promise<void>{switch(value){case 'new-writer':await newDraft('writer');break;case 'new-sheets':await newDraft('sheets');break;case 'new-slides':await newDraft('slides');break;case 'open':await openDialog();break;case 'save':await save(current());break;case 'save-as':await save(current(),true);break;case 'close-tab':if(active)await closeTab(active);break;case 'undo':case 'redo':if(!composing&&current()?.session){flushInputs();const s=current()!.session!;if(s.saveState!=='saving'){if(value==='undo')s.undo();else s.redo();await syncTabs();render();}}break;case 'settings':if(composing){notify('finishComposition');break;}flushInputs();$<HTMLSelectElement>('#theme').value=settings.theme;$<HTMLSelectElement>('#language').value=settings.language;$<HTMLSelectElement>('#region').value=settings.region;$<HTMLSelectElement>('#document-language').value=settings.documentLanguage;$<HTMLDialogElement>('#settings-dialog').showModal();break;case 'close-window':await closeWindow();break;}}
$('#home').onclick=()=>{if(composing){notify('finishComposition');return;}active=null;render();};$('#open').onclick=()=>run(openDialog);$('#new-menu').onclick=()=>{if(composing){notify('finishComposition');return;}active=null;render();$('.draft-card').focus();};$('#settings').onclick=()=>run(()=>action('settings'));$('#undo').onclick=()=>run(()=>action('undo'));$('#redo').onclick=()=>run(()=>action('redo'));$('#save').onclick=()=>run(()=>save(current()));$('#save-as').onclick=()=>run(()=>save(current(),true));
$<HTMLSelectElement>('#theme').onchange=()=>run(async()=>{settings=unwrap(await bridge.updateSettings({theme:$<HTMLSelectElement>('#theme').value as PublicSettings['theme']}));applyTheme();});
$<HTMLSelectElement>('#language').onchange=()=>run(async()=>{if(composing){$<HTMLSelectElement>('#language').value=settings.language;notify('finishComposition');return;}try{settings=unwrap(await bridge.updateSettings({language:$<HTMLSelectElement>('#language').value as PublicSettings['language']}));renderLocalized();}catch(error){$<HTMLSelectElement>('#language').value=settings.language;throw error;}});
$<HTMLSelectElement>('#region').onchange=()=>run(async()=>{if(composing){$<HTMLSelectElement>('#region').value=settings.region;notify('finishComposition');return;}try{settings=unwrap(await bridge.updateSettings({region:$<HTMLSelectElement>('#region').value as PublicSettings['region']}));renderLocalized();}catch(error){$<HTMLSelectElement>('#region').value=settings.region;throw error;}});
$<HTMLSelectElement>('#document-language').onchange=()=>run(async()=>{try{settings=unwrap(await bridge.updateSettings({documentLanguage:$<HTMLSelectElement>('#document-language').value as PublicSettings['documentLanguage']}));}catch(error){$<HTMLSelectElement>('#document-language').value=settings.documentLanguage;throw error;}});
$('#clear-recent').onclick=()=>run(async()=>{settings=unwrap(await bridge.updateSettings({clearRecent:true}));recent();notify('clearDone');});
matchMedia('(prefers-color-scheme:dark)').addEventListener('change',applyTheme);
document.addEventListener('dragover',event=>{event.preventDefault();$('.drop-zone')?.classList.add('drag-over');});document.addEventListener('dragleave',()=>$('.drop-zone')?.classList.remove('drag-over'));
document.addEventListener('drop',event=>{event.preventDefault();$('.drop-zone')?.classList.remove('drag-over');const files=event.dataTransfer?.files;if(files?.length===1)run(async()=>openFile(unwrap(await bridge.openDroppedFile(files[0]!))));else if(files?.length)notify('oneDroppedFile',true);});
bridge.onAction(value=>run(()=>action(value)));bridge.onFile(file=>run(()=>openFile(file)));
let checkpointSweepRunning=false;
setInterval(()=>run(async()=>{
  if(checkpointSweepRunning)return;
  checkpointSweepRunning=true;
  try{
    // Keep at most one recovery task active; all dirty tabs can checkpoint in one
    // interval without exhausting the shared four-worker budget.
    for(const tab of [...tabs.values()]){
      if(tasks.size>=4)break;
      try{await checkpoint(tab);}catch(error){
        const failure=asFailure(error);
        if(failure.code==='cancelled'||failure.code==='conflict'&&!tab.session?.dirty)continue;
        errorMessage(error);
      }
    }
  }finally{checkpointSweepRunning=false;}
}),30_000);
run(async()=>{const initial=unwrap(await bridge.initialize());settings=initial.settings;systemLocales=initial.systemLocales;fillLocaleSelects();recoveryItems=initial.recovery.entries;corruptRecovery=initial.recovery.corrupt;const shortcut=document.querySelector('kbd');if(shortcut)shortcut.textContent=initial.platform==='darwin'?'⌘ O':'Ctrl O';render();document.body.dataset.ready='true';});
