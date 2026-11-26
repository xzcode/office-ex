import { OfficeError, asFailure, EMU_PER_CSS_PIXEL, type DocumentKind, type Result } from '@office-ex/contracts';
import { DocumentSession, createDocument, validateDocument, isDocumentLanguageTag, type ImageResource, type DocumentCommand, type SaveTicket, type OfficeDocument } from '@office-ex/core';
import { writerWorkspace, type WriterViewState } from './writer-workspace.js';
import { sheetsWorkspace, releaseSheetCalculation, type SheetsViewState } from './sheets-workspace.js';
import { slidesWorkspace, type SlidesViewState } from './slides-workspace.js';
import { createPdfViewState, pdfWorkspace, disposePdfViewState, type PdfViewState } from './pdf-workspace.js';
import type { Inspection } from '@office-ex/format-service';
import type { DesktopBridge, OpenFileResult, PublicSettings, RecoveryItem, MenuAction } from './desktop-api.js';
import { availableLocales, TARGET_LOCALES, REGIONAL_LOCALES, resolveLocale, resolveRegion, localeDirection, localeLabel, isMessageKey, message as translate, diagnosticKey, errorKey, formatNumber, formatDate, type MessageKey } from '@office-ex/i18n';
import {decorateOfficeButton,officeIcon} from './office-ribbon.js';
import {xlsxValueView,type XlsxPreviewState} from './xlsx-value-preview.js';

type WorkerType='inspect'|'serialize'|'recover'|'image'|'convert-doc'|'convert-docx'|'convert-xls'|'convert-xlsx'|'convert-ppt'|'convert-pptx';
interface Tab {id:string;name:string;token?:string;session?:DocumentSession;inspection?:Inspection;xlsxPreviewState?:XlsxPreviewState;textPreviewPage?:number;busy:boolean;error?:{code:string;message:string};lastCheckpoint?:string;slidesState?:SlidesViewState;pdfState?:PdfViewState;writerState?:WriterViewState;sheetsState?:SheetsViewState;format?:'docx'|'doc'|'xls'|'xlsx'|'pptx'|'ppt'|'oex'}
const $=<T extends HTMLElement>(selector:string):T=>document.querySelector<T>(selector)!;
const bridge:DesktopBridge=window.office;
const tabs=new Map<string,Tab>();
const tasks=new Map<string,{worker:Worker;reject:(error:unknown)=>void;timer:ReturnType<typeof setTimeout>}>();
let active:string|null=null, taskCounter=0, recoveryItems:RecoveryItem[]=[], corruptRecovery=0, closing=false, composing=false;
let settings:PublicSettings={version:2,theme:'system',language:'system',region:'system',documentLanguage:'system',recent:[]};
let systemLocales:readonly string[]=['en-US'];
let appVersion='';
const uiLocale=()=>resolveLocale(settings.language,systemLocales);
const regionLocale=()=>resolveRegion(settings.region,systemLocales);
const t=(key:MessageKey,args?:Readonly<Record<string,string|number>>)=>translate(uiLocale(),key,args);
const number=(value:number,options?:Intl.NumberFormatOptions)=>formatNumber(regionLocale(),value,options);
function text<K extends keyof HTMLElementTagNameMap>(tag:K,value:string,className?:string):HTMLElementTagNameMap[K]{const el=document.createElement(tag);el.textContent=value;if(className)el.className=className;return el;}
function button(label:string,action:()=>void,className?:string):HTMLButtonElement{const el=text('button',label,className);el.type='button';el.onclick=action;return el;}
function unwrap<T>(result:Result<T>):T{if(!result.ok)throw new OfficeError(result.error.code,result.error.message);return result.value;}
function errorMessage(error:unknown):void{const failure=asFailure(error);const key:MessageKey=failure.message==='pdf.font-not-approved'||failure.message==='pdf.font-unavailable'?'pdfFontPolicy':failure.message==='pdf.encrypted'?'pdfEncrypted':failure.message==='pdf.sheets-print-rules-unavailable'?'pdfSheetsBlocked':errorKey(failure.code);notify(key,failure.code!=='cancelled',failure.code==='cancelled'?undefined:failure.message);}
let notification:{key:MessageKey;error:boolean;detail?:string}|undefined;
function message(value:string,error=false):void{notification=undefined;showMessage(value,error);}
function showMessage(value:string,error=false,detail?:string):void{const node=$('#message');node.hidden=!value;node.textContent=value;node.classList.toggle('error',error);node.title=detail??'';}
function notify(key:MessageKey,error=false,detail?:string):void{notification={key,error,detail};showMessage(t(key),error,detail);}
function run(work:()=>Promise<unknown>):void{void work().catch(errorMessage);}
function isSaving(session:DocumentSession|undefined):boolean{return session?.saveState==='saving';}
function current():Tab|undefined{return active?tabs.get(active):undefined;}
function displayTitle(tab:Tab):string{return tab.format==='docx'||tab.format==='doc'||tab.format==='xls'||tab.format==='xlsx'||tab.format==='pptx'||tab.format==='ppt'?tab.name:tab.session?.snapshot().document.title??tab.name;}
function formatLabel(tab:Tab):string{return t(tab.format==='docx'?'docxDocument':tab.format==='doc'?'docDocument':tab.format==='xls'?'xlsDocument':tab.format==='xlsx'?'xlsxDocument':tab.format==='pptx'?'pptxDocument':tab.format==='ppt'?'pptDocument':'draftLabel');}
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
type HomeSection='overview'|'new'|'recent'|'recovery';
let homeSection:HomeSection='overview';
const newButton=$<HTMLButtonElement>('#new-menu'),newOrigin=document.createComment('new menu origin'),fileActions=$('.titlebar-file-actions'),actions=$('.toolbar-actions');
newButton.before(newOrigin);
let editingShell=false;
function shellMode(editing:boolean):void{
 if(editing!==editingShell){editingShell=editing;if(editing)fileActions.prepend(newButton);else newOrigin.after(newButton);}
 $('#app').classList.toggle('office-editor-shell',editing);actions.hidden=!editing;
}
function stableIconButton(b:HTMLButtonElement,label:string,icon:string,compact=false):void{
 const identity=label+'|'+icon+'|'+compact;
 if(b.dataset.commandIdentity===identity&&b.querySelector('svg'))return;
 b.dataset.commandIdentity=identity;b.textContent=label;delete b.dataset.officeIcon;b.classList.remove('office-command','office-command-icon');decorateOfficeButton(b,icon,compact);
}
function quickButtonLabels():void{
 for(const [id,key,icon]of [['undo','undo','undo'],['redo','redo','redo'],['save-as','saveAs','save'],['save','save','save'],['open','open','folder'],['new-menu','homeNew','add'],['settings','settings','settings'],['about','productAbout','info'],['workspace-nav','workspace','home'],['recent-nav','recent','clock'],['recovery-nav','recoveries','recovery']] as const){
   const b=$<HTMLButtonElement>('#'+id);
   const label=id==='save'&&current()&&!current()?.session?t('copy'):t(key);
   stableIconButton(b,label,icon,['undo','redo','settings','about'].includes(id)||editingShell&&['open','new-menu'].includes(id));
 }
 for(const [id,section]of [['workspace-nav','overview'],['new-menu','new'],['recent-nav','recent'],['recovery-nav','recovery']] as const){
  const b=$('#'+id);if(!editingShell&&section===homeSection)b.setAttribute('aria-current','page');else b.removeAttribute('aria-current');
 }
}
function recent():void{
 const list=document.querySelector<HTMLElement>('#recent-list');if(!list)return;
 list.replaceChildren();
 const items=settings.recent.slice(0,homeSection==='recent'?20:6);
 if(!items.length){
  const empty=text('div','','home-empty');empty.append(officeIcon('clock'),text('span',t('noRecent')));list.append(empty);return;
 }
 for(const item of items){
  const ext=/\.([A-Za-z0-9]{1,8})$/.exec(item.name)?.[1]?.toLowerCase()??'';
  const type=['doc','docx'].includes(ext)?'writer':['xls','xlsx'].includes(ext)?'sheets':['ppt','pptx'].includes(ext)?'slides':ext==='pdf'?'pdf':'file';
  const b=button('',()=>run(async()=>openFile(unwrap(await bridge.openRecent(item.id)))),'recent-item home-recent-row');
  b.title=item.name;b.append(moduleMark(type,true),text('span',item.name,'home-recent-name'),text('span',ext.toUpperCase()||'—','home-recent-format'),text('time',item.lastUsedAt===undefined?'':formatDate(regionLocale(),item.lastUsedAt,{dateStyle:'medium',timeStyle:'short'}),'home-recent-date'));
  list.append(b);
 }
}
function moduleMark(kind:string,small=false):HTMLElement{
 const el=text('span','','home-module-mark');el.dataset.moduleKind=kind;el.classList.toggle('small',small);el.setAttribute('aria-hidden','true');
 el.append(officeIcon(kind==='writer'?'page':kind==='sheets'?'table':kind==='slides'?'play':kind==='pdf'?'page':'file'));
 return el;
}
function showHome(section:HomeSection):void{
 if(composing){notify('finishComposition');return;}if(!flushInputs())return;releaseSheetCalculation();active=null;homeSection=section;render();
}
function chrome():void{
  const list=$('#tablist');list.replaceChildren();
  for(const tab of tabs.values()){
    const element=document.createElement('div');element.className='tab';element.setAttribute('role','tab');element.setAttribute('aria-selected',String(tab.id===active));element.tabIndex=0;
    const title=displayTitle(tab);
    element.append(text('span',`${tab.session?.dirty?'• ':''}${title}`,'tab-title'));
    element.onclick=()=>{if(composing){notify('finishComposition');return;}if(!flushInputs())return;releaseSheetCalculation();active=tab.id;render();};element.onkeydown=e=>{if(composing){notify('finishComposition');return;}if(e.key==='Enter'||e.key===' '){e.preventDefault();if(!flushInputs())return;releaseSheetCalculation();active=tab.id;render();}else if(e.key==='ArrowRight'||e.key==='ArrowLeft'){e.preventDefault();if(!flushInputs())return;releaseSheetCalculation();const ids=[...tabs.keys()],i=ids.indexOf(tab.id);active=ids[(i+((e.key==='ArrowRight')===(localeDirection(uiLocale())==='ltr')?1:ids.length-1))%ids.length]!;render();const selected=document.querySelector<HTMLElement>('.tab[aria-selected=true]');selected?.focus();}};
    const close=button('×',()=>run(async()=>{await closeTab(tab.id);}),'tab-close');close.setAttribute('aria-label',`${t('close')} ${title}`);close.onclick=e=>{e.stopPropagation();run(async()=>{await closeTab(tab.id);});};element.append(close);list.append(element);
  }
  const tab=current(),session=tab?.session;
  shellMode(!!tab);
  $('#breadcrumb').textContent=tab?(session?`${t(session.snapshot().document.kind)} / ${formatLabel(tab)}`:t('diagnostics')):t('workspace');
  $<HTMLButtonElement>('#undo').disabled=!session?.canUndo||Boolean(tab?.busy)||session.saveState==='saving';$<HTMLButtonElement>('#redo').disabled=!session?.canRedo||Boolean(tab?.busy)||session.saveState==='saving';
  $<HTMLButtonElement>('#save').disabled=!tab||tab.busy||session?.saveState==='saving'||(!session&&!tab.inspection);
  $<HTMLButtonElement>('#save-as').disabled=!tab||tab.busy||session?.saveState==='saving'||(!session&&!tab.inspection);
  quickButtonLabels();
  $('#status').textContent=tab?.busy?t('working'):t('ready');$('#document-status').textContent=session?`${t('revision')} ${session.revision} · ${session.dirty?t('unsaved'):t('saved')}`:tab?.inspection?t('readOnly'):'';
}
function home():void{
 // The home page never carries a stale normal-save banner from a closed editor.
 if(notification?.key==='savedNote'&&!notification.error)message('');
 const content=$('#content');content.classList.add('home-content');content.replaceChildren();
 const workspace=text('div','','home-workspace');content.append(workspace);
 const hero=text('div','','home-hero');
 const title=homeSection==='overview'?t('welcome'):homeSection==='new'?t('homeNew'):homeSection==='recent'?t('recent'):t('recoveries');
 hero.append(text('h1',title));if(homeSection==='overview')hero.append(text('p',t('welcomeSub'),'intro'));workspace.append(hero);
 if(homeSection==='overview'||homeSection==='new'){
  const cards=text('div','','home-products');
  for(const [kind,desc,modern]of [['writer','writerSub','docx'],['sheets','sheetsSub','xlsx'],['slides','slidesSub','pptx']] as const){
   const card=button('',()=>run(()=>newDraft(kind,undefined,modern)),'product-card home-product');card.id='new-'+modern;card.dataset.product=kind;card.title=t(modern==='docx'?'docxNew':modern==='xlsx'?'xlsxNew':'pptxNew');
   const labels=text('span','','home-product-copy');labels.append(text('strong',t(kind)),text('small',t(desc)));
   const arrow=officeIcon('arrowNext');arrow.classList.add('home-product-arrow');card.append(moduleMark(kind),labels,arrow);cards.append(card);
  }
  workspace.append(cards);
  // Legacy and internal draft actions remain one click away under New.
  const more=text('details','','home-other-formats');more.open=homeSection==='new';
  more.append(text('summary',t('homeOtherFormats')));
  const legacy=text('div','','home-legacy-actions');
  for(const [kind,format,key]of [['writer','doc','docNew'],['sheets','xls','xlsNew'],['slides','ppt','pptNew']] as const){const b=button(t(key),()=>run(()=>newDraft(kind,undefined,format)));b.id='new-'+format;legacy.append(b);}
  more.append(legacy);
  const drafts=text('details','','internal-drafts');drafts.append(text('summary',t('draftHint')));
  const links=text('div','','draft-links');
  for(const kind of ['writer','sheets','slides'] as const){const b=button(t(kind),()=>run(()=>newDraft(kind)),'draft-card internal-draft-card');b.dataset.kind=kind;links.append(b);}
  drafts.append(links);more.append(drafts);workspace.append(more);
 }
 if(homeSection==='overview'||homeSection==='recent'){
  const recentSection=text('section','','home-recent-section');
  const heading=text('div','','home-section-heading');if(homeSection==='overview'){heading.append(text('h2',t('recent')));if(settings.recent.length>6)heading.append(button(t('homeMore'),()=>showHome('recent'),'home-text-button'));}
  else if(settings.recent.length)heading.append(button(t('clearRecent'),()=>run(async()=>{settings=unwrap(await bridge.updateSettings({clearRecent:true}));render();notify('clearDone');}),'home-text-button'));
  recentSection.append(heading,text('div','','home-recent-list'));recentSection.lastElementChild!.id='recent-list';workspace.append(recentSection);recent();
 }
 if(homeSection==='overview'&&(recoveryItems.length||corruptRecovery)||homeSection==='recovery'){
  const recovery=text('section','','home-recovery-section'),heading=text('div','','home-section-heading');
  if(homeSection==='overview')heading.append(text('h2',t('recoveries')));recovery.append(heading);
  if(!recoveryItems.length)recovery.append(text('div',t('homeNoRecovery'),'home-empty'));
  for(const entry of recoveryItems.slice(0,homeSection==='recovery'?20:4)){
   const row=text('div','','recovery-item');row.append(text('span',t('snapshot')+' · '+formatDate(regionLocale(),entry.createdAt)+' · '+number(Math.ceil(entry.byteLength/1024))+' KB'),button(t('restore'),()=>run(()=>restore(entry))),button(t('discard'),()=>run(async()=>{unwrap(await bridge.discardRecovery(entry.id));recoveryItems=recoveryItems.filter(e=>e.id!==entry.id);render();})));recovery.append(row);
  }
  if(corruptRecovery)recovery.append(text('p',t('recoveryCorrupt'),'diagnostic error'));workspace.append(recovery);
 }
 // File drop remains available across the entire app, without a giant drop card.
}
function execute(tab:Tab,command:DocumentCommand,rerender=true):boolean{try{tab.session!.execute(command);message('');if(rerender)render();else chrome();run(syncTabs);return true;}catch(error){errorMessage(error);return false;}}
function documentView(tab:Tab):void{
  const session=tab.session!,snapshot=session.snapshot(),model=snapshot.document,content=$('#content');content.replaceChildren();
  const header=document.createElement('div');header.className='document-header';const title=document.createElement('input');title.id='draft-title';title.dir='auto';title.value=displayTitle(tab);title.maxLength=1024;title.readOnly=tab.format==='docx'||tab.format==='doc'||tab.format==='xls'||tab.format==='xlsx'||tab.format==='pptx'||tab.format==='ppt';title.setAttribute('aria-label',t('draftTitle'));
  title.onchange=()=>{if(tab.format==='docx'||tab.format==='doc'||tab.format==='xls'||tab.format==='xlsx'||tab.format==='pptx'||tab.format==='ppt')return;execute(tab,{documentId:session.documentId,expectedRevision:session.revision,type:'document.rename',payload:{title:title.value}},false);};
  header.append(title,text('span',formatLabel(tab),'status-chip'));
  if(model.kind==='writer'&&(tab.format==='doc'||tab.format==='docx')){const convert=button(t(tab.format==='doc'?'convertDocx':'convertDoc'),()=>run(()=>convertWriter(tab)));convert.id='convert-writer';header.append(convert);}
  if(model.kind==='sheets'&&(tab.format==='xls'||tab.format==='xlsx')){const convert=button(t(tab.format==='xls'?'convertXlsx':'convertXls'),()=>run(()=>convertSheets(tab)));convert.id='convert-sheets';if(tab.inspection?.sheetsTemplateConversionBlocked){convert.disabled=true;convert.title=t('conversionBoundary');}header.append(convert);}
  if(model.kind==='slides'&&(tab.format==='ppt'||tab.format==='pptx')){const convert=button(t(tab.format==='ppt'?'convertPptx':'convertPpt'),()=>run(()=>convertSlides(tab)));convert.id='convert-slides';if(tab.inspection?.slidesTemplateConversionBlocked){convert.disabled=true;convert.title=t('errorUnsupported');}header.append(convert);}
  if(model.kind==='writer'||model.kind==='slides'){const output=button(t('pdfExport'),()=>run(()=>exportCurrentPdf(tab)));output.id='export-pdf';const print=button(t('pdfPrint'),()=>run(()=>printCurrent(tab)));print.id='print-document';header.append(output,print);}for(const b of header.querySelectorAll<HTMLButtonElement>('button'))decorateOfficeButton(b,b.id==='print-document'?'print':b.id==='export-pdf'?'page':'copy');content.append(header);
  content.classList.add('office-document-content');content.dataset.editor=model.kind;
  const card=document.createElement('div');card.className='workspace-card';
  if(model.kind==='writer'){
    tab.writerState??={start:0,end:0,pageIndex:0};
    card.classList.add('writer-workspace-card');
    card.append(writerWorkspace({snapshot:()=>session.snapshot(),execute:command=>execute(tab,command,false),render,error:errorMessage,composing:()=>composing,setComposing:value=>{composing=value;},t,number,prepareImage:bytes=>task<ImageResource>(tab.id+'-image','image',{bytes}),isCurrent:()=>current()===tab,legacyWriter:tab.format==='doc'},tab.writerState));
  }else if(model.kind==='sheets'){
    tab.sheetsState??={row:1,column:1,anchorRow:1,anchorColumn:1,scrollTop:0,scrollLeft:0,inputMode:'canonical',cachedValues:tab.inspection?.sheetsCachedValues};
    card.classList.add('sheets-workspace-card');
    card.append(sheetsWorkspace({snapshot:()=>session.snapshot(),execute:command=>execute(tab,command,false),render,error:errorMessage,composing:()=>composing,setComposing:value=>{composing=value;},t,locale:regionLocale,isCurrent:()=>current()===tab,format:tab.format==='xls'||tab.format==='xlsx'?tab.format:'oex'},tab.sheetsState));
  }else{
    tab.slidesState??={selectedIds:[],zoom:0};card.classList.add('slides-workspace-card');
    card.append(slidesWorkspace({snapshot:()=>session.snapshot(),execute:command=>execute(tab,command,false),render,error:errorMessage,composing:()=>composing,setComposing:value=>{composing=value;},t,number,prepareImage:bytes=>task<ImageResource>(tab.id+'-image','image',{bytes}),isCurrent:()=>current()===tab,legacyPresentation:tab.format==='ppt'},tab.slidesState));
  }
  content.append(card);
  const note=text('p',t(tab.format==='docx'?'docxNote':tab.format==='doc'?'docNote':tab.format==='xls'?'xlsNote':tab.format==='xlsx'?'xlsxNote':tab.format==='pptx'||tab.format==='ppt'?'slidesFormatNote':'draftNote'),'boundary-note');
  const technical=document.createElement('details');technical.className='office-technical';technical.append(text('summary',t('diagnostics')),note);content.append(technical);
  const details=document.createElement('div');details.className='info-row';details.append(text('span',`${t('revision')}: ${session.revision}`),text('span',`${t('history')}: ${number(session.historyUsage.items)} · ${number(Math.ceil(session.historyUsage.bytes/1024))} KB`));
  // A real browser font measurement verifies the native provider without claiming pagination fidelity.
  technical.append(details);
}
function inspectionView(tab:Tab):void{
  const content=$('#content'),view=tab.inspection!;content.replaceChildren(text('div',t(view.xlsxValuePreview?'xlsxPreviewTitle':view.officeTextPreview?'officeTextPreviewTitle':'readOnly'),'eyebrow'),text('h1',tab.name));
  const row=document.createElement('div');row.className='info-row';row.append(text('span',`${t('format')}: ${view.detection.format.toUpperCase()}`),text('span',`${t('version')}: ${view.detection.version}`));content.append(row);
  const diags=document.createElement('div');diags.className='diagnostics';for(const diagnostic of view.report.diagnostics){const label=t(diagnosticKey(diagnostic.messageKey));diags.append(text('div',`${label}  [${diagnostic.featureId}]`,`diagnostic ${diagnostic.severity==='error'?'error':''}`));}content.append(diags);
  if(view.semanticFailure){const detail=document.createElement('details');detail.className='compatibility-detail';detail.open=!view.xlsxValuePreview&&!view.officeTextPreview;detail.append(text('summary',t('semanticFailureDetail')),text('p',view.semanticFailure.detail,'muted'));content.append(detail);}
  if(view.xlsxValuePreview){const first=view.xlsxValuePreview.sheets[0]!;tab.xlsxPreviewState??={sheet:0,row:first.firstRow,column:first.firstColumn};content.append(xlsxValueView(view.xlsxValuePreview,tab.xlsxPreviewState,t));}
  if(view.officeTextPreview){const preview=document.createElement('section');preview.className='office-text-preview workspace-card';preview.id='office-text-preview';
    preview.append(text('p',t('officeTextPreviewNotice'),'boundary-note'));if(view.officeTextPreview.truncated)preview.append(text('p',t('xlsxPreviewLimited'),'diagnostic'));
    const pages=view.officeTextPreview.pages;let activePage=Math.min(pages.length-1,tab.textPreviewPage??0);const body=document.createElement('div');body.className='office-text-body';
    const select=document.createElement('select');select.id='office-text-preview-page';select.setAttribute('aria-label',t('officeTextPreviewTitle'));
    pages.forEach((p,i)=>{const o=text('option',number(i+1)+(p.title?' · '+p.title:''));o.value=String(i);select.append(o);});
    const paint=()=>{const p=pages[activePage]!;body.replaceChildren();if(!p.blocks.length)body.append(text('p','—','muted'));
      for(const block of p.blocks){if(block.kind==='paragraph'){const e=text('p',block.text||' ');e.dir='auto';body.append(e);}
        else{const viewport=document.createElement('div');viewport.className='office-text-table-wrap';const table=document.createElement('table');const tb=document.createElement('tbody');for(const row of block.rows){const tr=document.createElement('tr');for(const value of row){const cell=text('td',value);cell.dir='auto';tr.append(cell);}tb.append(tr);}table.append(tb);viewport.append(table);body.append(viewport);}}
    };
    if(view.officeTextPreview.kind==='slides'){select.value=String(activePage);select.onchange=()=>{activePage=Number(select.value);tab.textPreviewPage=activePage;paint();};preview.append(select);}preview.append(body);paint();content.append(preview);
  }
  if(view.legacyDocText){content.append(text('p',t('docTextPreview'),'boundary-note'));const preview=text('pre',view.legacyDocText.text.replace(/\r/g,'\n').replace(/[\x00-\x08\x0b-\x1f]/g,'�').slice(0,200000));preview.style.whiteSpace='pre-wrap';preview.dir='auto';preview.className='workspace-card';content.append(preview);}
  const parts=document.createElement('details');parts.className='compatibility-detail';parts.open=!view.xlsxValuePreview&&!view.officeTextPreview;
  parts.append(text('summary',`${t('parts')} (${view.entries.length}) · ${view.sha256.slice(0,16)}…`));content.append(parts);
  if(!view.entries.length){parts.append(text('p',t('noEntries'),'muted'));return;}
  const table=document.createElement('table');table.className='parts-table';const thead=document.createElement('thead'),tr=document.createElement('tr');for(const value of [t('part'),t('size'),t('digest')])tr.append(text('th',value));thead.append(tr);table.append(thead);const body=document.createElement('tbody');
  for(const entry of view.entries.slice(0,500)){const tr=document.createElement('tr');tr.append(text('td',entry.name),text('td',`${number(entry.size)} B`),text('td',entry.sha256));body.append(tr);}table.append(body);parts.append(table);if(view.entries.length>500)parts.append(text('p',t('entriesLimited',{count:number(view.entries.length)}),'muted'));
}
function render():void{if(composing)return;$('#content').classList.remove('office-document-content','home-content');delete $('#content').dataset.editor;for(const previous of tabs.values()){previous.slidesState?.dispose?.();previous.pdfState?.detach?.();}applyLanguage();chrome();const tab=current();if(!tab){home();return;}if(tab.busy){const content=$('#content');const loading=document.createElement('div');loading.className='loading';loading.append(text('span','','spinner'),text('span',t('reading')),button(t('cancel'),()=>{cancelTask(tab.id);tab.busy=false;render();}));content.replaceChildren(loading);return;}if(tab.error){$('#content').replaceChildren(text('h1',tab.name),text('p',t(errorKey(tab.error.code)),'diagnostic error'));return;}if(tab.session)documentView(tab);else if(tab.pdfState){const content=$('#content');content.replaceChildren(pdfWorkspace({t,error:errorMessage,isCurrent:()=>current()===tab,print:()=>printCurrent(tab)},tab.pdfState));}else if(tab.inspection)inspectionView(tab);}
async function refreshSettings():Promise<void>{const init=unwrap(await bridge.initialize());settings=init.settings;systemLocales=init.systemLocales;applyLanguage();recent();}
async function newDraft(kind:DocumentKind,model?:OfficeDocument,format:'oex'|'docx'|'doc'|'xls'|'xlsx'|'pptx'|'ppt'='oex'):Promise<void>{if(composing){notify('finishComposition');return;}if(!flushInputs())return;releaseSheetCalculation();if(tabs.size>=32)throw new OfficeError('limit',t('limitTabs'));const id=crypto.randomUUID();const document=model?validateDocument({...model,documentId:id,title:`${model.title} (${t('restore')})`}):createDocument(kind,id,`${t('untitled')} ${t(kind)}`);if((format==='doc'||format==='docx')&&document.kind!=='writer'||(format==='xls'||format==='xlsx')&&document.kind!=='sheets'||(format==='ppt'||format==='pptx')&&document.kind!=='slides')throw new OfficeError('invalid','Document kind does not match format');if(!model&&(document.kind==='writer'||document.kind==='slides')){const preference=settings.documentLanguage;const lang=preference==='system'?systemLocales.find(isDocumentLanguageTag)??'und':preference;document.defaultTextStyle.language=lang;if(format==='doc')document.defaultTextStyle.fontFamily='Arial';}
  const tab:Tab={id,name:document.title,session:new DocumentSession(document),busy:false,format};tabs.set(id,tab);active=id;await syncTabs();message('');render();}
async function convertWriter(tab:Tab):Promise<void>{
  if(composing||!tab.session||tab.busy||tab.session.saveState==='saving')return;
  if(tabs.size>=32)throw new OfficeError('limit',t('limitTabs'));
  // Do not discard opaque auxiliary parts on cross-format conversion.
  if(tab.format==='docx'&&tab.inspection&&(tab.inspection.entries.length!==3||tab.inspection.entries.some(e=>!['[Content_Types].xml','_rels/.rels','word/document.xml'].includes(e.name))))throw new OfficeError('unsupported',t('conversionBoundary'));
  const target=tab.format==='doc'?'docx':'doc',source=tab.session.snapshot();
  const id=crypto.randomUUID();
  const model=validateDocument({...source.document,documentId:id,title:displayTitle(tab).replace(/\.docx?$/i,'')});
  await task<boolean>(tab.id+'-convert',target==='doc'?'convert-doc':'convert-docx',{document:model,documentId:id});
  if(tab.session.revision!==source.revision||!tabs.has(tab.id))throw new OfficeError('conflict','Source changed during conversion');
  const next:Tab={id,name:model.title,session:new DocumentSession(model),busy:false,format:target};tabs.set(id,next);active=id;await syncTabs();render();
}
async function convertSheets(tab:Tab):Promise<void>{
  if(tab.inspection?.sheetsTemplateConversionBlocked)throw new OfficeError('unsupported',t('conversionBoundary'));
  if(composing||!tab.session||tab.busy||tab.session.saveState==='saving'||!flushInputs())return;
  if(tabs.size>=32)throw new OfficeError('limit',t('limitTabs'));
  const source=tab.session.snapshot(),target=tab.format==='xls'?'xlsx':'xls',id=crypto.randomUUID();
  const model=validateDocument({...source.document,documentId:id,title:displayTitle(tab).replace(/\.xlsx?$/i,'')});
  await task<boolean>(tab.id+'-convert',target==='xls'?'convert-xls':'convert-xlsx',{document:model,documentId:id});
  if(tab.session.revision!==source.revision||!tabs.has(tab.id))throw new OfficeError('conflict','Source changed during conversion');
  releaseSheetCalculation();tabs.set(id,{id,name:model.title,session:new DocumentSession(model),busy:false,format:target});active=id;await syncTabs();render();
}
async function convertSlides(tab:Tab):Promise<void>{
  if(tab.inspection?.slidesTemplateConversionBlocked)throw new OfficeError('unsupported','Conversion would discard source presentation layouts; save the original PPTX instead');
  if(composing||!tab.session||tab.busy||tab.session.saveState==='saving'||!flushInputs())return;
  if(tabs.size>=32)throw new OfficeError('limit',t('limitTabs'));
  const source=tab.session.snapshot(),target=tab.format==='ppt'?'pptx':'ppt',id=crypto.randomUUID();
  const model=validateDocument({...source.document,documentId:id,title:displayTitle(tab).replace(/\.pptx?$/i,'')});
  await task<boolean>(tab.id+'-convert',target==='ppt'?'convert-ppt':'convert-pptx',{document:model,documentId:id});
  if(tab.session.revision!==source.revision||!tabs.has(tab.id))throw new OfficeError('conflict','Source changed during conversion');
  tabs.set(id,{id,name:model.title,session:new DocumentSession(model),busy:false,format:target});active=id;await syncTabs();render();
}
async function exportCurrentPdf(tab:Tab|undefined):Promise<void>{
  if(!tab?.session||tab.busy||composing||tab.session.saveState==='saving'||!flushInputs())return;
  const snapshot=tab.session.snapshot();
  if(snapshot.document.kind==='sheets')throw new OfficeError('unsupported',t('pdfSheetsBlocked'));
  await syncTabs();tab.busy=true;chrome();
  try{
    const bytes=await task<Uint8Array>(tab.id+'-export','serialize',{document:snapshot.document});
    const result=unwrap(await bridge.exportPdf({documentId:tab.id,sourceToken:tab.token,bytes,revision:snapshot.revision,contentId:snapshot.contentId,saveAs:true,title:displayTitle(tab)}));
    if(result)notify(result.warnings.length?'pdfOutputWarning':'pdfExported');
  }finally{tab.busy=false;chrome();}
}
async function printCurrent(tab:Tab|undefined):Promise<void>{
  if(!tab||tab.busy||composing||tab.session?.saveState==='saving'||!flushInputs())return;
  const snapshot=tab.session?.snapshot();
  if(snapshot?.document.kind==='sheets')throw new OfficeError('unsupported',t('pdfSheetsBlocked'));
  if(!snapshot&&!tab.pdfState)return;
  await syncTabs();tab.busy=true;chrome();
  try{
    const result=snapshot?unwrap(await bridge.print({kind:'document',documentId:tab.id,sourceToken:tab.token,revision:snapshot.revision,contentId:snapshot.contentId,bytes:await task<Uint8Array>(tab.id+'-print','serialize',{document:snapshot.document})})):unwrap(await bridge.print({kind:'pdf',documentId:tab.id,sourceToken:tab.token!,revision:0}));
    notify(result.warnings.length?'pdfOutputWarning':'pdfPrinted');
  }finally{tab.busy=false;chrome();}
}
async function openDialog():Promise<void>{if(composing){notify('finishComposition');return;}const file=unwrap(await bridge.openFile());if(file)await openFile(file);}
async function openFile(file:OpenFileResult):Promise<void>{
  if(composing||!flushInputs()){if(!file.duplicate&&!tabs.has(file.documentId))unwrap(await bridge.releaseDocument(file.documentId,false));notify('finishComposition');return;}
  releaseSheetCalculation();if(file.duplicate&&tabs.has(file.documentId)){active=file.documentId;render();return;}
  if(tabs.size>=32){unwrap(await bridge.releaseDocument(file.documentId,false));throw new OfficeError('limit',t('limitTabs'));}
  const tab:Tab={id:file.documentId,name:file.name,token:file.token,busy:true};tabs.set(tab.id,tab);active=tab.id;await syncTabs();message('');render();
  try{const inspection=await task<Inspection>(tab.id,'inspect',{bytes:file.bytes,filename:file.name,documentId:file.documentId});if(!tabs.has(tab.id))return;tab.inspection=inspection;
    if(inspection.editableWriter){const model=validateDocument(inspection.editableWriter);if(model.kind!=='writer'||model.documentId!==tab.id)throw new OfficeError('invalid','Invalid DOCX model identity');tab.session=new DocumentSession(model,{saved:true});tab.format=inspection.editableWriterFormat??'docx';}
    else if(inspection.editableSheets){const model=validateDocument(inspection.editableSheets);if(model.kind!=='sheets'||model.documentId!==tab.id)throw new OfficeError('invalid','Invalid spreadsheet model identity');tab.session=new DocumentSession(model,{saved:true});tab.format=inspection.editableSheetsFormat??'xlsx';}
    else if(inspection.editableSlides){const model=validateDocument(inspection.editableSlides);if(model.kind!=='slides'||model.documentId!==tab.id)throw new OfficeError('invalid','Invalid presentation model identity');tab.session=new DocumentSession(model,{saved:true});tab.format=inspection.editableSlidesFormat??'pptx';}
    else if(inspection.detection.format==='pdf'){tab.pdfState=createPdfViewState(file.bytes);}
    else if(inspection.draft){const model=validateDocument(inspection.draft);tab.session=new DocumentSession(model,{saved:true});tab.format='oex';}
  }catch(error){if(!tabs.has(tab.id))return;const f=asFailure(error);tab.error=f;}
  finally{tab.busy=false;if(tabs.has(tab.id)){await syncTabs();await refreshSettings();render();}}
}
function flushInputs():boolean{if(composing)return false;const focused=document.activeElement;if(focused instanceof HTMLInputElement&&focused.id==='draft-title')focused.dispatchEvent(new Event('change'));return (current()?.sheetsState?.flush?.()??true)&&(current()?.slidesState?.flush?.()??true);}
async function save(tab:Tab|undefined,saveAs=false):Promise<boolean>{
  if(!tab||tab.busy||composing||tab.session?.saveState==='saving')return false;if(!flushInputs())return false;await syncTabs();
  if(!tab.session){const result=unwrap(await bridge.saveCopy({documentId:tab.id,sourceToken:tab.token!}));if(!result)return false;notify(result.warnings?.length?'savedMaintenanceWarning':result.durability==='synced'?'savedNote':'partialDirectorySync');return true;}
  const session=tab.session;let ticket:SaveTicket|undefined;
  try{ticket=session.beginSave();chrome();const bytes=await task<Uint8Array>(`${tab.id}-save`,'serialize',{document:ticket.document});
    const request={documentId:tab.id,sourceToken:tab.token,bytes,revision:ticket.revision,contentId:ticket.contentId,saveAs,title:displayTitle(tab)};
    const result=unwrap(await (tab.format==='docx'?bridge.saveDocx(request):tab.format==='doc'?bridge.saveDoc(request):tab.format==='xlsx'?bridge.saveXlsx(request):tab.format==='xls'?bridge.saveXls(request):tab.format==='pptx'?bridge.savePptx(request):tab.format==='ppt'?bridge.savePpt(request):bridge.saveDraft(request)));
    if(!result){session.cancelSave(ticket);return false;}tab.token=result.token;tab.name=result.name;session.completeSave(ticket);tab.lastCheckpoint=ticket.contentId;
    if((tab.format==='docx'||tab.format==='doc'||tab.format==='xls'||tab.format==='xlsx'||tab.format==='pptx'||tab.format==='ppt')&&current()===tab){const title=$<HTMLInputElement>('#draft-title');if(title)title.value=tab.name;}
    await syncTabs();await refreshSettings();chrome();notify(result.warnings?.length?'savedMaintenanceWarning':result.durability==='synced'?'savedNote':'partialDirectorySync');return true;
  }catch(error){if(ticket&&session.saveState==='saving')session.cancelSave(ticket);throw error;}finally{chrome();}
}
async function closeTab(id:string):Promise<boolean>{if(composing){notify('finishComposition');return false;}const tab=tabs.get(id);if(!tab)return true;if(tab.session?.saveState==='saving'){notify('working');return false;}if(id===active&&!flushInputs())return false;await syncTabs();const decision=unwrap(await bridge.confirmClose(id));if(decision==='cancel')return false;if(decision==='save'&&!await save(tab))return false;
  if(tab.session?.dirty&&decision==='save')return false; // A new edit during save must still be reviewed.
  releaseSheetCalculation(id);cancelTask(id);cancelTask(`${id}-save`);cancelTask(`${id}-recovery`);cancelTask(`${id}-image`);cancelTask(`${id}-convert`);cancelTask(`${id}-export`);cancelTask(`${id}-print`);unwrap(await bridge.releaseDocument(id,decision==='discard'));tab.slidesState?.dispose?.();if(tab.pdfState)await disposePdfViewState(tab.pdfState);tab.session?.close();tabs.delete(id);if(active===id)active=[...tabs.keys()].at(-1)??null;await syncTabs();render();return true;}
async function closeWindow():Promise<void>{if(closing)return;closing=true;let finished=false;try{for(const id of [...tabs.keys()])if(!await closeTab(id))return;unwrap(await bridge.finishClose());finished=true;}finally{closing=false;if(!finished)await bridge.cancelWindowClose();}}
async function checkpoint(tab:Tab):Promise<void>{const session=tab.session;if(!session||!session.dirty||session.contentId===tab.lastCheckpoint||session.saveState==='saving'||tab.busy)return;const snapshot=session.snapshot(),savedMarker=session.lastSavedContentId;const bytes=await task<Uint8Array>(`${tab.id}-recovery`,'serialize',{document:snapshot.document});if(!tabs.has(tab.id)||!session.dirty||isSaving(session)||session.lastSavedContentId!==savedMarker)return;unwrap(await bridge.checkpoint({documentId:tab.id,contentId:snapshot.contentId,revision:snapshot.revision,bytes}));tab.lastCheckpoint=snapshot.contentId;}
async function restore(entry:RecoveryItem):Promise<void>{const result=unwrap(await bridge.recover(entry.id));const model=await task<OfficeDocument>('recover','recover',{bytes:result.bytes});await newDraft(model.kind,model);notify('recovered');}
async function action(value:MenuAction):Promise<void>{switch(value){case 'new-writer':await newDraft('writer');break;case 'new-sheets':await newDraft('sheets');break;case 'new-slides':await newDraft('slides');break;case 'open':await openDialog();break;case 'save':await save(current());break;case 'save-as':await save(current(),true);break;case 'export-pdf':await exportCurrentPdf(current());break;case 'print':await printCurrent(current());break;case 'close-tab':if(active)await closeTab(active);break;case 'undo':case 'redo':if(!composing&&current()?.session){if(!flushInputs())return;const s=current()!.session!;if(s.saveState!=='saving'){if(value==='undo')s.undo();else s.redo();await syncTabs();render();}}break;case 'settings':if(composing){notify('finishComposition');break;}if(!flushInputs())return;$<HTMLSelectElement>('#theme').value=settings.theme;$<HTMLSelectElement>('#language').value=settings.language;$<HTMLSelectElement>('#region').value=settings.region;$<HTMLSelectElement>('#document-language').value=settings.documentLanguage;$<HTMLDialogElement>('#settings-dialog').showModal();break;case 'close-window':await closeWindow();break;}}
$('#home').onclick=()=>showHome('overview');$('#workspace-nav').onclick=()=>showHome('overview');$('#open').onclick=()=>run(openDialog);$('#new-menu').onclick=()=>showHome('new');$('#recent-nav').onclick=()=>showHome('recent');$('#recovery-nav').onclick=()=>showHome('recovery');$('#settings').onclick=()=>run(()=>action('settings'));$('#about').onclick=()=>{if(composing){notify('finishComposition');return;}if(!flushInputs())return;$<HTMLDialogElement>('#about-dialog').showModal();};$('#undo').onclick=()=>run(()=>action('undo'));$('#redo').onclick=()=>run(()=>action('redo'));$('#save').onclick=()=>run(()=>save(current()));$('#save-as').onclick=()=>run(()=>save(current(),true));
$<HTMLSelectElement>('#theme').onchange=()=>run(async()=>{settings=unwrap(await bridge.updateSettings({theme:$<HTMLSelectElement>('#theme').value as PublicSettings['theme']}));applyTheme();});
$<HTMLSelectElement>('#language').onchange=()=>run(async()=>{if(composing){$<HTMLSelectElement>('#language').value=settings.language;notify('finishComposition');return;}if(!flushInputs())return;try{settings=unwrap(await bridge.updateSettings({language:$<HTMLSelectElement>('#language').value as PublicSettings['language']}));renderLocalized();}catch(error){$<HTMLSelectElement>('#language').value=settings.language;throw error;}});
$<HTMLSelectElement>('#region').onchange=()=>run(async()=>{if(composing){$<HTMLSelectElement>('#region').value=settings.region;notify('finishComposition');return;}if(!flushInputs())return;try{settings=unwrap(await bridge.updateSettings({region:$<HTMLSelectElement>('#region').value as PublicSettings['region']}));renderLocalized();}catch(error){$<HTMLSelectElement>('#region').value=settings.region;throw error;}});
$<HTMLSelectElement>('#document-language').onchange=()=>run(async()=>{try{settings=unwrap(await bridge.updateSettings({documentLanguage:$<HTMLSelectElement>('#document-language').value as PublicSettings['documentLanguage']}));}catch(error){$<HTMLSelectElement>('#document-language').value=settings.documentLanguage;throw error;}});
$('#clear-recent').onclick=()=>run(async()=>{settings=unwrap(await bridge.updateSettings({clearRecent:true}));recent();notify('clearDone');});
matchMedia('(prefers-color-scheme:dark)').addEventListener('change',applyTheme);
document.addEventListener('dragover',event=>{event.preventDefault();if(event.dataTransfer?.types.includes('Files'))$('#content').classList.add('file-drag');});document.addEventListener('dragleave',event=>{if(!event.relatedTarget)$('#content').classList.remove('file-drag');});
document.addEventListener('drop',event=>{event.preventDefault();$('#content').classList.remove('file-drag');const files=event.dataTransfer?.files;if(files?.length===1)run(async()=>openFile(unwrap(await bridge.openDroppedFile(files[0]!))));else if(files?.length)notify('oneDroppedFile',true);});
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
run(async()=>{const initial=unwrap(await bridge.initialize());settings=initial.settings;systemLocales=initial.systemLocales;document.documentElement.dataset.platform=initial.platform;appVersion=initial.appVersion;$('#product-version').textContent=appVersion;$('#footer-version').textContent='Office EX · '+appVersion;$('#about-version').textContent=appVersion;fillLocaleSelects();recoveryItems=initial.recovery.entries;corruptRecovery=initial.recovery.corrupt;const shortcut=document.querySelector('kbd');if(shortcut)shortcut.textContent=initial.platform==='darwin'?'⌘ O':'Ctrl O';render();document.body.dataset.ready='true';});
