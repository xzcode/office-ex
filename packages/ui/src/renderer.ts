import { OfficeError, asFailure, EMU_PER_CSS_PIXEL, type DocumentKind, type Result } from '@office-ex/contracts';
import { DocumentSession, createDocument, validateDocument, type DocumentCommand, type SaveTicket, type OfficeDocument } from '@office-ex/core';
import { replaceText } from '@office-ex/writer';
import { setCells } from '@office-ex/sheets';
import { insertShapes } from '@office-ex/slides';
import { hitTest, shapeTransform, measureText } from '@office-ex/render';
import type { Inspection } from '@office-ex/format-service';
import type { DesktopBridge, OpenFileResult, PublicSettings, RecoveryItem, MenuAction } from './desktop-api.js';
import { CanvasFontMetrics } from './font-metrics.js';

type WorkerType='inspect'|'serialize'|'recover';
interface Tab {id:string;name:string;token?:string;session?:DocumentSession;inspection?:Inspection;busy:boolean;error?:string;lastCheckpoint?:string;selectedShape?:string}
const $=<T extends HTMLElement>(selector:string):T=>document.querySelector<T>(selector)!;
const bridge:DesktopBridge=window.office;
const tabs=new Map<string,Tab>();
const tasks=new Map<string,{worker:Worker;reject:(error:unknown)=>void;timer:ReturnType<typeof setTimeout>}>();
let active:string|null=null, taskCounter=0, recoveryItems:RecoveryItem[]=[], corruptRecovery=0, closing=false, composing=false;
let settings:PublicSettings={version:1,theme:'system',language:'zh-CN',recent:[]};
const zh={open:'打开文件',newDraft:'新建草稿',recent:'最近打开',local:'本地工作 · 离线优先',settings:'设置',workspace:'工作台',undo:'撤销',redo:'重做',saveAs:'另存为',save:'保存',ready:'就绪',theme:'外观',system:'跟随系统',light:'浅色',dark:'深色',language:'语言',clearRecent:'清除最近文件记录',settingsNote:'设置和恢复快照仅保存在本机。',welcome:'你的文档，从这里开始。',welcomeSub:'一个独立开发的桌面办公套件。当前为 M1 基础设施预览，可验证会话、事务、安全保存与格式容器。',newHeading:'创建本地草稿',draftHint:'内部草稿格式 · .oex',writer:'文字文档',sheets:'电子表格',slides:'演示文稿',writerSub:'文本模型与撤销事务',sheetsSub:'稀疏单元格与批量命令',slidesSub:'场景模型与坐标基础',dropTitle:'打开或拖入一个文件',dropHint:'DOC / DOCX · XLS / XLSX · PPT / PPTX · Office EX 草稿',choose:'选择文件',boundary:'M1 的 Office 文件以只读结构检查方式打开，可原样另存。完整编辑、分页、公式计算与新旧格式转换将在后续阶段接入。',formats:'格式基础设施',containerReady:'容器已接入',modern:'现代格式',legacy:'旧版格式',draft:'本地草稿',noRecent:'暂无最近文件',untitled:'未命名',draftLabel:'Office EX 草稿',readOnly:'只读 · 结构检查',unsaved:'未保存',saved:'已保存',reading:'正在检查文件结构…',cancel:'取消',close:'关闭',cancelled:'操作已取消',snapshot:'快照',revision:'版本',savedNote:'保存完成',draftNote:'这是用于验证 M1 会话和保存流程的内部草稿。Office 格式导入、排版和导出将在后续阶段实现。',textLabel:'草稿文本',cellsLabel:'单元格事务预览（A1:C3）',cellsNote:'支持值类型和公式源码存储；此阶段尚无公式计算或完整网格编辑器。',slideLabel:'场景与命中测试预览',addShape:'添加矩形',shapeNote:'点击图形可验证坐标命中；完整编辑器在 M4 实现。',parts:'包部件 / 流',part:'部件',size:'大小',digest:'SHA-256',version:'格式版本',format:'识别格式',copy:'原样另存',pending:'语义适配尚未实现',recoveries:'可恢复的草稿',restore:'恢复副本',discard:'放弃',recoveryNote:'恢复为未保存的新副本；原文件保持不变。',recoveryCorrupt:'有损坏的恢复快照，已隔离，未载入。',recovered:'已恢复为新草稿',working:'处理中',diagnostics:'兼容性报告',noEntries:'没有可展示的容器部件。',emptySlide:'空白幻灯片',clearDone:'已清除最近文件记录',draftTitle:'文档标题',limitTabs:'最多同时打开 32 个标签',stage:'独立引擎 · 基础设施阶段',partialDirectorySync:'文件已保存；当前文件系统未提供目录同步保证。',savedMaintenanceWarning:'文件已保存，但恢复记录或最近文件记录清理失败。'};
const en:Record<keyof typeof zh,string>={open:'Open file',newDraft:'New draft',recent:'RECENT FILES',local:'Local workspace · Offline first',settings:'Settings',workspace:'Workspace',undo:'Undo',redo:'Redo',saveAs:'Save As',save:'Save',ready:'Ready',theme:'Appearance',system:'System',light:'Light',dark:'Dark',language:'Language',clearRecent:'Clear recent files',settingsNote:'Settings and recovery snapshots stay on this computer.',welcome:'A place for your documents.',welcomeSub:'An independently developed desktop office suite. This M1 preview exercises sessions, transactions, safe saves, and format containers.',newHeading:'Create a local draft',draftHint:'Internal draft format · .oex',writer:'Document',sheets:'Spreadsheet',slides:'Presentation',writerSub:'Text model & undo transactions',sheetsSub:'Sparse cells & batch commands',slidesSub:'Scene model & coordinates',dropTitle:'Open a file, or drop it here',dropHint:'DOC / DOCX · XLS / XLSX · PPT / PPTX · Office EX drafts',choose:'Choose file',boundary:'M1 opens Office files as read-only structure inspections and can save exact copies. Editing, pagination, calculation, and format conversion arrive in later milestones.',formats:'Format infrastructure',containerReady:'Container connected',modern:'Modern formats',legacy:'Legacy formats',draft:'Local drafts',noRecent:'No recent files',untitled:'Untitled',draftLabel:'Office EX draft',readOnly:'Read only · Structure inspection',unsaved:'Unsaved',saved:'Saved',reading:'Inspecting file structure…',cancel:'Cancel',close:'Close',cancelled:'Operation cancelled',snapshot:'Snapshot',revision:'Revision',savedNote:'Saved successfully',draftNote:'This internal draft verifies the M1 session and save workflow. Office import, layout, and export are delivered in later milestones.',textLabel:'Draft text',cellsLabel:'Cell transaction preview (A1:C3)',cellsNote:'Typed values and formula source can be stored. Formula calculation and the full grid editor are not implemented yet.',slideLabel:'Scene & hit-test preview',addShape:'Add rectangle',shapeNote:'Click a shape to test coordinates. The full editor arrives in M4.',parts:'Package parts / streams',part:'Part',size:'Size',digest:'SHA-256',version:'Format version',format:'Detected format',copy:'Save exact copy',pending:'Semantic adapter pending',recoveries:'Recoverable drafts',restore:'Recover a copy',discard:'Discard',recoveryNote:'Recovery creates a new unsaved copy and preserves the source file.',recoveryCorrupt:'Corrupt recovery snapshots were isolated and not loaded.',recovered:'Recovered as a new draft',working:'Working',diagnostics:'Compatibility report',noEntries:'No container parts to display.',emptySlide:'Empty slide',clearDone:'Recent file history cleared',draftTitle:'Document title',limitTabs:'At most 32 tabs can be open',stage:'Independent engine · Infrastructure',partialDirectorySync:'File saved; this filesystem did not provide directory sync guarantees.',savedMaintenanceWarning:'File saved, but recovery or recent-file bookkeeping could not be updated.'};
const t=(key:keyof typeof zh):string=>(settings.language==='en'?en:zh)[key];
function text<K extends keyof HTMLElementTagNameMap>(tag:K,value:string,className?:string):HTMLElementTagNameMap[K]{const el=document.createElement(tag);el.textContent=value;if(className)el.className=className;return el;}
function button(label:string,action:()=>void,className?:string):HTMLButtonElement{const el=text('button',label,className);el.type='button';el.onclick=action;return el;}
function unwrap<T>(result:Result<T>):T{if(!result.ok)throw new OfficeError(result.error.code,result.error.message);return result.value;}
function errorMessage(error:unknown):void{const failure=asFailure(error);if(failure.code==='cancelled')message(t('cancelled'));else message(`${failure.code}: ${failure.message}`,true);}
function message(value:string,error=false):void{const node=$('#message');node.hidden=!value;node.textContent=value;node.classList.toggle('error',error);}
function run(work:()=>Promise<unknown>):void{void work().catch(errorMessage);}
function isSaving(session:DocumentSession|undefined):boolean{return session?.saveState==='saving';}
function current():Tab|undefined{return active?tabs.get(active):undefined;}
function cancelTask(id:string):void{const task=tasks.get(id);if(task){task.worker.terminate();clearTimeout(task.timer);tasks.delete(id);task.reject(new OfficeError('cancelled','Task cancelled'));}}
function task<T>(id:string,type:WorkerType,payload:{bytes?:Uint8Array;filename?:string;document?:unknown}):Promise<T>{
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
async function syncTabs():Promise<void>{unwrap(await bridge.updateTabs([...tabs.values()].map(tab=>({id:tab.id,title:tab.session?.snapshot().document.title??tab.name,dirty:tab.session?.dirty??false,revision:tab.session?.revision??0}))));}
function applyTheme():void{const dark=settings.theme==='dark'||settings.theme==='system'&&matchMedia('(prefers-color-scheme:dark)').matches;document.documentElement.dataset.theme=dark?'dark':'light';document.documentElement.lang=settings.language;}
function applyLanguage():void{for(const node of document.querySelectorAll<HTMLElement>('[data-i18n]')){const key=node.dataset.i18n as keyof typeof zh;if(key in zh)node.textContent=t(key);}applyTheme();$('#theme').setAttribute('aria-label',t('theme'));$('#language').setAttribute('aria-label',t('language'));}
function recent():void{const list=$('#recent-list');list.replaceChildren();if(!settings.recent.length){list.append(text('div',t('noRecent'),'recent-item'));return;}for(const item of settings.recent){const b=button(item.name,()=>run(async()=>openFile(unwrap(await bridge.openRecent(item.id)))),'recent-item');b.title=item.name;list.append(b);}}
function chrome():void{
  const list=$('#tablist');list.replaceChildren();
  for(const tab of tabs.values()){
    const element=document.createElement('div');element.className='tab';element.setAttribute('role','tab');element.setAttribute('aria-selected',String(tab.id===active));element.tabIndex=0;
    const title=tab.session?.snapshot().document.title??tab.name;
    element.append(text('span',`${tab.session?.dirty?'• ':''}${title}`,'tab-title'));
    element.onclick=()=>{active=tab.id;render();};element.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();active=tab.id;render();}else if(e.key==='ArrowRight'||e.key==='ArrowLeft'){e.preventDefault();const ids=[...tabs.keys()],i=ids.indexOf(tab.id);active=ids[(i+(e.key==='ArrowRight'?1:ids.length-1))%ids.length]!;render();const selected=document.querySelector<HTMLElement>('.tab[aria-selected=true]');selected?.focus();}};
    const close=button('×',()=>run(async()=>{await closeTab(tab.id);}),'tab-close');close.setAttribute('aria-label',`${t('close')} ${title}`);close.onclick=e=>{e.stopPropagation();run(async()=>{await closeTab(tab.id);});};element.append(close);list.append(element);
  }
  const tab=current(),session=tab?.session;
  $('#breadcrumb').textContent=tab?(session?`${t(session.snapshot().document.kind)} / ${t('draftLabel')}`:t('diagnostics')):t('workspace');
  $<HTMLButtonElement>('#undo').disabled=!session?.canUndo||Boolean(tab?.busy)||session.saveState==='saving';$<HTMLButtonElement>('#redo').disabled=!session?.canRedo||Boolean(tab?.busy)||session.saveState==='saving';
  $<HTMLButtonElement>('#save').disabled=!tab||tab.busy||session?.saveState==='saving'||(!session&&!tab.inspection);$<HTMLButtonElement>('#save').textContent=tab&&!session?t('copy'):t('save');
  $<HTMLButtonElement>('#save-as').disabled=!tab||tab.busy||session?.saveState==='saving'||(!session&&!tab.inspection);
  $('#status').textContent=tab?.busy?t('working'):t('ready');$('#document-status').textContent=session?`${t('revision')} ${session.revision} · ${session.dirty?t('unsaved'):t('saved')}`:tab?.inspection?t('readOnly'):'';
}
function home():void{
  const content=$('#content');content.replaceChildren(text('div','OFFICE EX / YOUR LOCAL WORKSPACE','eyebrow'),text('h1',t('welcome')),text('p',t('welcomeSub'),'intro'));
  if(recoveryItems.length){const heading=document.createElement('div');heading.className='section-heading';heading.append(text('h2',t('recoveries')));content.append(heading,text('p',t('recoveryNote'),'muted'));
    for(const entry of recoveryItems.slice(0,12)){const row=document.createElement('div');row.className='recovery-item';row.append(text('span',`${t('snapshot')} · ${new Date(entry.createdAt).toLocaleString(settings.language)} · ${Math.ceil(entry.byteLength/1024)} KB`),button(t('restore'),()=>run(()=>restore(entry))),button(t('discard'),()=>run(async()=>{unwrap(await bridge.discardRecovery(entry.id));recoveryItems=recoveryItems.filter(e=>e.id!==entry.id);render();})));content.append(row);}}
  if(corruptRecovery)content.append(text('p',t('recoveryCorrupt'),'diagnostic error'));
  const heading=document.createElement('div');heading.className='section-heading';heading.append(text('h2',t('newHeading')),text('span',t('draftHint')));content.append(heading);
  const cards=document.createElement('div');cards.className='cards';
  for(const [kind,mark,desc] of [['writer','W','writerSub'],['sheets','S','sheetsSub'],['slides','P','slidesSub']] as const){const b=button('',()=>run(()=>newDraft(kind)),'draft-card');b.dataset.kind=kind;b.append(text('span',mark,'card-mark'),text('strong',t(kind)),text('small',t(desc)));cards.append(b);}content.append(cards);
  const drop=document.createElement('div');drop.className='drop-zone';const labels=document.createElement('div');labels.append(text('strong',t('dropTitle')),text('p',t('dropHint')));drop.append(text('span','⇧','drop-icon'),labels,button(t('choose'),()=>run(openDialog)));content.append(drop);
  const note=document.createElement('div');note.className='boundary-note';note.append(text('span','ⓘ'),text('span',t('boundary')));content.append(note);
  const formats=document.createElement('div');formats.className='section-heading';formats.append(text('h2',t('formats')),text('span',t('stage')));content.append(formats);
  const grid=document.createElement('div');grid.className='format-grid';for(const [label,value] of [['modern','DOCX · XLSX · PPTX'],['legacy','DOC · XLS · PPT'],['draft','OEX']] as const){const row=document.createElement('div');row.className='format-item';row.append(text('strong',value),text('small',t(label)));grid.append(row);}content.append(grid);
}
function execute(tab:Tab,command:DocumentCommand,rerender=true):void{try{tab.session!.execute(command);message('');if(rerender)render();else chrome();run(syncTabs);}catch(error){errorMessage(error);}}
function documentView(tab:Tab):void{
  const session=tab.session!,snapshot=session.snapshot(),model=snapshot.document,content=$('#content');content.replaceChildren();
  const header=document.createElement('div');header.className='document-header';const title=document.createElement('input');title.id='draft-title';title.value=model.title;title.maxLength=1024;title.setAttribute('aria-label',t('draftTitle'));
  title.onchange=()=>execute(tab,{documentId:session.documentId,expectedRevision:session.revision,type:'document.rename',payload:{title:title.value}},false);
  header.append(title,text('span',t('draftLabel'),'status-chip'));content.append(header);
  const card=document.createElement('div');card.className='workspace-card';
  if(model.kind==='writer'){
    card.append(text('label',t('textLabel')));const paragraph=model.sections[0]?.blocks[0];const runValue=paragraph?.type==='paragraph'?paragraph.runs[0]:undefined;
    if(runValue?.type==='text'){
      const input=document.createElement('textarea');input.className='draft-text';input.id='draft-text';input.setAttribute('aria-label',t('textLabel'));input.value=runValue.text;
      const commit=():void=>{const s=session.snapshot();if(s.document.kind!=='writer')return;const p=s.document.sections[0]?.blocks[0];const runValue=p?.type==='paragraph'?p.runs[0]:undefined;if(runValue?.type==='text')execute(tab,replaceText(s,runValue.id,0,runValue.text.length,input.value),false);};
      input.addEventListener('compositionstart',()=>{composing=true;});input.addEventListener('compositionend',()=>{composing=false;commit();});input.oninput=event=>{if(!composing&&!(event as InputEvent).isComposing)commit();};card.append(input);
    }
  }else if(model.kind==='sheets'){
    card.append(text('label',t('cellsLabel')));const grid=document.createElement('div');grid.className='mini-grid';for(const h of ['', 'A','B','C'])grid.append(text('span',h));
    const sheet=model.sheets[0]!;
    for(let row=1;row<=3;row++){grid.append(text('span',String(row)));for(let column=1;column<=3;column++){
      const value=sheet.cells[`${row}:${column}`]?.value;const input=document.createElement('input');input.setAttribute('aria-label',`${String.fromCharCode(64+column)}${row}`);input.value=value?.kind==='formula'?value.source:value&&'value'in value?String(value.value):'';
      input.onchange=()=>{const source=input.value;const cell=source===''?null:{value:source.startsWith('=')?{kind:'formula' as const,source}:/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(source)&&Number.isFinite(Number(source))?{kind:'number' as const,value:Number(source)}:{kind:'string' as const,value:source}};execute(tab,setCells(session.snapshot(),sheet.id,[{row,column,cell}]),false);};grid.append(input);
    }}card.append(grid,text('p',t('cellsNote'),'muted'));
  }else{
    card.append(text('label',t('slideLabel')));const preview=document.createElement('div');preview.className='slide-preview';const canvas=document.createElement('canvas');canvas.width=960;canvas.height=540;canvas.setAttribute('aria-label',t('slideLabel'));canvas.setAttribute('role','img');preview.append(canvas);card.append(preview);
    const ctx=canvas.getContext('2d')!,scale=canvas.width/model.size.width,slide=model.slides[0]!;ctx.fillStyle='#ffffff';ctx.fillRect(0,0,canvas.width,canvas.height);
    for(const shape of slide.shapes){const matrix=shapeTransform(shape.geometry);ctx.save();ctx.setTransform(matrix[0]*scale,matrix[1]*scale,matrix[2]*scale,matrix[3]*scale,matrix[4]*scale,matrix[5]*scale);ctx.fillStyle=shape.fill?.kind==='rgba'?`rgba(${shape.fill.red},${shape.fill.green},${shape.fill.blue},${shape.fill.alpha/255})`:'#426fca';ctx.strokeStyle=tab.selectedShape===shape.id?'#ec903f':'#3159a5';ctx.lineWidth=(tab.selectedShape===shape.id?3:1)/scale;if(shape.type==='ellipse'){ctx.beginPath();ctx.ellipse(shape.geometry.width/2,shape.geometry.height/2,shape.geometry.width/2,shape.geometry.height/2,0,0,2*Math.PI);ctx.fill();ctx.stroke();}else if(shape.type==='rect'){ctx.fillRect(0,0,shape.geometry.width,shape.geometry.height);ctx.strokeRect(0,0,shape.geometry.width,shape.geometry.height);}ctx.restore();}
    canvas.onclick=event=>{const box=canvas.getBoundingClientRect();tab.selectedShape=hitTest(slide.shapes,{x:(event.clientX-box.left)/box.width*model.size.width,y:(event.clientY-box.top)/box.height*model.size.height});render();};
    card.append(text('p',t('shapeNote'),'muted'),button(t('addShape'),()=>{const s=session.snapshot();if(s.document.kind!=='slides')return;const list=s.document.slides[0]!;execute(tab,insertShapes(s,list.id,list.shapes.length,[{id:crypto.randomUUID(),type:'rect',geometry:{x:1_000_000+(list.shapes.length%5)*600_000,y:1_000_000+(list.shapes.length%5)*400_000,width:3_000_000,height:1_500_000,rotation:0},fill:{kind:'rgba',red:67,green:111,blue:202,alpha:255}}]));}));
  }
  content.append(card,text('p',t('draftNote'),'boundary-note'));
  const details=document.createElement('div');details.className='info-row';details.append(text('span',`${t('revision')}: ${session.revision}`),text('span',`History: ${session.historyUsage.items} · ${Math.ceil(session.historyUsage.bytes/1024)} KB`));
  // A real browser font measurement verifies the native provider without claiming pagination fidelity.
  const measured=measureText(new CanvasFontMetrics(),'Office EX 中文',{fontFamily:'sans-serif',fontSize:152400});details.append(text('span',`Canvas: ${(measured.metrics.advance/EMU_PER_CSS_PIXEL).toFixed(1)} px`));content.append(details);
}
const diagnosticText:Record<string,[string,string]>={
  'format.semanticPending':['当前仅检查容器结构；尚未导入可编辑的文档语义。','Container inspection only. Editable document semantics have not been imported.'],
  'format.macrosDisabled':['检测到宏相关内容；宏不会执行，原始内容仅随原样副本保留。','Macro-related content detected. Macros do not execute; bytes are retained only in exact copies.'],
  'format.signedReadOnly':['签名文件仅作结构检查；不修改签名或声称签名有效。','Signed file: inspection only. Signature validity is not asserted.'],
  'format.externalDisabled':['外部关系已记录；不会连接或加载外部内容。','External relationships are recorded; no external content is loaded.'],
  'format.strictReadOnly':['已识别 Strict OOXML；语义读写尚未验证。','Strict OOXML identified; semantic reading/writing is unverified.'],
  'format.encryptedUnsupported':['已识别加密容器；此阶段不解密、不编辑。','Encrypted container identified. Decryption and editing are unavailable.'],
  'format.extensionMismatch':['文件扩展名与检测到的内容不一致。','The extension differs from the detected content.'],
};
function inspectionView(tab:Tab):void{
  const content=$('#content'),view=tab.inspection!;content.replaceChildren(text('div',t('readOnly'),'eyebrow'),text('h1',tab.name));
  const row=document.createElement('div');row.className='info-row';row.append(text('span',`${t('format')}: ${view.detection.format.toUpperCase()}`),text('span',`${t('version')}: ${view.detection.version}`));content.append(row);
  const diags=document.createElement('div');diags.className='diagnostics';for(const diagnostic of view.report.diagnostics){const label=diagnosticText[diagnostic.messageKey]?.[settings.language==='en'?1:0]??diagnostic.messageKey;diags.append(text('div',`${label}  [${diagnostic.featureId}]`,`diagnostic ${diagnostic.severity==='error'?'error':''}`));}content.append(diags);
  const heading=document.createElement('div');heading.className='section-heading';heading.append(text('h2',`${t('parts')} (${view.entries.length})`),text('span',`${view.sha256.slice(0,16)}…`));content.append(heading);
  if(!view.entries.length){content.append(text('p',t('noEntries'),'muted'));return;}
  const table=document.createElement('table');table.className='parts-table';const thead=document.createElement('thead'),tr=document.createElement('tr');for(const value of [t('part'),t('size'),t('digest')])tr.append(text('th',value));thead.append(tr);table.append(thead);const body=document.createElement('tbody');
  for(const entry of view.entries.slice(0,500)){const tr=document.createElement('tr');tr.append(text('td',entry.name),text('td',`${entry.size.toLocaleString()} B`),text('td',entry.sha256));body.append(tr);}table.append(body);content.append(table);if(view.entries.length>500)content.append(text('p',`${view.entries.length} entries; first 500 shown.`,'muted'));
}
function render():void{applyLanguage();recent();chrome();const tab=current();if(!tab){home();return;}if(tab.busy){const content=$('#content');const loading=document.createElement('div');loading.className='loading';loading.append(text('span','','spinner'),text('span',t('reading')),button(t('cancel'),()=>{cancelTask(tab.id);tab.busy=false;render();}));content.replaceChildren(loading);return;}if(tab.error){$('#content').replaceChildren(text('h1',tab.name),text('p',tab.error,'diagnostic error'));return;}if(tab.session)documentView(tab);else if(tab.inspection)inspectionView(tab);}
async function refreshSettings():Promise<void>{const init=unwrap(await bridge.initialize());settings=init.settings;applyLanguage();recent();}
async function newDraft(kind:DocumentKind,model?:OfficeDocument):Promise<void>{if(tabs.size>=32)throw new OfficeError('limit',t('limitTabs'));const id=crypto.randomUUID();const document=model?validateDocument({...model,documentId:id,title:`${model.title} (${t('restore')})`}):createDocument(kind,id,`${t('untitled')} ${t(kind)}`);const tab:Tab={id,name:document.title,session:new DocumentSession(document),busy:false};tabs.set(id,tab);active=id;await syncTabs();message('');render();}
async function openDialog():Promise<void>{const file=unwrap(await bridge.openFile());if(file)await openFile(file);}
async function openFile(file:OpenFileResult):Promise<void>{
  if(file.duplicate&&tabs.has(file.documentId)){active=file.documentId;render();return;}
  if(tabs.size>=32){unwrap(await bridge.releaseDocument(file.documentId,false));throw new OfficeError('limit',t('limitTabs'));}
  const tab:Tab={id:file.documentId,name:file.name,token:file.token,busy:true};tabs.set(tab.id,tab);active=tab.id;await syncTabs();message('');render();
  try{const inspection=await task<Inspection>(tab.id,'inspect',{bytes:file.bytes,filename:file.name});if(!tabs.has(tab.id))return;tab.inspection=inspection;
    if(inspection.draft){const model=validateDocument(inspection.draft);tab.session=new DocumentSession(model,{saved:true});}
  }catch(error){if(!tabs.has(tab.id))return;const f=asFailure(error);tab.error=`${f.code}: ${f.message}`;}
  finally{tab.busy=false;if(tabs.has(tab.id)){await syncTabs();await refreshSettings();render();}}
}
function flushInputs():void{const focused=document.activeElement;if(focused instanceof HTMLInputElement)focused.dispatchEvent(new Event('change'));}
async function save(tab:Tab|undefined,saveAs=false):Promise<boolean>{
  if(!tab||tab.busy||composing||tab.session?.saveState==='saving')return false;flushInputs();await syncTabs();
  if(!tab.session){const result=unwrap(await bridge.saveCopy({documentId:tab.id,sourceToken:tab.token!}));if(!result)return false;message(result.warnings?.length?t('savedMaintenanceWarning'):result.durability==='synced'?t('savedNote'):t('partialDirectorySync'));return true;}
  const session=tab.session;let ticket:SaveTicket|undefined;
  try{ticket=session.beginSave();chrome();const bytes=await task<Uint8Array>(`${tab.id}-save`,'serialize',{document:ticket.document});
    const result=unwrap(await bridge.saveDraft({documentId:tab.id,sourceToken:tab.token,bytes,revision:ticket.revision,contentId:ticket.contentId,saveAs,title:ticket.document.title}));
    if(!result){session.cancelSave(ticket);return false;}tab.token=result.token;tab.name=result.name;session.completeSave(ticket);tab.lastCheckpoint=ticket.contentId;
    await syncTabs();await refreshSettings();chrome();message(result.warnings?.length?t('savedMaintenanceWarning'):result.durability==='synced'?t('savedNote'):t('partialDirectorySync'));return true;
  }catch(error){if(ticket&&session.saveState==='saving')session.cancelSave(ticket);throw error;}finally{chrome();}
}
async function closeTab(id:string):Promise<boolean>{const tab=tabs.get(id);if(!tab)return true;if(tab.session?.saveState==='saving'){message(t('working'));return false;}flushInputs();await syncTabs();const decision=unwrap(await bridge.confirmClose(id));if(decision==='cancel')return false;if(decision==='save'&&!await save(tab))return false;
  if(tab.session?.dirty&&decision==='save')return false; // A new edit during save must still be reviewed.
  cancelTask(id);cancelTask(`${id}-save`);cancelTask(`${id}-recovery`);unwrap(await bridge.releaseDocument(id,decision==='discard'));tab.session?.close();tabs.delete(id);if(active===id)active=[...tabs.keys()].at(-1)??null;await syncTabs();render();return true;}
async function closeWindow():Promise<void>{if(closing)return;closing=true;try{for(const id of [...tabs.keys()])if(!await closeTab(id))return;unwrap(await bridge.finishClose());}finally{closing=false;}}
async function checkpoint(tab:Tab):Promise<void>{const session=tab.session;if(!session||!session.dirty||session.contentId===tab.lastCheckpoint||session.saveState==='saving'||tab.busy)return;const snapshot=session.snapshot(),savedMarker=session.lastSavedContentId;const bytes=await task<Uint8Array>(`${tab.id}-recovery`,'serialize',{document:snapshot.document});if(!tabs.has(tab.id)||!session.dirty||isSaving(session)||session.lastSavedContentId!==savedMarker)return;unwrap(await bridge.checkpoint({documentId:tab.id,contentId:snapshot.contentId,revision:snapshot.revision,bytes}));tab.lastCheckpoint=snapshot.contentId;}
async function restore(entry:RecoveryItem):Promise<void>{const result=unwrap(await bridge.recover(entry.id));const model=await task<OfficeDocument>('recover','recover',{bytes:result.bytes});await newDraft(model.kind,model);message(t('recovered'));}
async function action(value:MenuAction):Promise<void>{switch(value){case 'new-writer':await newDraft('writer');break;case 'new-sheets':await newDraft('sheets');break;case 'new-slides':await newDraft('slides');break;case 'open':await openDialog();break;case 'save':await save(current());break;case 'save-as':await save(current(),true);break;case 'close-tab':if(active)await closeTab(active);break;case 'undo':case 'redo':if(!composing&&current()?.session){flushInputs();const s=current()!.session!;if(s.saveState!=='saving'){if(value==='undo')s.undo();else s.redo();await syncTabs();render();}}break;case 'settings':$<HTMLSelectElement>('#theme').value=settings.theme;$<HTMLSelectElement>('#language').value=settings.language;$<HTMLDialogElement>('#settings-dialog').showModal();break;case 'close-window':await closeWindow();break;}}
$('#home').onclick=()=>{active=null;render();};$('#open').onclick=()=>run(openDialog);$('#new-menu').onclick=()=>{active=null;render();$('.draft-card').focus();};$('#settings').onclick=()=>run(()=>action('settings'));$('#undo').onclick=()=>run(()=>action('undo'));$('#redo').onclick=()=>run(()=>action('redo'));$('#save').onclick=()=>run(()=>save(current()));$('#save-as').onclick=()=>run(()=>save(current(),true));
$<HTMLSelectElement>('#theme').onchange=()=>run(async()=>{settings=unwrap(await bridge.updateSettings({theme:$<HTMLSelectElement>('#theme').value as PublicSettings['theme']}));applyTheme();});
$<HTMLSelectElement>('#language').onchange=()=>run(async()=>{settings=unwrap(await bridge.updateSettings({language:$<HTMLSelectElement>('#language').value as PublicSettings['language']}));render();});
$('#clear-recent').onclick=()=>run(async()=>{settings=unwrap(await bridge.updateSettings({clearRecent:true}));recent();message(t('clearDone'));});
matchMedia('(prefers-color-scheme:dark)').addEventListener('change',applyTheme);
document.addEventListener('dragover',event=>{event.preventDefault();$('.drop-zone')?.classList.add('drag-over');});document.addEventListener('dragleave',()=>$('.drop-zone')?.classList.remove('drag-over'));
document.addEventListener('drop',event=>{event.preventDefault();$('.drop-zone')?.classList.remove('drag-over');const files=event.dataTransfer?.files;if(files?.length===1)run(async()=>openFile(unwrap(await bridge.openDroppedFile(files[0]!))));else if(files?.length)errorMessage(new OfficeError('limit','Open one dropped file at a time'));});
bridge.onAction(value=>run(()=>action(value)));bridge.onFile(file=>run(()=>openFile(file)));
setInterval(()=>{for(const tab of tabs.values())run(()=>checkpoint(tab));},30_000);
run(async()=>{const initial=unwrap(await bridge.initialize());settings=initial.settings;recoveryItems=initial.recovery.entries;corruptRecovery=initial.recovery.corrupt;const shortcut=document.querySelector('kbd');if(shortcut)shortcut.textContent=initial.platform==='darwin'?'⌘ O':'Ctrl O';render();document.body.dataset.ready='true';});
