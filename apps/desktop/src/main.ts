import { message as translate, resolveLocale, type MessageKey } from '@office-ex/i18n';
import { app, BrowserWindow, Menu, dialog, ipcMain, protocol, session, nativeTheme } from 'electron';
import type { IpcMainInvokeEvent, MenuItemConstructorOptions } from 'electron';
import { readFile, lstat, mkdir } from 'node:fs/promises';
import { randomUUID, createHash } from 'node:crypto';
import { resolve, join, basename, extname, isAbsolute } from 'node:path';
import { asFailure, OfficeError, type Result } from '@office-ex/contracts';
import type { TabState, MenuAction, OpenFileResult } from '@office-ex/ui';
import { FileAuthority, MAX_FILE_BYTES } from './file-service.js';
import { TRUSTED_APP_ORIGIN, isLocalAppUrl, assertTrustedSender, assertIpcArity, ipcId as id, parseDroppedPath, parseTabStates, parseDraftSave, parseDocxSave, parsePreservedCopy, parseRelease, parseSettingsPatch, parseRecoveryId, parseCheckpoint, parsePrintRequest, commandLineFiles } from './ipc-policy.js';
import { RecoveryStore } from './recovery.js';
import { SettingsStore } from './settings.js';
import { createDocxVerifier, type DocxVerifier } from './docx-verifier.js';
import { createPdfPrinter, type PdfPrinter } from './pdf-printer.js';
import { loadLocalResources, findLocalResource } from './local-resources.js';

const ORIGIN = TRUSTED_APP_ORIGIN;
const CSP = "default-src 'none'; script-src 'self'; style-src 'self'; img-src 'self' data:; font-src 'self' blob:; worker-src 'self'; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-src 'none'; form-action 'none'; frame-ancestors 'none'";
protocol.registerSchemesAsPrivileged([{ scheme:'office-ex', privileges:{ standard:true, secure:true, supportFetchAPI:true, corsEnabled:true } }]);
app.enableSandbox();
for (const flag of ['disable-background-networking','disable-component-update','disable-domain-reliability','no-pings']) app.commandLine.appendSwitch(flag);
app.commandLine.appendSwitch('disable-features','MediaRouter,OptimizationHints,AutofillServerCommunication,InterestFeedContentSuggestions');
if (process.env.OFFICE_EX_PROFILE_DIR && isAbsolute(process.env.OFFICE_EX_PROFILE_DIR)) app.setPath('userData',process.env.OFFICE_EX_PROFILE_DIR);
app.setName('Office EX');
const ownsProfile = app.requestSingleInstanceLock();
if(!ownsProfile)app.quit();
let window:BrowserWindow|null = null;
let allowClose = false;
let quitRequested = false;
let rendererReady = false;
let drainingFiles = false;
let createMainWindow:(()=>Promise<void>)|undefined;
let openingFiles:Promise<void>=Promise.resolve();
let queuedOpens=0;
let settings:SettingsStore;
let recovery:RecoveryStore;
let backups:RecoveryStore;
let verifyDocx:DocxVerifier;
let printPdf:PdfPrinter;
const authority = new FileAuthority();
const tabs = new Map<string,TabState>();
const originals = new Map<string,Uint8Array>();
// Imported PPTX IDs are derived from the original scene positions. Keep the
// immutable, main-owned import baseline for the lifetime of its document
// session, independently of the latest authorized disk-file bytes.
const pptxImportBaselines = new Map<string,Uint8Array>();
const pendingFiles:string[] = [];
const runningSaves = new Set<string>();
const MAX_ORIGINAL_CACHE_BYTES = 256 * 1024 * 1024;
const cacheReservations=new Map<string,number>();
function reserveCache(key:string,growth:number):void{
  const cached=[...new Set([...originals.values(),...pptxImportBaselines.values()])].reduce((sum,bytes)=>sum+bytes.byteLength,0);
  const reserved=[...cacheReservations.entries()].reduce((sum,[id,size])=>sum+(id===key?0:size),0);
  if(cached+reserved+growth>MAX_ORIGINAL_CACHE_BYTES)throw new OfficeError('limit','Open file cache exceeds the 256 MiB budget');
  cacheReservations.set(key,growth);
}
const byteHash = (value:Uint8Array):string => createHash('sha256').update(value).digest('hex');
const localOrigin = isLocalAppUrl;
const extensions = ['doc','docx','xls','xlsx','ppt','pptx','oex','pdf'];
function trusted(event:IpcMainInvokeEvent):BrowserWindow {
  const frame = event.senderFrame;
  assertTrustedSender({windowAlive:Boolean(window&&!window.isDestroyed()),expectedWebContents:Boolean(window&&event.sender===window.webContents),mainFrame:Boolean(frame&&frame===event.sender.mainFrame),frameOrigin:frame?.origin});
  return window!;
}
function handle<T>(channel:string,run:(event:IpcMainInvokeEvent,...args:unknown[])=>Promise<T>):void {
  ipcMain.handle(channel,async(event,...args):Promise<Result<T>> => {
    try { trusted(event); assertIpcArity(channel,args); return {ok:true,value:await run(event,...args)}; }
    catch(error) { return {ok:false,error:asFailure(error)}; }
  });
}
async function openSelected(path:string):Promise<OpenFileResult> {
  const owner=window;
  if(!owner||owner.isDestroyed())throw new OfficeError('conflict','No window is available');
  if(queuedOpens>=32)throw new OfficeError('limit','Too many pending file opens');
  queuedOpens++;
  const operation=openingFiles.then(()=>openSelectedNow(path,owner));
  openingFiles=operation.then(()=>undefined,()=>undefined);
  try{return await operation;}finally{queuedOpens--;}
}
async function openSelectedNow(path:string,owner:BrowserWindow):Promise<OpenFileResult>{
  if(window!==owner||owner.isDestroyed())throw new OfficeError('cancelled','The file window was closed');
  const reservation=`open:${randomUUID()}`;
  try{
    const stat=await lstat(path);
    if(!stat.isFile()||stat.isSymbolicLink())throw new OfficeError('unsafe','Only regular files can be opened');
    if(!Number.isSafeInteger(stat.size)||stat.size>MAX_FILE_BYTES)throw new OfficeError('limit','File exceeds the 100 MiB budget');
    let credited=0;
    for(const [token,bytes] of originals){
      for(const documentId of tabs.keys()){
        try{const grant=authority.get(token,owner.id,documentId);if(grant.path===resolve(path)||(grant.identity?.dev===String(stat.dev)&&grant.identity.ino===String(stat.ino))){credited=bytes.byteLength;break;}}catch{/* another session */}
      }
      if(credited)break;
    }
    // Reserve growth before the bounded read. A single open queue bounds read buffers as well.
    reserveCache(reservation,Math.max(0,stat.size-credited));
    const result=await authority.authorizeOpen(path,owner.id,randomUUID());
    if(window!==owner||owner.isDestroyed()){
      if(!result.duplicate)authority.revokeDocument(owner.id,result.documentId);
      throw new OfficeError('cancelled','The file window was closed');
    }
    if(result.duplicate&&authority.get(result.token,owner.id,result.documentId).identity?.sha256!==result.identity.sha256)throw new OfficeError('conflict','The already-open file changed outside Office EX');
    if(!result.duplicate){
      try{reserveCache(reservation,result.bytes.byteLength);}catch(error){authority.revokeDocument(owner.id,result.documentId);throw error;}
      const retained=result.bytes.slice();originals.set(result.token,retained);if(extname(result.name).toLowerCase()==='.pptx')pptxImportBaselines.set(result.documentId,retained);
    }
    cacheReservations.delete(reservation);
    // Recent-history persistence must not turn a successful file open into a failure.
    try{await settings.addRecent(authority.get(result.token,owner.id,result.documentId).path,result.name);}catch{/* document remains available */}
    return {token:result.token,documentId:result.documentId,name:result.name,bytes:result.bytes,duplicate:result.duplicate};
  }finally{cacheReservations.delete(reservation);}
}
function enqueueFile(path:string):void{
  const value=parseDroppedPath(path);
  if(pendingFiles.length>=32)throw new OfficeError('limit','Too many pending file opens');
  pendingFiles.push(value);
  void drainPendingFiles();
}
async function drainPendingFiles():Promise<void>{
  if(drainingFiles||!rendererReady||!window||window.isDestroyed())return;
  drainingFiles=true;
  try{
    while(pendingFiles.length&&rendererReady&&window&&!window.isDestroyed()){
      const owner:BrowserWindow=window,path=pendingFiles.shift()!;
      try{const file=await openSelected(path);if(window===owner&&!owner.isDestroyed())owner.webContents.send('office:file',file);}
      catch{if(window===owner&&!owner.isDestroyed())dialog.showErrorBox('Office EX',t('openFailed'));}
    }
  }finally{drainingFiles=false;}
}
function send(action:MenuAction):void { if (window && !window.isDestroyed()) window.webContents.send('office:action',action); }
const systemLanguages=():string[]=>app.isReady()?app.getPreferredSystemLanguages().slice(0,32):['en-US'];
const t=(key:MessageKey):string=>translate(resolveLocale(settings?.value().language??'system',systemLanguages()),key);
function installMenu():void {
  const menu:MenuItemConstructorOptions[] = [];
  if (process.platform === 'darwin') menu.push({role:'appMenu'});
  menu.push({label:t('menuFile'),submenu:[
    {label:t('menuNewWriter'),accelerator:'CmdOrCtrl+N',click:()=>send('new-writer')},
    {label:t('menuNewSheets'),click:()=>send('new-sheets')},
    {label:t('menuNewSlides'),click:()=>send('new-slides')},
    {type:'separator'},
    {label:t('menuOpen'),accelerator:'CmdOrCtrl+O',click:()=>send('open')},
    {label:t('save'),accelerator:'CmdOrCtrl+S',click:()=>send('save')},
    {label:t('menuSaveAs'),accelerator:'CmdOrCtrl+Shift+S',click:()=>send('save-as')},
    {label:t('pdfExport'),click:()=>send('export-pdf')},
    {label:t('pdfPrint'),accelerator:'CmdOrCtrl+P',click:()=>send('print')},
    {label:t('menuCloseTab'),accelerator:'CmdOrCtrl+W',click:()=>send('close-tab')},
    {type:'separator'},{label:t('settings'),accelerator:'CmdOrCtrl+,',click:()=>send('settings')},
  ]});
  menu.push({label:t('menuEdit'),submenu:[
    {label:t('undo'),accelerator:'CmdOrCtrl+Z',click:()=>send('undo')},
    {label:t('redo'),accelerator:process.platform==='darwin'?'Cmd+Shift+Z':'Ctrl+Y',click:()=>send('redo')},
    {type:'separator'},{role:'cut',label:t('menuCut')},{role:'copy',label:t('menuCopy')},{role:'paste',label:t('menuPaste')},{role:'selectAll',label:t('menuSelectAll')}
  ]});
  menu.push({label:t('menuView'),submenu:[{role:'togglefullscreen',label:t('menuFullscreen')},{role:'resetZoom',label:t('menuResetZoom')},{role:'zoomIn',label:t('menuZoomIn')},{role:'zoomOut',label:t('menuZoomOut')}]});
  Menu.setApplicationMenu(Menu.buildFromTemplate(menu));
}
/** Shared atomic commit path. DOCX output is created by a host-owned sandbox,
 * not trusted from the ordinary renderer. The existing OEX flow is unchanged.
 */
async function saveEditable(format:'oex'|'docx'|'doc'|'xlsx'|'xls'|'pptx'|'ppt',candidate:unknown,owner:BrowserWindow){
  const request=format!=='oex'?parseDocxSave(candidate):parseDraftSave(candidate);
  const tab=tabs.get(request.documentId);
  if(!tab)throw new OfficeError('invalid','Unknown document session');
  if(runningSaves.has(request.documentId))throw new OfficeError('conflict','Save is already running');
  if(format!=='oex'&&request.revision>tab.revision)throw new OfficeError('conflict','DOCX snapshot is newer than the registered session');
  runningSaves.add(request.documentId);
  let createdTarget:string|undefined,committed=false;
  const reservation='save:'+request.documentId;
  const assertOwner=()=>{if(window!==owner||owner.isDestroyed()||!tabs.has(request.documentId))throw new OfficeError('cancelled','The document window was closed');};
  try{
    assertOwner();
    let token=request.sourceToken;
    const grant=token?authority.get(token,owner.id,request.documentId):undefined;
    if(grant&&extname(grant.path).toLowerCase()!=='.'+format)throw new OfficeError('unsafe','Save format must match the source capability');
    let output=request.bytes;
    if(format!=='oex'){
      // Omitting an existing file capability cannot turn an imported file into
      // an unrestricted new-document save. The original is obtained only here.
      if(!token){
        for(const cachedToken of originals.keys()){
          let existing=false;
          try{authority.get(cachedToken,owner.id,request.documentId);existing=true;}catch{/* different document */}
          if(existing)throw new OfficeError('unsafe','Existing DOCX session requires its source capability');
        }
      }
      const original=token?originals.get(token):undefined;
      if(token&&(!original||!grant?.identity||byteHash(original)!==grant.identity.sha256))throw new OfficeError('conflict','DOCX source snapshot is unavailable or inconsistent');
      const semanticSource=format==='pptx'?pptxImportBaselines.get(request.documentId):original;
      if(format==='pptx'&&semanticSource&&!original)throw new OfficeError('unsafe','Imported presentation requires its current source capability');
      output=await verifyDocx({format,documentId:request.documentId,bytes:request.bytes,...(semanticSource?{sourceBytes:semanticSource}: {})});
      assertOwner();
    }
    if(!token||request.saveAs){
      const name=request.title.replace(/[<>:"/\\|?*\x00-\x1f]/g,'_').slice(0,160)||'Untitled';
      const baseName=name.toLowerCase().endsWith('.'+format)?name.slice(0,-format.length-1):name;
      const chosen=await dialog.showSaveDialog(owner,{defaultPath:baseName+'.'+format,filters:[{name:format==='docx'?t('docxDocument'):format==='doc'?t('docDocument'):format==='xlsx'?t('xlsxDocument'):format==='xls'?t('xlsDocument'):format==='pptx'?t('pptxDocument'):format==='ppt'?t('pptDocument'):t('draftLabel'),extensions:[format]}],properties:['showOverwriteConfirmation','createDirectory']});
      assertOwner();
      if(chosen.canceled||!chosen.filePath)return null;
      if(extname(chosen.filePath).toLowerCase()!=='.'+format)throw new OfficeError('invalid','Save destination must retain the selected file format');
      let exists=false;
      try{await lstat(chosen.filePath);exists=true;}catch(error){if((error as NodeJS.ErrnoException).code!=='ENOENT')throw error;}
      if(exists){
        const confirm=await dialog.showMessageBox(owner,{type:'warning',buttons:[t('cancel'),t('replace')],defaultId:0,cancelId:0,message:t('replaceTarget'),detail:basename(chosen.filePath)});
        assertOwner();if(confirm.response!==1)return null;
      }
      token=await authority.authorizeTarget(chosen.filePath,owner.id,request.documentId,exists,request.sourceToken);
      createdTarget=token;
    }
    assertOwner();
    const expected=byteHash(output),previous=originals.get(request.sourceToken??token),retainedByTemplate=previous&&[...pptxImportBaselines.values()].includes(previous),previousBytes=previous&&!retainedByTemplate?previous.byteLength:0;
    reserveCache(reservation,Math.max(0,output.byteLength-previousBytes));
    const result=await authority.commit(token,owner.id,request.documentId,output,{
      backup:async(previous,identity)=>{await backups.put(request.documentId,'backup:'+identity.sha256,request.revision,previous);},
      validate:async reread=>{
        // Verify exact bytes generated/reopened by the semantic worker before
        // the file service's final identity check and same-directory replace.
        if(byteHash(reread)!==expected)throw new OfficeError('invalid','Save verification failed');
      }
    });
    committed=true;originals.set(token,output.slice());
    if(request.sourceToken&&request.sourceToken!==token){authority.revokeToken(request.sourceToken,owner.id,request.documentId);originals.delete(request.sourceToken);}
    cacheReservations.delete(reservation);
    const warnings:string[]=[];
    try{await recovery.clearSaved(request.documentId,request.revision);}catch{warnings.push('recovery-cleanup-failed');}
    try{await settings.addRecent(authority.get(token,owner.id,request.documentId).path,result.name);}catch{warnings.push('recent-update-failed');}
    return {token,name:result.name,durability:result.durability,warnings};
  }finally{
    cacheReservations.delete(reservation);
    if(createdTarget&&!committed){try{authority.revokeToken(createdTarget,owner.id,request.documentId);}catch{/* already revoked */}}
    runningSaves.delete(request.documentId);
  }
}

async function exportPdf(candidate:unknown,owner:BrowserWindow){
  const request=parseDocxSave(candidate),tab=tabs.get(request.documentId);
  if(!tab||request.revision>tab.revision)throw new OfficeError('conflict','PDF snapshot is not registered');
  if(runningSaves.has(request.documentId))throw new OfficeError('conflict','A document output task is already running');
  if(request.sourceToken)authority.get(request.sourceToken,owner.id,request.documentId);
  runningSaves.add(request.documentId);let target:string|undefined;
  const assertOwner=()=>{if(window!==owner||owner.isDestroyed()||!tabs.has(request.documentId))throw new OfficeError('cancelled','Document window was closed');};
  try{
    const output=await printPdf({kind:'document',bytes:request.bytes,documentId:request.documentId,revision:request.revision,contentId:request.contentId},'export',owner);
    assertOwner();
    const name=request.title.replace(/[<>:"/\\|?*\x00-\x1f]/g,'_').replace(/\.(?:oex|docx?|pptx?)$/i,'').slice(0,160)||'Untitled';
    const chosen=await dialog.showSaveDialog(owner,{defaultPath:name+'.pdf',filters:[{name:t('pdfDocument'),extensions:['pdf']}],properties:['showOverwriteConfirmation','createDirectory']});
    assertOwner();if(chosen.canceled||!chosen.filePath)return null;
    if(extname(chosen.filePath).toLowerCase()!=='.pdf')throw new OfficeError('invalid','PDF destination must use the PDF extension');
    let exists=false;try{await lstat(chosen.filePath);exists=true;}catch(error){if((error as NodeJS.ErrnoException).code!=='ENOENT')throw error;}
    if(exists){const confirm=await dialog.showMessageBox(owner,{type:'warning',buttons:[t('cancel'),t('replace')],defaultId:0,cancelId:0,message:t('replaceTarget'),detail:basename(chosen.filePath)});assertOwner();if(confirm.response!==1)return null;}
    target=await authority.authorizeTarget(chosen.filePath,owner.id,request.documentId,exists,request.sourceToken);
    const bytes=output.bytes!,expected=byteHash(bytes);
    const result=await authority.commit(target,owner.id,request.documentId,bytes,{backup:async(previous,identity)=>{await backups.put(request.documentId,'backup:'+identity.sha256,request.revision,previous);},validate:async reread=>{if(byteHash(reread)!==expected)throw new OfficeError('invalid','PDF output verification failed');}});
    return {name:result.name,durability:result.durability,pages:output.pages,warnings:output.warnings};
  }finally{
    if(target){try{authority.revokeToken(target,owner.id,request.documentId);}catch{/* closed session */}}
    runningSaves.delete(request.documentId);
  }
}
async function printDocument(candidate:unknown,owner:BrowserWindow){
  const request=parsePrintRequest(candidate),tab=tabs.get(request.documentId);
  if(!tab||request.revision>tab.revision)throw new OfficeError('conflict','Print snapshot is not registered');
  if(runningSaves.has(request.documentId))throw new OfficeError('conflict','A document output task is already running');
  if(request.sourceToken)authority.get(request.sourceToken,owner.id,request.documentId);
  runningSaves.add(request.documentId);
  try{
    let bytes:Uint8Array;
    if(request.kind==='pdf'){
      const original=originals.get(request.sourceToken),grant=authority.get(request.sourceToken,owner.id,request.documentId);
      if(!original||!grant.identity||byteHash(original)!==grant.identity.sha256)throw new OfficeError('conflict','Original PDF snapshot is unavailable');
      bytes=original;
    }else bytes=request.bytes;
    const result=await printPdf({kind:request.kind,bytes,documentId:request.documentId,revision:request.revision,...(request.kind==='document'?{contentId:request.contentId}:{})},'print',owner);
    return {pages:result.pages,warnings:result.warnings};
  }finally{runningSaves.delete(request.documentId);}
}

function installIpc():void {
  handle('office:export-pdf',async(event,candidate)=>exportPdf(candidate,trusted(event)));
  handle('office:print',async(event,candidate)=>printDocument(candidate,trusted(event)));
  handle('office:initialize',async()=>{rendererReady=true;setTimeout(()=>{void drainPendingFiles();},0);return {settings:settings.publicValue(),recovery:await recovery.list(),platform:process.platform,systemLocales:systemLanguages()};});
  handle('office:open',async()=>{
    const result = await dialog.showOpenDialog(window!,{properties:['openFile'],filters:[{name:t('officeFiles'),extensions},{name:t('allFiles'),extensions:['*']}]});
    return result.canceled || !result.filePaths[0] ? null : openSelected(result.filePaths[0]);
  });
  handle('office:recent',async(_event,candidate)=>{id(candidate); return openSelected(settings.resolveRecent(candidate));});
  handle('office:drop',async(_event,candidate)=>{
    return openSelected(parseDroppedPath(candidate)); // preload only accepts real browser File objects.
  });
  handle('office:tabs',async(_event,candidate)=>{
    const next = parseTabStates(candidate);
    tabs.clear(); for(const value of next) tabs.set(value.id,value);
  });
  handle('office:save-draft',async(event,candidate)=>saveEditable('oex',candidate,trusted(event)));
  handle('office:save-docx',async(event,candidate)=>saveEditable('docx',candidate,trusted(event)));
  handle('office:save-doc',async(event,candidate)=>saveEditable('doc',candidate,trusted(event)));
  handle('office:save-xlsx',async(event,candidate)=>saveEditable('xlsx',candidate,trusted(event)));
  handle('office:save-xls',async(event,candidate)=>saveEditable('xls',candidate,trusted(event)));
  handle('office:save-pptx',async(event,candidate)=>saveEditable('pptx',candidate,trusted(event)));
  handle('office:save-ppt',async(event,candidate)=>saveEditable('ppt',candidate,trusted(event)));
  handle('office:save-copy',async(_event,candidate)=>{
    const request=parsePreservedCopy(candidate);
    if(!tabs.has(request.documentId))throw new OfficeError('unsafe','Unknown document session');
    if(runningSaves.has(request.documentId))throw new OfficeError('conflict','Save is already running');
    runningSaves.add(request.documentId);
    let target:string|undefined;
    try{
    const grant=authority.get(request.sourceToken,window!.id,request.documentId);
    const bytes=originals.get(request.sourceToken);
    if(!bytes)throw new OfficeError('invalid','Original bytes are unavailable');
    const suffix=extname(grant.path),extension=suffix.slice(1);
    const chosen=await dialog.showSaveDialog(window!,{defaultPath:`${basename(grant.path,suffix)}-copy${suffix}`,filters:[{name:extension?t('originalFormat'):t('allFiles'),extensions:[extension||'*']}],properties:['showOverwriteConfirmation','createDirectory']});
    if(chosen.canceled||!chosen.filePath)return null;
    if(extname(chosen.filePath).toLowerCase()!==extname(grant.path).toLowerCase())throw new OfficeError('invalid','A preserved copy must retain its source extension');
    let exists=false;try{await lstat(chosen.filePath);exists=true;}catch(error){if((error as NodeJS.ErrnoException).code!=='ENOENT')throw error;}
    if(exists){const confirmation=await dialog.showMessageBox(window!,{type:'warning',buttons:[t('cancel'),t('replace')],defaultId:0,cancelId:0,message:t('replaceTarget'),detail:basename(chosen.filePath)});if(confirmation.response!==1)return null;}
    target=await authority.authorizeTarget(chosen.filePath,window!.id,request.documentId,exists,request.sourceToken);
    const expected=byteHash(bytes);
    const result=await authority.commit(target,window!.id,request.documentId,bytes,{backup:async(previous,identity)=>{await backups.put(request.documentId as string,`backup:${identity.sha256}`,0,previous);},validate:async reread=>{if(byteHash(reread)!==expected)throw new OfficeError('invalid','Preserved copy verification failed');}});
    return {name:result.name,durability:result.durability,warnings:[] as string[]};
    }finally{if(target){try{authority.revokeToken(target,window!.id,request.documentId);}catch{/* already revoked on close */}}runningSaves.delete(request.documentId);}
  });
  handle('office:confirm-close',async(_event,candidate)=>{
    id(candidate);const tab=tabs.get(candidate);
    if(runningSaves.has(candidate))return 'cancel' as const;
    if(!tab||!tab.dirty)return 'close' as const;
    const choice=await dialog.showMessageBox(window!,{type:'question',buttons:[t('save'),t('discardChanges'),t('cancel')],defaultId:0,cancelId:2,message:t('saveBeforeClose'),detail:tab.title});
    return (['save','discard','cancel'] as const)[choice.response]??'cancel';
  });
  handle('office:release',async(_event,candidate,discard)=>{
    const release=parseRelease(candidate,discard);candidate=release.documentId;discard=release.discard;id(candidate);
    if(runningSaves.has(candidate))throw new OfficeError('conflict','Save is running');
    if(discard)await recovery.clearSaved(candidate,Number.MAX_SAFE_INTEGER);
    authority.revokeDocument(window!.id,candidate);tabs.delete(candidate);pptxImportBaselines.delete(candidate);
    // Released source tokens can no longer be looked up; remove any orphaned original payload.
    for(const token of originals.keys()){try{authority.get(token,window!.id,candidate);}catch{let live=false;for(const doc of tabs.keys()){try{authority.get(token,window!.id,doc);live=true;break;}catch{/* different session */}}if(!live)originals.delete(token);}}
  });
  handle('office:finish-close',async()=>{if(tabs.size||runningSaves.size)throw new OfficeError('conflict','Document sessions remain open');allowClose=true;window!.close();});
  handle('office:cancel-window-close',async()=>{quitRequested=false;});
  handle('office:settings',async(_event,candidate)=>{const patch=parseSettingsPatch(candidate);await settings.update(patch);nativeTheme.themeSource=settings.value().theme;installMenu();return settings.publicValue();});
  handle('office:recover',async(_event,candidate)=>{return recovery.load(parseRecoveryId(candidate));});
  handle('office:discard-recovery',async(_event,candidate)=>{await recovery.discard(parseRecoveryId(candidate));});
  handle('office:checkpoint',async(_event,candidate)=>{
    const r=parseCheckpoint(candidate);
    if(!tabs.has(r.documentId))throw new OfficeError('invalid','Unknown checkpoint session');
    await recovery.put(r.documentId,r.contentId,r.revision,r.bytes);
  });
}
app.on('open-file',(event,path)=>{event.preventDefault();try{enqueueFile(path);if(!window&&createMainWindow)void createMainWindow().catch(()=>dialog.showErrorBox('Office EX',t('windowFailed')));}catch{dialog.showErrorBox('Office EX',t('openFailed'));}});
app.on('second-instance',(_event,argv,cwd)=>{
  if(window&&!window.isDestroyed()){if(window.isMinimized())window.restore();window.focus();}
  for(const path of commandLineFiles(argv,Boolean(process.defaultApp),cwd)){try{enqueueFile(path);}catch{break;}}
  if(!window&&createMainWindow)void createMainWindow().catch(()=>dialog.showErrorBox('Office EX',t('windowFailed')));
});
// On macOS closing the last window is not quitting the app. Preserve an
// explicit Quit request across the asynchronous document close handshake.
app.on('before-quit',event=>{if(window&&!allowClose){event.preventDefault();quitRequested=true;send('close-window');}});
app.on('window-all-closed',()=>{if(process.platform!=='darwin'||quitRequested){quitRequested=false;app.quit();}});
if(ownsProfile)app.whenReady().then(async()=>{
  await mkdir(app.getPath('userData'),{recursive:true,mode:0o700});
  for(const path of commandLineFiles(process.argv,Boolean(process.defaultApp),process.cwd()))enqueueFile(path);
  settings=new SettingsStore(app.getPath('userData'));await settings.load();
  recovery=new RecoveryStore(join(app.getPath('userData'),'recovery'));
  backups=new RecoveryStore(join(app.getPath('userData'),'backups'));
  await recovery.cleanupStaging();await backups.cleanupStaging();
  nativeTheme.themeSource=settings.value().theme;
  const bundle=resolve(__dirname,'../ui');
  const resources=await loadLocalResources(bundle,['index.html','renderer.js','format-worker.js','sheets-recalc-worker.js','styles.css','pdf.worker.mjs','fonts.css']);
  protocol.handle('office-ex',request=>{
    const url=new URL(request.url);const resource=resources.get(url.pathname==='/'?'/index.html':url.pathname);
    if(!localOrigin(url)||!resource||request.method!=='GET'||url.search||url.hash)return new Response('',{status:404});
    return new Response(resource.data.slice().buffer,{headers:{'Content-Type':resource.type,'Content-Security-Policy':CSP,'X-Content-Type-Options':'nosniff','Cross-Origin-Resource-Policy':'same-origin'}});
  });
  session.defaultSession.setPermissionRequestHandler((_wc,_permission,callback)=>callback(false));
  session.defaultSession.setPermissionCheckHandler(()=>false);
  session.defaultSession.webRequest.onBeforeRequest((details,callback)=>{
    callback({cancel:!findLocalResource(resources,details.url)});
  });
  verifyDocx=await createDocxVerifier(bundle,join(__dirname,'docx-verifier-preload.cjs'));
  printPdf=await createPdfPrinter(bundle,join(__dirname,'pdf-printer-preload.cjs'));
  installIpc();installMenu();
  const createWindow=async():Promise<void>=>{
    if(window&&!window.isDestroyed())return;
    allowClose=false;quitRequested=false;rendererReady=false;
    window=new BrowserWindow({width:1280,height:840,minWidth:880,minHeight:600,show:false,title:'Office EX',backgroundColor:'#f7f8fa',webPreferences:{preload:join(__dirname,'preload.cjs'),sandbox:true,contextIsolation:true,nodeIntegration:false,nodeIntegrationInWorker:false,nodeIntegrationInSubFrames:false,webSecurity:true,webviewTag:false,spellcheck:false,devTools:!app.isPackaged}});
    window.webContents.setWindowOpenHandler(()=>({action:'deny'}));
    window.webContents.on('will-navigate',event=>event.preventDefault());
    window.webContents.on('will-attach-webview',event=>event.preventDefault());
    window.on('close',event=>{if(!allowClose){event.preventDefault();send('close-window');}});
    const currentId=window.id;
    window.on('closed',()=>{authority.revokeWindow(currentId);tabs.clear();originals.clear();pptxImportBaselines.clear();rendererReady=false;window=null;});
    window.once('ready-to-show',()=>window?.show());
    await window.loadURL(`${ORIGIN}/index.html`);
    // OS-open events queue until the renderer has installed its listener.
  };
  createMainWindow=createWindow;
  await createWindow();
  app.on('activate',()=>{if(!window)void createWindow();});
}).catch(error=>{console.error(error instanceof OfficeError?error.code:'startup-failed');app.exit(1);});
