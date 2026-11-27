// SPDX-License-Identifier: MIT
/** Main owns the source and save capability. A separate in-memory Chromium session,
 * hidden sandboxed window and worker build/verify the actual DOCX output.
 * The normal renderer never supplies a DOCX archive to the file commit path.
 */
import {BrowserWindow,ipcMain,session,type IpcMainEvent} from 'electron';
import {readFile} from 'node:fs/promises';
import {join} from 'node:path';
import {randomUUID} from 'node:crypto';
import {OfficeError,isRecord} from '@office-ex/contracts';

const URL='office-ex://app/verify-docx.html';
const ORIGIN='office-ex://app';
const MAX_BYTES=32*1024*1024;
const RESPONSE='office:docx-verify-result';
const CSP="default-src 'none'; script-src 'self'; worker-src 'self'; connect-src 'none'; object-src 'none'; base-uri 'none'; frame-src 'none'; form-action 'none'; frame-ancestors 'none'";
export interface VerifyDocxJob{format?:'doc'|'docx'|'xls'|'xlsx'|'ppt'|'pptx';documentId:string;bytes:Uint8Array;sourceBytes?:Uint8Array;compatibility?:boolean}
export type DocxVerifier=(job:VerifyDocxJob)=>Promise<Uint8Array>;

export async function createDocxVerifier(bundle:string,preload:string):Promise<DocxVerifier>{
  const isolated=session.fromPartition('office-ex-docx-verifier');
  const files=new Map<string,{data:Uint8Array;type:string}>();
  for(const [name,type] of [['verify-docx.html','text/html; charset=utf-8'],['verify-docx.js','text/javascript; charset=utf-8'],['format-worker.js','text/javascript; charset=utf-8']] as const)
    files.set('/'+name,{data:new Uint8Array(await readFile(join(bundle,name))),type});
  const local=(value:string):{data:Uint8Array;type:string}|undefined=>{
    try{
      const u=new globalThis.URL(value);
      if(u.protocol!=='office-ex:'||u.hostname!=='app'||u.port||u.username||u.password||u.search||u.hash)return undefined;
      return files.get(u.pathname);
    }catch{return undefined;}
  };
  isolated.protocol.handle('office-ex',request=>{
    const file=local(request.url);
    if(request.method!=='GET'||!file)return new Response('',{status:404});
    return new Response(file.data.slice().buffer,{headers:{'Content-Type':file.type,'Content-Security-Policy':CSP,'X-Content-Type-Options':'nosniff','Cross-Origin-Resource-Policy':'same-origin'}});
  });
  isolated.setPermissionRequestHandler((_wc,_permission,callback)=>callback(false));
  isolated.setPermissionCheckHandler(()=>false);
  isolated.webRequest.onBeforeRequest((details,callback)=>callback({cancel:!local(details.url)}));
  let active=false;
  return async(job:VerifyDocxJob):Promise<Uint8Array>=>{
    if(active)throw new OfficeError('conflict','Another DOCX verification is running');
    if(!job.bytes.byteLength||job.bytes.byteLength>MAX_BYTES||job.sourceBytes!==undefined&&(!job.sourceBytes.byteLength||job.sourceBytes.byteLength>MAX_BYTES))
      throw new OfficeError('limit','DOCX verification input exceeds budget');
    active=true;let child:BrowserWindow|undefined;
    try{
      child=new BrowserWindow({show:false,skipTaskbar:true,webPreferences:{session:isolated,preload,sandbox:true,contextIsolation:true,nodeIntegration:false,nodeIntegrationInWorker:false,nodeIntegrationInSubFrames:false,webSecurity:true,webviewTag:false,spellcheck:false,devTools:false}});
      const target=child,wc=target.webContents,id=randomUUID();
      wc.setWindowOpenHandler(()=>({action:'deny'}));
      wc.on('will-navigate',event=>event.preventDefault());
      wc.on('will-attach-webview',event=>event.preventDefault());
      return await new Promise<Uint8Array>((resolve,reject)=>{
        let settled=false;
        const finish=(error:OfficeError|undefined,bytes?:Uint8Array):void=>{
          if(settled)return;settled=true;clearTimeout(timer);ipcMain.removeListener(RESPONSE,onResult);
          if(error)reject(error);else resolve(bytes!);
        };
        const onResult=(event:IpcMainEvent,result:unknown,...extra:unknown[]):void=>{
          if(event.sender!==wc||event.senderFrame!==wc.mainFrame||event.senderFrame?.origin!==ORIGIN||event.senderFrame?.url!==URL)return;
          if(extra.length||!isRecord(result)||result.id!==id){finish(new OfficeError('invalid','Invalid verifier response'));return;}
          if(result.ok!==true){
            const code=isRecord(result.error)?result.error.code:undefined;
            const errorCode=code==='limit'||code==='invalid'||code==='unsafe'||code==='conflict'||code==='cancelled'?code:'unsupported';
            finish(new OfficeError(errorCode,'DOCX semantic verification rejected the snapshot or source'));return;
          }
          const bytes=result.bytes;
          if(!(bytes instanceof Uint8Array)||bytes.byteLength<4||bytes.byteLength>MAX_BYTES||((job.format==='doc'||job.format==='xls'||job.format==='ppt')?![0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1].every((n,i)=>bytes[i]===n):bytes[0]!==0x50||bytes[1]!==0x4b||bytes[2]!==3||bytes[3]!==4)){
            finish(new OfficeError('invalid','Verifier returned invalid DOCX output'));return;
          }
          finish(undefined,bytes.slice());
        };
        const timer=setTimeout(()=>finish(new OfficeError('limit','DOCX verification timed out')),20000);
        ipcMain.on(RESPONSE,onResult);
        wc.once('render-process-gone',()=>finish(new OfficeError('internal','DOCX verifier exited')));
        target.once('closed',()=>finish(new OfficeError('cancelled','DOCX verifier closed')));
        void target.loadURL(URL).then(()=>{
          if(!settled&&!target.isDestroyed())wc.send('office:docx-verify-job',{id,...job});
        }).catch(()=>finish(new OfficeError('internal','DOCX verifier could not load')));
      });
    }finally{if(child&&!child.isDestroyed())child.destroy();active=false;}
  };
}
