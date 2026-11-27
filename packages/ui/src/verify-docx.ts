// SPDX-License-Identifier: MIT
/** Trusted fixed renderer for one bounded verification job. No normal desktop API. */
interface Job {format?:'doc'|'docx'|'xls'|'xlsx'|'ppt'|'pptx';id:string;documentId:string;bytes:Uint8Array;sourceBytes?:Uint8Array}
declare global {interface Window {docxVerification:{onJob(callback:(job:unknown)=>void):void;finish(result:unknown):void}}}
window.docxVerification.onJob(value=>{
  const job=value as Job;
  const worker=new Worker(new URL('./format-worker.js',import.meta.url),{type:'module'});
  let finished=false;
  const done=(result:unknown):void=>{if(finished)return;finished=true;worker.terminate();window.docxVerification.finish(result);};
  worker.onmessage=event=>{
    const response=event.data;
    if(response?.id!==1){done({id:job.id,ok:false});return;}
    if(response.ok===true)done({id:job.id,ok:true,bytes:response.value});
    else done({id:job.id,ok:false,error:response.error});
  };
  worker.onerror=()=>done({id:job.id,ok:false});
  worker.onmessageerror=()=>done({id:job.id,ok:false});
  try{worker.postMessage({id:1,type:job.format==='ppt'?'ppt-save':job.format==='pptx'?'pptx-save':job.format==='xls'?'xls-save':job.format==='xlsx'?'xlsx-save':job.format==='doc'?'doc-save':'docx-save',documentId:job.documentId,bytes:job.bytes,sourceBytes:job.sourceBytes});}catch{done({id:job.id,ok:false});}
});
export {};
