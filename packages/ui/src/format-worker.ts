import { asFailure } from '@office-ex/contracts';
import { inspectFile, inspectForEditor, serializeDraft, parseDraft, prepareWriterDocxSave } from '@office-ex/format-service';
interface WorkerRequest {id:number;type:'inspect'|'serialize'|'recover'|'docx-save';bytes?:Uint8Array;filename?:string;documentId?:string;sourceBytes?:Uint8Array;document?:unknown}
self.onmessage = async (event:MessageEvent<WorkerRequest>):Promise<void> => {
  const request = event.data;
  try {
    if (!request || !Number.isSafeInteger(request.id)) throw new Error('Invalid worker request');
    const value = request.type === 'inspect' ? (request.documentId ? await inspectForEditor(request.bytes!,request.filename??'',request.documentId) : await inspectFile(request.bytes!,request.filename))
      : request.type === 'serialize' ? serializeDraft(request.document)
      : request.type === 'recover' ? parseDraft(request.bytes!)
      : request.type === 'docx-save' ? prepareWriterDocxSave(request.bytes!,request.documentId!,request.sourceBytes) : (() => {throw new Error('Invalid worker request');})();
    self.postMessage({id:request.id,ok:true,value});
  } catch (error) { self.postMessage({id:request.id,ok:false,error:asFailure(error)}); }
};
