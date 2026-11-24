import { asFailure } from '@office-ex/contracts';
import { inspectFile, serializeDraft, parseDraft } from '@office-ex/format-service';
interface WorkerRequest {id:number;type:'inspect'|'serialize'|'recover';bytes?:Uint8Array;filename?:string;document?:unknown}
self.onmessage = async (event:MessageEvent<WorkerRequest>):Promise<void> => {
  const request = event.data;
  try {
    if (!request || !Number.isSafeInteger(request.id)) throw new Error('Invalid worker request');
    const value = request.type === 'inspect' ? await inspectFile(request.bytes!,request.filename)
      : request.type === 'serialize' ? serializeDraft(request.document)
      : request.type === 'recover' ? parseDraft(request.bytes!) : (() => {throw new Error('Invalid worker request');})();
    self.postMessage({id:request.id,ok:true,value});
  } catch (error) { self.postMessage({id:request.id,ok:false,error:asFailure(error)}); }
};
