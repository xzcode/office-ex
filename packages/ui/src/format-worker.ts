import { makeImageResource, decodeImageBytes, validateImageResource, type OfficeDocument } from '@office-ex/core';
import { OfficeError, asFailure } from '@office-ex/contracts';
import { inspectFile, inspectForEditor, serializeDraft, parseDraft, prepareWriterDocxSave, prepareWriterDocSave, prepareSheetsSave, prepareSlidesSave } from '@office-ex/format-service';
interface WorkerRequest {id:number;type:'inspect'|'serialize'|'recover'|'docx-save'|'doc-save'|'convert-doc'|'convert-docx'|'xls-save'|'xlsx-save'|'convert-xls'|'convert-xlsx'|'ppt-save'|'pptx-save'|'convert-ppt'|'convert-pptx'|'image';bytes?:Uint8Array;filename?:string;documentId?:string;sourceBytes?:Uint8Array;document?:unknown}
async function verifyRaster(bytes:Uint8Array,mimeType:string,width:number,height:number):Promise<void>{
  try{
    const bitmap=await createImageBitmap(new Blob([bytes.slice().buffer],{type:mimeType}));
    try{if(bitmap.width!==width||bitmap.height!==height)throw new OfficeError('invalid','Raster dimensions disagree');}finally{bitmap.close();}
  }catch{throw new OfficeError('invalid','Raster decode failed');}
}
async function verifyResources(document:OfficeDocument|undefined):Promise<void>{
  for(const resource of Object.values(document?.resources??{})){validateImageResource(resource);await verifyRaster(decodeImageBytes(resource.dataBase64),resource.mimeType,resource.pixelWidth,resource.pixelHeight);}
}
self.onmessage = async (event:MessageEvent<WorkerRequest>):Promise<void> => {
  const request = event.data;
  try {
    if (!request || !Number.isSafeInteger(request.id)) throw new Error('Invalid worker request');
    let value = request.type === 'inspect' ? (request.documentId ? await inspectForEditor(request.bytes!,request.filename??'',request.documentId) : await inspectFile(request.bytes!,request.filename))
      : request.type === 'serialize' ? serializeDraft(request.document)
      : request.type === 'recover' ? parseDraft(request.bytes!)
       : request.type === 'docx-save' ? prepareWriterDocxSave(request.bytes!,request.documentId!,request.sourceBytes)
      : request.type === 'doc-save' ? prepareWriterDocSave(request.bytes!,request.documentId!,request.sourceBytes)
      : request.type === 'convert-doc' ? (prepareWriterDocSave(serializeDraft(request.document),request.documentId!),true)
      : request.type === 'xls-save'||request.type === 'xlsx-save' ? prepareSheetsSave(request.type==='xls-save'?'xls':'xlsx',request.bytes!,request.documentId!,request.sourceBytes)
      : request.type === 'convert-xls'||request.type === 'convert-xlsx' ? (prepareSheetsSave(request.type==='convert-xls'?'xls':'xlsx',serializeDraft(request.document),request.documentId!),true)
      : request.type === 'ppt-save'||request.type === 'pptx-save' ? prepareSlidesSave(request.type==='ppt-save'?'ppt':'pptx',request.bytes!,request.documentId!,request.sourceBytes)
      : request.type === 'convert-ppt'||request.type === 'convert-pptx' ? (prepareSlidesSave(request.type==='convert-ppt'?'ppt':'pptx',serializeDraft(request.document),request.documentId!),true)
      : request.type === 'convert-docx' ? (prepareWriterDocxSave(serializeDraft(request.document),request.documentId!),true)
      : request.type === 'image' ? makeImageResource(request.bytes!) : (() => {throw new Error('Invalid worker request');})();
    if(request.type==='image'){const r=makeImageResource(request.bytes!);await verifyRaster(request.bytes!,r.mimeType,r.pixelWidth,r.pixelHeight);}
    else if(request.type==='inspect'&&request.documentId){
      const inspection=value as Awaited<ReturnType<typeof inspectForEditor>>;
      try{await verifyResources(inspection.editableWriter??inspection.editableSlides??inspection.draft);}catch{
        if(inspection.editableWriter||inspection.editableSlides)value=await inspectFile(request.bytes!,request.filename);
        else throw new OfficeError('invalid','Draft image decode failed');
      }
    }else if(request.type==='serialize'||request.type==='docx-save'||request.type==='doc-save'||request.type==='ppt-save'||request.type==='pptx-save'){
      await verifyResources(parseDraft(request.type==='serialize'?value as Uint8Array:request.bytes!));
    }else if(request.type==='recover')await verifyResources(value as OfficeDocument);
    self.postMessage({id:request.id,ok:true,value});
  } catch (error) { self.postMessage({id:request.id,ok:false,error:asFailure(error)}); }
};
