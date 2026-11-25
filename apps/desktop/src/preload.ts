import { contextBridge, ipcRenderer, webUtils } from 'electron';
import type { DesktopBridge, MenuAction, OpenFileResult } from '@office-ex/ui';
const api: DesktopBridge = {
  exportPdf: request => ipcRenderer.invoke('office:export-pdf',request),
  print: request => ipcRenderer.invoke('office:print',request),
  initialize: () => ipcRenderer.invoke('office:initialize'),
  openFile: () => ipcRenderer.invoke('office:open'),
  openRecent: id => ipcRenderer.invoke('office:recent',id),
  openDroppedFile: async file => {
    // Electron resolves a native File capability, never a renderer-supplied path property.
    let path:string;
    try {path=webUtils.getPathForFile(file);} catch {return {ok:false,error:{code:'invalid',message:'A native dropped file is required'}};}
    if(!path)return {ok:false,error:{code:'invalid',message:'A native dropped file is required'}};
    return ipcRenderer.invoke('office:drop',path);
  },
  saveDraft: request => ipcRenderer.invoke('office:save-draft',request),
  saveDocx: request => ipcRenderer.invoke('office:save-docx',request),
  saveDoc: request => ipcRenderer.invoke('office:save-doc',request),
  saveXlsx: request => ipcRenderer.invoke('office:save-xlsx',request),
  saveXls: request => ipcRenderer.invoke('office:save-xls',request),
  savePptx: request => ipcRenderer.invoke('office:save-pptx',request),
  savePpt: request => ipcRenderer.invoke('office:save-ppt',request),
  saveCopy: request => ipcRenderer.invoke('office:save-copy',request),
  updateTabs: states => ipcRenderer.invoke('office:tabs',states),
  confirmClose: id => ipcRenderer.invoke('office:confirm-close',id),
  releaseDocument: (id,discard) => ipcRenderer.invoke('office:release',id,discard),
  finishClose: () => ipcRenderer.invoke('office:finish-close'),
  updateSettings: patch => ipcRenderer.invoke('office:settings',patch),
  recover: id => ipcRenderer.invoke('office:recover',id),
  discardRecovery: id => ipcRenderer.invoke('office:discard-recovery',id),
  checkpoint: request => ipcRenderer.invoke('office:checkpoint',request),
  onAction: callback => {
    const handler = (_event:Electron.IpcRendererEvent, action:MenuAction):void => callback(action);
    ipcRenderer.on('office:action',handler); return () => ipcRenderer.removeListener('office:action',handler);
  },
  onFile: callback => {
    const handler = (_event:Electron.IpcRendererEvent, file:OpenFileResult):void => callback(file);
    ipcRenderer.on('office:file',handler); return () => ipcRenderer.removeListener('office:file',handler);
  },
};
contextBridge.exposeInMainWorld('office',Object.freeze(api));
