// SPDX-License-Identifier: MIT
/** Only installed in the host-created, isolated verification window. */
import {contextBridge,ipcRenderer} from 'electron';
const bridge={
  onJob:(callback:(job:unknown)=>void):void=>{ipcRenderer.once('office:docx-verify-job',(_event,job)=>callback(job));},
  finish:(result:unknown):void=>{ipcRenderer.send('office:docx-verify-result',result);}
};
contextBridge.exposeInMainWorld('docxVerification',Object.freeze(bridge));
