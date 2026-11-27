// SPDX-License-Identifier: MIT
/** This capability exists only in the isolated, host-created print renderer. */
import { contextBridge, ipcRenderer } from 'electron';
contextBridge.exposeInMainWorld('pdfPrinter', Object.freeze({
  onJob: (callback: (job: unknown) => void): void => { ipcRenderer.once('office:pdf-print-job', (_event, job) => callback(job)); },
  finish: (result: unknown): void => { ipcRenderer.send('office:pdf-print-result', result); },
}));
