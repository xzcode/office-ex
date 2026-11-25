// SPDX-License-Identifier: MIT
/** Documents are laid out and PDF input is parsed in a separate sandbox. Main
 * owns the immutable job, native print operation and output bytes only. */
import { BrowserWindow, ipcMain, session, type IpcMainEvent } from 'electron';
import { randomUUID } from 'node:crypto';
import { OfficeError, isRecord } from '@office-ex/contracts';
import { loadLocalResources, findLocalResource } from './local-resources.js';

const URL = 'office-ex://app/pdf-print.html';
const ORIGIN = 'office-ex://app';
const RESPONSE = 'office:pdf-print-result';
const MAX_BYTES = 100 * 1024 * 1024;
const CSP = "default-src 'none'; script-src 'self'; style-src 'self'; img-src 'self' data: blob:; font-src 'self' blob:; worker-src 'self'; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-src 'none'; form-action 'none'; frame-ancestors 'none'";
export interface PdfPrintJob { kind: 'document' | 'pdf'; bytes: Uint8Array; documentId: string; revision: number; contentId?: string }
export interface PdfPrintResult { bytes?: Uint8Array; pages: number; warnings: string[] }
export type PdfPrinter = (job: PdfPrintJob, mode: 'export' | 'print', owner: BrowserWindow) => Promise<PdfPrintResult>;

export async function createPdfPrinter(bundle: string, preload: string): Promise<PdfPrinter> {
  const isolated = session.fromPartition('office-ex-pdf-printer');
  const resources = await loadLocalResources(bundle, ['pdf-print.html', 'pdf-print.js', 'format-worker.js', 'pdf.worker.mjs', 'pdf-print.css', 'fonts.css']);
  isolated.protocol.handle('office-ex', request => {
    const file = findLocalResource(resources, request.url);
    if (request.method !== 'GET' || !file) return new Response('', { status: 404 });
    return new Response(file.data.slice().buffer, { headers: { 'Content-Type': file.type, 'Content-Security-Policy': CSP, 'X-Content-Type-Options': 'nosniff', 'Cross-Origin-Resource-Policy': 'same-origin' } });
  });
  isolated.setPermissionRequestHandler((_wc, _permission, callback) => callback(false));
  isolated.setPermissionCheckHandler(() => false);
  isolated.webRequest.onBeforeRequest((details, callback) => callback({ cancel: !findLocalResource(resources, details.url) }));
  let active = false;
  return async (job, mode, owner) => {
    if (active) throw new OfficeError('conflict', 'Another print task is running');
    if (!job.bytes.byteLength || job.bytes.byteLength > (job.kind === 'document' ? 32 * 1024 * 1024 : MAX_BYTES)) throw new OfficeError('limit', 'Print input exceeds its byte budget');
    if (owner.isDestroyed()) throw new OfficeError('cancelled', 'Document window was closed');
    active = true; let child: BrowserWindow | undefined;
    try {
      child = new BrowserWindow({ show: false, parent: owner, skipTaskbar: true, webPreferences: { session: isolated, preload, sandbox: true, contextIsolation: true, nodeIntegration: false, nodeIntegrationInWorker: false, nodeIntegrationInSubFrames: false, webSecurity: true, webviewTag: false, spellcheck: false, devTools: false } });
      const target = child, wc = target.webContents, id = randomUUID();
      wc.setWindowOpenHandler(() => ({ action: 'deny' }));
      wc.on('will-navigate', event => event.preventDefault());
      wc.on('will-attach-webview', event => event.preventDefault());
      const metadata = await new Promise<{ pages: number; warnings: string[] }>((resolve, reject) => {
        let settled = false;
        const finish = (error?: OfficeError, result?: { pages: number; warnings: string[] }): void => {
          if (settled) return; settled = true; clearTimeout(timer); ipcMain.removeListener(RESPONSE, onResult);
          owner.removeListener('closed', onOwnerClosed); if (error) reject(error); else resolve(result!);
        };
        const onOwnerClosed = (): void => finish(new OfficeError('cancelled', 'Document window was closed'));
        const onResult = (event: IpcMainEvent, result: unknown, ...extra: unknown[]): void => {
          if (event.sender !== wc || event.senderFrame !== wc.mainFrame || event.senderFrame?.origin !== ORIGIN || event.senderFrame.url !== URL) return;
          if (extra.length || !isRecord(result) || result.id !== id) { finish(new OfficeError('invalid', 'Invalid print renderer response')); return; }
          if (result.ok !== true) {
            const code = isRecord(result.error) ? result.error.code : undefined;
            const message=isRecord(result.error)&&typeof result.error.message==='string'&&['pdf.font-not-approved','pdf.font-unavailable','pdf.font-glyph-unavailable','pdf.encrypted','pdf.sheets-print-rules-unavailable'].includes(result.error.message)?result.error.message:'Document cannot be safely printed or exported';
            finish(new OfficeError(code === 'limit' || code === 'unsafe' || code === 'invalid' || code === 'cancelled' ? code : 'unsupported', message)); return;
          }
          if (typeof result.pages !== 'number' || !Number.isSafeInteger(result.pages) || result.pages < 1 || result.pages > 500 || !Array.isArray(result.pageSizes) || result.pageSizes.length !== result.pages || result.pageSizes.some(size => !isRecord(size) || typeof size.width !== 'number' || typeof size.height !== 'number' || !Number.isFinite(size.width) || !Number.isFinite(size.height) || size.width <= 0 || size.height <= 0 || size.width > 20000 || size.height > 20000) || !Array.isArray(result.warnings) || result.warnings.length > 100 || result.warnings.some(warning => typeof warning !== 'string' || !/^[a-zA-Z0-9_.-]{1,128}$/.test(warning))) { finish(new OfficeError('invalid', 'Print metadata exceeds its budget')); return; }
          finish(undefined, { pages: result.pages, warnings: result.warnings as string[] });
        };
        const timer = setTimeout(() => finish(new OfficeError('limit', 'Print layout exceeded its time budget')), 45_000);
        ipcMain.on(RESPONSE, onResult); owner.once('closed', onOwnerClosed);
        wc.once('render-process-gone', () => finish(new OfficeError('internal', 'Print renderer exited')));
        target.once('closed', () => finish(new OfficeError('cancelled', 'Print renderer closed')));
        void target.loadURL(URL).then(() => { if (!settled && !target.isDestroyed()) wc.send('office:pdf-print-job', { id, ...job }); }).catch(() => finish(new OfficeError('internal', 'Print renderer failed to load')));
      });
      if (owner.isDestroyed() || target.isDestroyed()) throw new OfficeError('cancelled', 'Print window was closed');
      if (mode === 'print') {
        await new Promise<void>((resolve, reject) => wc.print({ silent: false, printBackground: true, margins: { marginType: 'none' } }, (success, reason) => success ? resolve() : reject(new OfficeError(/cancel/i.test(reason) ? 'cancelled' : 'io', 'Native print did not complete'))));
        return metadata;
      }
      const bytes = new Uint8Array(await wc.printToPDF({ preferCSSPageSize: true, printBackground: true, displayHeaderFooter: false, generateTaggedPDF: true, margins: { top: 0, bottom: 0, left: 0, right: 0 } }));
      if (!bytes.byteLength || bytes.byteLength > MAX_BYTES || new TextDecoder().decode(bytes.slice(0, 5)) !== '%PDF-') throw new OfficeError('invalid', 'Invalid or oversized PDF output');
      return { ...metadata, bytes };
    } finally { if (child && !child.isDestroyed()) child.destroy(); active = false; }
  };
}
