// SPDX-License-Identifier: MIT
import { assert, asFailure, isRecord } from '@office-ex/contracts';
import { validateDocument, type OfficeDocument } from '@office-ex/core';
import { prepareDocumentPdfOutput, prepareExistingPdfPrint } from './pdf-output.js';

interface PrintJob { id: string; kind: 'document' | 'pdf'; bytes: Uint8Array; documentId?: string; revision?: number; contentId?: string }
declare global { interface Window { pdfPrinter: { onJob(callback: (value: unknown) => void): void; finish(value: unknown): void } } }
function recover(bytes: Uint8Array): Promise<OfficeDocument> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL('./format-worker.js', import.meta.url), { type: 'module' });
    const done = (): void => { clearTimeout(timer); worker.terminate(); };
    const timer = setTimeout(() => { done(); reject(new Error('Print document parsing timed out')); }, 20000);
    worker.onerror = () => { done(); reject(new Error('Print document parsing failed')); };
    worker.onmessage = event => { done(); if (event.data?.id === 1 && event.data.ok === true) { try { validateDocument(event.data.value); resolve(event.data.value); } catch (error) { reject(error); } } else reject(new Error('Invalid print document')); };
    worker.postMessage({ id: 1, type: 'recover', bytes });
  });
}
let started = false;
window.pdfPrinter.onJob(value => {
  if (started) return; started = true;
  void (async () => {
    assert(isRecord(value) && typeof value.id === 'string' && value.id.length <= 128 && (value.kind === 'document' || value.kind === 'pdf') && value.bytes instanceof Uint8Array, 'invalid', 'Invalid print job');
    const job = value as unknown as PrintJob, container = document.createElement('main'); document.body.replaceChildren(container);
    let result;
    if (job.kind === 'pdf') result = await prepareExistingPdfPrint(job.bytes, container);
    else {
      assert(job.bytes.byteLength <= 32 * 1024 * 1024 && typeof job.documentId === 'string' && Number.isSafeInteger(job.revision) && job.revision! >= 0 && typeof job.contentId === 'string', 'invalid', 'Invalid document print request');
      const model = await recover(job.bytes); assert(model.documentId === job.documentId, 'conflict', 'Print document identity mismatch');
      result = await prepareDocumentPdfOutput({ document: model, revision: job.revision!, contentId: job.contentId! }, container);
    }
    window.pdfPrinter.finish({ id: job.id, ok: true, ...result });
  })().catch(error => window.pdfPrinter.finish({ id: isRecord(value) && typeof value.id === 'string' ? value.id : '', ok: false, error: asFailure(error) }));
});
export {};
