// SPDX-License-Identifier: MIT
/** PDF.js is a low-level parser/renderer. We deliberately do not instantiate
 * its viewer, scripting manager, attachment layer, link service or forms. */
import { assert, OfficeError } from '@office-ex/contracts';
import type { PDFDocumentLoadingTask, PDFDocumentProxy, PDFPageProxy, PDFWorker, TextLayer } from 'pdfjs-dist';
import type { TextContent } from 'pdfjs-dist/types/src/display/api.js';
import { PDF_LIMITS, validatePdfBytes, validatePdfPage, validatePdfViewport, type PdfMatch } from './policy.js';
import { PdfActualTextBuilder } from './actual-text.js';

type PdfLibrary = typeof import('pdfjs-dist');
export interface PdfOpenOptions { readonly assetBaseUrl: string; readonly signal?: AbortSignal }
export interface PdfPageText { readonly text: string; readonly segments: readonly { readonly start: number; readonly end: number }[]; readonly content: TextContent }
export interface PdfRenderedPage { readonly page: number; readonly width: number; readonly height: number; readonly text: string; highlight(matches: readonly PdfMatch[], active?: number): void; destroy(): void }

function cancelled(): OfficeError { return new OfficeError('cancelled', 'PDF operation cancelled'); }
function checkSignal(signal?: AbortSignal): void { if (signal?.aborted) throw cancelled(); }
function pdfFailure(error: unknown): OfficeError {
  if (error instanceof OfficeError) return error;
  if (error && typeof error === 'object' && 'name' in error && error.name === 'PasswordException') return new OfficeError('unsupported', 'pdf.encrypted');
  if (error && typeof error === 'object' && 'name' in error && ['RenderingCancelledException', 'AbortException'].includes(String(error.name))) return cancelled();
  return new OfficeError('invalid', 'PDF parsing or rendering failed');
}
async function deadline<T>(promise: Promise<T>, signal: AbortSignal | undefined, onCancel: (reason: 'abort' | 'timeout') => void): Promise<T> {
  checkSignal(signal);
  let timer: ReturnType<typeof setTimeout> | undefined;
  let abort: (() => void) | undefined;
  const guard = new Promise<never>((_, reject) => {
    timer = setTimeout(() => { onCancel('timeout'); reject(new OfficeError('limit', 'PDF operation time budget exceeded')); }, PDF_LIMITS.operationMilliseconds);
    abort = () => { onCancel('abort'); reject(cancelled()); };
    signal?.addEventListener('abort', abort, { once: true });
  });
  // Keep the work watchdog after a view detaches. A getPage message cannot be
  // individually cancelled, so it must still terminate the worker if it stalls.
  const settled = promise.finally(() => { if (timer !== undefined) clearTimeout(timer); });
  try { return await Promise.race([settled, guard]); }
  finally { if (abort) signal?.removeEventListener('abort', abort); }
}

function resourceBase(input: string): URL {
  const url = new URL(input, location.href), here = new URL(location.href);
  assert(url.protocol === here.protocol && url.host === here.host && !url.username && !url.password && !url.search && !url.hash && url.pathname.endsWith('/pdf-assets/'), 'unsafe', 'PDF resources must use the fixed local asset directory');
  assert(here.protocol === 'office-ex:' || (['http:', 'https:'].includes(here.protocol) && ['localhost', '127.0.0.1', '[::1]'].includes(here.hostname)), 'unsafe', 'PDF reader requires a trusted local page');
  return url;
}

async function readText(page: PDFPageProxy, signal: AbortSignal | undefined): Promise<PdfPageText> {
  // PDF encodes glyphs, not an Office source string. PDF.js's text extraction
  // normalization resolves compatibility glyph forms (for example Arabic
  // presentation forms and CJK radicals) for selection/search only. The
  // original PDF bytes and every Office model string remain untouched.
  checkSignal(signal);
  const stream = page.streamTextContent({ includeMarkedContent: true, disableNormalization: false });
  const reader = stream.getReader();
  // PDF.js MessageHandler requires an Error reason before marking its stream
  // controller closed. An undefined reason rejects upstream cancellation after
  // the browser has already closed the local stream, leaving in-flight chunks.
  const abort = () => { void reader.cancel(new Error('PDF text read cancelled')).catch(() => {}); };
  signal?.addEventListener('abort', abort, { once: true });
  const content: TextContent = { items: [], styles: Object.create(null), lang: null };
  const builder = new PdfActualTextBuilder();
  let text = '';
  const segments: { start: number; end: number }[] = [];
  try {
    while (true) {
      checkSignal(signal);
      const result = await reader.read();
      if (result.done) break;
      const chunk = result.value as TextContent;
      assert(Array.isArray(chunk.items), 'invalid', 'Invalid PDF text chunk');
      for (const item of chunk.items) builder.push(item);
      Object.assign(content.styles, chunk.styles);
      if (chunk.lang) content.lang = chunk.lang;
    }
    checkSignal(signal);
    content.items = builder.finish();
    for (const item of content.items) {
      if (!('str' in item)) continue;
      assert(text.length + item.str.length + 1 <= PDF_LIMITS.pageTextCharacters, 'limit', 'PDF page text budget exceeded');
      segments.push({ start: text.length, end: text.length + item.str.length }); text += item.str + (item.hasEOL ? '\n' : '');
    }
    return { text, segments, content };
  } catch (error) { await reader.cancel(error instanceof Error ? error : new Error('PDF text extraction failed')).catch(() => {}); throw error; }
  finally { signal?.removeEventListener('abort', abort); reader.releaseLock(); }
}

export class PdfReader {
  readonly pages: number;
  readonly fingerprint: string;
  private closed = false;
  private readonly texts = new Map<number, PdfPageText>();
  private textCharacters = 0;
  private constructor(private readonly library: PdfLibrary, private readonly document: PDFDocumentProxy, private readonly loading: PDFDocumentLoadingTask, private readonly worker: PDFWorker, private readonly nativeWorker: Worker) {
    this.pages = document.numPages; this.fingerprint = document.fingerprints[0] ?? '';
  }
  static async open(bytes: Uint8Array, options: PdfOpenOptions): Promise<PdfReader> {
    validatePdfBytes(bytes); checkSignal(options.signal);
    const base = resourceBase(options.assetBaseUrl);
    const library = await import('pdfjs-dist');
    // v6 removed the old eval-generated font path. No scripting bundle is loaded.
    // An explicit module Worker avoids PDF.js's cross-origin blob wrapper and
    // fake-worker fallback, including for Electron's custom standard protocol.
    const nativeWorker = new Worker(new URL('../pdf.worker.mjs', base), { type: 'module' });
    const worker = library.PDFWorker.create({ port: nativeWorker });
    const loading = library.getDocument({
      data: bytes.slice(), worker, cMapUrl: new URL('cmaps/', base).href, cMapPacked: true,
      standardFontDataUrl: new URL('standard_fonts/', base).href,
      iccUrl: new URL('iccs/', base).href, wasmUrl: new URL('wasm/', base).href,
      useWorkerFetch: true, useWasm: false, disableWorkerRendering: true,
      disableRange: true, disableStream: true, disableAutoFetch: true,
      enableXfa: false, stopAtErrors: true, useSystemFonts: false,
      maxImageSize: PDF_LIMITS.pagePixels, canvasMaxAreaInBytes: PDF_LIMITS.pagePixels * 4,
      verbosity: 0,
    });
    let passwordReject: ((reason: unknown) => void) | undefined;
    const passwordFailure = new Promise<never>((_, reject) => { passwordReject = reject; });
    const stop = () => { void loading.destroy().catch(() => {}); worker.destroy(); nativeWorker.terminate(); };
    loading.onPassword = () => { passwordReject?.(new OfficeError('unsupported', 'pdf.encrypted')); stop(); };
    try {
      const document = await deadline(Promise.race([loading.promise, passwordFailure]), options.signal, stop);
      validatePdfPage(1, document.numPages);
      const [permissions, metadata] = await deadline(Promise.all([document.getPermissions(), document.getMetadata()]), options.signal, stop);
      // Permissions are emitted only from an encryption dictionary. Empty-password
      // encrypted files are rejected too, rather than accidentally bypassing it.
      const info = metadata.info as Record<string, unknown>;
      assert(permissions === null && !info.EncryptFilterName, 'unsupported', 'pdf.encrypted');
      assert(!document.isPureXfa, 'unsupported', 'PDF XFA documents are not supported');
      checkSignal(options.signal);
      return new PdfReader(library, document, loading, worker, nativeWorker);
    } catch (error) { stop(); throw pdfFailure(error); }
  }
  private live(): void { assert(!this.closed, 'cancelled', 'PDF reader closed'); }
  get isClosed(): boolean { return this.closed; }
  async text(pageNumber: number, signal?: AbortSignal): Promise<PdfPageText> {
    this.live(); validatePdfPage(pageNumber, this.pages); checkSignal(signal);
    const saved = this.texts.get(pageNumber); if (saved) return saved;
    try {
      const page = await deadline(this.document.getPage(pageNumber), signal, reason => { if (reason === 'timeout') void this.destroy(); });
      const controller = new AbortController(), abort = () => controller.abort(); signal?.addEventListener('abort', abort, { once: true });
      let result: PdfPageText;
      try { result = await deadline(readText(page, controller.signal), signal, reason => { controller.abort(); if (reason === 'timeout') void this.destroy(); }); }
      finally { signal?.removeEventListener('abort', abort); }
      this.live(); checkSignal(signal);
      // Bound retained text; reading one page remains possible after a full search.
      if (result.text.length <= PDF_LIMITS.searchCharacters) {
        while (this.textCharacters + result.text.length > PDF_LIMITS.searchCharacters && this.texts.size) {
          const first = this.texts.keys().next().value!; this.textCharacters -= this.texts.get(first)!.text.length; this.texts.delete(first);
        }
        this.textCharacters -= this.texts.get(pageNumber)?.text.length ?? 0; this.texts.set(pageNumber, result); this.textCharacters += result.text.length;
      }
      return result;
    } catch (error) { throw pdfFailure(error); }
  }
  async pageSize(pageNumber: number, signal?: AbortSignal): Promise<{ width: number; height: number }> {
    this.live(); validatePdfPage(pageNumber, this.pages); checkSignal(signal);
    const page = await deadline(this.document.getPage(pageNumber), signal, reason => { if (reason === 'timeout') void this.destroy(); });
    const viewport = page.getViewport({ scale: 96 / 72 });
    validatePdfViewport(viewport.width, viewport.height, 1);
    return { width: viewport.width, height: viewport.height };
  }
  async render(pageNumber: number, canvas: HTMLCanvasElement, textContainer: HTMLElement | undefined, zoom = 1, signal?: AbortSignal, print = false): Promise<PdfRenderedPage> {
    this.live(); validatePdfPage(pageNumber, this.pages); checkSignal(signal);
    const page = await deadline(this.document.getPage(pageNumber), signal, reason => { if (reason === 'timeout') void this.destroy(); });
    const viewport = page.getViewport({ scale: zoom * 96 / 72 });
    validatePdfViewport(viewport.width, viewport.height, zoom);
    const device = print ? 1 : Math.min(2, Math.max(1, devicePixelRatio));
    validatePdfViewport(viewport.width * device, viewport.height * device, zoom);
    canvas.width = Math.ceil(viewport.width * device); canvas.height = Math.ceil(viewport.height * device);
    canvas.style.width = viewport.width + 'px'; canvas.style.height = viewport.height + 'px';
    const context = canvas.getContext('2d', { alpha: false }); assert(context, 'unsupported', 'PDF canvas is unavailable');
    const task = page.render({ canvas, viewport, transform: device === 1 ? undefined : [device, 0, 0, device, 0, 0], intent: print ? 'print' : 'display', annotationMode: this.library.AnnotationMode.ENABLE, background: '#ffffff' });
    let layer: TextLayer | undefined, text: PdfPageText | undefined, destroyed = false;
    try {
      const result = await deadline(Promise.all([task.promise, textContainer ? this.text(pageNumber, signal) : Promise.resolve(undefined)]), signal, reason => { task.cancel(); if (reason === 'timeout') void this.destroy(); });
      this.live(); checkSignal(signal); text = result[1];
      if (textContainer && text) {
        textContainer.replaceChildren(); textContainer.classList.add('textLayer');
        textContainer.style.setProperty('--scale-factor', String(viewport.scale));
        textContainer.style.setProperty('--total-scale-factor', String(viewport.scale));
        layer = new this.library.TextLayer({ textContentSource: text.content, container: textContainer, viewport });
        await deadline(layer.render(), signal, reason => { layer?.cancel(); if (reason === 'timeout') void this.destroy(); });
      }
      return {
        page: pageNumber, width: viewport.width, height: viewport.height, text: text?.text ?? '',
        highlight: (matches, active = -1) => {
          if (destroyed || !layer || !text) return;
          for (let i = 0; i < layer.textDivs.length; i++) {
            const div = layer.textDivs[i]!, source = layer.textContentItemsStr[i] ?? '', segment = text.segments[i];
            div.replaceChildren();
            if (!segment) { div.textContent = source; continue; }
            let cursor = 0;
            for (let j = 0; j < matches.length; j++) {
              const match = matches[j]!;
              const a = Math.max(segment.start, match.start), b = Math.min(segment.end, match.end);
              if (a >= b) continue;
              const start = a - segment.start, end = b - segment.start;
              if (start > cursor) div.append(document.createTextNode(source.slice(cursor, start)));
              const mark = document.createElement('mark'); mark.textContent = source.slice(start, end); mark.className = j === active ? 'pdf-match-current' : 'pdf-match'; div.append(mark); cursor = end;
            }
            if (cursor < source.length) div.append(document.createTextNode(source.slice(cursor)));
          }
        },
        destroy: () => { if (destroyed) return; destroyed = true; task.cancel(); layer?.cancel(); if (textContainer) textContainer.replaceChildren(); canvas.width = 0; canvas.height = 0; page.cleanup(); },
      };
    } catch (error) { task.cancel(); layer?.cancel(); canvas.width = 0; canvas.height = 0; throw pdfFailure(error); }
  }
  async destroy(): Promise<void> { if (this.closed) return; this.closed = true; this.texts.clear(); this.textCharacters = 0; void this.loading.destroy().catch(() => {}); this.worker.destroy(); this.nativeWorker.terminate(); }
}
