// SPDX-License-Identifier: MIT
import { OfficeError } from '@office-ex/contracts';
import { PdfReader, PDF_LIMITS, findPdfText, validatePdfBytes, type PdfRenderedPage, type PdfMatch } from '@office-ex/pdf';
import type { MessageKey } from '@office-ex/i18n';

export interface PdfViewState {
  readonly bytes: Uint8Array;
  page: number;
  zoom: number;
  fitWidth: boolean;
  query: string;
  matches: { page: number; match: PdfMatch }[];
  matchIndex: number;
  reader?: PdfReader;
  loading?: Promise<PdfReader>;
  readonly lifetime: AbortController;
  closed: boolean;
  /** Remove the view without releasing the document when switching tabs. */
  detach?: () => void;
}
export interface PdfHost {
  t(key: MessageKey, args?: Readonly<Record<string, string | number>>): string;
  error(error: unknown): void;
  isCurrent(): boolean;
  print(): Promise<void> | void;
  number?(value: number): string;
}
export function createPdfViewState(bytes: Uint8Array): PdfViewState {
  validatePdfBytes(bytes);
  return { bytes: bytes.slice(), page: 1, zoom: 1, fitWidth: true, query: '', matches: [], matchIndex: -1, lifetime: new AbortController(), closed: false };
}
export function disposePdfViewState(state: PdfViewState): void {
  if (state.closed) return;
  state.closed = true; state.detach?.(); state.detach = undefined; state.lifetime.abort();
  void state.reader?.destroy(); state.reader = undefined; state.matches = [];
}
function element<K extends keyof HTMLElementTagNameMap>(name: K, text = '', className?: string): HTMLElementTagNameMap[K] {
  const result = document.createElement(name); result.textContent = text; if (className) result.className = className; return result;
}

export function pdfWorkspace(host: PdfHost, state: PdfViewState): HTMLElement {
  state.detach?.();
  const root = element('section', '', 'pdf-workspace'), toolbar = element('div', '', 'pdf-toolbar'), searchbar = element('div', '', 'pdf-searchbar'), status = element('p', host.t('pdfLoading'), 'pdf-status'), viewport = element('div', '', 'pdf-viewport');
  root.setAttribute('aria-label', host.t('pdfReader')); root.id = 'pdf-workspace';
  toolbar.setAttribute('role', 'toolbar'); toolbar.setAttribute('aria-label', host.t('pdfReader'));
  status.setAttribute('role', 'status'); viewport.tabIndex = 0;
  root.append(toolbar, searchbar, element('p', host.t('pdfSafeMode'), 'pdf-safety-note'), status, viewport);
  const number = (n: number) => host.number?.(n) ?? String(n);
  let detached = false, rendered: PdfRenderedPage | undefined, renderAbort: AbortController | undefined, searchAbort: AbortController | undefined, serial = 0;
  const active = () => !detached && !state.closed && root.isConnected && host.isCurrent();
  const failure = (error: unknown): void => {
    if (!active() || (error instanceof OfficeError && error.code === 'cancelled')) return;
    status.textContent = error instanceof Error && error.message === 'pdf.encrypted' ? host.t('pdfEncrypted') : host.t('errorInvalid');
    host.error(error);
  };
  const button = (key: MessageKey, id: string, fn: () => void): HTMLButtonElement => { const b = element('button', host.t(key)); b.id = id; b.type = 'button'; b.onclick = fn; return b; };
  const pageText = element('span', '', 'pdf-page-position'), pageLabel = element('label', host.t('pdfPageNumber')), pageInput = element('input'), zoomLabel = element('label', host.t('pdfZoom')), zoomInput = element('select');
  pageInput.type = 'number'; pageInput.min = '1'; pageInput.step = '1'; pageInput.value = String(state.page); pageInput.id = 'pdf-page-number'; pageInput.setAttribute('aria-label', host.t('pdfPageNumber')); pageLabel.append(pageInput);
  const previous = button('writerPreviousPage', 'pdf-previous-page', () => { if (state.page > 1) { state.page--; void paint(); } });
  const next = button('writerNextPage', 'pdf-next-page', () => { if (state.reader && state.page < state.reader.pages) { state.page++; void paint(); } });
  for (const value of [0.25, 0.5, 0.75, 1, 1.25, 1.5, 2, 3, 4]) { const option = element('option', number(value * 100) + '%'); option.value = String(value); zoomInput.append(option); }
  zoomInput.value = String(state.zoom); zoomInput.id = 'pdf-zoom'; zoomInput.setAttribute('aria-label', host.t('pdfZoom')); zoomLabel.append(zoomInput);
  const fit = button('pdfFitWidth', 'pdf-fit-width', () => { state.fitWidth = true; void paint(); });
  const print = button('pdfPrint', 'pdf-print', () => { void Promise.resolve().then(() => host.print()).catch(failure); });
  toolbar.append(previous, next, pageText, pageLabel, zoomLabel, fit, print);
  const search = element('input'); search.type = 'search'; search.id = 'pdf-search'; search.maxLength = PDF_LIMITS.queryCharacters; search.value = state.query; search.setAttribute('aria-label', host.t('pdfSearch')); search.placeholder = host.t('pdfSearch');
  const searchStatus = element('span', '', 'pdf-search-status'); searchStatus.setAttribute('role', 'status');
  const find = button('pdfSearch', 'pdf-find', () => { void findAll(); });
  const matchPrevious = button('pdfSearchPrevious', 'pdf-match-previous', () => { void showMatch(-1); });
  const matchNext = button('pdfSearchNext', 'pdf-match-next', () => { void showMatch(1); });
  searchbar.append(search, find, matchPrevious, matchNext, searchStatus);
  function controls(): void {
    const count = state.reader?.pages ?? 0;
    pageText.textContent = count ? host.t('pdfPage', { current: number(state.page), total: number(count) }) : '';
    pageInput.value = String(state.page); pageInput.max = String(count || 1);
    previous.disabled = !count || state.page <= 1; next.disabled = !count || state.page >= count; pageInput.disabled = !count; zoomInput.disabled = !count; fit.disabled = !count; print.disabled = !count;
    matchPrevious.disabled = matchNext.disabled = state.matches.length === 0;
    if (state.matches.length) searchStatus.textContent = host.t('pdfMatches', { current: number(state.matchIndex + 1), total: number(state.matches.length) });
  }
  function highlight(): void {
    const onPage = state.matches.filter(value => value.page === state.page);
    const selected = state.matches[state.matchIndex];
    rendered?.highlight(onPage.map(value => value.match), onPage.findIndex(value => value === selected));
    viewport.querySelector<HTMLElement>('.pdf-match-current')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }
  async function reader(): Promise<PdfReader> {
    if (state.closed) throw new OfficeError('cancelled', 'PDF tab closed');
    if (state.reader?.isClosed) { state.reader = undefined; state.loading = undefined; }
    if (state.reader) return state.reader;
    if (!state.loading) state.loading = PdfReader.open(state.bytes, { assetBaseUrl: new URL('./pdf-assets/', import.meta.url).href, signal: state.lifetime.signal });
    try { const result = await state.loading; if (state.closed) { await result.destroy(); throw new OfficeError('cancelled', 'PDF tab closed'); } state.reader = result; return result; }
    catch (error) { state.loading = undefined; throw error; }
  }
  async function paint(): Promise<void> {
    const mine = ++serial;
    root.dataset.loaded = 'false';
    renderAbort?.abort(); renderAbort = new AbortController(); const signal = renderAbort.signal;
    rendered?.destroy(); rendered = undefined;
    const paper = element('div', '', 'pdf-paper'), canvas = element('canvas'), text = element('div', '', 'pdf-text-layer');
    canvas.setAttribute('aria-hidden', 'true'); text.setAttribute('aria-label', host.t('pdfReader')); paper.dir = 'ltr'; paper.append(canvas, text); viewport.replaceChildren(paper);
    status.textContent = host.t('pdfRendering');
    try {
      const document = await reader(); if (!active() || mine !== serial) return;
      state.page = Math.max(1, Math.min(document.pages, state.page));
      if (state.fitWidth) { const size = await document.pageSize(state.page, signal); state.zoom = Math.max(0.1, Math.min(4, (viewport.clientWidth - 36) / size.width)); }
      zoomInput.value = String(state.zoom);
      const result = await document.render(state.page, canvas, text, state.zoom, signal);
      if (!active() || mine !== serial) { result.destroy(); return; }
      rendered = result; paper.style.width = result.width + 'px'; paper.style.height = result.height + 'px';
      root.dataset.page = String(state.page); root.dataset.pages = String(document.pages); root.dataset.textCharacters = String(result.text.length); root.dataset.loaded = 'true';
      status.textContent = result.text.trim().length ? '' : host.t('pdfNoText'); controls(); highlight();
    } catch (error) { if (mine === serial) failure(error); }
  }
  async function findAll(): Promise<void> {
    searchAbort?.abort(); searchAbort = new AbortController(); const signal = searchAbort.signal;
    state.query = search.value; state.matches = []; state.matchIndex = -1; controls(); rendered?.highlight([]);
    if (!state.query) { searchStatus.textContent = ''; return; }
    find.disabled = true; searchStatus.textContent = host.t('pdfLoading'); let total = 0, limited = false;
    try {
      const document = await reader();
      for (let page = 1; page <= document.pages; page++) {
        if (signal.aborted || !active()) return;
        const text = await document.text(page, signal); total += text.text.length;
        if (total > PDF_LIMITS.searchCharacters || state.matches.length >= PDF_LIMITS.searchMatches) { limited = true; break; }
        const matches = findPdfText(text.text, state.query, PDF_LIMITS.searchMatches - state.matches.length);
        state.matches.push(...matches.map(match => ({ page, match })));
      }
      if (signal.aborted || !active()) return;
      if (state.matches.length) { state.matchIndex = 0; state.page = state.matches[0]!.page; await paint(); }
      else searchStatus.textContent = host.t('pdfNoMatches');
      controls(); if (limited) searchStatus.textContent += ' ' + host.t('pdfSearchLimited');
    } catch (error) { failure(error); }
    finally { if (active() && !signal.aborted) find.disabled = false; }
  }
  async function showMatch(delta: number): Promise<void> {
    if (!state.matches.length) return;
    state.matchIndex = (state.matchIndex + delta + state.matches.length) % state.matches.length;
    const page = state.matches[state.matchIndex]!.page;
    if (page !== state.page) { state.page = page; await paint(); } else highlight(); controls();
  }
  search.onkeydown = event => { if (event.key === 'Enter') { event.preventDefault(); if (state.query === search.value && state.matches.length) void showMatch(event.shiftKey ? -1 : 1); else void findAll(); } };
  search.oninput = () => { if (state.query !== search.value) { searchAbort?.abort(); find.disabled = false; } };
  pageInput.onchange = () => { const value = pageInput.valueAsNumber; if (Number.isSafeInteger(value) && value >= 1 && value <= (state.reader?.pages ?? 1)) { state.page = value; void paint(); } else pageInput.value = String(state.page); };
  zoomInput.onchange = () => { state.zoom = Number(zoomInput.value); state.fitWidth = false; void paint(); };
  viewport.onkeydown = event => { if (event.key === 'PageDown' || event.key === 'PageUp') { event.preventDefault(); if (event.key === 'PageDown') next.click(); else previous.click(); } };
  let lastWidth = 0, resizeTimer: ReturnType<typeof setTimeout> | undefined;
  const resize = new ResizeObserver(() => {
    const width = viewport.clientWidth; if (Math.abs(width - lastWidth) <= 2) return; lastWidth = width;
    if (active() && state.fitWidth && state.reader) { if (resizeTimer !== undefined) clearTimeout(resizeTimer); resizeTimer = setTimeout(() => { resizeTimer = undefined; if (active()) void paint(); }, 100); }
  });
  resize.observe(viewport);
  state.detach = () => { if (detached) return; detached = true; serial++; renderAbort?.abort(); searchAbort?.abort(); rendered?.destroy(); rendered = undefined; if (resizeTimer !== undefined) clearTimeout(resizeTimer); resize.disconnect(); };
  controls();
  requestAnimationFrame(() => { if (active()) void paint(); });
  return root;
}
