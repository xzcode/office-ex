// SPDX-License-Identifier: MIT
/** Trusted print projection. Office models use the same revision-bound scene /
 * native paragraph layout as the editors. No source markup is ever inserted. */
import { assert, EMU_PER_CSS_PIXEL, OfficeError } from '@office-ex/contracts';
import { resolveTextStyle, type DeepReadonly, type DocumentSnapshot, type Paragraph, type TextStyle, type WriterDocument } from '@office-ex/core';
import { locateParagraph } from '@office-ex/writer';
import { WriterPaginationEngine, buildSlideScene, type WriterLine } from '@office-ex/render';
import { PdfReader, PDF_EMBEDDABLE_FONT_FAMILIES, PDF_LIMITS, assertPdfFontFamily, assertPdfTextCoverage } from '@office-ex/pdf';
import { BrowserParagraphShaper } from './native-writer-shaper.js';
import { CanvasFontMetrics } from './font-metrics.js';
import { createSlideSceneElement, overflowingSlideText } from './slides-scene-dom.js';

const css = (value: number): string => value / EMU_PER_CSS_PIXEL + 'px';
export interface PdfOutputInfo { readonly pages: number; readonly pageSizes: readonly { readonly width: number; readonly height: number }[]; readonly warnings: readonly string[] }
const color = (style: DeepReadonly<TextStyle>): string => { const c = style.color; return c?.kind === 'rgba' ? `rgba(${c.red},${c.green},${c.blue},${c.alpha / 255})` : '#20242c'; };
function at(element: HTMLElement, x: number, y: number, width: number, height: number): void { Object.assign(element.style, { position: 'absolute', left: css(x), top: css(y), width: css(width), height: css(height), boxSizing: 'border-box' }); }

function fontStyle(style: DeepReadonly<TextStyle>): void {
  for (const name of ['fontFamily', 'fontAscii', 'fontHighAnsi', 'fontEastAsia', 'fontComplex'] as const) assertPdfFontFamily(style[name]);
}
function paragraphFonts(model: DeepReadonly<WriterDocument>, paragraph: DeepReadonly<Paragraph>): void {
  const named = paragraph.styleId ?? model.defaultParagraphStyleId;
  const base = resolveTextStyle(model.defaultTextStyle, model.styles, named ? [named] : [], paragraph.style ?? {}).style;
  fontStyle(base);
  for (const run of paragraph.runs) if (run.type === 'text') { fontStyle(resolveTextStyle(base, model.styles, run.styleId ? [run.styleId] : [], run.style ?? {}).style); assertPdfTextCoverage(run.text); }
}

/** CSSOM inserts numeric page sizes in the trusted, linked stylesheet. No
 * inline style element, caller CSS, document string or remote URL is accepted. */
function createPage(parent: HTMLElement, index: number, width: number, height: number): HTMLElement {
  assert(Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0 && width <= 20000 && height <= 20000, 'limit', 'PDF page size budget exceeded');
  const sheet = Array.from(document.styleSheets).find(item => item.href && new URL(item.href).pathname.endsWith('/pdf-print.css'));
  assert(sheet, 'internal', 'Missing trusted PDF print stylesheet');
  const name = `office-page-${index}`;
  sheet.insertRule(`@page ${name} { size: ${width}px ${height}px; margin: 0; }`, sheet.cssRules.length);
  const page = document.createElement('section'); page.className = 'office-print-page'; page.style.page = name; page.style.width = width + 'px'; page.style.height = height + 'px'; page.dataset.page = String(index + 1); page.dir = 'ltr'; parent.append(page); return page;
}

/** Exact local FontFace objects, not the existence of a system family name,
 * prove the embedding policy. Fonts are original OFL files with fsType=0. */
export async function ensurePdfFonts(): Promise<void> {
  const faces = await Promise.all(PDF_EMBEDDABLE_FONT_FAMILIES.map(family => document.fonts.load(`16px ${JSON.stringify(family)}`)));
  assert(faces.every(list => list.length > 0 && list.every(face => face.status === 'loaded')), 'unsupported', 'pdf.font-unavailable');
  await document.fonts.ready;
}

async function decodedImages(root: HTMLElement): Promise<void> {
  await Promise.all(Array.from(root.querySelectorAll('img')).map(async image => { try { await image.decode(); } catch { throw new OfficeError('invalid', 'PDF image decode failed'); } assert(image.naturalWidth > 0 && image.naturalHeight > 0 && image.naturalWidth * image.naturalHeight <= 40_000_000, 'limit', 'PDF image budget exceeded'); }));
}

export async function prepareDocumentPdfOutput(snapshot: DocumentSnapshot, container: HTMLElement): Promise<PdfOutputInfo> {
  const model = snapshot.document;
  assert(model.kind !== 'sheets', 'unsupported', 'pdf.sheets-print-rules-unavailable');
  const sizes: { width: number; height: number }[] = [], warnings = new Set<string>();
  container.replaceChildren(); container.className = 'office-print-document'; container.dataset.revision = String(snapshot.revision); container.dataset.contentId = snapshot.contentId;
  // Validate every used text face and every codepoint before anything is printed.
  if (model.kind === 'writer') {
    for (const section of model.sections) for (const block of section.blocks) {
      if (block.type === 'paragraph') paragraphFonts(model, block);
      else if (block.type === 'table') for (const row of block.rows) for (const cell of row.cells) for (const paragraph of cell.blocks) { assert(paragraph.type === 'paragraph', 'unsupported', 'Nested table printing is unsupported'); paragraphFonts(model, paragraph); }
    }
  } else for (const slide of model.slides) for (const shape of buildSlideScene(snapshot, slide.id).shapes) if (shape.type === 'text') for (const paragraph of shape.paragraphs) { fontStyle(paragraph.style); if(paragraph.marker){fontStyle(paragraph.marker.style);assertPdfTextCoverage(paragraph.marker.text);} for (const run of paragraph.runs) { fontStyle(run.style); assertPdfTextCoverage(run.text); } }
  await ensurePdfFonts();
  if (model.kind === 'writer') {
    const shaper = new BrowserParagraphShaper(), metrics = new CanvasFontMetrics(), engine = new WriterPaginationEngine();
    const layout = engine.layout(snapshot, metrics, { paragraphShaper: shaper, maxPages: 500, maxClusters: 100000 });
    for (const warning of layout.warnings) warnings.add(warning);
    for (const source of layout.pages) {
      const width = source.width / EMU_PER_CSS_PIXEL, height = source.height / EMU_PER_CSS_PIXEL;
      const page = createPage(container, sizes.length, width, height); sizes.push({ width, height });
      for (const cell of source.tableCells) { const box = document.createElement('div'); at(box, cell.x, cell.y, cell.width, cell.height); if (cell.border) box.style.border = css(cell.border.width) + ' solid ' + color({ color: cell.border.color }); page.append(box); }
      // Exact native line slices, including mixed-script shaping and bidi,
      // share the editor's paragraph projection rather than canvas screenshots.
      const groups = new Map<string, DeepReadonly<WriterLine>[]>();
      for (const line of source.lines) { assert(line.native, 'unsupported', 'PDF requires the final native paragraph layout'); const list = groups.get(line.paragraphId) ?? []; list.push(line); groups.set(line.paragraphId, list); }
      for (const [id, lines] of groups) {
        lines.sort((a, b) => a.y - b.y); const paragraph = locateParagraph(snapshot, id).paragraph;
        for (let i = 0; i < lines.length;) {
          const first = lines[i]!, native = first.native!; let j = i + 1;
          while (j < lines.length && lines[j]!.native!.originX === native.originX && Math.abs(lines[j]!.y - lines[j - 1]!.y - lines[j - 1]!.height) <= 1 && Math.abs(lines[j]!.native!.sourceTop - lines[j - 1]!.native!.sourceTop - lines[j - 1]!.height) <= 1) j++;
          const last = lines[j - 1]!, clip = document.createElement('div'); at(clip, native.originX, first.y, native.boxWidth, last.y + last.height - first.y); clip.style.overflow = 'hidden';
          const projection = shaper.render(model, paragraph, native.paragraphWidth); projection.style.position = 'absolute'; projection.style.top = css(-native.sourceTop); projection.style.left = '0'; clip.append(projection); page.append(clip); i = j;
        }
      }
    }
  } else {
    assert(model.slides.length <= PDF_LIMITS.pages, 'limit', 'PDF slide budget exceeded');
    for (const slide of model.slides) {
      const scene = buildSlideScene(snapshot, slide.id), width = scene.width / EMU_PER_CSS_PIXEL, height = scene.height / EMU_PER_CSS_PIXEL;
      assert(!scene.diagnostics.some(item => item.severity === 'error' || item.code === 'slides.image-unavailable'), 'unsupported', 'PDF slide resources are unavailable');
      for (const diagnostic of scene.diagnostics) warnings.add(diagnostic.code);
      const page = createPage(container, sizes.length, width, height); sizes.push({ width, height }); page.append(createSlideSceneElement(scene));
    }
  }
  await decodedImages(container); await document.fonts.ready;
  if (model.kind === 'slides') for (const page of container.children) if (overflowingSlideText(page as HTMLElement).length) warnings.add('slides.text-overflow');
  assert(sizes.length > 0 && sizes.length <= PDF_LIMITS.pages, 'limit', 'Invalid PDF output page count');
  await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
  return { pages: sizes.length, pageSizes: sizes, warnings: [...warnings].sort() };
}

/** Existing PDFs print their safe page appearances. No plugin, OpenAction,
 * JavaScript, document-controlled URL, attachment or annotation UI is loaded. */
export async function prepareExistingPdfPrint(bytes: Uint8Array, container: HTMLElement): Promise<PdfOutputInfo> {
  container.replaceChildren(); container.className = 'office-print-document';
  const reader = await PdfReader.open(bytes, { assetBaseUrl: new URL('./pdf-assets/', import.meta.url).href });
  const sizes: { width: number; height: number }[] = []; let pixels = 0;
  try {
    for (let number = 1; number <= reader.pages; number++) {
      const size = await reader.pageSize(number); pixels += Math.ceil(size.width * 1.5) * Math.ceil(size.height * 1.5);
      assert(pixels <= PDF_LIMITS.printPixels, 'limit', 'PDF print raster budget exceeded');
      const page = createPage(container, sizes.length, size.width, size.height); sizes.push(size); const canvas = document.createElement('canvas'); page.append(canvas);
      await reader.render(number, canvas, undefined, 1.5, undefined, true); canvas.style.width = size.width + 'px'; canvas.style.height = size.height + 'px';
    }
    await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
    return { pages: sizes.length, pageSizes: sizes, warnings: ['pdf.safe-appearance-print'] };
  } finally { await reader.destroy(); }
}
