// SPDX-License-Identifier: MIT
import { createDocument, DocumentSession, makeImageResource, type WriterDocument, type SlidesDocument } from '@office-ex/core';
import { PdfReader } from '@office-ex/pdf';
import { prepareDocumentPdfOutput, prepareExistingPdfPrint } from '../packages/ui/src/pdf-output.js';
import { createPdfViewState, pdfWorkspace, disposePdfViewState, type PdfViewState } from '../packages/ui/src/pdf-workspace.js';
import { message } from '@office-ex/i18n';

function image() { const c = document.createElement('canvas'); c.width = 60; c.height = 40; const x = c.getContext('2d')!; x.fillStyle = '#22a389'; x.fillRect(0, 0, 60, 40); x.fillStyle = '#efb139'; x.fillRect(10, 10, 30, 20); return makeImageResource(Uint8Array.from(atob(c.toDataURL('image/png').split(',')[1]!), s => s.charCodeAt(0))); }
function writer(): WriterDocument {
  const d = createDocument('writer', 'pdf-writer-fixture', 'PDF Writer fixture'); d.defaultTextStyle.fontFamily = 'Noto Sans';
  const asset = image(); d.resources[asset.id] = asset;
  d.sections[0]!.blocks = [
    { id: 'p-1', type: 'paragraph', style:{indent:228600,rightIndent:127000,firstLineIndent:-127000,paragraphBottomBorder:{color:{kind:'rgba',red:79,green:129,blue:189,alpha:255},width:12700,space:50800}}, runs: [{ id: 'r-1', type: 'text', text: 'Writer PDF — editable text and geometry', style: { fontSize: 304800, bold: true,characterSpacing:3175,kerningThreshold:177800 } }] },
    { id: 'p-2', type: 'paragraph', runs: [{ id: 'r-2', type: 'text', text: 'English 中文 繁體 日本語 한국어 العربية עברית हिन्दी ไทย Tiếng Việt Русский', style: { language: 'zh-CN' } }] },
    { id: 'p-image', type: 'paragraph', runs: [{ id: 'image-1', type: 'image', resourceId: asset.id, width: 1828800, height: 1219200 }] },
    { id: 'break-1', type: 'page-break' },
    { id: 'long', type: 'paragraph', runs: [{ id: 'long-r', type: 'text', text: Array.from({ length: 80 }, (_, i) => `UniqueLine${String(i + 1).padStart(3, '0')} — paragraph shared across pages`).join('\n') }] },
  ]; return d;
}
function slides(): SlidesDocument {
  const d = createDocument('slides', 'pdf-slides-fixture', 'PDF Slides fixture'), asset = image(); d.resources[asset.id] = asset;
  d.defaultTextStyle = { fontFamily: 'Noto Sans', fontSize: 304800 };
  d.slides = [0, 1, 2].map(i => ({ id: 's-' + i, background: { kind: 'rgba', red: [245, 235, 253][i]!, green: [248, 245, 239][i]!, blue: [255, 245, 235][i]!, alpha: 255 }, shapes: [
    { id: 't-' + i, type: 'text', geometry: { x: 914400, y: 700000, width: 10363200, height: 1700000, rotation: 0 }, runs: [{ id: 'tr-' + i, type: 'text', text: `Slide ${i + 1} — 文本 العربية हिन्दी`, style: { bold: true, language: 'zh-CN' } }] },
    { id: 'r-' + i, type: 'rect', geometry: { x: 1000000, y: 3200000, width: 2400000, height: 1600000, rotation: i * 15 * 60000 }, fill: { kind: 'rgba', red: 44, green: 110, blue: 196, alpha: 255 } },
    { id: 'e-' + i, type: 'ellipse', geometry: { x: 3900000, y: 3200000, width: 1800000, height: 1500000, rotation: 0 }, fill: { kind: 'rgba', red: 231, green: 139, blue: 49, alpha: 255 } },
    { id: 'i-' + i, type: 'image', resourceId: asset.id, geometry: { x: 6500000, y: 3200000, width: 2400000, height: 1600000, rotation: 0 } },
  ] }));
  d.slides[0]!.shapes.push({id:'list-output',type:'text',geometry:{x:9144000,y:3000000,width:2500000,height:2200000,rotation:0},textFrame:{autofit:{mode:'normal',fontScale:75000,lineReduction:0}},runs:[
    {id:'list-a',type:'text',text:'List alpha',style:{fontSize:228600,slideParagraph:{marginLeft:304800,firstLineIndent:-228600,bullet:{kind:'number',scheme:'arabicPeriod',startAt:3}}}},
    {id:'list-b',type:'text',text:'\nList beta',style:{fontSize:228600,slideParagraph:{marginLeft:304800,firstLineIndent:-228600,bullet:{kind:'number',scheme:'arabicPeriod'}}}}
  ]});
  const slide=d.slides[1]!,members=slide.shapes;
  const picture=members.find(s=>s.type==='image');if(picture?.type==='image')picture.crop={left:10000,right:5000,top:5000,bottom:0};
  members.push({id:'group-text',type:'text',geometry:{x:914400,y:5486400,width:4000000,height:914400,rotation:0},runs:[{id:'group-run',type:'text',text:'Grouped output',style:{fontSize:228600}}]});
  slide.shapes=[{id:'pdf-group',type:'group',geometry:{x:0,y:0,width:d.size.width,height:d.size.height,rotation:0},childSpace:{x:0,y:0,width:d.size.width,height:d.size.height},children:members}];
  return d;
}
let state: PdfViewState | undefined;
const errors: string[] = [];
const root = () => { const e = document.createElement('main'); document.body.replaceChildren(e); return e; };
const smoke = {
  async output(kind: 'writer' | 'slides') { if (state) disposePdfViewState(state); state = undefined; const model = kind === 'writer' ? writer() : slides(); return prepareDocumentPdfOutput(new DocumentSession(model).snapshot(), root()); },
  async open(bytes: number[]) { if (state) disposePdfViewState(state); state = createPdfViewState(new Uint8Array(bytes)); const host = root(); host.style.height = '95vh'; const view = pdfWorkspace({ t: (key, args) => message('en-US', key, args), error: e => errors.push(String(e)), isCurrent: () => true, print: () => {} }, state); host.append(view); },
  async password(bytes: number[]) { try { const r = await PdfReader.open(new Uint8Array(bytes), { assetBaseUrl: new URL('./pdf-assets/', location.href).href }); await r.destroy(); return 'accepted'; } catch (e) { return e instanceof Error ? e.message : String(e); } },
  async blockedFont() { const doc = writer(); doc.defaultTextStyle.fontFamily = 'Unreviewed Commercial Font'; try { await prepareDocumentPdfOutput(new DocumentSession(doc).snapshot(), root()); return false; } catch (e) { return e instanceof Error ? e.message : String(e); } },
  async sheetsBlocked() { try { await prepareDocumentPdfOutput(new DocumentSession(createDocument('sheets', 'pdf-sheets')).snapshot(), root()); return false; } catch (e) { return e instanceof Error ? e.message : String(e); } },
  async printPdf(bytes: number[]) { return prepareExistingPdfPrint(new Uint8Array(bytes), root()); },
  detach() { state?.detach?.(); },
  reattach() { if (!state) return; const host = root(); host.style.height = '95vh'; host.append(pdfWorkspace({ t: (key, args) => message('en-US', key, args), error: e => errors.push(String(e)), isCurrent: () => true, print: () => {} }, state)); },
  close() { if (state) disposePdfViewState(state); },
  errors() { return errors; },
};
Object.assign(window, { pdfSmoke: smoke });
