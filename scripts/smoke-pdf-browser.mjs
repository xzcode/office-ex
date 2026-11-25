// SPDX-License-Identifier: MIT
import { build } from 'esbuild';
import { chromium, _electron } from 'playwright-core';
import { createServer } from 'node:http';
import { readFile, mkdir, writeFile, readdir } from 'node:fs/promises';
import { resolve, join, extname, relative } from 'node:path';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';
import { preparePdfWorker } from '../packages/pdf/prepare-worker.mjs';
const base = process.cwd(), work = resolve('.work/pdf-browser'), output = resolve(process.env.OFFICE_EX_PDF_OUTPUT ?? 'out/pdf-browser');
await mkdir(work, { recursive: true }); await mkdir(output, { recursive: true });
await build({ entryPoints: ['tests/pdf-browser-entry.ts'], outfile: join(work, 'entry.js'), bundle: true, platform: 'browser', format: 'esm', target: 'chrome156', legalComments: 'external' });
await writeFile(join(work,'pdf.worker.mjs'), preparePdfWorker(await readFile('node_modules/pdfjs-dist/build/pdf.worker.mjs', 'utf8')));
const files = new Map([
  ['/entry.js', join(work, 'entry.js')], ['/fonts.css', resolve('packages/ui/fonts.css')], ['/pdf-print.css', resolve('packages/ui/pdf-print.css')], ['/pdf-workspace.css', resolve('packages/ui/src/pdf-workspace.css')], ['/pdf.worker.mjs', join(work,'pdf.worker.mjs')],
]);
for (const name of await readdir('assets/fonts')) files.set('/fonts/' + name, resolve('assets/fonts', name));
for (const directory of ['cmaps', 'standard_fonts', 'wasm', 'iccs']) for (const name of await readdir('node_modules/pdfjs-dist/' + directory).catch(() => [])) files.set('/pdf-assets/' + directory + '/' + name, resolve('node_modules/pdfjs-dist', directory, name));
const csp = "default-src 'none'; script-src 'self'; style-src 'self'; font-src 'self' blob:; img-src 'self' data: blob:; worker-src 'self'; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-src 'none'; form-action 'none'";
const server = createServer(async (req, res) => {
  try {
    const pathname = new URL(req.url, 'http://localhost').pathname;
    res.setHeader('Content-Security-Policy', csp); res.setHeader('X-Content-Type-Options', 'nosniff');
    if (pathname === '/') { res.setHeader('Content-Type', 'text/html'); res.end('<!doctype html><html lang="en"><head><meta charset="UTF-8"><link rel="stylesheet" href="/fonts.css"><link rel="stylesheet" href="/pdf-print.css"><link rel="stylesheet" href="/pdf-workspace.css"><script type="module" src="/entry.js"></script></head><body></body></html>'); return; }
    const file = files.get(pathname); if (!file) { res.statusCode = 404; res.end(); return; }
    res.setHeader('Content-Type', ({ '.js': 'text/javascript', '.mjs': 'text/javascript', '.css': 'text/css', '.ttf': 'font/ttf', '.otf': 'font/otf', '.wasm': 'application/wasm' })[extname(file)] ?? 'application/octet-stream'); res.end(await readFile(file));
  } catch { res.statusCode = 500; res.end(); }
});
await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
const address = server.address(), url = `http://127.0.0.1:${address.port}/`;
let browser, electron, page;
if (process.env.OFFICE_EX_BROWSER) {
  browser = await chromium.launch({ executablePath: process.env.OFFICE_EX_BROWSER, headless: true, chromiumSandbox: true });
  page = await browser.newPage({ viewport: { width: 1280, height: 900 }, deviceScaleFactor: 1 });
} else {
  const appPath = join(work, 'app.cjs');
  await writeFile(appPath, "const {app,BrowserWindow}=require('electron');app.enableSandbox();app.disableHardwareAcceleration();app.whenReady().then(()=>{const w=new BrowserWindow({show:true,width:1280,height:900,webPreferences:{sandbox:true,contextIsolation:true,nodeIntegration:false,webSecurity:true,backgroundThrottling:false}});w.loadURL('about:blank');});app.on('window-all-closed',()=>app.quit());\n");
  electron = await _electron.launch({ executablePath: resolve('node_modules/electron/dist/electron'), args: ['--headless', '--ozone-platform=headless', appPath], chromiumSandbox: true, timeout: 30000 });
  assert.equal(await electron.evaluate(() => process.argv.some(arg => arg === '--no-sandbox' || arg === '--disable-setuid-sandbox')), false);
  page = await electron.firstWindow(); await page.setViewportSize({ width: 1280, height: 900 });
}
const exportPdf = async path => {
  if (browser) return page.pdf({ path, preferCSSPageSize: true, printBackground: true, tagged: true });
  const values = await electron.evaluate(async ({ BrowserWindow }) => Array.from(await BrowserWindow.getAllWindows()[0].webContents.printToPDF({ preferCSSPageSize: true, printBackground: true, generateTaggedPDF: true, margins: { top: 0, bottom: 0, left: 0, right: 0 } })));
  const bytes = Buffer.from(values); await writeFile(path, bytes); return bytes;
};
const capture = async path => {
  if (browser) return page.screenshot({ path, fullPage: true });
  const values = await electron.evaluate(async ({ BrowserWindow }) => Array.from((await BrowserWindow.getAllWindows()[0].webContents.capturePage()).toPNG()));
  assert(values.length > 0, 'Native capture produced no pixels'); await writeFile(path, Buffer.from(values));
};
const errors = [], external = [], consoleErrors = [], checks = [];
page.on('pageerror', e => errors.push(e.message));
page.on('console', m => { if (m.type() === 'error') consoleErrors.push(m.text()); });
await page.route('**/*', route => { const target = route.request().url(); if (!target.startsWith(url) && !target.startsWith('blob:') && !target.startsWith('data:')) { external.push(target); return route.abort(); } return route.continue(); });
try {
  await page.goto(url); await page.waitForFunction(() => !!window.pdfSmoke);
  const writer = await page.evaluate(() => window.pdfSmoke.output('writer'));
  const writerBytes = await exportPdf(join(output, 'writer.pdf'));
  assert(writer.pages >= 3); checks.push({ name: 'Writer final native layout exports multiple pages', pages: writer.pages });
  const slides = await page.evaluate(() => window.pdfSmoke.output('slides'));
  const slidesBytes = await exportPdf(join(output, 'slides.pdf'));
  assert.equal(slides.pages, 3); checks.push({ name: 'Slides shared scene exports three landscape pages', pages: slides.pages });
  await capture(join(output, 'slides-print.png'));
  assert.equal(await page.evaluate(() => window.pdfSmoke.blockedFont()), 'pdf.font-not-approved');
  assert.equal(await page.evaluate(() => window.pdfSmoke.sheetsBlocked()), 'pdf.sheets-print-rules-unavailable'); checks.push({ name: 'Unreviewed fonts and Sheets export blocked' });
  await page.evaluate(bytes => window.pdfSmoke.open(bytes), [...writerBytes]);
  await page.waitForSelector('#pdf-workspace[data-loaded="true"]');
  const selection = await page.evaluate(() => { const text = document.querySelector('.pdf-text-layer'); const range = document.createRange(); range.selectNodeContents(text); getSelection().removeAllRanges(); getSelection().addRange(range); return { selected: getSelection().toString(), text: text.textContent, spans: text.children.length, size: [text.clientWidth,text.clientHeight], visible: getComputedStyle(text).display, html: text.innerHTML.slice(0,1200), errors: window.pdfSmoke.errors() }; });
  assert.match(selection.selected, /Writer PDF/, JSON.stringify(selection)); assert.match(selection.selected, /中文/); checks.push({ name: 'PDF selectable Unicode text layer' });
  await page.fill('#pdf-search', 'UniqueLine075'); await page.click('#pdf-find'); await page.waitForSelector('.pdf-match-current');
  const foundPage = Number(await page.locator('#pdf-workspace').getAttribute('data-page')); assert(foundPage >= 3); checks.push({ name: 'Full document literal search navigates to later page', page: foundPage });
  await page.selectOption('#pdf-zoom', '1.5'); await page.waitForFunction(() => document.querySelector('.pdf-paper')?.getBoundingClientRect().width > 1000);
  await page.click('#pdf-previous-page'); await page.waitForFunction(previous => Number(document.querySelector('#pdf-workspace')?.getAttribute('data-page')) === previous - 1, foundPage);
  checks.push({ name: 'Page navigation and zoom' });
  await page.evaluate(() => window.pdfSmoke.detach()); await page.evaluate(() => window.pdfSmoke.reattach()); await page.waitForSelector('#pdf-workspace[data-loaded="true"]'); checks.push({ name: 'Tab detach retains reader and reattaches safely' });
  await capture(join(output, 'pdf-reader.png'));
  execFileSync('python', ['-c', `import fitz\nd=fitz.open();p=d.new_page();p.insert_text((50,80),'Security fixture external actions must not execute');p.insert_link({'kind':fitz.LINK_URI,'from':fitz.Rect(50,50,200,100),'uri':'https://example.invalid/office-ex-test'});d.embfile_add('untrusted.txt',b'Not executed');x=d.get_new_xref();d.update_object(x,'<< /S /JavaScript /JS (app.alert("unsafe")) >>');d.xref_set_key(d.pdf_catalog(),'OpenAction',str(x)+' 0 R');d.save(${JSON.stringify(join(output, 'actions.pdf'))});d.save(${JSON.stringify(join(output, 'encrypted.pdf'))},encryption=fitz.PDF_ENCRYPT_AES_256,owner_pw='owner',user_pw='secret')`]);
  const encrypted = await readFile(join(output, 'encrypted.pdf')), actionBytes = await readFile(join(output, 'actions.pdf'));
  assert.equal(await page.evaluate(bytes => window.pdfSmoke.password(bytes), [...encrypted]), 'pdf.encrypted'); checks.push({ name: 'Encrypted PDF rejected' });
  await page.evaluate(bytes => window.pdfSmoke.open(bytes), [...actionBytes]); await page.waitForSelector('#pdf-workspace[data-loaded="true"]');
  assert.equal(await page.locator('.pdf-text-layer a').count(), 0); assert.equal(await page.locator('.annotationLayer').count(), 0); checks.push({ name: 'Scripts, external annotation actions and attachments have no active UI' });
  const print = await page.evaluate(bytes => window.pdfSmoke.printPdf(bytes), [...slidesBytes]); assert.equal(print.pages, 3); checks.push({ name: 'Existing PDF safe page print includes all pages' });
  await page.evaluate(() => window.pdfSmoke.close());
  assert.deepEqual(external, []); assert.deepEqual(errors, []); assert.deepEqual(await page.evaluate(() => window.pdfSmoke.errors()), []);
  const runtime = browser ? { browser: await browser.version() } : await electron.evaluate(() => ({ electron: process.versions.electron, chromium: process.versions.chrome }));
  const evidence = { timestamp: new Date().toISOString(), ...runtime, headless: true, sandbox: true, csp, checks, externalRequests: external, pageErrors: errors, consoleErrors, writer, slides, artifacts: {} };
  for (const name of ['writer.pdf', 'slides.pdf', 'slides-print.png', 'pdf-reader.png']) evidence.artifacts[name] = createHash('sha256').update(await readFile(join(output, name))).digest('hex');
  await writeFile(join(output, 'result.json'), JSON.stringify(evidence, null, 2) + '\n');
  console.log(JSON.stringify({ checks: checks.length, output: relative(base, output), writerPages: writer.pages, slidesPages: slides.pages, consoleErrors }, null, 2));
} finally { if (browser) await browser.close(); if (electron) await electron.close(); await new Promise(resolve => server.close(resolve)); }
