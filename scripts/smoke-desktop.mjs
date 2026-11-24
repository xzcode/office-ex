/** Runs the real Electron binary; no sandbox disabling and no product test backdoor. */
import { _electron as electron } from 'playwright-core';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, writeFile, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { createHash } from 'node:crypto';
import { docxParts, writeZip } from '@office-ex/ooxml';

const work=await mkdtemp(join(tmpdir(),'office-ex-native-'));
const profile=join(work,'profile');await mkdir(profile);
await mkdir('out/native',{recursive:true});
const name=`${process.platform}-${process.arch}`;
const evidence={platform:process.platform,arch:process.arch,node:process.versions.node,checks:[],startedAt:new Date().toISOString()};
let app,page;
const check=title=>{evidence.checks.push(title);console.log(`PASS ${title}`);};
const originalPath=join(work,'office-source.docx'),copyPath=join(work,'office-copy.docx'),draftPath=join(work,'draft.oex'),draftCopyPath=join(work,'draft-copy.oex');
const parts=docxParts([{runs:[{text:'Independent Office EX container smoke 中文 😀'}]}]);
const original=writeZip(Object.fromEntries(Object.entries(parts).map(([key,value])=>[key,new TextEncoder().encode(value)])));
await writeFile(originalPath,original);
const bytesHash=value=>createHash('sha256').update(value).digest('hex');
const waitFor=async(predicate,label,timeout=15_000)=>{const end=Date.now()+timeout;while(Date.now()<end){if(await predicate())return;await new Promise(r=>setTimeout(r,200));}throw new Error(`Timed out: ${label}`);};
async function launch(){const env={...process.env,OFFICE_EX_PROFILE_DIR:profile};delete env.ELECTRON_RUN_AS_NODE;app=await electron.launch({args:[resolve('apps/desktop')],env,timeout:30_000});page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]',{timeout:30_000});}
async function dialogSave(path){await app.evaluate(({dialog},filePath)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath});dialog.showMessageBox=async()=>({response:1,checkboxChecked:false});},path);}
async function dialogOpen(path){await app.evaluate(({dialog},filePath)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[filePath]});},path);await page.locator('#open').click();}
try{
  await launch();
  evidence.electron=await app.evaluate(({app})=>process.versions.electron);evidence.chromium=await page.evaluate(()=>navigator.userAgent);
  const prefs=await app.evaluate(({BrowserWindow})=>{const p=BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences();return {sandbox:p.sandbox,contextIsolation:p.contextIsolation,nodeIntegration:p.nodeIntegration,webSecurity:p.webSecurity};});
  assert.deepEqual(prefs,{sandbox:true,contextIsolation:true,nodeIntegration:false,webSecurity:true});
  assert.deepEqual(await page.evaluate(()=>[typeof window.require,typeof window.process]),['undefined','undefined']);check('Native window + sandbox, context isolation, no renderer Node');
  await page.screenshot({path:`out/native/${name}-home.png`});
  const denied=await page.evaluate(async()=>{const results=await Promise.allSettled([fetch('https://example.com/'),fetch('file:///etc/passwd')]);window.open('https://example.com/');return results.map(r=>r.status);});
  assert.deepEqual(denied,['rejected','rejected']);assert.equal(await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows().length),1);check('CSP rejects remote/local fetch; new windows denied');
  const invalid=await page.evaluate(()=>window.office.saveCopy({documentId:'fake',sourceToken:'fake'}));assert.equal(invalid.ok,false);assert.equal(invalid.error.code,'unsafe');
  const syntheticDrop=await page.evaluate(()=>window.office.openDroppedFile(new File(['data'],'forged.docx')));assert.equal(syntheticDrop.ok,false);check('Forged capability and synthetic dropped file rejected');
  await page.locator('[data-kind=writer]').click();await page.locator('#draft-text').fill('中文草稿 😀\nFirst saved revision');
  await dialogSave(draftPath);await page.locator('#save').click();await waitFor(async()=>{try{return JSON.parse(await readFile(draftPath,'utf8')).document.sections[0].blocks[0].runs[0].text.includes('First saved');}catch{return false;}},'first draft save');
  await page.waitForFunction(()=>document.querySelector('#document-status').textContent.includes('已保存'));
  const firstDocument=JSON.parse(await readFile(draftPath,'utf8')).document;check('Draft → worker serialization → atomic save → re-read');
  await page.locator('#draft-text').fill('Changed text');await page.locator('#undo').click();assert.equal(await page.locator('#draft-text').inputValue(),'中文草稿 😀\nFirst saved revision');await page.waitForFunction(()=>document.querySelector('#document-status').textContent.includes('已保存'));check('Undo restores saved content identity and clean state');
  // Actual DOM composition event path: one transaction for the completed input.
  const revisionBefore=await page.locator('#document-status').textContent();
  await page.locator('#draft-text').evaluate(el=>{el.dispatchEvent(new CompositionEvent('compositionstart',{bubbles:true,data:''}));el.value='中文输入法提交 😀';el.dispatchEvent(new InputEvent('input',{bubbles:true,inputType:'insertCompositionText',isComposing:true,data:'中'}));el.dispatchEvent(new CompositionEvent('compositionend',{bubbles:true,data:'中文输入法提交 😀'}));});
  assert.notEqual(await page.locator('#document-status').textContent(),revisionBefore);await page.locator('#undo').click();assert.equal(await page.locator('#draft-text').inputValue(),'中文草稿 😀\nFirst saved revision');check('Composition commit is one undo transaction (event simulation)');
  // Save As holds a native-dialog promise while a later edit arrives in the renderer.
  await app.evaluate(({dialog},filePath)=>{dialog.showSaveDialog=()=>new Promise(resolve=>{globalThis.__officeSmokeRelease=()=>resolve({canceled:false,filePath});});},draftCopyPath);
  await page.locator('#save-as').click();await waitFor(async()=>app.evaluate(()=>typeof globalThis.__officeSmokeRelease==='function'),'save snapshot reached dialog');
  await page.locator('#draft-text').fill('Edit made while saving');await app.evaluate(()=>globalThis.__officeSmokeRelease());
  await waitFor(async()=>{try{return Boolean(await readFile(draftCopyPath));}catch{return false;}},'save-as output');
  await page.waitForFunction(()=>!document.querySelector('#save').disabled);
  assert.equal(JSON.parse(await readFile(draftCopyPath,'utf8')).document.sections[0].blocks[0].runs[0].text,'中文草稿 😀\nFirst saved revision');
  assert.match(await page.locator('#document-status').textContent(),/未保存/);check('Concurrent edit remains dirty after snapshot save');
  // Existing target save requires and creates a previous-version private backup.
  await dialogSave(draftCopyPath);await page.locator('#save').click();await waitFor(async()=>JSON.parse(await readFile(draftCopyPath,'utf8')).document.sections[0].blocks[0].runs[0].text==='Edit made while saving','overwrite');
  assert((await readdir(join(profile,'backups'))).some(name=>name.endsWith('.oex-recovery')));check('Overwrite preserves previous version in private backups');
  await dialogOpen(originalPath);await page.waitForSelector('.parts-table');assert.match(await page.locator('#content').textContent(),/DOCX/);assert.match(await page.locator('#content').textContent(),/尚未导入/);
  await dialogSave(copyPath);await page.locator('#save').click();await waitFor(async()=>{try{return (await readFile(copyPath)).length>0;}catch{return false;}},'original copy');assert.equal(bytesHash(await readFile(copyPath)),bytesHash(original));assert.equal(bytesHash(await readFile(originalPath)),bytesHash(original));check('DOCX container inspection + byte-exact Save As with source unchanged');
  await page.screenshot({path:`out/native/${name}-inspection.png`});
  await page.locator('.tab').first().click();assert.equal(await page.locator('#draft-text').inputValue(),'Edit made while saving');
  // Bad save destination keeps the document dirty, usable and the old file unchanged.
  const stable=await readFile(draftCopyPath);await page.locator('#draft-text').fill('Dirty content before simulated crash');
  await dialogSave(join(work,'wrong.docx'));await page.locator('#save-as').click();await page.waitForSelector('#message.error');assert.equal(bytesHash(await readFile(draftCopyPath)),bytesHash(stable));assert.match(await page.locator('#document-status').textContent(),/未保存/);check('Wrong-extension draft save fails without changing source or dirty state');
  // Wait for the real 30-second checkpoint interval instead of calling a test-only API.
  await waitFor(async()=>{try{for(const name of await readdir(join(profile,'recovery'))){if(!name.endsWith('.oex-recovery'))continue;const b=await readFile(join(profile,'recovery',name));const doc=JSON.parse(b.subarray(4+b.readUInt32LE(0)).toString()).document;if(doc.sections?.[0]?.blocks?.[0]?.runs?.[0]?.text==='Dirty content before simulated crash')return true;}return false;}catch{return false;}},'automatic recovery checkpoint',40_000);check('Real 30-second automatic recovery checkpoint');
  await page.locator('#settings').click();await page.locator('#theme').selectOption('dark');await page.locator('#language').selectOption('en');await page.locator('#settings-dialog button[value=close]').click();
  await page.screenshot({path:`out/native/${name}-draft-dark.png`});
  // Hard process exit leaves the checkpoint for the next launch, bypassing normal close UI.
  await app.evaluate(({app})=>app.exit(0));await app.close().catch(()=>undefined);app=undefined;
  await launch();assert.equal(await page.locator('html').getAttribute('data-theme'),'dark');assert.equal(await page.locator('html').getAttribute('lang'),'en');
  await page.getByRole('button',{name:'Recover a copy',exact:true}).first().click();await page.waitForSelector('#draft-text');assert.equal(await page.locator('#draft-text').inputValue(),'Dirty content before simulated crash');assert.match(await page.locator('#document-status').textContent(),/Unsaved/);assert.equal(bytesHash(await readFile(draftCopyPath)),bytesHash(stable));check('Crash recovery reopens an unsaved copy; settings persist; source remains unchanged');
  await app.evaluate(({dialog})=>{dialog.showMessageBox=async()=>({response:2,checkboxChecked:false});});await page.locator('.tab-close').click();assert.equal(await page.locator('.tab').count(),1);check('Dirty close cancellation keeps document open');
  await app.evaluate(({dialog})=>{dialog.showMessageBox=async()=>({response:1,checkboxChecked:false});});await page.locator('.tab-close').click();await page.waitForFunction(()=>document.querySelectorAll('.tab').length===0);check('Confirmed discard closes the session');
  evidence.status='passed';evidence.finishedAt=new Date().toISOString();
}catch(error){evidence.status='failed';evidence.error=String(error?.stack??error);console.error(evidence.error);if(page)await page.screenshot({path:`out/native/${name}-failure.png`}).catch(()=>undefined);process.exitCode=1;}
finally{await writeFile(`out/native/${name}.json`,JSON.stringify(evidence,null,2)+'\n');if(app)await app.evaluate(({app})=>app.exit(process.exitCode?1:0)).catch(()=>undefined);if(app)await app.close().catch(()=>undefined);await rm(work,{recursive:true,force:true});}
