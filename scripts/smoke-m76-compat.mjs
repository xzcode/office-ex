// SPDX-License-Identifier: MIT
/** Three source-preserving editors through the real sandboxed save IPC. */
import {_electron as electron} from 'playwright-core';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,readFile,writeFile,copyFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {openOpc,openWriterDocxCompatibility,openCompatibleOfficeValues} from '@office-ex/ooxml';
const fixtures=resolve(process.argv[2]??'.work/m76-office-corpus');
const work=await mkdtemp(join(tmpdir(),'office-ex-m76-')),profile=join(work,'profile');
await mkdir(profile);
let app;
const checks=[];
const check=name=>{checks.push(name);console.log('PASS',name);};
async function wait(fn,label,timeout=30000){const end=Date.now()+timeout;let last;while(Date.now()<end){try{const r=await fn();if(r)return r;}catch(e){last=String(e);}await new Promise(r=>setTimeout(r,100));}throw Error('Timed out '+label+(last?': '+last:''));}
try{
 app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env:{...process.env,OFFICE_EX_PROFILE_DIR:profile},timeout:30000});
 const home=await app.firstWindow();await home.waitForSelector('body[data-ready=true]');
 await home.locator('#settings').click();await home.locator('#language').selectOption('en-US');await home.locator('#settings-dialog button[value=close]').click();
 async function open(path){const opening=app.waitForEvent('window',{timeout:30000});await app.evaluate(({dialog},file)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[file]});},path);await home.locator('#open').click();const page=await opening;await page.waitForSelector('#compatibility-edit-start');return page;}
 const cases=[{format:'docx',before:'Original Word',after:'Native Word',part:'word/document.xml'},{format:'xlsx',before:'20',after:'78.5',part:'xl/worksheets/sheet1.xml'},{format:'pptx',before:'Original slide',after:'Native PPT',part:'ppt/slides/slide1.xml'}];
 for(const {format,before,after,part}of cases){
  const path=join(work,'Document '+format+'.'+format);await copyFile(join(fixtures,'compat.'+format),path);const original=await readFile(path);
  let page=await open(path),errors=[];page.on('pageerror',e=>errors.push(String(e)));
  await page.locator('#compatibility-edit-start').click();await page.waitForSelector('#compatibility-workspace');
  const id=await page.locator('[data-compat-id]').evaluateAll((els,old)=>els.find(el=>el.value===old)?.dataset.compatId,before);assert.notEqual(id,undefined);
  const field=()=>page.locator('[data-compat-id="'+id+'"]');
  await field().fill(after);assert.equal(await field().inputValue(),after);
  await page.locator('#undo').click();assert.equal(await field().inputValue(),before);
  await page.locator('#redo').click();assert.equal(await field().inputValue(),after);
  const status=page.locator('#document-status');assert.match(await status.textContent(),/Unsaved|Not saved/i);
  await page.locator('#save').click();
  const value=bytes=>{if(format==='docx'){const v=openWriterDocxCompatibility(bytes);return v.blocks[0].items.find(x=>x.kind==='text').segment.text;}const v=openCompatibleOfficeValues(bytes,format);return v.pages[0].fields.find(f=>f.id===Number(id)).text;};
  await wait(async()=>value(await readFile(path))===after,'real compatible '+format+' save');
  await wait(async()=>/Saved/i.test(await status.textContent())&&!/Unsaved|Not saved/i.test(await status.textContent()),'saved status');
  const saved=await readFile(path),a=openOpc(original),b=openOpc(saved);
  for(const [name,data]of a.parts)if(name!==part)assert.deepEqual(b.parts.get(name),data,name);
  // Repeat save from the original baseline to verify stable IDs and
  // cumulative edits, not a stale newly imported snapshot.
  await field().fill(after+'2');if(format==='xlsx')await field().fill('81.5');
  const next=format==='xlsx'?'81.5':after+'2';
  await page.locator('#save').click();await wait(async()=>value(await readFile(path))===next,'second save');
  await wait(async()=>/Saved/i.test(await status.textContent())&&!/Unsaved|Not saved/i.test(await status.textContent()),'second saved status');
  const closing=page.waitForEvent('close',{timeout:30000});await page.locator('.tab-close').click();await closing;
  page=await open(path);await page.locator('#compatibility-edit-start').click();await page.waitForSelector('#compatibility-workspace');
  assert.ok(await page.locator('[data-compat-id]').evaluateAll((els,v)=>els.some(el=>el.value===v),next));
  const closed=page.waitForEvent('close',{timeout:30000});await page.locator('.tab-close').click();await closed;
  assert.deepEqual(errors,[]);check(format.toUpperCase()+' native compatibility edit, undo/redo, two saves, exact unrelated parts and reopen');
 }
 const prefs=await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows().map(w=>w.webContents.getLastWebPreferences()));
 assert.ok(prefs.every(p=>p.sandbox&&p.contextIsolation&&!p.nodeIntegration));
 await app.close();app=undefined;await mkdir('out/m7',{recursive:true});await writeFile('out/m7/native-m76-compat.json',JSON.stringify({schema:1,platform:process.platform,arch:process.arch,sandbox:true,checks,status:'passed'},null,2)+'\n');
}finally{if(app)await app.evaluate(({app})=>app.exit()).catch(()=>{});await rm(work,{recursive:true,force:true,maxRetries:5,retryDelay:100});}
