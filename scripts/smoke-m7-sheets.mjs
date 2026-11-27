import {prepareTabbedProfile} from './tabbed-test-profile.mjs';
// SPDX-License-Identifier: MIT
/** Native rich text display/edit and named style controls using independent files. */
import {revealControl} from './ribbon-test-controls.mjs';
import {_electron as electron} from 'playwright-core';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,copyFile,readFile,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {openSheetsXlsx} from '@office-ex/ooxml';
const fixtures=resolve(process.argv[2]??'.work/m7-sheet-corpus'),work=await mkdtemp(join(tmpdir(),'office-ex-m7-sheet-')),profile=join(work,'profile');
await prepareTabbedProfile(profile);const file=join(work,'inline.xlsx');await copyFile(join(fixtures,'inline.xlsx'),file);
const checks=[];let app;const check=name=>{checks.push(name);console.log('PASS',name);};
async function wait(fn,label){const end=Date.now()+30000;while(Date.now()<end){try{if(await fn())return;}catch{}await new Promise(r=>setTimeout(r,100));}throw Error('Timed out: '+label);}
try{
 app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env:{...process.env,OFFICE_EX_PROFILE_DIR:profile},timeout:30000});
 const page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]');const errors=[];page.on('pageerror',e=>errors.push(String(e)));
 const prefs=await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences());assert.equal(prefs.sandbox,true);assert.equal(prefs.contextIsolation,true);assert.equal(prefs.nodeIntegration,false);
 await page.locator('#settings').click();await page.locator('#language').selectOption('en-US');await page.locator('#settings-dialog button[value=close]').click();
 const open=async()=>{await app.evaluate(({dialog},path)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[path]});},file);await page.locator('#open').click();await page.waitForSelector('#sheets-scroll');};
 const cell=addr=>page.locator('#sheets-scroll .sheets-cell[data-address="'+addr+'"]');
 await open();const rich=cell('A1').locator('.sheets-rich-run');await rich.first().waitFor();
 assert.equal(await rich.count(),2);assert.equal(await rich.first().evaluate(e=>getComputedStyle(e).fontWeight),'700');assert.equal(await rich.first().evaluate(e=>getComputedStyle(e).color),'rgb(170, 34, 68)');assert.equal(await rich.last().evaluate(e=>getComputedStyle(e).fontStyle),'italic');
 check('Native rich cell renders its separate font and color runs');
 await cell('A1').dblclick();const editor=page.locator('#sheets-cell-editor');await editor.waitFor({state:'visible'});await editor.fill('Title New Tail');await editor.press('Enter');
 await page.waitForFunction(()=>document.querySelector('[data-address=A1]')?.textContent==='Title New Tail');
 assert.equal(await rich.count(),2);assert.equal(await rich.first().textContent(),'Title');assert.equal(await rich.last().textContent(),' New Tail');
 check('Native cell editing preserves unaffected rich text runs');
 await cell('C1').click();const style=await revealControl(page,'#sheet-named-style');await style.selectOption('Report Format');
 await page.locator('#save').click();await wait(async()=>{const d=openSheetsXlsx(await readFile(file),'native').document;return d.sheets[0].cells['1:3']?.namedStyle==='Report Format'&&d.sheets[0].cells['1:1'].value.value==='Title New Tail';},'native rich and named style save');
 await page.locator('.tab[aria-selected=true] .tab-close').click();await page.waitForFunction(()=>document.querySelectorAll('.tab').length===0);await open();
 assert.equal(await cell('A1').textContent(),'Title New Tail');assert.equal(await cell('A1').locator('.sheets-rich-run').count(),2);
 await cell('C1').click();assert.equal(await (await revealControl(page,'#sheet-named-style')).inputValue(),'Report Format');assert.deepEqual(errors,[]);
 check('Named style applied by actual Ribbon control and both features reopen from native XLSX');
 await app.close();app=undefined;await mkdir('out/m6',{recursive:true});await writeFile('out/m6/native-m7-sheets.json',JSON.stringify({schema:1,platform:process.platform,arch:process.arch,sandbox:true,checks,status:'passed'},null,2)+'\n');
}finally{if(app)await app.evaluate(({app})=>app.exit()).catch(()=>{});await rm(work,{recursive:true,force:true,maxRetries:5,retryDelay:100});}
