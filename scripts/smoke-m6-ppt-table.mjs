import {prepareTabbedProfile} from './tabbed-test-profile.mjs';
// SPDX-License-Identifier: MIT
import {revealControl} from './ribbon-test-controls.mjs';
import {_electron as electron} from 'playwright-core';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,copyFile,readFile,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {openSlidesPptx,openOpc} from '@office-ex/ooxml';
const fixture=resolve(process.argv[2]??'.work/m6-conditional-corpus/advanced.pptx');
const work=await mkdtemp(join(tmpdir(),'office-ex-m6-ppt-table-')),profile=join(work,'profile'),file=join(work,'table.pptx');await prepareTabbedProfile(profile);await copyFile(fixture,file);
const original=await readFile(file),checks=[];let app;
const check=name=>{checks.push(name);console.log('PASS',name);};
async function wait(fn,label){const end=Date.now()+30000;while(Date.now()<end){try{if(await fn())return;}catch{}await new Promise(r=>setTimeout(r,100));}throw Error('Timed out: '+label);}
try{
 app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env:{...process.env,OFFICE_EX_PROFILE_DIR:profile},timeout:30000});
 const page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]');const errors=[];page.on('pageerror',e=>errors.push(String(e)));
 const prefs=await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences());assert.equal(prefs.sandbox,true);assert.equal(prefs.contextIsolation,true);assert.equal(prefs.nodeIntegration,false);
 await page.locator('#settings').click();await page.locator('#language').selectOption('en-US');await page.locator('#settings-dialog button[value=close]').click();
 const open=async()=>{await app.evaluate(({dialog},path)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[path]});},file);await page.locator('#open').click();await page.waitForSelector('.slides-stage .slides-native-table');};
 await open();const table=page.locator('.slides-stage .slides-native-table').first();await table.dblclick();await page.locator('#slide-table-dialog[open]').waitFor();
 await page.locator('#slide-table-cells textarea[data-row="1"][data-column="0"]').fill('已编辑表格');await page.locator('#slide-table-apply').click();await page.waitForFunction(()=>!document.querySelector('#slide-table-dialog[open]'));
 await page.locator('#save').click();await wait(async()=>openSlidesPptx(await readFile(file),'native').document.slides[0].shapes.find(s=>s.type==='table').table.cells[1][0]==='已编辑表格','edited cell saved');
 const after=openOpc(await readFile(file));for(const [name,part]of openOpc(original).parts)if(!name.match(/^ppt\/slides\/slide1\.xml$/))assert.deepEqual(after.parts.get(name),part,name);
 check('Double-click real PPTX table, edit cell and preserve all other source parts');
 await (await revealControl(page,'#slide-insert-table')).click();await page.locator('#slide-table-rows').fill('2');await page.locator('#slide-table-columns').fill('3');await page.locator('#slide-table-apply').click();await page.waitForFunction(()=>document.querySelectorAll('.slides-stage .slides-native-table').length===2);
 await page.locator('#save').click();await wait(async()=>openSlidesPptx(await readFile(file),'native').document.slides[0].shapes.filter(s=>s.type==='table').length===2,'new native table saved');check('Real insert dialog creates native rectangular table');
 await page.locator('.tab[aria-selected=true] .tab-close').click();await page.waitForFunction(()=>document.querySelectorAll('.tab').length===0);await open();assert.equal(await page.locator('.slides-stage .slides-native-table').count(),2);assert.deepEqual(errors,[]);check('Both native tables reopen editable with no renderer errors');
 await app.close();app=undefined;await mkdir('out/m6',{recursive:true});await writeFile('out/m6/native-ppt-table.json',JSON.stringify({schema:1,platform:process.platform,arch:process.arch,sandbox:true,checks,status:'passed'},null,2)+'\n');
}finally{if(app)await app.evaluate(({app})=>app.exit()).catch(()=>{});await rm(work,{recursive:true,force:true,maxRetries:5,retryDelay:100});}
