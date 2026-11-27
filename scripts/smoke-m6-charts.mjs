import {prepareTabbedProfile} from './tabbed-test-profile.mjs';
// SPDX-License-Identifier: MIT
/** Real desktop chart overlay and conditional rule controls, original XLSX save. */
import {revealControl} from './ribbon-test-controls.mjs';
import {_electron as electron} from 'playwright-core';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,copyFile,readFile,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {openSheetsXlsx,openOpc} from '@office-ex/ooxml';
const fixtures=resolve(process.argv[2]??'.work/m6-chart-corpus');
const work=await mkdtemp(join(tmpdir(),'office-ex-m6-chart-')),profile=join(work,'profile');await prepareTabbedProfile(profile);
const file=join(work,'charts.xlsx');await copyFile(join(fixtures,'m6-charts.xlsx'),file);
const original=await readFile(file),checks=[];let app;
const check=name=>{checks.push(name);console.log('PASS',name);};
async function wait(fn,label){const end=Date.now()+30000;while(Date.now()<end){try{if(await fn())return;}catch{}await new Promise(r=>setTimeout(r,100));}throw Error('Timed out: '+label);}
try{
 app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env:{...process.env,OFFICE_EX_PROFILE_DIR:profile},timeout:30000});
 const page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]');const errors=[];page.on('pageerror',e=>errors.push(String(e)));
 const prefs=await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences());assert.equal(prefs.sandbox,true);assert.equal(prefs.contextIsolation,true);assert.equal(prefs.nodeIntegration,false);
 await page.locator('#settings').click();await page.locator('#language').selectOption('en-US');await page.locator('#settings-dialog button[value=close]').click();
 const open=async()=>{await app.evaluate(({dialog},path)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[path]});},file);await page.locator('#open').click();await page.waitForSelector('#sheets-scroll');};
 const cell=addr=>page.locator('#sheets-scroll .sheets-cell[data-address="'+addr+'"]');
 const selected=async addr=>{await page.locator('#sheets-address').fill(addr);await page.locator('#sheets-address').press('Enter');};
 await open();const column=page.locator('.sheets-chart[data-chart-type=column]');await column.waitFor();
 assert.ok(await column.locator('svg rect').count()>=3);assert.equal(await cell('B3').evaluate(el=>getComputedStyle(el).backgroundColor),'rgb(255, 170, 0)');
 const originalSVG=await column.locator('svg').innerHTML();
 await cell('B2').dblclick();const editor=page.locator('#sheets-cell-editor');await editor.waitFor({state:'visible'});await editor.fill('17');await editor.press('Enter');
 await page.waitForFunction(()=>document.querySelector('[data-address=B2]')?.textContent==='17');
 assert.equal(await cell('B2').evaluate(el=>getComputedStyle(el).backgroundColor),'rgb(255, 170, 0)');assert.notEqual(await column.locator('svg').innerHTML(),originalSVG);
 check('Imported native chart and conditional fill respond to a real cell edit');
 await page.locator('#save').click();await wait(async()=>openSheetsXlsx(await readFile(file),'native','Charts').document.sheets[0].cells['2:2'].value.value===17,'source edit');
 const preserved=openOpc(await readFile(file));for(const [name,part]of openOpc(original).parts)if(name.startsWith('xl/drawings/')||name.startsWith('xl/charts/'))assert.deepEqual(preserved.parts.get(name),part,name);
 await selected('A2:A4');await (await revealControl(page,'#sheet-conditional-open')).click();await page.locator('#sheet-conditional-value').fill('0');await page.locator('#sheet-conditional-color').fill('#22cc44');await page.locator('#sheet-conditional-add').click();
 await page.waitForFunction(()=>!document.querySelector('#sheet-conditional-dialog[open]'));
 await selected('A1:B4');await (await revealControl(page,'#sheet-chart-open')).click();await page.locator('#sheet-chart-type').selectOption('line');await page.locator('#sheet-chart-add').click();
 await page.waitForFunction(()=>!document.querySelector('#sheet-chart-dialog[open]'));
 await page.locator('#save').click();await wait(async()=>{const doc=openSheetsXlsx(await readFile(file),'native','Charts').document;return doc.sheets[0].charts.length===4&&doc.sheets[0].conditionalRules.length===2;},'chart and rules save');
 check('Native controls add a rule and a real chart; saved worksheet contains the chart graph and differential rules');
 await page.locator('.tab[aria-selected=true] .tab-close').click();await page.waitForFunction(()=>document.querySelectorAll('.tab').length===0);await open();await column.waitFor();
 assert.equal(await cell('B2').textContent(),'17');assert.deepEqual(errors,[]);check('Saved chart and rule workbook reopens editable without renderer exceptions');
 await app.close();app=undefined;await mkdir('out/m6',{recursive:true});await writeFile('out/m6/native-charts.json',JSON.stringify({schema:1,platform:process.platform,arch:process.arch,sandbox:true,checks,status:'passed'},null,2)+'\n');
}finally{if(app)await app.evaluate(({app})=>app.exit()).catch(()=>{});await rm(work,{recursive:true,force:true,maxRetries:5,retryDelay:100});}
