import {prepareTabbedProfile} from './tabbed-test-profile.mjs';
// SPDX-License-Identifier: MIT
/** Actual sandboxed desktop merged-table editor and atomic save/reopen. */
import {_electron as electron} from 'playwright-core';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,copyFile,readFile,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {openSlidesPptx,openOpc} from '@office-ex/ooxml';
const fixture=resolve(process.argv[2]??'.work/m7-ppt-table-corpus/merged.pptx');
const work=await mkdtemp(join(tmpdir(),'office-ex-m7-table-')),profile=join(work,'profile'),file=join(work,'merged.pptx');await prepareTabbedProfile(profile);await copyFile(fixture,file);
const original=await readFile(file),checks=[];let app;
const check=name=>{checks.push(name);console.log('PASS',name);};
async function wait(fn,label){const end=Date.now()+30000;while(Date.now()<end){try{if(await fn())return;}catch{}await new Promise(r=>setTimeout(r,100));}throw Error('Timed out: '+label);}
try{
 app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env:{...process.env,OFFICE_EX_PROFILE_DIR:profile},timeout:30000});
 const page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]');const errors=[];page.on('pageerror',e=>errors.push(String(e)));
 const prefs=await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences());assert.equal(prefs.sandbox,true);assert.equal(prefs.contextIsolation,true);assert.equal(prefs.nodeIntegration,false);
 await page.locator('#settings').click();await page.locator('#language').selectOption('en-US');await page.locator('#settings-dialog button[value=close]').click();
 const open=async()=>{await app.evaluate(({dialog},path)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[path]});},file);await page.locator('#open').click();await page.waitForSelector('.slides-stage .slides-native-table');};
 const grid=page.locator('.slides-stage .slides-native-table').first(),cell=(r,c)=>grid.locator('.slides-table-cell[data-row="'+r+'"][data-column="'+c+'"]'),editor=(r,c)=>page.locator('#slide-table-cells textarea[data-row="'+r+'"][data-column="'+c+'"]');
 const dialog=page.locator('#slide-table-dialog');
 await open();assert.equal(await grid.locator('.slides-table-cell').count(),10);
 assert.equal(await cell(0,0).evaluate(e=>e.style.gridColumnEnd),'span 3');
 assert.equal(await cell(0,0).evaluate(e=>e.style.gridRowEnd),'span 2');
 assert.equal(await cell(2,2).evaluate(e=>getComputedStyle(e).backgroundColor),'rgb(20, 40, 120)');
 assert.equal(await cell(2,2).evaluate(e=>getComputedStyle(e).fontWeight),'700');
 check('Imported merged native grid displays actual spanning cells and direct cell paint');
 await cell(0,0).dblclick();await dialog.waitFor({state:'visible'});
 await editor(0,0).fill('Cancelled edit');await page.locator('#slide-table-cancel').click();assert.equal(await cell(0,0).textContent(),'0,0\n0,1\n0,2\n1,0\n1,1\n1,2');
 check('Cancel discards the entire table edit draft');
 await cell(0,0).dblclick();await dialog.waitFor({state:'visible'});
 await editor(0,0).click();await page.locator('#slide-table-add-row').click();
 assert.equal(await editor(0,0).evaluate(e=>e.style.gridRowEnd),'span 3');
 await editor(3,1).click();await page.locator('#slide-table-split').click();
 assert.equal(await editor(4,1).count(),1);
 await editor(3,3).click();await page.locator('#slide-table-add-column').click();
 await editor(3,2).fill('Edited native');
 await editor(3,2).click();await page.locator('#slide-table-italic').click();
 await editor(3,4).click();await page.locator('#slide-table-fill').fill('#64c828');
 await page.locator('#slide-table-border-color').fill('#334455');
 await page.locator('#slide-table-border-width').fill('1.5');await page.locator('#slide-table-apply-border').click();
 await page.locator('#slide-table-apply').click();await page.waitForFunction(()=>!document.querySelector('#slide-table-dialog[open]'));
 assert.equal(await cell(0,0).evaluate(e=>e.style.gridRowEnd),'span 3');
 assert.equal(await cell(3,2).textContent(),'Edited native');assert.equal(await cell(3,2).evaluate(e=>getComputedStyle(e).fontStyle),'italic');
 assert.equal(await cell(3,4).evaluate(e=>getComputedStyle(e).backgroundColor),'rgb(100, 200, 40)');
 await page.locator('#undo').click();assert.equal(await cell(0,0).evaluate(e=>e.style.gridRowEnd),'span 2');assert.equal(await cell(2,2).textContent(),'2,2');
 await page.locator('#redo').click();assert.equal(await cell(3,2).textContent(),'Edited native');
 check('Native row/column insert, split, font, fill and border edits apply in one undoable transaction');
 await page.locator('#save').click();await wait(async()=>{const d=openSlidesPptx(await readFile(file),'native').document,t=d.slides[0].shapes[0].table;return t.rowHeights.length===5&&t.columnWidths.length===5&&t.cells[3][2]==='Edited native';},'native table save');
 const saved=await readFile(file),out=openSlidesPptx(saved,'native').document.slides[0].shapes[0].table;assert.equal(out.merges.length,1);assert.equal(out.merges[0].rowCount,3);assert.equal(out.formats['3:4'].borders.left.width,19050);
 const a=openOpc(original),b=openOpc(saved);for(const [name,data]of a.parts)if(!['ppt/slides/slide1.xml','[Content_Types].xml'].includes(name))assert.deepEqual(b.parts.get(name),data,name);
 await page.locator('.tab[aria-selected=true] .tab-close').click();await page.waitForFunction(()=>document.querySelectorAll('.tab').length===0);await open();
 assert.equal(await cell(3,2).textContent(),'Edited native');assert.equal(await cell(0,0).evaluate(e=>e.style.gridRowEnd),'span 3');assert.deepEqual(errors,[]);
 check('Saved PPTX reopens with merged structure, direct formatting and unchanged source parts');
 await app.close();app=undefined;await mkdir('out/m6',{recursive:true});await writeFile('out/m6/native-m7-ppt-tables.json',JSON.stringify({schema:1,platform:process.platform,arch:process.arch,sandbox:true,checks,status:'passed'},null,2)+'\n');
}finally{if(app)await app.evaluate(({app})=>app.exit()).catch(()=>{});await rm(work,{recursive:true,force:true,maxRetries:5,retryDelay:100});}
