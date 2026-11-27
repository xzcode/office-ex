import {prepareTabbedProfile} from './tabbed-test-profile.mjs';
// SPDX-License-Identifier: MIT
/** Independent workbook import, editable grid, live recalculation and native save. */
import {revealControl} from './ribbon-test-controls.mjs';
import {_electron as electron} from 'playwright-core';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,copyFile,readFile,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {openSheetsXlsx,openOpc} from '@office-ex/ooxml';
const fixtures=resolve(process.argv[2]??'.work/m6-sheet-corpus');
const work=await mkdtemp(join(tmpdir(),'office-ex-m6-sheet-')),profile=join(work,'profile');await prepareTabbedProfile(profile);
const file=join(work,'report.xlsx');await copyFile(join(fixtures,'m6-sheet.xlsx'),file);
const original=await readFile(file),checks=[];let app;
const check=name=>{checks.push(name);console.log('PASS',name);};
async function wait(fn,label){const end=Date.now()+30000;while(Date.now()<end){try{if(await fn())return;}catch{}await new Promise(r=>setTimeout(r,100));}throw Error('Timed out: '+label);}
try{
 app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env:{...process.env,OFFICE_EX_PROFILE_DIR:profile},timeout:30000});
 const page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]');const errors=[];page.on('pageerror',e=>errors.push(String(e)));
 const prefs=await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences());
 assert.equal(prefs.sandbox,true);assert.equal(prefs.contextIsolation,true);assert.equal(prefs.nodeIntegration,false);
 await page.locator('#settings').click();await page.locator('#language').selectOption('zh-CN');await page.locator('#settings-dialog button[value=close]').click();
 const open=async()=>{await app.evaluate(({dialog},path)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[path]});},file);await page.locator('#open').click();await page.waitForSelector('#sheets-scroll');};
 const cell=addr=>page.locator('#sheets-scroll .sheets-cell[data-address="'+addr+'"]');
 const value=async(addr,expected)=>page.waitForFunction(({addr,expected})=>document.querySelector('#sheets-scroll .sheets-cell[data-address="'+addr+'"]')?.textContent===expected,{addr,expected});
 await open();await value('B2','24');await value('B3','3');await value('B4','Office');await value('B5','1');
 const css=await cell('A1').evaluate(el=>{const s=getComputedStyle(el),b=el.getBoundingClientRect();return {bottom:s.borderBottomStyle,left:s.borderLeftStyle,bottomColor:s.borderBottomColor,leftColor:s.borderLeftColor,width:b.width,height:b.height};});
 assert.equal(css.bottom,'double');assert.equal(css.left,'dashed');assert.equal(css.bottomColor,'rgb(17, 136, 68)');assert.equal(css.leftColor,'rgb(204, 34, 51)');
 assert.ok(Math.abs(css.width-126)<1,JSON.stringify(css));assert.ok(Math.abs(css.height-26*12700/9525)<1,JSON.stringify(css));
 check('Independent workbook opens editable with mixed border edges and default row/column geometry');
 await cell('B1').dblclick();const editor=page.locator('#sheets-cell-editor');await editor.waitFor({state:'visible'});
 await editor.fill('21');await editor.press('Enter');await value('B2','42');
 await page.locator('#undo').click();await value('B2','24');await page.locator('#redo').click();await value('B2','42');
 check('Native double-click edit, recalculation of new functions, undo and redo');
 await page.locator('#save').click();await wait(async()=>openSheetsXlsx(await readFile(file),'native','Report').document.sheets[0].cells['1:2'].value.value===21,'source-backed XLSX save');
 await page.waitForFunction(()=>document.querySelector('#document-status')?.textContent.includes('已保存'));
 const bytes=await readFile(file),source=openSheetsXlsx(bytes,'native','Report'),sheet=source.document.sheets[0];
 assert.equal(sheet.cells['1:1'].style.borderEdges.bottom.line,'double');assert.equal(sheet.defaultGeometry.rowHeight,26*12700);assert.equal(sheet.defaultGeometry.columnWidth,18*9525*7);
 assert.ok(source.cachedValues.some(([k,v])=>k==='sheet-1!2:2'&&v.kind==='number'&&v.value===42));
 assert.ok(source.cachedValues.some(([k,v])=>k==='sheet-2!1:1'&&v.kind==='number'&&v.value===24));
 for(const name of ['xl/styles.xml','xl/theme/theme1.xml','docProps/core.xml','docProps/app.xml'])assert.deepEqual(openOpc(bytes).parts.get(name),openOpc(original).parts.get(name),name);
 await page.locator('.tab[aria-selected=true] .tab-close').click();await page.waitForFunction(()=>document.querySelectorAll('.tab').length===0);
 await open();await value('B2','42');assert.deepEqual(errors,[]);check('Native save and reopen preserve original style/theme metadata and recalculate cross-sheet caches');
 await app.close();app=undefined;await mkdir('out/m6',{recursive:true});
 await writeFile('out/m6/native-sheets.json',JSON.stringify({schema:1,platform:process.platform,arch:process.arch,sandbox:true,source:'Independent openpyxl report',checks,geometry:css,status:'passed'},null,2)+'\n');
}finally{if(app)await app.evaluate(({app})=>app.exit()).catch(()=>{});await rm(work,{recursive:true,force:true,maxRetries:5,retryDelay:100});}
