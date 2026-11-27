import {prepareTabbedProfile} from './tabbed-test-profile.mjs';
// SPDX-License-Identifier: MIT
/** Real sandboxed desktop: imported cross-sheet lookups, conditional aggregates and native XLSX save. */
import {_electron as electron} from 'playwright-core';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,copyFile,readFile,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {openSheetsXlsx,openOpc} from '@office-ex/ooxml';
const fixtures=resolve(process.argv[2]??'.work/m6-lookup-corpus');
const work=await mkdtemp(join(tmpdir(),'office-ex-m6-lookup-')),profile=join(work,'profile');await prepareTabbedProfile(profile);
const file=join(work,'lookup.xlsx');await copyFile(join(fixtures,'m6-lookups.xlsx'),file);
const original=await readFile(file),checks=[];let app;
const check=name=>{checks.push(name);console.log('PASS',name);};
async function wait(fn,label){const end=Date.now()+30000;while(Date.now()<end){try{if(await fn())return;}catch{}await new Promise(r=>setTimeout(r,100));}throw Error('Timed out: '+label);}
try{
 app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env:{...process.env,OFFICE_EX_PROFILE_DIR:profile},timeout:30000});
 const page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]');const errors=[];page.on('pageerror',e=>errors.push(String(e)));
 const prefs=await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences());
 assert.equal(prefs.sandbox,true);assert.equal(prefs.contextIsolation,true);assert.equal(prefs.nodeIntegration,false);
 await page.locator('#settings').click();await page.locator('#language').selectOption('en-US');await page.locator('#settings-dialog button[value=close]').click();
 const open=async()=>{await app.evaluate(({dialog},path)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[path]});},file);await page.locator('#open').click();await page.waitForSelector('#sheets-scroll');};
 const tab=name=>page.locator('.sheet-tabs [role=tab]').filter({hasText:name});
 const cell=addr=>page.locator('#sheets-scroll .sheets-cell[data-address="'+addr+'"]');
 const value=async(addr,expected)=>page.waitForFunction(({addr,expected})=>document.querySelector('#sheets-scroll .sheets-cell[data-address="'+addr+'"]')?.textContent===expected,{addr,expected});
 const showReport=async(expected)=>{await tab('Report').click();await value('B1',expected);await value('B4',expected);await value('B2','2');await value('B5','30');await value('B6','500');await value('B7','50');};
 await open();await showReport('10');
 check('Independent workbook opens editable and worker calculates conditional, vertical and horizontal lookup results');
 await tab('Data').click();await value('C1','10');
 await cell('C1').dblclick();const editor=page.locator('#sheets-cell-editor');await editor.waitFor({state:'visible'});await editor.fill('20');await editor.press('Enter');await value('C1','20');
 await showReport('20');await value('B3','30');await value('B9','40');await value('B10','30');
 await page.locator('#undo').click();await value('B1','10');await page.locator('#redo').click();await value('B1','20');await value('B4','20');
 check('Cell edit on another sheet invalidates affected formulas, undo and redo restore caches');
 await page.locator('#save').click();await wait(async()=>openSheetsXlsx(await readFile(file),'native-lookup','Report').document.sheets[1].cells['1:3'].value.value===20,'source-backed XLSX save');
 const bytes=await readFile(file),source=openSheetsXlsx(bytes,'native-lookup','Report'),cache=new Map(source.cachedValues);
 for(const [i,n]of [20,2,30,20,30,500,50,1048572,40,30,1].entries())assert.deepEqual(cache.get('sheet-1!'+(i+1)+':2'),{kind:'number',value:n});
 for(const name of ['xl/styles.xml','xl/theme/theme1.xml','docProps/core.xml','docProps/app.xml'])assert.deepEqual(openOpc(bytes).parts.get(name),openOpc(original).parts.get(name),name);
 await page.locator('.tab[aria-selected=true] .tab-close').click();await page.waitForFunction(()=>document.querySelectorAll('.tab').length===0);
 await open();await showReport('20');assert.deepEqual(errors,[]);
 check('Native save/reopen preserves formula source, all eleven caches and source metadata');
 await app.close();app=undefined;await mkdir('out/m6',{recursive:true});
 await writeFile('out/m6/native-lookups.json',JSON.stringify({schema:1,platform:process.platform,arch:process.arch,sandbox:true,source:'Independent openpyxl conditional/lookup workbook',checks,status:'passed'},null,2)+'\n');
}finally{if(app)await app.evaluate(({app})=>app.exit()).catch(()=>{});await rm(work,{recursive:true,force:true,maxRetries:5,retryDelay:100});}
