// SPDX-License-Identifier: MIT
/** Focused native regression using independently generated python-docx files.
 * Generate via tests/office_external_compat.py; no production test backdoor. */
import {revealControl} from './ribbon-test-controls.mjs';
import {_electron as electron} from 'playwright-core';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,copyFile,readFile,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {openWriterDocx,openOpc} from '@office-ex/ooxml';
const fixtures=resolve(process.argv[2]??'.work/m6-external-final');
const work=await mkdtemp(join(tmpdir(),'office-ex-m6-native-')),profile=join(work,'profile');
await mkdir(profile);
const file=join(work,'ordinary-table.docx');await copyFile(join(fixtures,'ordinary-table.docx'),file);
const before=openOpc(await readFile(file)),checks=[];
let app;
function check(title){checks.push(title);console.log('PASS',title);}
try{
 app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env:{...process.env,OFFICE_EX_PROFILE_DIR:profile},timeout:30000});
 const page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]');
 const prefs=await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences());assert.equal(prefs.sandbox,true);assert.equal(prefs.contextIsolation,true);assert.equal(prefs.nodeIntegration,false);
 const errors=[];page.on('pageerror',e=>errors.push(String(e)));
 await page.locator('#settings').click();await page.locator('#language').selectOption('zh-CN');await page.locator('#settings-dialog button[value=close]').click();
 const open=async path=>{await app.evaluate(({dialog},filePath)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[filePath]});},path);await page.locator('#open').click();};
 await open(file);await revealControl(page,'.writer-edit-table');await page.waitForSelector('.writer-edit-table');
 const cell=(r,c)=>page.locator('.writer-cell-input[data-row="'+r+'"][data-column="'+c+'"]').first();
 assert.equal(await cell(1,0).inputValue(),'办公软件');check('Ordinary external TableGrid DOCX opens in actual editable Writer');
 await cell(1,0).fill('原生编辑测试');await cell(1,0).click();
 await (await revealControl(page,'#writer-table-add-row')).click();
 await page.waitForFunction(()=>document.querySelectorAll('.writer-edit-table tbody tr').length===3);
 await cell(2,0).fill('新增行');await cell(2,0).press('Tab');
 await page.locator('#save').click();
 await page.waitForFunction(()=>document.querySelector('#document-status')?.textContent.includes('已保存'),null,{timeout:30000});
 const bytes=await readFile(file),model=openWriterDocx(bytes,'m6-native','Table').document;
 const table=model.sections[0].blocks.find(b=>b.type==='table');
 assert.equal(table.rows.length,3);assert.equal(table.rows[1].cells[0].blocks[0].runs[0].text,'原生编辑测试');assert.equal(table.rows[2].cells[0].blocks[0].runs[0].text,'新增行');
 assert.equal(table.nativeTable.styleId,'TableGrid');assert.equal(table.nativeTable.autoLayout,true);assert.equal(table.nativeTable.autoWidth,true);
 const after=openOpc(bytes);
 for(const [name,data]of before.parts)if(name!==before.mainPart)assert.deepEqual(after.parts.get(name),data,name);
 check('Real desktop edit and row insertion save native DOCX with style and all ancillary parts intact');
 await page.locator('.tab[aria-selected=true] .tab-close').click();await page.waitForFunction(()=>document.querySelectorAll('.tab').length===0);
 await open(file);await revealControl(page,'.writer-edit-table');await page.waitForSelector('.writer-edit-table');assert.equal(await cell(1,0).inputValue(),'原生编辑测试');
 assert.deepEqual(errors,[]);check('Saved DOCX reopens editable without renderer errors');
 await app.close();app=undefined;
 await mkdir('out/m6',{recursive:true});await writeFile('out/m6/native-word-table.json',JSON.stringify({schema:1,platform:process.platform,arch:process.arch,source:'Independent python-docx default template',checks,sandbox:true,status:'passed'},null,2)+'\n');
}finally{if(app)await app.evaluate(({app})=>app.exit()).catch(()=>{});await rm(work,{recursive:true,force:true,maxRetries:5,retryDelay:100});}
