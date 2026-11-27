import {prepareTabbedProfile} from './tabbed-test-profile.mjs';
// SPDX-License-Identifier: MIT
import {_electron as electron} from 'playwright-core';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,rm,writeFile,readFile} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {openWriterDocx,openSheetsXlsx,openSlidesPptx} from '@office-ex/ooxml';
import {revealControl} from './ribbon-test-controls.mjs';
const work=await mkdtemp(join(tmpdir(),'office-ex-ribbon-')),profile=join(work,'profile'),out=resolve(process.env.OFFICE_EX_RIBBON_OUTPUT??'out/m5-ribbon');
await prepareTabbedProfile(profile);await mkdir(out,{recursive:true});
const report={platform:process.platform,arch:process.arch,checks:[],screenshots:[],startedAt:new Date().toISOString()};let app;
const check=name=>{report.checks.push(name);console.log('PASS',name);};
try{
 app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env:{...process.env,OFFICE_EX_PROFILE_DIR:profile},timeout:30000});
 const page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]');
 await app.evaluate(({BrowserWindow})=>{BrowserWindow.getAllWindows()[0].setSize(1440,900);});
 const prefs=await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences());assert.equal(prefs.sandbox,true);assert.equal(prefs.contextIsolation,true);assert.equal(prefs.nodeIntegration,false);
 await page.locator('#settings').click();await page.locator('#language').selectOption('zh-CN');await page.waitForFunction(()=>document.documentElement.lang==='zh-CN');await page.locator('#theme').selectOption('light');await page.locator('#settings-dialog button[value=close]').click();
 const openNew=async(format)=>{await page.locator('#home').click();await page.locator('#new-'+format).click();};
 const nav=async(tab)=>{await page.locator('.office-ribbon-tab[data-ribbon-tab="'+tab+'"]').click();assert.equal(await page.locator('.office-ribbon-panel:not([hidden])').count(),1);};
 const snap=async(name)=>{await page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));const filename=join(out,name+'.png');await page.screenshot({path:filename});report.screenshots.push(name+'.png');};
 const save=async(format)=>{const file=join(work,'ribbon.'+format);await app.evaluate(({dialog},filePath)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath});dialog.showMessageBox=async()=>({response:1,checkboxChecked:false});},file);await page.locator('#save').click();await page.waitForFunction(()=>document.querySelector('#document-status')?.textContent.includes('已保存'));return file;};
 const typeNative=async(text)=>{const input=page.locator('#writer-page-input');await input.waitFor({state:'visible'});await input.focus();await page.keyboard.insertText(text);};
 const checkShell=async(kind)=>{
  assert.equal(await page.locator('.office-editor-shell').count(),1);assert.equal(await page.locator('.sidebar').isVisible(),false);
  assert.equal(await page.locator('.office-ribbon-panel:not([hidden])').count(),1);
  assert.equal(await page.locator('.ribbon-unclassified').count(),0);
  const result=await page.evaluate(()=>{const r=document.querySelector('.office-ribbon').getBoundingClientRect(),body=document.querySelector('#content').getBoundingClientRect();return {ribbon:r.height,body:body.height};});assert.ok(result.ribbon<160&&result.body>400,JSON.stringify(result));
  const ids=await page.evaluate(()=>{const ids=[...document.querySelectorAll('[id]')].map(e=>e.id);return ids.filter((id,i)=>ids.indexOf(id)!==i);});assert.deepEqual(ids,[]);console.log(kind,'ribbon',result.ribbon,'content',result.body);
 };
 await openNew('docx');await page.waitForSelector('#writer-native-page:not([hidden])');await checkShell('Writer');
 await typeNative('Office EX 产品方案');
 await page.keyboard.press('ControlOrMeta+a');await page.locator('#writer-bold').click();await page.locator('#writer-font-size').fill('24');await page.locator('#writer-font-size').press('Tab');
 await page.locator('#writer-page-input').focus();await page.keyboard.press('End');await page.keyboard.press('Enter');await typeNative('这是直接在页面中编辑的正文。');
 // Direct page-mode commands remain accessible in their actual ribbons.
 await nav('layout');await page.locator('#writer-spaceAfter').fill('8');await page.locator('#writer-spaceAfter').press('Tab');
 await nav('home');await snap('writer');
 const docx=await save('docx'),word=openWriterDocx(await readFile(docx),'ribbon','Writer').document;
 assert.ok(word.sections[0].blocks.some(p=>p.type==='paragraph'&&p.runs.some(r=>r.type==='text'&&r.text.includes('产品方案')&&r.style?.bold===true&&r.style.fontSize===304800)));
 assert.ok(word.sections[0].blocks.some(p=>p.type==='paragraph'&&p.style?.spaceAfter===101600));
 check('Writer default direct paper editing, formatting and paragraph layout commands save to DOCX');
 await nav('view');await page.locator('#writer-exit-page-edit').click();await page.waitForSelector('#draft-text');await revealControl(page,'#writer-page-edit');await page.locator('#writer-page-edit').click();await page.waitForSelector('#writer-native-page:not([hidden])');check('Writer alternate text mode and page view roundtrip');
 await openNew('xlsx');await page.waitForSelector('#sheets-scroll');await checkShell('Sheets');
 const enter=async(address,text)=>{await page.locator('#sheets-address').fill(address);await page.locator('#sheets-address').press('Enter');await page.locator('#sheets-formula').fill(text);await page.locator('#sheets-formula').press('Enter');};
 for(const [address,value] of [['A1','项目'],['B1','预算'],['C1','实际'],['A2','研发'],['B2','320'],['C2','298'],['D2','=B2-C2'],['A3','市场'],['B3','180'],['C3','195'],['D3','=B3-C3']])await enter(address,value);
 await nav('cells');await page.locator('#sheet-column-width').fill('140');await page.locator('#sheet-column-width').press('Tab');await nav('home');await snap('sheets');
 const sheetFile=await save('xlsx'),sheet=openSheetsXlsx(await readFile(sheetFile),'ribbon').document.sheets[0];assert.equal(sheet.cells['2:4'].value.source,'=B2-C2');
 check('Sheets native formula bar, cell management ribbon and workbook save');
 await openNew('pptx');await page.waitForSelector('#slides-workspace');await checkShell('Slides');await nav('insert');await page.locator('#slide-insert-rect').click();
 await nav('home');await page.locator('#slide-insert-text').click();await page.locator('#slide-text-editor').fill('Office EX 演示文稿');await page.locator('#slide-text-editor').press('Tab');
 await nav('arrange');assert.equal(await page.locator('#slide-group').count(),1);await nav('view');await page.locator('#slide-properties-toggle').click();assert.equal(await page.locator('.slides-inspector').isVisible(),false);await page.locator('#slide-properties-toggle').click();assert.equal(await page.locator('.slides-inspector').isVisible(),true);
 await nav('home');await snap('slides');const pptx=await save('pptx'),deck=openSlidesPptx(await readFile(pptx),'ribbon').document;assert.equal(deck.slides[0].shapes.length,2);assert.ok(deck.slides[0].shapes.some(s=>s.type==='text'&&s.runs.some(r=>r.text.includes('演示文稿'))));check('Slides insert, arrange, contextual inspector and real PPTX save');
 report.status='passed';report.completedAt=new Date().toISOString();await writeFile(join(out,'result.json'),JSON.stringify(report,null,2)+'\n');
 await app.close();app=undefined;
}catch(e){report.status='failed';report.error=String(e);await writeFile(join(out,'result.json'),JSON.stringify(report,null,2)+'\n');if(app){const p=await app.firstWindow().catch(()=>null);if(p)await p.screenshot({path:join(out,'failure.png')}).catch(()=>{});}throw e;}
finally{if(app)await app.evaluate(({app})=>app.exit()).catch(()=>{});await rm(work,{recursive:true,force:true,maxRetries:5,retryDelay:100});}
