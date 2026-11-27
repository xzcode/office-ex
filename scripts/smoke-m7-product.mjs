// SPDX-License-Identifier: MIT
/** Unseeded profile: actual independent windows, OS path delivery, chrome
 * selection policy and pointer-driven panel resizing. */
import {_electron as electron} from 'playwright-core';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,readFile,writeFile,rm} from 'node:fs/promises';
import {spawn} from 'node:child_process';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {createDocument} from '@office-ex/core';
import {writeSheetsXlsx,writeSlidesPptx,openSheetsXlsx} from '@office-ex/ooxml';
import {revealControl} from './ribbon-test-controls.mjs';
const work=await mkdtemp(join(tmpdir(),'office-ex-product-')),profile=join(work,'profile');await mkdir(profile);
const xlsx=join(work,'预算表 Test.xlsx'),pptx=join(work,'演示 Test.pptx'),second=join(work,'第二份演示.pptx');
const w=createDocument('sheets','native-product');w.sheets[0].cells['1:1']={value:{kind:'number',value:5}};
await writeFile(xlsx,writeSheetsXlsx(w));
const deck=createDocument('slides','native-product-slides');deck.slides[0].shapes=[{id:'box',type:'rect',geometry:{x:914400,y:914400,width:1828800,height:914400,rotation:0}}];
const ppt=writeSlidesPptx(deck);await writeFile(pptx,ppt);await writeFile(second,ppt);
const checks=[],measurements={};let app;
const check=s=>{checks.push(s);console.log('PASS',s);};
async function wait(fn,label,timeout=30000){const deadline=Date.now()+timeout;let last;while(Date.now()<deadline){try{const value=await fn();if(value)return value;}catch(error){last=String(error);}await new Promise(r=>setTimeout(r,100));}throw Error('Timed out '+label+(last?': '+last:''));}
async function findWindow(selector,title){
 return wait(async()=>{for(const p of app.windows()){if(p.isClosed())continue;if(await p.locator(selector).count()&&(!title||await p.title()===title)&&await p.locator('body[data-ready=true]').count())return p;}},selector+' '+(title??''));
}
async function openFrom(page,path){
 await app.evaluate(({dialog},file)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[file]});},path);await page.locator('#open').click();
}
async function closeDocument(page){await page.locator('.tab-close').click();await wait(()=>page.isClosed(),'document window close');}
try{
 const args=[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop'),xlsx];
 const env={...process.env,OFFICE_EX_PROFILE_DIR:profile};delete env.ELECTRON_RUN_AS_NODE;
 app=await electron.launch({chromiumSandbox:true,args,env,timeout:30000});
 const home=await findWindow('.home-workspace','Office EX'),sheet=await findWindow('#sheets-scroll','预算表 Test.xlsx — Office EX');
 assert.notEqual(home,sheet);assert.equal(await sheet.locator('.tab').count(),1);assert.equal(await home.locator('.tab').count(),0);
 const prefs=await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows().filter(w=>w.webContents.getURL().startsWith('office-ex://app/index.html')).map(w=>{const p=w.webContents.getLastWebPreferences();return {sandbox:p.sandbox,contextIsolation:p.contextIsolation,nodeIntegration:p.nodeIntegration,webSecurity:p.webSecurity};}));
 assert.equal(prefs.length,2);assert.ok(prefs.every(p=>p.sandbox&&p.contextIsolation&&!p.nodeIntegration&&p.webSecurity));
 await sheet.locator('#settings').click();assert.equal(await sheet.locator('#document-window-mode').inputValue(),'windows');await sheet.locator('#settings-dialog button[value=close]').click();
 await openFrom(home,xlsx);await new Promise(r=>setTimeout(r,500));assert.equal(app.windows().filter(p=>!p.isClosed()).length,2);
 check('Cold OS document argument opens a separate sandboxed window; duplicate opens reuse it');
 // Real second process passes a file to the running single-instance app.
 const exe=app.process().spawnfile;
 const child=spawn(exe,[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop'),pptx],{env,stdio:'ignore'});
 await new Promise((resolve,reject)=>{child.once('error',reject);child.once('exit',code=>code===0?resolve():reject(Error('Secondary process '+code)));});
 const slides=await findWindow('#slide-canvas','演示 Test.pptx — Office EX');
 assert.equal(app.windows().filter(p=>!p.isClosed()).length,3);
 check('Second-instance command-line file is delivered to another independent window');
 // Native titlebar hit-region geometry and browser text selection behavior.
 const geometry=await sheet.evaluate(()=>{
  const space=document.querySelector('.titlebar-drag-space'),tabs=document.querySelector('#tablist'),button=document.querySelector('#open');
  const region=e=>getComputedStyle(e).getPropertyValue('-webkit-app-region')||getComputedStyle(e).getPropertyValue('app-region');
  return {spacer:space.getBoundingClientRect().width,drag:region(space),tabs:region(tabs),button:region(button),select:getComputedStyle(document.querySelector('.office-ribbon-tab')).userSelect,input:getComputedStyle(document.querySelector('#sheets-formula')).userSelect};
 });
 assert.ok(geometry.spacer>100,JSON.stringify(geometry));assert.equal(geometry.drag,'drag');assert.equal(geometry.tabs,'no-drag');assert.equal(geometry.button,'no-drag');assert.equal(geometry.select,'none');assert.equal(geometry.input,'text');
 const tab=sheet.locator('.office-ribbon-tab').first(),b=await tab.boundingBox();await sheet.evaluate(()=>getSelection()?.removeAllRanges());await sheet.mouse.move(b.x+3,b.y+b.height/2);await sheet.mouse.down();await sheet.mouse.move(b.x+b.width-3,b.y+b.height/2,{steps:8});await sheet.mouse.up();assert.equal(await sheet.evaluate(()=>getSelection()?.toString()),'');
 const formula=sheet.locator('#sheets-formula');await formula.fill('42');await formula.press('ControlOrMeta+a');assert.deepEqual(await formula.evaluate(e=>[e.selectionStart,e.selectionEnd]),[0,2]);await formula.press('Enter');
 await sheet.waitForFunction(()=>document.querySelector('[data-address=A1]')?.textContent==='42');
 await sheet.locator('#save').click();await wait(async()=>openSheetsXlsx(await readFile(xlsx),'native').document.sheets[0].cells['1:1'].value.value===42,'worksheet save');
 measurements.titlebarDragWidth=geometry.spacer;
 check('UI labels cannot be selected; titlebar spacer is draggable; input text and real editing still work');
 async function width(page,selector){return (await page.locator(selector).boundingBox()).width;}
 async function drag(page,selector,dx){const box=await page.locator(selector).boundingBox();await page.mouse.move(box.x+box.width/2,box.y+Math.min(90,box.height/2));await page.mouse.down();await page.mouse.move(box.x+box.width/2+dx,box.y+Math.min(90,box.height/2),{steps:10});await page.mouse.up();}
 const left=await width(slides,'.slides-thumbnails'),right=await width(slides,'.slides-inspector');
 await drag(slides,'[data-splitter=slide-rail]',50);await drag(slides,'[data-splitter=slide-inspector]',-60);
 const left2=await width(slides,'.slides-thumbnails'),right2=await width(slides,'.slides-inspector');
 assert.ok(left2>=left+45&&right2>=right+55,JSON.stringify({left,left2,right,right2}));
 measurements.panels={leftBefore:left,leftAfter:left2,rightBefore:right,rightAfter:right2};
 await openFrom(home,second);const other=await findWindow('#slide-canvas','第二份演示.pptx — Office EX');
 assert.ok(Math.abs(await width(other,'.slides-thumbnails')-left2)<2);assert.ok(Math.abs(await width(other,'.slides-inspector')-right2)<2);
 await closeDocument(other);check('Pointer resize changes both panels and the next document inherits their widths');
 await closeDocument(sheet);assert.ok(!home.isClosed()&&!slides.isClosed());
 await home.locator('#new-docx').click();const writer=await findWindow('#writer-native-page');assert.notEqual(writer,home);
 const text=writer.locator('#writer-page-input');await text.focus();await writer.keyboard.insertText('Independent document');
 await app.evaluate(({dialog})=>{dialog.showMessageBox=async()=>({response:2,checkboxChecked:false});});
 await writer.locator('.tab-close').click();await new Promise(r=>setTimeout(r,300));assert.ok(!writer.isClosed());
 await app.evaluate(({dialog})=>{dialog.showMessageBox=async()=>({response:1,checkboxChecked:false});});await closeDocument(writer);
 assert.ok(!slides.isClosed()&&!home.isClosed());await closeDocument(slides);
 check('New documents use independent windows; cancelling close keeps dirty content; closing one window leaves the others intact');
 await home.locator('#settings').click();await home.locator('#document-window-mode').selectOption('tabs');await home.locator('#settings-dialog button[value=close]').click();
 await home.locator('#new-xlsx').click();await home.waitForSelector('#sheets-scroll');
 assert.equal(app.windows().filter(p=>!p.isClosed()).length,1);
 check('Explicit tabbed preference remains available');
 await app.close();app=undefined;
 await mkdir('out/m7',{recursive:true});await writeFile('out/m7/native-product.json',JSON.stringify({schema:1,platform:process.platform,arch:process.arch,checks,measurements,status:'passed'},null,2)+'\n');
}finally{if(app)await app.evaluate(({app})=>app.exit()).catch(()=>{});await rm(work,{recursive:true,force:true,maxRetries:5,retryDelay:100});}
