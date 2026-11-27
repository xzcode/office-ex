import {prepareTabbedProfile} from './tabbed-test-profile.mjs';
// SPDX-License-Identifier: MIT
import {_electron as electron} from 'playwright-core';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,rm,readFile,writeFile} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {openWriterDocx,openSheetsXlsx,openSlidesPptx} from '@office-ex/ooxml';
import {revealControl} from './ribbon-test-controls.mjs';
const work=await mkdtemp(join(tmpdir(),'office-ex-home-')),profile=join(work,'profile'),out=resolve(process.env.OFFICE_EX_HOME_OUTPUT??'out/m5-home');
await prepareTabbedProfile(profile);await mkdir(out,{recursive:true});
const report={platform:process.platform,arch:process.arch,checks:[],screenshots:[],startedAt:new Date().toISOString()};let app;
const check=name=>{report.checks.push(name);console.log('PASS',name);};
try{
 app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env:{...process.env,OFFICE_EX_PROFILE_DIR:profile},timeout:30000});
 const page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]');
 await app.evaluate(({BrowserWindow})=>{BrowserWindow.getAllWindows()[0].setSize(1440,900);});
 const snap=async name=>{await page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));await page.screenshot({path:join(out,name+'.png')});report.screenshots.push(name+'.png');};
 await page.locator('#settings').click();await page.locator('#language').selectOption('zh-CN');await page.waitForFunction(()=>document.documentElement.lang==='zh-CN');await page.locator('#theme').selectOption('light');await page.locator('#settings-dialog button[value=close]').click();
 const assertHeader=async()=>{
  const layout=await page.evaluate(()=>{const b=document.querySelector('#office-titlebar'),r=b.getBoundingClientRect(),home=document.querySelector('#home');return {height:r.height,top:r.top,brandCount:document.querySelectorAll('#home').length,homeLeft:home.getBoundingClientRect().left,viewport:innerWidth,barWidth:r.width};});
  assert.equal(layout.height,50);assert.equal(layout.top,0);assert.equal(layout.brandCount,1);assert.equal(layout.viewport,layout.barWidth);
 };
 await assertHeader();assert.equal(await page.locator('.home-product').count(),3);
 assert.equal(await page.locator('.home-other-formats').evaluate(e=>e.open),false);
 assert.equal(await page.locator('.toolbar-actions').isVisible(),false);
 assert.equal(await page.locator('.home-product svg').count(),6);
 assert.equal(await page.locator('#recent-list .home-recent-row').count(),0);
 assert.equal(await page.locator('#recent-list .home-empty').isVisible(),true);
 check('Single compact titlebar, original three-module marks, clean empty-state home');
 await page.locator('#new-menu').click();assert.equal(await page.locator('#new-doc').isVisible(),true);assert.equal(await page.locator('#new-xls').isVisible(),true);assert.equal(await page.locator('#new-ppt').isVisible(),true);
 await page.locator('#recovery-nav').click();assert.equal(await page.locator('.home-empty').isVisible(),true);
 await page.locator('#recent-nav').click();assert.equal(await page.locator('#recent-list').isVisible(),true);
 await page.locator('#workspace-nav').click();assert.equal(await page.locator('#new-docx').isVisible(),true);
 check('Home navigation, older formats, draft recovery and recent views work');
 const files=[['docx','产品方案.docx'],['xlsx','成本统计表.xlsx'],['pptx','产品演示.pptx']];
 for(const [format,name] of files){
  await page.locator('#new-'+format).click();await page.waitForSelector('.office-editor-shell');await assertHeader();assert.equal(await page.locator('.sidebar').isVisible(),false);assert.equal(await page.locator('#save').isVisible(),true);
  if(format==='docx'){const input=page.locator('#writer-page-input');await input.waitFor();await input.focus();await page.keyboard.insertText('项目产品方案');}
  else if(format==='xlsx'){await page.locator('#sheets-formula').fill('123');await page.locator('#sheets-formula').press('Enter');}
  else {await (await revealControl(page,'#slide-insert-text')).click();await page.locator('#slide-text-editor').fill('产品演示');}
  const file=join(work,name);
  await app.evaluate(({dialog},path)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:path});dialog.showMessageBox=async()=>({response:1,checkboxChecked:false});},file);
  await page.locator('#save').click();
  try{await page.waitForFunction(name=>document.querySelector('.tab[aria-selected=true] .tab-title')?.textContent===name,name,{timeout:12000});}
  catch(e){console.log(await page.evaluate(()=>({notice:document.querySelector('#message')?.textContent,status:document.querySelector('#document-status')?.textContent,active:document.activeElement?.id,save:document.querySelector('#save')?.disabled})));throw e;}
  const bytes=await readFile(file);
  if(format==='docx')assert.ok(openWriterDocx(bytes,'home-check').document.sections[0].blocks[0].runs.some(r=>r.text==='项目产品方案'));
  else if(format==='xlsx')assert.equal(openSheetsXlsx(bytes,'home-check').document.sheets[0].cells['1:1'].value.value,123);
  else assert.ok(openSlidesPptx(bytes,'home-check').document.slides[0].shapes.some(s=>s.type==='text'&&s.runs.some(r=>r.text==='产品演示')));
  await page.locator('.tab[aria-selected=true] .tab-close').click();await page.waitForSelector('.home-workspace');console.log('Saved',format);
 }
 check('All three new-file cards and merged header save actual edited native documents, including PPT text blur');
 assert.equal(await page.locator('#recent-list .home-recent-row').count(),3);
 assert.deepEqual(await page.locator('.home-recent-name').allTextContents(),['产品演示.pptx','成本统计表.xlsx','产品方案.docx']);
 assert.ok((await page.locator('.home-recent-date').allTextContents()).every(s=>s.length>0));
 await snap('home-light');
 await page.locator('#recent-list .home-recent-row').first().click();await page.waitForSelector('#slides-workspace');assert.equal(await page.locator('.tab[aria-selected=true] .tab-title').textContent(),'产品演示.pptx');
 await page.locator('#home').click();await page.waitForSelector('.home-workspace');assert.equal(await page.locator('.tab').count(),1);assert.equal(await page.locator('.toolbar-actions').isVisible(),false);
 await page.locator('.tab').first().click();await page.waitForSelector('#slides-workspace');await page.locator('.tab-close').click();await page.waitForSelector('.home-workspace');
 check('Recent files reflect real local saves and open by opaque capability; open tabs remain reachable from home');
 await page.locator('#settings').click();await page.locator('#theme').selectOption('dark');await page.locator('#settings-dialog button[value=close]').click();await snap('home-dark');
 // Keep the physical macOS titlebar reservation separate from OS behavior.
 // This simulated CSS check is not a native macOS window test.
 if(process.platform==='darwin'){
  const native=await app.evaluate(({BrowserWindow})=>{const w=BrowserWindow.getAllWindows()[0];return {buttons:w.getWindowButtonPosition(),window:w.getBounds(),content:w.getContentBounds()};});
  assert.ok(native.buttons&&native.buttons.x>=8&&native.buttons.x<=26&&native.buttons.y>=8&&native.buttons.y<=28);
  assert.ok(Math.abs(native.window.height-native.content.height)<=2);
  report.nativeTitlebar=native;
 }else{
  const geometry=await page.evaluate(()=>{const root=document.documentElement,platform=root.dataset.platform;root.dataset.platform='darwin';const bar=document.querySelector('#office-titlebar'),brand=document.querySelector('#home');const result={left:brand.getBoundingClientRect().left,padding:parseFloat(getComputedStyle(bar).paddingLeft),region:getComputedStyle(brand).getPropertyValue('-webkit-app-region')||getComputedStyle(brand).getPropertyValue('app-region')};root.dataset.platform=platform;return result;});
  assert.ok(geometry.left>=88&&geometry.padding>=88);assert.equal(geometry.region,'no-drag');report.macCssReservation=geometry;
 }
 check('Light/dark home and platform-specific titlebar safe-area geometry');
 report.status='passed';report.completedAt=new Date().toISOString();await writeFile(join(out,'result.json'),JSON.stringify(report,null,2)+'\n');
 await app.close();app=undefined;
}catch(error){report.status='failed';report.error=String(error);await writeFile(join(out,'result.json'),JSON.stringify(report,null,2)+'\n');if(app){const p=await app.firstWindow().catch(()=>null);if(p)await p.screenshot({path:join(out,'failure.png')}).catch(()=>{});}throw error;}
finally{if(app)await app.evaluate(({app})=>app.exit()).catch(()=>{});await rm(work,{recursive:true,force:true,maxRetries:5,retryDelay:100});}
