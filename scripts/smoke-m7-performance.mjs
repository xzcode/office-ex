import {prepareTabbedProfile} from './tabbed-test-profile.mjs';
// SPDX-License-Identifier: MIT
/** Native desktop: value-only recalculation transport and bounded thumbnails. */
import {_electron as electron} from 'playwright-core';
import {revealControl} from './ribbon-test-controls.mjs';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {createDocument} from '@office-ex/core';
import {writeSheetsXlsx,writeSlidesPptx} from '@office-ex/ooxml';
import {createCalculationUpdate,setCells} from '@office-ex/sheets';
const work=await mkdtemp(join(tmpdir(),'office-ex-m7-perf-')),profile=join(work,'profile');await prepareTabbedProfile(profile);
const workbook=join(work,'large.xlsx'),deck=join(work,'large.pptx');
const d=createDocument('sheets','performance-native'),s=d.sheets[0];
for(let r=1;r<=2000;r++)s.cells[r+':2']={value:{kind:'number',value:r}};
s.cells['1:1']={value:{kind:'number',value:5}};s.cells['1:3']={value:{kind:'formula',source:'=A1+1'}};
s.cells['2:3']={value:{kind:'formula',source:'=SUM(A:A)'}};
await writeFile(workbook,writeSheetsXlsx(d));
const p=createDocument('slides','performance-deck');
p.slides=Array.from({length:60},(_,i)=>({id:'slide-'+(i+1),background:{kind:'rgba',red:255,green:255,blue:255,alpha:255},shapes:[{id:'rect-'+(i+1),type:'rect',geometry:{x:914400,y:914400,width:3657600,height:1828800,rotation:0},fill:{kind:'theme',name:'accent1'}},{id:'text-'+(i+1),type:'text',geometry:{x:914400,y:3200400,width:5486400,height:914400,rotation:0},runs:[{id:'run-'+(i+1),type:'text',text:'Page '+(i+1)}]}]}));
await writeFile(deck,writeSlidesPptx(p));
let app;const checks=[],measurements={};
const check=name=>{checks.push(name);console.log('PASS',name);};
try{
 app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env:{...process.env,OFFICE_EX_PROFILE_DIR:profile},timeout:30000});
 const page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]');const errors=[];page.on('pageerror',e=>errors.push(String(e)));
 const prefs=await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences());assert.equal(prefs.sandbox,true);assert.equal(prefs.contextIsolation,true);assert.equal(prefs.nodeIntegration,false);
 await page.locator('#settings').click();await page.locator('#language').selectOption('en-US');await page.locator('#settings-dialog button[value=close]').click();
 async function open(file){await app.evaluate(({dialog},path)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[path]});},file);await page.locator('#open').click();}
 const cell=address=>page.locator('#sheets-scroll .sheets-cell[data-address="'+address+'"]');
 const expectCalc=async(kind,count,values)=>page.waitForFunction(({kind,count,values})=>{
  const root=document.querySelector('#sheets-workspace');return root?.dataset.calcTransport===kind&&Number(root.dataset.calcChangedCells)===count&&root.dataset.calculating==='false'&&Object.entries(values).every(([a,v])=>root.querySelector('[data-address="'+a+'"]')?.textContent===v);
 },{kind,count,values},{timeout:30000});
 await open(workbook);await expectCalc('full',2003,{C1:'6',C2:'5'});
 await cell('A1').dblclick();const editor=page.locator('#sheets-cell-editor');await editor.waitFor({state:'visible'});await editor.fill('9');await editor.press('Enter');
 await expectCalc('delta',1,{C1:'10',C2:'9'});
 await page.locator('#undo').click();await expectCalc('delta',1,{C1:'6',C2:'5'});
 await page.locator('#redo').click();await expectCalc('delta',1,{C1:'10',C2:'9'});
 await cell('A1').click();await (await revealControl(page,'#sheet-bold')).click();await expectCalc('delta',0,{C1:'10',C2:'9'});
 check('Real worker sends only changed scalar cells; format-only edit sends no value changes; undo/redo recalculates');
 // Opening the deck terminates the dedicated sheet worker.
 await open(deck);await page.waitForSelector('#slide-canvas');
 const rail=page.locator('.slides-thumbnails');
 await page.waitForFunction(()=>document.querySelectorAll('.slides-thumbnail .slides-scene').length>=2);
 assert.equal(await rail.locator('.slides-thumbnail').count(),60);
 const initial=await rail.evaluate(el=>({renders:Number(el.dataset.thumbnailRenders),live:el.querySelectorAll('.slides-scene').length,scroll:el.scrollTop}));
 assert.ok(initial.live<60);
 await page.evaluate(()=>{window.__thumbFirst=document.querySelector('.slides-thumbnail[data-slide-id="slide-1"]');window.__thumbSecond=document.querySelector('.slides-thumbnail[data-slide-id="slide-2"] .slides-scene');});
 await page.locator('#slide-canvas [data-shape-type=rect]').click();
 const x=page.getByLabel('X (pt)',{exact:true});await x.fill('80');await x.dispatchEvent('change');
 await page.waitForFunction(()=>document.querySelector('#slide-canvas [data-shape-type=rect]')?.style.left!=='96px');
 const after=await rail.evaluate(el=>({renders:Number(el.dataset.thumbnailRenders),live:el.querySelectorAll('.slides-scene').length,scroll:el.scrollTop}));
 assert.equal(after.renders,initial.renders+1);
 assert.equal(await page.evaluate(()=>window.__thumbFirst===document.querySelector('.slides-thumbnail[data-slide-id="slide-1"]')&&window.__thumbSecond===document.querySelector('.slides-thumbnail[data-slide-id="slide-2"] .slides-scene')),true);
 assert.equal(after.scroll,initial.scroll);measurements.initialThumbnails=initial.live;measurements.redrawnAfterOneEdit=after.renders-initial.renders;
 check('One slide edit redraws exactly one visible preview, retains unchanged thumbnail nodes and rail scroll');
 await rail.evaluate(el=>{el.scrollTop=el.scrollHeight;});
 await page.waitForFunction(()=>document.querySelector('.slides-thumbnail[data-slide-id="slide-60"] .slides-scene')&&!document.querySelector('.slides-thumbnail[data-slide-id="slide-1"] .slides-scene'));
 const tail=await rail.evaluate(el=>el.querySelectorAll('.slides-scene').length);assert.ok(tail<60);measurements.tailThumbnails=tail;
 await rail.locator('.slides-thumbnail[data-slide-id="slide-60"]').click();await page.waitForFunction(()=>document.querySelector('#slide-canvas')?.dataset.slideId==='slide-60');
 assert.deepEqual(errors,[]);
 check('Scrolling evicts offscreen preview DOM; last slide renders on demand and is selectable');
 await app.close();app=undefined;await mkdir('out/m7',{recursive:true});await writeFile('out/m7/native-performance.json',JSON.stringify({schema:1,platform:process.platform,arch:process.arch,sandbox:true,checks,measurements,status:'passed'},null,2)+'\n');
}finally{if(app)await app.evaluate(({app})=>app.exit()).catch(()=>{});await rm(work,{recursive:true,force:true,maxRetries:5,retryDelay:100});}
