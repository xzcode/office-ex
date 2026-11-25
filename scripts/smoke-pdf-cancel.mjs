// SPDX-License-Identifier: MIT
// Focused real desktop regression for PDF.js cancellation and rapid page changes.
import { _electron as electron } from 'playwright-core';
import { mkdtemp, mkdir, readFile, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import assert from 'node:assert/strict';
const fixture=resolve(process.argv[2]??'out/m4-native/slides-native.pdf');
await readFile(fixture);const work=await mkdtemp(join(tmpdir(),'office-ex-pdf-cancel-')),output=resolve('out/pdf-cancel');await mkdir(output,{recursive:true});
const errors=[],evidence={scope:'Real sandboxed Electron PDF tab; rapid page changes, cancelled text reads and search',checks:[],errors};let app;
try{
  const env={...process.env,OFFICE_EX_PROFILE_DIR:work};delete env.ELECTRON_RUN_AS_NODE;
  app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env,timeout:30000});
  const page=await app.firstWindow();page.on('pageerror',e=>errors.push(e.stack??e.message));await page.waitForSelector('body[data-ready=true]');
  await app.evaluate(({dialog},filePath)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[filePath]});},fixture);await page.locator('#open').click();await page.waitForSelector('#pdf-workspace[data-loaded=true]',{timeout:30000});
  assert.equal(await page.locator('#pdf-workspace').getAttribute('data-pages'),'3');
  for(let i=0;i<5;i++){
    await page.evaluate(()=>{const n=document.querySelector('#pdf-next-page'),p=document.querySelector('#pdf-previous-page');n.click();p.click();n.click();p.click();});
    await page.waitForSelector('#pdf-workspace[data-loaded=true]');
  }
  evidence.checks.push('Repeated rapid navigation settles with live local text layer');
  await page.locator('#pdf-search').fill('Native presentation');await page.locator('#pdf-find').click();await page.waitForSelector('.pdf-match-current');
  await page.locator('#pdf-next-page').click();await page.waitForFunction(()=>document.querySelector('#pdf-workspace')?.dataset.page==='2'&&document.querySelector('#pdf-workspace')?.dataset.loaded==='true');
  evidence.checks.push('Search and subsequent navigation remain functional');
  await page.locator('.tab[aria-selected=true] .tab-close').click();await page.waitForSelector('#pdf-workspace',{state:'detached'});await page.waitForTimeout(600);assert.deepEqual(errors,[]);evidence.checks.push('Tab close leaves no renderer stream exceptions');
  console.log('PDF cancellation regression passed: '+evidence.checks.length+' checks, 0 renderer errors');
}catch(e){evidence.failure=String(e);throw e;}
finally{
  try{if(app){if(evidence.failure)await app.evaluate(({app})=>app.exit(1)).catch(()=>{});else{let timer;try{await Promise.race([app.close(),new Promise((_,reject)=>{timer=setTimeout(()=>reject(new Error('Orderly desktop quit exceeded 20 seconds')),20000);})]);evidence.orderlyQuit=true;}finally{clearTimeout(timer);}}}}
  catch(error){evidence.failure??=String(error);await app?.evaluate(({app})=>app.exit(1)).catch(()=>{});throw error;}
  finally{await writeFile(join(output,'result.json'),JSON.stringify(evidence,null,2)+'\n');await rm(work,{recursive:true,force:true});}
}
