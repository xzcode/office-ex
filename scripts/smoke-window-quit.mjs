// SPDX-License-Identifier: MIT
// Real Electron quit handshake; native confirmation dialog choices are stubbed.
import { _electron as electron } from 'playwright-core';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
const work=await mkdtemp(join(tmpdir(),'office-ex-quit-')),output=resolve('out/window-quit');await mkdir(output,{recursive:true});
const evidence={platform:process.platform,checks:[],startedAt:new Date().toISOString()};let app;
const bounded=async(promise,label)=>{let timer;try{return await Promise.race([promise,new Promise((_,reject)=>{timer=setTimeout(()=>reject(new Error(label)),20000);})]);}finally{clearTimeout(timer);}};
try{
 const env={...process.env,OFFICE_EX_PROFILE_DIR:work};delete env.ELECTRON_RUN_AS_NODE;
 app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env,timeout:30000});
 const page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]');
 await page.locator('[data-kind=slides]').click();await page.locator('#slide-insert-rect').click();
 await page.waitForSelector('.tab[aria-selected=true]');assert.match(await page.locator('.tab[aria-selected=true] .tab-title').textContent(),/^• /);
 await app.evaluate(({dialog,app})=>{globalThis.__quitDialogCount=0;dialog.showMessageBox=async()=>{globalThis.__quitDialogCount++;return {response:2,checkboxChecked:false};};app.quit();});
 const start=Date.now();while(await app.evaluate(()=>globalThis.__quitDialogCount)!==1){if(Date.now()-start>10000)throw new Error('Quit did not request dirty-document confirmation');await new Promise(r=>setTimeout(r,50));}
 await page.waitForSelector('.tab[aria-selected=true]');assert.match(await page.locator('.tab[aria-selected=true] .tab-title').textContent(),/^• /);evidence.checks.push('Cancelling explicit Quit preserves dirty document and running application');
 await app.evaluate(({dialog})=>{dialog.showMessageBox=async()=>({response:1,checkboxChecked:false});});
 await bounded(app.close(),'Confirmed orderly Quit did not exit');evidence.checks.push('Confirmed discard during explicit Quit releases sessions and terminates process');evidence.completedAt=new Date().toISOString();console.log('Window quit regression passed: 2/2');
}catch(error){evidence.failure=String(error);await app?.evaluate(({app})=>app.exit(1)).catch(()=>{});throw error;}
finally{await writeFile(join(output,'result.json'),JSON.stringify(evidence,null,2)+'\n');await rm(work,{recursive:true,force:true});}
