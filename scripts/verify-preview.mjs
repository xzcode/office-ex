// SPDX-License-Identifier: MIT
/** Verify and launch a locally generated trusted development-preview archive.
 * No arbitrary downloaded archives, no disabled sandbox, no actual user data. */
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {createReadStream} from 'node:fs';
import {mkdir,mkdtemp,readFile,readdir,lstat,readlink,rm,writeFile} from 'node:fs/promises';
import {basename,dirname,isAbsolute,join,relative,resolve,sep} from 'node:path';
import {tmpdir} from 'node:os';
import {execFileSync} from 'node:child_process';
import {_electron as electron} from 'playwright-core';
import {inspectForEditor} from '@office-ex/format-service';

const manifestPath=resolve(process.argv[2]??'');
assert(process.argv[2]?.endsWith('.manifest.json'),'Pass the generated manifest path');
const manifest=JSON.parse(await readFile(manifestPath,'utf8')),prefix=manifestPath.slice(0,-'.manifest.json'.length);
assert(manifest.schema===1&&manifest.channel==='development-preview'&&manifest.formalRelease===false,'Not a generated preview');
assert(manifest.platform===process.platform&&manifest.arch===process.arch,'Wrong native platform');
assert(/^Office-EX-\d+\.\d+\.\d+-dev\.\d+-[a-z0-9]+-[a-z0-9]+$/.test(basename(prefix)),'Invalid preview name');
assert(manifest.appFolder===(process.platform==='darwin'?'Office EX.app':basename(prefix)),'Unexpected app folder');
const expectedExecutable=process.platform==='darwin'?'Contents/MacOS/Office EX':process.platform==='win32'?'Office EX.exe':'office-ex';
assert(manifest.executable===expectedExecutable,'Unexpected executable');
const hash=async path=>{const h=createHash('sha256');for await(const b of createReadStream(path))h.update(b);return h.digest('hex');};
const sums=await readFile(prefix+'.SHA256SUMS','utf8'),seen=new Set();
for(const line of sums.trim().split('\n')){
 const match=/^([a-f0-9]{64})  ([A-Za-z0-9. _-]+)$/.exec(line);assert(match,'Invalid SHA256SUMS');const [,digest,name]=match;
 assert(name.startsWith(basename(prefix))&&!seen.has(name),'Unexpected checksum entry');seen.add(name);
 assert.equal(await hash(join(dirname(prefix),name)),digest,'Checksum mismatch: '+name);
}
const archive=prefix+(process.platform==='linux'?'.tar.gz':'.zip');assert(seen.has(basename(archive))&&seen.has(basename(manifestPath)),'Incomplete checksum list');
const work=await mkdtemp(join(tmpdir(),'office-ex-preview-check-')),profile=join(work,'profile'),checks=[],startedAt=new Date().toISOString(),errors=[];let app;
const report={schema:1,scope:'Locally generated archive extraction and actual packaged Electron launch; not installer, publisher signature, physical OS or reference Office acceptance.',sourceRevision:manifest.sourceRevision,sourceClean:manifest.sourceClean,platform:process.platform,arch:process.arch,version:manifest.version,archive:basename(archive),archiveSha256:await hash(archive),checks,startedAt};
const record=name=>{checks.push(name);console.log('PASS '+name);};
const safePath=value=>typeof value==='string'&&value.length>0&&!isAbsolute(value)&&!value.includes('\\')&&value.split('/').every(p=>p&&p!=='.'&&p!=='..');
try{
 const tool=process.platform==='win32'?'tar.exe':'tar',names=execFileSync(tool,['-tf',archive],{encoding:'utf8',maxBuffer:8*1024*1024}).split(/\r?\n/).filter(Boolean);
 for(const raw of names){const name=raw.replace(/\/$/,'');assert(safePath(name),'Unsafe archive path');assert(name===manifest.appFolder||name.startsWith(manifest.appFolder+'/')||process.platform==='darwin'&&(name==='__MACOSX'||name.startsWith('__MACOSX/')),'Unexpected archive root');}
 if(process.platform==='darwin')execFileSync('ditto',['-x','-k',archive,work],{stdio:'inherit'});
 else execFileSync(tool,['-xf',archive,'-C',work],{stdio:'inherit'});
 const folder=join(work,manifest.appFolder);
 // Each file or link is compared to the independently computed staging manifest.
 const actual=[];
 async function visit(dir){for(const name of (await readdir(dir)).sort()){
  const path=join(dir,name),info=await lstat(path),rel=relative(folder,path).split(sep).join('/');
  if(info.isDirectory())await visit(path);
  else if(info.isSymbolicLink()){const target=await readlink(path),resolved=resolve(dirname(path),target);assert(!isAbsolute(target)&&(resolved===folder||resolved.startsWith(folder+sep)),'Unsafe package symlink');actual.push({path:rel,symlink:target});}
  else{assert(info.isFile(),'Unexpected special file');actual.push({path:rel,bytes:info.size,sha256:await hash(path)});}
 }}
 await visit(folder);assert.deepEqual(actual,manifest.files,'Extracted file tree mismatch');record('Archive checksum and every packaged file/link verified');
 const appDir=process.platform==='darwin'?join(folder,'Contents','Resources','app'):join(folder,'resources','app');
 const appPackage=JSON.parse(await readFile(join(appDir,'package.json'),'utf8'));assert.equal(appPackage.version,manifest.version);
 assert.equal(appPackage.main,'dist/desktop/main.cjs');assert.ok(actual.some(f=>f.path.endsWith('/dist/licenses/THIRD_PARTY_NOTICES.md')));
 await mkdir(profile);
 if(process.platform==='linux'&&process.env.OFFICE_EX_CONFIGURE_TEST_SANDBOX==='1'){
  // CI-only setup in the freshly extracted private directory. No privilege
  // changes to user files or distributed archives.
  execFileSync('sudo',['chown','root:root',join(folder,'chrome-sandbox')],{stdio:'inherit'});
  execFileSync('sudo',['chmod','4755',join(folder,'chrome-sandbox')],{stdio:'inherit'});
 }
 const env={...process.env,OFFICE_EX_PROFILE_DIR:profile};delete env.ELECTRON_RUN_AS_NODE;
 const start=performance.now();
 app=await electron.launch({executablePath:join(folder,manifest.executable),chromiumSandbox:true,args:process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[],env,timeout:30000});
 const page=await app.firstWindow();page.on('pageerror',e=>errors.push(String(e)));await page.waitForSelector('body[data-ready=true]',{timeout:30000});
 report.startupMs=performance.now()-start;
 report.runtime=await app.evaluate(({app,BrowserWindow})=>{const w=BrowserWindow.getAllWindows()[0],p=w.webContents.getLastWebPreferences();return {packaged:app.isPackaged,name:app.getName(),version:app.getVersion(),electron:process.versions.electron,chromium:process.versions.chrome,sandbox:p.sandbox,contextIsolation:p.contextIsolation,nodeIntegration:p.nodeIntegration,webSecurity:p.webSecurity,devToolsOpened:(w.webContents.openDevTools({mode:'detach'}),w.webContents.isDevToolsOpened()),noSandbox:app.commandLine.hasSwitch('no-sandbox')||app.commandLine.hasSwitch('disable-setuid-sandbox')};});
 assert.deepEqual(report.runtime,{...report.runtime,packaged:true,name:'Office EX',version:manifest.version,electron:manifest.electron,sandbox:true,contextIsolation:true,nodeIntegration:false,webSecurity:true,devToolsOpened:false,noSandbox:false});
 assert.deepEqual(await page.evaluate(()=>[typeof window.require,typeof window.process]),['undefined','undefined']);record('Actual packaged executable launches with sandbox, restricted preload and disabled DevTools');
 const brandingPath=join(appDir,'dist','desktop','branding','manifest.json'),branding=JSON.parse(await readFile(brandingPath,'utf8'));
 assert.equal(manifest.branding.sha256,await hash(brandingPath));assert.ok(branding.files.some(f=>f.name==='office-ex.ico')&&branding.files.some(f=>f.name==='office-ex.icns'));
 for(const item of branding.files)assert.equal(await hash(join(appDir,'dist','desktop','branding',item.name)),item.sha256,item.name);
 await page.waitForFunction(()=>{const icon=document.querySelector('.brand-icon');return icon instanceof HTMLImageElement&&icon.complete&&icon.naturalWidth===128;});
 assert.equal(await page.locator('#product-version').textContent(),manifest.version);
 assert.equal(await page.locator('#footer-version').textContent(),'Office EX · '+manifest.version);
 await page.locator('#about').click();await page.waitForSelector('#about-dialog[open]');
 assert.equal(await page.locator('#about-version').textContent(),manifest.version);
 await page.locator('#about-dialog button[value=close]').click();
 assert.equal(await page.locator('.product-card').count(),3);assert.equal(await page.locator('.internal-drafts').evaluate(el=>el.open),false);
 if(process.platform==='darwin')assert.equal(await hash(join(folder,'Contents','Resources','office-ex.icns')),await hash(join(appDir,'dist','desktop','branding','office-ex.icns')));
 record('Product branding, native icon assets, real version, About and native-format home verified');

 await page.locator('#settings').click();await page.locator('#language').selectOption('en-US');await page.locator('#document-language').selectOption('en-US');await page.locator('#settings-dialog button[value=close]').click();
 for(const format of ['doc','docx','xls','xlsx','ppt','pptx']){
  console.log('CHECK '+format);const name='preview-check.'+format,path=join(work,name);await page.locator('#new-'+format).click();
  await page.waitForSelector('.tab[aria-selected=true]');
  if(format==='docx'){await page.locator('#draft-text').fill('Packaged Writer test');}
  else if(format==='ppt'||format==='pptx'){await page.locator('#slide-insert-text').click();await page.locator('#slide-text-editor').fill('Packaged Slides test');}
  await app.evaluate(({dialog},filePath)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath});dialog.showOpenDialog=async()=>({canceled:false,filePaths:[filePath]});dialog.showMessageBox=async()=>({response:1,checkboxChecked:false});},path);
  await page.locator('#save').click();try{await page.waitForFunction(name=>document.querySelector('.tab[aria-selected=true] .tab-title')?.textContent===name,name,{timeout:30000});}catch(error){const ui=await page.evaluate(()=>({title:document.querySelector('.tab[aria-selected=true] .tab-title')?.textContent,status:document.querySelector('#document-status')?.textContent,notice:document.querySelector('#notice')?.textContent,save:document.querySelector('#save')?.disabled}));throw new Error(format+' save: '+JSON.stringify(ui)+'; '+error);}
  const bytes=await readFile(path),inspection=await inspectForEditor(bytes,name,'package-check-'+format);assert.equal(inspection.report.readOnly,false,format);
  assert.equal(inspection.editableWriterFormat??inspection.editableSheetsFormat??inspection.editableSlidesFormat,format);
  await page.locator('.tab[aria-selected=true] .tab-close').click();await page.waitForFunction(()=>document.querySelectorAll('.tab').length===0);
  await page.locator('#open').click();await page.waitForFunction(name=>document.querySelector('.tab[aria-selected=true] .tab-title')?.textContent===name,name);
  await page.waitForFunction(()=>document.querySelector('#save')&&!document.querySelector('#save').disabled);
  await page.locator('.tab[aria-selected=true] .tab-close').click();await page.waitForFunction(()=>document.querySelectorAll('.tab').length===0);
  record('Packaged create, native save, parse and reopen '+format.toUpperCase());
 }
 const blocked=await page.evaluate(async()=>{try{await fetch('https://example.invalid/office-ex-preview-check');return false;}catch{return true;}});
 assert.equal(blocked,true);assert.deepEqual(errors,[]);record('Packaged renderer blocks external requests and reports no page errors');
 let timer;try{await Promise.race([app.close(),new Promise((_,reject)=>{timer=setTimeout(()=>reject(Error('Packaged Quit timeout')),20000);})]);}finally{clearTimeout(timer);}app=undefined;
 record('Packaged app exits normally');report.completedAt=new Date().toISOString();
}catch(error){report.failure=String(error);if(app){const child=app.process();const exited=new Promise(resolve=>child.once('exit',resolve));await app.evaluate(({app})=>app.exit(1)).catch(()=>{});if(child.exitCode===null)await Promise.race([exited,new Promise(resolve=>setTimeout(resolve,3000))]);}throw error;}
finally{await writeFile(prefix+'.verification.json',JSON.stringify(report,null,2)+'\n');await rm(work,{recursive:true,force:true,maxRetries:8,retryDelay:200});}
console.log(JSON.stringify(report,null,2));
