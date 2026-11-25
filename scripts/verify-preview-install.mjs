// SPDX-License-Identifier: MIT
/** Install and remove our own trusted native installer on a disposable CI host.
 * Linux DEB is a system install: never run this on an ordinary workstation. */
import assert from 'node:assert/strict';
import {execFileSync,spawnSync} from 'node:child_process';
import {createHash} from 'node:crypto';
import {createReadStream} from 'node:fs';
import {lstat,mkdir,mkdtemp,readFile,readlink,rm,writeFile} from 'node:fs/promises';
import {basename,dirname,join,resolve} from 'node:path';
import {tmpdir} from 'node:os';
import {_electron as electron} from 'playwright-core';

assert.equal(process.env.GITHUB_ACTIONS,'true','Installer smoke is restricted to disposable Actions runners');
assert.equal(process.env.OFFICE_EX_INSTALL_TEST,'1','Explicit installer-test opt-in is required');
const manifestPath=resolve(process.argv[2]??'');
assert(process.argv[2]?.endsWith('.manifest.json'),'Pass the generated manifest');
const manifest=JSON.parse(await readFile(manifestPath,'utf8')),prefix=manifestPath.slice(0,-'.manifest.json'.length);
assert(manifest.schema===1&&manifest.channel==='development-preview'&&!manifest.formalRelease&&manifest.sourceClean&&manifest.sourceRevision===process.env.GITHUB_SHA,'Expected a clean source-correlated development preview');
assert(manifest.platform===process.platform&&manifest.arch===process.arch,'Installer platform mismatch');
assert(/^Office-EX-\d+\.\d+\.\d+-dev\.\d+-[a-z0-9]+-[a-z0-9]+$/.test(basename(prefix)),'Unexpected installer name');
const platform=process.platform,installer=prefix+(platform==='linux'?'.deb':platform==='darwin'?'.dmg':'-setup.exe');
const hash=async path=>{const h=createHash('sha256');for await(const b of createReadStream(path))h.update(b);return h.digest('hex');};
const sums=await readFile(prefix+'.SHA256SUMS','utf8'),sha=await hash(installer),manifestSha=await hash(manifestPath);
assert(sums.split(/\r?\n/).includes(sha+'  '+basename(installer))&&sums.split(/\r?\n/).includes(manifestSha+'  '+basename(manifestPath)),'Installer or manifest checksum mismatch');
const work=await mkdtemp(join(tmpdir(),'office-ex-installer-ci-')),profile=join(work,'profile'),mount=join(work,'mounted');
const checks=[],report={scope:'Disposable CI native installer installation, manifest verification, actual launch and removal; not publisher trust or reference Office acceptance.',sourceRevision:manifest.sourceRevision,platform,arch:process.arch,version:manifest.version,installer:basename(installer),sha256:sha,checks,startedAt:new Date().toISOString()};
const run=(program,args,options={})=>execFileSync(program,args,{stdio:'inherit',timeout:180000,...options});
const exists=async path=>{try{await lstat(path);return true;}catch(error){if(error.code==='ENOENT')return false;throw error;}};
let app,installed,attempted=false,mounted=false;
const record=name=>{checks.push(name);console.log('PASS '+name);};
async function verifyFiles(root){
 for(const entry of manifest.files){
  assert(typeof entry.path==='string'&&!entry.path.includes('\\')&&entry.path.split('/').every(p=>p&&p!=='.'&&p!=='..'),'Invalid manifest path');
  const path=join(root,...entry.path.split('/')),info=await lstat(path);
  if(entry.symlink!==undefined){assert(info.isSymbolicLink());assert.equal(await readlink(path),entry.symlink);}
  else {assert(info.isFile());assert.equal(info.size,entry.bytes,entry.path);assert.equal(await hash(path),entry.sha256,entry.path);}
 }
}
async function uninstall(){
 if(!attempted)return;
 if(platform==='linux'){
  run('sudo',['dpkg','--purge','office-ex-preview']);assert(!await exists(installed),'DEB installation directory remains');
 }else if(platform==='win32'){
  const executable=join(installed,'unins000.exe');assert(await exists(executable),'Missing installed uninstaller');
  run(executable,['/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART']);
  const until=Date.now()+20000;while(await exists(installed)&&Date.now()<until)await new Promise(r=>setTimeout(r,200));
  assert(!await exists(installed),'Windows installation directory remains');
 }else{
  await rm(installed,{recursive:true,force:true,maxRetries:8,retryDelay:200});assert(!await exists(installed));
 }
 attempted=false;record('Native uninstall / removal completes');
}
try{
 if(platform==='linux'){
  installed='/opt/office-ex-preview';
  assert(!await exists(installed),'Refusing to replace an existing installation');
  const existing=spawnSync('dpkg-query',['-W','-f=${db:Status-Status}','office-ex-preview'],{encoding:'utf8'});
  assert(existing.status!==0||existing.stdout.trim()!=='installed','Existing DEB installation');
  assert.equal(execFileSync('dpkg-deb',['-f',installer,'Package'],{encoding:'utf8'}).trim(),'office-ex-preview');
  attempted=true;run('sudo',['dpkg','-i',installer]);
  const helper=await lstat(join(installed,'chrome-sandbox'));assert.equal(helper.uid,0);assert.equal(helper.mode&0o7777,0o4755);
 }else if(platform==='win32'){
  installed=join(work,'Office EX Preview');
  assert(!await exists(installed));attempted=true;
  run(installer,['/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART','/SP-','/CURRENTUSER','/DIR='+installed]);
 }else{
  await mkdir(mount);run('hdiutil',['attach','-readonly','-nobrowse','-mountpoint',mount,installer]);mounted=true;
  installed=join(work,'Applications','Office EX.app');await mkdir(dirname(installed),{recursive:true});
  attempted=true;run('ditto',[join(mount,manifest.appFolder),installed]);run('codesign',['--verify','--deep','--strict',installed]);
  run('hdiutil',['detach',mount]);mounted=false;
 }
 record('Native installer creates a fresh installation');
 await verifyFiles(installed);record('Installed application files match the source-correlated content manifest');
 if(platform==='linux'){
  const desktop=await readFile('/usr/share/applications/office-ex-preview.desktop','utf8');assert.match(desktop,/^Icon=office-ex-preview$/m);
  for(const size of [16,24,32,48,64,128,256,512]){
   const icon='/usr/share/icons/hicolor/'+size+'x'+size+'/apps/office-ex-preview.png';
   assert.equal(await hash(icon),await hash(join(installed,'resources','app','dist','desktop','branding','office-ex-'+size+'.png')));
  }
 }else if(platform==='darwin'){
  const plist=execFileSync('/usr/libexec/PlistBuddy',['-c','Print :CFBundleIconFile',join(installed,'Contents','Info.plist')],{encoding:'utf8'}).trim();assert.equal(plist,'office-ex.icns');
  assert.equal(await hash(join(installed,'Contents','Resources','office-ex.icns')),await hash(join(installed,'Contents','Resources','app','dist','desktop','branding','office-ex.icns')));
 }

 await mkdir(profile);const env={...process.env,OFFICE_EX_PROFILE_DIR:profile};delete env.ELECTRON_RUN_AS_NODE;
 const start=performance.now();
 app=await electron.launch({executablePath:join(installed,...manifest.executable.split('/')),chromiumSandbox:true,env,timeout:30000});
 const page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]',{timeout:30000});
 report.startupMs=performance.now()-start;
 const runtime=await app.evaluate(({app,BrowserWindow})=>{const prefs=BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences();return {packaged:app.isPackaged,version:app.getVersion(),sandbox:prefs.sandbox,contextIsolation:prefs.contextIsolation,nodeIntegration:prefs.nodeIntegration,webSecurity:prefs.webSecurity,noSandbox:app.commandLine.hasSwitch('no-sandbox')||app.commandLine.hasSwitch('disable-setuid-sandbox')};});
 assert.deepEqual(runtime,{packaged:true,version:manifest.version,sandbox:true,contextIsolation:true,nodeIntegration:false,webSecurity:true,noSandbox:false});report.runtime=runtime;
 await page.waitForFunction(()=>{const icon=document.querySelector('.brand-icon');return icon?.complete&&icon.naturalWidth===128;});assert.equal(await page.locator('#product-version').textContent(),manifest.version);record('Installed product identity and icon assets verified');
 await page.locator('#new-xlsx').click();await page.waitForSelector('#sheets-scroll [role=gridcell]');record('Installed executable launches with sandbox and opens an editable native worksheet');
 await app.evaluate(({dialog})=>{dialog.showMessageBox=async()=>({response:1,checkboxChecked:false});});
 let timer;try{await Promise.race([app.close(),new Promise((_,reject)=>{timer=setTimeout(()=>reject(Error('Installed application Quit timed out')),20000);})]);}finally{clearTimeout(timer);}app=undefined;record('Installed application exits normally');
 await uninstall();report.completedAt=new Date().toISOString();
}catch(error){
 report.failure=String(error);if(app){const child=app.process();const exited=new Promise(resolve=>child.once('exit',resolve));await app.evaluate(({app})=>app.exit(1)).catch(()=>{});if(child.exitCode===null)await Promise.race([exited,new Promise(resolve=>setTimeout(resolve,3000))]);}
 // Remove only the installation created by this disposable CI test.
 if(attempted)try{await uninstall();}catch(cleanup){report.cleanupFailure=String(cleanup);}
 throw error;
}finally{
 if(mounted)try{run('hdiutil',['detach',mount]);}catch(error){report.cleanupFailure=String(error);}
 await writeFile(prefix+'.installation.json',JSON.stringify(report,null,2)+'\n');
 await rm(work,{recursive:true,force:true,maxRetries:8,retryDelay:200});
}
console.log(JSON.stringify(report,null,2));
