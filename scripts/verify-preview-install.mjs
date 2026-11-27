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
import {fileTypes,desktopEntry,mimePackage,progId,windowsCapabilities,applicationName} from './file-associations.mjs';

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
const regQuery=(key,name)=>{const result=spawnSync('reg.exe',['query','HKCU\\'+key,...(name===undefined?['/ve']:['/v',name])],{encoding:'utf8',windowsHide:true});return result.status===0?result.stdout.trim():null;};
const winDefaultSnapshot=()=>JSON.parse(execFileSync('pwsh',['-NoProfile','-NonInteractive','-Command',[
 '$ErrorActionPreference="Stop";',
 'function Read-RegValue($Path,$Name) { $k=Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue; if ($null -eq $k) { return $null }; return $k.GetValue($Name,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) };',
 '$types=$env:OEX_TEST_FILE_TYPES | ConvertFrom-Json;',
 '$result=@(foreach ($ext in $types) { $class="Registry::HKEY_CURRENT_USER\\Software\\Classes\\."+$ext; $choice="Registry::HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\."+$ext+"\\UserChoice"; [ordered]@{ext=$ext;default=(Read-RegValue $class "");choice=(Read-RegValue $choice "ProgId");hash=(Read-RegValue $choice "Hash")} });',
 'ConvertTo-Json -InputObject $result -Compress -Depth 4'
].join(' ')],{encoding:'utf8',env:{...process.env,OEX_TEST_FILE_TYPES:JSON.stringify(fileTypes.map(f=>f.extension))}}));
const priorDefaults=platform==='win32'?winDefaultSnapshot():undefined;
async function verifyAssociations(){
 if(platform==='linux'){
  assert.equal(await readFile('/usr/share/applications/office-ex-preview.desktop','utf8'),desktopEntry());
  assert.equal(await readFile('/usr/share/mime/packages/office-ex-preview.xml','utf8'),mimePackage());
 }else if(platform==='darwin'){
  const plist=JSON.parse(execFileSync('/usr/bin/plutil',['-convert','json','-o','-',join(installed,'Contents','Info.plist')],{encoding:'utf8'}));
  assert.equal(plist.CFBundleDocumentTypes.length,fileTypes.length);
  for(const f of fileTypes){
   const type=plist.CFBundleDocumentTypes.find(v=>v.LSItemContentTypes.includes(f.uti));
   assert(type&&type.CFBundleTypeRole===f.role&&type.CFBundleTypeExtensions.includes(f.extension)&&type.LSHandlerRank===(f.extension==='oex'?'Owner':'Alternate'),f.extension);
  }
  assert(plist.UTExportedTypeDeclarations.some(v=>v.UTTypeIdentifier==='com.officex.preview.draft'));
 }else{
  const exe=join(installed,'Office EX.exe');
  assert.ok(regQuery('Software\\RegisteredApplications',applicationName)?.includes(windowsCapabilities));
  for(const f of fileTypes){
   assert.ok(regQuery('Software\\Classes\\.'+f.extension+'\\OpenWithProgids',progId(f.extension)),f.extension);
   assert.ok(regQuery(windowsCapabilities+'\\FileAssociations','.'+f.extension)?.includes(progId(f.extension)),f.extension);
   assert.ok(regQuery('Software\\Classes\\'+progId(f.extension)+'\\shell\\open\\command')?.includes('"'+exe+'" "%1"'),f.extension);
  }
  assert.deepEqual(winDefaultSnapshot(),priorDefaults,'Installer changed user default associations');
 }
 record('Native installer advertises all eight supported document types without forcing a default application');
}
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
  run('sudo',['dpkg','--purge','office-ex-preview']);assert(!await exists(installed),'DEB installation directory remains');assert(!await exists('/usr/share/applications/office-ex-preview.desktop')&&!await exists('/usr/share/mime/packages/office-ex-preview.xml'),'Association registration remains');
 }else if(platform==='win32'){
  const executable=join(installed,'unins000.exe');assert(await exists(executable),'Missing installed uninstaller');
  run(executable,['/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART']);
  const until=Date.now()+20000;while(await exists(installed)&&Date.now()<until)await new Promise(r=>setTimeout(r,200));
  assert(!await exists(installed),'Windows installation directory remains');
  for(const f of fileTypes){assert.equal(regQuery('Software\\Classes\\.'+f.extension+'\\OpenWithProgids',progId(f.extension)),null);assert.equal(regQuery('Software\\Classes\\'+progId(f.extension)+'\\shell\\open\\command'),null);}
  assert.equal(regQuery('Software\\RegisteredApplications',applicationName),null);assert.deepEqual(winDefaultSnapshot(),priorDefaults);
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
 assert.deepEqual(manifest.fileAssociations,fileTypes.map(f=>f.extension));await verifyAssociations();
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
 await page.locator('#settings').click();await page.locator('#language').selectOption('en-US');await page.locator('#settings-dialog button[value=close]').click();
 await page.waitForFunction(()=>{const icon=document.querySelector('.brand-icon');return icon?.complete&&icon.naturalWidth===128;});assert.equal(await page.locator('#product-version').textContent(),manifest.version);record('Installed product identity and icon assets verified');
 const created=app.waitForEvent('window',{timeout:30000});await page.locator('#new-xlsx').click();const document=await created;
 await document.waitForSelector('#sheets-scroll [role=gridcell]');assert.equal(await document.locator('html').getAttribute('data-window-role'),'document');assert.equal(await page.locator('.tab').count(),0);
 record('Installed executable launches with sandbox and opens an editable native worksheet');
 const fixture=join(work,'Native association test.xlsx');
 await document.locator('#sheets-formula').fill('42');await document.locator('#sheets-formula').press('Enter');
 await app.evaluate(({dialog},path)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:path});dialog.showMessageBox=async()=>({response:1,checkboxChecked:false});},fixture);
 await document.locator('#save').click();await document.waitForFunction(name=>document.title.startsWith(name),basename(fixture));
 const closed=document.waitForEvent('close',{timeout:30000});await document.locator('.tab-close').click();await closed;
 const opening=app.waitForEvent('window',{timeout:30000});
 // Invoke the installed handler's actual executable (the Linux wrapper
 // uses the registered desktop entry). This is not a forced OS default.
 if(platform==='linux')run('gio',['launch','/usr/share/applications/office-ex-preview.desktop',fixture],{env,timeout:30000});
 else run(join(installed,...manifest.executable.split('/')),[fixture],{env,timeout:30000});
 const reopened=await opening;await reopened.waitForSelector('#sheets-scroll [role=gridcell]');await reopened.waitForFunction(()=>document.querySelector('[data-address=A1]')?.textContent==='42');
 assert.equal(await reopened.locator('html').getAttribute('data-window-role'),'document');assert.equal(await page.locator('.tab').count(),0);
 record('Installed handler accepts a real filename with spaces into a separate window and reopens the saved value');
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
