// SPDX-License-Identifier: MIT
/** Native, offline development-preview packaging. No publishing or signing keys. */
import assert from 'node:assert/strict';
import {cp,copyFile,chmod,lstat,mkdir,mkdtemp,readFile,readdir,readlink,realpath,rename,rm,stat,symlink,writeFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {createReadStream} from 'node:fs';
import {execFileSync,spawnSync} from 'node:child_process';
import {basename,dirname,isAbsolute,join,relative,resolve,sep} from 'node:path';
import {tmpdir} from 'node:os';

const root=resolve('.'),out=resolve(process.env.OFFICE_EX_PACKAGE_OUT??'out/m5-preview');
const platform=process.platform,arch=process.arch;
assert(['linux','win32','darwin'].includes(platform),'Unsupported platform');
assert((platform==='darwin'&&arch==='arm64')||(platform!=='darwin'&&arch==='x64'),'Preview builds target macOS ARM64, Linux/Windows x64');
const pkg=JSON.parse(await readFile('package.json','utf8')),version=pkg.version;
assert(/^\d+\.\d+\.\d+-dev\.\d+$/.test(version),'This packager is limited to development previews');
const source=process.env.OFFICE_EX_SOURCE_SHA??process.env.GITHUB_SHA??'';
assert(/^[a-f0-9]{40}$/.test(source),'Set OFFICE_EX_SOURCE_SHA to the source commit (or use CI)');
const sourceClean=process.env.GITHUB_ACTIONS==='true'&&process.env.GITHUB_SHA===source&&spawnSync('git',['diff','--quiet','HEAD','--'],{cwd:root}).status===0&&spawnSync('git',['status','--porcelain','--untracked-files=normal'],{cwd:root,encoding:'utf8'}).stdout?.trim()==='';
const slug='Office-EX-'+version+'-'+platform+'-'+arch;
const executableName=platform==='win32'?'Office EX.exe':platform==='linux'?'office-ex':'Office EX';
const hash=async file=>{const h=createHash('sha256');for await(const b of createReadStream(file))h.update(b);return h.digest('hex');};
const run=(program,args,opts={})=>execFileSync(program,args,{stdio:'inherit',...opts});
async function files(dir){
 const result=[];
 async function visit(base){for(const name of (await readdir(base)).sort()){
  const path=join(base,name),info=await lstat(path),rel=relative(dir,path).split(sep).join('/');
  if(info.isSymbolicLink()){
   const target=await readlink(path),destination=resolve(dirname(path),target);
   assert(!isAbsolute(target)&&(destination===dir||destination.startsWith(dir+sep)),'Unsafe archive symlink: '+rel);
   result.push({path:rel,symlink:target});
  }else if(info.isDirectory())await visit(path);
  else{assert(info.isFile(),'Unexpected special file: '+rel);result.push({path:rel,bytes:info.size,sha256:await hash(path)});}
 }}
 await visit(dir);return result;
}
await mkdir(out,{recursive:true});
const work=await mkdtemp(join(out,'.stage-')),folder=platform==='darwin'?'Office EX.app':slug,stage=join(work,folder);
let success=false;
try{
 const electronRoot=await realpath('node_modules/electron/dist');
 const actualElectron=JSON.parse(await readFile('node_modules/electron/package.json','utf8')).version;
 if(platform!=='darwin')assert((await readFile(join(electronRoot,'version'),'utf8')).trim()===actualElectron,'Electron distribution version mismatch');
 else{const v=execFileSync('/usr/libexec/PlistBuddy',['-c','Print :CFBundleShortVersionString',join(electronRoot,'Electron.app','Contents','Info.plist')],{encoding:'utf8'}).trim();assert.equal(v,actualElectron,'Electron application version mismatch');}
 assert(actualElectron===pkg.devDependencies.electron,'Electron version mismatch');
 const fonts=JSON.parse(await readFile('assets/fonts/manifest.json','utf8'));
 for(const file of fonts.files){const path=join('dist/ui/fonts',file.file);assert((await stat(path)).size===file.bytes&&await hash(path)===file.sha256,'Bundled font mismatch: '+file.file);}
 const licenseFiles=['Office-EX-LICENSE','THIRD_PARTY_NOTICES.md','DEPENDENCIES.lock.json','electron-distribution.json'];
 for(const name of licenseFiles)assert((await stat(join('dist/licenses',name))).size>0,'Missing license: '+name);
 const branding=JSON.parse(await readFile('dist/desktop/branding/manifest.json','utf8'));
 for(const file of branding.files){const path=join('dist/desktop/branding',file.name);assert((await stat(path)).size===file.bytes&&await hash(path)===file.sha256,'Branding asset mismatch: '+file.name);}
 const built=JSON.parse(await readFile('dist/build-meta.json','utf8'));assert(built.version===version,'Stale desktop build');
 if(platform==='darwin')await cp(join(electronRoot,'Electron.app'),stage,{recursive:true,verbatimSymlinks:true});
 else{
  await cp(electronRoot,stage,{recursive:true,verbatimSymlinks:true});
  await rename(join(stage,platform==='win32'?'electron.exe':'electron'),join(stage,executableName));
  if(platform==='win32')run('pwsh',['-NoProfile','-NonInteractive','-File',resolve('scripts/set-windows-icon.ps1'),'-Executable',join(stage,executableName),'-Icon',resolve('dist/desktop/branding/office-ex.ico')]);
  if(platform==='linux')await chmod(join(stage,'chrome-sandbox'),0o755); // do not ship a setuid root executable
 }
 const resources=platform==='darwin'?join(stage,'Contents','Resources'):join(stage,'resources'),appDir=join(resources,'app');
 await rm(join(resources,'default_app.asar'),{force:true});
 await mkdir(appDir,{recursive:true});
 await writeFile(join(appDir,'package.json'),JSON.stringify({name:'office-ex',productName:'Office EX',version,private:true,main:'dist/desktop/main.cjs',license:'MIT'},null,2)+'\n');
 for(const part of ['desktop','ui','licenses'])await cp(join('dist',part),join(appDir,'dist',part),{recursive:true,filter:src=>!src.endsWith('.map')});
 await copyFile('LICENSE',join(appDir,'LICENSE'));
 for(const path of ['desktop/main.cjs','desktop/preload.cjs','desktop/docx-verifier-preload.cjs','desktop/pdf-printer-preload.cjs','ui/index.html','ui/renderer.js','ui/format-worker.js','ui/sheets-recalc-worker.js','ui/pdf.worker.mjs','ui/pdf-print.html','ui/verify-docx.html'])assert((await stat(join(appDir,'dist',path))).size>0,'Missing runtime file: '+path);
 const readme=platform==='darwin'?join(resources,'PREVIEW.txt'):join(stage,'PREVIEW.txt');
 await writeFile(readme,'Office EX '+version+' — DEVELOPMENT PREVIEW\nNot a certified Office-compatible release. No telemetry, account or online updater.\nWindows/macOS preview is not publisher-signed or notarized; OS trust warnings may apply.\nLinux requires an available Chromium sandbox and system desktop libraries. Do not disable the sandbox.\nLicense and third-party notices: resources/app/dist/licenses (macOS: Contents/Resources/app/dist/licenses).\nSee https://github.com/xzcode/office-ex for source and limitations.\n');
 if(platform==='darwin'){
  const plist=join(stage,'Contents','Info.plist');
  // Electron treats a main executable still named Electron as an unpackaged app on macOS.
  await rename(join(stage,'Contents','MacOS','Electron'),join(stage,'Contents','MacOS',executableName));
  for(const [key,value]of [['CFBundleExecutable',executableName],['CFBundleName','Office EX'],['CFBundleDisplayName','Office EX'],['CFBundleIconFile','office-ex.icns'],['CFBundleIdentifier','com.officex.preview'],['CFBundleShortVersionString',version.split('-')[0]],['CFBundleVersion',version.split('-')[0]]]){if(spawnSync('/usr/libexec/PlistBuddy',['-c','Print :'+key,plist],{stdio:'ignore'}).status===0)run('/usr/libexec/PlistBuddy',['-c','Set :'+key+' '+value,plist]);else run('/usr/libexec/PlistBuddy',['-c','Add :'+key+' string '+value,plist]);}
  await copyFile('dist/desktop/branding/office-ex.icns',join(resources,'office-ex.icns'));
  // The pinned runtime's nested helpers stay intact. Ad-hoc signing is not a
  // publisher signature or notarization.
  run('codesign',['--force','--deep','--sign','-',stage]);
  run('codesign',['--verify','--deep','--strict',stage]);
 }
 const entries=await files(stage);
 const manifest={schema:1,product:'Office EX',channel:'development-preview',formalRelease:false,version,sourceRevision:source,sourceClean,platform,arch,electron:actualElectron,signing:platform==='darwin'?'ad-hoc':'unsigned',branding:{sha256:await hash(join('dist/desktop/branding','manifest.json')),icon:platform==='darwin'?'Contents/Resources/office-ex.icns':platform==='win32'?executableName:'resources/app/dist/desktop/branding/office-ex-512.png'},appFolder:folder,executable:platform==='darwin'?'Contents/MacOS/'+executableName:executableName,files:entries};
 const manifestPath=join(out,slug+'.manifest.json');await writeFile(manifestPath,JSON.stringify(manifest,null,2)+'\n');
 const archive=join(out,slug+(platform==='linux'?'.tar.gz':'.zip'));
 await rm(archive,{force:true});
 if(platform==='darwin')run('ditto',['-c','-k','--sequesterRsrc','--keepParent',stage,archive]);
 else if(platform==='win32')run('tar.exe',['-a','-cf',archive,'-C',work,folder]);
 else run('tar',['-czf',archive,'-C',work,folder]);
 const outputs=[archive,manifestPath];
 if(process.argv.includes('--installer')){
  if(platform==='darwin'){
   const dmg=join(out,slug+'.dmg');await rm(dmg,{force:true});
   await symlink('/Applications',join(work,'Applications'));
   run('hdiutil',['create','-ov','-format','UDZO','-volname','Office EX Preview','-srcfolder',work,dmg]);
   run('hdiutil',['verify',dmg]);outputs.push(dmg);
  }else if(platform==='linux'){
   const debRoot=join(work,'deb'),target=join(debRoot,'opt','office-ex-preview');await mkdir(join(debRoot,'DEBIAN'),{recursive:true});await mkdir(dirname(target),{recursive:true});
   await cp(stage,target,{recursive:true,verbatimSymlinks:true});
   // The privileged DEB installation can establish the root-owned setuid
   // Chromium sandbox. The unprivileged portable archive remains mode 0755.
   await chmod(join(target,'chrome-sandbox'),0o4755);await mkdir(join(debRoot,'usr','bin'),{recursive:true});await mkdir(join(debRoot,'usr','share','applications'),{recursive:true});
   const wrapper=join(debRoot,'usr','bin','office-ex-preview');await writeFile(wrapper,'#!/bin/sh\nexec /opt/office-ex-preview/office-ex "$@"\n',{mode:0o755});
   await writeFile(join(debRoot,'usr','share','applications','office-ex-preview.desktop'),'[Desktop Entry]\nType=Application\nName=Office EX Preview\nExec=office-ex-preview %F\nTerminal=false\nIcon=office-ex-preview\nStartupWMClass=Office EX\nCategories=Office;\n');
   for(const size of [16,24,32,48,64,128,256,512]){const directory=join(debRoot,'usr','share','icons','hicolor',size+'x'+size,'apps');await mkdir(directory,{recursive:true});await copyFile('dist/desktop/branding/office-ex-'+size+'.png',join(directory,'office-ex-preview.png'));}
   await writeFile(join(debRoot,'DEBIAN','control'),'Package: office-ex-preview\nVersion: '+version.replace('-dev.','~dev.')+'\nArchitecture: amd64\nMaintainer: Office EX contributors\nDepends: libgtk-3-0t64 | libgtk-3-0, libnss3, libgbm1, libasound2t64 | libasound2, libxss1, libx11-xcb1\nSection: editors\nPriority: optional\nDescription: Office EX offline development preview\n Not a certified Office-compatible release.\n');
   const deb=join(out,slug+'.deb');await rm(deb,{force:true});run('dpkg-deb',['--root-owner-group','--build',debRoot,deb]);run('dpkg-deb',['--info',deb]);outputs.push(deb);
  }else{
   // Inno Setup is a native build-time compiler, never bundled in the app.
   const compiler=process.env.OFFICE_EX_ISCC??join(process.env['ProgramFiles(x86)']??'C:\\Program Files (x86)','Inno Setup 6','ISCC.exe');
   const iss=join(work,'preview.iss'),safe=value=>{assert(!/["\r\n]/.test(value),'Invalid installer path');return value;};
   await writeFile(iss,`[Setup]\r\nAppId=OfficeEXPreview\r\nAppName=Office EX Preview\r\nAppVersion=${version}\r\nDefaultDirName={localappdata}\\Programs\\Office EX Preview\r\nDisableDirPage=yes\r\nDisableProgramGroupPage=yes\r\nUsePreviousAppDir=no\r\nCloseApplications=no\r\nRestartApplications=no\r\nPrivilegesRequired=lowest\r\nArchitecturesAllowed=x64compatible\r\nArchitecturesInstallIn64BitMode=x64compatible\r\nOutputDir=${safe(out)}\r\nOutputBaseFilename=${slug}-setup\r\nCompression=lzma2\r\nSolidCompression=yes\r\nSetupIconFile=${safe(resolve('dist/desktop/branding/office-ex.ico'))}\r\nUninstallDisplayIcon={app}\\Office EX.exe\r\n[Files]\r\nSource: "${safe(stage)}\\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs\r\n[Icons]\r\nName: "{autoprograms}\\Office EX Preview"; Filename: "{app}\\Office EX.exe"\r\n`);
   run(compiler,[iss]);outputs.push(join(out,slug+'-setup.exe'));
  }
 }
 const sums=[];for(const path of outputs)sums.push(await hash(path)+'  '+basename(path));
 await writeFile(join(out,slug+'.SHA256SUMS'),sums.join('\n')+'\n');
 const report={archive,manifest:manifestPath,outputs:outputs.map(path=>basename(path)),sourceClean};
 console.log(JSON.stringify(report,null,2));success=true;
}finally{await rm(work,{recursive:true,force:true});if(!success)console.error('Preview packaging failed; incomplete output must not be distributed.');}
