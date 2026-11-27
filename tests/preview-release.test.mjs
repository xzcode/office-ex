// SPDX-License-Identifier: MIT
import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,rm,writeFile,readFile} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {createHash} from 'node:crypto';
import {preparePreviewRelease,validatePreviewTag} from '../scripts/prepare-preview-release.mjs';
const version='0.1.0-dev.1',tag='v'+version,source='a'.repeat(40),runId='123456';
const digest=bytes=>createHash('sha256').update(bytes).digest('hex');
const platformList=[['linux','x64','.tar.gz','.deb'],['win32','x64','.zip','-setup.exe'],['darwin','arm64','.zip','.dmg']];
async function fixture(root,mutate){
 for(const [platform,arch,archiveExt,installerExt]of platformList){
  const stem='Office-EX-'+version+'-'+platform+'-'+arch,dir=join(root,'office-ex-preview-'+platform+'-'+arch);await mkdir(dir,{recursive:true});
  const manifest={schema:1,version,sourceRevision:source,sourceClean:true,formalRelease:false,channel:'development-preview',platform,arch,signing:platform==='darwin'?'ad-hoc':'unsigned',branding:{sha256:'b'.repeat(64)}};
  const common={sourceRevision:source,version,platform,arch,completedAt:'2026-11-25T00:00:00Z',runtime:{packaged:true,sandbox:true,contextIsolation:true,nodeIntegration:false,noSandbox:false}};
  const archive=stem+archiveExt,installer=stem+installerExt,a=Buffer.from('archive '+platform),i=Buffer.from('installer '+platform);
  const verification={...common,sourceClean:true,archive,archiveSha256:digest(a),checks:['Archive checksum and every packaged file/link verified','Actual packaged executable launches with sandbox, restricted preload and disabled DevTools','Product branding, native icon assets, real version, About and native-format home verified',...['DOC','DOCX','XLS','XLSX','PPT','PPTX'].map(x=>'Packaged create, native save, parse and reopen '+x),'Packaged renderer blocks external requests and reports no page errors','Packaged app exits normally']};
  const installation={...common,installer,sha256:digest(i),checks:['Native installer creates a fresh installation','Installed application files match the source-correlated content manifest','Installed product identity and icon assets verified','Installed executable launches with sandbox and opens an editable native worksheet','Installed application exits normally','Native uninstall / removal completes']};
  mutate?.(platform,{manifest,verification,installation});
  const m=Buffer.from(JSON.stringify(manifest));
  for(const [name,data]of [[archive,a],[installer,i],[stem+'.manifest.json',m],[stem+'.verification.json',JSON.stringify(verification)],[stem+'.installation.json',JSON.stringify(installation)]])await writeFile(join(dir,name),data);
  await writeFile(join(dir,stem+'.SHA256SUMS'),[[archive,a],[installer,i],[stem+'.manifest.json',m]].map(([name,data])=>digest(data)+'  '+name).join('\n')+'\n');
 }
}
test('preview release assembly only admits exact same-source native evidence and stays draft-only',async()=>{
 const root=await mkdtemp(join(tmpdir(),'office-ex-release-test-'));
 try{
  const input=join(root,'input'),out=join(root,'output');await fixture(input);
  const result=await preparePreviewRelease(input,out,version,tag,source,runId);
  assert.equal(result.formalRelease,false);assert.equal(result.publication,'draft-only');assert.equal(result.platforms.length,3);
  assert.equal((await readFile(join(out,'SHA256SUMS'),'utf8')).trim().split('\n').length,19);
  assert.match(await readFile(join(out,'RELEASE-NOTES.md'),'utf8'),/not notarized/);
  assert.throws(()=>validatePreviewTag('v0.1.0',version,source));assert.throws(()=>validatePreviewTag('v0.1.0','0.1.0',source));
  for(const [name,mutate,pattern]of [
   ['source', (platform,v)=>{if(platform==='win32')v.manifest.sourceRevision='c'.repeat(40);}, /provenance/],
   ['sandbox',(platform,v)=>{if(platform==='darwin')v.installation.runtime.sandbox=false;},/native verification/],
   ['incomplete',(platform,v)=>{if(platform==='linux')v.verification.checks.pop();},/Archive evidence/],
   ['brand',(platform,v)=>{if(platform==='win32')v.manifest.branding.sha256='c'.repeat(64);},/branding mismatch/],
   ['cleanup',(platform,v)=>{if(platform==='linux')v.installation.cleanupFailure='removal failed';},/native verification/]
  ]){
   const sourceDir=join(root,name);await fixture(sourceDir,mutate);
   await assert.rejects(()=>preparePreviewRelease(sourceDir,join(root,name+'-out'),version,tag,source,runId),pattern);
  }
 }finally{await rm(root,{recursive:true,force:true});}
});

test('product version matches all workspace manifests and audited lock',async()=>{
 const root=new URL('../',import.meta.url);
 const read=async path=>JSON.parse(await readFile(new URL(path,root),'utf8'));
 const pkg=await read('package.json'),lockBytes=await readFile(new URL('package-lock.json',root)),lock=JSON.parse(lockBytes);
 const inventory=await read('docs/DEPENDENCIES.lock.json');
 assert.match(pkg.version,/^\d+\.\d+\.\d+(?:-[\w.-]+)?$/);
 assert.equal(lock.version,pkg.version);assert.equal(lock.packages[''].version,pkg.version);
 assert.equal(inventory.lockSha256,digest(lockBytes));
 const workspaces=Object.entries(lock.packages).filter(([path])=>path&&!path.startsWith('node_modules/'));
 assert.ok(workspaces.length>=14);
 for(const [path,entry] of workspaces){
  const manifest=await read(path+'/package.json');
  assert.equal(manifest.version,pkg.version,path);
  assert.equal(entry.version,pkg.version,path+' lock version');
  for(const [dep,requested] of Object.entries(manifest.dependencies??{})){
   if(!dep.startsWith('@office-ex/'))continue;
   assert.equal(requested,pkg.version,path+' '+dep);
   assert.equal(entry.dependencies[dep],requested,path+' '+dep+' lock dependency');
  }
 }
});
