// SPDX-License-Identifier: MIT
/** Offline release assembly. Accept only the three successful same-source
 * installer and archive artifacts. Does not create tags or publish releases. */
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {createReadStream} from 'node:fs';
import {mkdir,readdir,readFile,copyFile,writeFile,lstat} from 'node:fs/promises';
import {join,resolve,basename} from 'node:path';
import {pathToFileURL} from 'node:url';
const platforms=[['linux','x64','.tar.gz','.deb'],['win32','x64','.zip','-setup.exe'],['darwin','arm64','.zip','.dmg']];
const hash=async file=>{const h=createHash('sha256');for await(const chunk of createReadStream(file))h.update(chunk);return h.digest('hex');};
const json=async path=>JSON.parse(await readFile(path,'utf8'));
export function validatePreviewTag(tag,version,source){
 assert(/^\d+\.\d+\.\d+-dev\.\d+$/.test(version)&&tag==='v'+version,'Preview tag must match the development version exactly');
 assert(/^[a-f0-9]{40}$/.test(source),'Missing exact source revision');
}
export async function preparePreviewRelease(artifacts,target,version,tag,source,runId){
 validatePreviewTag(tag,version,source);assert(/^\d+$/.test(String(runId)),'Invalid workflow run');
 const dir=resolve(target);await mkdir(dir,{recursive:true});assert((await readdir(dir)).length===0,'Release assembly output must be empty');
 const output=[],summaries=[];let brandingDigest;
 for(const [platform,arch,archiveExt,installerExt]of platforms){
  const stem='Office-EX-'+version+'-'+platform+'-'+arch,artifact=join(artifacts,'office-ex-preview-'+platform+'-'+arch);
  const expected=[stem+archiveExt,stem+installerExt,stem+'.manifest.json',stem+'.SHA256SUMS',stem+'.verification.json',stem+'.installation.json'];
  assert.deepEqual((await readdir(artifact)).sort(),[...expected].sort(),'Incomplete or unexpected native artifact set');
  for(const name of expected)assert((await lstat(join(artifact,name))).isFile(),'Expected regular release artifact');
  const manifest=await json(join(artifact,stem+'.manifest.json')),verification=await json(join(artifact,stem+'.verification.json')),install=await json(join(artifact,stem+'.installation.json'));
  assert(manifest.schema===1&&manifest.version===version&&manifest.sourceRevision===source&&manifest.sourceClean===true&&manifest.formalRelease===false&&manifest.channel==='development-preview'&&manifest.platform===platform&&manifest.arch===arch,'Mismatched package provenance');
  assert(manifest.signing===(platform==='darwin'?'ad-hoc':'unsigned')&&/^[a-f0-9]{64}$/.test(manifest.branding?.sha256),'Missing branding or incorrect signing declaration');
  // PNG generation is deterministic in a pinned Node build; require the same
  // mark and generated native icon manifests on every platform.
  if(brandingDigest===undefined)brandingDigest=manifest.branding.sha256;else assert.equal(brandingDigest,manifest.branding.sha256,'Cross-platform branding mismatch');
  const sums=new Map();
  for(const line of (await readFile(join(artifact,stem+'.SHA256SUMS'),'utf8')).trim().split(/\r?\n/)){
   const match=/^([a-f0-9]{64})  ([A-Za-z0-9._-]+)$/.exec(line);assert(match&&!sums.has(match[2]),'Invalid checksum entry');sums.set(match[2],match[1]);
  }
  const checkFiles=expected.slice(0,3);assert.deepEqual([...sums.keys()].sort(),checkFiles.sort(),'Incomplete artifact checksums');
  for(const name of checkFiles)assert.equal(await hash(join(artifact,name)),sums.get(name),'Artifact checksum mismatch: '+name);
  for(const report of [verification,install]){
   assert(report.sourceRevision===source&&report.version===version&&report.platform===platform&&report.arch===arch&&report.completedAt&&!report.failure&&!report.cleanupFailure&&report.runtime?.packaged===true&&report.runtime.sandbox===true&&report.runtime.contextIsolation===true&&report.runtime.nodeIntegration===false&&report.runtime.noSandbox===false,'Missing successful native verification');
  }
  assert(verification.sourceClean===true&&verification.archive===stem+archiveExt&&verification.archiveSha256===sums.get(stem+archiveExt)&&verification.checks.length>=11,'Archive evidence mismatch');
  for(const format of ['DOC','DOCX','XLS','XLSX','PPT','PPTX'])assert(verification.checks.includes('Packaged create, native save, parse and reopen '+format),'Missing native format check '+format);
  assert(verification.checks.includes('Product branding, native icon assets, real version, About and native-format home verified'),'Missing product identity check');
  assert(install.installer===stem+installerExt&&install.sha256===sums.get(stem+installerExt)&&install.checks.includes('Native uninstall / removal completes')&&install.checks.includes('Installed product identity and icon assets verified'),'Installer evidence mismatch');
  for(const name of expected){await copyFile(join(artifact,name),join(dir,name));output.push(name);}
  if(manifest.fileAssociations!==undefined){
   assert.deepEqual(manifest.fileAssociations,['doc','docx','xls','xlsx','ppt','pptx','pdf','oex'],'Association coverage mismatch');
   assert(install.checks.includes('Native installer advertises all eight supported document types without forcing a default application')&&install.checks.includes('Installed handler accepts a real filename with spaces into a separate window and reopens the saved value'),'Missing native document association evidence');
  }
  summaries.push({platform,arch,archive:stem+archiveExt,installer:stem+installerExt,signing:manifest.signing,archiveChecks:verification.checks.length,installerChecks:install.checks.length});
 }
 const report={schema:1,version,tag,sourceRevision:source,workflowRun:String(runId),channel:'development-preview',formalRelease:false,publication:'draft-only',platforms:summaries};
 await writeFile(join(dir,'preview-release-evidence.json'),JSON.stringify(report,null,2)+'\n');output.push('preview-release-evidence.json');
 const notes='# Office EX '+version+' — Development Preview\n\nThis is an early open-source preview, not a certified Office-compatible release.\n\nSource: `'+source+'` · CI run '+runId+'\n\n| Platform | Recommended installer | Signing |\n|---|---|---|\n'+summaries.map(s=>'| '+s.platform+' '+s.arch+' | `'+s.installer+'` | '+s.signing+' |').join('\n')+'\n\n- Supports bounded DOC/DOCX, XLS/XLSX and PPT/PPTX editing plus offline PDF reading and Writer/Slides PDF output. Unsupported structures may be protected read-only.\n- The Windows installer is unsigned; macOS is ad-hoc signed, not notarized. Do not disable operating-system or Chromium security protections.\n- Linux needs compatible desktop libraries and a working sandbox.\n- Keep backups of important documents. Full reference Microsoft Office compatibility, legacy profiles, performance and human locale review remain open.\n- MIT source license and third-party notices are bundled. Per-platform manifests and checksums are included.\n';
 await writeFile(join(dir,'RELEASE-NOTES.md'),notes);
 const sums=[];for(const name of output.sort())sums.push(await hash(join(dir,name))+'  '+name);
 await writeFile(join(dir,'SHA256SUMS'),sums.join('\n')+'\n');return report;
}
if(import.meta.url===pathToFileURL(resolve(process.argv[1]??'')).href){
 const version=(await json('package.json')).version,tag=process.env.GITHUB_REF_NAME??'',source=process.env.GITHUB_SHA??'',runId=process.env.GITHUB_RUN_ID??'';
 console.log(JSON.stringify(await preparePreviewRelease(process.argv[2]??'out/release-input',process.argv[3]??'out/release-draft',version,tag,source,runId),null,2));
}
