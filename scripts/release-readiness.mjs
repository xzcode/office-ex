// SPDX-License-Identifier: MIT
/** Machine-readable release boundary. Never turn a preview into a formal release. */
import {readFile,mkdir,writeFile} from 'node:fs/promises';
import assert from 'node:assert/strict';
const matrix=JSON.parse(await readFile('docs/specs/features.json','utf8'));
const pkg=JSON.parse(await readFile('package.json','utf8'));
const evidence=JSON.parse(await readFile('docs/verification/m5/release-gates.json','utf8'));
const requiredIds=['FMT-04','LEG-06','LEG-02','LEG-03','LEG-04','PDF-01','PDF-02','QA-03','QA-04','QA-05','QA-06','QA-07','I18N-01','I18N-02','I18N-03','I18N-04','I18N-05'];
const byId=new Map(matrix.features.map(f=>[f.id,f]));
const featureBlockers=requiredIds.flatMap(id=>{
 const f=byId.get(id);assert(f,'Required feature missing: '+id);
 return f.status==='verified'?[]:[{id,status:f.status,acceptance:f.acceptance}];
});
const required=['six-format-reference-office','legacy-doc','legacy-xls','legacy-ppt','native-installation','publisher-trust','performance','locale-human-review','native-ime-and-print'];
assert(evidence.schema===1&&Array.isArray(evidence.gates),'Invalid release gate evidence');
const gates=new Map(evidence.gates.map(g=>[g.id,g]));
assert(gates.size===evidence.gates.length,'Duplicate release gate');
const blockers=required.flatMap(id=>{const g=gates.get(id);assert(g,'Required gate missing: '+id);assert(['open','passed','blocked'].includes(g.status),'Invalid gate status: '+id);
 if(g.status==='passed')assert(typeof g.sourceRevision==='string'&&/^[a-f0-9]{40}$/.test(g.sourceRevision)&&g.evidence?.length,'A passed gate requires source-correlated evidence: '+id);
 return g.status==='passed'?[]:[{id,status:g.status,reason:g.reason}];});
const previewVersion=/-dev\.\d+$/.test(pkg.version);
const ready=!previewVersion&&!featureBlockers.length&&!blockers.length;
const report={schema:1,version:pkg.version,channel:previewVersion?'development-preview':'release-candidate',formalReleaseReady:ready,featureBlockers,blockers};
await mkdir('out/m5-readiness',{recursive:true});await writeFile('out/m5-readiness/report.json',JSON.stringify(report,null,2)+'\n');
console.log(JSON.stringify(report,null,2));
if(process.argv.includes('--require-formal')&&!ready)process.exitCode=1;
