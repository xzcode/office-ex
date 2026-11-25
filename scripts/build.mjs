import { auditDependencies } from './dependency-inventory.mjs';
import { build } from 'esbuild';
import { mkdir, copyFile, writeFile, readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
await mkdir('dist/desktop',{recursive:true});await mkdir('dist/ui',{recursive:true});
const common={bundle:true,sourcemap:'external',legalComments:'external',logLevel:'warning'};
const results=await Promise.all([
  build({...common,entryPoints:['apps/desktop/src/main.ts'],outfile:'dist/desktop/main.cjs',platform:'node',format:'cjs',target:'node24',external:['electron'],metafile:true}),
  build({...common,entryPoints:['apps/desktop/src/preload.ts'],outfile:'dist/desktop/preload.cjs',platform:'node',format:'cjs',target:'node24',external:['electron'],metafile:true}),
  build({...common,entryPoints:['apps/desktop/src/docx-verifier-preload.ts'],outfile:'dist/desktop/docx-verifier-preload.cjs',platform:'node',format:'cjs',target:'node24',external:['electron'],metafile:true}),
  build({...common,entryPoints:['packages/ui/src/verify-docx.ts'],outfile:'dist/ui/verify-docx.js',platform:'browser',format:'esm',target:'chrome156',metafile:true}),
  build({...common,entryPoints:['packages/ui/src/renderer.ts'],outfile:'dist/ui/renderer.js',platform:'browser',format:'esm',target:'chrome156',metafile:true}),
  build({...common,entryPoints:['packages/ui/src/sheets-recalc-worker.ts'],outfile:'dist/ui/sheets-recalc-worker.js',platform:'browser',format:'esm',target:'chrome156',metafile:true}),
  build({...common,entryPoints:['packages/ui/src/format-worker.ts'],outfile:'dist/ui/format-worker.js',platform:'browser',format:'esm',target:'chrome156',metafile:true})
]);
await copyFile('packages/ui/index.html','dist/ui/index.html');await copyFile('packages/ui/verify-docx.html','dist/ui/verify-docx.html');await copyFile('packages/ui/styles.css','dist/ui/styles.css');
await writeFile('dist/build-meta.json',JSON.stringify({version:JSON.parse(await readFile('package.json','utf8')).version,outputs:results.map(x=>x.metafile)},null,2)+'\n');
await auditDependencies({stageDir:'dist/licenses'});
console.log(`Built Office EX desktop + sandboxed renderer/worker in ${resolve('dist')}`);
