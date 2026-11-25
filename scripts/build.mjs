import { auditDependencies } from './dependency-inventory.mjs';
import { preparePdfWorker } from '../packages/pdf/prepare-worker.mjs';
import { build } from 'esbuild';
import { mkdir, copyFile, writeFile, readFile, cp } from 'node:fs/promises';
import { resolve } from 'node:path';
await mkdir('dist/desktop',{recursive:true});await mkdir('dist/ui',{recursive:true});
const common={bundle:true,sourcemap:'external',legalComments:'external',logLevel:'warning'};
const results=await Promise.all([
  build({...common,entryPoints:['apps/desktop/src/main.ts'],outfile:'dist/desktop/main.cjs',platform:'node',format:'cjs',target:'node24',external:['electron'],metafile:true}),
  build({...common,entryPoints:['apps/desktop/src/preload.ts'],outfile:'dist/desktop/preload.cjs',platform:'node',format:'cjs',target:'node24',external:['electron'],metafile:true}),
  build({...common,entryPoints:['apps/desktop/src/docx-verifier-preload.ts'],outfile:'dist/desktop/docx-verifier-preload.cjs',platform:'node',format:'cjs',target:'node24',external:['electron'],metafile:true}),
  build({...common,entryPoints:['apps/desktop/src/pdf-printer-preload.ts'],outfile:'dist/desktop/pdf-printer-preload.cjs',platform:'node',format:'cjs',target:'node24',external:['electron'],metafile:true}),
  build({...common,entryPoints:['packages/ui/src/pdf-print.ts'],outfile:'dist/ui/pdf-print.js',platform:'browser',format:'esm',target:'chrome156',metafile:true}),
  build({...common,entryPoints:['packages/ui/src/verify-docx.ts'],outfile:'dist/ui/verify-docx.js',platform:'browser',format:'esm',target:'chrome156',metafile:true}),
  build({...common,entryPoints:['packages/ui/src/renderer.ts'],outfile:'dist/ui/renderer.js',platform:'browser',format:'esm',target:'chrome156',metafile:true}),
  build({...common,entryPoints:['packages/ui/src/sheets-recalc-worker.ts'],outfile:'dist/ui/sheets-recalc-worker.js',platform:'browser',format:'esm',target:'chrome156',metafile:true}),
  build({...common,entryPoints:['packages/ui/src/format-worker.ts'],outfile:'dist/ui/format-worker.js',platform:'browser',format:'esm',target:'chrome156',metafile:true})
]);
await copyFile('packages/ui/index.html','dist/ui/index.html');await copyFile('packages/ui/verify-docx.html','dist/ui/verify-docx.html');await writeFile('dist/ui/styles.css',(await Promise.all(['packages/ui/styles.css','packages/ui/slides-workspace.css','packages/ui/src/pdf-workspace.css'].map(path=>readFile(path,'utf8')))).join('\n'));
for(const name of ['pdf-print.html','pdf-print.css','fonts.css'])await copyFile('packages/ui/'+name,'dist/ui/'+name);
await writeFile('dist/ui/pdf.worker.mjs',preparePdfWorker(await readFile('node_modules/pdfjs-dist/build/pdf.worker.mjs','utf8')));
for(const name of ['cmaps','standard_fonts','iccs'])await cp('node_modules/pdfjs-dist/'+name,'dist/ui/pdf-assets/'+name,{recursive:true});
await cp('assets/fonts','dist/ui/fonts',{recursive:true});
await writeFile('dist/build-meta.json',JSON.stringify({version:JSON.parse(await readFile('package.json','utf8')).version,outputs:results.map(x=>x.metafile)},null,2)+'\n');
await auditDependencies({stageDir:'dist/licenses'});
console.log(`Built Office EX desktop + sandboxed renderer/worker in ${resolve('dist')}`);
