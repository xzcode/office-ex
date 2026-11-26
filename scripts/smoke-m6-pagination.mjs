// SPDX-License-Identifier: MIT
/** Native pagination controls, source-backed DOCX save, and real PDF export. */
import {revealControl} from './ribbon-test-controls.mjs';
import {_electron as electron} from 'playwright-core';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,readFile,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {createHash} from 'node:crypto';
import {execFileSync} from 'node:child_process';
import {createDocument} from '@office-ex/core';
import {writeWriterDocx,openWriterDocx} from '@office-ex/ooxml';
const work=await mkdtemp(join(tmpdir(),'office-ex-pagination-')),profile=join(work,'profile');
await mkdir(profile);
const file=join(work,'pagination.docx'),pdf=join(work,'pagination.pdf');
const model=createDocument('writer','pagination-native');
model.defaultTextStyle={fontFamily:'Noto Sans',fontSize:127000};
const section=model.sections[0];
section.page={width:4572000,height:1841500,margins:{left:254000,right:254000,top:508000,bottom:508000}};
section.pageExtras={headerDistance:127000,footerDistance:127000};
section.pageNumberStart=5;
section.storyRefs={header:'rIdOexStory_header',footer:'rIdOexStory_footer'};
model.pageStories={
 rIdOexStory_header:{kind:'header',text:'Independent PDF header',alignment:'center'},
 rIdOexStory_footer:{kind:'footer',text:'Page ',pageNumber:true,after:' / report',alignment:'end'}
};
const p=(id,text,style)=>({id,type:'paragraph',runs:[{id:id+'-run',type:'text',text}],...(style?{style}:{})});
section.blocks=[
 p('lead','Lead one\nLead two\nLead three'),
 p('heading','Keep heading',{keepNext:true,keepLines:true}),
 p('body','Body first\nBody last'),
 p('forced','Forced page',{pageBreakBefore:true})
];
await writeFile(file,writeWriterDocx(model));
const checks=[],hash=b=>createHash('sha256').update(b).digest('hex');
const check=name=>{checks.push(name);console.log('PASS',name);};let app;
async function wait(fn,label,timeout=30000){const until=Date.now()+timeout;let last;while(Date.now()<until){try{const result=await fn();if(result)return result;}catch(e){last=String(e);}await new Promise(r=>setTimeout(r,100));}throw Error('Timed out: '+label+(last?'; '+last:''));}
try{
 app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env:{...process.env,OFFICE_EX_PROFILE_DIR:profile},timeout:30000});
 const page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]');
 const errors=[];page.on('pageerror',e=>errors.push(String(e)));
 const prefs=await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences());
 assert.equal(prefs.sandbox,true);assert.equal(prefs.contextIsolation,true);assert.equal(prefs.nodeIntegration,false);assert.equal(prefs.webSecurity,true);
 await page.locator('#settings').click();await page.locator('#language').selectOption('zh-CN');await page.locator('#settings-dialog button[value=close]').click();
 await app.evaluate(({dialog},path)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[path]});},file);
 await page.locator('#open').click();await page.waitForSelector('.writer-native-running-story[data-story=footer]');
 const footer=page.locator('.writer-native-running-story[data-story=footer]');
 const sheet=page.locator('.writer-native-paper');
 assert.equal(await page.locator('[data-writer-page]').count(),3);
 assert.equal(await footer.textContent(),'Page 5 / report');
 assert.match(await sheet.textContent(),/Lead three/);
 assert.doesNotMatch(await sheet.textContent(),/Keep heading/);
 await page.locator('[data-writer-page="1"]').click();
 assert.match(await sheet.textContent(),/Keep heading/);assert.match(await sheet.textContent(),/Body last/);
 assert.equal(await footer.textContent(),'Page 6 / report');
 await page.locator('[data-writer-page="2"]').click();
 assert.match(await sheet.textContent(),/Forced page/);assert.equal(await footer.textContent(),'Page 7 / report');
 check('Actual native layout keeps heading and body together, starts forced paragraph on next page');
 // Enter text mode and change the real paragraph controls. Source export uses
 // the authorized main-process DOCX save capability, never raw test IPC.
 await revealControl(page,'#draft-text');await page.waitForSelector('#draft-text');
 const heading=page.locator('.writer-paragraph').filter({hasText:'Keep heading'});
 const input=page.locator('textarea[data-paragraph-id]').filter({hasText:'Keep heading'});
 // Textareas expose their current value via inputValue, not textContent.
 let target;
 for(const candidate of await page.locator('textarea[data-paragraph-id]').all()){
  if(await candidate.inputValue()==='Keep heading'){target=candidate;break;}
 }
 assert.ok(target,'Heading textarea available');await target.click();
 const keep=await revealControl(page,'#writer-keep-next');assert.equal(await keep.isChecked(),true);
 await keep.uncheck();await page.locator('#save').click();
 await wait(async()=>{const doc=openWriterDocx(await readFile(file),'pagination-native','Pagination').document;const b=doc.sections[0].blocks.find(v=>v.type==='paragraph'&&v.runs.some(r=>r.type==='text'&&r.text==='Keep heading'));return b?.style?.keepNext===false;},'persisted keep-next false');
 await target.click();const restored=await revealControl(page,'#writer-keep-next');await restored.check();
 await page.locator('#save').click();
 await wait(async()=>{const doc=openWriterDocx(await readFile(file),'pagination-native','Pagination').document;const b=doc.sections[0].blocks.find(v=>v.type==='paragraph'&&v.runs.some(r=>r.type==='text'&&r.text==='Keep heading'));return b?.style?.keepNext===true;},'persisted keep-next true');
 await page.waitForFunction(()=>document.querySelector('#document-status')?.textContent.includes('已保存'));
 check('Native ribbon pagination controls edit and save explicit false/true DOCX overrides');
 const before=hash(await readFile(file)),status=await page.locator('#document-status').textContent();
 await app.evaluate(({dialog},path)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:path});dialog.showMessageBox=async()=>({response:1,checkboxChecked:false});},pdf);
 await page.locator('#export-pdf').click();
 const bytes=await wait(async()=>{const b=await readFile(pdf);return b.subarray(0,5).toString()==='%PDF-'?b:false;},'real PDF export');
 assert.equal(hash(await readFile(file)),before);assert.equal(await page.locator('#document-status').textContent(),status);
 const extracted=JSON.parse(execFileSync(process.env.OFFICE_EX_PDF_PYTHON??'python3',['-c','from pypdf import PdfReader; import json,sys; r=PdfReader(sys.argv[1]); print(json.dumps([p.extract_text() for p in r.pages]))',pdf],{encoding:'utf8'}));
 assert.equal(extracted.length,3);
 for(let i=0;i<3;i++){
  assert.match(extracted[i],/Independent PDF header/);
  assert.ok(extracted[i].includes('Page '+(i+5)+' / report'),extracted[i]);
 }
 assert.match(extracted[0],/Lead three/);assert.doesNotMatch(extracted[0],/Keep heading/);
 assert.match(extracted[1],/Keep heading/);assert.match(extracted[1],/Body last/);
 assert.match(extracted[2],/Forced page/);assert.doesNotMatch(extracted[2],/Keep heading/);
 assert.deepEqual(errors,[]);
 check('Independent PDF text readback proves three pages, body placement and repeated running page numbers');
 await app.close();app=undefined;
 await mkdir('out/m6',{recursive:true});
 await writeFile('out/m6/native-pagination.json',JSON.stringify({schema:1,platform:process.platform,arch:process.arch,sandbox:true,checks,pdfSha256:hash(bytes),pages:extracted,status:'passed'},null,2)+'\n');
}finally{if(app)await app.evaluate(({app})=>app.exit()).catch(()=>{});await rm(work,{recursive:true,force:true,maxRetries:5,retryDelay:100});}
