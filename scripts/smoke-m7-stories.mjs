// SPDX-License-Identifier: MIT
/** Native variants/multiple story paragraphs, modal edit, save and PDF readback. */
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
const work=await mkdtemp(join(tmpdir(),'office-ex-m7-stories-')),profile=join(work,'profile');
await mkdir(profile);const file=join(work,'variants.docx'),pdf=join(work,'variants.pdf');
const model=createDocument('writer','m7-native');
model.defaultTextStyle={fontFamily:'Noto Sans',fontSize:127000};model.differentEvenOddPages=true;
const section=model.sections[0];
section.page={width:4572000,height:4572000,margins:{left:254000,right:254000,top:1016000,bottom:1016000}};
section.pageExtras={headerDistance:127000,footerDistance:127000};section.pageNumberStart=4;section.differentFirstPage=true;
section.storyRefs={header:'oddH',footer:'oddF',firstHeader:'firstH',firstFooter:'firstF',evenHeader:'evenH',evenFooter:'evenF'};
model.pageStories={
 oddH:{kind:'header',text:'Odd header',additionalParagraphs:[{text:'Odd second',alignment:'end'}]},
 firstH:{kind:'header',text:'First header',additionalParagraphs:[{text:'First subtitle',alignment:'center'}]},
 evenH:{kind:'header',text:'Even header'},
 oddF:{kind:'footer',text:'Odd ',pageNumber:true,additionalParagraphs:[{text:'Footer second'}]},
 firstF:{kind:'footer',text:'First ',pageNumber:true},
 evenF:{kind:'footer',text:'Even ',pageNumber:true}
};
const p=(id,text)=>({id,type:'paragraph',runs:[{id:id+'-r',type:'text',text}]});
section.blocks=[p('a','Body A'),{id:'break1',type:'page-break'},p('b','Body B'),{id:'break2',type:'page-break'},p('c','Body C')];
model.sections.push({id:'section-2',page:section.page,pageExtras:section.pageExtras,pageNumberStart:10,blocks:[p('d','Body D'),{id:'break3',type:'page-break'},p('e','Body E')]});
await writeFile(file,writeWriterDocx(model));
const checks=[],hash=b=>createHash('sha256').update(b).digest('hex');let app;
const check=name=>{checks.push(name);console.log('PASS',name);};
async function wait(fn,label,timeout=30000){const end=Date.now()+timeout;let last;while(Date.now()<end){try{const result=await fn();if(result)return result;}catch(e){last=String(e);}await new Promise(r=>setTimeout(r,100));}throw Error('Timed out: '+label+(last?'; '+last:''));}
try{
 app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env:{...process.env,OFFICE_EX_PROFILE_DIR:profile},timeout:30000});
 const page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]');const errors=[];page.on('pageerror',e=>errors.push(String(e)));
 const prefs=await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences());assert.equal(prefs.sandbox,true);assert.equal(prefs.contextIsolation,true);assert.equal(prefs.nodeIntegration,false);assert.equal(prefs.webSecurity,true);
 await page.locator('#settings').click();await page.locator('#language').selectOption('en-US');await page.locator('#settings-dialog button[value=close]').click();
 await app.evaluate(({dialog},path)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[path]});},file);
 await page.locator('#open').click();await page.waitForSelector('.writer-native-running-story[data-story=footer]');
 const header=page.locator('.writer-native-running-story[data-story=header]'),footer=page.locator('.writer-native-running-story[data-story=footer]');
 const show=async(i,heads,feet)=>{await page.locator('[data-writer-page="'+i+'"]').click();assert.deepEqual(await header.allTextContents(),heads);assert.deepEqual(await footer.allTextContents(),feet);};
 assert.equal(await page.locator('[data-writer-page]').count(),5);
 await show(0,['First header','First subtitle'],['First 4']);
 await show(1,['Odd header','Odd second'],['Odd 5','Footer second']);
 await show(2,['Even header'],['Even 6']);await show(3,['Even header'],['Even 10']);await show(4,['Odd header','Odd second'],['Odd 11','Footer second']);
 check('Native five-page rendering follows first/even/odd variants and page-number restart');
 await (await revealControl(page,'#writer-running-stories')).click();
 const dialog=page.locator('#writer-story-dialog');await page.waitForSelector('#writer-story-dialog[open]');
 await page.locator('#writer-story-section').selectOption({index:0});
 await page.locator('#writer-story-kind').selectOption('header');await page.locator('#writer-story-variant').selectOption('first');
 assert.equal(await page.locator('#writer-story-first-page').isChecked(),true);assert.equal(await page.locator('#writer-story-even-odd').isChecked(),true);
 await page.locator('#writer-story-paragraph').selectOption({value:'1'});assert.equal(await page.locator('#writer-story-text').inputValue(),'First subtitle');
 await page.locator('#writer-story-text').fill('Updated subtitle');
 await page.locator('#writer-story-add-paragraph').click();await page.locator('#writer-story-text').fill('Third line');
 await page.locator('#writer-story-apply').click();await page.waitForFunction(()=>!document.querySelector('#writer-story-dialog[open]'));
 await show(0,['First header','Updated subtitle','Third line'],['First 4']);
 await page.locator('#save').click();await wait(async()=>{const d=openWriterDocx(await readFile(file),'m7-native','Variants').document,s=d.pageStories[d.sections[0].storyRefs.firstHeader];return s.additionalParagraphs.map(p=>p.text).join('|')==='Updated subtitle|Third line';},'persisted multi-paragraph story');
 await page.waitForFunction(()=>document.querySelector('#document-status')?.textContent.includes('Saved'));
 check('Native Ribbon paragraph controls edit and append a first-page story and save DOCX');
 const before=hash(await readFile(file)),status=await page.locator('#document-status').textContent();
 await app.evaluate(({dialog},path)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:path});dialog.showMessageBox=async()=>({response:1,checkboxChecked:false});},pdf);
 await page.locator('#export-pdf').click();
 const bytes=await wait(async()=>{const b=await readFile(pdf);return b.subarray(0,5).toString()==='%PDF-'?b:false;},'real PDF output');
 const extracted=JSON.parse(execFileSync(process.env.OFFICE_EX_PDF_PYTHON??'python3',['-c','from pypdf import PdfReader; import json,sys; r=PdfReader(sys.argv[1]); print(json.dumps([p.extract_text() for p in r.pages]))',pdf],{encoding:'utf8'}));
 assert.equal(extracted.length,5);
 const expected=[['First header','Updated subtitle','Third line','First 4','Body A'],['Odd header','Odd second','Odd 5','Footer second','Body B'],['Even header','Even 6','Body C'],['Even header','Even 10','Body D'],['Odd header','Odd second','Odd 11','Footer second','Body E']];
 for(const [i,values]of expected.entries())for(const text of values)assert.ok(extracted[i].includes(text),i+': '+text+' absent from '+extracted[i]);
 assert.equal(hash(await readFile(file)),before);assert.equal(await page.locator('#document-status').textContent(),status);assert.deepEqual(errors,[]);
 check('Real five-page PDF readback matches body, all running variants and PAGE fields without modifying DOCX');
 await app.close();app=undefined;await mkdir('out/m6',{recursive:true});await writeFile('out/m6/native-m7-stories.json',JSON.stringify({schema:1,platform:process.platform,arch:process.arch,sandbox:true,checks,pdfSha256:hash(bytes),pages:extracted,status:'passed'},null,2)+'\n');
}finally{if(app)await app.evaluate(({app})=>app.exit()).catch(()=>{});await rm(work,{recursive:true,force:true,maxRetries:5,retryDelay:100});}
