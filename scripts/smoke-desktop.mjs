/** Runs the real Electron binary; no sandbox disabling and no product test backdoor. */
import { _electron as electron } from 'playwright-core';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, writeFile, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { createHash } from 'node:crypto';
import { docxParts, writeZip, writeWriterDocx, openWriterDocx, openSheetsXlsx, writeSheetsXlsx } from '@office-ex/ooxml';
import { createDocument } from '@office-ex/core';
import { serializeDraft, openLegacyDoc, inspectLegacyDocText, openLegacyXls } from '@office-ex/format-service';
import { styledFixture } from '../tests/writer-styled-fixture.ts';
import {PNG,JPEG} from '../tests/writer-image-fixture.ts';
import { openOpc } from '@office-ex/ooxml';

const work=await mkdtemp(join(tmpdir(),'office-ex-native-'));
const profile=join(work,'profile');await mkdir(profile);
await mkdir('out/native',{recursive:true});
const name=`${process.platform}-${process.arch}`;
const evidence={platform:process.platform,arch:process.arch,node:process.versions.node,checks:[],startedAt:new Date().toISOString()};
let app,page;
const check=title=>{evidence.checks.push(title);console.log(`PASS ${title}`);};
const originalPath=join(work,'office-source.docx'),copyPath=join(work,'office-copy.docx'),draftPath=join(work,'draft.oex'),draftCopyPath=join(work,'draft-copy.oex');
const parts=docxParts([{runs:[{text:'Independent Office EX container smoke 中文 😀'}]}]);
const original=writeZip(Object.fromEntries(Object.entries(parts).map(([key,value])=>[key,new TextEncoder().encode(value)])));
await writeFile(originalPath,original);
const bytesHash=value=>createHash('sha256').update(value).digest('hex');
const waitFor=async(predicate,label,timeout=15_000)=>{const end=Date.now()+timeout;while(Date.now()<end){if(await predicate())return;await new Promise(r=>setTimeout(r,200));}throw new Error(`Timed out: ${label}`);};
const inputEquals=(selector,value,label)=>page.waitForFunction(({selector,value})=>document.querySelector(selector)?.value===value,{selector,value},{timeout:15_000}).catch(error=>{throw new Error(`${label}: ${error.message}`);});
async function launch(){const env={...process.env,OFFICE_EX_PROFILE_DIR:profile};delete env.ELECTRON_RUN_AS_NODE;app=await electron.launch({chromiumSandbox:true,args:[...(process.env.OFFICE_EX_HEADLESS==='1'?['--headless','--ozone-platform=headless']:[]),resolve('apps/desktop')],env,timeout:30_000});page=await app.firstWindow();await page.waitForSelector('body[data-ready=true]',{timeout:30_000});await waitFor(()=>app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0]?.isVisible()===true),'native window visible');}
async function screenshot(label){
  // A ready DOM can precede the first compositor surface on Xvfb. Require an
  // actual visible window and rendered frames; retry only that capture error.
  for(let attempt=1;attempt<=3;attempt++){
    await page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));
    try{await page.screenshot({path:`out/native/${name}-${label}.png`});return;}
    catch(error){if(attempt===3||!String(error).includes('Unable to capture screenshot'))throw error;(evidence.captureRetries??=[]).push({label,attempt,reason:'Compositor surface not ready'});await new Promise(resolve=>setTimeout(resolve,250));}
  }
}
async function dialogSave(path){await app.evaluate(({dialog},filePath)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath});dialog.showMessageBox=async()=>({response:1,checkboxChecked:false});},path);}
async function dialogOpen(path){await app.evaluate(({dialog},filePath)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[filePath]});},path);await page.locator('#open').click();}
try{
  await launch();
  evidence.electron=await app.evaluate(({app})=>process.versions.electron);evidence.chromium=await page.evaluate(()=>navigator.userAgent);
  const prefs=await app.evaluate(({BrowserWindow})=>{const p=BrowserWindow.getAllWindows()[0].webContents.getLastWebPreferences();return {sandbox:p.sandbox,contextIsolation:p.contextIsolation,nodeIntegration:p.nodeIntegration,webSecurity:p.webSecurity};});
  assert.deepEqual(prefs,{sandbox:true,contextIsolation:true,nodeIntegration:false,webSecurity:true});
  assert.deepEqual(await page.evaluate(()=>[typeof window.require,typeof window.process]),['undefined','undefined']);check('Native window + sandbox, context isolation, no renderer Node');
  await page.locator('#settings').click();await page.locator('#language').selectOption('zh-CN');await page.waitForFunction(()=>document.documentElement.lang==='zh-CN');await page.locator('#settings-dialog button[value=close]').click();
  await screenshot('home');
  const denied=await page.evaluate(async()=>{const results=await Promise.allSettled([fetch('https://example.com/'),fetch('file:///etc/passwd')]);window.open('https://example.com/');return results.map(r=>r.status);});
  assert.deepEqual(denied,['rejected','rejected']);assert.equal(await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows().length),1);check('CSP rejects remote/local fetch; new windows denied');
  const invalid=await page.evaluate(()=>window.office.saveCopy({documentId:'fake',sourceToken:'fake'}));assert.equal(invalid.ok,false);assert.equal(invalid.error.code,'unsafe');
  const syntheticDrop=await page.evaluate(()=>window.office.openDroppedFile(new File(['data'],'forged.docx')));assert.equal(syntheticDrop.ok,false);check('Forged capability and synthetic dropped file rejected');
  await page.locator('[data-kind=writer]').click();await page.locator('#draft-text').fill('中文草稿 😀\nFirst saved revision');
  await dialogSave(draftPath);await page.locator('#save').click();await waitFor(async()=>{try{return JSON.parse(await readFile(draftPath,'utf8')).document.sections[0].blocks[0].runs[0].text.includes('First saved');}catch{return false;}},'first draft save');
  await page.waitForFunction(()=>document.querySelector('#document-status').textContent.includes('已保存'));
  const firstDocument=JSON.parse(await readFile(draftPath,'utf8')).document;check('Draft → worker serialization → atomic save → re-read');
  const languageBefore=await page.locator('#document-status').textContent();
  await page.locator('#settings').click();
  await page.locator('#language').selectOption('ja-JP');await page.waitForFunction(()=>document.documentElement.lang==='ja-JP');
  assert.equal(await page.locator('#draft-text').inputValue(),'中文草稿 😀\nFirst saved revision');
  await page.locator('#region').selectOption('de-DE');await page.waitForFunction(()=>document.querySelector('#region-sample')?.textContent?.includes('12.345,67'));
  const firstBatch=[['ko-KR','설정','파일'],['fr-FR','Paramètres','Fichier'],['de-DE','Einstellungen','Datei'],['es-ES','Configuración','Archivo'],['pt-BR','Configurações','Arquivo'],['ru-RU','Настройки','Файл'],['it-IT','Impostazioni','File'],['en-US','Settings','File'],['zh-TW','設定','檔案']];
  for(const [locale,settingsLabel,fileMenu] of firstBatch){
    await page.locator('#language').selectOption(locale);
    await page.waitForFunction(expected=>document.documentElement.lang===expected,locale);
    assert.equal(await page.locator('#settings-dialog h2').textContent(),settingsLabel);
    assert.ok((await app.evaluate(({Menu})=>Menu.getApplicationMenu().items.map(x=>x.label))).includes(fileMenu));
    assert.equal(await page.locator('#draft-text').inputValue(),'中文草稿 😀\nFirst saved revision');
    assert.ok(await page.locator('#settings-dialog').evaluate(el=>el.scrollWidth<=el.clientWidth+2));
  }
  check('I18N all eleven first-batch catalogs switch with localized settings and native menus');
  await page.locator('#language').selectOption('zh-CN');await page.waitForFunction(()=>document.documentElement.lang==='zh-CN');
  await page.locator('#settings-dialog button[value=close]').click();
  assert.equal(await page.locator('#draft-text').inputValue(),'中文草稿 😀\nFirst saved revision');
  assert.equal(await page.locator('#document-status').textContent(),languageBefore);
  assert.deepEqual(JSON.parse(await readFile(draftPath,'utf8')).document,firstDocument);check('I18N live language and region switching preserves draft, revision and saved state');
  await page.locator('#draft-text').fill('Changed text');await page.locator('#undo').click();await inputEquals('#draft-text','中文草稿 😀\nFirst saved revision','Writer undo rendered');await page.waitForFunction(()=>document.querySelector('#document-status').textContent.includes('已保存'));check('Undo restores saved content identity and clean state');
  // Actual DOM composition event path: one transaction for the completed input.
  const revisionBefore=await page.locator('#document-status').textContent();
  await page.locator('#draft-text').evaluate(el=>{el.dispatchEvent(new CompositionEvent('compositionstart',{bubbles:true,data:''}));el.value='中文输入法提交 😀';el.dispatchEvent(new InputEvent('input',{bubbles:true,inputType:'insertCompositionText',isComposing:true,data:'中'}));el.dispatchEvent(new CompositionEvent('compositionend',{bubbles:true,data:'中文输入法提交 😀'}));});
  assert.notEqual(await page.locator('#document-status').textContent(),revisionBefore);await page.locator('#undo').click();await inputEquals('#draft-text','中文草稿 😀\nFirst saved revision','Composition undo rendered');check('Composition commit is one undo transaction (event simulation)');
  // M2 Writer input controls and derived page preview share the same draft.
  await page.waitForFunction(()=>{const c=document.querySelector('#writer-page-preview');return c&&!c.hidden&&c.width>100;});
  await page.locator('#draft-text').evaluate(el=>{el.focus();el.setSelectionRange(0,4);el.dispatchEvent(new Event('select'));});
  await page.locator('#writer-bold').click();
  await page.waitForFunction(()=>document.querySelector('#writer-bold').getAttribute('aria-pressed')==='true');
  await page.locator('#draft-text').evaluate(el=>{el.focus();el.setSelectionRange(4,4);el.dispatchEvent(new Event('select'));});
  await page.locator('#writer-split').click();
  await page.waitForSelector('#writer-paragraph-1');
  assert.equal(await page.locator('#draft-text').inputValue(),'中文草稿');
  assert.equal(await page.locator('#writer-paragraph-1').inputValue(),' 😀\nFirst saved revision');
  await page.locator('#writer-paragraph-1').evaluate(el=>{el.focus();el.setSelectionRange(0,0);el.dispatchEvent(new Event('select'));});
  await page.locator('#writer-paragraph-1').press('Backspace');
  await page.waitForFunction(()=>document.querySelectorAll('.writer-paragraph').length===1);
  assert.equal(await page.locator('#draft-text').inputValue(),'中文草稿 😀\nFirst saved revision');
  await page.locator('#writer-page-break').click();
  await page.waitForFunction(()=>!document.querySelector('#writer-next-page').disabled);
  await page.locator('#writer-next-page').click();assert.match(await page.locator('.writer-pagebar span').textContent(),/2/);
  await page.locator('#writer-previous-page').click();
  await page.locator('#writer-page-preview').click({position:{x:50,y:50}});
  assert.equal(await page.locator('#draft-text').evaluate(el=>document.activeElement===el),true);
  // Revert the four structural/style operations, preserving the earlier saved baseline.
  for(let i=0;i<4;i++){const before=await page.locator('#document-status').textContent();await page.locator('#undo').click();await page.waitForFunction(old=>document.querySelector('#document-status').textContent!==old,before);}
  await inputEquals('#draft-text','中文草稿 😀\nFirst saved revision','Writer structural undo rendered');
  await page.waitForFunction(()=>document.querySelector('#document-status').textContent.includes('已保存'));
  check('Writer multi-paragraph controls, selected formatting, split/join, explicit page navigation and caret hit-test');
  // Each domain uses the same session and persistence contract in its own tab.
  await page.locator('#home').click();await page.locator('[data-kind=sheets]').click();
  // A real queued viewport repaint used to replace the address draft before Enter.
  await page.locator('#sheets-address').fill('B1');
  await page.evaluate(async()=>{document.querySelector('#sheets-scroll').dispatchEvent(new Event('scroll'));await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));});
  assert.equal(await page.locator('#sheets-address').inputValue(),'B1');
  await page.locator('#sheets-address').press('Enter');
  assert.equal(await page.locator('#sheets-address').inputValue(),'B1');
  check('Sheet address survives asynchronous viewport repaint');
  const enterSheet=async(address,value)=>{await page.locator('#sheets-address').fill(address);await page.locator('#sheets-address').press('Enter');assert.equal(await page.locator('#sheets-address').inputValue(),address);await page.locator('#sheets-formula').fill(value);await page.locator('#sheets-formula').press('Enter');};
  const sheetValue=(address,value)=>page.waitForFunction(({address,value})=>document.querySelector('[data-address="'+address+'"]')?.textContent===value,{address,value});
  await enterSheet('A1','单元格 中文');await enterSheet('B1','12');await enterSheet('C1','=SUM(B1:B3)');await sheetValue('C1','12');
  await page.locator('#undo').click();await sheetValue('C1','');
  await page.locator('#redo').click();await sheetValue('C1','12');
  const sheetPath=join(work,'sheets.oex');await dialogSave(sheetPath);await page.locator('#save').click();
  await waitFor(async()=>{try{return Boolean(await readFile(sheetPath));}catch{return false;}},'sheet draft save');
  const sheet=JSON.parse(await readFile(sheetPath,'utf8')).document.sheets[0];assert.equal(sheet.cells['1:1'].value.value,'单元格 中文');assert.equal(sheet.cells['1:2'].value.value,12);assert.equal(sheet.cells['1:3'].value.source,'=SUM(B1:B3)');
  await page.waitForFunction(()=>!document.querySelector('#save').disabled);await page.locator('.tab[aria-selected=true] .tab-close').click();
  await page.locator('#home').click();await page.locator('#new-xlsx').click();
  await enterSheet('A1','10');await enterSheet('B1','=A1*2');await sheetValue('B1','20');
  const xlsxPath=join(work,'native.xlsx');await dialogSave(xlsxPath);await page.locator('#save').click();
  await waitFor(async()=>{try{return openSheetsXlsx(new Uint8Array(await readFile(xlsxPath)),'native-check').document.sheets[0].cells['1:2'].value.kind==='formula';}catch{return false;}},'native XLSX save');
  await page.waitForFunction(()=>!document.querySelector('#save').disabled);await page.locator('#convert-sheets').click();
  await page.waitForFunction(()=>document.querySelector('#convert-sheets')?.textContent.includes('XLSX'));await sheetValue('B1','20');
  const xlsPath=join(work,'native.xls');await dialogSave(xlsPath);await page.locator('#save').click();
  await waitFor(async()=>{try{return openLegacyXls(new Uint8Array(await readFile(xlsPath)),'native-check').document.sheets[0].cells['1:2'].value.kind==='formula';}catch{return false;}},'native BIFF8 save');
  await page.waitForFunction(()=>!document.querySelector('#save').disabled);await page.locator('.tab[aria-selected=true] .tab-close').click();
  await page.locator('.tab[aria-selected=true] .tab-close').click();
  await dialogOpen(xlsPath);await sheetValue('B1','20');await page.locator('.tab[aria-selected=true] .tab-close').click();
  check('Spreadsheet semantic worker + guarded native XLSX save + BIFF8 conversion + reopen');
  // Imported worksheet identities must remain tied to the original package,
  // not the positional IDs obtained by reparsing the first edited output.
  const template=createDocument('sheets','native-xlsx-source','Imported workbook');
  template.sheets[0].name='First';template.sheets[0].cells['1:1']={value:{kind:'number',value:11}};
  template.sheets.push({id:'second',name:'Second',cells:{'1:1':{value:{kind:'number',value:22}},'1:2':{value:{kind:'formula',source:'=First!A1*2'}}},rows:{},columns:{},merges:[]});
  const templateParts=new Map(openOpc(writeSheetsXlsx(template)).parts),encode=value=>new TextEncoder().encode(value),decode=value=>new TextDecoder().decode(value);
  const rootRels='_rels/.rels',contentTypes='[Content_Types].xml',firstSheet='xl/worksheets/sheet1.xml',secondSheet='xl/worksheets/sheet2.xml',corePart='docProps/core.xml';
  templateParts.set(corePart,encode('<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Native source identity</dc:title></cp:coreProperties>'));
  templateParts.set(rootRels,encode(decode(templateParts.get(rootRels)).replace('</Relationships>','<Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>')));
  templateParts.set(contentTypes,encode(decode(templateParts.get(contentTypes)).replace('</Types>','<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>')));
  templateParts.set(firstSheet,encode(decode(templateParts.get(firstSheet)).replace('</worksheet>','<pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.1" footer="0.1"/></worksheet>')));
  templateParts.set(secondSheet,encode(decode(templateParts.get(secondSheet)).replace('</worksheet>','<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/></worksheet>')));
  const importPath=join(work,'native-imported.xlsx');await writeFile(importPath,writeZip(templateParts));await dialogOpen(importPath);
  await page.waitForFunction(()=>document.querySelector('.tab[aria-selected=true] .tab-title')?.textContent==='native-imported.xlsx');
  await sheetValue('A1','11');assert.equal(await page.locator('#convert-sheets').isDisabled(),true);
  await page.locator('.sheets-advanced summary').click();await page.locator('#sheet-move-right').click();
  await page.waitForFunction(()=>Array.from(document.querySelectorAll('.sheet-tabs [role=tab]')).map(el=>el.textContent).join(',')==='Second,First');
  await page.locator('#save').click();
  const readImported=async(expected)=>{
    const result=openSheetsXlsx(await readFile(importPath),'native-import-check');
    assert.deepEqual(result.document.sheets.map(s=>s.name),['Second','First']);
    assert.deepEqual(result.native.sheetParts.map(s=>s.path),[secondSheet,firstSheet]);
    assert.equal(result.document.sheets[1].cells['1:1'].value.value,expected);
    assert.ok(result.cachedValues.some(([key,value])=>key==='sheet-1!1:2'&&value.kind==='number'&&value.value===expected*2));
    assert.deepEqual(result.native.opc.parts.get(corePart),templateParts.get(corePart));
    assert.match(decode(result.native.opc.parts.get(firstSheet)),/left="0.25"/);
    assert.match(decode(result.native.opc.parts.get(secondSheet)),/left="0.75"/);
    return true;
  };
  await waitFor(async()=>{try{return await readImported(11);}catch{return false;}},'first imported XLSX native save');
  await page.waitForFunction(()=>!document.querySelector('#save').disabled);
  await enterSheet('A1','33');await sheetValue('A1','33');await page.locator('#save').click();
  await waitFor(async()=>{try{return await readImported(33);}catch{return false;}},'second imported XLSX native save');
  await page.waitForFunction(()=>!document.querySelector('#save').disabled);await page.locator('.tab[aria-selected=true] .tab-close').click();
  check('Consecutive imported XLSX saves retain stable worksheet paths, metadata, print margins and formula caches after reorder');

  await page.locator('#home').click();await page.locator('[data-kind=slides]').click();await page.locator('#slide-insert-rect').click();
  const shapes=page.locator('#slide-canvas [data-shape-id]');assert.equal(await shapes.count(),1);await page.locator('#undo').click();await waitFor(async()=>await shapes.count()===0,'Slide undo rendered');await page.locator('#redo').click();await waitFor(async()=>await shapes.count()===1,'Slide redo rendered');
  await shapes.first().click();await screenshot('slides');
  await page.locator('.tab[aria-selected=true] .tab-close').click();await page.locator('.tab').first().click();
  assert.equal(await page.locator('#draft-text').inputValue(),'中文草稿 😀\nFirst saved revision');check('Three independent draft tabs: typed cells/formula source, scene shape/undo, isolated histories');
  // Save As holds a native-dialog promise while a later edit arrives in the renderer.
  await app.evaluate(({dialog},filePath)=>{dialog.showSaveDialog=()=>new Promise(resolve=>{globalThis.__officeSmokeRelease=()=>resolve({canceled:false,filePath});});},draftCopyPath);
  await page.locator('#save-as').click();await waitFor(async()=>app.evaluate(()=>typeof globalThis.__officeSmokeRelease==='function'),'save snapshot reached dialog');
  await page.locator('#draft-text').fill('Edit made while saving');await app.evaluate(()=>globalThis.__officeSmokeRelease());
  await waitFor(async()=>{try{return Boolean(await readFile(draftCopyPath));}catch{return false;}},'save-as output');
  await page.waitForFunction(()=>!document.querySelector('#save').disabled);
  assert.equal(JSON.parse(await readFile(draftCopyPath,'utf8')).document.sections[0].blocks[0].runs[0].text,'中文草稿 😀\nFirst saved revision');
  assert.match(await page.locator('#document-status').textContent(),/未保存/);check('Concurrent edit remains dirty after snapshot save');
  // Existing target save requires and creates a previous-version private backup.
  await dialogSave(draftCopyPath);await page.locator('#save').click();await waitFor(async()=>JSON.parse(await readFile(draftCopyPath,'utf8')).document.sections[0].blocks[0].runs[0].text==='Edit made while saving','overwrite');
  assert((await readdir(join(profile,'backups'))).some(name=>name.endsWith('.oex-recovery')));check('Overwrite preserves previous version in private backups');
  await dialogOpen(originalPath);await page.waitForSelector('.parts-table');assert.match(await page.locator('#content').textContent(),/DOCX/);assert.match(await page.locator('#content').textContent(),/尚未导入/);
  await dialogSave(copyPath);await page.locator('#save').click();await waitFor(async()=>{try{return (await readFile(copyPath)).length>0;}catch{return false;}},'original copy');assert.equal(bytesHash(await readFile(copyPath)),bytesHash(original));assert.equal(bytesHash(await readFile(originalPath)),bytesHash(original));check('DOCX container inspection + byte-exact Save As with source unchanged');
  await screenshot('inspection');
  // A valid file token alone cannot bypass the unsupported DOCX profile.
  const unauthorized=await page.evaluate(async()=>{
    const opened=await window.office.openFile();
    if(!opened.ok||!opened.value)throw Error('fixture open failed');
    return {documentId:opened.value.documentId,sourceToken:opened.value.token};
  });
  const blockedModel=createDocument('writer',unauthorized.documentId,'Blocked');
  const blocked=await page.evaluate(async({request,data})=>window.office.saveDocx({...request,bytes:Uint8Array.from(data)}),{
    request:{...unauthorized,revision:0,contentId:unauthorized.documentId+':content:0',saveAs:false,title:'Blocked'},
    data:Array.from(serializeDraft(blockedModel))
  });
  assert.equal(blocked.ok,false);assert.equal(blocked.error.code,'unsupported');
  assert.equal(bytesHash(await readFile(originalPath)),bytesHash(original));
  assert.equal(await page.evaluate(()=>typeof window.docxVerification),'undefined');
  check('Host verifier rejects semantic overwrite of unsupported DOCX even with an authorized source token');

  // Create a real DOCX through the native editor and independent save verifier.
  const docxPath=join(work,'editable.docx'),docxCopyPath=join(work,'editable-copy.docx');
  const docxText=async path=>{
    const bytes=await readFile(path);
    const model=openWriterDocx(bytes,'test-reader','Verification').document;
    return model.sections[0].blocks.filter(b=>b.type==='paragraph').map(p=>p.runs.filter(r=>r.type==='text').map(r=>r.text).join('')).join('\n');
  };
  await page.locator('#home').click();await page.locator('#new-docx').click();
  await inputEquals('#draft-text','','New native DOCX editor');
  await page.locator('#draft-text').fill('Writer DOCX 中文 😀');
  await page.locator('#draft-text').evaluate(el=>{el.focus();el.setSelectionRange(0,6);el.dispatchEvent(new Event('select'));});
  await page.locator('#writer-font-family').fill('Arial');await page.locator('#writer-font-family').press('Tab');
  await page.locator('#writer-font-size').fill('14.5');await page.locator('#writer-font-size').press('Tab');
  await page.locator('#writer-text-color').fill('#336699');
  await dialogSave(docxPath);await page.locator('#save').click();
  await waitFor(async()=>{try{return await docxText(docxPath)==='Writer DOCX 中文 😀';}catch{return false;}},'native DOCX creation',30000);
  await page.waitForFunction(()=>!document.querySelector('#save').disabled&&document.querySelector('#document-status').textContent.includes('已保存'));
  const firstDocx=await readFile(docxPath);
  const firstRun=openWriterDocx(firstDocx,'test-reader','Verification').document.sections[0].blocks[0].runs[0];
  assert.equal(firstRun.text,'Writer');assert.equal(firstRun.style.fontFamily,'Arial');
  assert.equal(firstRun.style.fontSize,184150);assert.deepEqual(firstRun.style.color,{kind:'rgba',red:51,green:102,blue:153,alpha:255});
  check('Writer font, half-point size and RGB toolbar edits survive real DOCX save');
  assert.equal(await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows().length),1);
  await screenshot('writer-docx');
  check('New DOCX commits a real independently readable Office archive through an isolated host verifier');

  await dialogSave(docxCopyPath);await page.locator('#save-as').click();
  await waitFor(async()=>{try{return bytesHash(await readFile(docxCopyPath))===bytesHash(firstDocx);}catch{return false;}},'unchanged DOCX Save As',30000);
  await page.waitForFunction(()=>!document.querySelector('#save').disabled);
  assert.equal(bytesHash(await readFile(docxPath)),bytesHash(firstDocx));
  const beforeClose=await page.locator('.tab').count();
  await page.locator('.tab[aria-selected=true] .tab-close').click();
  await page.waitForFunction(count=>document.querySelectorAll('.tab').length===count-1,beforeClose);
  await dialogOpen(docxCopyPath);await inputEquals('#draft-text','Writer DOCX 中文 😀','Reopened native DOCX');
  const backupCount=(await readdir(join(profile,'backups'))).length;
  await page.locator('#draft-text').fill('Edited DOCX 中文');
  await page.locator('#save').click();
  await waitFor(async()=>{try{return await docxText(docxCopyPath)==='Edited DOCX 中文';}catch{return false;}},'edited DOCX native save',30000);
  await page.waitForFunction(()=>!document.querySelector('#save').disabled&&document.querySelector('#document-status').textContent.includes('已保存'));
  assert.ok((await readdir(join(profile,'backups'))).length>backupCount);
  assert.equal(bytesHash(await readFile(docxPath)),bytesHash(firstDocx));
  check('DOCX Save As is byte-exact when unchanged; reopen/edit/overwrite preserves the source and backup');

  // An external writer changed the authorized original: commit must fail closed.
  const externalModel=createDocument('writer','external','External');
  externalModel.sections[0].blocks=[{id:'p',type:'paragraph',runs:[{id:'r',type:'text',text:'External change'}]}];
  const external=writeWriterDocx(externalModel);
  await page.locator('#draft-text').fill('Must not replace the external change');
  await writeFile(docxCopyPath,external);
  await page.locator('#save').click();await page.waitForSelector('#message.error',{timeout:30000});
  assert.equal(bytesHash(await readFile(docxCopyPath)),bytesHash(external));
  assert.match(await page.locator('#document-status').textContent(),/未保存/);
  await app.evaluate(({dialog})=>{dialog.showMessageBox=async()=>({response:1,checkboxChecked:false});});
  const beforeDiscard=await page.locator('.tab').count();
  await page.locator('.tab[aria-selected=true] .tab-close').click();
  await page.waitForFunction(count=>document.querySelectorAll('.tab').length===count-1,beforeDiscard);
  check('External DOCX change blocks native overwrite and retains unsaved model');

  // Native open and apply supported named styles while retaining source styles/theme.
  const styledPath=join(work,'styled.docx'),styledCopy=join(work,'styled-copy.docx');
  const styledBytes=styledFixture();await writeFile(styledPath,styledBytes);
  await dialogOpen(styledPath);await inputEquals('#draft-text','Styled 中文 😀','Open styled native DOCX');
  await page.locator('#draft-text').evaluate(el=>{el.focus();el.setSelectionRange(0,6);el.dispatchEvent(new Event('select'));});
  assert.equal(await page.locator('#writer-paragraph-style').inputValue(),'Heading');
  assert.equal(await page.locator('#writer-character-style').inputValue(),'Accent');
  await page.locator('#writer-paragraph-style').selectOption('Normal');
  await page.locator('#draft-text').fill('Modified styled text');
  await dialogSave(styledCopy);await page.locator('#save-as').click();
  await waitFor(async()=>{try{return (await docxText(styledCopy)).startsWith('Modified styled text');}catch{return false;}},'styled DOCX Save As',30000);
  await page.waitForFunction(()=>!document.querySelector('#save').disabled&&document.querySelector('#document-status').textContent.includes('已保存'));
  const styledOut=await readFile(styledCopy),originalOpc=openOpc(styledBytes),editedOpc=openOpc(styledOut);
  for(const [name,data] of originalOpc.parts)if(name!==originalOpc.mainPart)assert.equal(bytesHash(editedOpc.parts.get(name)),bytesHash(data));
  const styledModel=openWriterDocx(styledOut,'test-reader','Verification').document;
  assert.equal(styledModel.sections[0].blocks[0].styleId,'Normal');assert.equal(styledModel.styles.Accent.kind,'character');
  assert.equal(bytesHash(await readFile(styledPath)),bytesHash(styledBytes));
  await screenshot('writer-styled-docx');
  const styledCount=await page.locator('.tab').count();await page.locator('.tab[aria-selected=true] .tab-close').click();
  await page.waitForFunction(count=>document.querySelectorAll('.tab').length===count-1,styledCount);
  check('Styled DOCX opens in Writer, applies named paragraph style and preserves every non-main part');

  // Table editing goes through the real desktop editor and source-preserving save.
  const tablePath=join(work,'table.docx');
  await page.locator('#home').click();await page.locator('#new-docx').click();
  await inputEquals('#draft-text','','New table document');
  await page.locator('#draft-text').fill('Table example');
  await page.locator('#writer-insert-table').click();
  await page.waitForFunction(()=>document.querySelectorAll('.writer-edit-table tbody tr').length===2);
  const cell=(row,column)=>page.locator('.writer-cell-input[data-row="'+row+'"][data-column="'+column+'"]').first();
  await cell(0,0).fill('中文 表格 😀');
  await cell(0,1).fill('日本語');
  await cell(1,0).fill('Second row');
  await cell(0,0).click();
  await page.locator('#writer-table-add-row').click();
  await page.waitForFunction(()=>document.querySelectorAll('.writer-edit-table tbody tr').length===3);
  await cell(0,0).click();
  await page.locator('#writer-table-add-column').click();
  await page.waitForFunction(()=>document.querySelectorAll('.writer-edit-table tbody tr:first-child td').length===3);
  await cell(0,1).fill('Inserted column');
  await cell(1,0).fill('Inserted row');
  await page.locator('#writer-table-delete-row').click();
  await page.waitForFunction(()=>document.querySelectorAll('.writer-edit-table tbody tr').length===2);
  await page.locator('#undo').click();
  await page.waitForFunction(()=>document.querySelectorAll('.writer-edit-table tbody tr').length===3);
  assert.equal(await cell(1,0).inputValue(),'Inserted row');
  await cell(0,0).click();
  await page.waitForFunction(()=>{const c=document.querySelector('#writer-page-preview');return c&&!c.hidden&&c.width>100;});
  await dialogSave(tablePath);await page.locator('#save').click();
  await waitFor(async()=>{try{return (await readFile(tablePath)).length>100;}catch{return false;}},'real DOCX table save',30000);
  await page.waitForFunction(()=>!document.querySelector('#save').disabled&&document.querySelector('#document-status').textContent.includes('已保存'));
  const tableBytes=await readFile(tablePath);
  const tableModel=openWriterDocx(tableBytes,'table-reader','Table').document;
  const table=tableModel.sections[0].blocks.find(b=>b.type==='table');
  assert.equal(table.rows.length,3);assert.equal(table.columnWidths.length,3);
  const cellText=(r,c)=>table.rows[r].cells[c].blocks.map(p=>p.runs.map(run=>run.type==='text'?run.text:'').join('')).join('\n');
  assert.equal(cellText(0,0),'中文 表格 😀');assert.equal(cellText(0,1),'Inserted column');
  assert.equal(cellText(0,2),'日本語');assert.equal(cellText(1,0),'Inserted row');assert.ok(table.border);
  await screenshot('writer-table-docx');
  const tableCount=await page.locator('.tab').count();await page.locator('.tab[aria-selected=true] .tab-close').click();
  await page.waitForFunction(count=>document.querySelectorAll('.tab').length===count-1,tableCount);
  await dialogOpen(tablePath);await page.waitForSelector('.writer-edit-table');
  assert.equal(await cell(0,0).inputValue(),'中文 表格 😀');assert.equal(await cell(0,1).inputValue(),'Inserted column');
  const reopenCount=await page.locator('.tab').count();await page.locator('.tab[aria-selected=true] .tab-close').click();
  await page.waitForFunction(count=>document.querySelectorAll('.tab').length===count-1,reopenCount);
  check('Writer table creates rows and columns, edits cells, undoes row deletion, saves and reopens a real DOCX');

  const imagePath=join(work,'images.docx');
  await page.locator('#home').click();await page.locator('#new-docx').click();
  await inputEquals('#draft-text','','New image document');
  await page.locator('#draft-text').fill('Image example');
  await page.locator('#writer-paper').selectOption('letter');
  await page.locator('#writer-orientation').selectOption('landscape');
  await page.locator('#writer-apply-page').click();
  await page.locator('#draft-text').click();
  await page.locator('#writer-spaceAfter').fill('6');await page.locator('#writer-spaceAfter').press('Tab');
  await page.locator('#writer-image-file').setInputFiles({name:'example.png',mimeType:'image/png',buffer:PNG});
  await page.waitForSelector('.writer-image-card img',{timeout:30000});
  await page.locator('#writer-image-file').setInputFiles({name:'example.jpeg',mimeType:'image/jpeg',buffer:JPEG});
  await page.waitForFunction(()=>document.querySelectorAll('.writer-image-card img').length===2);
  await page.locator('.writer-image-card').first().locator('.writer-image-width').fill('72');
  await page.locator('.writer-image-card').first().locator('.writer-image-width').press('Tab');
  await page.waitForFunction(()=>document.querySelector('#writer-page-preview')?.width>100&&!document.querySelector('#writer-page-preview')?.hidden);
  await dialogSave(imagePath);await page.locator('#save').click();
  await waitFor(async()=>{try{return (await readFile(imagePath)).length>100;}catch{return false;}},'DOCX image save',30000);
  await page.waitForFunction(()=>!document.querySelector('#save').disabled&&document.querySelector('#document-status').textContent.includes('已保存'));
  const imageDoc=openWriterDocx(await readFile(imagePath),'image-reader','Images').document;
  assert.equal(Object.keys(imageDoc.resources).length,2);assert.ok(imageDoc.sections[0].page.width>imageDoc.sections[0].page.height);
  const pictures=imageDoc.sections[0].blocks.flatMap(b=>b.type==='paragraph'?b.runs.filter(r=>r.type==='image'):[]);
  assert.equal(pictures.length,2);assert.equal(pictures[0].width,914400);
  assert.ok(Object.values(imageDoc.resources).some(r=>r.sha256===bytesHash(PNG)));
  assert.ok(Object.values(imageDoc.resources).some(r=>r.sha256===bytesHash(JPEG)));
  await screenshot('writer-images-docx');
  const imageCount=await page.locator('.tab').count();await page.locator('.tab[aria-selected=true] .tab-close').click();
  await page.waitForFunction(n=>document.querySelectorAll('.tab').length===n-1,imageCount);
  await dialogOpen(imagePath);await page.waitForFunction(()=>document.querySelectorAll('.writer-image-card img').length===2);
  await page.locator('.writer-image-card').first().locator('button').last().click();
  await page.waitForFunction(()=>document.querySelectorAll('.writer-image-card').length===1);
  await page.locator('#undo').click();await page.waitForFunction(()=>document.querySelectorAll('.writer-image-card').length===2);
  const reopenedImages=await page.locator('.tab').count();await page.locator('.tab[aria-selected=true] .tab-close').click();
  await page.waitForFunction(n=>document.querySelectorAll('.tab').length===n-1,reopenedImages);
  check('Writer inserts bounded PNG/JPEG, resizes, saves landscape DOCX, reopens and undoes removal');

  // Native continuous input edits a real range spanning multiple paragraphs.
  const continuousPath=join(work,'continuous.docx');
  await page.locator('#home').click();await page.locator('#new-docx').click();
  await inputEquals('#draft-text','','New continuous document');
  await page.locator('#draft-text').fill('First 中文');
  await page.locator('#draft-text').evaluate(el=>{el.focus();el.setSelectionRange(el.value.length,el.value.length);el.dispatchEvent(new Event('select'));});
  await page.locator('#writer-split').click();
  await page.locator('#writer-paragraph-1').fill('Second 😀');
  await page.locator('#draft-text').click();
  await page.locator('#writer-continuous').click();
  await inputEquals('#writer-continuous-text','First 中文\nSecond 😀','Continuous paragraphs');
  const continuous=page.locator('#writer-continuous-text');
  await continuous.evaluate(el=>{el.focus();el.setSelectionRange(6,15);el.dispatchEvent(new Event('select'));});
  await page.locator('#writer-continuous-bold').click();
  await continuous.evaluate(el=>{el.focus();el.setSelectionRange(6,15);el.dispatchEvent(new Event('select'));});
  await continuous.press('Backspace');
  await inputEquals('#writer-continuous-text','First  😀','Cross-paragraph deletion');
  await page.locator('#undo').click();
  await inputEquals('#writer-continuous-text','First 中文\nSecond 😀','Cross-paragraph undo');
  await page.locator('#writer-continuous').click();
  await inputEquals('#draft-text','First 中文','Return to paragraph mode');
  // Add an image inside a text run using the native file input.
  await page.locator('#draft-text').evaluate(el=>{el.focus();el.setSelectionRange(6,6);el.dispatchEvent(new Event('select'));});
  const chooserPromise=page.waitForEvent('filechooser');
  await page.locator('#writer-insert-inline-image').click();
  await (await chooserPromise).setFiles({name:'inline.png',mimeType:'image/png',buffer:PNG});
  await page.waitForSelector('.writer-image-card img',{timeout:30000});
  await inputEquals('#draft-text','First \uFFFC中文','Inline object projection');
  // Set explicit document text metadata and create an independent next-page section.
  await page.locator('#draft-text').evaluate(el=>{el.focus();el.setSelectionRange(7,9);el.dispatchEvent(new Event('select'));});
  await page.locator('#writer-text-language').fill('zh-TW');await page.locator('#writer-text-language').press('Tab');
  await page.locator('#writer-section-break').click();
  await page.waitForFunction(()=>document.querySelector('#writer-section')?.options.length===2);
  await page.locator('#writer-paper').selectOption('letter');await page.locator('#writer-orientation').selectOption('landscape');await page.locator('#writer-apply-page').click();
  await dialogSave(continuousPath);await page.locator('#save').click();
  await waitFor(async()=>{try{return (await readFile(continuousPath)).length>100;}catch{return false;}},'mixed DOCX save',30000);
  await page.waitForFunction(()=>!document.querySelector('#save').disabled&&document.querySelector('#document-status').textContent.includes('已保存'));
  const mixed=openWriterDocx(await readFile(continuousPath),'continuous-reader','Continuous').document;
  assert.equal(mixed.sections.length,2);assert.ok(mixed.sections[0].page.width<mixed.sections[0].page.height&&mixed.sections[1].page.width>mixed.sections[1].page.height);
  const first=mixed.sections[0].blocks[0];assert.equal(first.runs.filter(r=>r.type==='image').length,1);
  assert.equal(first.runs.filter(r=>r.type==='text').map(r=>r.text).join(''),'First 中文');
  assert.ok(first.runs.some(r=>r.type==='text'&&r.text==='中文'&&r.style.bold&&r.style.language==='zh-TW'));
  await screenshot('writer-inline-range');
  const continuousCount=await page.locator('.tab').count();await page.locator('.tab[aria-selected=true] .tab-close').click();
  await page.waitForFunction(n=>document.querySelectorAll('.tab').length===n-1,continuousCount);
  check('Native continuous selection, undo, inline image, section geometry and text language DOCX save');

  // Page surface uses Chromium's native shaping geometry and an IME bridge.
  const pagedPath=join(work,'paged.docx');
  await page.locator('#home').click();await page.locator('#new-docx').click();
  const complex='English 中文 العربية עברית हिन्दी ไทย 😀 e\u0301';
  await page.locator('#draft-text').fill(complex);
  await page.locator('#writer-page-edit').click();
  await page.waitForSelector('#writer-native-page:not([hidden]) .writer-native-slice');
  const pageInput=page.locator('#writer-page-input');
  await page.waitForSelector('#writer-page-input:not([hidden])');
  await pageInput.focus();
  await page.keyboard.insertText('测');
  await page.waitForFunction(()=>document.querySelector('#document-status')?.textContent.includes('修订 2')||document.querySelector('#writer-native-page')?.textContent.includes('测'));
  assert.ok((await page.locator('#writer-native-page').textContent()).includes('测'));
  await screenshot('writer-native-multilingual');
  await dialogSave(pagedPath);await page.locator('#save').click();
  await waitFor(async()=>{try{return openWriterDocx(await readFile(pagedPath),'page-native','Page').document.sections[0].blocks[0].runs.map(r=>r.type==='text'?r.text:'').join('').includes('测');}catch{return false;}},'native page edit DOCX save',30000);
  await page.waitForFunction(()=>document.querySelector('#document-status').textContent.includes('已保存'));
  await page.locator('.tab[aria-selected=true] .tab-close').click();
  check('Native page surface shapes multilingual text and commits input to DOCX');

  // Exercise the new format controls, native line projection and source save.
  const wordM5Path=join(work,'word-m5.docx');
  await page.locator('#home').click();await page.locator('#new-docx').click();
  const advancedText='Character spacing and hanging indent';
  const draft=page.locator('#draft-text');await draft.fill(advancedText);
  const chooseText=()=>draft.evaluate(el=>{el.focus();el.setSelectionRange(0,9);el.dispatchEvent(new Event('select'));});
  await chooseText();await page.locator('#writer-character-spacing').fill('0.25');await page.locator('#writer-character-spacing').press('Tab');
  await chooseText();await page.locator('#writer-kerning-threshold').fill('14');await page.locator('#writer-kerning-threshold').press('Tab');
  await page.locator('#writer-indent').fill('36');await page.locator('#writer-indent').press('Tab');
  await page.locator('#writer-rightIndent').fill('18');await page.locator('#writer-rightIndent').press('Tab');
  await page.locator('#writer-firstLineIndent').fill('-18');await page.locator('#writer-firstLineIndent').press('Tab');
  await page.locator('#writer-contextual-spacing').check();
  await page.locator('#writer-bottom-rule').check();
  await page.locator('#writer-page-edit').click();
  const rule=page.locator('#writer-native-page .writer-paragraph-bottom-rule');
  await rule.waitFor();
  assert.equal(await rule.evaluate(el=>getComputedStyle(el).backgroundColor),'rgb(79, 129, 189)');
  const rendered=await page.locator('#writer-native-page .writer-native-paragraph').first().evaluate(el=>({indent:parseFloat(getComputedStyle(el).textIndent),pad:parseFloat(getComputedStyle(el).paddingLeft),text:el.textContent,spacing:getComputedStyle(el.querySelector('span')).letterSpacing}));
  assert.equal(rendered.indent,-24);assert.equal(rendered.pad,24);assert.ok(rendered.text.includes(advancedText));assert.ok(parseFloat(rendered.spacing)>0);
  await dialogSave(wordM5Path);await page.locator('#save').click();
  await waitFor(async()=>{try{
   const p=openWriterDocx(await readFile(wordM5Path),'m5-word-ui','M5 Word').document.sections[0].blocks[0];
   assert.equal(p.style.indent,457200);assert.equal(p.style.rightIndent,228600);assert.equal(p.style.firstLineIndent,-228600);assert.equal(p.style.contextualSpacing,true);assert.equal(p.style.paragraphBottomBorder.width,12700);assert.equal(p.runs[0].style.characterSpacing,3175);assert.equal(p.runs[0].style.kerningThreshold,177800);return true;
  }catch{return false;}},'M5 Writer native formatting save',30000);
  await page.waitForFunction(()=>document.querySelector('#document-status').textContent.includes('已保存'));await page.locator('.tab[aria-selected=true] .tab-close').click();
  check('Writer advanced spacing, kerning, hanging indentation and bottom rule controls persist to native DOCX');

  // Native DOC save uses the same isolated verifier and commit capability.
  const legacyPath=join(work,'writer-legacy.doc'),convertedPath=join(work,'writer-converted.docx');
  await page.locator('#home').click();
  await page.locator('#settings').click();await page.locator('#document-language').selectOption('zh-CN');await page.locator('#settings-dialog button[value=close]').click();
  await page.locator('#home').click();await page.locator('#new-doc').click();
  await page.locator('#draft-text').fill('Native DOC 中文');
  await dialogSave(legacyPath);await page.locator('#save').click();
  await waitFor(async()=>{try{return inspectLegacyDocText(await readFile(legacyPath)).text==='Native DOC 中文\r';}catch{return false;}},'native DOC save',30000);
  await page.waitForFunction(()=>document.querySelector('#document-status').textContent.includes('已保存'));
  await page.locator('.tab[aria-selected=true] .tab-close').click();
  await dialogOpen(legacyPath);await inputEquals('#draft-text','Native DOC 中文','native DOC reopen');
  const beforeConversion=await page.locator('.tab').count();await page.locator('#convert-writer').click();
  await page.waitForFunction(n=>document.querySelectorAll('.tab').length===n+1,beforeConversion);
  await inputEquals('#draft-text','Native DOC 中文','conversion model');
  await dialogSave(convertedPath);await page.locator('#save').click();
  await waitFor(async()=>{try{return openWriterDocx(await readFile(convertedPath),'converted-native','Converted').document.sections[0].blocks[0].runs[0].text==='Native DOC 中文';}catch{return false;}},'native DOC to DOCX conversion',30000);
  await page.waitForFunction(()=>document.querySelector('#document-status').textContent.includes('已保存'));
  await page.locator('.tab[aria-selected=true] .tab-close').click();
  await page.locator('.tab[aria-selected=true] .tab-close').click();
  check('Native DOC creation, verified save, reopen and conversion to DOCX');

  await page.locator('.tab').first().click();assert.equal(await page.locator('#draft-text').inputValue(),'Edit made while saving');
  // Bad save destination keeps the document dirty, usable and the old file unchanged.
  const stable=await readFile(draftCopyPath);await page.locator('#draft-text').fill('Dirty content before simulated crash');
  await dialogSave(join(work,'wrong.docx'));await page.locator('#save-as').click();await page.waitForSelector('#message.error');assert.equal(bytesHash(await readFile(draftCopyPath)),bytesHash(stable));assert.match(await page.locator('#document-status').textContent(),/未保存/);check('Wrong-extension draft save fails without changing source or dirty state');
  // Wait for the real 30-second checkpoint interval instead of calling a test-only API.
  await waitFor(async()=>{try{for(const name of await readdir(join(profile,'recovery'))){if(!name.endsWith('.oex-recovery'))continue;const b=await readFile(join(profile,'recovery',name));const doc=JSON.parse(b.subarray(4+b.readUInt32LE(0)).toString()).document;if(doc.sections?.[0]?.blocks?.[0]?.runs?.[0]?.text==='Dirty content before simulated crash')return true;}return false;}catch{return false;}},'automatic recovery checkpoint',40_000);check('Real 30-second automatic recovery checkpoint');
  await page.locator('#settings').click();await page.locator('#theme').selectOption('dark');await page.locator('#language').selectOption('en-US');await page.waitForFunction(()=>document.documentElement.lang==='en-US');await page.locator('#settings-dialog button[value=close]').click();
  await screenshot('draft-dark');
  // Hard process exit leaves the checkpoint for the next launch, bypassing normal close UI.
  await app.evaluate(({app})=>app.exit(0));await app.close().catch(()=>undefined);app=undefined;
  await launch();assert.equal(await page.locator('html').getAttribute('data-theme'),'dark');assert.equal(await page.locator('html').getAttribute('lang'),'en-US');
  await page.getByRole('button',{name:'Recover a copy',exact:true}).first().click();await page.waitForSelector('#draft-text');assert.equal(await page.locator('#draft-text').inputValue(),'Dirty content before simulated crash');assert.match(await page.locator('#document-status').textContent(),/Unsaved/);assert.equal(bytesHash(await readFile(draftCopyPath)),bytesHash(stable));check('Crash recovery reopens an unsaved copy; settings persist; source remains unchanged');
  await app.evaluate(({dialog})=>{dialog.showMessageBox=async()=>({response:2,checkboxChecked:false});});await page.locator('.tab-close').click();assert.equal(await page.locator('.tab').count(),1);check('Dirty close cancellation keeps document open');
  await app.evaluate(({dialog})=>{dialog.showMessageBox=async()=>({response:1,checkboxChecked:false});});await page.locator('.tab-close').click();await page.waitForFunction(()=>document.querySelectorAll('.tab').length===0);check('Confirmed discard closes the session');
  evidence.status='passed';evidence.finishedAt=new Date().toISOString();
}catch(error){evidence.status='failed';evidence.error=String(error?.stack??error);console.error(evidence.error);if(page)await screenshot('failure').catch(()=>undefined);process.exitCode=1;}
finally{await writeFile(`out/native/${name}.json`,JSON.stringify(evidence,null,2)+'\n');if(app)await app.evaluate(({app})=>app.exit(process.exitCode?1:0)).catch(()=>undefined);if(app)await app.close().catch(()=>undefined);await rm(work,{recursive:true,force:true});}
