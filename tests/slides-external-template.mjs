// SPDX-License-Identifier: MIT
import assert from 'node:assert/strict';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { openSlidesPptx, saveSlidesPptx, openOpc } from '@office-ex/ooxml';
import { inspectForEditor, prepareSlidesSave, serializeDraft, writeLegacyPpt } from '@office-ex/format-service';
import { DocumentSession, resolveSlideThemeFont } from '@office-ex/core';
import { buildSlideScene } from '@office-ex/render';
const dir=process.argv[2],source=await readFile(join(dir,'external-default.pptx'));
const opened=openSlidesPptx(source,'external-default');const model=opened.document;
assert.equal(opened.native.layoutPaths.length,11);assert.equal(opened.native.metadataPaths.length,3);assert.equal(opened.native.auxiliaryPaths.length,4);
assert.deepEqual(model.slides.map(s=>s.shapes.map(x=>x.type)),[['text','image','rect','ellipse','line'],['text','text'],['text','text']]);
assert.equal(model.slides[0].shapes[0].runs.map(r=>r.text).join(''),'External fixed 中文\u2028Soft line\nNew paragraph');
assert.equal(model.slides[0].shapes[0].runs[0].style.lineMultiple,360);
assert.equal(model.slides[1].shapes[0].geometry.x,685800);assert.equal(model.slides[1].shapes[0].runs[0].style.fontSize,558800);assert.equal(model.slides[1].shapes[0].runs[0].style.fontFamily,'+mj-lt');
assert.equal(model.slides[1].shapes[1].runs[0].style.color.tintSpace,'linear');
assert.equal(resolveSlideThemeFont('+mn-ea',model.theme,'中文','zh-CN'),model.theme.minorFontDetails.scripts.Hans);
assert.equal(resolveSlideThemeFont('+mn-ea',model.theme,'繁體','zh-TW'),model.theme.minorFontDetails.scripts.Hant);
const scene=buildSlideScene(new DocumentSession(model).snapshot(),model.slides[1].id);assert.equal(scene.shapes[0].paragraphs[0].runs[0].style.fontFamily,model.theme.majorFont);
assert.ok(Buffer.from(saveSlidesPptx(source,model)).equals(source));
const inspection=await inspectForEditor(source,'external-default.pptx','external-default');assert.equal(inspection.report.readOnly,false);assert.equal(inspection.slidesTemplateConversionBlocked,true);assert.throws(()=>writeLegacyPpt(model),/script font scheme/);
model.slides[0].shapes[0].geometry.x+=914400;model.slides[0].shapes[4].geometry.x+=914400;model.slides[1].shapes[0].runs[0].text='Edited independent title';model.slides[2].shapes[1].runs[0].text='Edited body';model.slides=[model.slides[1],model.slides[0],model.slides[2],{id:'added-slide',background:{kind:'theme',name:'bg1'},shapes:[]}];model.theme.colors.accent1={kind:'rgba',red:100,green:110,blue:120,alpha:255};
const output=prepareSlidesSave('pptx',serializeDraft(model),model.documentId,source);const saved=openSlidesPptx(output,model.documentId);assert.equal(saved.document.slides.length,4);assert.equal(saved.document.slides[0].shapes[0].runs.map(r=>r.text).join(''),'Edited independent title');assert.equal(saved.document.slides[2].shapes[1].runs.map(r=>r.text).join(''),'Edited body');assert.deepEqual(saved.document.theme.majorFontDetails,model.theme.majorFontDetails);assert.deepEqual(saved.document.theme.colors.accent1,model.theme.colors.accent1);
const oldParts=openOpc(source).parts,newParts=openOpc(output).parts;for(const path of [...opened.native.metadataPaths,...opened.native.auxiliaryPaths,opened.native.masterPath,...opened.native.layoutPaths,'_rels/.rels'])assert.deepEqual(newParts.get(path),oldParts.get(path),path);
await writeFile(join(dir,'external-default-edited.pptx'),output);
for(const file of ['external-autofit.pptx','external-bullets.pptx']){const bytes=await readFile(join(dir,file));const result=await inspectForEditor(bytes,file,'protected-external');assert.equal(result.report.readOnly,true,file);assert.throws(()=>openSlidesPptx(bytes,'protected-external'));}
console.log('Standard template: text, title/body inheritance, picture, vector shapes, connector, soft break, theme and native source preservation; 2 unsupported profiles protected.');
