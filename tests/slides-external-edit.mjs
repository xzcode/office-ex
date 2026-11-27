// SPDX-License-Identifier: MIT
import assert from 'node:assert/strict';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { openSlidesPptx, saveSlidesPptx, openOpc } from '@office-ex/ooxml';
const dir=process.argv[2];
const source=await readFile(join(dir,'external-python.pptx'));
const opened=openSlidesPptx(source,'external-python');
assert.equal(opened.document.slides[0].shapes[0].geometry.x,1828800);
assert.deepEqual(opened.native.metadataPaths,['docProps/core.xml']);
opened.document.slides[0].shapes[0].geometry.x+=914400;
const saved=saveSlidesPptx(source,opened.document);
assert.deepEqual(openOpc(saved).parts.get('docProps/core.xml'),openOpc(source).parts.get('docProps/core.xml'));
await writeFile(join(dir,'external-edited.pptx'),saved);
