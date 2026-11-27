// SPDX-License-Identifier: MIT
import assert from 'node:assert/strict';
import {readFile,writeFile} from 'node:fs/promises';
import {join} from 'node:path';
import {inspectForEditor,prepareWriterDocxSave,serializeDraft} from '@office-ex/format-service';
import {openWriterDocx,saveWriterDocx} from '@office-ex/ooxml';
import {DocumentSession,buildWriterNumberingMarkers} from '@office-ex/core';
import {replaceParagraphRange,paragraphText} from '@office-ex/writer';
const dir=process.argv[2],b=await readFile(join(dir,'ordinary-word.docx'));
const i=await inspectForEditor(b,'ordinary-word.docx','word-compat');assert.equal(i.report.readOnly,false,i.semanticFailure?.detail);
const src=openWriterDocx(b,'word-compat','Report');assert.ok(Buffer.from(saveWriterDocx(src,src.document)).equals(b));
assert.equal(src.document.sections[0].nativeImplicitMargins,true);
const marker=[...buildWriterNumberingMarkers(src.document).values()].at(-1);assert.equal(marker.text,'•');assert.equal(marker.style.fontSize,127000);assert.deepEqual(marker.style.color,{kind:'rgba',red:51,green:112,blue:255,alpha:255});
const s=new DocumentSession(src.document),p=src.document.sections[0].blocks[2];assert.equal(p.nativeMissingStyleId,'9999');assert.equal(p.type,'paragraph');assert.equal(paragraphText(s.snapshot(),p.id),'普通文档 的正文及样式。');
s.execute(replaceParagraphRange(s.snapshot(),p.id,0,4,'已编辑文档','compat-edit'));
await writeFile(join(dir,'ordinary-word-edited.docx'),prepareWriterDocxSave(serializeDraft(s.snapshot().document),'word-compat',b));
