import {writeFileSync} from 'node:fs';
import {imageFixture} from './writer-image-fixture.ts';
import {writeWriterDocx,openWriterDocx,saveWriterDocx} from '@office-ex/ooxml';
import {DocumentSession} from '@office-ex/core';
import {resizeInlineImage} from '@office-ex/writer';
const original=writeWriterDocx(imageFixture().snapshot().document),source=openWriterDocx(original,'image-interop','Pictures'),s=new DocumentSession(source.document);
s.execute(resizeInlineImage(s.snapshot(),source.document.sections[0].blocks[1].runs[0].id,1270000));
writeFileSync(process.argv[2],original);writeFileSync(process.argv[3],saveWriterDocx(source,s.snapshot().document));
