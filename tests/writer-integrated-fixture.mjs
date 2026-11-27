import {writeFileSync,readFileSync} from 'node:fs';
import {createDocument,DocumentSession,makeImageResource,projectParagraph} from '@office-ex/core';
import {replaceParagraphRange,formatParagraphRange,splitWriterSectionAfter,setWriterPage,insertImageAt,insertTableAfter,locateParagraph} from '@office-ex/writer';
import {writeWriterDocx,openWriterDocx,saveWriterDocx} from '@office-ex/ooxml';
import {PNG,JPEG} from './writer-image-fixture.ts';
const [mode,input,output]=process.argv.slice(2);
if(mode==='generate'){
  const d=createDocument('writer','integrated','Integrated'),s=new DocumentSession(d),first='paragraph-1';
  s.execute(replaceParagraphRange(s.snapshot(),first,0,0,'Before after','text'));
  s.execute(formatParagraphRange(s.snapshot(),first,7,12,{language:'zh-TW',bold:true},'fmt'));
  s.execute(insertImageAt(s.snapshot(),first,7,7,makeImageResource(PNG),'pic'));
  const split=splitWriterSectionAfter(s.snapshot(),first,'second');s.execute(split.command);
  const page={width:15840*635,height:12240*635,margins:{top:720*635,right:720*635,bottom:720*635,left:720*635}};
  s.execute(setWriterPage(s.snapshot(),split.sectionId,page));
  s.execute(replaceParagraphRange(s.snapshot(),split.paragraphId,0,0,'Second section','second-text'));
  s.execute(insertTableAfter(s.snapshot(),split.paragraphId,1,2,'table'));
  const table=s.snapshot().document.sections[1].blocks[1],id=table.rows[0].cells[0].blocks[0].id;
  s.execute(replaceParagraphRange(s.snapshot(),id,0,0,'Cell content','cell'));
  s.execute(insertImageAt(s.snapshot(),id,5,5,makeImageResource(JPEG),'cell-pic'));
  writeFileSync(input,writeWriterDocx(s.snapshot().document));
}else if(mode==='edit'){
  const bytes=readFileSync(input),source=openWriterDocx(bytes,'python-authored','Independent');
  const s=new DocumentSession(source.document,{saved:true});
  const last=source.document.sections.at(-1).blocks.at(-1);
  if(last.type!=='paragraph'||projectParagraph(last)!=='Independent edit')throw Error('Independent author content not imported');
  s.execute(replaceParagraphRange(s.snapshot(),last.id,0,11,'Office EX','edit'));
  writeFileSync(output,saveWriterDocx(source,s.snapshot().document));
}else throw Error('Unknown fixture mode');
