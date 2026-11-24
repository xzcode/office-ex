// SPDX-License-Identifier: MIT
import {createDocument,DocumentSession,type WriterDocument,type DeepReadonly,type Table} from '@office-ex/core';
import {insertTableAfter,replaceParagraphRange,formatParagraphRange,splitParagraphRange,insertTableRow,insertTableColumn} from '@office-ex/writer';
export function tableOf(doc:DeepReadonly<WriterDocument>):DeepReadonly<Table>{
  const table=doc.sections[0]!.blocks.find(b=>b.type==='table');
  if(table?.type!=='table')throw Error('Missing table');return table;
}
export function writerModel(session:DocumentSession):DeepReadonly<WriterDocument>{
  const d=session.snapshot().document;if(d.kind!=='writer')throw Error('Expected Writer');return d;
}
export function cellParagraph(session:DocumentSession,row:number,column:number,index=0):string{
  return tableOf(writerModel(session)).rows[row]!.cells[column]!.blocks[index]!.id;
}
export function makeTableFixture():DocumentSession{
  const s=new DocumentSession(createDocument('writer','table-fixture','Table fixture'),{saved:true});
  s.execute(replaceParagraphRange(s.snapshot(),'paragraph-1',0,0,'Table example','heading'));
  s.execute(insertTableAfter(s.snapshot(),'paragraph-1',2,2,'table'));
  const p=cellParagraph(s,0,0);
  s.execute(replaceParagraphRange(s.snapshot(),p,0,0,'中文 😀','cell-a'));
  s.execute(formatParagraphRange(s.snapshot(),p,0,2,{bold:true},'bold-a'));
  s.execute(splitParagraphRange(s.snapshot(),p,2,2,'split-a','cell-a-second'));
  const b=cellParagraph(s,0,1);
  s.execute(replaceParagraphRange(s.snapshot(),b,0,0,'日本語','cell-b'));
  const c=cellParagraph(s,1,0);
  s.execute(replaceParagraphRange(s.snapshot(),c,0,0,'Second row','cell-c'));
  return s;
}
