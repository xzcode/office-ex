import {writeFileSync} from 'node:fs';
import {createDocument,DocumentSession} from '@office-ex/core';
import {writeWriterDocx,openWriterDocx,saveWriterDocx} from '@office-ex/ooxml';
import {replaceParagraphRange} from '@office-ex/writer';
const doc=createDocument('writer','interop-writer','Interop');
doc.defaultTextStyle={fontFamily:'Arial',fontSize:152400};
doc.sections[0].blocks=[
 {id:'p',type:'paragraph',style:{alignment:'center',spaceAfter:25400},runs:[{id:'r',type:'text',text:' 中文 & <Office> 😀 ',style:{bold:true}},{id:'r2',type:'text',text:'日本語\n한국어\tالعربية',style:{italic:true,underline:true,fontSize:177800}}]},
 {id:'break',type:'page-break'},
 {id:'p2',type:'paragraph',runs:[{id:'r3',type:'text',text:'e\u0301 हिन्दी русский'}]}
];
const original=writeWriterDocx(doc);
const source=openWriterDocx(original,doc.documentId,doc.title);
const session=new DocumentSession(source.document,{saved:true});
session.execute(replaceParagraphRange(session.snapshot(),source.document.sections[0].blocks[0].id,1,3,'繁體','edit'));
writeFileSync(process.argv[2],original);
writeFileSync(process.argv[3],saveWriterDocx(source,session.snapshot().document));
