// SPDX-License-Identifier: MIT
/** Controlled WordprocessingML profile fixture, not claimed to be Office-authored. */
import {createDocument} from '@office-ex/core';
import {openOpc,writeZip,writeWriterDocx} from '@office-ex/ooxml';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const PKG='http://schemas.openxmlformats.org/package/2006/relationships';
const enc=new TextEncoder(),dec=new TextDecoder();
export function styledFixture():Uint8Array{
  const model=createDocument('writer','fixture-style','Styled document');model.defaultTextStyle={};
  model.sections[0]!.blocks=[
    {id:'p1',type:'paragraph',runs:[{id:'r1',type:'text',text:'Styled 中文 😀'}]},
    {id:'p2',type:'paragraph',runs:[{id:'r2',type:'text',text:'Second paragraph'}]}
  ];
  const pkg=openOpc(writeWriterDocx(model)),parts=new Map(pkg.parts),main=pkg.mainPart!;
  let xml=dec.decode(parts.get(main)!);
  xml=xml.replace('<w:p><w:r><w:t','<w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Accent"/></w:rPr><w:t');
  xml=xml.replace('w:header="0" w:footer="0"','w:header="720" w:footer="720"');parts.set(main,enc.encode(xml));
  const defaults='<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorAscii" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults>';
  const styles='<w:styles xmlns:w="'+W+'">'+defaults+
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rPr><w:b/></w:rPr></w:style>'+
    '<w:style w:type="paragraph" w:styleId="Heading"><w:name w:val="Heading"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>'+
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'+
    '<w:style w:type="character" w:styleId="Accent"><w:name w:val="Accent"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:b/><w:color w:val="336699" w:themeColor="accent1"/></w:rPr></w:style>'+
    '<w:style w:type="paragraph" w:styleId="UnusedList"><w:name w:val="Unused list"/><w:pPr><w:numPr><w:numId w:val="4"/></w:numPr></w:pPr></w:style>'+
    '<w:style w:type="paragraph" w:styleId="UnusableChild"><w:name w:val="Unusable child"/><w:basedOn w:val="UnusedList"/></w:style>'+
    '</w:styles>';
  parts.set('word/styles.xml',enc.encode(styles));
  const colors=[['dk1','000000'],['lt1','FFFFFF'],['dk2','112233'],['lt2','EEEEEE'],['accent1','336699'],['accent2','CC0000'],['accent3','00CC00'],['accent4','0000CC'],['accent5','FFAA00'],['accent6','663399'],['hlink','0000FF'],['folHlink','800080']];
  const fonts=(kind:string)=>'<a:'+kind+'Font><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:'+kind+'Font>';
  const solid='<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  const lines=[6350,12700,19050].map(w=>'<a:ln w="'+w+'">'+solid+'<a:prstDash val="solid"/></a:ln>').join('');
  const fmt='<a:fmtScheme name="Fixture"><a:fillStyleLst>'+solid.repeat(3)+'</a:fillStyleLst><a:lnStyleLst>'+lines+'</a:lnStyleLst><a:effectStyleLst>'+('<a:effectStyle><a:effectLst/></a:effectStyle>').repeat(3)+'</a:effectStyleLst><a:bgFillStyleLst>'+solid.repeat(3)+'</a:bgFillStyleLst></a:fmtScheme>';
  const theme='<a:theme xmlns:a="'+A+'" name="Fixture"><a:themeElements><a:clrScheme name="Fixture">'+colors.map(([key,hex])=>'<a:'+key+'><a:srgbClr val="'+hex+'"/></a:'+key+'>').join('')+'</a:clrScheme><a:fontScheme name="Fixture">'+fonts('major')+fonts('minor')+'</a:fontScheme>'+fmt+'</a:themeElements></a:theme>';
  parts.set('word/theme/theme1.xml',enc.encode(theme));
  parts.set('word/_rels/document.xml.rels',enc.encode('<Relationships xmlns="'+PKG+'"><Relationship Id="rStyles" Type="'+R+'styles" Target="styles.xml"/><Relationship Id="rTheme" Type="'+R+'theme" Target="theme/theme1.xml"/></Relationships>'));
  const types=dec.decode(parts.get('[Content_Types].xml')!).replace('</Types>','<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>');
  parts.set('[Content_Types].xml',enc.encode(types));
  return writeZip(parts);
}
export function alterPart(source:Uint8Array,path:string,change:(xml:string)=>string):Uint8Array{
  const parts=new Map(openOpc(source).parts);const before=parts.get(path);if(!before)throw Error('Fixture part missing');
  const text=dec.decode(before),after=change(text);if(after===text)throw Error('Fixture mutation did not match');
  parts.set(path,enc.encode(after));return writeZip(parts);
}
