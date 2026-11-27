import {createDocument,type SheetsDocument} from '@office-ex/core';
export function fixture():SheetsDocument{
  const d=createDocument('sheets','sheet-fixture','Test');d.dateSystem='1904';
  d.sheets[0].cells={'1:1':{value:{kind:'number',value:12}},'1:2':{value:{kind:'string',value:'中文 Excel'}},'2:1':{value:{kind:'formula',source:'=SUM(A1,3)'}},'2:2':{value:{kind:'boolean',value:true}},'3:1':{value:{kind:'formula',source:'="你好"'}},'4:1':{value:{kind:'number',value:123.45},numberFormat:'0.00',style:{bold:true,italic:true,underline:true,alignment:'center',wrapText:true,fill:{kind:'rgba',red:255,green:200,blue:100,alpha:255},border:{width:9525,color:{kind:'rgba',red:10,green:20,blue:30,alpha:255}}}},'6:1':{value:{kind:'error',value:'#DIV/0!'}},'7:1':{value:{kind:'formula',source:'=1/0'}},'8:1':{value:{kind:'formula',source:'=IF(A1>0,TRUE,FALSE)'}}};
  d.sheets[0].merges=[{row:5,column:1,rowCount:1,columnCount:2}];
  d.sheets[0].rows={'1':{height:254000}};d.sheets[0].columns={'1':{width:952500}};d.sheets[0].view={frozenRows:1,frozenColumns:1,rightToLeft:true,showGridLines:false};
  d.sheets.push({id:'sheet-2',name:'Other',cells:{'1:1':{value:{kind:'formula',source:'=Sheet1!A1+2'}},'2:1':{value:{kind:'formula',source:'=SUM(Sheet1!A1:A2)'}}},rows:{},columns:{},merges:[]});return d;
}
