/** Test-only host adapter, not included in production renderer. */
import {createDocument,DocumentSession,type SheetsDocument} from '@office-ex/core';
import {message} from '@office-ex/i18n';
import {sheetsWorkspace,releaseSheetCalculation,type SheetsViewState} from '../packages/ui/src/sheets-workspace.js';
let session:DocumentSession,state:SheetsViewState,composing=false,region='en-US',rejectEdits=false,readOnly=false;
const errors:string[]=[];
const root=document.querySelector('#test-root')!;
function render(){root.replaceChildren(sheetsWorkspace({snapshot:()=>session.snapshot(),execute:command=>{if(rejectEdits)return false;try{session.execute(command);return true;}catch(e){errors.push(String(e));return false;}},render,error:e=>errors.push(String(e)),t:(key,args)=>message('en-US',key,args),locale:()=>region,composing:()=>composing,setComposing:value=>{composing=value;},isCurrent:()=>true,format:'xlsx',readOnly},state));}
function load(doc:SheetsDocument){releaseSheetCalculation();session=new DocumentSession(doc);readOnly=false;state={row:1,column:1,anchorRow:1,anchorColumn:1,scrollTop:0,scrollLeft:0,inputMode:'canonical'};render();}
const controls={load,snapshot:()=>session.snapshot(),errors,setRegion:(value:string)=>{region=value;render();},reject:(value:boolean)=>{rejectEdits=value;},opaquePreview:()=>{releaseSheetCalculation();const model=createDocument('sheets','opaque-preview'),sheet=model.sheets[0];sheet.cells['1:2']={value:{kind:'formula',source:'=A1+1'}};session=new DocumentSession(model,{capabilities:{canEdit:false,allowedFeatures:[],diagnostics:[]}});readOnly=true;state={row:1,column:1,anchorRow:1,anchorColumn:1,scrollTop:0,scrollLeft:0,inputMode:'canonical',unsupportedFormulas:[sheet.id+'!1:1'],cachedValues:[[sheet.id+'!1:1',{kind:'number',value:3}],[sheet.id+'!1:2',{kind:'number',value:4}]]};render();},undo:()=>{session.undo();render();},redo:()=>{session.redo();render();},fixture:(size:number)=>{const model=createDocument('sheets','browser-book');model.defaultTextStyle={fontFamily:'Arial',fontSize:139700};for(let i=0;i<size;i++)model.sheets[0].cells[String(Math.floor(i/10)+1)+':'+(i%10+1)]={value:{kind:'number',value:i+1}};return model;}};
Object.assign(window,{testSheets:controls});
load(createDocument('sheets','browser-book'));
