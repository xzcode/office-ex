// SPDX-License-Identifier: MIT
/** Read-only values from the original workbook; no mutable document session. */
import type {Inspection} from '@office-ex/format-service';
import type {MessageKey} from '@office-ex/i18n';
import {columnName,columnNumber} from '@office-ex/sheets';
type Preview=NonNullable<Inspection['xlsxValuePreview']>;
export interface XlsxPreviewState {sheet:number;row:number;column:number}
const ROWS=50,COLUMNS=12;
function el<K extends keyof HTMLElementTagNameMap>(tag:K,text='',className?:string):HTMLElementTagNameMap[K]{
 const e=document.createElement(tag);e.textContent=text;if(className)e.className=className;return e;
}
export function xlsxValueView(preview:Preview,state:XlsxPreviewState,t:(key:MessageKey,args?:Readonly<Record<string,string|number>>)=>string):HTMLElement{
 const root=el('section','','xlsx-value-preview workspace-card');root.id='xlsx-value-preview';
 const intro=el('p',t('xlsxPreviewNotice'),'boundary-note');root.append(intro);
 if(preview.truncated)root.append(el('p',t('xlsxPreviewLimited'),'diagnostic'));
 const control=el('div','','xlsx-preview-controls'),sheetChoice=el('select');sheetChoice.id='xlsx-preview-sheet';sheetChoice.setAttribute('aria-label',t('xlsxPreviewSheet'));
 preview.sheets.forEach((sheet,i)=>{const o=el('option',sheet.name);o.value=String(i);sheetChoice.append(o);});
 const previousRows=el('button','↑'),nextRows=el('button','↓'),previousCols=el('button','←'),nextCols=el('button','→');
 for(const [b,id,key]of [[previousRows,'xlsx-preview-up','xlsxPreviewPreviousRows'],[nextRows,'xlsx-preview-down','xlsxPreviewNextRows'],[previousCols,'xlsx-preview-left','xlsxPreviewPreviousColumns'],[nextCols,'xlsx-preview-right','xlsxPreviewNextColumns']] as const){b.type='button';b.id=id;b.title=t(key);b.setAttribute('aria-label',t(key));}
 const jump=el('input');jump.id='xlsx-preview-address';jump.type='text';jump.maxLength=10;jump.placeholder='A1';jump.setAttribute('aria-label',t('xlsxPreviewAddress'));
 const go=el('button',t('xlsxPreviewGo'));go.type='button';
 const range=el('span','','muted');range.id='xlsx-preview-range';
 control.append(sheetChoice,previousRows,nextRows,previousCols,nextCols,jump,go,range);root.append(control);
 const viewport=el('div','','xlsx-preview-viewport');const table=el('table','','xlsx-preview-grid');table.id='xlsx-preview-grid';table.setAttribute('aria-label',t('xlsxPreviewSheet'));viewport.append(table);root.append(viewport);
 const formulaInfo=el('div','','xlsx-preview-cell-detail muted');formulaInfo.id='xlsx-preview-cell-detail';formulaInfo.setAttribute('aria-live','polite');root.append(formulaInfo);
 const reset=()=>{const s=preview.sheets[state.sheet]!;state.row=s.firstRow;state.column=s.firstColumn;};
 function render(){
  state.sheet=Math.max(0,Math.min(preview.sheets.length-1,state.sheet));
  const s=preview.sheets[state.sheet]!;
  state.row=Math.max(1,Math.min(s.lastRow,state.row||s.firstRow));state.column=Math.max(1,Math.min(s.lastColumn,state.column||s.firstColumn));
  const r1=state.row,c1=state.column,r2=Math.min(s.lastRow,r1+ROWS-1),c2=Math.min(s.lastColumn,c1+COLUMNS-1);
  sheetChoice.value=String(state.sheet);jump.value=columnName(c1)+r1;jump.removeAttribute('aria-invalid');
  range.textContent=columnName(c1)+r1+' – '+columnName(c2)+r2;
  previousRows.disabled=r1<=1;nextRows.disabled=r2===s.lastRow;previousCols.disabled=c1<=1;nextCols.disabled=c2===s.lastColumn;
  const head=el('thead'),header=el('tr');header.append(el('th',''));
  for(let c=c1;c<=c2;c++){const th=el('th',columnName(c));th.scope='col';header.append(th);}head.append(header);
  const body=el('tbody');for(let r=r1;r<=r2;r++){
   const row=el('tr'),th=el('th',String(r));th.scope='row';row.append(th);
   for(let c=c1;c<=c2;c++){
    const value=s.cells[r+':'+c],cell=el('td',value?.value??(value?'—':''));cell.dataset.address=columnName(c)+r;
    if(value){cell.title=value.value??'';cell.tabIndex=0;cell.dataset.kind=value.kind;
     const select=()=>{formulaInfo.textContent=cell.dataset.address+' · '+(value.kind==='formula-cache'?t('xlsxPreviewCached'):value.kind==='formula-uncached'?t('xlsxPreviewNoCache'):value.kind==='unsupported'?t('xlsxPreviewUnsupported'):t('xlsxPreviewValue'))+(value.formula?' · '+value.formula:'')+(value.value!==null?' · '+value.value:'');};
     cell.onfocus=cell.onclick=select;
    }
    row.append(cell);
   }
   body.append(row);
  }
  table.replaceChildren(head,body);formulaInfo.textContent='';viewport.scrollTop=0;viewport.scrollLeft=0;
 }
 sheetChoice.onchange=()=>{state.sheet=Number(sheetChoice.value);reset();render();};
 previousRows.onclick=()=>{state.row=Math.max(1,state.row-ROWS);render();};nextRows.onclick=()=>{state.row+=ROWS;render();};
 previousCols.onclick=()=>{state.column=Math.max(1,state.column-COLUMNS);render();};nextCols.onclick=()=>{state.column+=COLUMNS;render();};
 const navigate=()=>{const m=/^([A-Z]{1,3})([1-9]\d*)$/.exec(jump.value.trim().toUpperCase()),s=preview.sheets[state.sheet]!;
  try{if(!m)throw Error();const r=Number(m[2]),c=columnNumber(m[1]!);if(!Number.isSafeInteger(r)||r>s.lastRow||c>s.lastColumn)throw Error();state.row=r;state.column=c;render();}
  catch{jump.setAttribute('aria-invalid','true');formulaInfo.textContent=t('xlsxPreviewAddressInvalid');}
 };
 go.onclick=navigate;jump.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();navigate();}};
 render();return root;
}
