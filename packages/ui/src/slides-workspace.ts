// SPDX-License-Identifier: MIT
import { EMU_PER_CSS_PIXEL, EMU_PER_POINT, ROTATION_UNITS_PER_DEGREE, OfficeError } from '@office-ex/contracts';
import { cloneData,slideTableCoverage,resolveSlideThemeColor,resolveSlideThemeFont,walkSlideShapes, isConnectorTarget, resolveSlideConnector, previewModelPatch, type PatchOperation, type DocumentSnapshot, type DocumentCommand, type ImageResource, type Shape, type Geometry, type TextStyle, type SlideParagraphFormat, type Color, type DeepReadonly, type SlidePreset, type SlideCrop, type SlideStroke, type SlideDash,type SlideTableData,type SlideTableRange,type SlideTableCellFormat } from '@office-ex/core';
import { addSlide, duplicateSlide, deleteSlide, moveSlide, locateSlide, locateShape, presentation, insertShapes, insertSlideTable, replaceSlideTable, editSlideTableGrid, type SlideTableEdit, distributeSlideShapes, removeShapes, duplicateShapes, setShapeGeometries, setShapePaint, reorderShapes, alignShapes, insertSlideImage, aspectFitImage, shapeText, replaceShapeText, slideTextEditDiff, formatShapeText, formatSlideParagraphs, setShapeTextFrame, setSlideBackground, setSlideSize, setSlideTheme, DEFAULT_SLIDE_THEME, pasteShapes, groupShapes, ungroupShape, canUngroupShape, setSlideImageCrop, flipSlideShapes, connectSlideShapes, setSlideConnections, type ShapeSnapshot, type ShapeAlignment, type LayerMove } from '@office-ex/slides';
import { buildSlideScene, walkSlideScene, hitTest, shapeTransform, inverse, transform, multiply, IDENTITY, groupChildTransform, slideShapeParentTransform, type Matrix, type SlideScene, type Point } from '@office-ex/render';
import type { MessageKey } from '@office-ex/i18n';
import {installRibbon,ribbonNodes,type RibbonState,decorateOfficeButton} from './office-ribbon.js';
import {createPanelSplitter,panelWidth,panelCollapsed,persistPanelCollapsed} from './panel-splitter.js';
import { createSlideSceneElement, slideCssColor, overflowingSlideText, measureSlideTextFit, measureSlideShapeGrowth, updateSlideShapeGeometry } from './slides-scene-dom.js';

export interface SlidesViewState extends RibbonState {inspectorMode?:'object'|'text'|'slide';inspectorCollapsed?:boolean;slideId?:string;groupId?:string;selectedIds:string[];zoom:number;lockAspect?:boolean;flush?:()=>boolean;dispose?:()=>void}
export interface SlidesHost {
  snapshot():DocumentSnapshot;execute(command:DocumentCommand):boolean;render():void;error(error:unknown):void;
  composing():boolean;setComposing(value:boolean):void;t(key:MessageKey,args?:Readonly<Record<string,string|number>>):string;number(value:number):string;
  prepareImage(bytes:Uint8Array):Promise<ImageResource>;isCurrent():boolean;readOnly?:boolean;legacyPresentation?:boolean;
}
interface Gesture {pointerId:number;revision:number;kind:'move'|'resize'|'rotate';start:Point;geometries:Record<string,Geometry>;preview:Record<string,Geometry>;shapeId?:string;corner?:string;angle?:number;frame?:number}
let copiedShapes:{nonce:string;shapes:readonly ShapeSnapshot[];resources:Record<string,DeepReadonly<ImageResource>>}|undefined;
const rgb=(value:string):Color=>({kind:'rgba',red:parseInt(value.slice(1,3),16),green:parseInt(value.slice(3,5),16),blue:parseInt(value.slice(5,7),16),alpha:255});
const colorHex=(color:DeepReadonly<Color>|undefined,fallback='#426fca')=>color?.kind==='rgba'?'#'+[color.red,color.green,color.blue].map(v=>v.toString(16).padStart(2,'0')).join(''):fallback;
const id=()=>crypto.randomUUID();
const emu=(points:string)=>Math.round(Number(points)*EMU_PER_POINT);
function element<K extends keyof HTMLElementTagNameMap>(tag:K,value='',className=''):HTMLElementTagNameMap[K]{const node=document.createElement(tag);node.textContent=value;if(className)node.className=className;return node;}
function isTyping(target:EventTarget|null):boolean{return target instanceof HTMLElement&&(target.isContentEditable||['INPUT','TEXTAREA','SELECT'].includes(target.tagName));}

export function slidesWorkspace(host:SlidesHost,state:SlidesViewState):HTMLElement {
  state.dispose?.();state.selectedIds??=[];state.zoom??=0;state.lockAspect??=true;state.inspectorCollapsed??=panelCollapsed('slide-inspector');
  const t=host.t,root=element('section','','slides-workspace');root.id='slides-workspace';root.setAttribute('aria-label',t('slides'));
  const toolbar=element('div','','slides-toolbar'),body=element('div','','slides-workbench'),rail=element('nav','','slides-thumbnails'),stage=element('div','','slides-stage'),inspector=element('aside','','slides-inspector'),surface=element('div','','slides-surface'),scaled=element('div','','slides-scaled-scene'),overlay=element('div','','slides-selection-overlay'),status=element('div','','slides-status');
  toolbar.setAttribute('role','toolbar');toolbar.setAttribute('aria-label',t('slides'));rail.setAttribute('aria-label',t('slides'));stage.tabIndex=0;stage.setAttribute('role','application');stage.setAttribute('aria-label',t('canvas'));root.append(toolbar,body,status);body.classList.add('resizable-slides');
  let railWidth=panelWidth('slide-rail',158,110,360),inspectorWidth=panelWidth('slide-inspector',270,200,560);
  const available=()=>body.clientWidth||window.innerWidth;
  function layoutPanels(){
   const room=available(),right=state.inspectorCollapsed?0:Math.min(inspectorWidth,Math.max(200,room-railWidth-260-16));
   const left=Math.min(railWidth,Math.max(110,room-right-260-16));
   body.style.setProperty('--slide-rail-width',left+'px');body.style.setProperty('--slide-inspector-width',right+'px');
  }
  const leftSplitter=createPanelSplitter({key:'slide-rail',label:t('ribbonPages'),side:'start',panel:rail,width:()=>railWidth,setWidth:v=>{railWidth=v;layoutPanels();},limits:()=>({min:110,max:Math.max(110,Math.min(360,available()-(state.inspectorCollapsed?0:Math.min(inspectorWidth,560))-276))})});
  const rightSplitter=createPanelSplitter({key:'slide-inspector',label:t('ribbonProperties'),side:'end',panel:inspector,width:()=>inspectorWidth,setWidth:v=>{inspectorWidth=v;layoutPanels();},limits:()=>({min:200,max:Math.max(200,Math.min(560,available()-railWidth-276))})});
  body.append(rail,leftSplitter,stage,rightSplitter,inspector);layoutPanels();stage.append(surface);surface.append(scaled);scaled.append(overlay);
  const help=element('p',t('slideHelp'),'muted slides-help'),overflow=element('p','','slides-overflow'),position=element('span','');status.append(position,overflow,help);overflow.hidden=true;
  let scene:SlideScene,sceneElement:HTMLDivElement,gesture:Gesture|undefined,disposed=false,zoom=1,textTimer:ReturnType<typeof setTimeout>|undefined,textDirty=false,editingId:string|undefined,observer:ResizeObserver|undefined,thumbObserver:IntersectionObserver|undefined,presentationDialog:HTMLDialogElement|undefined,closePresentation:(()=>void)|undefined;
  const editControls:(HTMLButtonElement|HTMLInputElement|HTMLSelectElement|HTMLTextAreaElement)[]=[];const selectionControls:(HTMLButtonElement|HTMLInputElement|HTMLSelectElement)[]=[];
  const snap=()=>host.snapshot(),doc=()=>presentation(snap()),selected=()=>scopeShapes().filter(shape=>state.selectedIds.includes(shape.id));
  function safe(fn:()=>void):void{try{fn();}catch(error){host.error(error);}}
  function button(key:MessageKey,fn:()=>void,target:HTMLElement=toolbar,identifier?:string,needsSelection=false,mutates=true):HTMLButtonElement {
    const b=element('button',t(key));b.type='button';if(identifier)b.id=identifier;b.onpointerdown=event=>{if(event.pointerType==='mouse')event.preventDefault();};b.onclick=()=>safe(()=>{if(flush())fn();});target.append(b);if(mutates)editControls.push(b);if(needsSelection)selectionControls.push(b);return b;
  }
  function field(key:MessageKey,control:HTMLInputElement|HTMLSelectElement|HTMLTextAreaElement,parent:HTMLElement=inspector):void {const label=element('label','','slides-field');label.append(element('span',t(key)),control);control.setAttribute('aria-label',t(key));parent.append(label);editControls.push(control);}
  function input(key:MessageKey,type:string,value:string,fn:(value:string)=>void,needsSelection=true):HTMLInputElement {
    const control=element('input');control.type=type;control.value=value;field(key,control);control.onchange=()=>safe(()=>{if(flush())fn(control.value);});if(needsSelection)selectionControls.push(control);return control;
  }
  function fitCommand(command:DocumentCommand):DocumentCommand{
    if(host.legacyPresentation||command.type!=='model.patch'||!state.slideId)return command;
    const ops=command.payload.operations.filter(op=>op.op==='set'&&['runs','geometry','textStyle','textFrame'].includes(String(op.path.at(-1))));
    if(!ops.length)return command;
    const snapshot=snap(),slide=locateSlide(snapshot,state.slideId),affected=[...walkSlideShapes(slide.slide.shapes)].filter(shape=>{
      if(shape.type!=='text'||!shape.textFrame?.autofit||shape.textFrame.autofit.mode==='none')return false;
      const path=locateShape(snapshot,slide.slide.id,shape.id).path;
      return ops.some(op=>op.path.length===path.length+1&&path.every((v,i)=>op.path[i]===v));
    });
    if(!affected.length)return command;
    if(affected.length>32)throw new OfficeError('limit','Too many fitted text boxes in one edit');
    const candidate=previewModelPatch(snapshot,command),projected=new Map([...walkSlideScene(buildSlideScene(candidate,slide.slide.id).shapes)].map(shape=>[shape.id,shape])),extra:PatchOperation[]=[];
    for(const old of affected){
      if(old.type!=='text')continue;
      const location=locateShape(candidate,slide.slide.id,old.id),model=location.shape,shape=projected.get(old.id);
      if(model.type!=='text'||!model.textFrame?.autofit||model.textFrame.autofit.mode==='none'||shape?.type!=='text')continue;
      const newFit=model.textFrame.autofit;
      if(newFit.mode==='shape'){
        if(JSON.stringify(old)===JSON.stringify(model))continue;
        if(location.parent||model.geometry.rotation!==0||model.textFrame.verticalAlignment==='center'||model.textFrame.verticalAlignment==='bottom')throw new OfficeError('unsupported','Shape autofit requires a direct unrotated top-anchored textbox');
        const growth=measureSlideShapeGrowth(shape);
        if(growth.width!==model.geometry.width||growth.height!==model.geometry.height)extra.push({op:'set',path:[...location.path,'geometry'],value:{...model.geometry,...growth}});
        continue;
      }
      // Explicit Fit/Reset already carries its selected scale.
      if(JSON.stringify(old.textFrame?.autofit)!==JSON.stringify(newFit)&&ops.every(op=>op.path.at(-1)==='textFrame'))continue;
      if(JSON.stringify(old)===JSON.stringify(model))continue;
      const fontScale=measureSlideTextFit(shape)||10000;
      if(fontScale!==newFit.fontScale)extra.push({op:'set',path:[...location.path,'textFrame'],value:{...model.textFrame,autofit:{...newFit,fontScale}}});
    }
    return extra.length?{...command,payload:{...command.payload,operations:[...command.payload.operations,...extra]}}:command;
  }
  function apply(command:DocumentCommand,repaint=true):boolean {if(host.readOnly||disposed)return false;const ok=host.execute(fitCommand(command));if(ok&&repaint)refresh();return ok;}
  function cancelGesture():void {if(!gesture)return;const previous=gesture;gesture=undefined;if(previous.frame!==undefined)cancelAnimationFrame(previous.frame);if(stage.hasPointerCapture(previous.pointerId))stage.releasePointerCapture(previous.pointerId);paintGeometry();}
  function flush():boolean {if(disposed)return true;if(host.composing())return false;if(gesture)cancelGesture();if(textTimer!==undefined){clearTimeout(textTimer);textTimer=undefined;}return commitText();}
  state.flush=flush;
  function selectedText():Extract<ShapeSnapshot,{type:'text'}>|undefined {if(state.selectedIds.length!==1)return undefined;const s=selected()[0];return s?.type==='text'?s:undefined;}
  function commitText():boolean {
    if(!textDirty||!editingId)return true;if(host.readOnly||host.composing())return false;
    try{const shape=locateShape(snap(),state.slideId!,editingId).shape;if(shape.type!=='text')return false;const before=shapeText(shape);if(before===textEditor.value){textDirty=false;return true;}const difference=slideTextEditDiff(before,textEditor.value);const ok=apply(replaceShapeText(snap(),state.slideId!,editingId,difference.start,difference.end,difference.text,id),false);if(ok){textDirty=false;refreshScene();refreshThumbnails();}return ok;}catch(error){host.error(error);return false;}
  }
  function changeSlide(slideId:string):void {if(!flush())return;state.slideId=slideId;state.groupId=undefined;state.selectedIds=[];editingId=undefined;refresh();stage.focus({preventScroll:true});}
  function activeGroup():Extract<ShapeSnapshot,{type:'group'}>|undefined{
    if(!state.groupId||!state.slideId)return undefined;
    const group=[...walkSlideShapes(locateSlide(snap(),state.slideId).slide.shapes)].find(s=>s.id===state.groupId);
    if(group?.type==='group')return group;state.groupId=undefined;return undefined;
  }
  function scopeShapes():readonly ShapeSnapshot[]{return activeGroup()?.children??locateSlide(snap(),state.slideId!).slide.shapes;}
  function scopeMatrix():Matrix{
    const group=activeGroup();if(!group)return IDENTITY;
    return multiply(slideShapeParentTransform(locateSlide(snap(),state.slideId!).slide.shapes,group.id)??IDENTITY,groupChildTransform(group));
  }
  function choose(ids:string[]):void{
    if(!flush())return;
    if(ids.length){const locations=ids.map(shapeId=>locateShape(snap(),state.slideId!,shapeId));if(!locations.every(l=>l.siblings===locations[0]!.siblings))return;state.groupId=locations[0]!.parent?.id;}
    state.selectedIds=ids;syncInspector();paintGeometry();
  }
  function enterGroup(group:Extract<ShapeSnapshot,{type:'group'}>):void{if(!flush())return;state.groupId=group.id;state.selectedIds=[];syncInspector();paintGeometry();stage.focus({preventScroll:true});}
  function leaveGroup():void{const group=activeGroup();if(group)choose([group.id]);}
  function duplicateSelection():void{const count=scopeShapes().length;if(apply(duplicateShapes(snap(),state.slideId!,state.selectedIds,id)))choose(scopeShapes().slice(count).map(s=>s.id));}
  function selectedHasBindings():boolean{return selected().some(s=>s.type==='line'&&s.connections);}

  function addShape(type:'text'|SlidePreset,connector=false):void {
    const m=doc(),slide=locateSlide(snap(),state.slideId!).slide,shapeId=id(),geometry={x:Math.round(m.size.width*.2),y:Math.round(m.size.height*.2),width:Math.round(m.size.width*.4),height:Math.round(m.size.height*(type==='line'?.15:type==='text'?.2:.3)),rotation:0};let shape:Shape;
    if(type==='text')shape={id:shapeId,type,runs:[{id:id(),type:'text',text:''}],textStyle:{fontSize:304800},geometry};
    else shape={id:shapeId,type,geometry,...(connector?{connector:true as const}:{}),...(type==='line'?{}:{fill:{kind:'theme',name:'accent1'} as Color}),stroke:{color:{kind:'theme',name:type==='line'?'accent1':'dk1'},width:12700}};
    if(apply(insertShapes(snap(),state.slideId!,slide.shapes.length,[shape]))){choose([shapeId]);if(type==='text')focusTextEditor();}
  }
  button('slideAdd',()=>{const current=locateSlide(snap(),state.slideId!),newId=id();if(apply(addSlide(snap(),current.index+1,newId,current.slide.background as Color))){state.slideId=newId;state.groupId=undefined;state.selectedIds=[];refresh();}},toolbar,'slide-add');
  button('slideDuplicate',()=>{const current=locateSlide(snap(),state.slideId!);if(apply(duplicateSlide(snap(),state.slideId!,id))){state.slideId=doc().slides[current.index+1]!.id;state.groupId=undefined;state.selectedIds=[];refresh();}},toolbar,'slide-duplicate');
  const deletePage=button('slideDelete',()=>{const current=locateSlide(snap(),state.slideId!);if(apply(deleteSlide(snap(),state.slideId!))){state.slideId=doc().slides[Math.min(current.index,doc().slides.length-1)]!.id;state.groupId=undefined;state.selectedIds=[];refresh();}},toolbar,'slide-delete');
  const earlier=button('slideEarlier',()=>{const current=locateSlide(snap(),state.slideId!);apply(moveSlide(snap(),state.slideId!,current.index-1));},toolbar,'slide-earlier');
  const later=button('slideLater',()=>{const current=locateSlide(snap(),state.slideId!);apply(moveSlide(snap(),state.slideId!,current.index+1));},toolbar,'slide-later');
  for(const [type,key] of [['text','slideText'],['rect','slideRectangle'],['ellipse','slideEllipse'],['line','slideLine']] as const)button(key,()=>addShape(type),toolbar,'slide-insert-'+type);
  const tableDialog=element('dialog');tableDialog.id='slide-table-dialog';
  const tableHeader=element('header');tableHeader.append(element('h2',t('slideTable')),button('close',()=>tableDialog.close(),tableHeader,'slide-table-close',false,false));tableDialog.append(tableHeader);
  const dims=element('div');const rowInput=element('input'),colInput=element('input');for(const [input,identifier,value,max]of [[rowInput,'slide-table-rows','3','200'],[colInput,'slide-table-columns','3','64']] as const){input.type='number';input.min='1';input.max=max;input.value=value;input.id=identifier;}
  field('writerTableRows',rowInput,dims);field('writerTableColumns',colInput,dims);tableDialog.append(dims);
  const tableTools=element('div','','slide-table-tools'),rangeBar=element('div','','slide-table-bar'),structureBar=element('div','','slide-table-bar'),formatBar=element('div','','slide-table-bar');
  tableTools.append(rangeBar,structureBar,formatBar);tableDialog.append(tableTools);
  const matrix=element('div');matrix.id='slide-table-cells';tableDialog.append(matrix);
  let tableId:string|undefined,tableDraft:SlideTableData|undefined,tableRevision=0;
  const numberField=(key:MessageKey,name:string,parent:HTMLElement,value='1')=>{
    const input=element('input');input.id=name;input.type='number';input.min='1';input.step='1';input.value=value;field(key,input,parent);return input;
  };
  const tableRow=numberField('slideTableStartRow','slide-table-row',rangeBar),tableColumn=numberField('slideTableStartColumn','slide-table-column',rangeBar);
  const spanRows=numberField('slideTableRowCount','slide-table-span-rows',rangeBar),spanColumns=numberField('slideTableColumnCount','slide-table-span-columns',rangeBar);
  const chosenRange=():SlideTableRange=>({row:tableRow.valueAsNumber-1,column:tableColumn.valueAsNumber-1,rowCount:spanRows.valueAsNumber,columnCount:spanColumns.valueAsNumber});
  function chooseTableCell(r:number,c:number,rs=1,cs=1){tableRow.value=String(r+1);tableColumn.value=String(c+1);spanRows.value=String(rs);spanColumns.value=String(cs);syncTableRange();}
  function syncTableRange(){
    if(!tableDraft)return;const rows=tableDraft.rowHeights.length,cols=tableDraft.columnWidths.length;
    tableRow.max=String(rows);tableColumn.max=String(cols);
    const clamp=(v:number,min:number,max:number)=>Math.min(max,Math.max(min,Number.isFinite(v)?Math.trunc(v):min));
    tableRow.value=String(clamp(tableRow.valueAsNumber,1,rows));tableColumn.value=String(clamp(tableColumn.valueAsNumber,1,cols));
    spanRows.max=String(rows-tableRow.valueAsNumber+1);spanColumns.max=String(cols-tableColumn.valueAsNumber+1);
    spanRows.value=String(clamp(spanRows.valueAsNumber,1,Number(spanRows.max)));spanColumns.value=String(clamp(spanColumns.valueAsNumber,1,Number(spanColumns.max)));
    const range=chosenRange(),coverage=slideTableCoverage(tableDraft),origin=coverage.get(range.row+':'+range.column);
    for(const cell of matrix.querySelectorAll<HTMLTextAreaElement>('textarea[data-row]')){
      const r=Number(cell.dataset.row),c=Number(cell.dataset.column),m=coverage.get(r+':'+c);
      const selected=r<range.row+range.rowCount&&range.row<r+(m?.rowCount??1)&&c<range.column+range.columnCount&&range.column<c+(m?.columnCount??1);
      cell.classList.toggle('table-range-selected',selected);
    }
    const f=tableDraft.formats?.[(origin?.row??range.row)+':'+(origin?.column??range.column)];
    const visibleColor=(c:DeepReadonly<Color>|null|undefined,fallback:string)=>colorHex(c?resolveSlideThemeColor(c,doc().theme):undefined,fallback);
    tableFill.value=visibleColor(f?.fill,'#ffffff');tableInk.value=visibleColor(f?.textStyle?.color,'#182030');
    tableBorder.value=visibleColor(f?.borders?.left?.color,'#182030');tableBorderWidth.value=String((f?.borders?.left?.width??12700)/EMU_PER_POINT);
    tableFamily.value=f?.textStyle?.fontFamily??'';tableSize.value=String((f?.textStyle?.fontSize??228600)/EMU_PER_POINT);
    tableBold.setAttribute('aria-pressed',String(!!f?.textStyle?.bold));tableItalic.setAttribute('aria-pressed',String(!!f?.textStyle?.italic));tableUnderline.setAttribute('aria-pressed',String(!!f?.textStyle?.underline));
    tableAlignment.value=f?.alignment??'start';tableVertical.value=f?.verticalAlignment??'top';
    mergeTable.disabled=range.rowCount*range.columnCount<=1;splitTable.disabled=!origin;
    deleteRow.disabled=rows<=1;deleteCol.disabled=cols<=1;addRow.disabled=rows>=200||(rows+1)*cols>2000;addCol.disabled=cols>=64||rows*(cols+1)>2000;
  }
  for(const control of [tableRow,tableColumn,spanRows,spanColumns])control.onchange=syncTableRange;
  function changeTableDraft(action:SlideTableEdit){
    if(!tableDraft)return;const range=chosenRange();tableDraft=editSlideTableGrid(tableDraft,action);
    renderTableDraft();
    const keepRange=action.kind==='merge'||action.kind==='format'||action.kind==='resetFormat';
    chooseTableCell(Math.min(range.row,tableDraft.rowHeights.length-1),Math.min(range.column,tableDraft.columnWidths.length-1),keepRange?range.rowCount:1,keepRange?range.columnCount:1);
  }
  const mergeTable=button('writerMergeCells',()=>changeTableDraft({kind:'merge',range:chosenRange()}),structureBar,'slide-table-merge');
  const splitTable=button('writerSplitCell',()=>changeTableDraft({kind:'split',row:chosenRange().row,column:chosenRange().column}),structureBar,'slide-table-split');
  const addRow=button('writerAddRow',()=>changeTableDraft({kind:'insert',axis:'row',index:chosenRange().row+1}),structureBar,'slide-table-add-row');
  const deleteRow=button('writerDeleteRow',()=>changeTableDraft({kind:'delete',axis:'row',index:chosenRange().row}),structureBar,'slide-table-delete-row');
  const addCol=button('writerAddColumn',()=>changeTableDraft({kind:'insert',axis:'column',index:chosenRange().column+1}),structureBar,'slide-table-add-column');
  const deleteCol=button('writerDeleteColumn',()=>changeTableDraft({kind:'delete',axis:'column',index:chosenRange().column}),structureBar,'slide-table-delete-column');
  const applyFormat=(format:SlideTableCellFormat)=>changeTableDraft({kind:'format',range:chosenRange(),format});
  const colorField=(name:string,value:string)=>{const c=element('input');c.type='color';c.id=name;c.value=value;return c;};
  const tableFill=colorField('slide-table-fill','#ffffff'),tableInk=colorField('slide-table-text-color','#182030'),tableBorder=colorField('slide-table-border-color','#182030');
  field('sheetFillColor',tableFill,formatBar);tableFill.onchange=()=>safe(()=>applyFormat({fill:rgb(tableFill.value)}));
  field('writerTextColor',tableInk,formatBar);tableInk.onchange=()=>safe(()=>applyFormat({textStyle:{color:rgb(tableInk.value)}}));
  field('writerTableBorder',tableBorder,formatBar);const tableBorderWidth=numberField('slideStrokeWidth','slide-table-border-width',formatBar);tableBorderWidth.min='0';tableBorderWidth.max='100';tableBorderWidth.step='.25';
  button('writerTableBorder',()=>{
    const width=Math.round(tableBorderWidth.valueAsNumber*EMU_PER_POINT);if(!Number.isFinite(width)||width<0||width>12700*100)throw new OfficeError('invalid','Invalid table border width');
    const border=width?{color:rgb(tableBorder.value),width}:null;
    applyFormat({borders:{left:border,right:border,top:border,bottom:border}});
  },formatBar,'slide-table-apply-border');
  const tableFamily=element('input');tableFamily.id='slide-table-font';tableFamily.placeholder=t('ribbonFont');field('writerFontFamily',tableFamily,formatBar);
  tableFamily.onchange=()=>safe(()=>{if(tableFamily.value.trim())applyFormat({textStyle:{fontFamily:tableFamily.value.trim()}});});
  const tableSize=numberField('writerFontSize','slide-table-font-size',formatBar,'18');tableSize.max='1638';tableSize.step='.01';tableSize.onchange=()=>safe(()=>applyFormat({textStyle:{fontSize:Math.round(tableSize.valueAsNumber*100)*127}}));
  const tableBold=button('writerBold',()=>applyFormat({textStyle:{bold:tableBold.getAttribute('aria-pressed')!=='true'}}),formatBar,'slide-table-bold');
  const tableItalic=button('writerItalic',()=>applyFormat({textStyle:{italic:tableItalic.getAttribute('aria-pressed')!=='true'}}),formatBar,'slide-table-italic');
  const tableUnderline=button('writerUnderline',()=>applyFormat({textStyle:{underline:tableUnderline.getAttribute('aria-pressed')!=='true'}}),formatBar,'slide-table-underline');
  const tableAlignment=element('select');tableAlignment.id='slide-table-align';for(const [v,k]of [['start','writerAlignLeft'],['center','writerAlignCenter'],['end','writerAlignRight']] as const){const o=element('option',t(k));o.value=v;tableAlignment.append(o);}field('ribbonAlignment',tableAlignment,formatBar);tableAlignment.onchange=()=>safe(()=>applyFormat({alignment:tableAlignment.value as SlideTableCellFormat['alignment']}));
  const tableVertical=element('select');tableVertical.id='slide-table-vertical';for(const [v,k]of [['top','slideAlignTop'],['center','slideAlignMiddle'],['bottom','slideAlignBottom']] as const){const o=element('option',t(k));o.value=v;tableVertical.append(o);}field('ribbonAlignment',tableVertical,formatBar);tableVertical.onchange=()=>safe(()=>applyFormat({verticalAlignment:tableVertical.value as SlideTableCellFormat['verticalAlignment']}));
  button('slideTableResetFormat',()=>changeTableDraft({kind:'resetFormat',range:chosenRange()}),formatBar,'slide-table-reset-format');
  function renderTableDraft(){
    matrix.replaceChildren();if(!tableDraft)return;
    const table=tableDraft,coverage=slideTableCoverage(table);
    Object.assign(matrix.style,{display:'grid',gridTemplateColumns:'repeat('+table.columnWidths.length+', minmax(90px,1fr))',gridTemplateRows:'repeat('+table.rowHeights.length+',minmax(64px,auto))',gap:'4px',maxHeight:'45vh',overflow:'auto'});
    for(const [r,row]of table.cells.entries())for(const [c,value]of row.entries()){
      const merge=coverage.get(r+':'+c);if(merge&&(merge.row!==r||merge.column!==c))continue;
      const input=element('textarea');input.dataset.row=String(r);input.dataset.column=String(c);input.value=value;input.maxLength=4096;input.rows=2;input.setAttribute('aria-label',t('writerCell',{row:host.number(r+1),column:host.number(c+1)}));
      input.style.gridRow=(r+1)+' / span '+(merge?.rowCount??1);input.style.gridColumn=(c+1)+' / span '+(merge?.columnCount??1);
      const f=table.formats?.[r+':'+c],cssColor=(v:Color)=>slideCssColor(resolveSlideThemeColor(v,doc().theme));
      if(f?.fill!==undefined)input.style.background=f.fill===null?'transparent':cssColor(f.fill);
      if(f?.textStyle){const s=f.textStyle;if(s.color)input.style.color=cssColor(s.color);if(s.bold!==undefined)input.style.fontWeight=s.bold?'700':'400';if(s.italic!==undefined)input.style.fontStyle=s.italic?'italic':'normal';if(s.underline!==undefined)input.style.textDecoration=s.underline?'underline':'none';if(s.fontFamily)input.style.fontFamily=resolveSlideThemeFont(s.fontFamily,doc().theme)??s.fontFamily;if(s.fontSize)input.style.fontSize=Math.min(32,s.fontSize/EMU_PER_CSS_PIXEL)+'px';}
      input.style.textAlign=f?.alignment??'start';
      input.onfocus=()=>chooseTableCell(r,c,merge?.rowCount??1,merge?.columnCount??1);
      input.oninput=()=>{table.cells[r]![c]=input.value;};
      matrix.append(input);
    }
    syncTableRange();
  }
  const applyTable=button('writerApplyPage',()=>{
    if(tableId&&tableDraft){
      if(snap().revision!==tableRevision)throw new OfficeError('conflict','Table changed while dialog was open');
      if(apply(replaceSlideTable(snap(),state.slideId!,tableId,tableDraft)))tableDialog.close();
    }else{const shapeId=id();if(apply(insertSlideTable(snap(),state.slideId!,rowInput.valueAsNumber,colInput.valueAsNumber,shapeId))){tableDialog.close();choose([shapeId]);}}
  },tableDialog,'slide-table-apply');
  button('cancel',()=>tableDialog.close(),tableDialog,'slide-table-cancel',false,false);
  root.append(tableDialog);
  function openTable(editing:boolean){
    tableId=undefined;tableDraft=undefined;matrix.replaceChildren();dims.hidden=editing;tableTools.hidden=!editing;matrix.hidden=!editing;
    if(editing){const s=selected()[0];if(s?.type!=='table')return;tableId=s.id;tableRevision=snap().revision;tableDraft=cloneData(s.table) as SlideTableData;tableRow.value=tableColumn.value=spanRows.value=spanColumns.value='1';renderTableDraft();}
    tableDialog.showModal();
  }
  const insertTableButton=button('slideTable',()=>openTable(false),toolbar,'slide-insert-table');
  const editTableButton=button('slideEditTable',()=>openTable(true),toolbar,'slide-edit-table',true);
  const distributeH=button('slideDistributeHorizontal',()=>apply(distributeSlideShapes(snap(),state.slideId!,state.selectedIds,'horizontal')),toolbar,'slide-distribute-h',true);
  const distributeV=button('slideDistributeVertical',()=>apply(distributeSlideShapes(snap(),state.slideId!,state.selectedIds,'vertical')),toolbar,'slide-distribute-v',true);
  const advancedPresets:[SlidePreset,MessageKey][]=[['roundRect','slideRoundRect'],['triangle','slideTriangle'],['diamond','slideDiamond'],['parallelogram','slideParallelogram']];
  const newGraphics:HTMLButtonElement[]=advancedPresets.map(([type,key])=>button(key,()=>addShape(type),toolbar,'slide-insert-'+type));
  newGraphics.push(button('slideConnector',()=>addShape('line',true),toolbar,'slide-insert-connector'));
  const fileInput=element('input');fileInput.type='file';fileInput.accept='image/png,image/jpeg';fileInput.hidden=true;root.append(fileInput);
  let imageRequest=0;
  button('writerInsertImage',()=>fileInput.click(),toolbar,'slide-insert-image');
  fileInput.onchange=()=>safe(()=>{
    const file=fileInput.files?.[0];fileInput.value='';if(!file||disposed||host.readOnly||!host.isCurrent())return;
    // File selection can arrive before a pending text debounce or without the toolbar click.
    if(!flush())throw new OfficeError('conflict','Finish the current text edit before inserting an image');
    const request=++imageRequest,{document:{documentId},revision}=snap(),slideId=state.slideId!;
    void(async()=>{try{
      if(file.size>8*1024*1024)throw new OfficeError('limit','Image byte budget exceeded');
      const resource=await host.prepareImage(new Uint8Array(await file.arrayBuffer()));
      if(disposed||request!==imageRequest||!host.isCurrent()||snap().document.documentId!==documentId)return;
      // Never insert into a different slide or silently overwrite newer editing state.
      if(host.readOnly||state.slideId!==slideId||!flush()||snap().revision!==revision)throw new OfficeError('conflict','The slide changed while preparing the image; please insert it again');
      const m=doc(),fit=aspectFitImage(resource,m.size.width*.6,m.size.height*.6),shapeId=id(),slide=locateSlide(snap(),slideId).slide;
      if(apply(insertSlideImage(snap(),slideId,slide.shapes.length,shapeId,resource,{x:Math.round((m.size.width-fit.width)/2),y:Math.round((m.size.height-fit.height)/2),...fit,rotation:0})))choose([shapeId]);
    }catch(error){if(!disposed&&request===imageRequest&&host.isCurrent())host.error(error);}})();
  });
  button('slidePresent',()=>present(),toolbar,'slide-present',false,false);
  button('menuZoomOut',()=>{state.zoom=Math.max(.1,zoom/1.25);resize();},toolbar,'slide-zoom-out',false,false);
  button('menuResetZoom',()=>{state.zoom=0;resize();},toolbar,'slide-zoom-fit',false,false);
  button('menuZoomIn',()=>{state.zoom=Math.min(4,zoom*1.25);resize();},toolbar,'slide-zoom-in',false,false);

  const geometryGroup=element('fieldset'),geometryLegend=element('legend',t('slideArrange'));geometryGroup.append(geometryLegend);inspector.append(geometryGroup);
  function geometryInput(key:MessageKey,property:keyof Geometry):HTMLInputElement {
    const control=element('input');control.type='number';control.step=property==='rotation'?'1':'.1';field(key,control,geometryGroup);selectionControls.push(control);control.onchange=()=>safe(()=>{if(!flush()||state.selectedIds.length!==1)return;const shape=selected()[0]!,g={...shape.geometry};const value=property==='rotation'?Math.round(Number(control.value)*ROTATION_UNITS_PER_DEGREE):emu(control.value);if(!Number.isFinite(value))throw new OfficeError('invalid','Invalid geometry');g[property]=value;
      if((property==='width'||property==='height')&&(state.lockAspect||shape.type==='image')&&shape.geometry.width>0&&shape.geometry.height>0){if(property==='width')g.height=Math.max(1,Math.round(value*shape.geometry.height/shape.geometry.width));else g.width=Math.max(1,Math.round(value*shape.geometry.width/shape.geometry.height));}
      apply(setShapeGeometries(snap(),state.slideId!,{[shape.id]:g}));});return control;
  }
  const geometryFields={x:geometryInput('slideX','x'),y:geometryInput('slideY','y'),width:geometryInput('slideWidth','width'),height:geometryInput('slideHeight','height'),rotation:geometryInput('slideRotation','rotation')};geometryFields.width.min='0';geometryFields.height.min='0';
  const aspect=element('input');aspect.type='checkbox';aspect.checked=state.lockAspect;field('slideAspect',aspect,geometryGroup);aspect.onchange=()=>{state.lockAspect=aspect.checked;};
  const objectActions=element('div','','slides-tool-buttons');geometryGroup.append(objectActions);
  button('slideDuplicateObject',duplicateSelection,objectActions,'slide-object-duplicate',true);
  button('slideDeleteObject',()=>{if(apply(removeShapes(snap(),state.slideId!,state.selectedIds)))choose([]);},objectActions,'slide-object-delete',true);
  const groupButton=button('slideGroupObjects',()=>{const groupId=id();if(apply(groupShapes(snap(),state.slideId!,state.selectedIds,()=>groupId)))choose([groupId]);},objectActions,'slide-group',true);
  const ungroupButton=button('slideUngroupObjects',()=>{const shape=selected()[0];if(shape?.type==='group'){const ids=shape.children.map(child=>child.id);if(apply(ungroupShape(snap(),state.slideId!,shape.id)))choose(ids);}},objectActions,'slide-ungroup',true);
  const flipButtons=(['flipH','flipV'] as const).map((axis,index)=>button(index?'slideFlipV':'slideFlipH',()=>apply(flipSlideShapes(snap(),state.slideId!,state.selectedIds,axis)),objectActions,'slide-'+axis,true));
  const editGroup=button('slideEditGroup',()=>{const shape=selected()[0];if(shape?.type==='group')enterGroup(shape);},objectActions,'slide-edit-group',true);
  const exitGroup=button('slideExitGroup',leaveGroup,objectActions,'slide-exit-group');
  const connectButton=button('slideConnectObjects',()=>{const newId=id();if(apply(connectSlideShapes(snap(),state.slideId!,state.selectedIds,()=>newId)))choose([newId]);},objectActions,'slide-connect-objects',true);
  const detachButton=button('slideDetachConnector',()=>{const shape=selected()[0];if(shape?.type==='line'&&shape.connections)apply(setSlideConnections(snap(),state.slideId!,shape.id,null));},objectActions,'slide-detach-connector',true);
  const objectTree=element('select');objectTree.id='slide-object-tree';field('slideObjectTree',objectTree,geometryGroup);
  objectTree.onchange=()=>safe(()=>{if(objectTree.value)choose([objectTree.value]);});
  const arrangeGroup=element('div','','slides-tool-buttons');geometryGroup.append(arrangeGroup);
  for(const [move,key] of [['front','slideFront'],['back','slideBack'],['forward','slideForward'],['backward','slideBackward']] as const)button(key,()=>apply(reorderShapes(snap(),state.slideId!,state.selectedIds,move as LayerMove)),arrangeGroup,'slide-layer-'+move,true);
  for(const [alignment,key] of [['left','writerAlignLeft'],['center','writerAlignCenter'],['right','writerAlignRight'],['top','slideAlignTop'],['middle','slideAlignMiddle'],['bottom','slideAlignBottom']] as const)button(key,()=>apply(alignShapes(snap(),state.slideId!,state.selectedIds,alignment as ShapeAlignment)),arrangeGroup,'slide-align-'+alignment,true);
  const paintGroup=element('fieldset');paintGroup.append(element('legend',t('slideFill')));inspector.append(paintGroup);
  const fill=element('input');fill.type='color';field('slideFill',fill,paintGroup);selectionControls.push(fill);fill.onchange=()=>safe(()=>{if(flush())apply(setShapePaint(snap(),state.slideId!,state.selectedIds,{fill:rgb(fill.value)}));});
  function currentStroke():SlideStroke{const single=selected()[0];return single?.stroke?structuredClone(single.stroke):{color:{kind:'rgba',red:24,green:32,blue:48,alpha:255},width:12700};}
  function updateStroke(patch:Partial<SlideStroke>,remove?:'head'|'tail'):void{
    if(!flush()||state.selectedIds.length!==1)return;const value={...currentStroke(),...patch};if(remove)delete value[remove];
    apply(setShapePaint(snap(),state.slideId!,state.selectedIds,{stroke:value}));
  }
  const stroke=element('input');stroke.type='color';field('slideStroke',stroke,paintGroup);selectionControls.push(stroke);stroke.onchange=()=>safe(()=>updateStroke({color:rgb(stroke.value)}));
  const strokeWidth=element('input');strokeWidth.type='number';strokeWidth.min='0';strokeWidth.max='100';strokeWidth.step='.25';field('slideStrokeWidth',strokeWidth,paintGroup);selectionControls.push(strokeWidth);strokeWidth.onchange=()=>safe(()=>updateStroke({width:emu(strokeWidth.value)}));
  const dash=element('select');dash.id='slide-stroke-dash';for(const [value,label]of [['solid','━━━━'],['dash','━ ━ ━'],['dot','• • •'],['dashDot','━ • ━ •'],['lgDash','━━ ━━'],['lgDashDot','━━ • ━━ •'],['lgDashDotDot','━━ • •'],['sysDash','– – –'],['sysDot','· · ·'],['sysDashDot','– · – ·'],['sysDashDotDot','– · ·']] as const){const option=element('option',label);option.value=value;dash.append(option);}field('slideStrokePattern',dash,paintGroup);dash.onchange=()=>safe(()=>updateStroke({dash:dash.value as SlideDash}));
  function arrowSelect(end:'head'|'tail'):HTMLSelectElement{
    const select=element('select');select.id='slide-arrow-'+end;for(const [value,label]of [['none',t('slideListNone')],['triangle','▶'],['arrow','➝']] as const){const option=element('option',label);option.value=value;select.append(option);}field(end==='head'?'slideArrowHead':'slideArrowTail',select,paintGroup);
    select.onchange=()=>safe(()=>{const type=select.value;if(type==='none')updateStroke({},end);else updateStroke({[end]:{...currentStroke()[end],type:type as 'triangle'|'arrow'}});});return select;
  }
  const arrowHead=arrowSelect('head'),arrowTail=arrowSelect('tail');
  const cropGroup=element('fieldset');cropGroup.append(element('legend',t('slideImageCrop')));inspector.append(cropGroup);
  const cropFields={} as Record<keyof SlideCrop,HTMLInputElement>;
  for(const [key,label]of [['left','slideCropLeft'],['top','slideCropTop'],['right','slideCropRight'],['bottom','slideCropBottom']] as const){const control=element('input');control.id='slide-crop-'+key;control.type='number';control.min='0';control.max='99.999';control.step='.1';field(label,control,cropGroup);cropFields[key]=control;}
  const cropApply=button('slideApplyCrop',()=>{const shape=selected()[0];if(shape?.type==='image'){const values=Object.fromEntries(Object.entries(cropFields).map(([k,v])=>[k,Math.round(Number(v.value)*1000)])) as unknown as SlideCrop;apply(setSlideImageCrop(snap(),state.slideId!,shape.id,values));}},cropGroup,'slide-crop-apply',true);
  const cropReset=button('slideResetCrop',()=>{const shape=selected()[0];if(shape?.type==='image')apply(setSlideImageCrop(snap(),state.slideId!,shape.id,{left:0,top:0,right:0,bottom:0}));},cropGroup,'slide-crop-reset',true);
  const textGroup=element('fieldset');textGroup.append(element('legend',t('slideTextEditor')));inspector.append(textGroup);
  const textEditor=element('textarea');textEditor.id='slide-text-editor';textEditor.rows=5;textEditor.dir='auto';textEditor.spellcheck=false;field('slideTextEditor',textEditor,textGroup);
  textEditor.oninput=()=>{textDirty=true;if(host.composing())return;if(textTimer!==undefined)clearTimeout(textTimer);textTimer=setTimeout(()=>{textTimer=undefined;safe(()=>commitText());},300);};
  textEditor.addEventListener('compositionstart',()=>{host.setComposing(true);});textEditor.addEventListener('compositionend',()=>{host.setComposing(false);textDirty=true;queueMicrotask(()=>safe(()=>commitText()));});textEditor.onblur=()=>safe(()=>{if(!host.composing())commitText();});
  function format(style:TextStyle):void {const shape=selectedText();if(!shape)return;const range=style.alignment===undefined&&editingId===shape.id?{start:textEditor.selectionStart,end:textEditor.selectionEnd}:undefined;apply(formatShapeText(snap(),state.slideId!,shape.id,style,range,id));}
  const textTools=element('div','','slides-tool-buttons');textGroup.append(textTools);
  for(const [property,key] of [['bold','writerBold'],['italic','writerItalic'],['underline','writerUnderline']] as const)button(key,()=>{const shape=selectedText();if(shape)format({[property]:!shape.textStyle?.[property]});},textTools,'slide-text-'+property,true);
  for(const [alignment,key] of [['start','writerAlignLeft'],['center','writerAlignCenter'],['end','writerAlignRight']] as const)button(key,()=>format({alignment}),textTools,'slide-text-'+alignment,true);
  const font=element('input');font.id='slide-font-family';font.type='text';font.maxLength=512;field('writerFontFamily',font,textGroup);font.onchange=()=>safe(()=>{if(flush()&&font.value.trim())format({fontFamily:font.value.trim()});});
  const size=element('input');size.id='slide-font-size';size.type='number';size.min='1';size.max='1638';size.step='.5';field('writerFontSize',size,textGroup);size.onchange=()=>safe(()=>{if(flush())format({fontSize:emu(size.value)});});
  const textColor=element('input');textColor.id='slide-font-color';textColor.type='color';field('writerTextColor',textColor,textGroup);textColor.onchange=()=>safe(()=>{if(flush())format({color:rgb(textColor.value)});});
  const language=element('input');language.type='text';language.maxLength=128;field('writerTextLanguage',language,textGroup);language.onchange=()=>safe(()=>{if(flush())format({language:language.value.trim()});});
  const direction=element('select');for(const [value,key] of [['ltr','writerLtr'],['rtl','writerRtl']] as const){const option=element('option',t(key));option.value=value;direction.append(option);}field('writerParagraphDirection',direction,textGroup);direction.onchange=()=>safe(()=>{const shape=selectedText();if(shape&&flush())apply(setShapeTextFrame(snap(),state.slideId!,shape.id,{direction:direction.value as 'ltr'|'rtl'}));});
  const wrap=element('input');wrap.type='checkbox';field('sheetWrap',wrap,textGroup);wrap.onchange=()=>safe(()=>{const shape=selectedText();if(shape&&flush())apply(setShapeTextFrame(snap(),state.slideId!,shape.id,{wrap:wrap.checked}));});
  const vertical=element('select');for(const [value,key] of [['top','slideAlignTop'],['center','slideAlignMiddle'],['bottom','slideAlignBottom']] as const){const option=element('option',t(key));option.value=value;vertical.append(option);}field('slideAlignMiddle',vertical,textGroup);vertical.onchange=()=>safe(()=>{const shape=selectedText();if(shape&&flush())apply(setShapeTextFrame(snap(),state.slideId!,shape.id,{verticalAlignment:vertical.value as 'top'|'center'|'bottom'}));});
  const listType=element('select');listType.id='slide-list-type';for(const [value,key]of [['none','slideListNone'],['char','slideListBullets'],['number','slideListNumbers']] as const){const option=element('option',t(key));option.value=value;listType.append(option);}field('slideParagraphList',listType,textGroup);
  const listLevel=element('select');listLevel.id='slide-list-level';for(let n=1;n<=9;n++){const option=element('option',host.number(n));option.value=String(n);listLevel.append(option);}field('slideListLevel',listLevel,textGroup);
  function paragraphFormat():DeepReadonly<SlideParagraphFormat>|undefined{
    const shape=selectedText();if(!shape)return undefined;const at=textEditor.selectionStart;let offset=0;
    for(const run of shape.runs){const end=offset+run.text.length;if(at<end||at===end&&run.text.length&&run.text.endsWith('\n'))return run.style?.slideParagraph??shape.textStyle?.slideParagraph;offset=end;}
    return shape.runs.at(-1)?.style?.slideParagraph??shape.textStyle?.slideParagraph;
  }
  function applyList(kind:string,level:number):void{
    if(host.legacyPresentation||!flush())return;const shape=selectedText();if(!shape)return;
    if(!Number.isSafeInteger(level)||level<0||level>8)throw new OfficeError('invalid','Invalid paragraph list level');
    const prior=paragraphFormat(),margin=kind==='none'?0:342900+level*457200;
    const bullet:SlideParagraphFormat['bullet']=kind==='char'?(prior?.bullet?.kind==='char'?structuredClone(prior.bullet):{kind:'char',character:'•'}):kind==='number'?(prior?.bullet?.kind==='number'?structuredClone(prior.bullet):{kind:'number',scheme:'arabicPeriod'}):{kind:'none'};
    apply(formatSlideParagraphs(snap(),state.slideId!,shape.id,{start:textEditor.selectionStart,end:textEditor.selectionEnd},{level:kind==='none'?0:level,marginLeft:margin,firstLineIndent:kind==='none'?0:-228600,bullet},id));
  }
  listType.onchange=()=>safe(()=>applyList(listType.value,Math.max(0,Number(listLevel.value)-1)));
  listLevel.onchange=()=>safe(()=>applyList(listType.value,Number(listLevel.value)-1));
  const fitButton=button('slideFitText',()=>{
    const shape=selectedText();if(!shape||host.legacyPresentation)return;const revision=snap().revision,shapeId=shape.id,slideId=state.slideId!;
    void document.fonts.ready.then(()=>safe(()=>{
      if(disposed||!host.isCurrent()||snap().revision!==revision||state.slideId!==slideId||selectedText()?.id!==shapeId)return;
      const candidate=[...walkSlideScene(buildSlideScene(snap(),slideId).shapes)].find(s=>s.id===shapeId);if(candidate?.type!=='text')return;
      const fontScale=measureSlideTextFit(candidate);if(!fontScale)throw new OfficeError('limit','Text cannot fit at the minimum supported scale');
      apply(setShapeTextFrame(snap(),slideId,shapeId,{autofit:{mode:'normal',fontScale,lineReduction:candidate.autofit?.mode==='normal'?candidate.autofit.lineReduction:0}}));
    }));
  },textGroup,'slide-fit-text',true);
  const resetFit=button('slideResetFit',()=>{const shape=selectedText();if(shape&&!host.legacyPresentation)apply(setShapeTextFrame(snap(),state.slideId!,shape.id,{autofit:{mode:'none'}}));},textGroup,'slide-reset-fit',true);
  function syncListControls():void{
    const style=paragraphFormat(),disabled=!!host.readOnly||!!host.legacyPresentation||!selectedText();listType.disabled=disabled;listLevel.disabled=disabled;fitButton.disabled=disabled;resetFit.disabled=disabled;
    if(document.activeElement!==listType)listType.value=style?.bullet?.kind??'none';if(document.activeElement!==listLevel)listLevel.value=String((style?.level??0)+1);
  }
  textEditor.addEventListener('select',syncListControls);
  const slideGroup=element('fieldset');slideGroup.append(element('legend',t('slides')));inspector.append(slideGroup);
  const background=element('input');background.type='color';field('slideBackground',background,slideGroup);background.onchange=()=>safe(()=>{if(flush())apply(setSlideBackground(snap(),state.slideId!,rgb(background.value)));});
  const themeAccent=element('input');themeAccent.type='color';field('slideThemeAccent',themeAccent,slideGroup);themeAccent.onchange=()=>safe(()=>{if(!flush())return;const theme=doc().theme??DEFAULT_SLIDE_THEME;apply(setSlideTheme(snap(),{...theme,colors:{...theme.colors,accent1:rgb(themeAccent.value) as Extract<Color,{kind:'rgba'}>}}));});
  const slideWidth=element('input');slideWidth.type='number';slideWidth.min='1';slideWidth.max='10000';slideWidth.step='.1';field('writerPageWidth',slideWidth,slideGroup);
  const slideHeight=element('input');slideHeight.type='number';slideHeight.min='1';slideHeight.max='10000';slideHeight.step='.1';field('writerPageHeight',slideHeight,slideGroup);
  button('writerApplyPage',()=>apply(setSlideSize(snap(),emu(slideWidth.value),emu(slideHeight.value))),slideGroup,'slide-size-apply');

  let lastInspectorSelection:string|undefined;
  const ribbonTextControls=[font,size,textColor,...Array.from(textTools.querySelectorAll<HTMLButtonElement>('button'))];
  const inspectorHead=element('div','','office-properties-head'),inspectorTitle=element('strong',t('ribbonProperties'));
  const collapse=element('button',t('close'));collapse.type='button';collapse.id='slide-properties-close';decorateOfficeButton(collapse,'close',true);inspectorHead.append(inspectorTitle,collapse);
  const inspectorTabs=element('div','','office-properties-tabs'),propertyButtons=new Map<string,HTMLButtonElement>();
  for(const [mode,key] of [['object','ribbonObject'],['text','ribbonText'],['slide','ribbonSlide']] as const){
    const b=element('button',t(key));b.type='button';b.dataset.propertyTab=mode;b.onclick=()=>safe(()=>{if(flush()){state.inspectorMode=mode;showInspector();}});inspectorTabs.append(b);propertyButtons.set(mode,b);
  }
  inspector.prepend(inspectorHead,inspectorTabs);
  geometryGroup.dataset.propertyPanel=paintGroup.dataset.propertyPanel=cropGroup.dataset.propertyPanel='object';textGroup.dataset.propertyPanel='text';slideGroup.dataset.propertyPanel='slide';
  const propertiesButton=element('button',t('ribbonProperties'));propertiesButton.type='button';propertiesButton.id='slide-properties-toggle';
  propertiesButton.onclick=()=>safe(()=>{if(!flush())return;state.inspectorCollapsed=!state.inspectorCollapsed;showInspector();resize();});
  collapse.onclick=()=>safe(()=>{if(!flush())return;state.inspectorCollapsed=true;showInspector();resize();});
  state.inspectorMode??='slide';
  function focusTextEditor():void{state.inspectorMode='text';state.inspectorCollapsed=false;showInspector();resize();textEditor.focus();}
  function showInspector():void{
    persistPanelCollapsed('slide-inspector',!!state.inspectorCollapsed);inspector.hidden=!!state.inspectorCollapsed;rightSplitter.hidden=!!state.inspectorCollapsed;layoutPanels();
    body.classList.toggle('properties-collapsed',!!state.inspectorCollapsed);
    propertiesButton.setAttribute('aria-pressed',String(!state.inspectorCollapsed));
    for(const [key,b] of propertyButtons)b.setAttribute('aria-pressed',String(key===state.inspectorMode));
    geometryGroup.hidden=state.inspectorMode!=='object';
    paintGroup.hidden=state.inspectorMode!=='object';
    textGroup.hidden=state.inspectorMode!=='text';
    slideGroup.hidden=state.inspectorMode!=='slide';
    cropGroup.hidden=state.inspectorMode!=='object'||selected()[0]?.type!=='image';
  }
  function syncInspector():void {
    const m=doc(),current=locateSlide(snap(),state.slideId!);const scope=scopeShapes();state.selectedIds=state.selectedIds.filter(shapeId=>scope.some(s=>s.id===shapeId));const shapes=selected(),single=shapes.length===1?shapes[0]:undefined,text=single?.type==='text'?single:undefined;
    for(const control of editControls)control.disabled=!!host.readOnly;for(const control of selectionControls)control.disabled=!!host.readOnly||!shapes.length;
    deletePage.disabled=!!host.readOnly||m.slides.length<2;earlier.disabled=!!host.readOnly||current.index===0;later.disabled=!!host.readOnly||current.index===m.slides.length-1;
    const attached=single?.type==='line'&&Boolean(single.connections),displayGeometry=single?([...walkSlideScene(scene.shapes)].find(s=>s.id===single.id)?.geometry??single.geometry):undefined;
    for(const [key,control] of Object.entries(geometryFields)){control.disabled=!!host.readOnly||!single||attached;control.value=displayGeometry?String(Math.round(displayGeometry[key as keyof Geometry]/(key==='rotation'?ROTATION_UNITS_PER_DEGREE:EMU_PER_POINT)*100)/100):'';}
    aspect.checked=single?.type==='image'||!!state.lockAspect;aspect.disabled=single?.type==='image'||!!host.readOnly;
    fill.value=colorHex(single?.fill);stroke.value=colorHex(single?.stroke?.color,'#182030');strokeWidth.value=String((single?.stroke?.width??12700)/EMU_PER_POINT);
    const graphicsDisabled=!!host.readOnly||!!host.legacyPresentation;
    insertTableButton.disabled=graphicsDisabled;editTableButton.disabled=graphicsDisabled||single?.type!=='table';
    distributeH.disabled=distributeV.disabled=!!host.readOnly||shapes.length<3||shapes.some(s=>s.geometry.rotation!==0||s.type==='line'&&!!s.connections);
    geometryFields.rotation.disabled||=single?.type==='table';
    for(const control of newGraphics)control.disabled=graphicsDisabled;
    const selectedIndices=scope.flatMap((shape,i)=>state.selectedIds.includes(shape.id)?[i]:[]);
    groupButton.disabled=graphicsDisabled||selectedIndices.length<2||selectedIndices.at(-1)!-selectedIndices[0]!+1!==selectedIndices.length;
    ungroupButton.disabled=graphicsDisabled||!single||!canUngroupShape(single);
    for(const control of flipButtons)control.disabled=graphicsDisabled||!shapes.length||shapes.some(shape=>shape.type==='text'||shape.type==='table'||shape.type==='line'&&shape.connections);
    editGroup.disabled=graphicsDisabled||single?.type!=='group';exitGroup.disabled=graphicsDisabled||!state.groupId;connectButton.disabled=graphicsDisabled||!!state.groupId||shapes.length!==2||!shapes.every(isConnectorTarget);detachButton.disabled=graphicsDisabled||!attached;
    groupButton.disabled||=shapes.some(s=>s.type==='table');
    groupButton.disabled||=scope.some(s=>s.type==='line'&&s.connections&&(state.selectedIds.includes(s.id)||state.selectedIds.includes(s.connections.start.shapeId)||state.selectedIds.includes(s.connections.end.shapeId)));
    const options=document.createDocumentFragment(),placeholder=element('option','—');placeholder.value='';options.append(placeholder);
    const names:Record<Shape['type'],MessageKey>={table:'slideTable',text:'slideText',rect:'slideRectangle',ellipse:'slideEllipse',line:'slideLine',image:'writerInsertImage',group:'slideGroupObjects',roundRect:'slideRoundRect',triangle:'slideTriangle',diamond:'slideDiamond',parallelogram:'slideParallelogram'};
    const addOptions=(nodes:readonly ShapeSnapshot[],depth:number)=>{for(const [index,shape]of nodes.entries()){const label=t(names[shape.type])+' '+host.number(index+1)+(shape.type==='text'?' — '+shapeText(shape).slice(0,32):'');const option=element('option','　'.repeat(depth)+label);option.value=shape.id;options.append(option);if(shape.type==='group')addOptions(shape.children,depth+1);}};
    addOptions(current.slide.shapes,0);objectTree.replaceChildren(options);objectTree.value=single?.id??'';objectTree.disabled=!!host.readOnly||!current.slide.shapes.length;

    fill.disabled=!!host.readOnly||!single||single.type==='group'||single.type==='table';
    stroke.disabled=strokeWidth.disabled=!!host.readOnly||!single||single.type==='group'||single.type==='table';
    dash.disabled=graphicsDisabled||!single||['text','image','group','table'].includes(single.type);
    arrowHead.disabled=arrowTail.disabled=graphicsDisabled||single?.type!=='line';
    dash.value=single?.stroke?.dash??'solid';arrowHead.value=single?.stroke?.head?.type??'none';arrowTail.value=single?.stroke?.tail?.type??'none';
    cropGroup.hidden=state.inspectorMode!=='object'||single?.type!=='image';cropGroup.disabled=graphicsDisabled||single?.type!=='image';cropApply.disabled=cropReset.disabled=cropGroup.disabled;
    if(single?.type==='image')for(const key of Object.keys(cropFields) as (keyof SlideCrop)[])cropFields[key].value=String((single.crop?.[key]??0)/1000);
    textGroup.disabled=!text||!!host.readOnly;
    for(const c of ribbonTextControls)c.disabled=!text||!!host.readOnly;
    if(state.inspectorMode!==undefined&&state.selectedIds.join(',')!==lastInspectorSelection){
      lastInspectorSelection=state.selectedIds.join(',');state.inspectorMode=text?'text':single||shapes.length?'object':'slide';showInspector();
    }

    if(text){if(editingId!==text.id||!textDirty){const focused=document.activeElement===textEditor,start=textEditor.selectionStart,end=textEditor.selectionEnd;textEditor.value=shapeText(text);if(focused)textEditor.setSelectionRange(Math.min(start,textEditor.value.length),Math.min(end,textEditor.value.length));}editingId=text.id;const style={...m.defaultTextStyle,...text.textStyle,...text.runs[0]?.style};font.value=style.fontFamily??'sans-serif';size.value=String((style.fontSize??304800)/EMU_PER_POINT);textColor.value=colorHex(style.color,'#182030');language.value=style.language??'und';direction.value=text.textFrame?.direction==='rtl'?'rtl':'ltr';wrap.checked=text.textFrame?.wrap??true;vertical.value=text.textFrame?.verticalAlignment??'top';textEditor.dir=text.textFrame?.direction??style.paragraphDirection??'auto';}
    else {editingId=undefined;textEditor.value='';textDirty=false;}
    syncListControls();
    background.value=colorHex(current.slide.background,'#ffffff');themeAccent.value=colorHex(m.theme?.colors.accent1??DEFAULT_SLIDE_THEME.colors.accent1);slideWidth.value=String(Math.round(m.size.width/EMU_PER_POINT*100)/100);slideHeight.value=String(Math.round(m.size.height/EMU_PER_POINT*100)/100);
    position.textContent=t('slidePageNumber',{current:host.number(current.index+1),total:host.number(m.slides.length)});for(const thumb of rail.querySelectorAll<HTMLButtonElement>('[data-slide-id]')){thumb.setAttribute('aria-current',String(thumb.dataset.slideId===state.slideId));thumb.classList.toggle('is-current',thumb.dataset.slideId===state.slideId);}
  }
  function checkOverflow():void {if(disposed||!sceneElement?.isConnected||scene.revision!==snap().revision)return;const overflowIds=overflowingSlideText(sceneElement);overflow.hidden=overflowIds.length===0;overflow.textContent=overflowIds.length?t('slideTextOverflow'):'';for(const node of sceneElement.querySelectorAll<HTMLElement>('[data-shape-id]'))node.classList.toggle('slides-text-overflow',overflowIds.includes(node.dataset.shapeId!));}
  function refreshScene():void {
    const snapshot=snap(),m=presentation(snapshot);if(!m.slides.length)throw new OfficeError('invalid','Presentation has no slides');if(!m.slides.some(s=>s.id===state.slideId))state.slideId=m.slides[0]!.id;
    activeGroup();scene=buildSlideScene(snapshot,state.slideId!);sceneElement=createSlideSceneElement(scene);sceneElement.id='slide-canvas';scaled.replaceChildren(sceneElement,overlay);resize();paintGeometry();requestAnimationFrame(checkOverflow);void document.fonts.ready.then(()=>{if(!disposed)checkOverflow();});
  }
  // Stable thumbnail shells retain scroll/focus/order; only visible
  // preview content is retained. A local shape edit must not rebuild
  // every thumbnail or accumulate a whole presentation's DOM.
  type ThumbSource=ReturnType<typeof presentation>['slides'][number];
  type ThumbEntry={button:HTMLButtonElement;image:HTMLDivElement;label:HTMLSpanElement;visible:boolean;source?:ThumbSource;globalVersion?:number};
  const thumbnails=new Map<string,ThumbEntry>();
  let thumbnailSnapshot:DocumentSnapshot|undefined,thumbnailGlobals:unknown[]|undefined,thumbnailVersion=0,thumbnailRenderCount=0;
  function renderThumbnail(entry:ThumbEntry){
    if(disposed||!thumbnailSnapshot||!entry.visible)return;
    const snapshot=thumbnailSnapshot,m=presentation(snapshot),slideId=entry.button.dataset.slideId!,source=m.slides.find(s=>s.id===slideId);
    if(!source)return;
    if(entry.source===source&&entry.globalVersion===thumbnailVersion&&entry.image.firstElementChild){
      // The projection is identical, but expose the current revision
      // for accessibility/debugging and existing revision checks.
      const node=entry.image.firstElementChild as HTMLElement;
      node.dataset.revision=String(snapshot.revision);node.dataset.contentId=snapshot.contentId;
      return;
    }
    const preview=buildSlideScene(snapshot,slideId),node=createSlideSceneElement(preview),ratio=128/(preview.width/EMU_PER_CSS_PIXEL);
    node.style.transform='scale('+ratio+')';node.style.transformOrigin='top left';entry.image.replaceChildren(node);
    entry.source=source;entry.globalVersion=thumbnailVersion;thumbnailRenderCount++;
    rail.dataset.thumbnailRenders=String(thumbnailRenderCount);
  }
  function refreshThumbnails():void {
    const snapshot=snap(),m=presentation(snapshot);thumbnailSnapshot=snapshot;
    const globals=[m.size,m.theme,m.defaultTextStyle,m.styles,m.resources];
    if(!thumbnailGlobals||globals.some((v,i)=>v!==thumbnailGlobals![i])){thumbnailGlobals=globals;thumbnailVersion++;}
    if(!thumbObserver&&typeof IntersectionObserver!=='undefined')thumbObserver=new IntersectionObserver(entries=>{
      if(disposed)return;
      for(const observation of entries){
        const button=observation.target as HTMLButtonElement,entry=thumbnails.get(button.dataset.slideId!);
        if(!entry||entry.button!==button)continue;
        entry.visible=observation.isIntersecting;
        if(entry.visible)renderThumbnail(entry);
        else{entry.image.replaceChildren();entry.source=undefined;entry.globalVersion=undefined;}
      }
    },{root:rail,rootMargin:'200px'});
    const wanted=new Set(m.slides.map(s=>s.id));
    for(const [id,entry]of thumbnails)if(!wanted.has(id)){thumbObserver?.unobserve(entry.button);entry.button.remove();thumbnails.delete(id);}
    for(const [index,slide]of m.slides.entries()){
      const slideId=slide.id;let entry=thumbnails.get(slideId);
      if(!entry){
        const thumb=element('button','','slides-thumbnail');thumb.type='button';thumb.dataset.slideId=slideId;
        const image=element('div','','slides-thumbnail-image'),label=element('span');image.setAttribute('aria-hidden','true');
        thumb.append(image,label);
        thumb.onclick=()=>safe(()=>changeSlide(slideId));
        thumb.ondragstart=event=>{if(!flush()){event.preventDefault();return;}event.dataTransfer?.setData('application/x-office-ex-slide',slideId);if(event.dataTransfer)event.dataTransfer.effectAllowed='move';};
        thumb.ondragover=event=>{if(!host.readOnly&&event.dataTransfer?.types.includes('application/x-office-ex-slide'))event.preventDefault();};
        thumb.ondrop=event=>{event.preventDefault();const source=event.dataTransfer?.getData('application/x-office-ex-slide'),target=doc().slides.findIndex(s=>s.id===slideId);if(source&&target>=0&&doc().slides.some(s=>s.id===source))safe(()=>{if(flush())apply(moveSlide(snap(),source,target));});};
        entry={button:thumb,image,label,visible:!thumbObserver};thumbnails.set(slideId,entry);rail.append(thumb);thumbObserver?.observe(thumb);
      }
      const {button:thumb,image,label}=entry;
      thumb.draggable=!host.readOnly;thumb.setAttribute('aria-label',t('slidePageNumber',{current:host.number(index+1),total:host.number(m.slides.length)}));thumb.setAttribute('aria-current',String(slideId===state.slideId));thumb.classList.toggle('is-current',slideId===state.slideId);
      image.style.width='128px';image.style.height=128*m.size.height/m.size.width+'px';label.textContent=host.number(index+1);
      if(rail.children[index]!==thumb)rail.insertBefore(thumb,rail.children[index]??null);
      renderThumbnail(entry);
    }
  }
  function refresh():void {if(disposed)return;refreshScene();refreshThumbnails();syncInspector();}
  function resize():void {
    layoutPanels();
    if(!scene)return;const width=scene.width/EMU_PER_CSS_PIXEL,height=scene.height/EMU_PER_CSS_PIXEL,style=getComputedStyle(stage),
      innerWidth=Math.max(1,stage.clientWidth-parseFloat(style.paddingLeft)-parseFloat(style.paddingRight)-2),
      innerHeight=Math.max(1,stage.clientHeight-parseFloat(style.paddingTop)-parseFloat(style.paddingBottom)-2),
      fit=Math.max(.05,Math.min(1,innerWidth/width,innerHeight/height));zoom=state.zoom>0?state.zoom:fit;surface.style.width=width*zoom+'px';surface.style.height=height*zoom+'px';scaled.style.width=width+'px';scaled.style.height=height+'px';scaled.style.transform=`scale(${zoom})`;stage.dataset.zoom=String(zoom);paintGeometry();
  }
  function paintGeometry():void {
    if(!sceneElement)return;const values=gesture?.preview??{},projected=new Map([...walkSlideScene(scene.shapes)].map(s=>[s.id,s]));
    const roots=locateSlide(snap(),state.slideId!).slide.shapes,siblings=new Map(roots.map(s=>[s.id,values[s.id]?{...s,geometry:values[s.id]!}:s]));
    for(const node of sceneElement.querySelectorAll<HTMLElement>('[data-shape-id]')){
      let shape=projected.get(node.dataset.shapeId!);if(!shape)continue;
      const original=siblings.get(shape.id);
      if(original?.type==='line'&&original.connections){const current=resolveSlideConnector(original,siblings);shape={...shape,geometry:current.geometry,flipH:current.flipH,flipV:current.flipV};}
      updateSlideShapeGeometry(node,shape,values[shape.id]??shape.geometry);projected.set(shape.id,shape);
    }
    overlay.replaceChildren();const parent=scopeMatrix();
    for(const shapeId of state.selectedIds){const shape=projected.get(shapeId);if(!shape)continue;const g=values[shapeId]??shape.geometry,box=element('div','','slides-selection-box');box.dataset.selectedShape=shapeId;
      const matrix=multiply(parent,shapeTransform(g));Object.assign(box.style,{left:'0',top:'0',width:g.width/EMU_PER_CSS_PIXEL+'px',height:g.height/EMU_PER_CSS_PIXEL+'px',transformOrigin:'0 0',transform:`matrix(${matrix[0]},${matrix[1]},${matrix[2]},${matrix[3]},${matrix[4]/EMU_PER_CSS_PIXEL},${matrix[5]/EMU_PER_CSS_PIXEL})`,borderWidth:1.5/zoom+'px'});overlay.append(box);
      if(state.selectedIds.length===1&&!host.readOnly&&!selectedHasBindings())for(const corner of (shape.type==='table'?['nw','ne','se','sw']:['nw','ne','se','sw','rotate'])){const handle=element('span','','slides-handle slides-handle-'+corner);handle.dataset.handle=corner;handle.dataset.shapeId=shapeId;handle.style.width=9/zoom+'px';handle.style.height=9/zoom+'px';handle.style.borderWidth=1/zoom+'px';if(corner==='rotate')handle.style.top=-24/zoom+'px';handle.setAttribute('aria-hidden','true');box.append(handle);}
    }
  }
  function point(event:PointerEvent|MouseEvent):Point {const rect=sceneElement.getBoundingClientRect();return transform(inverse(scopeMatrix()),{x:(event.clientX-rect.left)/zoom*EMU_PER_CSS_PIXEL,y:(event.clientY-rect.top)/zoom*EMU_PER_CSS_PIXEL});}
  function scopeHit(p:Point,tolerance:number):string|undefined{
    const shapes=scopeShapes(),map=new Map(shapes.map(s=>[s.id,s]));return hitTest(shapes.map(s=>resolveSlideConnector(s,map)),p,tolerance);
  }
  stage.onpointerdown=event=>safe(()=>{
    if(event.button!==0||host.readOnly||host.composing())return;const target=event.target as HTMLElement,handle=target.closest<HTMLElement>('[data-handle]'),p=point(event);if(!flush())return;
    if(handle){const shape=locateShape(snap(),state.slideId!,handle.dataset.shapeId!).shape,g={...shape.geometry};gesture={pointerId:event.pointerId,revision:snap().revision,kind:handle.dataset.handle==='rotate'?'rotate':'resize',start:p,shapeId:shape.id,corner:handle.dataset.handle,geometries:{[shape.id]:g},preview:{[shape.id]:g},angle:Math.atan2(p.y-g.y-g.height/2,p.x-g.x-g.width/2)};}
    else {const shapeId=scopeHit(p,6/zoom*EMU_PER_CSS_PIXEL);if(!shapeId){choose([]);stage.focus();return;}if(event.shiftKey){choose(state.selectedIds.includes(shapeId)?state.selectedIds.filter(value=>value!==shapeId):[...state.selectedIds,shapeId]);return;}if(!state.selectedIds.includes(shapeId))choose([shapeId]);if(selectedHasBindings()){stage.focus({preventScroll:true});return;}const geometries=Object.fromEntries(selected().map(s=>[s.id,{...s.geometry}]));gesture={pointerId:event.pointerId,revision:snap().revision,kind:'move',start:p,geometries,preview:geometries};}
    event.preventDefault();stage.focus({preventScroll:true});stage.setPointerCapture(event.pointerId);
  });
  stage.onpointermove=event=>safe(()=>{
    if(!gesture||gesture.pointerId!==event.pointerId)return;const current=gesture,p=point(event),delta={x:p.x-current.start.x,y:p.y-current.start.y};
    if(current.kind==='move')current.preview=Object.fromEntries(Object.entries(current.geometries).map(([key,g])=>[key,{...g,x:Math.round(g.x+delta.x),y:Math.round(g.y+delta.y)}]));
    else {const g=current.geometries[current.shapeId!]!;
      if(current.kind==='rotate'){const angle=Math.atan2(p.y-g.y-g.height/2,p.x-g.x-g.width/2),step=event.shiftKey?15:1,rotation=Math.round((g.rotation/ROTATION_UNITS_PER_DEGREE+(angle-current.angle!)*180/Math.PI)/step)*step*ROTATION_UNITS_PER_DEGREE;current.preview={[current.shapeId!]:{...g,rotation}};}
      else {const local=transform(inverse(shapeTransform(g)),p),start=transform(inverse(shapeTransform(g)),current.start),east=current.corner!.includes('e'),south=current.corner!.includes('s'),min=EMU_PER_CSS_PIXEL;let width=Math.max(min,g.width+(local.x-start.x)*(east?1:-1)),height=Math.max(min,g.height+(local.y-start.y)*(south?1:-1));const shape=locateShape(snap(),state.slideId!,current.shapeId!).shape;if((state.lockAspect||shape.type==='image'||event.shiftKey)&&g.width>0&&g.height>0){const ratio=Math.max(width/g.width,height/g.height);width=g.width*ratio;height=g.height*ratio;}width=Math.round(width);height=Math.round(height);const fixed=transform(shapeTransform(g),{x:east?0:g.width,y:south?0:g.height}),angle=g.rotation/ROTATION_UNITS_PER_DEGREE*Math.PI/180,u={x:Math.cos(angle),y:Math.sin(angle)},v={x:-Math.sin(angle),y:Math.cos(angle)},fx=east?-.5:.5,fy=south?-.5:.5,cx=fixed.x-u.x*fx*width-v.x*fy*height,cy=fixed.y-u.y*fx*width-v.y*fy*height;current.preview={[current.shapeId!]:{...g,x:Math.round(cx-width/2),y:Math.round(cy-height/2),width,height}};}
    }
    if(current.frame===undefined)current.frame=requestAnimationFrame(()=>{current.frame=undefined;if(gesture===current)paintGeometry();});
  });
  stage.onpointerup=event=>safe(()=>{if(!gesture||gesture.pointerId!==event.pointerId)return;const current=gesture;gesture=undefined;if(current.frame!==undefined)cancelAnimationFrame(current.frame);if(stage.hasPointerCapture(event.pointerId))stage.releasePointerCapture(event.pointerId);if(current.revision!==snap().revision){refresh();return;}if(Object.entries(current.preview).every(([id,g])=>JSON.stringify(g)===JSON.stringify(current.geometries[id]))){paintGeometry();return;}apply(setShapeGeometries(snap(),state.slideId!,current.preview));});
  stage.onpointercancel=()=>cancelGesture();stage.onlostpointercapture=()=>{if(gesture)cancelGesture();};stage.ondblclick=event=>safe(()=>{const shapeId=scopeHit(point(event),4/zoom*EMU_PER_CSS_PIXEL);if(shapeId){choose([shapeId]);const shape=selected()[0];if(shape?.type==='group')enterGroup(shape);else if(shape?.type==='table'&&!host.readOnly)openTable(true);else if(selectedText()){focusTextEditor();textEditor.setSelectionRange(0,textEditor.value.length);}}});
  root.addEventListener('keydown',event=>safe(()=>{
    if(tableDialog.open)return;
    if(event.key==='Escape'&&gesture){event.preventDefault();event.stopPropagation();cancelGesture();return;}if(host.composing()||isTyping(event.target)||host.readOnly)return;
    if(event.key==='Escape'){if(state.selectedIds.length)choose([]);else leaveGroup();event.preventDefault();return;}
    if(event.key==='F5'){event.preventDefault();if(flush())present();return;}
    if((event.ctrlKey||event.metaKey)&&event.key.toLowerCase()==='a'){choose(scopeShapes().map(s=>s.id));event.preventDefault();return;}
    if((event.ctrlKey||event.metaKey)&&event.key.toLowerCase()==='d'){event.preventDefault();if(flush()&&state.selectedIds.length)duplicateSelection();return;}
    if(!state.selectedIds.length)return;if(['Delete','Backspace'].includes(event.key)){event.preventDefault();if(flush()&&apply(removeShapes(snap(),state.slideId!,state.selectedIds)))choose([]);return;}
    if(event.key==='Enter'&&state.selectedIds.length===1&&selected()[0]?.type==='table'){event.preventDefault();openTable(true);return;}
    if(event.key==='Enter'&&selectedText()){event.preventDefault();focusTextEditor();return;}
    if(['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(event.key)){event.preventDefault();if(!flush()||selectedHasBindings())return;const step=(event.shiftKey?10:1)*EMU_PER_CSS_PIXEL,dx=event.key==='ArrowLeft'?-step:event.key==='ArrowRight'?step:0,dy=event.key==='ArrowUp'?-step:event.key==='ArrowDown'?step:0;apply(setShapeGeometries(snap(),state.slideId!,Object.fromEntries(selected().map(s=>[s.id,{...s.geometry,x:s.geometry.x+dx,y:s.geometry.y+dy}]))));}
  }));
  root.addEventListener('copy',event=>safe(()=>{if(isTyping(event.target)||!state.selectedIds.length||!flush())return;const source=scopeShapes(),siblings=new Map(source.map(s=>[s.id,s])),shapes=selected().map(shape=>resolveSlideConnector(shape,siblings)),resources:Record<string,DeepReadonly<ImageResource>>={};for(const shape of walkSlideShapes(shapes))if(shape.type==='image'){const r=doc().resources[shape.resourceId];if(r?.dataBase64)resources[shape.resourceId]=r as DeepReadonly<ImageResource>;}copiedShapes={nonce:id(),shapes,resources};event.clipboardData?.setData('application/x-office-ex-shapes',copiedShapes.nonce);event.clipboardData?.setData('text/plain',shapes.filter(s=>s.type==='text').map(shapeText).join('\n'));event.preventDefault();}));
  root.addEventListener('cut',event=>safe(()=>{if(isTyping(event.target)||host.readOnly||!state.selectedIds.length||!flush())return;const copy=new ClipboardEvent('copy',{bubbles:true,clipboardData:event.clipboardData});stage.dispatchEvent(copy);if(copiedShapes&&event.clipboardData?.getData('application/x-office-ex-shapes')===copiedShapes.nonce){event.preventDefault();if(apply(removeShapes(snap(),state.slideId!,state.selectedIds)))choose([]);}}));
  root.addEventListener('paste',event=>safe(()=>{if(isTyping(event.target)||!copiedShapes||event.clipboardData?.getData('application/x-office-ex-shapes')!==copiedShapes.nonce||host.readOnly||!flush())return;event.preventDefault();if(host.legacyPresentation&&copiedShapes.shapes.some(s=>!['rect','ellipse','line','image','text'].includes(s.type)||s.flipH||s.flipV||s.type==='image'&&s.crop||s.type==='line'&&s.connector||s.stroke?.dash&&s.stroke.dash!=='solid'||s.stroke?.head||s.stroke?.tail))throw new OfficeError('unsupported','Clipboard graphics cannot be preserved in legacy PPT');const count=scopeShapes().length;if(apply(pasteShapes(snap(),state.slideId!,copiedShapes.shapes,copiedShapes.resources,id,state.groupId)))choose(scopeShapes().slice(count).map(s=>s.id));}));

  function present():void {
    if(!flush())return;const snapshot=snap(),m=presentation(snapshot);let page=Math.max(0,m.slides.findIndex(s=>s.id===state.slideId));const dialog=element('dialog','','slides-presentation'),screen=element('div','','slides-presentation-screen'),bar=element('div','','slides-presentation-bar'),caption=element('span','');presentationDialog=dialog;dialog.tabIndex=-1;dialog.dataset.revision=String(snapshot.revision);root.append(dialog);dialog.append(screen,bar);
    const previous=element('button',t('slidePrevious')),next=element('button',t('slideNext')),exit=element('button',t('slideExitPresentation'));for(const b of [previous,next,exit])b.type='button';bar.append(previous,caption,next,exit);
    const resizePresentation=()=>{const node=screen.firstElementChild as HTMLElement|null;if(!node)return;const scale=Math.min(screen.clientWidth/(m.size.width/EMU_PER_CSS_PIXEL),screen.clientHeight/(m.size.height/EMU_PER_CSS_PIXEL));node.style.transform=`scale(${Math.max(.01,scale)})`;};
    const paint=()=>{const node=createSlideSceneElement(buildSlideScene(snapshot,m.slides[page]!.id));node.style.transformOrigin='center center';screen.replaceChildren(node);caption.textContent=t('slidePageNumber',{current:host.number(page+1),total:host.number(m.slides.length)});previous.disabled=page===0;next.disabled=page===m.slides.length-1;resizePresentation();};
    const close=()=>{closePresentation=undefined;if(document.fullscreenElement===dialog)void document.exitFullscreen().catch(()=>{});dialog.close();dialog.remove();presentationDialog=undefined;window.removeEventListener('resize',resizePresentation);stage.focus();};closePresentation=close;previous.onclick=()=>{page=Math.max(0,page-1);paint();};next.onclick=()=>{page=Math.min(m.slides.length-1,page+1);paint();};exit.onclick=close;dialog.oncancel=event=>{event.preventDefault();close();};dialog.onkeydown=event=>{if(['ArrowRight','ArrowDown','PageDown',' '].includes(event.key)){event.preventDefault();page=Math.min(m.slides.length-1,page+1);paint();}else if(['ArrowLeft','ArrowUp','PageUp'].includes(event.key)){event.preventDefault();page=Math.max(0,page-1);paint();}else if(event.key==='Home'){event.preventDefault();page=0;paint();}else if(event.key==='End'){event.preventDefault();page=m.slides.length-1;paint();}else if(event.key==='Escape'){event.preventDefault();close();}};dialog.showModal();dialog.focus();paint();window.addEventListener('resize',resizePresentation);if(dialog.requestFullscreen)void dialog.requestFullscreen().then(()=>{resizePresentation();if(dialog.isConnected)dialog.focus();}).catch(()=>{});
  }
  state.dispose=()=>{if(disposed)return;cancelGesture();disposed=true;if(textTimer!==undefined)clearTimeout(textTimer);observer?.disconnect();thumbObserver?.disconnect();thumbnails.clear();thumbnailSnapshot=undefined;thumbnailGlobals=undefined;closePresentation?.();presentationDialog?.close();presentationDialog?.remove();};
  const pick=(...ids:string[])=>ribbonNodes(root,...ids);
  installRibbon(root,state,t,[
    {id:'home',label:'ribbonHome',groups:[
      {label:'ribbonSlide',rows:[pick('slide-add','slide-duplicate','slide-delete'),pick('slide-earlier','slide-later')]},
      {label:'ribbonFont',rows:[pick('slide-font-family','slide-font-size'),pick('slide-text-bold','slide-text-italic','slide-text-underline','slide-font-color')]},
      {label:'ribbonParagraph',rows:[pick('slide-text-start','slide-text-center','slide-text-end')]},
      {label:'ribbonInsert',rows:[pick('slide-insert-text','slide-insert-image')]},
      {label:'ribbonPresentation',rows:[pick('slide-present')]}
    ]},
    {id:'insert',label:'ribbonInsert',groups:[
      {label:'slideTable',rows:[pick('slide-insert-table','slide-edit-table')]},
      {label:'ribbonShapes',rows:[pick('slide-insert-rect','slide-insert-ellipse','slide-insert-line','slide-insert-roundRect'),pick('slide-insert-triangle','slide-insert-diamond','slide-insert-parallelogram','slide-insert-connector')]},
      {label:'ribbonObject',rows:[pick('slide-object-duplicate','slide-object-delete')]}
    ]},
    {id:'arrange',label:'ribbonArrange',groups:[
      {label:'ribbonLayer',rows:[pick('slide-layer-front','slide-layer-back'),pick('slide-layer-forward','slide-layer-backward')]},
      {label:'ribbonAlignment',rows:[pick('slide-align-left','slide-align-center','slide-align-right'),pick('slide-align-top','slide-align-middle','slide-align-bottom'),pick('slide-distribute-h','slide-distribute-v')]},
      {label:'ribbonGrouping',rows:[pick('slide-group','slide-ungroup','slide-edit-group','slide-exit-group'),pick('slide-flipH','slide-flipV','slide-connect-objects','slide-detach-connector')]}
    ]},
    {id:'design',label:'ribbonDesign',groups:[
      {label:'ribbonSlide',rows:[[background.closest('label')!,themeAccent.closest('label')!],[slideWidth.closest('label')!,slideHeight.closest('label')!,...pick('slide-size-apply')]]}
    ]},
    {id:'view',label:'ribbonView',groups:[
      {label:'ribbonWindow',rows:[[propertiesButton]]}
    ]}
  ],()=>!host.composing()&&flush());
  // Formatting fields remain controlled by the same selection sync.
  const zoomBar=element('div','','slides-bottom-zoom');
  for(const id of ['slide-zoom-out','slide-zoom-fit','slide-zoom-in']){const b=root.querySelector<HTMLButtonElement>('#'+id)!;decorateOfficeButton(b);zoomBar.append(b);}
  status.append(zoomBar);
  for(const c of [toolbar,textTools,objectActions,arrangeGroup])if(!c.childElementCount)c.remove();
  // The Design ribbon owns the page-level fields, so the right panel shows
  // an explicit useful status rather than an empty fieldset.
  const slideInfo=element('p',t('ribbonSlideSettingsHint'),'muted');slideGroup.append(slideInfo);
  for(const b of inspector.querySelectorAll<HTMLButtonElement>('button[id]'))decorateOfficeButton(b);
  showInspector();
  refresh();observer=new ResizeObserver(()=>{if(!disposed)resize();});observer.observe(stage);return root;
}
