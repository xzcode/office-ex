// SPDX-License-Identifier: MIT
import { EMU_PER_CSS_PIXEL, EMU_PER_POINT, ROTATION_UNITS_PER_DEGREE, OfficeError } from '@office-ex/contracts';
import { walkSlideShapes, type DocumentSnapshot, type DocumentCommand, type ImageResource, type Shape, type Geometry, type TextStyle, type SlideParagraphFormat, type Color, type DeepReadonly, type SlidePreset, type SlideCrop, type SlideStroke, type SlideDash } from '@office-ex/core';
import { addSlide, duplicateSlide, deleteSlide, moveSlide, locateSlide, locateShape, presentation, insertShapes, removeShapes, duplicateShapes, setShapeGeometries, setShapePaint, reorderShapes, alignShapes, insertSlideImage, aspectFitImage, shapeText, replaceShapeText, slideTextEditDiff, formatShapeText, formatSlideParagraphs, setShapeTextFrame, setSlideBackground, setSlideSize, setSlideTheme, DEFAULT_SLIDE_THEME, pasteShapes, groupShapes, ungroupShape, canUngroupShape, setSlideImageCrop, flipSlideShapes, type ShapeSnapshot, type ShapeAlignment, type LayerMove } from '@office-ex/slides';
import { buildSlideScene, hitTest, shapeTransform, inverse, transform, type SlideScene, type Point } from '@office-ex/render';
import type { MessageKey } from '@office-ex/i18n';
import { createSlideSceneElement, slideCssColor, overflowingSlideText, measureSlideTextFit, updateSlideShapeGeometry } from './slides-scene-dom.js';

export interface SlidesViewState {slideId?:string;selectedIds:string[];zoom:number;lockAspect?:boolean;flush?:()=>boolean;dispose?:()=>void}
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
  state.dispose?.();state.selectedIds??=[];state.zoom??=0;state.lockAspect??=true;
  const t=host.t,root=element('section','','slides-workspace');root.id='slides-workspace';root.setAttribute('aria-label',t('slides'));
  const toolbar=element('div','','slides-toolbar'),body=element('div','','slides-workbench'),rail=element('nav','','slides-thumbnails'),stage=element('div','','slides-stage'),inspector=element('aside','','slides-inspector'),surface=element('div','','slides-surface'),scaled=element('div','','slides-scaled-scene'),overlay=element('div','','slides-selection-overlay'),status=element('div','','slides-status');
  toolbar.setAttribute('role','toolbar');toolbar.setAttribute('aria-label',t('slides'));rail.setAttribute('aria-label',t('slides'));stage.tabIndex=0;stage.setAttribute('role','application');stage.setAttribute('aria-label',t('canvas'));root.append(toolbar,body,status);body.append(rail,stage,inspector);stage.append(surface);surface.append(scaled);scaled.append(overlay);
  const help=element('p',t('slideHelp'),'muted slides-help'),overflow=element('p','','slides-overflow'),position=element('span','');status.append(position,overflow,help);overflow.hidden=true;
  let scene:SlideScene,sceneElement:HTMLDivElement,gesture:Gesture|undefined,disposed=false,zoom=1,textTimer:ReturnType<typeof setTimeout>|undefined,textDirty=false,editingId:string|undefined,observer:ResizeObserver|undefined,thumbObserver:IntersectionObserver|undefined,presentationDialog:HTMLDialogElement|undefined,closePresentation:(()=>void)|undefined;
  const editControls:(HTMLButtonElement|HTMLInputElement|HTMLSelectElement|HTMLTextAreaElement)[]=[];const selectionControls:(HTMLButtonElement|HTMLInputElement|HTMLSelectElement)[]=[];
  const snap=()=>host.snapshot(),doc=()=>presentation(snap()),selected=()=>state.selectedIds.map(shapeId=>locateShape(snap(),state.slideId!,shapeId).shape);
  function safe(fn:()=>void):void{try{fn();}catch(error){host.error(error);}}
  function button(key:MessageKey,fn:()=>void,target:HTMLElement=toolbar,identifier?:string,needsSelection=false,mutates=true):HTMLButtonElement {
    const b=element('button',t(key));b.type='button';if(identifier)b.id=identifier;b.onpointerdown=event=>{if(event.pointerType==='mouse')event.preventDefault();};b.onclick=()=>safe(()=>{if(flush())fn();});target.append(b);if(mutates)editControls.push(b);if(needsSelection)selectionControls.push(b);return b;
  }
  function field(key:MessageKey,control:HTMLInputElement|HTMLSelectElement|HTMLTextAreaElement,parent:HTMLElement=inspector):void {const label=element('label',t(key),'slides-field');label.append(control);control.setAttribute('aria-label',t(key));parent.append(label);editControls.push(control);}
  function input(key:MessageKey,type:string,value:string,fn:(value:string)=>void,needsSelection=true):HTMLInputElement {
    const control=element('input');control.type=type;control.value=value;field(key,control);control.onchange=()=>safe(()=>{if(flush())fn(control.value);});if(needsSelection)selectionControls.push(control);return control;
  }
  function apply(command:DocumentCommand,repaint=true):boolean {if(host.readOnly||disposed)return false;const ok=host.execute(command);if(ok&&repaint)refresh();return ok;}
  function cancelGesture():void {if(!gesture)return;const previous=gesture;gesture=undefined;if(previous.frame!==undefined)cancelAnimationFrame(previous.frame);if(stage.hasPointerCapture(previous.pointerId))stage.releasePointerCapture(previous.pointerId);paintGeometry();}
  function flush():boolean {if(disposed)return true;if(host.composing())return false;if(gesture)cancelGesture();if(textTimer!==undefined){clearTimeout(textTimer);textTimer=undefined;}return commitText();}
  state.flush=flush;
  function selectedText():Extract<ShapeSnapshot,{type:'text'}>|undefined {if(state.selectedIds.length!==1)return undefined;const s=selected()[0];return s?.type==='text'?s:undefined;}
  function commitText():boolean {
    if(!textDirty||!editingId)return true;if(host.readOnly||host.composing())return false;
    try{const shape=locateShape(snap(),state.slideId!,editingId).shape;if(shape.type!=='text')return false;const before=shapeText(shape);if(before===textEditor.value){textDirty=false;return true;}const difference=slideTextEditDiff(before,textEditor.value);const ok=apply(replaceShapeText(snap(),state.slideId!,editingId,difference.start,difference.end,difference.text,id),false);if(ok){textDirty=false;refreshScene();refreshThumbnails();}return ok;}catch(error){host.error(error);return false;}
  }
  function changeSlide(slideId:string):void {if(!flush())return;state.slideId=slideId;state.selectedIds=[];editingId=undefined;refresh();stage.focus({preventScroll:true});}
  function choose(ids:string[]):void {if(!flush())return;state.selectedIds=ids;syncInspector();paintGeometry();}
  function addShape(type:'text'|SlidePreset,connector=false):void {
    const m=doc(),slide=locateSlide(snap(),state.slideId!).slide,shapeId=id(),geometry={x:Math.round(m.size.width*.2),y:Math.round(m.size.height*.2),width:Math.round(m.size.width*.4),height:Math.round(m.size.height*(type==='line'?.15:type==='text'?.2:.3)),rotation:0};let shape:Shape;
    if(type==='text')shape={id:shapeId,type,runs:[{id:id(),type:'text',text:''}],textStyle:{fontSize:304800},geometry};
    else shape={id:shapeId,type,geometry,...(connector?{connector:true as const}:{}),...(type==='line'?{}:{fill:{kind:'theme',name:'accent1'} as Color}),stroke:{color:{kind:'theme',name:type==='line'?'accent1':'dk1'},width:12700}};
    if(apply(insertShapes(snap(),state.slideId!,slide.shapes.length,[shape]))){choose([shapeId]);if(type==='text')textEditor.focus();}
  }
  button('slideAdd',()=>{const current=locateSlide(snap(),state.slideId!),newId=id();if(apply(addSlide(snap(),current.index+1,newId,current.slide.background as Color))){state.slideId=newId;state.selectedIds=[];refresh();}},toolbar,'slide-add');
  button('slideDuplicate',()=>{const current=locateSlide(snap(),state.slideId!);if(apply(duplicateSlide(snap(),state.slideId!,id))){state.slideId=doc().slides[current.index+1]!.id;state.selectedIds=[];refresh();}},toolbar,'slide-duplicate');
  const deletePage=button('slideDelete',()=>{const current=locateSlide(snap(),state.slideId!);if(apply(deleteSlide(snap(),state.slideId!))){state.slideId=doc().slides[Math.min(current.index,doc().slides.length-1)]!.id;state.selectedIds=[];refresh();}},toolbar,'slide-delete');
  const earlier=button('slideEarlier',()=>{const current=locateSlide(snap(),state.slideId!);apply(moveSlide(snap(),state.slideId!,current.index-1));},toolbar,'slide-earlier');
  const later=button('slideLater',()=>{const current=locateSlide(snap(),state.slideId!);apply(moveSlide(snap(),state.slideId!,current.index+1));},toolbar,'slide-later');
  for(const [type,key] of [['text','slideText'],['rect','slideRectangle'],['ellipse','slideEllipse'],['line','slideLine']] as const)button(key,()=>addShape(type),toolbar,'slide-insert-'+type);
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
  button('slideDuplicateObject',()=>{const count=locateSlide(snap(),state.slideId!).slide.shapes.length;if(apply(duplicateShapes(snap(),state.slideId!,state.selectedIds,id)))choose(locateSlide(snap(),state.slideId!).slide.shapes.slice(count).map(s=>s.id));},objectActions,'slide-object-duplicate',true);
  button('slideDeleteObject',()=>{if(apply(removeShapes(snap(),state.slideId!,state.selectedIds)))choose([]);},objectActions,'slide-object-delete',true);
  const groupButton=button('slideGroupObjects',()=>{const groupId=id();if(apply(groupShapes(snap(),state.slideId!,state.selectedIds,()=>groupId)))choose([groupId]);},objectActions,'slide-group',true);
  const ungroupButton=button('slideUngroupObjects',()=>{const shape=selected()[0];if(shape?.type==='group'){const ids=shape.children.map(child=>child.id);if(apply(ungroupShape(snap(),state.slideId!,shape.id)))choose(ids);}},objectActions,'slide-ungroup',true);
  const flipButtons=(['flipH','flipV'] as const).map((axis,index)=>button(index?'slideFlipV':'slideFlipH',()=>apply(flipSlideShapes(snap(),state.slideId!,state.selectedIds,axis)),objectActions,'slide-'+axis,true));
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
  const font=element('input');font.type='text';font.maxLength=512;field('writerFontFamily',font,textGroup);font.onchange=()=>safe(()=>{if(flush()&&font.value.trim())format({fontFamily:font.value.trim()});});
  const size=element('input');size.type='number';size.min='1';size.max='1638';size.step='.5';field('writerFontSize',size,textGroup);size.onchange=()=>safe(()=>{if(flush())format({fontSize:emu(size.value)});});
  const textColor=element('input');textColor.type='color';field('writerTextColor',textColor,textGroup);textColor.onchange=()=>safe(()=>{if(flush())format({color:rgb(textColor.value)});});
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
      const candidate=buildSlideScene(snap(),slideId).shapes.find(s=>s.id===shapeId);if(candidate?.type!=='text')return;
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

  function syncInspector():void {
    const m=doc(),current=locateSlide(snap(),state.slideId!);state.selectedIds=state.selectedIds.filter(shapeId=>current.slide.shapes.some(s=>s.id===shapeId));const shapes=selected(),single=shapes.length===1?shapes[0]:undefined,text=single?.type==='text'?single:undefined;
    for(const control of editControls)control.disabled=!!host.readOnly;for(const control of selectionControls)control.disabled=!!host.readOnly||!shapes.length;
    deletePage.disabled=!!host.readOnly||m.slides.length<2;earlier.disabled=!!host.readOnly||current.index===0;later.disabled=!!host.readOnly||current.index===m.slides.length-1;
    for(const [key,control] of Object.entries(geometryFields)){control.disabled=!!host.readOnly||!single;control.value=single?String(Math.round(single.geometry[key as keyof Geometry]/(key==='rotation'?ROTATION_UNITS_PER_DEGREE:EMU_PER_POINT)*100)/100):'';}
    aspect.checked=single?.type==='image'||!!state.lockAspect;aspect.disabled=single?.type==='image'||!!host.readOnly;
    fill.value=colorHex(single?.fill);stroke.value=colorHex(single?.stroke?.color,'#182030');strokeWidth.value=String((single?.stroke?.width??12700)/EMU_PER_POINT);
    const graphicsDisabled=!!host.readOnly||!!host.legacyPresentation;
    for(const control of newGraphics)control.disabled=graphicsDisabled;
    const selectedIndices=current.slide.shapes.flatMap((shape,i)=>state.selectedIds.includes(shape.id)?[i]:[]);
    groupButton.disabled=graphicsDisabled||selectedIndices.length<2||selectedIndices.at(-1)!-selectedIndices[0]!+1!==selectedIndices.length;
    ungroupButton.disabled=graphicsDisabled||!single||!canUngroupShape(single);
    for(const control of flipButtons)control.disabled=graphicsDisabled||!shapes.length||shapes.some(shape=>shape.type==='text');
    fill.disabled=!!host.readOnly||!single||single.type==='group';
    stroke.disabled=strokeWidth.disabled=!!host.readOnly||!single||single.type==='group';
    dash.disabled=graphicsDisabled||!single||['text','image','group'].includes(single.type);
    arrowHead.disabled=arrowTail.disabled=graphicsDisabled||single?.type!=='line';
    dash.value=single?.stroke?.dash??'solid';arrowHead.value=single?.stroke?.head?.type??'none';arrowTail.value=single?.stroke?.tail?.type??'none';
    cropGroup.hidden=single?.type!=='image';cropGroup.disabled=graphicsDisabled||single?.type!=='image';cropApply.disabled=cropReset.disabled=cropGroup.disabled;
    if(single?.type==='image')for(const key of Object.keys(cropFields) as (keyof SlideCrop)[])cropFields[key].value=String((single.crop?.[key]??0)/1000);
    textGroup.disabled=!text||!!host.readOnly;
    if(text){if(editingId!==text.id||!textDirty){const focused=document.activeElement===textEditor,start=textEditor.selectionStart,end=textEditor.selectionEnd;textEditor.value=shapeText(text);if(focused)textEditor.setSelectionRange(Math.min(start,textEditor.value.length),Math.min(end,textEditor.value.length));}editingId=text.id;const style={...m.defaultTextStyle,...text.textStyle,...text.runs[0]?.style};font.value=style.fontFamily??'sans-serif';size.value=String((style.fontSize??304800)/EMU_PER_POINT);textColor.value=colorHex(style.color,'#182030');language.value=style.language??'und';direction.value=text.textFrame?.direction==='rtl'?'rtl':'ltr';wrap.checked=text.textFrame?.wrap??true;vertical.value=text.textFrame?.verticalAlignment??'top';textEditor.dir=text.textFrame?.direction??style.paragraphDirection??'auto';}
    else {editingId=undefined;textEditor.value='';textDirty=false;}
    syncListControls();
    background.value=colorHex(current.slide.background,'#ffffff');themeAccent.value=colorHex(m.theme?.colors.accent1??DEFAULT_SLIDE_THEME.colors.accent1);slideWidth.value=String(Math.round(m.size.width/EMU_PER_POINT*100)/100);slideHeight.value=String(Math.round(m.size.height/EMU_PER_POINT*100)/100);
    position.textContent=t('slidePageNumber',{current:host.number(current.index+1),total:host.number(m.slides.length)});for(const thumb of rail.querySelectorAll<HTMLButtonElement>('[data-slide-id]')){thumb.setAttribute('aria-current',String(thumb.dataset.slideId===state.slideId));thumb.classList.toggle('is-current',thumb.dataset.slideId===state.slideId);}
  }
  function checkOverflow():void {if(disposed||!sceneElement?.isConnected||scene.revision!==snap().revision)return;const overflowIds=overflowingSlideText(sceneElement);overflow.hidden=overflowIds.length===0;overflow.textContent=overflowIds.length?t('slideTextOverflow'):'';for(const node of sceneElement.querySelectorAll<HTMLElement>('[data-shape-id]'))node.classList.toggle('slides-text-overflow',overflowIds.includes(node.dataset.shapeId!));}
  function refreshScene():void {
    const snapshot=snap(),m=presentation(snapshot);if(!m.slides.length)throw new OfficeError('invalid','Presentation has no slides');if(!m.slides.some(s=>s.id===state.slideId))state.slideId=m.slides[0]!.id;
    scene=buildSlideScene(snapshot,state.slideId!);sceneElement=createSlideSceneElement(scene);sceneElement.id='slide-canvas';scaled.replaceChildren(sceneElement,overlay);resize();paintGeometry();requestAnimationFrame(checkOverflow);void document.fonts.ready.then(()=>{if(!disposed)checkOverflow();});
  }
  function refreshThumbnails():void {
    thumbObserver?.disconnect();rail.replaceChildren();const snapshot=snap(),m=presentation(snapshot),seen=new Set<string>();
    const renderThumb=(thumb:HTMLElement)=>{const slideId=thumb.dataset.slideId!;if(seen.has(slideId)||disposed)return;seen.add(slideId);const thumbnailScene=buildSlideScene(snapshot,slideId),thumbScene=createSlideSceneElement(thumbnailScene),ratio=128/(thumbnailScene.width/EMU_PER_CSS_PIXEL);thumbScene.style.transform=`scale(${ratio})`;thumbScene.style.transformOrigin='top left';thumb.firstElementChild!.append(thumbScene);};
    if(typeof IntersectionObserver!=='undefined')thumbObserver=new IntersectionObserver(entries=>{for(const entry of entries)if(entry.isIntersecting){renderThumb(entry.target as HTMLElement);thumbObserver?.unobserve(entry.target);}},{root:rail,rootMargin:'200px'});
    for(const [index,slide] of m.slides.entries()){const thumb=element('button','','slides-thumbnail');thumb.type='button';thumb.dataset.slideId=slide.id;thumb.setAttribute('aria-label',t('slidePageNumber',{current:host.number(index+1),total:host.number(m.slides.length)}));thumb.setAttribute('aria-current',String(slide.id===state.slideId));thumb.classList.toggle('is-current',slide.id===state.slideId);const image=element('div','','slides-thumbnail-image');image.style.width='128px';image.style.height=128*m.size.height/m.size.width+'px';image.setAttribute('aria-hidden','true');thumb.append(image,element('span',host.number(index+1)));thumb.onclick=()=>safe(()=>changeSlide(slide.id));thumb.draggable=!host.readOnly;thumb.ondragstart=event=>{if(!flush()){event.preventDefault();return;}event.dataTransfer?.setData('application/x-office-ex-slide',slide.id);if(event.dataTransfer)event.dataTransfer.effectAllowed='move';};thumb.ondragover=event=>{if(!host.readOnly&&event.dataTransfer?.types.includes('application/x-office-ex-slide'))event.preventDefault();};thumb.ondrop=event=>{event.preventDefault();const source=event.dataTransfer?.getData('application/x-office-ex-slide');if(source&&doc().slides.some(s=>s.id===source))safe(()=>{if(flush())apply(moveSlide(snap(),source,index));});};rail.append(thumb);if(thumbObserver)thumbObserver.observe(thumb);else renderThumb(thumb);}
  }
  function refresh():void {if(disposed)return;refreshScene();refreshThumbnails();syncInspector();}
  function resize():void {
    if(!scene)return;const width=scene.width/EMU_PER_CSS_PIXEL,height=scene.height/EMU_PER_CSS_PIXEL,fit=Math.max(.05,Math.min(1,(stage.clientWidth-56)/width,(Math.max(360,stage.clientHeight)-56)/height));zoom=state.zoom>0?state.zoom:fit;surface.style.width=width*zoom+'px';surface.style.height=height*zoom+'px';scaled.style.width=width+'px';scaled.style.height=height+'px';scaled.style.transform=`scale(${zoom})`;stage.dataset.zoom=String(zoom);paintGeometry();
  }
  function paintGeometry():void {
    if(!sceneElement)return;const values=gesture?.preview??{};for(const node of sceneElement.querySelectorAll<HTMLElement>('[data-shape-id]')){const shape=scene.shapes.find(s=>s.id===node.dataset.shapeId);if(!shape)continue;const g=values[shape.id]??shape.geometry;updateSlideShapeGeometry(node,shape,g);}
    overlay.replaceChildren();for(const shapeId of state.selectedIds){const shape=scene.shapes.find(s=>s.id===shapeId);if(!shape)continue;const g=values[shapeId]??shape.geometry,box=element('div','','slides-selection-box');box.dataset.selectedShape=shapeId;Object.assign(box.style,{left:g.x/EMU_PER_CSS_PIXEL+'px',top:g.y/EMU_PER_CSS_PIXEL+'px',width:g.width/EMU_PER_CSS_PIXEL+'px',height:g.height/EMU_PER_CSS_PIXEL+'px',transform:`rotate(${g.rotation/ROTATION_UNITS_PER_DEGREE}deg)`,borderWidth:1.5/zoom+'px'});overlay.append(box);
      if(state.selectedIds.length===1&&!host.readOnly)for(const corner of ['nw','ne','se','sw','rotate']){const handle=element('span','','slides-handle slides-handle-'+corner);handle.dataset.handle=corner;handle.dataset.shapeId=shapeId;handle.style.width=9/zoom+'px';handle.style.height=9/zoom+'px';handle.style.borderWidth=1/zoom+'px';if(corner==='rotate')handle.style.top=-24/zoom+'px';handle.setAttribute('aria-hidden','true');box.append(handle);}}
  }
  function point(event:PointerEvent|MouseEvent):Point {const rect=sceneElement.getBoundingClientRect();return {x:(event.clientX-rect.left)/zoom*EMU_PER_CSS_PIXEL,y:(event.clientY-rect.top)/zoom*EMU_PER_CSS_PIXEL};}
  stage.onpointerdown=event=>safe(()=>{
    if(event.button!==0||host.readOnly||host.composing())return;const target=event.target as HTMLElement,handle=target.closest<HTMLElement>('[data-handle]'),p=point(event);if(!flush())return;
    if(handle){const shape=locateShape(snap(),state.slideId!,handle.dataset.shapeId!).shape,g={...shape.geometry};gesture={pointerId:event.pointerId,revision:snap().revision,kind:handle.dataset.handle==='rotate'?'rotate':'resize',start:p,shapeId:shape.id,corner:handle.dataset.handle,geometries:{[shape.id]:g},preview:{[shape.id]:g},angle:Math.atan2(p.y-g.y-g.height/2,p.x-g.x-g.width/2)};}
    else {const shapeId=hitTest(locateSlide(snap(),state.slideId!).slide.shapes,p,6/zoom*EMU_PER_CSS_PIXEL);if(!shapeId){choose([]);stage.focus();return;}if(event.shiftKey){choose(state.selectedIds.includes(shapeId)?state.selectedIds.filter(value=>value!==shapeId):[...state.selectedIds,shapeId]);return;}if(!state.selectedIds.includes(shapeId))choose([shapeId]);const geometries=Object.fromEntries(selected().map(s=>[s.id,{...s.geometry}]));gesture={pointerId:event.pointerId,revision:snap().revision,kind:'move',start:p,geometries,preview:geometries};}
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
  stage.onpointerup=event=>safe(()=>{if(!gesture||gesture.pointerId!==event.pointerId)return;const current=gesture;gesture=undefined;if(current.frame!==undefined)cancelAnimationFrame(current.frame);if(stage.hasPointerCapture(event.pointerId))stage.releasePointerCapture(event.pointerId);if(current.revision!==snap().revision){refresh();return;}apply(setShapeGeometries(snap(),state.slideId!,current.preview));});
  stage.onpointercancel=()=>cancelGesture();stage.onlostpointercapture=()=>{if(gesture)cancelGesture();};stage.ondblclick=event=>safe(()=>{const shapeId=hitTest(locateSlide(snap(),state.slideId!).slide.shapes,point(event),4/zoom*EMU_PER_CSS_PIXEL);if(shapeId){choose([shapeId]);if(selectedText()){textEditor.focus();textEditor.setSelectionRange(0,textEditor.value.length);}}});
  root.addEventListener('keydown',event=>safe(()=>{
    if(event.key==='Escape'&&gesture){event.preventDefault();event.stopPropagation();cancelGesture();return;}if(host.composing()||isTyping(event.target)||host.readOnly)return;
    if(event.key==='Escape'){choose([]);event.preventDefault();return;}
    if(event.key==='F5'){event.preventDefault();if(flush())present();return;}
    if((event.ctrlKey||event.metaKey)&&event.key.toLowerCase()==='a'){choose(locateSlide(snap(),state.slideId!).slide.shapes.map(s=>s.id));event.preventDefault();return;}
    if((event.ctrlKey||event.metaKey)&&event.key.toLowerCase()==='d'){event.preventDefault();if(flush()){const count=locateSlide(snap(),state.slideId!).slide.shapes.length;if(apply(duplicateShapes(snap(),state.slideId!,state.selectedIds,id)))choose(locateSlide(snap(),state.slideId!).slide.shapes.slice(count).map(s=>s.id));}return;}
    if(!state.selectedIds.length)return;if(['Delete','Backspace'].includes(event.key)){event.preventDefault();if(flush()&&apply(removeShapes(snap(),state.slideId!,state.selectedIds)))choose([]);return;}
    if(event.key==='Enter'&&selectedText()){event.preventDefault();textEditor.focus();return;}
    if(['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(event.key)){event.preventDefault();if(!flush())return;const step=(event.shiftKey?10:1)*EMU_PER_CSS_PIXEL,dx=event.key==='ArrowLeft'?-step:event.key==='ArrowRight'?step:0,dy=event.key==='ArrowUp'?-step:event.key==='ArrowDown'?step:0;apply(setShapeGeometries(snap(),state.slideId!,Object.fromEntries(selected().map(s=>[s.id,{...s.geometry,x:s.geometry.x+dx,y:s.geometry.y+dy}]))));}
  }));
  root.addEventListener('copy',event=>safe(()=>{if(isTyping(event.target)||!state.selectedIds.length||!flush())return;const shapes=selected(),resources:Record<string,DeepReadonly<ImageResource>>={};for(const shape of walkSlideShapes(shapes))if(shape.type==='image'){const r=doc().resources[shape.resourceId];if(r?.dataBase64)resources[shape.resourceId]=r as DeepReadonly<ImageResource>;}copiedShapes={nonce:id(),shapes,resources};event.clipboardData?.setData('application/x-office-ex-shapes',copiedShapes.nonce);event.clipboardData?.setData('text/plain',shapes.filter(s=>s.type==='text').map(shapeText).join('\n'));event.preventDefault();}));
  root.addEventListener('cut',event=>safe(()=>{if(isTyping(event.target)||host.readOnly||!state.selectedIds.length||!flush())return;const copy=new ClipboardEvent('copy',{bubbles:true,clipboardData:event.clipboardData});stage.dispatchEvent(copy);if(copiedShapes&&event.clipboardData?.getData('application/x-office-ex-shapes')===copiedShapes.nonce){event.preventDefault();if(apply(removeShapes(snap(),state.slideId!,state.selectedIds)))choose([]);}}));
  root.addEventListener('paste',event=>safe(()=>{if(isTyping(event.target)||!copiedShapes||event.clipboardData?.getData('application/x-office-ex-shapes')!==copiedShapes.nonce||host.readOnly||!flush())return;event.preventDefault();if(host.legacyPresentation&&copiedShapes.shapes.some(s=>!['rect','ellipse','line','image','text'].includes(s.type)||s.flipH||s.flipV||s.type==='image'&&s.crop||s.type==='line'&&s.connector||s.stroke?.dash&&s.stroke.dash!=='solid'||s.stroke?.head||s.stroke?.tail))throw new OfficeError('unsupported','Clipboard graphics cannot be preserved in legacy PPT');const count=locateSlide(snap(),state.slideId!).slide.shapes.length;if(apply(pasteShapes(snap(),state.slideId!,copiedShapes.shapes,copiedShapes.resources,id)))choose(locateSlide(snap(),state.slideId!).slide.shapes.slice(count).map(s=>s.id));}));

  function present():void {
    if(!flush())return;const snapshot=snap(),m=presentation(snapshot);let page=Math.max(0,m.slides.findIndex(s=>s.id===state.slideId));const dialog=element('dialog','','slides-presentation'),screen=element('div','','slides-presentation-screen'),bar=element('div','','slides-presentation-bar'),caption=element('span','');presentationDialog=dialog;dialog.tabIndex=-1;dialog.dataset.revision=String(snapshot.revision);root.append(dialog);dialog.append(screen,bar);
    const previous=element('button',t('slidePrevious')),next=element('button',t('slideNext')),exit=element('button',t('slideExitPresentation'));for(const b of [previous,next,exit])b.type='button';bar.append(previous,caption,next,exit);
    const resizePresentation=()=>{const node=screen.firstElementChild as HTMLElement|null;if(!node)return;const scale=Math.min(screen.clientWidth/(m.size.width/EMU_PER_CSS_PIXEL),screen.clientHeight/(m.size.height/EMU_PER_CSS_PIXEL));node.style.transform=`scale(${Math.max(.01,scale)})`;};
    const paint=()=>{const node=createSlideSceneElement(buildSlideScene(snapshot,m.slides[page]!.id));node.style.transformOrigin='center center';screen.replaceChildren(node);caption.textContent=t('slidePageNumber',{current:host.number(page+1),total:host.number(m.slides.length)});previous.disabled=page===0;next.disabled=page===m.slides.length-1;resizePresentation();};
    const close=()=>{closePresentation=undefined;if(document.fullscreenElement===dialog)void document.exitFullscreen().catch(()=>{});dialog.close();dialog.remove();presentationDialog=undefined;window.removeEventListener('resize',resizePresentation);stage.focus();};closePresentation=close;previous.onclick=()=>{page=Math.max(0,page-1);paint();};next.onclick=()=>{page=Math.min(m.slides.length-1,page+1);paint();};exit.onclick=close;dialog.oncancel=event=>{event.preventDefault();close();};dialog.onkeydown=event=>{if(['ArrowRight','ArrowDown','PageDown',' '].includes(event.key)){event.preventDefault();page=Math.min(m.slides.length-1,page+1);paint();}else if(['ArrowLeft','ArrowUp','PageUp'].includes(event.key)){event.preventDefault();page=Math.max(0,page-1);paint();}else if(event.key==='Home'){event.preventDefault();page=0;paint();}else if(event.key==='End'){event.preventDefault();page=m.slides.length-1;paint();}else if(event.key==='Escape'){event.preventDefault();close();}};dialog.showModal();dialog.focus();paint();window.addEventListener('resize',resizePresentation);if(dialog.requestFullscreen)void dialog.requestFullscreen().then(()=>{resizePresentation();if(dialog.isConnected)dialog.focus();}).catch(()=>{});
  }
  state.dispose=()=>{if(disposed)return;cancelGesture();disposed=true;if(textTimer!==undefined)clearTimeout(textTimer);observer?.disconnect();thumbObserver?.disconnect();closePresentation?.();presentationDialog?.close();presentationDialog?.remove();};
  refresh();observer=new ResizeObserver(()=>{if(!disposed)resize();});observer.observe(stage);return root;
}
