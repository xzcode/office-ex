// SPDX-License-Identifier: MIT
/** Presentation commands use stable IDs and one atomic patch per user action. */
import { assert, assertId, assertUnicode, graphemeBoundaries } from '@office-ex/contracts';
import { cloneData, patchCommand, DEFAULT_SLIDE_THEME_COLORS, validateColor, validateGeometry, validateTextStyle, validateImageResource, type DocumentSnapshot, type DocumentCommand, type DeepReadonly, type SlidesDocument, type Slide, type Shape, type Geometry, type Color, type TextRun, type TextStyle, type SlideTextFrame, type SlideTheme, type ImageResource, type PatchOperation } from '@office-ex/core';

export type PresentationSnapshot = DeepReadonly<SlidesDocument>;
export type SlideSnapshot = DeepReadonly<Slide>;
export type ShapeSnapshot = DeepReadonly<Shape>;
export type IdFactory = () => string;
export function presentation(snapshot: DocumentSnapshot): PresentationSnapshot {
  assert(snapshot.document.kind === 'slides', 'invalid', 'Expected presentation'); return snapshot.document;
}
export function locateSlide(snapshot: DocumentSnapshot, slideId: string): { slide: SlideSnapshot; index: number } {
  const doc = presentation(snapshot), index = doc.slides.findIndex(slide => slide.id === slideId);
  assert(index >= 0, 'invalid', 'Slide does not exist'); return { slide: doc.slides[index]!, index };
}
export function locateShape(snapshot: DocumentSnapshot, slideId: string, shapeId: string): { shape: ShapeSnapshot; slideIndex: number; index: number } {
  const { slide, index: slideIndex } = locateSlide(snapshot, slideId), index = slide.shapes.findIndex(shape => shape.id === shapeId);
  assert(index >= 0, 'invalid', 'Shape does not exist'); return { shape: slide.shapes[index]!, slideIndex, index };
}
function command(snapshot: DocumentSnapshot, operations: readonly PatchOperation[]): DocumentCommand { return patchCommand(snapshot, 'SLD-02', operations); }
function uniqueIds(snapshot:DocumentSnapshot, factory:IdFactory):IdFactory {
  const doc=presentation(snapshot),ids=new Set(doc.slides.flatMap(s=>[s.id,...s.shapes.flatMap(shape=>[shape.id,...(shape.type==='text'?shape.runs.map(r=>r.id):[])])]));
  return ()=>{const id=factory();assertId(id);assert(!ids.has(id),'invalid','New node ID is not unique');ids.add(id);return id;};
}
export function addSlide(snapshot: DocumentSnapshot, index: number, id: string, background: Color = { kind:'rgba',red:255,green:255,blue:255,alpha:255 }): DocumentCommand {
  const doc=presentation(snapshot);assertId(id);validateColor(background);assert(Number.isSafeInteger(index)&&index>=0&&index<=doc.slides.length,'invalid','Invalid slide insertion index');
  return command(snapshot,[{op:'splice',path:['slides'],index,deleteCount:0,items:[{id,background,shapes:[]}]}]);
}
export function duplicateSlide(snapshot:DocumentSnapshot,slideId:string,factory:IdFactory):DocumentCommand {
  const {slide,index}=locateSlide(snapshot,slideId),id=uniqueIds(snapshot,factory),copy=cloneData(slide) as Slide;copy.id=id();
  for(const shape of copy.shapes){shape.id=id();if(shape.type==='text')for(const run of shape.runs)run.id=id();}
  return command(snapshot,[{op:'splice',path:['slides'],index:index+1,deleteCount:0,items:[copy]}]);
}
export function deleteSlide(snapshot:DocumentSnapshot,slideId:string):DocumentCommand {
  const doc=presentation(snapshot),{index}=locateSlide(snapshot,slideId);assert(doc.slides.length>1,'invalid','A presentation must retain one slide');
  return command(snapshot,[{op:'splice',path:['slides'],index,deleteCount:1,items:[]}]);
}
export function moveSlide(snapshot:DocumentSnapshot,slideId:string,targetIndex:number):DocumentCommand {
  const doc=presentation(snapshot),{slide,index}=locateSlide(snapshot,slideId);assert(Number.isSafeInteger(targetIndex)&&targetIndex>=0&&targetIndex<doc.slides.length,'invalid','Invalid slide destination');
  return command(snapshot,index===targetIndex?[]:[{op:'splice',path:['slides'],index,deleteCount:1,items:[]},{op:'splice',path:['slides'],index:targetIndex,deleteCount:0,items:[slide]}]);
}
export function setSlideBackground(snapshot:DocumentSnapshot,slideId:string,color:Color):DocumentCommand {validateColor(color);return command(snapshot,[{op:'set',path:['slides',locateSlide(snapshot,slideId).index,'background'],value:color}]);}
export function setSlideSize(snapshot:DocumentSnapshot,width:number,height:number):DocumentCommand {
  presentation(snapshot);assert(Number.isSafeInteger(width)&&Number.isSafeInteger(height)&&width>0&&height>0,'invalid','Invalid slide dimensions');return command(snapshot,[{op:'set',path:['size'],value:{width,height}}]);
}
export function setSlideTheme(snapshot:DocumentSnapshot,theme:SlideTheme):DocumentCommand {presentation(snapshot);return command(snapshot,[{op:'set',path:['theme'],value:theme}]);}
export const DEFAULT_SLIDE_THEME: SlideTheme = {name:'Office EX',majorFont:'sans-serif',minorFont:'sans-serif',colors:{...DEFAULT_SLIDE_THEME_COLORS}};
export function insertShapes(snapshot:DocumentSnapshot,slideId:string,index:number,shapes:readonly Shape[]):DocumentCommand {
  const located=locateSlide(snapshot,slideId);assert(Number.isSafeInteger(index)&&index>=0&&index<=located.slide.shapes.length,'invalid','Invalid shape insertion index');
  return command(snapshot,[{op:'splice',path:['slides',located.index,'shapes'],index,deleteCount:0,items:shapes}]);
}
export function removeShapes(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[]):DocumentCommand {
  const {slide,index}=locateSlide(snapshot,slideId),set=new Set(shapeIds);assert(set.size===shapeIds.length&&shapeIds.every(id=>slide.shapes.some(s=>s.id===id)),'invalid','Invalid shape selection');
  return command(snapshot,slide.shapes.flatMap((s,i)=>set.has(s.id)?[{op:'splice' as const,path:['slides',index,'shapes'],index:i,deleteCount:1,items:[]}]:[]).reverse());
}
export function duplicateShapes(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[],factory:IdFactory,offset={x:190500,y:190500}):DocumentCommand {
  const {slide}=locateSlide(snapshot,slideId),set=new Set(shapeIds),id=uniqueIds(snapshot,factory);assert(shapeIds.every(value=>slide.shapes.some(s=>s.id===value)),'invalid','Invalid shape selection');
  const copies=slide.shapes.filter(s=>set.has(s.id)).map(s=>{const copy=cloneData(s) as Shape;copy.id=id();copy.geometry.x+=offset.x;copy.geometry.y+=offset.y;if(copy.type==='text')for(const r of copy.runs)r.id=id();return copy;});
  return insertShapes(snapshot,slideId,slide.shapes.length,copies);
}
/** Transient pointer previews are outside this API; a completed gesture uses one command. */
export function setShapeGeometries(snapshot:DocumentSnapshot,slideId:string,values:Readonly<Record<string,Geometry>>):DocumentCommand {
  return command(snapshot,Object.entries(values).map(([id,geometry])=>{validateGeometry(geometry);const located=locateShape(snapshot,slideId,id);return {op:'set',path:['slides',located.slideIndex,'shapes',located.index,'geometry'],value:geometry};}));
}
export function setShapeGeometry(snapshot:DocumentSnapshot,slideId:string,shapeId:string,geometry:Geometry):DocumentCommand {return setShapeGeometries(snapshot,slideId,{[shapeId]:geometry});}
export function nudgeShapes(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[],dx:number,dy:number):DocumentCommand {
  assert(Number.isSafeInteger(dx)&&Number.isSafeInteger(dy),'invalid','Invalid shape movement');
  return setShapeGeometries(snapshot,slideId,Object.fromEntries(shapeIds.map(id=>{const g=locateShape(snapshot,slideId,id).shape.geometry;return [id,{...g,x:g.x+dx,y:g.y+dy}];})));
}
export function setShapePaint(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[],paint:{fill?:Color|null;stroke?:{color:Color;width:number}|null}):DocumentCommand {
  const operations:PatchOperation[]=[];if(paint.fill)validateColor(paint.fill);if(paint.stroke){validateColor(paint.stroke.color);assert(Number.isSafeInteger(paint.stroke.width)&&paint.stroke.width>=0,'invalid','Invalid stroke width');}
  for(const id of shapeIds){const l=locateShape(snapshot,slideId,id),base=['slides',l.slideIndex,'shapes',l.index];for(const key of ['fill','stroke'] as const)if(paint[key]!==undefined)operations.push(paint[key]===null?{op:'delete',path:[...base,key]}:{op:'set',path:[...base,key],value:paint[key]});}
  return command(snapshot,operations);
}
export type LayerMove='front'|'back'|'forward'|'backward';
export function reorderShapes(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[],move:LayerMove):DocumentCommand {
  const {slide,index}=locateSlide(snapshot,slideId),selected=new Set(shapeIds),shapes=[...slide.shapes];assert(shapeIds.every(id=>shapes.some(s=>s.id===id)),'invalid','Invalid shape selection');
  if(move==='front'||move==='back'){const chosen=shapes.filter(s=>selected.has(s.id)),rest=shapes.filter(s=>!selected.has(s.id));return command(snapshot,[{op:'set',path:['slides',index,'shapes'],value:move==='front'?[...rest,...chosen]:[...chosen,...rest]}]);}
  if(move==='forward')for(let i=shapes.length-2;i>=0;i--){if(selected.has(shapes[i]!.id)&&!selected.has(shapes[i+1]!.id))[shapes[i],shapes[i+1]]=[shapes[i+1]!,shapes[i]!];}
  else for(let i=1;i<shapes.length;i++){if(selected.has(shapes[i]!.id)&&!selected.has(shapes[i-1]!.id))[shapes[i],shapes[i-1]]=[shapes[i-1]!,shapes[i]!];}
  return command(snapshot,[{op:'set',path:['slides',index,'shapes'],value:shapes}]);
}
export type ShapeAlignment='left'|'center'|'right'|'top'|'middle'|'bottom';
export function alignShapes(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[],alignment:ShapeAlignment,toSlide=false):DocumentCommand {
  const doc=presentation(snapshot),shapes=shapeIds.map(id=>locateShape(snapshot,slideId,id).shape);assert(shapes.length>0,'invalid','No selected shapes');
  const bounds=toSlide||shapes.length===1?{left:0,top:0,right:doc.size.width,bottom:doc.size.height}:{left:Math.min(...shapes.map(s=>s.geometry.x)),top:Math.min(...shapes.map(s=>s.geometry.y)),right:Math.max(...shapes.map(s=>s.geometry.x+s.geometry.width)),bottom:Math.max(...shapes.map(s=>s.geometry.y+s.geometry.height))};
  return setShapeGeometries(snapshot,slideId,Object.fromEntries(shapes.map(s=>{const g={...s.geometry};switch(alignment){case 'left':g.x=bounds.left;break;case 'center':g.x=Math.round((bounds.left+bounds.right-g.width)/2);break;case 'right':g.x=bounds.right-g.width;break;case 'top':g.y=bounds.top;break;case 'middle':g.y=Math.round((bounds.top+bounds.bottom-g.height)/2);break;case 'bottom':g.y=bounds.bottom-g.height;break;}return [s.id,g];})));
}
export function insertSlideImage(snapshot:DocumentSnapshot,slideId:string,index:number,shapeId:string,resource:ImageResource,geometry:Geometry):DocumentCommand {
  validateImageResource(resource);validateGeometry(geometry);const doc=presentation(snapshot),{slide,index:slideIndex}=locateSlide(snapshot,slideId);assert(Number.isSafeInteger(index)&&index>=0&&index<=slide.shapes.length,'invalid','Invalid shape insertion index');
  const existing=doc.resources[resource.id];assert(!existing||existing.sha256===resource.sha256&&existing.dataBase64===resource.dataBase64,'conflict','Image resource mismatch');
  const operations:PatchOperation[]=[];if(!existing)operations.push({op:'set',path:['resources',resource.id],value:resource});
  operations.push({op:'splice',path:['slides',slideIndex,'shapes'],index,deleteCount:0,items:[{id:shapeId,type:'image',resourceId:resource.id,geometry}]});return command(snapshot,operations);
}
export function aspectFitImage(resource:Pick<ImageResource,'pixelWidth'|'pixelHeight'>,maxWidth:number,maxHeight:number):{width:number;height:number}{
  assert(Number.isFinite(maxWidth)&&Number.isFinite(maxHeight)&&maxWidth>0&&maxHeight>0&&resource.pixelWidth>0&&resource.pixelHeight>0,'invalid','Invalid image fit geometry');const scale=Math.min(maxWidth/resource.pixelWidth,maxHeight/resource.pixelHeight);return {width:Math.max(1,Math.round(resource.pixelWidth*scale)),height:Math.max(1,Math.round(resource.pixelHeight*scale))};
}
export function shapeText(shape:ShapeSnapshot):string {assert(shape.type==='text','invalid','Expected a text shape');return shape.runs.map(run=>run.text).join('');}
function textRange(shape:Extract<ShapeSnapshot,{type:'text'}>,start:number,end:number):void {const text=shapeText(shape),boundaries=new Set(graphemeBoundaries(text));assert(Number.isSafeInteger(start)&&Number.isSafeInteger(end)&&start<=end&&boundaries.has(start)&&boundaries.has(end),'invalid','Text range splits a grapheme');}
/** Preserve existing rich runs outside the logical replacement range. */
export function replaceShapeText(snapshot:DocumentSnapshot,slideId:string,shapeId:string,start:number,end:number,text:string,factory:IdFactory):DocumentCommand {
  const l=locateShape(snapshot,slideId,shapeId),shape=l.shape;assert(shape.type==='text','invalid','Expected a text shape');textRange(shape,start,end);assertUnicode(text);assert(text.length<=1_000_000,'limit','Text shape budget exceeded');
  if(start===end&&!text)return command(snapshot,[]);
  const id=uniqueIds(snapshot,factory),before:TextRun[]=[],after:TextRun[]=[];let at=0,inherit:DeepReadonly<TextRun>|undefined;
  for(const run of shape.runs){const stop=at+run.text.length;if(at<=start&&start<=stop)inherit??=run;if(!run.text.length)continue;
    if(stop<=start)before.push(cloneData(run) as TextRun);else if(at<start)before.push({...cloneData(run),text:run.text.slice(0,start-at)} as TextRun);
    if(at>=end)after.push(cloneData(run) as TextRun);else if(stop>end)after.push({...cloneData(run),id:id(),text:run.text.slice(end-at)} as TextRun);at=stop;
  }
  const insertion:TextRun[]=[];if(text.length)insertion.push({id:id(),type:'text',text,...(inherit?.styleId?{styleId:inherit.styleId}:{}),...(inherit?.style?{style:cloneData(inherit.style) as TextStyle}:{})});
  const runs=[...before,...insertion,...after];if(!runs.length)runs.push({id:id(),type:'text',text:''});
  return command(snapshot,[{op:'set',path:['slides',l.slideIndex,'shapes',l.index,'runs'],value:runs}]);
}
/** Difference boundaries expand to complete graphemes in both old and new text. */
export function slideTextEditDiff(before:string,after:string):{start:number;end:number;text:string}{
  let start=0;while(start<before.length&&start<after.length&&before[start]===after[start])start++;
  const oldBoundaries=new Set(graphemeBoundaries(before)),newBoundaries=new Set(graphemeBoundaries(after));while(start>0&&(!oldBoundaries.has(start)||!newBoundaries.has(start)))start--;
  let end=before.length,newEnd=after.length;while(end>start&&newEnd>start&&before[end-1]===after[newEnd-1]){end--;newEnd--;}
  while(end<before.length&&newEnd<after.length&&(!oldBoundaries.has(end)||!newBoundaries.has(newEnd))){end++;newEnd++;}
  return {start,end,text:after.slice(start,newEnd)};
}
export function formatShapeText(snapshot:DocumentSnapshot,slideId:string,shapeId:string,style:TextStyle,range?:{start:number;end:number},factory?:IdFactory):DocumentCommand {
  validateTextStyle(style);const l=locateShape(snapshot,slideId,shapeId),shape=l.shape;assert(shape.type==='text','invalid','Expected text shape');
  if(!range||range.start===range.end)return command(snapshot,[{op:'set',path:['slides',l.slideIndex,'shapes',l.index,'textStyle'],value:{...shape.textStyle,...style}},{op:'set',path:['slides',l.slideIndex,'shapes',l.index,'runs'],value:shape.runs.map(run=>({...run,style:{...run.style,...style}}))}]);
  textRange(shape,range.start,range.end);assert(factory,'invalid','A split-run ID factory is required');const id=uniqueIds(snapshot,factory),runs:TextRun[]=[];let at=0;
  for(const run of shape.runs){const stop=at+run.text.length;if(stop<=range.start||at>=range.end){runs.push(cloneData(run) as TextRun);at=stop;continue;}
    const first=Math.max(0,range.start-at),last=Math.min(run.text.length,range.end-at);if(first>0)runs.push({...cloneData(run),text:run.text.slice(0,first)} as TextRun);
    runs.push({...cloneData(run),id:first>0?id():run.id,text:run.text.slice(first,last),style:{...run.style,...style}} as TextRun);
    if(last<run.text.length)runs.push({...cloneData(run),id:id(),text:run.text.slice(last)} as TextRun);at=stop;
  }
  return command(snapshot,[{op:'set',path:['slides',l.slideIndex,'shapes',l.index,'runs'],value:runs}]);
}
export function setShapeTextFrame(snapshot:DocumentSnapshot,slideId:string,shapeId:string,frame:SlideTextFrame):DocumentCommand {
  const l=locateShape(snapshot,slideId,shapeId);assert(l.shape.type==='text','invalid','Expected text shape');return command(snapshot,[{op:'set',path:['slides',l.slideIndex,'shapes',l.index,'textFrame'],value:{...l.shape.textFrame,...frame}}]);
}

/** Paste copies model objects and immutable image assets, never active clipboard markup. */
export function pasteShapes(snapshot:DocumentSnapshot,slideId:string,shapes:readonly ShapeSnapshot[],resources:Readonly<Record<string,DeepReadonly<ImageResource>>>,factory:IdFactory):DocumentCommand {
  const doc=presentation(snapshot),{slide,index}=locateSlide(snapshot,slideId),id=uniqueIds(snapshot,factory),operations:PatchOperation[]=[],added=new Set<string>();
  const copies=shapes.map(s=>{const copy=cloneData(s) as Shape;copy.id=id();copy.geometry.x+=190500;copy.geometry.y+=190500;if(copy.type==='text')for(const run of copy.runs)run.id=id();if(copy.type==='image'){const resource=resources[copy.resourceId];assert(resource,'invalid','Copied image payload is missing');validateImageResource(resource);const existing=doc.resources[copy.resourceId];assert(!existing||existing.sha256===resource.sha256&&existing.dataBase64===resource.dataBase64,'conflict','Copied image payload differs');if(!existing&&!added.has(copy.resourceId)){added.add(copy.resourceId);operations.push({op:'set',path:['resources',copy.resourceId],value:resource});}}return copy;});
  operations.push({op:'splice',path:['slides',index,'shapes'],index:slide.shapes.length,deleteCount:0,items:copies});return command(snapshot,operations);
}
