// SPDX-License-Identifier: MIT
/** Presentation commands use stable IDs and one atomic patch per user action. */
import { assert, assertId, assertUnicode, graphemeBoundaries } from '@office-ex/contracts';
import { DEFAULT_SLIDE_TABLE_STYLE, slideTableCoverage, validateSlideTableData, cloneData, patchCommand, DEFAULT_SLIDE_THEME_COLORS, validateColor, validateGeometry, validateSlideCrop, walkSlideShapes, resolveSlideConnector, isConnectorTarget, nearestSlideConnections, connectorGeometry, validateTextStyle, validateSlideParagraph, resolveTextStyle, splitSlideParagraphs, firstSlideParagraphRun, validateImageResource, type DocumentSnapshot, type DocumentCommand, type DeepReadonly, type SlidesDocument, type Slide, type Shape, type Geometry, type Color, type TextRun, type TextStyle, type SlideTextFrame, type SlideTableData, type SlideTableRange, type SlideTableCellFormat, type SlideCrop, type SlideStroke, type SlideConnections, type ModelPath, type SlideParagraphFormat, type SlideTheme, type ImageResource, type PatchOperation } from '@office-ex/core';

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
export interface ShapeLocation {shape:ShapeSnapshot;slideIndex:number;index:number;path:ModelPath;containerPath:ModelPath;siblings:readonly ShapeSnapshot[];parent?:Extract<ShapeSnapshot,{type:'group'}>}
export function locateShape(snapshot:DocumentSnapshot,slideId:string,shapeId:string):ShapeLocation{
 const {slide,index:slideIndex}=locateSlide(snapshot,slideId);
 const search=(shapes:readonly ShapeSnapshot[],containerPath:ModelPath,parent?:Extract<ShapeSnapshot,{type:'group'}>):ShapeLocation|undefined=>{
  for(const [index,shape]of shapes.entries()){const path=[...containerPath,index];if(shape.id===shapeId)return {shape,slideIndex,index,path,containerPath,siblings:shapes,...(parent?{parent}:{})};if(shape.type==='group'){const found=search(shape.children,[...path,'children'],shape);if(found)return found;}}
 };
 const found=search(slide.shapes,['slides',slideIndex,'shapes']);assert(found,'invalid','Shape does not exist');return found;
}
function selectionContainer(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[]):{shapes:readonly ShapeSnapshot[];path:ModelPath;parent?:Extract<ShapeSnapshot,{type:'group'}>}{
 const l=locateSlide(snapshot,slideId);if(!shapeIds.length)return {shapes:l.slide.shapes,path:['slides',l.index,'shapes']};
 const locations=shapeIds.map(id=>locateShape(snapshot,slideId,id)),first=locations[0]!;
 assert(new Set(shapeIds).size===shapeIds.length&&locations.every(x=>x.siblings===first.siblings),'unsupported','Select objects within one group level');
 return {shapes:first.siblings,path:first.containerPath,parent:first.parent};
}
function resolvedShapes(shapes:readonly ShapeSnapshot[]):ShapeSnapshot[]{const map=new Map(shapes.map(s=>[s.id,s]));return shapes.map(s=>resolveSlideConnector(s,map));}
function reidentify(copies:Shape[],factory:IdFactory):void{
 const mapping=new Map<string,string>();for(const item of walkSlideShapes(copies)){const old=item.id;item.id=factory();mapping.set(old,item.id);if(item.type==='text')for(const run of item.runs)run.id=factory();}
 for(const item of walkSlideShapes(copies))if(item.type==='line'&&item.connections){
  const start=mapping.get(item.connections.start.shapeId),end=mapping.get(item.connections.end.shapeId);
  if(start&&end)item.connections={start:{...item.connections.start,shapeId:start},end:{...item.connections.end,shapeId:end}};
  else delete item.connections;
 }
}
function command(snapshot: DocumentSnapshot, operations: readonly PatchOperation[]): DocumentCommand { return patchCommand(snapshot, 'SLD-02', operations); }
function uniqueIds(snapshot:DocumentSnapshot, factory:IdFactory):IdFactory {
  const doc=presentation(snapshot),ids=new Set(doc.slides.flatMap(s=>[s.id,...[...walkSlideShapes(s.shapes)].flatMap(shape=>[shape.id,...(shape.type==='text'?shape.runs.map(r=>r.id):[])])]));
  return ()=>{const id=factory();assertId(id);assert(!ids.has(id),'invalid','New node ID is not unique');ids.add(id);return id;};
}
export function addSlide(snapshot: DocumentSnapshot, index: number, id: string, background: Color = { kind:'rgba',red:255,green:255,blue:255,alpha:255 }): DocumentCommand {
  const doc=presentation(snapshot);assertId(id);validateColor(background);assert(Number.isSafeInteger(index)&&index>=0&&index<=doc.slides.length,'invalid','Invalid slide insertion index');
  return command(snapshot,[{op:'splice',path:['slides'],index,deleteCount:0,items:[{id,background,shapes:[]}]}]);
}
export function duplicateSlide(snapshot:DocumentSnapshot,slideId:string,factory:IdFactory):DocumentCommand {
  const {slide,index}=locateSlide(snapshot,slideId),id=uniqueIds(snapshot,factory),copy=cloneData(slide) as Slide;copy.id=id();
  copy.shapes=cloneData(resolvedShapes(slide.shapes)) as Shape[];reidentify(copy.shapes,id);
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
export function removeShapes(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[]):DocumentCommand{
 const {shapes,path,parent}=selectionContainer(snapshot,slideId,shapeIds),set=new Set(shapeIds);assert(!parent||shapes.length>set.size,'unsupported','A group must retain at least one child');
 const resolved=new Map(resolvedShapes(shapes).map(shape=>[shape.id,shape]));const kept=shapes.filter(s=>!set.has(s.id)).map(shape=>{
  if(shape.type==='line'&&shape.connections&&(set.has(shape.connections.start.shapeId)||set.has(shape.connections.end.shapeId))){const copy=cloneData(resolved.get(shape.id)!) as Shape;if(copy.type==='line')delete copy.connections;return copy;}
  return shape;
 });
 return command(snapshot,[{op:'set',path,value:kept}]);
}
export function duplicateShapes(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[],factory:IdFactory,offset={x:190500,y:190500}):DocumentCommand{
 const {shapes,path}=selectionContainer(snapshot,slideId,shapeIds),set=new Set(shapeIds),copies=cloneData(resolvedShapes(shapes).filter(s=>set.has(s.id))) as Shape[];
 for(const copy of copies){copy.geometry.x+=offset.x;copy.geometry.y+=offset.y;}
 reidentify(copies,uniqueIds(snapshot,factory));
 return command(snapshot,[{op:'splice',path,index:shapes.length,deleteCount:0,items:copies}]);
}
/** Transient pointer previews are outside this API; a completed gesture uses one command. */
export function setShapeGeometries(snapshot:DocumentSnapshot,slideId:string,values:Readonly<Record<string,Geometry>>):DocumentCommand {
  const operations:PatchOperation[]=[];
  for(const [id,geometry]of Object.entries(values)){
    validateGeometry(geometry);const located=locateShape(snapshot,slideId,id),shape=located.shape;
    assert(!(shape.type==='line'&&shape.connections),'unsupported','Detach the connector before changing its geometry');
    if(shape.type==='table'){
      assert(geometry.rotation===0,'unsupported','Tables cannot be rotated');
      const scale=(parts:readonly number[],total:number)=>{
        assert(total>=parts.length,'invalid','Table dimensions too small');
        const old=parts.reduce((a,b)=>a+b,0);let cursor=0,last=0;
        return parts.map((p,i)=>{cursor+=p;const next=i===parts.length-1?total:Math.max(last+1,Math.min(total-(parts.length-i-1),Math.round(cursor/old*total)));const size=next-last;last=next;return size;});
      };
      operations.push({op:'set',path:[...located.path,'table','columnWidths'],value:scale(shape.table.columnWidths,geometry.width)},{op:'set',path:[...located.path,'table','rowHeights'],value:scale(shape.table.rowHeights,geometry.height)});
    }
    operations.push({op:'set',path:[...located.path,'geometry'],value:geometry});
  }
  return command(snapshot,operations);
}
export function setShapeGeometry(snapshot:DocumentSnapshot,slideId:string,shapeId:string,geometry:Geometry):DocumentCommand {return setShapeGeometries(snapshot,slideId,{[shapeId]:geometry});}
export function nudgeShapes(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[],dx:number,dy:number):DocumentCommand {
  assert(Number.isSafeInteger(dx)&&Number.isSafeInteger(dy),'invalid','Invalid shape movement');
  return setShapeGeometries(snapshot,slideId,Object.fromEntries(shapeIds.map(id=>{const g=locateShape(snapshot,slideId,id).shape.geometry;return [id,{...g,x:g.x+dx,y:g.y+dy}];})));
}
export function setShapePaint(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[],paint:{fill?:Color|null;stroke?:SlideStroke|null}):DocumentCommand {
  const operations:PatchOperation[]=[];if(paint.fill)validateColor(paint.fill);if(paint.stroke){validateColor(paint.stroke.color);assert(Number.isSafeInteger(paint.stroke.width)&&paint.stroke.width>=0,'invalid','Invalid stroke width');}
  for(const id of shapeIds){const l=locateShape(snapshot,slideId,id),base=l.path;assert(l.shape.type!=='group'&&l.shape.type!=='table','unsupported','This object requires its own paint controls');for(const key of ['fill','stroke'] as const)if(paint[key]!==undefined)operations.push(paint[key]===null?{op:'delete',path:[...base,key]}:{op:'set',path:[...base,key],value:paint[key]});}
  return command(snapshot,operations);
}
export type LayerMove='front'|'back'|'forward'|'backward';
export function reorderShapes(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[],move:LayerMove):DocumentCommand {
  const container=selectionContainer(snapshot,slideId,shapeIds),selected=new Set(shapeIds),shapes=[...container.shapes];
  if(move==='front'||move==='back'){const chosen=shapes.filter(s=>selected.has(s.id)),rest=shapes.filter(s=>!selected.has(s.id));return command(snapshot,[{op:'set',path:container.path,value:move==='front'?[...rest,...chosen]:[...chosen,...rest]}]);}
  if(move==='forward')for(let i=shapes.length-2;i>=0;i--){if(selected.has(shapes[i]!.id)&&!selected.has(shapes[i+1]!.id))[shapes[i],shapes[i+1]]=[shapes[i+1]!,shapes[i]!];}
  else for(let i=1;i<shapes.length;i++){if(selected.has(shapes[i]!.id)&&!selected.has(shapes[i-1]!.id))[shapes[i],shapes[i-1]]=[shapes[i-1]!,shapes[i]!];}
  return command(snapshot,[{op:'set',path:container.path,value:shapes}]);
}
export type ShapeAlignment='left'|'center'|'right'|'top'|'middle'|'bottom';
export function alignShapes(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[],alignment:ShapeAlignment,toSlide=false):DocumentCommand {
  const doc=presentation(snapshot),container=selectionContainer(snapshot,slideId,shapeIds),shapes=shapeIds.map(id=>locateShape(snapshot,slideId,id).shape);assert(shapes.length>0,'invalid','No selected shapes');
  const space=container.parent?.childSpace??{x:0,y:0,width:doc.size.width,height:doc.size.height};const bounds=toSlide||shapes.length===1?{left:space.x,top:space.y,right:space.x+space.width,bottom:space.y+space.height}:{left:Math.min(...shapes.map(s=>s.geometry.x)),top:Math.min(...shapes.map(s=>s.geometry.y)),right:Math.max(...shapes.map(s=>s.geometry.x+s.geometry.width)),bottom:Math.max(...shapes.map(s=>s.geometry.y+s.geometry.height))};
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
  const id=uniqueIds(snapshot,factory),before:TextRun[]=[],after:TextRun[]=[];let at=0,inherit:DeepReadonly<TextRun>|undefined,previous:DeepReadonly<TextRun>|undefined;
  for(const run of shape.runs){const stop=at+run.text.length;if(at<=start&&start<stop)inherit??=run;if(stop<=start)previous=run;if(!run.text.length){if(at===start)inherit??=run;continue;}
    if(stop<=start)before.push(cloneData(run) as TextRun);else if(at<start)before.push({...cloneData(run),text:run.text.slice(0,start-at)} as TextRun);
    if(at>=end)after.push(cloneData(run) as TextRun);else if(stop>end)after.push({...cloneData(run),id:id(),text:run.text.slice(end-at)} as TextRun);at=stop;
  }
  inherit??=previous;
  const insertion:TextRun[]=[];if(text.length)insertion.push({id:id(),type:'text',text,...(inherit?.styleId?{styleId:inherit.styleId}:{}),...(inherit?.style?{style:cloneData(inherit.style) as TextStyle}:{})});
  const runs=[...before,...insertion,...after];if(!runs.length)runs.push({id:id(),type:'text',text:'',...(inherit?.styleId?{styleId:inherit.styleId}:{}),...(inherit?.style?{style:cloneData(inherit.style) as TextStyle}:{})});
  return command(snapshot,[{op:'set',path:[...l.path,'runs'],value:runs}]);
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
  if(!range||range.start===range.end)return command(snapshot,[{op:'set',path:[...l.path,'textStyle'],value:{...shape.textStyle,...style}},{op:'set',path:[...l.path,'runs'],value:shape.runs.map(run=>({...run,style:{...run.style,...style}}))}]);
  textRange(shape,range.start,range.end);assert(factory,'invalid','A split-run ID factory is required');const id=uniqueIds(snapshot,factory),runs:TextRun[]=[];let at=0;
  for(const run of shape.runs){const stop=at+run.text.length;if(stop<=range.start||at>=range.end){runs.push(cloneData(run) as TextRun);at=stop;continue;}
    const first=Math.max(0,range.start-at),last=Math.min(run.text.length,range.end-at);if(first>0)runs.push({...cloneData(run),text:run.text.slice(0,first)} as TextRun);
    runs.push({...cloneData(run),id:first>0?id():run.id,text:run.text.slice(first,last),style:{...run.style,...style}} as TextRun);
    if(last<run.text.length)runs.push({...cloneData(run),id:id(),text:run.text.slice(last)} as TextRun);at=stop;
  }
  return command(snapshot,[{op:'set',path:[...l.path,'runs'],value:runs}]);
}
/** A selection is expanded to complete logical paragraphs. Formatting changes
 * never insert marker characters into the document text. */
export function formatSlideParagraphs(snapshot:DocumentSnapshot,slideId:string,shapeId:string,range:{start:number;end:number},patch:Partial<SlideParagraphFormat>,factory:IdFactory):DocumentCommand{
 const l=locateShape(snapshot,slideId,shapeId),shape=l.shape,doc=presentation(snapshot);assert(shape.type==='text','invalid','Expected text shape');textRange(shape,range.start,range.end);
 validateSlideParagraph(patch);
 const paragraphs=splitSlideParagraphs(shape.runs),ids=uniqueIds(snapshot,factory),runs:TextRun[]=[];
 let at=0;
 const chosen=(start:number,end:number)=>range.start===range.end?range.start>=start&&range.start<=end:range.start<=end&&range.end>start;
 for(const [i,items] of paragraphs.entries()){
  const actual=firstSlideParagraphRun(items),effective=resolveTextStyle({...doc.defaultTextStyle,...shape.textStyle},doc.styles,actual?.styleId?[actual.styleId]:[],actual?.style??{}).style;
  const end=at+items.reduce((sum,r)=>sum+r.text.length,0),change=chosen(at,end);
  const format=change?{...effective.slideParagraph,...patch}:undefined;
  if(i){const separator:TextRun={id:ids(),type:'text',text:'\n',...(actual?.styleId?{styleId:actual.styleId}:{}),...(actual?.style?{style:cloneData(actual.style) as TextStyle}:{})};if(format){separator.style={...separator.style,slideParagraph:format};if(format.beforePercent!==undefined)delete separator.style.spaceBefore;if(format.afterPercent!==undefined)delete separator.style.spaceAfter;}runs.push(separator);}
  const kept=items.filter(r=>r.text.length);
  if(!kept.length&&actual)kept.push({...actual,text:''});
  for(const item of kept.length?kept:[{id:ids(),type:'text' as const,text:''}]){
    const value=cloneData(item) as TextRun;value.id=ids();
    if(format){const style={...value.style,slideParagraph:format};if(format.beforePercent!==undefined)delete style.spaceBefore;if(format.afterPercent!==undefined)delete style.spaceAfter;value.style=style;}
    runs.push(value);
  }
  at=end+1;
 }
 return command(snapshot,[{op:'set',path:[...l.path,'runs'],value:runs}]);
}
export function setShapeTextFrame(snapshot:DocumentSnapshot,slideId:string,shapeId:string,frame:SlideTextFrame):DocumentCommand {
  const l=locateShape(snapshot,slideId,shapeId);assert(l.shape.type==='text','invalid','Expected text shape');return command(snapshot,[{op:'set',path:[...l.path,'textFrame'],value:{...l.shape.textFrame,...frame}}]);
}

/** Paste copies model objects and immutable image assets, never active clipboard markup. */
export function pasteShapes(snapshot:DocumentSnapshot,slideId:string,shapes:readonly ShapeSnapshot[],resources:Readonly<Record<string,DeepReadonly<ImageResource>>>,factory:IdFactory,parentId?:string):DocumentCommand {
  const doc=presentation(snapshot),{slide,index}=locateSlide(snapshot,slideId),id=uniqueIds(snapshot,factory),operations:PatchOperation[]=[],added=new Set<string>();
  const parent=parentId?locateShape(snapshot,slideId,parentId):undefined;assert(!parent||parent.shape.type==='group','invalid','Expected group paste target');
  const target=parent&&parent.shape.type==='group'?{path:[...parent.path,'children'],count:parent.shape.children.length}:{path:['slides',index,'shapes'],count:slide.shapes.length};
  const copies=shapes.map(s=>{const copy=cloneData(s) as Shape;copy.geometry.x+=190500;copy.geometry.y+=190500;for(const item of walkSlideShapes([copy])){if(item.type==='image'){const resource=resources[item.resourceId];assert(resource,'invalid','Copied image payload is missing');validateImageResource(resource);const existing=doc.resources[item.resourceId];assert(!existing||existing.sha256===resource.sha256&&existing.dataBase64===resource.dataBase64,'conflict','Copied image payload differs');if(!existing&&!added.has(item.resourceId)){added.add(item.resourceId);operations.push({op:'set',path:['resources',item.resourceId],value:resource});}}}return copy;});
  reidentify(copies,id);if(parent)for(const shape of copies)if(shape.type==='line')delete shape.connections;operations.push({op:'splice',path:target.path,index:target.count,deleteCount:0,items:copies});return command(snapshot,operations);
}

/** Group contiguous paint-order objects without changing their world geometry. */
export function groupShapes(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[],factory:IdFactory):DocumentCommand{
 const container=selectionContainer(snapshot,slideId,shapeIds),set=new Set(shapeIds),indices=container.shapes.flatMap((shape,i)=>set.has(shape.id)?[i]:[]);
 assert(set.size===shapeIds.length&&indices.length===shapeIds.length&&indices.length>=2,'invalid','Select at least two distinct shapes');
 assert(indices.at(-1)!-indices[0]!+1===indices.length,'unsupported','Grouping noncontiguous paint layers would change their stacking order');
 assert(!container.shapes.some(shape=>shape.type==='line'&&shape.connections&&(set.has(shape.id)||set.has(shape.connections.start.shapeId)||set.has(shape.connections.end.shapeId))),'unsupported','Detach connectors before grouping their targets');
 const selected=container.shapes.slice(indices[0]!,indices.at(-1)!+1);assert(!selected.some(s=>s.type==='table'),'unsupported','Tables cannot be grouped');const points=selected.flatMap(({geometry:g})=>{
  const a=g.rotation/60000*Math.PI/180,c=Math.cos(a),s=Math.sin(a),cx=g.x+g.width/2,cy=g.y+g.height/2;
  return [[-g.width/2,-g.height/2],[g.width/2,-g.height/2],[g.width/2,g.height/2],[-g.width/2,g.height/2]].map(([x,y])=>({x:cx+x!*c-y!*s,y:cy+x!*s+y!*c}));
 });
 const x=Math.floor(Math.min(...points.map(p=>p.x))),y=Math.floor(Math.min(...points.map(p=>p.y))),width=Math.max(1,Math.ceil(Math.max(...points.map(p=>p.x)))-x),height=Math.max(1,Math.ceil(Math.max(...points.map(p=>p.y)))-y);
 const group:Shape={type:'group',id:uniqueIds(snapshot,factory)(),geometry:{x,y,width,height,rotation:0},childSpace:{x,y,width,height},children:cloneData(selected) as Shape[]};
 return command(snapshot,[{op:'splice',path:container.path,index:indices[0]!,deleteCount:indices.length,items:[group]}]);
}
/** Exact rigid ungrouping; refuse scaled/flipped groups instead of flattening
 * text metrics, strokes or non-uniform transforms incorrectly. */
export function canUngroupShape(shape:ShapeSnapshot):boolean{
 return shape.type==='group'&&!shape.flipH&&!shape.flipV&&shape.geometry.width===shape.childSpace.width&&shape.geometry.height===shape.childSpace.height;
}
export function ungroupShape(snapshot:DocumentSnapshot,slideId:string,shapeId:string):DocumentCommand{
 const l=locateShape(snapshot,slideId,shapeId),shape=l.shape;assert(shape.type==='group'&&canUngroupShape(shape),'unsupported','Restore group scale and flips before ungrouping');
 const g=shape.geometry,space=shape.childSpace,a=g.rotation/60000*Math.PI/180,c=Math.cos(a),s=Math.sin(a),cx=g.x+g.width/2,cy=g.y+g.height/2;
 const children=cloneData(shape.children) as Shape[];
 for(const child of children){const v=child.geometry,x=v.x+v.width/2-space.x-space.width/2,y=v.y+v.height/2-space.y-space.height/2;child.geometry={...v,x:Math.round(cx+c*x-s*y-v.width/2),y:Math.round(cy+s*x+c*y-v.height/2),rotation:((v.rotation+g.rotation)%21600000+21600000)%21600000};}
 return command(snapshot,[{op:'splice',path:l.containerPath,index:l.index,deleteCount:1,items:children}]);
}
export function setSlideImageCrop(snapshot:DocumentSnapshot,slideId:string,shapeId:string,crop:SlideCrop):DocumentCommand{
 validateSlideCrop(crop);const l=locateShape(snapshot,slideId,shapeId);assert(l.shape.type==='image','invalid','Expected a picture');
 const path=[...l.path,'crop'];
 return command(snapshot,Object.values(crop).some(v=>v!==0)?[{op:'set',path,value:crop}]:l.shape.crop?[{op:'delete',path}]:[]);
}
export function flipSlideShapes(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[],axis:'flipH'|'flipV'):DocumentCommand{
 return command(snapshot,shapeIds.map(shapeId=>{const l=locateShape(snapshot,slideId,shapeId);assert(l.shape.type!=='text'&&l.shape.type!=='table'&&!(l.shape.type==='line'&&l.shape.connections),'unsupported','Cannot flip text or an attached connector');return {op:'set',path:[...l.path,axis],value:!l.shape[axis]};}));
}

/** Create one attached straight connector; endpoints are stable model IDs. */
export function connectSlideShapes(snapshot:DocumentSnapshot,slideId:string,shapeIds:readonly string[],factory:IdFactory):DocumentCommand{
 assert(shapeIds.length===2,'invalid','Select two rectangular objects');
 const {slide}=locateSlide(snapshot,slideId),map=new Map(slide.shapes.map(s=>[s.id,s])),a=map.get(shapeIds[0]!),b=map.get(shapeIds[1]!);
 assert(isConnectorTarget(a)&&isConnectorTarget(b),'unsupported','Connections require top-level rectangular objects');
 const connections=nearestSlideConnections(a,b),shape:Shape={id:uniqueIds(snapshot,factory)(),type:'line',connector:true,connections,...connectorGeometry(connections,map),stroke:{color:{kind:'theme',name:'accent1'},width:25400,tail:{type:'triangle'}}};
 return insertShapes(snapshot,slideId,slide.shapes.length,[shape]);
}
export function setSlideConnections(snapshot:DocumentSnapshot,slideId:string,shapeId:string,connections:SlideConnections|null):DocumentCommand{
 const l=locateShape(snapshot,slideId,shapeId),shape=l.shape;assert(!l.parent&&shape.type==='line'&&shape.connector,'invalid','Expected top-level straight connector');
 const map=new Map(l.siblings.map(s=>[s.id,s])),resolved=connections?{...shape,connections,...connectorGeometry(connections,map)}:resolveSlideConnector(shape,map);
 const next=cloneData(resolved) as Shape;if(!connections&&next.type==='line')delete next.connections;
 return command(snapshot,[{op:'set',path:l.path,value:next}]);
}

/** Native DrawingML table with a bounded rectangular text-only grid. */
export function insertSlideTable(snapshot:DocumentSnapshot,slideId:string,rows:number,columns:number,shapeId:string):DocumentCommand{
 const {slide}=locateSlide(snapshot,slideId),doc=presentation(snapshot);
 assert(Number.isSafeInteger(rows)&&Number.isSafeInteger(columns)&&rows>=1&&rows<=200&&columns>=1&&columns<=64&&rows*columns<=2000,'invalid','Invalid slide table dimensions');
 assertId(shapeId);const width=Math.max(columns,Math.round(doc.size.width*.6)),height=Math.max(rows,Math.round(doc.size.height*.3));
 const split=(total:number,count:number)=>Array.from({length:count},(_,i)=>Math.floor(total*(i+1)/count)-Math.floor(total*i/count));
 const table:Shape={id:shapeId,type:'table',geometry:{x:Math.round(doc.size.width*.2),y:Math.round(doc.size.height*.3),width,height,rotation:0},table:{columnWidths:split(width,columns),rowHeights:split(height,rows),cells:Array.from({length:rows},()=>Array<string>(columns).fill('')),styleId:DEFAULT_SLIDE_TABLE_STYLE,firstRow:true,bandRow:true}};
 return insertShapes(snapshot,slideId,slide.shapes.length,[table]);
}
export function editSlideTable(snapshot:DocumentSnapshot,slideId:string,shapeId:string,cells:readonly (readonly string[])[]):DocumentCommand{
 const l=locateShape(snapshot,slideId,shapeId);assert(l.shape.type==='table','invalid','Expected slide table');
 const table=cloneData(l.shape.table) as SlideTableData;table.cells=cells.map(r=>[...r]);
 return replaceSlideTable(snapshot,slideId,shapeId,table);
}
export function replaceSlideTable(snapshot:DocumentSnapshot,slideId:string,shapeId:string,table:DeepReadonly<SlideTableData>):DocumentCommand{
 const l=locateShape(snapshot,slideId,shapeId);assert(l.shape.type==='table','invalid','Expected slide table');validateSlideTableData(table);
 const geometry={...l.shape.geometry,width:table.columnWidths.reduce((a,b)=>a+b,0),height:table.rowHeights.reduce((a,b)=>a+b,0)};
 return command(snapshot,[{op:'set',path:[...l.path,'table'],value:table},{op:'set',path:[...l.path,'geometry'],value:geometry}]);
}
export type SlideTableEdit=
 |{kind:'merge';range:SlideTableRange}
 |{kind:'split';row:number;column:number}
 |{kind:'insert'|'delete';axis:'row'|'column';index:number}
 |{kind:'format';range:SlideTableRange;format:SlideTableCellFormat}
 |{kind:'resetFormat';range:SlideTableRange};
const sameCellFormat=(a:unknown,b:unknown)=>JSON.stringify(a??{})===JSON.stringify(b??{});
/** Pure shared operations for a dialog draft or a standalone atomic command. */
export function editSlideTableGrid(input:DeepReadonly<SlideTableData>,action:SlideTableEdit):SlideTableData{
 validateSlideTableData(input);const t=cloneData(input) as SlideTableData,rows=t.rowHeights.length,cols=t.columnWidths.length,coverage=slideTableCoverage(t);
 const key=(r:number,c:number)=>r+':'+c;
 const check=(r:number,c:number)=>assert(Number.isSafeInteger(r)&&Number.isSafeInteger(c)&&r>=0&&r<rows&&c>=0&&c<cols,'invalid','Invalid table cell');
 const rangeCheck=(r:SlideTableRange)=>{check(r.row,r.column);assert(Number.isSafeInteger(r.rowCount)&&r.rowCount>0&&Number.isSafeInteger(r.columnCount)&&r.columnCount>0,'invalid','Invalid table range');check(r.row+r.rowCount-1,r.column+r.columnCount-1);};
 const format=(r:number,c:number)=>t.formats?.[key(r,c)];
 const setFormat=(r:number,c:number,value:SlideTableCellFormat|undefined)=>{if(value&&Object.keys(value).length)(t.formats??={})[key(r,c)]=value;else if(t.formats)delete t.formats[key(r,c)];};
 if(action.kind==='merge'){
  const r=action.range;rangeCheck(r);assert(r.rowCount*r.columnCount>1,'invalid','Select multiple cells');
  const parts:string[]=[],anchor=format(r.row,r.column);
  for(let i=r.row;i<r.row+r.rowCount;i++)for(let j=r.column;j<r.column+r.columnCount;j++){
   assert(!coverage.has(key(i,j)),'unsupported','Split existing merges before merging this range');
   const other=format(i,j);assert(!other||sameCellFormat(other,anchor),'unsupported','Reset conflicting cell formatting before merging');
   if(t.cells[i]![j]!)parts.push(t.cells[i]![j]!);
   t.cells[i]![j]='';if(i!==r.row||j!==r.column)setFormat(i,j,undefined);
  }
  t.cells[r.row]![r.column]=parts.join('\n');(t.merges??=[]).push({...r});
 }else if(action.kind==='split'){
  check(action.row,action.column);const m=coverage.get(key(action.row,action.column));assert(m,'invalid','Cell is not merged');
  t.merges=t.merges!.filter(r=>r.row!==m.row||r.column!==m.column);
 }else if(action.kind==='format'||action.kind==='resetFormat'){
  const r=action.range;rangeCheck(r);const touched=new Set<string>();
  for(let i=r.row;i<r.row+r.rowCount;i++)for(let j=r.column;j<r.column+r.columnCount;j++){
   const m=coverage.get(key(i,j)),rr=m?.row??i,cc=m?.column??j,k=key(rr,cc);if(touched.has(k))continue;touched.add(k);
   if(action.kind==='resetFormat')setFormat(rr,cc,undefined);
   else{const prior=format(rr,cc)??{},f=action.format;setFormat(rr,cc,{...prior,...f,...(f.textStyle?{textStyle:{...prior.textStyle,...f.textStyle}}:{}),...(f.borders?{borders:{...prior.borders,...f.borders}}:{})});}
  }
 }else{
  const axis=action.axis,dim=axis==='row'?'row':'column',count=axis==='row'?'rowCount':'columnCount',sizes=axis==='row'?t.rowHeights:t.columnWidths,index=action.index,insert=action.kind==='insert';
  assert(Number.isSafeInteger(index)&&index>=0&&index<(insert?sizes.length+1:sizes.length),'invalid','Invalid table dimension index');
  assert(insert||sizes.length>1,'invalid','Retain at least one table row and column');
  const map=(v:number)=>insert?(v>=index?v+1:v):v===index?undefined:v>index?v-1:v;
  // Capture surviving merged origins before deleting any grid slots.
  const merges:SlideTableRange[]=[],rescued:{row:number;column:number;text:string;format?:SlideTableCellFormat}[]=[];
  for(const m of t.merges??[]){
   const start=m[dim],last=start+m[count]-1;let next=start,length=m[count];
   if(insert){if(index<=start)next++;else if(index<=last)length++;}
   else if(index<start)next--;
   else if(index<=last){length--;if(length>0&&index===start)rescued.push({row:dim==='row'?start:m.row,column:dim==='column'?start:m.column,text:t.cells[m.row]![m.column]!,format:format(m.row,m.column)});}
   if(length>0){const n={...m,[dim]:next,[count]:length};if(n.rowCount*n.columnCount>1)merges.push(n);}
  }
  const formats:Record<string,SlideTableCellFormat>={};
  for(const [k,f]of Object.entries(t.formats??{})){const [r,c]=k.split(':').map(Number),v=map(axis==='row'?r!:c!);if(v!==undefined)formats[key(axis==='row'?v:r!,axis==='column'?v:c!)]=f;}
  if(insert){sizes.splice(index,0,sizes[Math.max(0,Math.min(sizes.length-1,index-1))]!);if(axis==='row')t.cells.splice(index,0,Array<string>(cols).fill(''));else for(const r of t.cells)r.splice(index,0,'');}
  else{sizes.splice(index,1);if(axis==='row')t.cells.splice(index,1);else for(const r of t.cells)r.splice(index,1);}
  t.merges=merges;t.formats=formats;
  for(const item of rescued){t.cells[item.row]![item.column]=item.text;setFormat(item.row,item.column,item.format);}
 }
 if(!t.merges?.length)delete t.merges;else t.merges.sort((a,b)=>a.row-b.row||a.column-b.column);
 if(!t.formats||!Object.keys(t.formats).length)delete t.formats;
 validateSlideTableData(t);return t;
}
/** Equalize spacing in the local coordinate space. Rotated objects are
 * intentionally outside this profile; use existing manual alignment. */
export function distributeSlideShapes(snapshot:DocumentSnapshot,slideId:string,ids:readonly string[],axis:'horizontal'|'vertical'):DocumentCommand{
 const container=selectionContainer(snapshot,slideId,ids),selected=container.shapes.filter(s=>ids.includes(s.id));
 assert(selected.length>=3&&selected.length===ids.length&&selected.every(s=>s.geometry.rotation===0&&!(s.type==='line'&&s.connections)),'unsupported','Select at least three unrotated unattached shapes');
 const pos=axis==='horizontal'?'x':'y',size=axis==='horizontal'?'width':'height';
 const sorted=[...selected].sort((a,b)=>a.geometry[pos]-b.geometry[pos]||a.id.localeCompare(b.id));
 const first=sorted[0]!,last=sorted.at(-1)!,span=last.geometry[pos]+last.geometry[size]-first.geometry[pos];
 const total=sorted.reduce((n,s)=>n+s.geometry[size],0),gap=(span-total)/(sorted.length-1);
 let cursor=first.geometry[pos];const values:Record<string,Geometry>={};
 for(const s of sorted){values[s.id]={...s.geometry,[pos]:Math.round(cursor)};cursor+=s.geometry[size]+gap;}
 return setShapeGeometries(snapshot,slideId,values);
}
