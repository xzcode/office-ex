// SPDX-License-Identifier: MIT
/** A revision-bound scene shared by editing, thumbnails, presentation and print. */
import { assert, type Diagnostic } from '@office-ex/contracts';
import { freezeData, resolveTextStyle, resolveSlideThemeColor, resolveSlideThemeFont, type DeepReadonly, type Color, type DocumentSnapshot, type Geometry, type ResourceRef, type Shape, type SlideTheme, type TextStyle } from '@office-ex/core';

type Rgba=Extract<Color,{kind:'rgba'}>;
export interface SlideSceneRun {readonly id:string;readonly text:string;readonly style:DeepReadonly<TextStyle>}
export interface SlideSceneParagraph {readonly runs:readonly SlideSceneRun[];readonly style:DeepReadonly<TextStyle>;readonly direction:'ltr'|'rtl'|'auto'}
interface SceneBase {readonly id:string;readonly geometry:DeepReadonly<Geometry>;readonly fill?:Rgba;readonly stroke?:{readonly color:Rgba;readonly width:number}}
export type SlideSceneShape=(SceneBase&{readonly type:'rect'|'ellipse'|'line'})|(SceneBase&{readonly type:'image';readonly resourceId:string;readonly resource?:DeepReadonly<ResourceRef>})|(SceneBase&{readonly type:'text';readonly paragraphs:readonly SlideSceneParagraph[];readonly textStyle:DeepReadonly<TextStyle>;readonly wrap:boolean;readonly verticalAlignment:'top'|'center'|'bottom';readonly margins:{readonly top:number;readonly right:number;readonly bottom:number;readonly left:number};readonly direction:'ltr'|'rtl'|'auto'});
export interface SlideScene {
  readonly documentId:string;readonly revision:number;readonly contentId:string;readonly slideId:string;readonly width:number;readonly height:number;readonly background:Rgba;readonly shapes:readonly SlideSceneShape[];readonly diagnostics:readonly Diagnostic[];
}
export const resolveSlideColor=resolveSlideThemeColor;
export function buildSlideScene(snapshot:DocumentSnapshot,slideId:string):SlideScene {
  const doc=snapshot.document;assert(doc.kind==='slides','invalid','Expected presentation snapshot');const slide=doc.slides.find(item=>item.id===slideId);assert(slide,'invalid','Slide does not exist');
  const diagnostics:Diagnostic[]=[],color=(value:DeepReadonly<Color>):Rgba=>resolveSlideColor(value,doc.theme);
  const resolve=(shape:DeepReadonly<Extract<Shape,{type:'text'}>>,styleId?:string,style:DeepReadonly<TextStyle>={}):DeepReadonly<TextStyle>=>{
    const result=resolveTextStyle({...doc.defaultTextStyle,...shape.textStyle},doc.styles,styleId?[styleId]:[],style).style;
    return {...result,...(result.color?{color:color(result.color)}:{}),...(result.fontFamily?{fontFamily:resolveSlideThemeFont(result.fontFamily,doc.theme)}:{})};
  };
  const shapes:SlideSceneShape[]=slide.shapes.map(shape=>{
    const base:SceneBase={id:shape.id,geometry:shape.geometry,...(shape.fill?{fill:color(shape.fill)}:{}),...(shape.stroke?{stroke:{color:color(shape.stroke.color),width:shape.stroke.width}}:{})};
    if(shape.type==='image'){
      const resource=doc.resources[shape.resourceId];if(!resource?.dataBase64)diagnostics.push({code:'slides.image-unavailable',featureId:'SLD-02',severity:'warning',messageKey:'slides.imageUnavailable',safeAction:'none'});
      return {...base,type:'image',resourceId:shape.resourceId,...(resource?{resource}:{})};
    }
    if(shape.type!=='text')return {...base,type:shape.type};
    const paragraphs:SlideSceneParagraph[]=[],textStyle=resolve(shape),direction=shape.textFrame?.direction??textStyle.paragraphDirection??'auto';let runs:SlideSceneRun[]=[];
    const flush=()=>{const style=runs[0]?.style??textStyle;paragraphs.push({runs,style,direction:shape.textFrame?.direction&&shape.textFrame.direction!=='auto'?shape.textFrame.direction:style.paragraphDirection??direction});runs=[];};
    for(const run of shape.runs){const pieces=run.text.split('\n'),style=resolve(shape,run.styleId,run.style);for(let i=0;i<pieces.length;i++){runs.push({id:run.id+':'+i,text:pieces[i]!,style});if(i<pieces.length-1)flush();}}
    flush();
    return {...base,type:'text',paragraphs,textStyle,wrap:shape.textFrame?.wrap??true,verticalAlignment:shape.textFrame?.verticalAlignment??'top',margins:shape.textFrame?.margins??{top:45720,right:91440,bottom:45720,left:91440},direction};
  });
  return freezeData({documentId:doc.documentId,revision:snapshot.revision,contentId:snapshot.contentId,slideId,width:doc.size.width,height:doc.size.height,background:color(slide.background),shapes,diagnostics});
}
