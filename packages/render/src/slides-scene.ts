// SPDX-License-Identifier: MIT
/** A revision-bound scene shared by editing, thumbnails, presentation and print. */
import { assert, type Diagnostic } from '@office-ex/contracts';
import { freezeData, resolveTextStyle, resolveSlideThemeColor, resolveSlideThemeFont, splitSlideParagraphs, firstSlideParagraphRun, type SlideTextFrame, type SlidePreset, type SlideCrop, type SlideStroke, type SlideChildSpace, type DeepReadonly, type Color, type DocumentSnapshot, type Geometry, type ResourceRef, type Shape, type SlideTheme, type TextStyle } from '@office-ex/core';

type Rgba=Extract<Color,{kind:'rgba'}>;
export interface SlideSceneRun {readonly id:string;readonly text:string;readonly style:DeepReadonly<TextStyle>}
export interface SlideSceneParagraph {readonly runs:readonly SlideSceneRun[];readonly style:DeepReadonly<TextStyle>;readonly direction:'ltr'|'rtl'|'auto';readonly marker?:{readonly text:string;readonly style:DeepReadonly<TextStyle>}}
interface SceneBase {readonly id:string;readonly geometry:DeepReadonly<Geometry>;readonly fill?:Rgba;readonly stroke?:DeepReadonly<Omit<SlideStroke,'color'>&{color:Rgba}>;readonly flipH?:boolean;readonly flipV?:boolean}
export type SlideSceneShape=(SceneBase&{readonly type:SlidePreset;readonly connector?:true})|(SceneBase&{readonly type:'group';readonly childSpace:DeepReadonly<SlideChildSpace>;readonly children:readonly SlideSceneShape[]})|(SceneBase&{readonly type:'image';readonly resourceId:string;readonly resource?:DeepReadonly<ResourceRef>;readonly crop?:DeepReadonly<SlideCrop>})|(SceneBase&{readonly type:'text';readonly paragraphs:readonly SlideSceneParagraph[];readonly textStyle:DeepReadonly<TextStyle>;readonly wrap:boolean;readonly verticalAlignment:'top'|'center'|'bottom';readonly margins:{readonly top:number;readonly right:number;readonly bottom:number;readonly left:number};readonly direction:'ltr'|'rtl'|'auto';readonly autofit?:DeepReadonly<NonNullable<SlideTextFrame['autofit']>>});
export interface SlideScene {
  readonly documentId:string;readonly revision:number;readonly contentId:string;readonly slideId:string;readonly width:number;readonly height:number;readonly background:Rgba;readonly shapes:readonly SlideSceneShape[];readonly diagnostics:readonly Diagnostic[];
}
export const resolveSlideColor=resolveSlideThemeColor;
export function buildSlideScene(snapshot:DocumentSnapshot,slideId:string):SlideScene {
  const doc=snapshot.document;assert(doc.kind==='slides','invalid','Expected presentation snapshot');const slide=doc.slides.find(item=>item.id===slideId);assert(slide,'invalid','Slide does not exist');
  const diagnostics:Diagnostic[]=[],color=(value:DeepReadonly<Color>):Rgba=>resolveSlideColor(value,doc.theme);
  const resolve=(shape:DeepReadonly<Extract<Shape,{type:'text'}>>,styleId?:string,style:DeepReadonly<TextStyle>={},text=''):DeepReadonly<TextStyle>=>{
    const result=resolveTextStyle({...doc.defaultTextStyle,...shape.textStyle},doc.styles,styleId?[styleId]:[],style).style;
    const resolved={...result,...(result.color?{color:color(result.color)}:{})};for(const key of ['fontFamily','fontAscii','fontHighAnsi','fontEastAsia','fontComplex'] as const)if(result[key])resolved[key]=resolveSlideThemeFont(result[key],doc.theme,text,result.language??result.eastAsiaLanguage??'und');return resolved;
  };
  const project=(shape:DeepReadonly<Shape>):SlideSceneShape=>{
    const base:SceneBase={id:shape.id,geometry:shape.geometry,...(shape.fill?{fill:color(shape.fill)}:{}),...(shape.stroke?{stroke:{...shape.stroke,color:color(shape.stroke.color)}}:{}),...(shape.flipH?{flipH:true}:{}),...(shape.flipV?{flipV:true}:{})};
    if(shape.type==='group')return {...base,type:'group',childSpace:shape.childSpace,children:shape.children.map(project)};
    if(shape.type==='image'){
      const resource=doc.resources[shape.resourceId];if(!resource?.dataBase64)diagnostics.push({code:'slides.image-unavailable',featureId:'SLD-02',severity:'warning',messageKey:'slides.imageUnavailable',safeAction:'none'});
      return {...base,type:'image',resourceId:shape.resourceId,...(resource?{resource}:{}),...(shape.crop?{crop:shape.crop}:{})};
    }
    if(shape.type!=='text')return {...base,type:shape.type,...(shape.connector?{connector:true as const}:{})};
    const textStyle=resolve(shape),direction=shape.textFrame?.direction??textStyle.paragraphDirection??'auto';
    const counters=new Map<number,{scheme:string;value:number;start:number}>();
    const paragraphs:SlideSceneParagraph[]=splitSlideParagraphs(shape.runs.map(run=>({id:run.id,text:run.text,style:resolve(shape,run.styleId,run.style,run.text)}))).map((source,index)=>{
      const runs=source.map((run,i)=>({...run,id:run.id+':'+index+':'+i})),style=firstSlideParagraphRun(runs)?.style??textStyle,format=style.slideParagraph,level=format?.level??0,bullet=format?.bullet;
      for(const key of counters.keys())if(key>level)counters.delete(key);
      let marker:SlideSceneParagraph['marker'];
      if(bullet&&bullet.kind!=='none'){
        let text:string;
        if(bullet.kind==='char'){text=bullet.character;counters.delete(level);}
        else{const prior=counters.get(level),start=bullet.startAt??prior?.start??1;const value=prior&&prior.scheme===bullet.scheme&&prior.start===start?prior.value+1:start;counters.set(level,{scheme:bullet.scheme,value,start});text=bullet.scheme==='arabicPeriod'?value+'.':bullet.scheme==='arabicParenR'?value+')':bullet.scheme==='arabicParenBoth'?'('+value+')':String(value);}
        const size=bullet.sizeEmu??Math.round((style.fontSize??304800)*(bullet.sizePercent??100000)/100000);
        const markerStyle:TextStyle={fontFamily:resolveSlideThemeFont(bullet.font??style.fontFamily,doc.theme,text,style.language??'und'),fontSize:size,color:bullet.color?color(bullet.color):style.color,bold:style.bold,italic:style.italic,language:style.language};
        marker={text,style:markerStyle};
      }else counters.delete(level);
      return {runs,style,direction:shape.textFrame?.direction&&shape.textFrame.direction!=='auto'?shape.textFrame.direction:style.paragraphDirection??direction,...(marker?{marker}:{})};
    });
    return {...base,type:'text',paragraphs,textStyle,wrap:shape.textFrame?.wrap??true,verticalAlignment:shape.textFrame?.verticalAlignment??'top',margins:shape.textFrame?.margins??{top:45720,right:91440,bottom:45720,left:91440},direction,...(shape.textFrame?.autofit?{autofit:shape.textFrame.autofit}:{})};
  };
  const shapes=slide.shapes.map(project);
  return freezeData({documentId:doc.documentId,revision:snapshot.revision,contentId:snapshot.contentId,slideId,width:doc.size.width,height:doc.size.height,background:color(slide.background),shapes,diagnostics});
}

/** Recursively enumerate all leaf and group scene nodes for output validation. */
export function* walkSlideScene(shapes:readonly SlideSceneShape[]):Generator<SlideSceneShape>{for(const shape of shapes){yield shape;if(shape.type==='group')yield* walkSlideScene(shape.children);}}
