// SPDX-License-Identifier: MIT
/** Only trusted model values become DOM. Source markup is never inserted. */
import { EMU_PER_CSS_PIXEL, ROTATION_UNITS_PER_DEGREE } from '@office-ex/contracts';
import type { Color, DeepReadonly, TextStyle } from '@office-ex/core';
import type { SlideScene, SlideSceneShape } from '@office-ex/render';
import { bundledFontFamily } from '@office-ex/pdf';
const px=(emu:number)=>emu/EMU_PER_CSS_PIXEL+'px';
export function slideCssColor(color:DeepReadonly<Color>):string {return color.kind==='rgba'?`rgba(${color.red},${color.green},${color.blue},${color.alpha/255})`:'#182030';}
export function applySlideTextStyle(element:HTMLElement,style:DeepReadonly<TextStyle>,text=''):void {
  const complex=/[\u0590-\u08ff\u0900-\u0e7f]/u.test(text),eastAsian=/[\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff]/u.test(text),font=(complex?style.fontComplex:undefined)??(eastAsian?style.fontEastAsia:undefined)??style.fontFamily??style.fontAscii??style.fontHighAnsi??'sans-serif';
  element.style.fontFamily=bundledFontFamily(font,style.language??style.eastAsiaLanguage??'und');element.style.fontSize=px((complex?style.complexFontSize:undefined)??style.fontSize??304800);element.style.fontWeight=style.bold?'700':'400';element.style.fontStyle=style.italic?'italic':'normal';element.style.textDecoration=style.underline?'underline':'none';element.style.color=style.color?slideCssColor(style.color):'#182030';
  if(style.language)element.lang=style.language;
  if(style.runDirection){element.dir=style.runDirection;element.style.unicodeBidi='isolate';}
}
export function createSlideShapeElement(shape:SlideSceneShape):HTMLElement {
  const element=document.createElement('div');element.className='slides-scene-shape';element.dataset.shapeId=shape.id;element.dataset.shapeType=shape.type;const g=shape.geometry;
  Object.assign(element.style,{position:'absolute',left:px(g.x),top:px(g.y),width:px(g.width),height:px(g.height),transform:`rotate(${g.rotation/ROTATION_UNITS_PER_DEGREE}deg)`,transformOrigin:'50% 50%',boxSizing:'border-box',overflow:'visible'});
  if(shape.type==='line'){
    const svg=document.createElementNS('http://www.w3.org/2000/svg','svg'),line=document.createElementNS('http://www.w3.org/2000/svg','line');svg.setAttribute('viewBox',`0 0 ${Math.max(EMU_PER_CSS_PIXEL,g.width)} ${Math.max(EMU_PER_CSS_PIXEL,g.height)}`);svg.setAttribute('preserveAspectRatio','none');Object.assign(svg.style,{width:px(Math.max(EMU_PER_CSS_PIXEL,g.width)),height:px(Math.max(EMU_PER_CSS_PIXEL,g.height)),overflow:'visible'});line.setAttribute('x1','0');line.setAttribute('y1','0');line.setAttribute('x2',String(g.width));line.setAttribute('y2',String(g.height));line.setAttribute('stroke',shape.stroke?slideCssColor(shape.stroke.color):shape.fill?slideCssColor(shape.fill):'#182030');line.setAttribute('stroke-width',String(shape.stroke?.width??12700));svg.append(line);element.append(svg);
  } else {
    if(shape.fill)element.style.backgroundColor=slideCssColor(shape.fill);
    if(shape.stroke&&shape.stroke.width>0){element.style.outline=px(shape.stroke.width)+' solid '+slideCssColor(shape.stroke.color);element.style.outlineOffset=px(-shape.stroke.width/2);}
    if(shape.type==='ellipse'){element.style.borderRadius='50%';if(shape.stroke){element.style.outline='none';element.style.border=px(shape.stroke.width)+' solid '+slideCssColor(shape.stroke.color);}}
    if(shape.type==='image'){
      if(shape.resource?.dataBase64&&(shape.resource.mimeType==='image/png'||shape.resource.mimeType==='image/jpeg')){const img=document.createElement('img');img.alt='';img.draggable=false;img.src='data:'+shape.resource.mimeType+';base64,'+shape.resource.dataBase64;Object.assign(img.style,{width:'100%',height:'100%',display:'block',objectFit:'fill'});element.append(img);}else element.dataset.missingImage='true';
    } else if(shape.type==='text'){
      const frame=document.createElement('div');frame.className='slides-text-frame';frame.dataset.textFrame=shape.id;frame.dir=shape.direction;
      Object.assign(frame.style,{width:'100%',height:'100%',boxSizing:'border-box',padding:`${px(shape.margins.top)} ${px(shape.margins.right)} ${px(shape.margins.bottom)} ${px(shape.margins.left)}`,display:'flex',flexDirection:'column',justifyContent:shape.verticalAlignment==='center'?'center':shape.verticalAlignment==='bottom'?'flex-end':'flex-start',overflow:'visible'});
      const body=document.createElement('div');body.className='slides-text-body';Object.assign(body.style,{flex:'0 0 auto',minWidth:'0',whiteSpace:shape.wrap?'pre-wrap':'pre',overflowWrap:shape.wrap?'break-word':'normal',wordBreak:'normal',lineBreak:'strict'});applySlideTextStyle(body,shape.textStyle);
      for(const paragraph of shape.paragraphs){const p=document.createElement('p');p.dir=paragraph.direction;const style=paragraph.style;applySlideTextStyle(p,style,paragraph.runs.map(r=>r.text).join(''));Object.assign(p.style,{margin:'0',textAlign:style.alignment==='center'?'center':style.alignment==='end'?'end':style.alignment==='justify'?'justify':'start',textIndent:px(style.indent??0),paddingTop:px(style.spaceBefore??0),paddingBottom:px(style.spaceAfter??0),lineHeight:style.lineHeight?px(style.lineHeight):style.lineMultiple?String(style.lineMultiple/240):'1.2',minHeight:px(style.fontSize??shape.textStyle.fontSize??304800)});
        for(const run of paragraph.runs){const span=document.createElement('span');span.dataset.runId=run.id;span.textContent=run.text;applySlideTextStyle(span,run.style,run.text);p.append(span);}if(!paragraph.runs.some(r=>r.text))p.append(document.createElement('br'));body.append(p);}
      frame.append(body);element.append(frame);
    }
  }
  return element;
}
export function createSlideSceneElement(scene:SlideScene):HTMLDivElement {
  const element=document.createElement('div');element.className='slides-scene';element.dataset.slideId=scene.slideId;element.dataset.revision=String(scene.revision);element.dataset.contentId=scene.contentId;Object.assign(element.style,{position:'relative',width:px(scene.width),height:px(scene.height),overflow:'hidden',backgroundColor:slideCssColor(scene.background),isolation:'isolate',direction:'ltr',flex:'0 0 auto'});
  for(const shape of scene.shapes)element.append(createSlideShapeElement(shape));return element;
}
/** Measured only after attachment/font readiness; no measured data mutates the model. */
export function overflowingSlideText(sceneElement:HTMLElement):string[] {
  const result:string[]=[];for(const frame of sceneElement.querySelectorAll<HTMLElement>('[data-text-frame]')){const body=frame.firstElementChild as HTMLElement|null;if(!body)continue;const css=getComputedStyle(frame),height=frame.clientHeight-parseFloat(css.paddingTop)-parseFloat(css.paddingBottom),width=frame.clientWidth-parseFloat(css.paddingLeft)-parseFloat(css.paddingRight);if(body.scrollHeight>height+1||body.scrollWidth>width+1)result.push(frame.dataset.textFrame!);}return result;
}
