// SPDX-License-Identifier: MIT
/** Only trusted model values become DOM. Source markup is never inserted. */
import { EMU_PER_CSS_PIXEL, ROTATION_UNITS_PER_DEGREE } from '@office-ex/contracts';
import type { Color, DeepReadonly, TextStyle, Geometry, SlideDash, SlideArrow } from '@office-ex/core';
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
type TextScene=Extract<SlideSceneShape,{type:'text'}>;
function fittedStyle(style:DeepReadonly<TextStyle>,shape:TextScene,override?:number):DeepReadonly<TextStyle>{
 const fit=shape.autofit?.mode==='normal'?shape.autofit:undefined,scale=(override??fit?.fontScale??100000)/100000;
 return {...style,...(scale!==1?{fontSize:Math.max(1,Math.round((style.fontSize??304800)*scale)),...(style.complexFontSize!==undefined?{complexFontSize:Math.max(1,Math.round(style.complexFontSize*scale))}:{})}:{})};
}
const SVG='http://www.w3.org/2000/svg';
const dashRatios:Record<SlideDash,number[]>={solid:[],dash:[4,3],dot:[1,3],dashDot:[4,3,1,3],lgDash:[8,3],lgDashDot:[8,3,1,3],lgDashDotDot:[8,3,1,3,1,3],sysDash:[3,1],sysDot:[1,1],sysDashDot:[3,1,1,1],sysDashDotDot:[3,1,1,1,1,1]};
function svgNode<K extends keyof SVGElementTagNameMap>(tag:K,attrs:Record<string,string|number>={}):SVGElementTagNameMap[K]{const e=document.createElementNS(SVG,tag);for(const [k,v]of Object.entries(attrs))e.setAttribute(k,String(v));return e;}
/** The same geometry projection drives editing previews and the final scene. */
export function updateSlideShapeGeometry(element:HTMLElement,shape:SlideSceneShape,g:DeepReadonly<Geometry>=shape.geometry):void{
 Object.assign(element.style,{left:px(g.x),top:px(g.y),width:px(g.width),height:px(g.height),transform:`rotate(${g.rotation/ROTATION_UNITS_PER_DEGREE}deg) scale(${shape.flipH?-1:1},${shape.flipV?-1:1})`});
 if(shape.type==='group'){
  const inner=element.firstElementChild as HTMLElement|null,space=shape.childSpace;if(inner)inner.style.transform=`scale(${g.width/space.width},${g.height/space.height}) translate(${px(-space.x)},${px(-space.y)})`;
 }else if(shape.type!=='text'&&shape.type!=='image'){
  const svg=element.firstElementChild as SVGSVGElement|null;if(svg&&(svg.dataset.width!==String(g.width)||svg.dataset.height!==String(g.height))){paintVector(svg,shape,g);svg.dataset.width=String(g.width);svg.dataset.height=String(g.height);}
 }
}
function arrowGeometry(svg:SVGSVGElement,shape:SlideSceneShape,arrow:DeepReadonly<SlideArrow>,end:'head'|'tail',g:DeepReadonly<Geometry>):void{
 const w=g.width,h=g.height,n=Math.hypot(w,h);if(!n||!shape.stroke?.width)return;
 const stroke=shape.stroke.width,unit={sm:2,med:3,lg:5},length=stroke*unit[arrow.length??'med'],half=stroke*unit[arrow.width??'med']/2,ux=w/n*(end==='head'?-1:1),uy=h/n*(end==='head'?-1:1),tx=end==='head'?0:w,ty=end==='head'?0:h,bx=tx-ux*length,by=ty-uy*length;
 const points=`${bx-uy*half},${by+ux*half} ${tx},${ty} ${bx+uy*half},${by-ux*half}`;
 svg.append(svgNode(arrow.type==='triangle'?'polygon':'polyline',{points,fill:arrow.type==='triangle'?slideCssColor(shape.stroke.color):'none',stroke:arrow.type==='arrow'?slideCssColor(shape.stroke.color):'none','stroke-width':stroke,'stroke-linejoin':'miter'}));
}
function paintVector(svg:SVGSVGElement,shape:SlideSceneShape,g:DeepReadonly<Geometry>):void{
 const w=g.width,h=g.height,min=Math.min(w,h);svg.replaceChildren();svg.setAttribute('viewBox',`0 0 ${Math.max(EMU_PER_CSS_PIXEL,w)} ${Math.max(EMU_PER_CSS_PIXEL,h)}`);svg.setAttribute('preserveAspectRatio','none');Object.assign(svg.style,{width:px(Math.max(EMU_PER_CSS_PIXEL,w)),height:px(Math.max(EMU_PER_CSS_PIXEL,h)),overflow:'visible',display:'block'});
 let node:SVGElement;
 switch(shape.type){
  case 'line':node=svgNode('line',{x1:0,y1:0,x2:w,y2:h});break;
  case 'ellipse':node=svgNode('ellipse',{cx:w/2,cy:h/2,rx:w/2,ry:h/2});break;
  case 'roundRect':node=svgNode('rect',{x:0,y:0,width:w,height:h,rx:min*.16667,ry:min*.16667});break;
  case 'triangle':node=svgNode('polygon',{points:`${w/2},0 ${w},${h} 0,${h}`});break;
  case 'diamond':node=svgNode('polygon',{points:`${w/2},0 ${w},${h/2} ${w/2},${h} 0,${h/2}`});break;
  case 'parallelogram':{const x=min*.25;node=svgNode('polygon',{points:`${x},0 ${w},0 ${w-x},${h} 0,${h}`});break;}
  default:node=svgNode('rect',{x:0,y:0,width:w,height:h});
 }
 const stroke=shape.stroke;node.setAttribute('fill',shape.type==='line'?'none':shape.fill?slideCssColor(shape.fill):'none');node.setAttribute('stroke',stroke?slideCssColor(stroke.color):'none');node.setAttribute('stroke-width',String(stroke?.width??0));
 if(stroke?.dash&&stroke.dash!=='solid')node.setAttribute('stroke-dasharray',dashRatios[stroke.dash].map(v=>v*stroke.width).join(' '));
 svg.append(node);
 if(shape.type==='line'&&stroke){if(stroke.head)arrowGeometry(svg,shape,stroke.head,'head',g);if(stroke.tail)arrowGeometry(svg,shape,stroke.tail,'tail',g);}
}
export function createSlideShapeElement(shape:SlideSceneShape,fitOverride?:number):HTMLElement {
  const element=document.createElement('div');element.className='slides-scene-shape';element.dataset.shapeId=shape.id;element.dataset.shapeType=shape.type;const g=shape.geometry;
  Object.assign(element.style,{position:'absolute',left:px(g.x),top:px(g.y),width:px(g.width),height:px(g.height),transform:`rotate(${g.rotation/ROTATION_UNITS_PER_DEGREE}deg)`,transformOrigin:'50% 50%',boxSizing:'border-box',overflow:'visible'});
  if(shape.type==='group'){
    const inner=document.createElement('div');inner.className='slides-group-content';Object.assign(inner.style,{position:'absolute',left:'0',top:'0',width:px(shape.childSpace.width),height:px(shape.childSpace.height),transformOrigin:'0 0',overflow:'visible'});
    for(const child of shape.children)inner.append(createSlideShapeElement(child));element.append(inner);
  }else if(shape.type!=='text'&&shape.type!=='image'){
    const svg=svgNode('svg');element.append(svg);
  }else {
    if(shape.fill)element.style.backgroundColor=slideCssColor(shape.fill);
    if(shape.stroke&&shape.stroke.width>0){element.style.outline=px(shape.stroke.width)+' '+(shape.stroke.dash==='dot'||shape.stroke.dash==='sysDot'?'dotted':shape.stroke.dash&&shape.stroke.dash!=='solid'?'dashed':'solid')+' '+slideCssColor(shape.stroke.color);element.style.outlineOffset=px(-shape.stroke.width/2);}
    if(shape.type==='image'){
      if(shape.resource?.dataBase64&&(shape.resource.mimeType==='image/png'||shape.resource.mimeType==='image/jpeg')){
       const img=document.createElement('img');img.alt='';img.draggable=false;img.src='data:'+shape.resource.mimeType+';base64,'+shape.resource.dataBase64;
       const c=shape.crop??{left:0,top:0,right:0,bottom:0},w=100000-c.left-c.right,h=100000-c.top-c.bottom,clip=document.createElement('div');clip.className='slides-image-crop';Object.assign(clip.style,{position:'relative',width:'100%',height:'100%',overflow:'hidden'});
       Object.assign(img.style,{position:'absolute',left:-c.left/w*100+'%',top:-c.top/h*100+'%',width:100000/w*100+'%',height:100000/h*100+'%',maxWidth:'none',display:'block',objectFit:'fill'});clip.append(img);element.append(clip);
      }else element.dataset.missingImage='true';
    } else if(shape.type==='text'){
      element.dataset.textFit=shape.autofit?.mode??'none';
      const frame=document.createElement('div');frame.className='slides-text-frame';frame.dataset.textFrame=shape.id;frame.dir=shape.direction;
      Object.assign(frame.style,{width:'100%',height:'100%',boxSizing:'border-box',padding:`${px(shape.margins.top)} ${px(shape.margins.right)} ${px(shape.margins.bottom)} ${px(shape.margins.left)}`,display:'flex',flexDirection:'column',justifyContent:shape.verticalAlignment==='center'?'center':shape.verticalAlignment==='bottom'?'flex-end':'flex-start',overflow:'visible'});
      const body=document.createElement('div');body.className='slides-text-body';Object.assign(body.style,{flex:'0 0 auto',minWidth:'0',whiteSpace:shape.wrap?'pre-wrap':'pre',overflowWrap:shape.wrap?'break-word':'normal',wordBreak:'normal',lineBreak:'strict'});applySlideTextStyle(body,fittedStyle(shape.textStyle,shape,fitOverride));
      for(const paragraph of shape.paragraphs){
        const p=document.createElement('p');p.dir=paragraph.direction;const style=fittedStyle(paragraph.style,shape,fitOverride),format=style.slideParagraph,indent=format?.firstLineIndent??style.indent??0,baseFont=style.fontSize??304800;
        applySlideTextStyle(p,style,paragraph.runs.map(r=>r.text).join(''));
        const reduction=shape.autofit?.mode==='normal'?shape.autofit.lineReduction/100000:0;
        Object.assign(p.style,{position:'relative',margin:'0',marginLeft:px(format?.marginLeft??0),marginRight:px(format?.marginRight??0),textAlign:style.alignment==='center'?'center':style.alignment==='end'?'end':style.alignment==='justify'?'justify':'start',textIndent:px(paragraph.marker?Math.max(0,indent):indent),paddingTop:px(format?.beforePercent!==undefined?baseFont*format.beforePercent/100000:style.spaceBefore??0),paddingBottom:px(format?.afterPercent!==undefined?baseFont*format.afterPercent/100000:style.spaceAfter??0),lineHeight:style.lineHeight?px(style.lineHeight):String(Math.max(.01,(style.lineMultiple!==undefined?style.lineMultiple/240:1.2)*(1-reduction))),minHeight:px(baseFont)});
        if(paragraph.marker){
          const mark=document.createElement('span');mark.className='slides-list-marker';mark.textContent=paragraph.marker.text;mark.dataset.marker=paragraph.marker.text;applySlideTextStyle(mark,fittedStyle(paragraph.marker.style,shape,fitOverride),mark.textContent);
          if(indent<0)Object.assign(mark.style,{position:'absolute',insetInlineStart:px(indent),textIndent:'0',whiteSpace:'nowrap',textAlign:'start'});else mark.style.marginInlineEnd='0.35em';p.append(mark);
        }
        for(const run of paragraph.runs){const span=document.createElement('span');span.dataset.runId=run.id;span.textContent=run.text;applySlideTextStyle(span,fittedStyle(run.style,shape,fitOverride),run.text);p.append(span);}
        if(!paragraph.runs.some(r=>r.text))p.append(document.createElement('br'));body.append(p);
      }
      frame.append(body);element.append(frame);
    }
  }
  updateSlideShapeGeometry(element,shape);
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


/** Explicit Fit Text command, using the same DOM projection as print/slideshow.
 * A bounded measurement never mutates the model or starts a background task. */
export function measureSlideTextFit(shape:TextScene):number{
 const probe=document.createElement('div');Object.assign(probe.style,{position:'fixed',left:'-100000px',top:'0',visibility:'hidden',pointerEvents:'none',width:px(shape.geometry.width),height:px(shape.geometry.height)});document.body.append(probe);
 const fits=(scale:number)=>{const node=createSlideShapeElement(shape,scale);node.style.left='0';node.style.top='0';node.style.transform='none';probe.replaceChildren(node);return overflowingSlideText(probe).length===0;};
 try{
  if(fits(100000))return 100000;
  if(!fits(10000))return 0;
  let low=10000,high=100000;
  for(let i=0;i<12;i++){const mid=Math.floor((low+high)/2);if(fits(mid))low=mid;else high=mid-1;}
  return low;
 }finally{probe.remove();}
}
