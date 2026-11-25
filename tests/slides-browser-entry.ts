// SPDX-License-Identifier: MIT
/** Test-only browser entry; it is never included in the desktop product. */
import {createDocument,DocumentSession,makeImageResource,decodeImageBytes,type SlidesDocument,type DocumentSnapshot} from '@office-ex/core';
import {DEFAULT_SLIDE_THEME,type IdFactory} from '@office-ex/slides';
import {message,localeDirection,type UiLocale} from '@office-ex/i18n';
import {slidesWorkspace,type SlidesViewState} from '../packages/ui/src/slides-workspace.js';
const PNG=decodeImageBytes('iVBORw0KGgoAAAANSUhEUgAAAAgAAAAGCAIAAABxZ0isAAAAHklEQVR4nGP8z8Cgl7mKAQMw6mITZWBgYMIqSl0JAOavA4uzO/5KAAAAAElFTkSuQmCC');
let imageGate: (()=>void)|undefined;let imageDelayed=false;
let locale:UiLocale='en-US',session:DocumentSession,state:SlidesViewState={selectedIds:[],zoom:0},composing=false,rejected=false;const errors:string[]=[];
function fixture(count=3):SlidesDocument {
  const doc=createDocument('slides','slides-browser');doc.defaultTextStyle={fontFamily:'sans-serif',fontSize:304800,language:'zh-CN'};doc.theme=structuredClone(DEFAULT_SLIDE_THEME);const image=makeImageResource(PNG);doc.resources[image.id]=image;
  doc.slides=Array.from({length:count},(_,i)=>({id:'slide-'+i,background:{kind:'rgba',red:255-i%3*10,green:255-i%3*6,blue:255,alpha:255},shapes:[
    {id:'rectangle-'+i,type:'rect',geometry:{x:1000000,y:1000000,width:3000000,height:1500000,rotation:0},fill:{kind:'theme',name:'accent1'},stroke:{color:{kind:'theme',name:'dk1'},width:12700}},
    {id:'text-'+i,type:'text',geometry:{x:5000000,y:1000000,width:5000000,height:1800000,rotation:0},textStyle:{fontSize:304800},textFrame:{wrap:true,direction:'auto'},runs:[{id:'run-'+i,type:'text',text:'Office EX '+(i+1)+'\n中文 العربية עברית हिन्दी ไทย',style:{language:'zh-CN'}}]},
    {id:'ellipse-'+i,type:'ellipse',geometry:{x:1000000,y:3800000,width:2200000,height:1700000,rotation:2700000},fill:{kind:'rgba',red:210,green:89,blue:83,alpha:255}},
    {id:'image-'+i,type:'image',geometry:{x:5200000,y:3800000,width:1600000,height:1200000,rotation:0},resourceId:image.id},
    {id:'line-'+i,type:'line',geometry:{x:8000000,y:4000000,width:2000000,height:1000000,rotation:0},stroke:{color:{kind:'theme',name:'accent2'},width:25400}}
  ]}));return doc;
}
function render():void {state.dispose?.();document.documentElement.lang=locale;document.documentElement.dir=localeDirection(locale);document.querySelector('#test-root')!.replaceChildren(slidesWorkspace({snapshot:()=>session.snapshot(),execute:command=>{if(rejected)return false;try{session.execute(command);return true;}catch(error){errors.push(String(error));return false;}},render,error:error=>errors.push(String(error)),composing:()=>composing,setComposing:value=>{composing=value;},t:(key,args)=>message(locale,key,args),number:value=>new Intl.NumberFormat(locale).format(value),prepareImage:async bytes=>{if(imageDelayed){await new Promise<void>(resolve=>{imageGate=resolve;});imageGate=undefined;}return makeImageResource(bytes);},isCurrent:()=>true},state));}
function load(doc:SlidesDocument):void {state.dispose?.();session=new DocumentSession(doc,{saved:true});state={selectedIds:[],zoom:0};render();}
const api={snapshot:()=>session.snapshot(),history:()=>session.historyUsage,load,fixture,undo:()=>{if(state.flush?.()!==false){session.undo();render();}},redo:()=>{if(state.flush?.()!==false){session.redo();render();}},setLocale:(value:UiLocale)=>{if(state.flush?.()!==false){locale=value;render();}},flush:()=>state.flush?.(),select:(ids:string[])=>{if(state.flush?.()!==false){state.selectedIds=ids;render();}},errors,delayImage:(value:boolean)=>{imageDelayed=value;},imagePending:()=>Boolean(imageGate),releaseImage:()=>imageGate?.(),selected:()=>state.selectedIds,reject:(value:boolean)=>{rejected=value;},setText:(text:string)=>{const input=document.querySelector<HTMLTextAreaElement>('#slide-text-editor')!;input.value=text;input.dispatchEvent(new InputEvent('input',{bubbles:true}));},state:()=>({slideId:state.slideId,zoom:state.zoom,selectedIds:state.selectedIds})};
declare global {interface Window {testSlides:typeof api}}
window.testSlides=api;load(fixture());
