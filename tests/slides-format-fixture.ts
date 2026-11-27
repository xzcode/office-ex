import {createDocument,makeImageResource,type SlidesDocument,type Color} from '@office-ex/core';
import {PNG,JPEG} from './writer-image-fixture.js';
export function slidesFormatFixture():SlidesDocument{
 const d=createDocument('slides','slides-fixture','Three slide scene');d.defaultTextStyle={fontFamily:'Arial',fontSize:228600,language:'zh-CN'};
 const png=makeImageResource(PNG),jpeg=makeImageResource(JPEG);d.resources={[png.id]:png,[jpeg.id]:jpeg};
 const colors:Color[]=[{kind:'rgba',red:244,green:247,blue:251,alpha:255},{kind:'rgba',red:235,green:246,blue:238,alpha:255},{kind:'rgba',red:255,green:242,blue:229,alpha:255}];
 d.slides=colors.map((background,i)=>({id:'page-'+i,background,shapes:[
  {id:'rect-'+i,type:'rect',geometry:{x:914400,y:914400,width:2743200,height:1371600,rotation:15*60000},fill:{kind:'rgba',red:40+i*50,green:100,blue:180,alpha:255},stroke:{width:12700,color:{kind:'rgba',red:20,green:30,blue:40,alpha:255}}},
  {id:'ellipse-'+i,type:'ellipse',geometry:{x:3657600,y:914400,width:1828800,height:1371600,rotation:0},fill:{kind:'rgba',red:200,green:80+i*30,blue:70,alpha:255}},
  {id:'line-'+i,type:'line',geometry:{x:914400,y:2743200,width:4572000,height:914400,rotation:0},stroke:{width:25400,color:{kind:'rgba',red:30,green:40,blue:70,alpha:255}}},
  {id:'text-'+i,type:'text',geometry:{x:914400,y:4114800,width:5486400,height:1828800,rotation:0},textStyle:{alignment:'center',paragraphDirection:i===1?'rtl':'ltr'},textFrame:{wrap:true,verticalAlignment:'center',margins:{left:91440,right:91440,top:45720,bottom:45720}},runs:[{id:'run-'+i,type:'text',text:['第一张 — Hello 🙂\nEditable text','مرحبا بالعالم — שלום','Unicode देवनागरी ไทย Việt'][i]!,style:{bold:true,fontSize:254000,color:{kind:'rgba',red:60,green:40,blue:100,alpha:255},language:i===1?'ar-SA':'zh-CN'}}]},
  {id:'image-'+i,type:'image',resourceId:i===1?jpeg.id:png.id,geometry:{x:7315200,y:1828800,width:2743200,height:2057400,rotation:0}}
 ]}));return d;
}
