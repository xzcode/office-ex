// SPDX-License-Identifier: MIT
/** Canonical presentation theme defaults used by scene and file adapters. */
import {assert} from '@office-ex/contracts';
import {freezeData,type Color,type SlideTheme} from './model.js';
import type {DeepReadonly} from './session.js';
type Rgba=Extract<Color,{kind:'rgba'}>;
export const DEFAULT_SLIDE_THEME_COLORS:Readonly<Record<string,Rgba>>=freezeData({
  dk1:{kind:'rgba',red:24,green:32,blue:48,alpha:255},lt1:{kind:'rgba',red:255,green:255,blue:255,alpha:255},dk2:{kind:'rgba',red:48,green:58,blue:80,alpha:255},lt2:{kind:'rgba',red:240,green:244,blue:250,alpha:255},
  accent1:{kind:'rgba',red:66,green:111,blue:202,alpha:255},accent2:{kind:'rgba',red:24,green:140,blue:123,alpha:255},accent3:{kind:'rgba',red:239,green:143,blue:54,alpha:255},accent4:{kind:'rgba',red:139,green:91,blue:182,alpha:255},accent5:{kind:'rgba',red:211,green:75,blue:105,alpha:255},accent6:{kind:'rgba',red:73,green:158,blue:193,alpha:255},hlink:{kind:'rgba',red:31,green:83,blue:196,alpha:255},folHlink:{kind:'rgba',red:123,green:63,blue:151,alpha:255}
});
const aliases:Readonly<Record<string,string>>={text1:'dk1',text2:'dk2',background1:'lt1',background2:'lt2',tx1:'dk1',tx2:'dk2',bg1:'lt1',bg2:'lt2'};
export function resolveSlideThemeColor(color:DeepReadonly<Color>,theme?:DeepReadonly<SlideTheme>):Rgba {
  if(color.kind==='rgba')return {...color};const name=aliases[color.name]??color.name,base=theme?.colors[name]??DEFAULT_SLIDE_THEME_COLORS[name];assert(base,'unsupported','Unknown presentation theme color');
  const tint=color.tint??0,channel=(n:number)=>Math.max(0,Math.min(255,Math.round(tint<0?n*(1+tint):n+(255-n)*tint)));
  return {kind:'rgba',red:channel(base.red),green:channel(base.green),blue:channel(base.blue),alpha:base.alpha};
}

export function resolveSlideThemeFont(family:string|undefined,theme?:DeepReadonly<SlideTheme>):string|undefined {return family==='+mj-lt'?theme?.majorFont??'sans-serif':family==='+mn-lt'?theme?.minorFont??'sans-serif':family;}
