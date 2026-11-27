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
  const tint=color.tint??0,channel=(n:number)=>{
    if(!tint)return n;
    if(color.tintSpace==='linear'){const srgb=n/255,linear=srgb<=0.04045?srgb/12.92:((srgb+0.055)/1.055)**2.4;const adjusted=tint<0?linear*(1+tint):linear+(1-linear)*tint;const encoded=adjusted<=0.0031308?adjusted*12.92:1.055*adjusted**(1/2.4)-0.055;return Math.max(0,Math.min(255,Math.round(encoded*255)));}
    return Math.max(0,Math.min(255,Math.round(tint<0?n*(1+tint):n+(255-n)*tint)));
  };
  return {kind:'rgba',red:channel(base.red),green:channel(base.green),blue:channel(base.blue),alpha:base.alpha};
}

/** Resolve DrawingML script font aliases without fetching or installing fonts. */
const SCRIPT_TESTS:readonly [string,RegExp][]=[['Jpan',/[\p{Script=Hiragana}\p{Script=Katakana}]/u],['Hang',/\p{Script=Hangul}/u],['Arab',/\p{Script=Arabic}/u],['Hebr',/\p{Script=Hebrew}/u],['Deva',/\p{Script=Devanagari}/u],['Thai',/\p{Script=Thai}/u],['Beng',/\p{Script=Bengali}/u],['Gujr',/\p{Script=Gujarati}/u],['Guru',/\p{Script=Gurmukhi}/u],['Knda',/\p{Script=Kannada}/u],['Taml',/\p{Script=Tamil}/u],['Telu',/\p{Script=Telugu}/u],['Mlym',/\p{Script=Malayalam}/u],['Orya',/\p{Script=Oriya}/u],['Khmr',/\p{Script=Khmer}/u],['Laoo',/\p{Script=Lao}/u],['Sinh',/\p{Script=Sinhala}/u],['Tibt',/\p{Script=Tibetan}/u],['Syrc',/\p{Script=Syriac}/u],['Ethi',/\p{Script=Ethiopic}/u],['Mong',/\p{Script=Mongolian}/u]];
export function resolveSlideThemeFont(family:string|undefined,theme?:DeepReadonly<SlideTheme>,text='',language=''):string|undefined {
  const match=family&&/^\+(mj|mn)-(lt|ea|cs)$/.exec(family);if(!match)return family;
  const major=match[1]==='mj',latin=(major?theme?.majorFont:theme?.minorFont)??'sans-serif',details=major?theme?.majorFontDetails:theme?.minorFontDetails;
  if(match[2]==='lt')return latin;
  const explicit=match[2]==='ea'?details?.eastAsia:details?.complex;if(explicit)return explicit;
  let script=SCRIPT_TESTS.find(([key,test])=>(match[2]==='ea'?['Jpan','Hang'].includes(key):!['Jpan','Hang'].includes(key))&&test.test(text))?.[0];
  if(!script&&match[2]==='ea'&&/\p{Script=Han}/u.test(text))script=/^ja(?:-|$)/i.test(language)?'Jpan':/^ko(?:-|$)/i.test(language)?'Hang':/^zh-(?:Hant|TW|HK|MO)/i.test(language)?'Hant':'Hans';
  if(!script){const locale=language.toLowerCase();const fromLocale:Record<string,string>={ja:'Jpan',ko:'Hang',zh:'Hans',ar:'Arab',he:'Hebr',hi:'Deva',th:'Thai',bn:'Beng',gu:'Gujr',pa:'Guru',ta:'Taml',te:'Telu',ml:'Mlym',kn:'Knda',km:'Khmr',lo:'Laoo',si:'Sinh'};const code=fromLocale[locale.split('-')[0]!];if(code&&(match[2]==='ea'?['Jpan','Hang','Hans','Hant'].includes(code):!['Jpan','Hang','Hans','Hant'].includes(code)))script=code==='Hans'&&/^zh-(?:hant|tw|hk|mo)/.test(locale)?'Hant':code;}
  return script?details?.scripts?.[script]??latin:latin;
}
