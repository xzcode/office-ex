// SPDX-License-Identifier: MIT
import { recoveredCatalogs } from './recovered-catalogs.js';
import { m4Catalogs } from './m4-catalogs.js';
const english = { ...recoveredCatalogs['en-US'], ...m4Catalogs['en-US'] };
export { formatTemplate, formatPlural, formatSelect, pseudoLocalize } from './messages.js';
import { formatTemplate } from './messages.js';
export type MessageKey = keyof typeof english;
export type Catalog = Readonly<Record<MessageKey,string>>;
const catalogs = Object.fromEntries(Object.entries(recoveredCatalogs).map(([locale,catalog])=>[locale,{
  ...catalog,...m4Catalogs[locale as keyof typeof m4Catalogs],previewTag:catalog.previewTag.replace(/^M\d+\s*·\s*/,''),stage:'Writer · Sheets · Slides · PDF'
}])) as Record<keyof typeof recoveredCatalogs,Catalog>;
export type UiLocale = keyof typeof catalogs;
export type LanguagePreference = UiLocale | 'system';
export const TARGET_LOCALES = [
  ['zh-CN','简体中文','ltr','A'],['zh-TW','繁體中文','ltr','A'],['en-US','English','ltr','A'],['ja-JP','日本語','ltr','A'],
  ['ko-KR','한국어','ltr','A'],['fr-FR','Français','ltr','A'],['de-DE','Deutsch','ltr','A'],['es-ES','Español','ltr','A'],
  ['pt-BR','Português (Brasil)','ltr','A'],['ru-RU','Русский','ltr','A'],['it-IT','Italiano','ltr','A'],
  ['ar','العربية','rtl','B'],['he-IL','עברית','rtl','B'],['hi-IN','हिन्दी','ltr','B'],['th-TH','ไทย','ltr','B'],
  ['vi-VN','Tiếng Việt','ltr','B'],['id-ID','Bahasa Indonesia','ltr','B'],['ms-MY','Bahasa Melayu','ltr','B'],
  ['tr-TR','Türkçe','ltr','B'],['nl-NL','Nederlands','ltr','B'],['pl-PL','Polski','ltr','B'],['uk-UA','Українська','ltr','B']
] as const;
export type TargetLocale = typeof TARGET_LOCALES[number][0];
export const REGIONAL_LOCALES = [...TARGET_LOCALES.map(x=>x[0]),'en-GB','en-AU','es-419','pt-PT','zh-HK'] as const;
export type RegionalPreference = typeof REGIONAL_LOCALES[number] | 'system';
export type DocumentLanguagePreference = TargetLocale | 'system' | 'und';
export const availableLocales = (): readonly UiLocale[] => TARGET_LOCALES.map(x=>x[0]).filter(isUiLocale);
export function isUiLocale(value:unknown):value is UiLocale {return typeof value==='string'&&Object.hasOwn(catalogs,value);}
export function isLanguagePreference(value:unknown):value is LanguagePreference{return value==='system'||isUiLocale(value);}
export function isRegionalPreference(value:unknown):value is RegionalPreference{return value==='system'||typeof value==='string'&&REGIONAL_LOCALES.some(x=>x===value);}
export function isDocumentLanguagePreference(value:unknown):value is DocumentLanguagePreference{return value==='system'||value==='und'||typeof value==='string'&&TARGET_LOCALES.some(x=>x[0]===value);}
function canonical(value:string):string|undefined{try{return Intl.getCanonicalLocales(value)[0];}catch{return undefined;}}
export function resolveLocale(preference:LanguagePreference, system:readonly string[]):UiLocale {
  if(preference!=='system')return preference;
  for(const requested of system.slice(0,32)){
    const name=canonical(requested);if(!name)continue;
    if(isUiLocale(name))return name;
    const loc=new Intl.Locale(name);
    if(loc.language==='zh'){return loc.script==='Hant'||(!loc.script&&['TW','HK','MO'].includes(loc.region??''))?'zh-TW':'zh-CN';}
    if(loc.language==='en')return 'en-US';
    if(loc.language==='ja')return 'ja-JP';
    const match=availableLocales().find(x=>new Intl.Locale(x).language===loc.language);if(match)return match;
  }
  return 'en-US';
}
export function resolveRegion(preference:RegionalPreference, system:readonly string[]):string{
  if(preference!=='system')return preference;
  for(const requested of system.slice(0,32)){const c=canonical(requested);if(c&&Intl.NumberFormat.supportedLocalesOf([c]).length&&Intl.DateTimeFormat.supportedLocalesOf([c]).length)return c;}
  return 'en-US';
}
export function localeDirection(locale:string):'ltr'|'rtl'{const c=canonical(locale);if(!c)return 'ltr';return ['ar','he','fa','ur'].includes(new Intl.Locale(c).language)?'rtl':'ltr';}
export function isMessageKey(key:unknown):key is MessageKey{return typeof key==='string'&&Object.hasOwn(english,key);}
export function message(locale:UiLocale,key:MessageKey,args:Readonly<Record<string,string|number>>={}):string {
  if(!isMessageKey(key))throw new Error('Unknown message key');
  const value:string=isUiLocale(locale)?catalogs[locale][key]:english[key];
  return formatTemplate(localeDirection(locale)==='rtl'?value.replace(/\{[A-Za-z][A-Za-z0-9_]*\}/g,token=>'\u2068'+token+'\u2069'):value,args);
}
export function localeLabel(locale:string,ui?:UiLocale):string{
  const name=TARGET_LOCALES.find(x=>x[0]===locale)?.[1];
  if(!ui)return name??locale;
  try{const translated=new Intl.DisplayNames([ui],{type:'language',languageDisplay:'standard',fallback:'code'}).of(locale);return name&&translated&&name!==translated?name+' · '+translated:name??translated??locale;}catch{return name??locale;}
}
export function validateCatalogs():void{
  const required=Object.keys(english).sort().join('|');
  const placeholders=(s:string)=>[...s.matchAll(/\{([A-Za-z][A-Za-z0-9_]*)\}/g)].map(m=>m[1]).sort().join('|');
  for(const [locale,catalog] of Object.entries(catalogs)){
    if(Object.keys(catalog).sort().join('|')!==required)throw new Error('Catalog keys differ: '+locale);
    for(const key of Object.keys(english) as MessageKey[])if(!catalog[key].trim()||placeholders(catalog[key])!==placeholders(english[key]))throw new Error('Catalog value or parameters differ: '+locale+' '+key);
  }
}
const diagnosticKeys:Readonly<Record<string,MessageKey>>={
  'format.semanticPending':'diagSemanticPending','format.xlsxValuePreview':'xlsxPreviewReadOnly','format.officeTextPreview':'officeTextPreviewReadOnly','format.macrosDisabled':'diagMacrosDisabled','format.signedReadOnly':'diagSignedReadOnly',
  'format.externalDisabled':'diagExternalDisabled','format.strictReadOnly':'diagStrictReadOnly','format.encryptedUnsupported':'diagEncryptedUnsupported',
  'format.extensionMismatch':'diagExtensionMismatch','format.docxProfile':'docxNote','format.docProfile':'docNote','format.xlsxProfile':'xlsxNote','format.xlsProfile':'xlsNote'
};
export function diagnosticKey(key:string):MessageKey{return Object.hasOwn(diagnosticKeys,key)?diagnosticKeys[key]!:'diagUnknown';}
const errorKeys:Readonly<Record<string,MessageKey>>={invalid:'errorInvalid',unsupported:'errorUnsupported',unsafe:'errorUnsafe',limit:'errorLimit',conflict:'errorConflict',cancelled:'cancelled',io:'errorIo',internal:'errorInternal'};
export function errorKey(code:string):MessageKey{return Object.hasOwn(errorKeys,code)?errorKeys[code]!:'errorInternal';}
export function formatNumber(locale:string,value:number,options?:Intl.NumberFormatOptions):string{return new Intl.NumberFormat(locale,options).format(value);}
export function formatDate(locale:string,value:Date|number,options?:Intl.DateTimeFormatOptions):string{return new Intl.DateTimeFormat(locale,options??{dateStyle:'medium',timeStyle:'short'}).format(value);}
