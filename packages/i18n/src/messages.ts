// SPDX-License-Identifier: MIT
/** A deliberately small structured message API; no runtime evaluation of translation code.
 * Not a general ICU MessageFormat parser. New messages use typed branches and Intl rules.
 */
export type MessageArgs = Readonly<Record<string,string|number>>;
const PARAM = /\{([A-Za-z][A-Za-z0-9_]*)\}/g;
export function formatTemplate(template:string,args:MessageArgs={}):string {
  if(template.length>16384)throw new Error('Message template too large');
  const output=template.replace(PARAM,(_match:string,name:string)=>{
    const descriptor=Object.getOwnPropertyDescriptor(args,name);
    if(!descriptor||!Object.hasOwn(descriptor,'value'))throw new Error('Missing or unsafe message parameter: '+name);
    const value:unknown=descriptor.value;
    if(typeof value!=='string'&&(typeof value!=='number'||!Number.isFinite(value)))throw new Error('Invalid message parameter: '+name);
    if(typeof value==='string'&&value.length>8192)throw new Error('Message parameter too large');
    return String(value);
  });
  if(output.length>65536)throw new Error('Formatted message too large');
  return output;
}
type PluralCategory=Intl.LDMLPluralRule;
type PluralForms=Readonly<{other:string}&Partial<Record<Exclude<PluralCategory,'other'>,string>>>;
/** Caller supplies reviewed localized branches. The number shown can use an independent region. */
export function formatPlural(locale:string,count:number,forms:PluralForms,args:MessageArgs={},region=locale):string{
  if(!Number.isFinite(count))throw new Error('Invalid plural count');
  const category=new Intl.PluralRules(locale).select(count);
  const branch=Object.hasOwn(forms,category)?forms[category]:forms.other;
  if(typeof branch!=='string'||typeof forms.other!=='string')throw new Error('Invalid plural branches');
  return formatTemplate(branch,{...args,count:new Intl.NumberFormat(region).format(count)});
}
export function formatSelect(value:string,forms:Readonly<Record<string,string>&{other:string}>,args:MessageArgs={}):string{
  const branch=Object.hasOwn(forms,value)?forms[value]:forms.other;
  if(typeof branch!=='string'||typeof forms.other!=='string')throw new Error('Invalid select branches');
  return formatTemplate(branch,args);
}
/** Development-only stress fixture: preserve parameter tokens, expand visible text, no document mutation. */
export function pseudoLocalize(template:string,direction:'ltr'|'rtl'='ltr'):string{
  const accents:Readonly<Record<string,string>>={a:'á',e:'ë',i:'ï',o:'ö',u:'ü',A:'Á',E:'Ë',I:'Ï',O:'Ö',U:'Ü'};
  let visible=0;
  const transformed=template.split(/(\{[A-Za-z][A-Za-z0-9_]*\})/g).map(part=>{
    if(/^\{[A-Za-z][A-Za-z0-9_]*\}$/.test(part))return part;
    visible+=Array.from(part).length;return part.replace(/[aeiouAEIOU]/g,c=>accents[c]??c);
  }).join('');
  const label='⟦'+transformed+' '+ '~'.repeat(Math.ceil(visible*0.4))+'⟧';
  return direction==='rtl'?'\u2067'+label+'\u2069':label;
}
