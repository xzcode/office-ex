// SPDX-License-Identifier: MIT
/** Closed editable numbering subset. Unknown unused definitions stay in the
 * source XML; any active reference outside this subset is protected. */
import type {Element} from '@xmldom/xmldom';
import {OfficeError,EMU_PER_TWIP} from '@office-ex/contracts';
import type {WriterNumbering,WriterListLevel,WriterListReference,DeepReadonly} from '@office-ex/core';
import {parseXml,serializeXml} from './xml.js';
import type {ContainerOptions} from './limits.js';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',PKG='http://schemas.openxmlformats.org/package/2006/relationships',CT='http://schemas.openxmlformats.org/package/2006/content-types',XMLNS='http://www.w3.org/2000/xmlns/',MC='http://schemas.openxmlformats.org/markup-compatibility/2006';
const MIME='application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
function valid(condition:unknown,message:string):asserts condition{if(!condition)throw new OfficeError('unsupported','DOCX numbering: '+message);}
const attr=(e:Element,k:string)=>e.hasAttributeNS(W,k)?e.getAttributeNS(W,k)!:undefined;
function children(e:Element):Element[]{const out:Element[]=[];for(let n=e.firstChild;n;n=n.nextSibling){if(n.nodeType===1)out.push(n as Element);else valid(n.nodeType===3&&!(n.nodeValue??'').trim()||n.nodeType===8,'Unexpected element content');}return out;}
function attrs(e:Element,keys:readonly string[]=[],root=false){for(let i=0;i<e.attributes.length;i++){const a=e.attributes.item(i)!;if(a.namespaceURI===XMLNS)continue;if(root&&a.namespaceURI===MC&&a.localName==='Ignorable')continue;valid(a.namespaceURI===W&&keys.includes(a.localName??''),'Unsupported '+e.localName+' attribute');}}
function nodes(e:Element,keys:readonly string[]):Map<string,Element>{const out=new Map<string,Element>();for(const c of children(e)){const key=c.localName??'';valid(c.namespaceURI===W&&keys.includes(key)&&!out.has(key),'Unknown or repeated '+key);out.set(key,c);}return out;}
function number(e:Element,key:string,min=0,max=1000000000):number{const raw=attr(e,key);valid(raw&&/^\d+$/.test(raw),'Invalid '+key);const n=Number(raw);valid(Number.isSafeInteger(n)&&n>=min&&n<=max,'Invalid '+key);return n;}
function value(e:Element|undefined):string{valid(e,'Missing numbering property');attrs(e,['val']);valid(!children(e).length,'Unexpected property children');const v=attr(e,'val');valid(v!==undefined,'Missing property value');return v;}
function intValue(e:Element|undefined,min:number,max:number):number{valid(e,'Missing numeric property');value(e);return number(e,'val',min,max);}
export function readNumberingReference(e:Element):WriterListReference{
 attrs(e);const c=nodes(e,['ilvl','numId']),numId=intValue(c.get('numId'),0,1000000000),level=c.has('ilvl')?intValue(c.get('ilvl'),0,8):0;
 valid(numId!==0||level===0,'Disabled numbering level');return {numId,level};
}
function readLevel(e:Element,index:number,single:boolean):WriterListLevel{
 attrs(e,['ilvl','tentative']);valid(attr(e,'ilvl')===undefined?single&&index===0:number(e,'ilvl',0,8)===index,'Noncontiguous numbering levels');
 const c=nodes(e,['start','numFmt','lvlText','lvlJc','pPr','rPr','pStyle','suff','lvlRestart']),format=value(c.get('numFmt'));
 valid(format==='decimal'||format==='bullet','Unsupported list format');
 const start=!c.has('start')&&format==='bullet'?1:intValue(c.get('start'),1,32767),alignment=c.has('lvlJc')?value(c.get('lvlJc')):'left';
 valid(['left','center','right'].includes(alignment),'Unsupported marker alignment');
 valid(!c.has('suff')||value(c.get('suff'))==='tab','Only tab suffix supported');
 // Style-associated numbering is supported for the single level default
 // templates. Multi-level style-linked numbering needs a separate resolver.
 if(c.has('pStyle')){valid(index===0,'Higher level style association');value(c.get('pStyle'));}
 const restart=c.has('lvlRestart')?intValue(c.get('lvlRestart'),0,8):undefined;
 valid(restart===undefined||restart===0||restart===index,'Unsupported list restart trigger');
 let pattern=value(c.get('lvlText')),symbolBullet=false;
 const rpr=c.get('rPr');
 const paint:Pick<WriterListLevel,'markerColor'|'markerFontSize'>={};
 if(rpr){
  attrs(rpr);const props=nodes(rpr,['rFonts','color','sz']);const fonts=props.get('rFonts');
  if(fonts){valid(format==='bullet'&&pattern==='\uf0b7','Unsupported numbering font formatting');attrs(fonts,['ascii','hAnsi','hint']);valid(!children(fonts).length&&attr(fonts,'ascii')==='Symbol'&&attr(fonts,'hAnsi')==='Symbol'&&(!attr(fonts,'hint')||attr(fonts,'hint')==='default'),'Unsupported symbol font');pattern='•';symbolBullet=true;}
  if(props.has('color')){const hex=value(props.get('color'));valid(/^[0-9a-fA-F]{6}$/.test(hex),'Only explicit RGB numbering color supported');paint.markerColor={kind:'rgba',red:parseInt(hex.slice(0,2),16),green:parseInt(hex.slice(2,4),16),blue:parseInt(hex.slice(4,6),16),alpha:255};}
  if(props.has('sz'))paint.markerFontSize=intValue(props.get('sz'),1,3276)*6350;
 }
 if(format==='bullet')valid(Array.from(pattern).length===1&&!/[\s%\x00-\x1f\x7f\uE000-\uF8FF]/u.test(pattern),'Unsupported bullet character');
 else valid(pattern.length<=64&&!/[\x00-\x1f\x7f]/.test(pattern)&&/^(?:[^%]|%[1-9])+$/.test(pattern)&&pattern.includes('%'+(index+1))&&[...pattern.matchAll(/%([1-9])/g)].every(m=>Number(m[1])<=index+1),'Unsupported decimal pattern');
 const p=c.get('pPr');let left=360,hanging=180;
 if(p){attrs(p);const pc=nodes(p,['ind','tabs']),ind=pc.get('ind');valid(ind,'Missing list indentation');
  attrs(ind,['left','hanging']);valid(!children(ind).length,'Nested indentation');left=number(ind,'left',1,31680);hanging=number(ind,'hanging',1,left);
  const tabs=pc.get('tabs');if(tabs){attrs(tabs);const tab=nodes(tabs,['tab']).get('tab');valid(tab,'Missing list tab');attrs(tab,['val','pos']);valid(!children(tab).length&&['num','left'].includes(attr(tab,'val')??'')&&number(tab,'pos',1,31680)===left,'Unsupported list tab stop');}
 }else valid(single&&format==='bullet','Missing list paragraph geometry');
 // For the simple producer bullet profile, the preview fallback is 18pt / 9pt.
 // Existing definitions are immutable and the source numbering XML is preserved.
 return {format,start,pattern,indent:left*EMU_PER_TWIP,hanging:hanging*EMU_PER_TWIP,alignment:alignment==='left'?'start':alignment==='right'?'end':'center',...(restart===0&&index>0?{neverRestart:true}:{}),...(symbolBullet?{symbolBullet:true}:{}),...paint};
}
export interface NumberingSource {path:string;numbering:WriterNumbering;maxAbstractId:number}
export function readNumbering(path:string,bytes:Uint8Array,options:ContainerOptions):NumberingSource{
 const doc=parseXml(bytes,options),root=doc.documentElement!;valid(root.namespaceURI===W&&root.localName==='numbering','Unexpected numbering root');attrs(root,[],true);
 const abstracts=new Map<number,WriterListLevel[]>(),seenAbstract=new Set<number>(),seenNum=new Set<number>(),nums:Element[]=[];
 let maxAbstractId=-1,maxId=0;
 for(const e of children(root)){
  valid(e.namespaceURI===W,'Unexpected numbering namespace');
  if(e.localName==='abstractNum'){
   const id=number(e,'abstractNumId');valid(!seenAbstract.has(id)&&seenAbstract.size<2048,'Duplicate or excessive abstract numbering');seenAbstract.add(id);maxAbstractId=Math.max(maxAbstractId,id);
   try{
    attrs(e,['abstractNumId']);const levelNodes:Element[]=[],metadata=new Set<string>();
    for(const child of children(e)){
     valid(child.namespaceURI===W,'Unexpected abstract child');const key=child.localName??'';
     if(key==='lvl')levelNodes.push(child);
     else{valid(['nsid','multiLevelType','tmpl','name'].includes(key)&&!metadata.has(key),'Unsupported abstract property');metadata.add(key);const v=value(child);if(key==='multiLevelType')valid(['singleLevel','multilevel','hybridMultilevel'].includes(v),'Unsupported numbering kind');}
    }
    valid(levelNodes.length>0&&levelNodes.length<=9,'Invalid abstract levels');if(levelNodes.length>1)levelNodes.sort((a,b)=>number(a,'ilvl',0,8)-number(b,'ilvl',0,8));
    const levels=levelNodes.map((e,i)=>readLevel(e,i,levelNodes.length===1));
    // Decimal composite patterns cannot refer to bullet ancestors.
    for(const l of levels)if(l.format==='decimal')valid([...l.pattern.matchAll(/%([1-9])/g)].every(m=>levels[Number(m[1])-1]?.format==='decimal'),'Mixed composite numbering unsupported');
    abstracts.set(id,levels);
   }catch(error){if(!(error instanceof OfficeError)||error.code!=='unsupported')throw error;}
  }else if(e.localName==='num'){const id=number(e,'numId',1);valid(!seenNum.has(id)&&seenNum.size<2048,'Duplicate or excessive numbering instances');seenNum.add(id);maxId=Math.max(maxId,id);nums.push(e);}
  else valid(['numPicBullet','numIdMacAtCleanup'].includes(e.localName??''),'Unknown numbering root child');
 }
 const lists:WriterNumbering['lists']=Object.create(null);
 for(const e of nums){
  try{
   attrs(e,['numId']);const id=number(e,'numId',1),parts=children(e),abstract=parts.filter(c=>c.namespaceURI===W&&c.localName==='abstractNumId');valid(abstract.length===1,'Invalid abstract reference');
   const base=abstracts.get(intValue(abstract[0],0,1000000000));valid(base,'Unsupported abstract definition');
   const levels=base.map(l=>({...l})),overrides=new Set<number>();
   for(const c of parts){if(c===abstract[0])continue;valid(c.namespaceURI===W&&c.localName==='lvlOverride','Unsupported instance property');attrs(c,['ilvl']);const level=number(c,'ilvl',0,levels.length-1);valid(!overrides.has(level),'Duplicate start override');overrides.add(level);const start=nodes(c,['startOverride']).get('startOverride');levels[level]!.start=intValue(start,1,32767);}
   lists[String(id)]={levels};
  }catch(error){if(!(error instanceof OfficeError)||error.code!=='unsupported')throw error;}
 }
 return {path,numbering:{nextId:maxId+1,lists},maxAbstractId};
}
export function supportsNumberingReference(ref:WriterListReference|undefined,numbering:WriterNumbering|undefined):boolean{return !ref?.numId||!!numbering?.lists[String(ref.numId)]?.levels[ref.level];}
const canonical=(v:unknown):string=>v===null||typeof v!=='object'?JSON.stringify(v):Array.isArray(v)?'['+v.map(canonical).join(',')+']':'{'+Object.entries(v).sort(([a],[b])=>a.localeCompare(b)).map(([k,v])=>JSON.stringify(k)+':'+canonical(v)).join(',')+'}';
export function numberingChanged(source:NumberingSource|undefined,numbering:DeepReadonly<WriterNumbering>|undefined):boolean{
 if(source){valid(numbering,'Source numbering registry must be preserved');for(const [id,def] of Object.entries(source.numbering.lists))valid(canonical(numbering.lists[id])===canonical(def),'Existing numbering definition is immutable');valid(numbering.nextId>=source.numbering.nextId,'Numbering allocation regressed');}
 const additions=Object.keys(numbering?.lists??{}).filter(id=>!source?.numbering.lists[id]);
 for(const id of additions)valid(Number(id)>=(source?.numbering.nextId??1),'Numbering identity collides with opaque source definition');
 return additions.length>0;
}
export function addNumberingParts(parts:Map<string,Uint8Array>,main:string,source:NumberingSource|undefined,numbering:DeepReadonly<WriterNumbering>|undefined,options:ContainerOptions):void{
 if(!numberingChanged(source,numbering))return;
 const path=source?.path??(()=>{let p='word/numbering.xml',i=1;const used=new Set([...parts.keys()].map(k=>k.toLowerCase()));while(used.has(p.toLowerCase()))p='word/office-ex-numbering-'+i+++'.xml';return p;})();
 const doc=parseXml(parts.get(path)??'<w:numbering xmlns:w="'+W+'"/>',options),root=doc.documentElement!;
 const node=(name:string,attributes:Record<string,string|number>={},children:Element[]=[]):Element=>{const e=doc.createElementNS(W,'w:'+name);for(const [key,v]of Object.entries(attributes))e.setAttributeNS(W,'w:'+key,String(v));for(const c of children)e.appendChild(c);return e;};
 const property=(name:string,val:string|number)=>node(name,{val});
 let abstractId=(source?.maxAbstractId??-1)+1;
 for(const [id,definition]of Object.entries(numbering!.lists)){
  if(source?.numbering.lists[id])continue;
  valid(abstractId<=1000000000,'Abstract numbering allocation exhausted');
  const levels=definition.levels.map((l,i)=>{
   valid(l.indent%EMU_PER_TWIP===0&&l.hanging%EMU_PER_TWIP===0,'List indentation precision');
   const elements=[property('start',l.start),property('numFmt',l.format)];
   if(l.neverRestart)elements.push(property('lvlRestart',0));
   elements.push(property('suff','tab'),property('lvlText',l.symbolBullet?'\uf0b7':l.pattern),property('lvlJc',l.alignment==='start'?'left':l.alignment==='end'?'right':'center'),node('pPr',{},[node('tabs',{},[node('tab',{val:'num',pos:l.indent/EMU_PER_TWIP})]),node('ind',{left:l.indent/EMU_PER_TWIP,hanging:l.hanging/EMU_PER_TWIP})]));
   const rPr:Element[]=[];
   if(l.symbolBullet)rPr.push(node('rFonts',{ascii:'Symbol',hAnsi:'Symbol',hint:'default'}));
   if(l.markerColor){const c=l.markerColor;valid(c.kind==='rgba'&&c.alpha===255,'Marker color requires opaque RGB');rPr.push(property('color',[c.red,c.green,c.blue].map(n=>n.toString(16).padStart(2,'0')).join('').toUpperCase()));}
   if(l.markerFontSize!==undefined){valid(l.markerFontSize%6350===0,'Marker size requires half points');rPr.push(property('sz',l.markerFontSize/6350));}
   if(rPr.length)elements.push(node('rPr',{},rPr));
   return node('lvl',{ilvl:i},elements);
  });
  const abstract=node('abstractNum',{abstractNumId:abstractId},[property('multiLevelType',levels.length===1?'singleLevel':'multilevel'),...levels]);
  const firstInstance=children(root).find(e=>e.localName==='num'||e.localName==='numIdMacAtCleanup');
  root.insertBefore(abstract,firstInstance??null);
  const instance=node('num',{numId:id},[property('abstractNumId',abstractId++)]);root.insertBefore(instance,children(root).find(e=>e.localName==='numIdMacAtCleanup')??null);
 }
 parts.set(path,serializeXml(doc,options));
 if(!source){
  const i=main.lastIndexOf('/'),expected=main.slice(0,i+1)+'_rels/'+main.slice(i+1)+'.rels',actual=[...parts.keys()].find(p=>p.toLowerCase()===expected.toLowerCase())??expected;
  const rels=parseXml(parts.get(actual)??'<Relationships xmlns="'+PKG+'"/>',options),relRoot=rels.documentElement!,ids=new Set(children(relRoot).map(e=>e.getAttribute('Id')));let j=1;while(ids.has('rIdOexNumbering'+j))j++;
  const rel=rels.createElementNS(PKG,'Relationship');rel.setAttribute('Id','rIdOexNumbering'+j);rel.setAttribute('Type',R+'/numbering');rel.setAttribute('Target','/'+path);relRoot.appendChild(rel);parts.set(actual,serializeXml(rels,options));
  const types=parseXml(parts.get('[Content_Types].xml')!,options),override=types.createElementNS(CT,'Override');override.setAttribute('PartName','/'+path);override.setAttribute('ContentType',MIME);types.documentElement!.appendChild(override);parts.set('[Content_Types].xml',serializeXml(types,options));
 }
}
