// SPDX-License-Identifier: MIT
/** Source-owned Word style/theme subset. The original ancillary parts are never
 * regenerated. Unknown unused styles may be kept opaque; any reference to them
 * (including a default or basedOn ancestor) prevents semantic editing.
 */
import type {Element} from '@xmldom/xmldom';
import {OfficeError,assertId} from '@office-ex/contracts';
import type {NamedStyle,TextStyle} from '@office-ex/core';
import type {OpcPackage,OpcRelationship} from './opc.js';
import {parseXml} from './xml.js';
import {verifySettings,verifyAuxiliaryXml} from './writer-docx-auxiliary.js';
import type {ContainerOptions} from './limits.js';
import {readNumbering,supportsNumberingReference,type NumberingSource} from './writer-docx-numbering.js';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const NS='http://www.w3.org/2000/xmlns/';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const STYLES='application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const THEME='application/vnd.openxmlformats-officedocument.theme+xml';
const CORE='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const REL_TYPE='application/vnd.openxmlformats-package.relationships+xml';
type StyleKind='paragraph'|'character';
export interface ThemeProfile {colors:Readonly<Record<string,string>>;fonts:Readonly<Record<string,string>>}
export interface DocxStyleProfile{
  styles:Record<string,NamedStyle>;
  defaults:TextStyle;
  defaultParagraph?:string;
  tableDefaultsSafe:boolean;
  builtinTableStyles?:ReadonlySet<'TableNormal'|'TableGrid'>;
  builtinStoryStyles?:ReadonlySet<'Header'|'Footer'>;
  numbering?:NumberingSource;
}
type RunReader=(el:Element,reference?:((id:string)=>void),theme?:ThemeProfile,named?:boolean,toggles?:('bold'|'italic')[])=>TextStyle;
type ParagraphReader=(el:Element,reference?:((id:string)=>void),sectionEnd?:Element,theme?:ThemeProfile)=>TextStyle;
function allowed(value:unknown,reason:string):asserts value{if(!value)throw new OfficeError('unsupported','DOCX style profile: '+reason);}
function name(el:Element,ns:string,n:string):boolean{return el.namespaceURI===ns&&el.localName===n;}
function children(el:Element):Element[]{
  const result:Element[]=[];
  for(let n=el.firstChild;n;n=n.nextSibling){
    if(n.nodeType===1)result.push(n as Element);
    else allowed(n.nodeType===3&&!(n.nodeValue??'').trim()||n.nodeType===8,'Unexpected non-element content');
  }
  return result;
}
function attribute(el:Element,n:string):string|undefined{return el.hasAttributeNS(W,n)?el.getAttributeNS(W,n)??undefined:undefined;}
function attrs(el:Element,ns:string,keys:readonly string[]):void{
  for(let i=0;i<el.attributes.length;i++){
    const a=el.attributes.item(i)!;
    if(a.namespaceURI===NS)continue;
    if(a.namespaceURI==='http://schemas.openxmlformats.org/markup-compatibility/2006'&&a.localName==='Ignorable'&&el.localName==='styles'&&a.value.split(/\s+/).every(prefix=>prefix&&el.lookupNamespaceURI(prefix)?.startsWith('http://schemas.microsoft.com/office/')))continue;
    allowed((a.namespaceURI??'')===ns&&a.localName!==null&&keys.includes(a.localName),'Unexpected attribute on '+el.localName);
  }
}
function value(el:Element):string{
  attrs(el,W,['val']);allowed(!children(el).length,'Unexpected nested style value');
  const v=attribute(el,'val');allowed(v!==undefined&&v.length>0,'Missing style value');return v;
}
function flag(value:string|undefined):boolean{
  if(value===undefined)return true;
  allowed(['0','1','true','false','on','off'].includes(value),'Invalid style flag');
  return ['1','true','on'].includes(value);
}
function stableId(id:string):boolean{return /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,127}$/.test(id)&&!['__proto__','constructor','prototype'].includes(id);}
function only(el:Element,ns:string,n:string):Element{
  const matches=children(el).filter(c=>name(c,ns,n));allowed(matches.length===1,'Expected unique '+n);return matches[0]!;
}
const colorNames=['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'] as const;
function readTheme(bytes:Uint8Array,options:ContainerOptions):ThemeProfile{
  const root=parseXml(bytes,options).documentElement!;
  allowed(name(root,A,'theme'),'Expected DrawingML theme');
  const elements=only(root,A,'themeElements'),scheme=only(elements,A,'clrScheme'),fonts=only(elements,A,'fontScheme');
  const colors:Record<string,string>=Object.create(null) as Record<string,string>;
  for(const n of colorNames){
    const entry=only(scheme,A,n),parts=children(entry);allowed(parts.length===1,'Unsupported theme color');
    const c=parts[0]!;allowed(!children(c).length,'Theme color transforms not supported');
    let hex:string|null=null;
    if(name(c,A,'srgbClr')){attrs(c,'',['val']);hex=c.getAttribute('val');}
    else if(name(c,A,'sysClr')){attrs(c,'',['val','lastClr']);hex=c.getAttribute('lastClr');}
    allowed(hex&&/^[0-9a-fA-F]{6}$/.test(hex),'Unsupported theme RGB or system-color fallback');
    colors[n]=hex.toUpperCase();
  }
  // No settings/color remapping part is admitted by this profile.
  for(const [alias,n] of Object.entries({dark1:'dk1',light1:'lt1',dark2:'dk2',light2:'lt2',text1:'dk1',text2:'dk2',background1:'lt1',background2:'lt2',hyperlink:'hlink',followedHyperlink:'folHlink'}))
    colors[alias]=colors[n]!;
  const mapping:Record<string,string>=Object.create(null) as Record<string,string>;
  for(const key of ['major','minor'] as const){
    const group=only(fonts,A,key+'Font');
    for(const [element,suffixes] of [['latin',['Ascii','HAnsi']],['ea',['EastAsia']],['cs',['Bidi']]] as const){
      const node=only(group,A,element),face=node.getAttribute('typeface');
      // Empty typefaces rely on script-specific/theme fallback and are not
      // guessed. Styles using them stay outside the editable subset.
      allowed(face!==null&&face.length<=512&&!/[\x00-\x1f\x7f]/.test(face),'Invalid theme typeface');
      for(const suffix of suffixes)mapping[key+suffix]=face;
    }
  }
  return {colors,fonts:mapping};
}
export function requireStyleReference(profile:DocxStyleProfile|undefined,id:string,kind:StyleKind):string{
  allowed(profile&&stableId(id)&&Object.hasOwn(profile.styles,id)&&profile.styles[id]?.kind===kind,'Missing or unsupported '+kind+' style '+id);
  return id;
}
/** Recognize the actual default-template definitions, not merely their names.
 * Conditional formatting, shading, custom inheritance and any other style
 * payload remain outside the editable profile.
 */
function builtinTableStyle(node:Element,id:'TableNormal'|'TableGrid'):boolean{
 try{
  const grid=id==='TableGrid';
  attrs(node,W,['styleId','type','default']);
  allowed(attribute(node,'styleId')===id&&attribute(node,'type')==='table'&&(grid?attribute(node,'default')===undefined:attribute(node,'default')==='1'),'Unexpected built-in table style identity');
  const items=new Map<string,Element>();
  for(const child of children(node)){const n=child.localName;allowed(child.namespaceURI===W&&n&&!items.has(n),'Invalid built-in table style child');items.set(n,child);}
  const take=(key:string)=>{const item=items.get(key);allowed(item,'Missing table style '+key);items.delete(key);return item;};
  value(take('name'));
  if(grid)allowed(value(take('basedOn'))==='TableNormal','Unexpected table ancestor');
  for(const key of ['uiPriority','rsid'])if(items.has(key)){const v=value(take(key));allowed(key==='rsid'?/^[0-9a-f]{8}$/i.test(v):/^\d+$/.test(v),'Unexpected table style metadata');}
  for(const key of ['semiHidden','unhideWhenUsed'])if(items.has(key)){const n=take(key);attrs(n,W,['val']);allowed(!children(n).length,'Nested table style flag');flag(attribute(n,'val'));}
  if(grid){const p=take('pPr');attrs(p,W,[]);const a=children(p);allowed(a.length===1&&name(a[0]!,W,'spacing'),'Unexpected Grid paragraph properties');const s=a[0]!;attrs(s,W,['after','line','lineRule']);allowed(!children(s).length&&attribute(s,'after')==='0'&&attribute(s,'line')==='240'&&attribute(s,'lineRule')==='auto','Unexpected Grid spacing');}
  const tbl=take('tblPr');attrs(tbl,W,[]);const props=children(tbl);
  allowed(props.length===(grid?3:2),'Unexpected built-in table property count');
  const ind=only(tbl,W,'tblInd');attrs(ind,W,['w','type']);allowed(!children(ind).length&&attribute(ind,'w')==='0'&&attribute(ind,'type')==='dxa','Unsupported table style indent');
  const pad=only(tbl,W,'tblCellMar');attrs(pad,W,[]);allowed(children(pad).length===4,'Unexpected table style margins');
  for(const edge of ['top','left','bottom','right']){const n=only(pad,W,edge);attrs(n,W,['w','type']);allowed(!children(n).length&&attribute(n,'type')==='dxa'&&attribute(n,'w')===(edge==='left'||edge==='right'?'108':'0'),'Unexpected table style margin');}
  if(grid){const b=only(tbl,W,'tblBorders');attrs(b,W,[]);allowed(children(b).length===6,'Unexpected Grid border');for(const edge of ['top','left','bottom','right','insideH','insideV']){const n=only(b,W,edge);attrs(n,W,['val','sz','space','color']);allowed(!children(n).length&&attribute(n,'val')==='single'&&attribute(n,'sz')==='4'&&attribute(n,'space')==='0'&&attribute(n,'color')==='auto','Unexpected Grid border style');}}
  allowed(items.size===0,'Unsupported built-in table style property');
  return true;
 }catch(error){if(error instanceof OfficeError&&error.code==='unsupported')return false;throw error;}
}
/** Default-template Header/Footer definitions have tab stops which are
 * inactive in the admitted tab-free single paragraph profile. Validate
 * the whole definition, not just the built-in name. */
function builtinStoryStyle(node:Element,id:'Header'|'Footer'):boolean{
 try{
  attrs(node,W,['styleId','type']);allowed(attribute(node,'styleId')===id&&attribute(node,'type')==='paragraph','Invalid story style');
  const items=new Map<string,Element>();
  for(const c of children(node)){const n=c.localName;allowed(c.namespaceURI===W&&n&&!items.has(n),'Unexpected story style property');items.set(n,c);}
  const take=(key:string)=>{const e=items.get(key);allowed(e,'Missing story style '+key);items.delete(key);return e;};
  value(take('name'));allowed(value(take('basedOn'))==='Normal','Unsupported story style ancestor');
  if(items.has('link'))allowed(value(take('link'))===id+'Char','Unsupported story style link');
  for(const key of ['uiPriority','rsid'])if(items.has(key)){const v=value(take(key));allowed(key==='rsid'?/^[0-9a-f]{8}$/i.test(v):/^\d+$/.test(v),'Unexpected story style metadata');}
  if(items.has('unhideWhenUsed')){const e=take('unhideWhenUsed');attrs(e,W,['val']);allowed(!children(e).length,'Nested style flag');flag(attribute(e,'val'));}
  const p=take('pPr');attrs(p,W,[]);const a=children(p);allowed(a.length===2&&name(a[0]!,W,'tabs')&&name(a[1]!,W,'spacing'),'Unsupported story paragraph style');
  const tabs=a[0]!;attrs(tabs,W,[]);const stops=children(tabs);allowed(stops.length===2,'Unexpected story tab stops');
  for(const [i,stop]of stops.entries()){allowed(name(stop,W,'tab'),'Unexpected story tab');attrs(stop,W,['val','pos']);allowed(!children(stop).length&&attribute(stop,'val')===(i?'right':'center')&&attribute(stop,'pos')===(i?'9360':'4680'),'Unsupported story tab geometry');}
  const sp=a[1]!;attrs(sp,W,['after','line','lineRule']);allowed(!children(sp).length&&attribute(sp,'after')==='0'&&attribute(sp,'line')==='240'&&attribute(sp,'lineRule')==='auto'&&items.size===0,'Unsupported story spacing or properties');
  return true;
 }catch(error){if(error instanceof OfficeError&&error.code==='unsupported')return false;throw error;}
}
export interface DocxStorySource{kind:'header'|'footer';path:string}
function readStyles(bytes:Uint8Array,options:ContainerOptions,readRun:RunReader,readParagraph:ParagraphReader,theme?:ThemeProfile,numbering?:NumberingSource):DocxStyleProfile{
  const root=parseXml(bytes,options).documentElement!;
  allowed(name(root,W,'styles'),'Expected Word styles root');attrs(root,W,[]);
  const defaults:TextStyle={},styles:Record<string,NamedStyle>=Object.create(null) as Record<string,NamedStyle>;
  const entries=new Map<string,{node:Element;kind:string;isDefault:boolean}>();
  let seenDefaults=false,seenLatent=false,defaultParagraph:string|undefined,defaultCharacter:string|undefined;
  for(const child of children(root)){
    if(name(child,W,'docDefaults')){
      allowed(!seenDefaults,'Duplicate document defaults');seenDefaults=true;attrs(child,W,[]);
      const kinds=new Set<string>();
      for(const item of children(child)){
        const n=item.localName;allowed(item.namespaceURI===W&&n!==null&&!kinds.has(n),'Duplicate or foreign default property');kinds.add(n);
        attrs(item,W,[]);const inner=children(item);
        allowed(inner.length<=1,'Unexpected default style payload');
        if(n==='rPrDefault'){if(inner.length)Object.assign(defaults,readRun(inner[0]!,undefined,theme));}
        else if(n==='pPrDefault'){if(inner.length)Object.assign(defaults,readParagraph(inner[0]!,undefined,undefined,theme));}
        else allowed(false,'Unsupported document defaults');
      }
    }else if(name(child,W,'latentStyles')){
      // Latent style metadata is not applied to runs/paragraphs in this profile.
      // It remains byte-for-byte in the original style part.
      allowed(!seenLatent,'Duplicate latent styles');seenLatent=true;
    }else if(name(child,W,'style')){
      const id=attribute(child,'styleId'),kind=attribute(child,'type');
      allowed(id&&id.length<=512&&kind&&!entries.has(id),'Duplicate or malformed style identifier');
      allowed(entries.size<2048,'Style count limit');
      const isDefault=attribute(child,'default')!==undefined&&flag(attribute(child,'default'));
      entries.set(id,{node:child,kind,isDefault});
      if(isDefault&&kind==='paragraph'){allowed(!defaultParagraph,'Duplicate default paragraph style');defaultParagraph=id;}
      if(isDefault&&kind==='character'){allowed(!defaultCharacter,'Duplicate default character style');defaultCharacter=id;}
    }else allowed(false,'Unsupported styles-root child');
  }
  // Parse individual definitions. An unused unsupported style is retained in
  // the source part but not exposed as a selectable model style.
  const raw=new Map<string,NamedStyle>();
  for(const [id,entry] of entries){
    if(!stableId(id)||(entry.kind!=='paragraph'&&entry.kind!=='character'))continue;
    try{
      attrs(entry.node,W,['styleId','type','default','customStyle']);
      if(attribute(entry.node,'customStyle')!==undefined)flag(attribute(entry.node,'customStyle'));
      let label=id,basedOn:string|undefined;const properties:TextStyle={},toggles:('bold'|'italic')[]=[],seen=new Set<string>();
      for(const child of children(entry.node)){
        const n=child.localName;allowed(child.namespaceURI===W&&n!==null&&!seen.has(n),'Duplicate or foreign style property');seen.add(n);
        if(n==='name')label=value(child);
        else if(n==='basedOn')basedOn=value(child);
        else if(n==='pPr'){allowed(entry.kind==='paragraph','Character style with paragraph properties');Object.assign(properties,readParagraph(child,undefined,undefined,theme));}
        else if(n==='rPr')Object.assign(properties,readRun(child,undefined,theme,true,toggles));
        else if(n==='rsid'){allowed(/^[0-9a-fA-F]{8}$/.test(value(child)),'Invalid style revision metadata');}
        else if(n==='next'||n==='link'||n==='aliases'||n==='uiPriority'){value(child);/* Preserved metadata; explicit style application only. */}
        else if(['qFormat','semiHidden','unhideWhenUsed','hidden','locked','autoRedefine','personal','personalCompose','personalReply'].includes(n)){
          attrs(child,W,['val']);allowed(!children(child).length,'Nested style flag');
          const enabled=flag(attribute(child,'val'));
          // Editing constraints/auto-redefinition cannot be treated as mere UI metadata.
          if(n==='locked'||n==='autoRedefine')allowed(!enabled,'Locked or auto-redefining style');
        }else allowed(false,'Unsupported named style property '+n);
      }
      allowed(label.length>0&&label.length<=512,'Style name limit');
      if(basedOn)allowed(stableId(basedOn),'Unsupported ancestor identifier');
      allowed(supportsNumberingReference(properties.writerList,numbering?.numbering),'Unsupported named numbering');
      raw.set(id,{id,name:label,kind:entry.kind,...(basedOn?{basedOn}:{}),...(toggles.length?{toggles}:{}),properties});
    }catch(error){if(!(error instanceof OfficeError)||error.code!=='unsupported')throw error;}
  }
  const failed=new Set<string>();
  function include(id:string,ancestors:Set<string>):boolean{
    if(Object.hasOwn(styles,id))return true;
    if(failed.has(id)||ancestors.has(id)||ancestors.size>=128)return false;
    const style=raw.get(id);if(!style){failed.add(id);return false;}
    ancestors.add(id);
    if(style.basedOn&&(raw.get(style.basedOn)?.kind!==style.kind||!include(style.basedOn,ancestors))){ancestors.delete(id);failed.add(id);return false;}
    ancestors.delete(id);styles[id]=style;return true;
  }
  for(const id of raw.keys())include(id,new Set());
  if(defaultParagraph)allowed(Object.hasOwn(styles,defaultParagraph),'Default paragraph style unsupported');
  if(defaultCharacter){
    const style=styles[defaultCharacter];allowed(style&&!style.basedOn&&!style.toggles?.length&&!Object.keys(style.properties).length,'Default character style must be empty');
  }
  allowed(supportsNumberingReference(defaults.writerList,numbering?.numbering),'Unsupported default numbering');
  const tables=[...entries.values()].filter(e=>e.kind==='table'&&e.isDefault);
  const builtinTableStyles=new Set<'TableNormal'|'TableGrid'>();
  const normal=entries.get('TableNormal'),grid=entries.get('TableGrid');
  if(normal&&tables.length===1&&tables[0]===normal&&builtinTableStyle(normal.node,'TableNormal')){
    builtinTableStyles.add('TableNormal');
    if(grid&&builtinTableStyle(grid.node,'TableGrid'))builtinTableStyles.add('TableGrid');
  }
  const builtinStoryStyles=new Set<'Header'|'Footer'>();
  if(defaultParagraph==='Normal')for(const kind of ['Header','Footer'] as const){const entry=entries.get(kind);if(entry&&builtinStoryStyle(entry.node,kind))builtinStoryStyles.add(kind);}
  return {styles,defaults,...(defaultParagraph?{defaultParagraph}:{}),tableDefaultsSafe:tables.length===0||builtinTableStyles.has('TableNormal'),builtinTableStyles,builtinStoryStyles};
}
function relationPath(main:string):string{
  const i=main.lastIndexOf('/');return main.slice(0,i+1)+'_rels/'+main.slice(i+1)+'.rels';
}
/** Permit only the explicit style/theme relationship graph and passive package properties.
 * Running story parts are admitted without any auxiliary relationships;
 * their semantic payload is verified by the main DOCX adapter. Unknown
 * active content, external links, signatures and other graphs remain blocked.
 */
export function readStyleProfile(pkg:OpcPackage,options:ContainerOptions,readRun:RunReader,readParagraph:ParagraphReader):{part:string;profile?:DocxStyleProfile;stories:ReadonlyMap<string,DocxStorySource>}{
  allowed(pkg.kind==='docx'&&pkg.conformance==='transitional'&&pkg.mainPart,'Expected Transitional DOCX');
  allowed(!pkg.hasMacros&&!pkg.hasSignatures&&!pkg.externalRelationships,'Protected or linked document');
  const main=pkg.mainPart,expected=new Set<string>(['[Content_Types].xml','_rels/.rels',main]);
  const root=pkg.relationships.get('')??[];
  let mainCount=0,stylePath:string|undefined,themePath:string|undefined,hasMetadata=false,settingsPath:string|undefined,numberingPath:string|undefined;
  const customParts:string[]=[];const extraSources=new Set<string>(),stories=new Map<string,DocxStorySource>();
  const typesSeen=new Set<string>();
  function admit(rel:OpcRelationship,type:string):string{
    const path=rel.targetPart;allowed(rel.mode==='Internal'&&path&&rel.fragment===null&&!expected.has(path)&&pkg.contentTypes.get(path)===type,'Unsupported or duplicate relationship target');
    expected.add(path);return path;
  }
  for(const rel of root){
    allowed(!typesSeen.has(rel.type),'Duplicate root relationship type');typesSeen.add(rel.type);
    if(rel.type===R+'officeDocument'){allowed(rel.targetPart===main&&rel.fragment===null&&rel.mode==='Internal','Invalid main part');mainCount++;}
    else if(rel.type===CORE){admit(rel,'application/vnd.openxmlformats-package.core-properties+xml');hasMetadata=true;}
    else if(rel.type===R+'extended-properties'){admit(rel,'application/vnd.openxmlformats-officedocument.extended-properties+xml');hasMetadata=true;}
    else if(rel.type==='http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail'){
      const type=rel.targetPart?pkg.contentTypes.get(rel.targetPart):undefined;allowed(type==='image/png'||type==='image/jpeg','Unsupported thumbnail');admit(rel,type);hasMetadata=true;
    }
    else allowed(false,'Unsupported root relationship');
  }
  allowed(mainCount===1,'Missing main relationship');
  const mainRels=pkg.relationships.get(main);
  if(mainRels){
    const relFile=[...pkg.parts.keys()].find(path=>path.toLowerCase()===relationPath(main).toLowerCase());
    allowed(relFile&&pkg.contentTypes.get(relFile)===REL_TYPE,'Main relationship part missing');expected.add(relFile);
    for(const rel of mainRels){
      if(rel.type===R+'styles'){allowed(!stylePath,'Duplicate styles part');stylePath=admit(rel,STYLES);}
      else if(rel.type===R+'theme'){allowed(!themePath,'Duplicate theme part');themePath=admit(rel,THEME);}
      else if(rel.type===R+'settings'){allowed(!settingsPath,'Duplicate settings');settingsPath=admit(rel,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');verifySettings(pkg.parts.get(settingsPath)!,options);hasMetadata=true;}
      else if(rel.type===R+'webSettings'){const path=admit(rel,'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml');verifyAuxiliaryXml(pkg.parts.get(path)!,options,'webSettings');hasMetadata=true;}
      else if(rel.type===R+'fontTable'){const path=admit(rel,'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml');verifyAuxiliaryXml(pkg.parts.get(path)!,options,'fonts');hasMetadata=true;}
      else if(rel.type===R+'numbering'){allowed(!numberingPath,'Duplicate numbering');numberingPath=admit(rel,'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');hasMetadata=true;}
      else if(rel.type==='http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects'){const path=admit(rel,'application/vnd.ms-word.stylesWithEffects+xml');verifyAuxiliaryXml(pkg.parts.get(path)!,options,'styles');hasMetadata=true;}
      else if(rel.type===R+'customXml'){const type=rel.targetPart?pkg.contentTypes.get(rel.targetPart):undefined;allowed(type==='application/xml'||type==='text/xml','Unsupported custom XML type');const path=admit(rel,type);verifyAuxiliaryXml(pkg.parts.get(path)!,options);customParts.push(path);hasMetadata=true;}
      else if(rel.type===R+'header'||rel.type===R+'footer'){
        const kind=rel.type===R+'header'?'header':'footer';
        allowed(stories.size<256&&stableId(rel.id)&&/^[A-Za-z][A-Za-z0-9_.-]{0,127}$/.test(rel.id),'Invalid running story relationship');
        const path=admit(rel,'application/vnd.openxmlformats-officedocument.wordprocessingml.'+kind+'+xml');stories.set(rel.id,{kind,path});
      }
      else if(rel.type===R+'image'){const type=rel.targetPart?pkg.contentTypes.get(rel.targetPart):undefined;allowed(type==='image/png'||type==='image/jpeg','Unsupported image type');admit(rel,type);}
      else allowed(false,'Unsupported document relationship');
    }
  }
  for(const path of customParts){
    const rels=pkg.relationships.get(path);if(!rels)continue;
    const relFile=[...pkg.parts.keys()].find(p=>p.toLowerCase()===relationPath(path).toLowerCase());
    allowed(relFile&&pkg.contentTypes.get(relFile)===REL_TYPE&&rels.length===1&&rels[0]!.type===R+'customXmlProps','Unsupported custom XML dependency');expected.add(relFile);extraSources.add(path);
    const props=admit(rels[0]!,'application/vnd.openxmlformats-officedocument.customXmlProperties+xml');verifyAuxiliaryXml(pkg.parts.get(props)!,options,'datastoreItem','http://schemas.openxmlformats.org/officeDocument/2006/customXml');
  }
  allowed([...pkg.relationships.keys()].every(k=>k===''||k===main||extraSources.has(k)),'Unexpected auxiliary relationship graph');
  allowed(pkg.parts.size===expected.size&&[...pkg.parts.keys()].every(k=>expected.has(k)),'Additional package parts require a later adapter');
  if(!stylePath&&!themePath&&!hasMetadata)return {part:main,stories};
  const theme=themePath?readTheme(pkg.parts.get(themePath)!,options):undefined;
  const numbering=numberingPath?readNumbering(numberingPath,pkg.parts.get(numberingPath)!,options):undefined;
  const profile:DocxStyleProfile=stylePath?readStyles(pkg.parts.get(stylePath)!,options,readRun,readParagraph,theme,numbering):{styles:{},defaults:{},tableDefaultsSafe:true};
  if(numbering)profile.numbering=numbering;
  return {part:main,profile,stories};
}
