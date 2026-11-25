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
import type {ContainerOptions} from './limits.js';
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
}
type RunReader=(el:Element,reference?:((id:string)=>void),theme?:ThemeProfile,named?:boolean,toggles?:('bold'|'italic')[])=>TextStyle;
type ParagraphReader=(el:Element,reference?:((id:string)=>void))=>TextStyle;
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
      if(face&&face.length<=512&&!/[\x00-\x1f\x7f]/.test(face))
        for(const suffix of suffixes)mapping[key+suffix]=face;
    }
  }
  return {colors,fonts:mapping};
}
export function requireStyleReference(profile:DocxStyleProfile|undefined,id:string,kind:StyleKind):string{
  allowed(profile&&stableId(id)&&Object.hasOwn(profile.styles,id)&&profile.styles[id]?.kind===kind,'Missing or unsupported '+kind+' style '+id);
  return id;
}
function readStyles(bytes:Uint8Array,options:ContainerOptions,readRun:RunReader,readParagraph:ParagraphReader,theme?:ThemeProfile):DocxStyleProfile{
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
        else if(n==='pPrDefault'){if(inner.length)Object.assign(defaults,readParagraph(inner[0]!));}
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
        else if(n==='pPr'){allowed(entry.kind==='paragraph','Character style with paragraph properties');Object.assign(properties,readParagraph(child));}
        else if(n==='rPr')Object.assign(properties,readRun(child,undefined,theme,true,toggles));
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
  return {styles,defaults,...(defaultParagraph?{defaultParagraph}:{}),tableDefaultsSafe:![...entries.values()].some(e=>e.kind==='table'&&e.isDefault)};
}
function relationPath(main:string):string{
  const i=main.lastIndexOf('/');return main.slice(0,i+1)+'_rels/'+main.slice(i+1)+'.rels';
}
/** Permit only the explicit style/theme relationship graph and passive package properties.
 * No settings, numbering, drawings, comments, macros, signatures, external links,
 * headers, footers or other relationship dependencies are admitted.
 */
export function readStyleProfile(pkg:OpcPackage,options:ContainerOptions,readRun:RunReader,readParagraph:ParagraphReader):{part:string;profile?:DocxStyleProfile}{
  allowed(pkg.kind==='docx'&&pkg.conformance==='transitional'&&pkg.mainPart,'Expected Transitional DOCX');
  allowed(!pkg.hasMacros&&!pkg.hasSignatures&&!pkg.externalRelationships,'Protected or linked document');
  const main=pkg.mainPart,expected=new Set<string>(['[Content_Types].xml','_rels/.rels',main]);
  const root=pkg.relationships.get('')??[];
  let mainCount=0,stylePath:string|undefined,themePath:string|undefined,hasMetadata=false;
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
      else if(rel.type===R+'image'){const type=rel.targetPart?pkg.contentTypes.get(rel.targetPart):undefined;allowed(type==='image/png'||type==='image/jpeg','Unsupported image type');admit(rel,type);}
      else allowed(false,'Unsupported document relationship');
    }
  }
  allowed([...pkg.relationships.keys()].every(k=>k===''||k===main),'Unexpected auxiliary relationship graph');
  allowed(pkg.parts.size===expected.size&&[...pkg.parts.keys()].every(k=>expected.has(k)),'Additional package parts require a later adapter');
  if(!stylePath&&!themePath&&!hasMetadata)return {part:main};
  const theme=themePath?readTheme(pkg.parts.get(themePath)!,options):undefined;
  const profile=stylePath?readStyles(pkg.parts.get(stylePath)!,options,readRun,readParagraph,theme):{styles:{},defaults:{},tableDefaultsSafe:true};
  return {part:main,profile};
}
