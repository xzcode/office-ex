// SPDX-License-Identifier: MIT
/** Admit passive or understood auxiliary Word parts only when the main and
 * active style parsers reject features that would use opaque dependencies.
 * These parts are never executed, modified, or fetched. */
import type {Element} from '@xmldom/xmldom';
import {OfficeError} from '@office-ex/contracts';
import {parseXml} from './xml.js';
import type {ContainerOptions} from './limits.js';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC='http://schemas.openxmlformats.org/markup-compatibility/2006';
const W14='http://schemas.microsoft.com/office/word/2010/wordml';
const M='http://schemas.openxmlformats.org/officeDocument/2006/math';
const O='urn:schemas-microsoft-com:office:office';
const XMLNS='http://www.w3.org/2000/xmlns/';
function valid(v:unknown,reason:string):asserts v{if(!v)throw new OfficeError('unsupported','DOCX auxiliary profile: '+reason);}
function elements(el:Element):Element[]{
  const output:Element[]=[];for(let n=el.firstChild;n;n=n.nextSibling){if(n.nodeType===1)output.push(n as Element);else valid(n.nodeType===3&&!(n.nodeValue??'').trim()||n.nodeType===8,'Unexpected auxiliary content');}return output;
}
function attrs(el:Element,allowed:readonly string[],namespace=W):void{
  for(let i=0;i<el.attributes.length;i++){const a=el.attributes.item(i)!;if(a.namespaceURI===XMLNS)continue;
    if(a.namespaceURI===MC&&a.localName==='Ignorable'&&el.localName==='settings'&&a.value.split(/\s+/).every(prefix=>prefix&&el.lookupNamespaceURI(prefix)?.startsWith('http://schemas.microsoft.com/office/')))continue;
    valid(a.namespaceURI===namespace&&allowed.includes(a.localName??''),'Unknown auxiliary attribute on '+el.localName);
  }
}
const prop=(el:Element,k:string)=>el.getAttributeNS(W,k);
function empty(el:Element):void{valid(elements(el).length===0,'Unexpected nested settings content');}
const COMPAT_NAMES=new Set(['compatibilityMode','overrideTableStyleFontSizeAndJustification','enableOpenTypeFeatures','doNotFlipMirrorIndents']);
export function verifySettings(bytes:Uint8Array,options:ContainerOptions):boolean{
  const root=parseXml(bytes,options).documentElement!;valid(root.namespaceURI===W&&root.localName==='settings','Expected settings');attrs(root,[]);
  const seen=new Set<string>();let differentEvenOddPages=false;
  for(const el of elements(root)){
    const name=el.localName??'',key=el.namespaceURI+':'+name;valid(!seen.has(key),'Duplicate setting');seen.add(key);
    if(el.namespaceURI===W){
      switch(name){
        case 'evenAndOddHeaders':{attrs(el,['val']);empty(el);const value=el.hasAttributeNS(W,'val')?prop(el,'val'):'1';valid(['0','1','true','false','on','off'].includes(value??''),'Invalid even/odd header option');differentEvenOddPages=['1','true','on'].includes(value??'');break;}
        case 'zoom':attrs(el,['val','percent']);empty(el);break;
        case 'proofState':attrs(el,['spelling','grammar']);empty(el);break;
        case 'defaultTabStop':attrs(el,['val']);empty(el);valid(prop(el,'val')==='720','Nondefault tab stop unsupported');break;
        case 'characterSpacingControl':attrs(el,['val']);empty(el);valid(prop(el,'val')==='doNotCompress','Character compression unsupported');break;
        case 'savePreviewPicture':case 'doNotAutoCompressPictures':attrs(el,['val']);empty(el);break;
        case 'themeFontLang':attrs(el,['val','eastAsia','bidi']);empty(el);break;
        case 'decimalSymbol':case 'listSeparator':attrs(el,['val']);empty(el);break;
        case 'clrSchemeMapping':{
          const defaults={bg1:'light1',t1:'dark1',bg2:'light2',t2:'dark2',accent1:'accent1',accent2:'accent2',accent3:'accent3',accent4:'accent4',accent5:'accent5',accent6:'accent6',hyperlink:'hyperlink',followedHyperlink:'followedHyperlink'};
          attrs(el,Object.keys(defaults));empty(el);
          for(const [k,v] of Object.entries(defaults))valid(!el.hasAttributeNS(W,k)||prop(el,k)===v,'Nondefault color scheme mapping');
          break;
        }
        case 'compat':{
          attrs(el,[]);
          for(const c of elements(el)){
            valid(c.namespaceURI===W,'Foreign compatibility setting');
            if(c.localName==='useFELayout'){attrs(c,['val']);empty(c);}
            else{valid(c.localName==='compatSetting','Unsupported compatibility flag');attrs(c,['name','uri','val']);empty(c);
              const name=prop(c,'name'),v=prop(c,'val');
              valid(name&&COMPAT_NAMES.has(name)&&prop(c,'uri')==='http://schemas.microsoft.com/office/word','Unknown compatibility setting');
              valid(name==='compatibilityMode'?['12','14','15'].includes(v??''):v==='0'||v==='1','Unsupported compatibility value');}
          }break;
        }
        case 'rsids':{
          attrs(el,[]);for(const c of elements(el)){valid(c.namespaceURI===W&&['rsidRoot','rsid'].includes(c.localName??''),'Unknown edit-session identifier');attrs(c,['val']);empty(c);valid(/^[0-9A-Fa-f]{8}$/.test(prop(c,'val')??''),'Invalid edit-session identifier');}break;
        }
        case 'shapeDefaults':{
          // No VML/shape content is admitted in the main document.
          attrs(el,[]);for(const c of elements(el))valid(c.namespaceURI===O&&['shapedefaults','shapelayout'].includes(c.localName??''),'Unknown shape defaults');
          break;
        }
        default:valid(false,'Unsupported setting '+name);
      }
    }else if(el.namespaceURI===M&&name==='mathPr'){
      // Equation objects are rejected in main XML and active styles.
    }else if(el.namespaceURI===W14&&['docId','defaultImageDpi'].includes(name)){
      attrs(el,['val'],W14);empty(el);
    }else valid(false,'Unknown settings namespace '+name);
  }
  return differentEvenOddPages;
}
export function verifyAuxiliaryXml(bytes:Uint8Array,options:ContainerOptions,rootName?:string,namespace=W):void{
  const root=parseXml(bytes,options).documentElement!;
  if(rootName)valid(root.localName===rootName&&root.namespaceURI===namespace,'Unexpected auxiliary root '+rootName);
}

