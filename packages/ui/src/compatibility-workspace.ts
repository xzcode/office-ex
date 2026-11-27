// SPDX-License-Identifier: MIT
import type {CompatDocxView,CompatParagraph,CompatibleOfficeView,CompatibleValueKind} from '@office-ex/format-service';
import type {MessageKey} from '@office-ex/i18n';
import {OfficeError} from '@office-ex/contracts';
type Change={id:number;before:string;after:string};
type View=CompatDocxView|CompatibleOfficeView;
const isOffice=(v:View):v is CompatibleOfficeView=>'pages'in v;
export class CompatibilitySession{
 readonly format:'docx'|'xlsx'|'pptx';
 readonly original=new Map<number,string>();
 private edits=new Map<number,string>();
 private saved=new Map<number,string>();
 private past:Change[]=[];private future:Change[]=[];
 revision=0;saving=false;page=0;chunk=0;
 constructor(readonly view:View){
  this.format=isOffice(view)?view.format:'docx';
  const add=(id:number,text:string)=>this.original.set(id,text);
  if(isOffice(view))for(const p of view.pages)for(const f of p.fields)add(f.id,f.text);
  else{
   const paragraph=(p:CompatParagraph)=>{for(const item of p.items)if(item.kind==='text')add(item.segment.id,item.segment.text);};
   for(const b of view.blocks)if(b.kind==='paragraph')paragraph(b);else if(b.kind==='table')for(const row of b.rows)for(const cell of row)for(const p of cell)paragraph(p);
  }
 }
 get dirty(){return this.edits.size!==this.saved.size||[...this.edits].some(([id,v])=>this.saved.get(id)!==v);}
 get canUndo(){return this.past.length>0&&!this.saving;}get canRedo(){return this.future.length>0&&!this.saving;}
 get(id:number){return this.edits.get(id)??this.original.get(id)??'';}
 private change(id:number,value:string){if(value===this.original.get(id))this.edits.delete(id);else this.edits.set(id,value);this.revision++;}
 set(id:number,value:string){
  if(this.saving||!this.original.has(id)||value.length>32767)throw new OfficeError('invalid','Invalid compatible text value');
  const before=this.get(id);if(before===value)return;
  const total=[...this.edits].reduce((n,[k,v])=>n+(k===id?0:v.length),0)+value.length;
  if(total>2*1024*1024)throw new OfficeError('limit','Compatible edit budget exceeded');
  this.change(id,value);this.past.push({id,before,after:value});this.future=[];
  let historyBytes=this.past.reduce((n,c)=>n+c.before.length+c.after.length,0);
  while(this.past.length>128||historyBytes>2*1024*1024){const c=this.past.shift()!;historyBytes-=c.before.length+c.after.length;}
 }
 undo(){if(!this.canUndo)return;const c=this.past.pop()!;this.change(c.id,c.before);this.future.push(c);}
 redo(){if(!this.canRedo)return;const c=this.future.pop()!;this.change(c.id,c.after);this.past.push(c);}
 bytes(documentId:string){return new TextEncoder().encode(JSON.stringify({documentId,sourceSha256:this.view.sourceSha256,edits:[...this.edits].map(([id,text])=>({id,text}))}));}
 markSaved(){this.saved=new Map(this.edits);}
}
interface Host {t(key:MessageKey,args?:Record<string,string|number>):string;change():void;render():void;error(error:unknown):void;setComposing(value:boolean):void;}
export function compatibilityWorkspace(session:CompatibilitySession,host:Host):HTMLElement{
 const {t}=host,view=session.view,root=document.createElement('section');root.className='compat-edit-workspace workspace-card';root.id='compatibility-workspace';
 const el=<K extends keyof HTMLElementTagNameMap>(tag:K,text?:string,className?:string)=>{const e=document.createElement(tag);if(text!==undefined)e.textContent=text;if(className)e.className=className;return e;};
 root.append(el('h2',t('compatibilityEditTitle')),el('p',t('compatibilityEditNotice'),'boundary-note'),el('p',t('compatibilityEditCount',{count:view.editableSegments}),'muted'));
 const bar=el('div',undefined,'compat-edit-navigation'),body=el('div',undefined,'compat-edit-body');root.append(bar,body);
 function input(id:number,kind:CompatibleValueKind,label:string){
  const field=kind==='text'?el('textarea'):el('input');field.className='compat-value';field.dataset.compatId=String(id);field.setAttribute('aria-label',label);field.value=session.get(id);field.maxLength=32767;field.disabled=session.saving;
  if(field instanceof HTMLTextAreaElement)field.rows=Math.min(4,Math.max(1,Math.ceil(field.value.length/90)));
  else field.inputMode=kind==='number'?'decimal':'numeric';
  field.addEventListener('compositionstart',()=>host.setComposing(true));field.addEventListener('compositionend',()=>host.setComposing(false));
  field.oninput=()=>{try{session.set(id,field.value);host.change();}catch(e){field.value=session.get(id);host.error(e);}};
  return field;
 }
 const paragraph=(p:CompatParagraph)=>{
  const line=el('div',undefined,'compat-paragraph');if(p.alignment)line.style.textAlign=p.alignment;if(p.indent)line.style.paddingInlineStart=p.indent+'px';
  let editable=false;
  for(const item of p.items){
   if(item.kind==='text'){
    editable=true;const field=input(item.segment.id,'text',t('compatibilityEditSegment',{count:item.segment.id+1})),s=item.segment.style;
    if(s.font)field.style.fontFamily=s.font;if(s.size)field.style.fontSize=s.size+'pt';if(s.color)field.style.color=s.color;if(s.fill)field.style.backgroundColor=s.fill;
    if(s.bold)field.style.fontWeight='bold';if(s.italic)field.style.fontStyle='italic';field.style.textDecoration=[s.strike?'line-through':'',s.underline?'underline':''].filter(Boolean).join(' ');
    line.append(field);
   }else if(item.kind==='image'&&!isOffice(view)){
    const resource=view.resources[item.resourceId];
    if(resource){const img=el('img');img.src='data:'+resource.mimeType+';base64,'+resource.dataBase64;img.alt=t('compatibilityLockedObject');if(item.width)img.style.width=item.width+'px';line.append(img);}
    if(item.embedded)line.append(el('span',t('compatibilityLockedObject'),'compat-locked'));
   }else line.append(el('span',t('compatibilityLockedObject'),'compat-locked'));
  }
  if(!editable&&p.text)line.prepend(el('p',p.text,'copyable-text'));
  return line;
 };
 function select(values:string[],index:number,change:(n:number)=>void){
  const s=el('select');s.setAttribute('aria-label',t('compatibilityEditTitle'));s.disabled=session.saving;
  values.forEach((value,i)=>{const o=el('option',value);o.value=String(i);s.append(o);});s.value=String(index);s.onchange=()=>change(Number(s.value));bar.append(s);
 }
 const perPage=60;
 if(isOffice(view)){
  session.page=Math.min(session.page,view.pages.length-1);const page=view.pages[session.page]!;
  select(view.pages.map(p=>p.name),session.page,i=>{session.page=i;session.chunk=0;host.render();});
  const chunks=Math.max(1,Math.ceil(page.fields.length/perPage));session.chunk=Math.min(session.chunk,chunks-1);
  if(chunks>1)select(Array.from({length:chunks},(_,i)=>String(i+1)+' / '+chunks),session.chunk,i=>{session.chunk=i;host.render();});
  if(page.locked)body.append(el('p',t('compatibilityLockedCount',{count:page.locked}),'muted'));
  for(const f of page.fields.slice(session.chunk*perPage,(session.chunk+1)*perPage)){
   const row=el('label',undefined,'compat-value-row');row.append(el('span',f.label),input(f.id,f.kind,f.label));body.append(row);
  }
 }else{
  const chunks=Math.max(1,Math.ceil(view.blocks.length/perPage));session.chunk=Math.min(session.chunk,chunks-1);
  if(chunks>1)select(Array.from({length:chunks},(_,i)=>String(i+1)+' / '+chunks),session.chunk,i=>{session.chunk=i;host.render();});
  for(const b of view.blocks.slice(session.chunk*perPage,(session.chunk+1)*perPage)){
   if(b.kind==='paragraph')body.append(paragraph(b));
   else if(b.kind==='table'){
    const table=el('table'),tb=el('tbody');for(const r of b.rows){const tr=el('tr');for(const c of r){const td=el('td');for(const p of c)td.append(paragraph(p));tr.append(td);}tb.append(tr);}table.append(tb);body.append(table);
   }else body.append(el('p',t('compatibilityLockedObject'),'compat-locked'));
  }
 }
 return root;
}
