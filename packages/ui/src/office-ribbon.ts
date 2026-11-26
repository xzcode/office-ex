// SPDX-License-Identifier: MIT
/** Compact, localized ribbon presentation for existing live controls.
 * Moving nodes (not cloning them) preserves handlers, state and permissions. */
import type {MessageKey} from '@office-ex/i18n';
type Translate=(key:MessageKey,args?:Readonly<Record<string,string|number>>)=>string;
export interface RibbonState {ribbonTab?:string}
export interface RibbonGroup {label:MessageKey;rows:(HTMLElement|undefined)[][]}
export interface RibbonTab {id:string;label:MessageKey;groups:RibbonGroup[]}
const NS='http://www.w3.org/2000/svg';
// Original 24-unit line pictograms, using common desktop command metaphors.
const paths:Record<string,string>={
 file:'M6 3h9l4 4v14H6z M15 3v5h4',
 home:'M3 11l9-8 9 8 M6 10v11h12V10 M10 21v-7h4v7',
 clock:'M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18 M12 7v5l4 2',
 recovery:'M4 9a8 8 0 1 1 0 6 M4 4v5h5 M12 7v5l4 2',
 info:'M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18 M12 10v7 M12 7h.1',
 arrowNext:'M9 5l7 7-7 7',
 bold:'M7 4h6a4 4 0 0 1 0 8H7z M7 12h7a4 4 0 0 1 0 8H7z',
 italic:'M10 4h10 M4 20h10 M15 4L9 20',
 underline:'M6 4v8a6 6 0 0 0 12 0V4 M4 21h16',
 left:'M4 5h16 M4 10h11 M4 15h16 M4 20h11',
 center:'M4 5h16 M7 10h10 M4 15h16 M7 20h10',
 right:'M4 5h16 M9 10h11 M4 15h16 M9 20h11',
 top:'M4 4h16 M12 7v13 M8 11l4-4 4 4',
 bottom:'M4 20h16 M12 4v13 M8 13l4 4 4-4',
 middle:'M4 12h16 M12 2v7 M9 6l3 3 3-3 M12 22v-7 M9 18l3-3 3 3',
 table:'M3 4h18v16H3z M3 9h18 M3 14h18 M9 4v16 M15 4v16',
 rowAdd:'M3 4h18v10H3z M3 9h18 M9 4v10 M15 4v10 M12 16v6 M9 19h6',
 rowDelete:'M3 3h18v11H3z M3 8h18 M9 3v11 M15 3v11 M9 17l6 6 M15 17l-6 6',
 colAdd:'M3 3h12v18H3z M3 9h12 M3 15h12 M9 3v18 M19 9v6 M16 12h6',
 colDelete:'M3 3h12v18H3z M3 9h12 M3 15h12 M9 3v18 M17 9l5 6 M22 9l-5 6',
 image:'M3 4h18v16H3z M3 17l6-6 4 4 3-3 5 5 M15 8a1.5 1.5 0 1 0 3 0 1.5 1.5 0 1 0-3 0',
 page:'M6 3h9l4 4v14H6z M15 3v5h4 M9 13h7 M9 17h7',
 break:'M5 3h14v6H5z M5 15h14v6H5z M3 12h4 M10 12h4 M17 12h4',
 split:'M5 4h14 M5 8h11 M5 16h14 M5 20h11 M4 12h16',
 list:'M4 5h1 M4 12h1 M4 19h1 M9 5h11 M9 12h11 M9 19h11',
 number:'M5 3v5 M3 4l2-1 M3 8h4 M3 12c4-2 5 1 0 5h4 M10 5h10 M10 14h10',
 wrap:'M4 5h16 M4 10h11a4 4 0 0 1 0 8h-3 M15 15l-3 3 3 3 M4 18h4',
 border:'M3 3h18v18H3z M3 12h18 M12 3v18',
 merge:'M3 4h18v16H3z M3 12h18 M12 4v4 M12 16v4 M5 8l4 4-4 4 M19 8l-4 4 4 4',
 unmerge:'M3 4h18v16H3z M3 12h18 M12 4v16 M9 8l-4 4 4 4 M15 8l4 4-4 4',
 erase:'M4 14l9-10 7 7-9 10H9z M8 10l8 8 M12 21h9',
 down:'M12 3v17 M6 14l6 6 6-6',
 rightArrow:'M3 12h17 M14 6l6 6-6 6',
 up:'M12 21V4 M6 10l6-6 6 6',
 leftArrow:'M21 12H4 M10 6l-6 6 6 6',
 add:'M12 5v14 M5 12h14',
 remove:'M5 6h14 M9 6V3h6v3 M7 6l1 15h8l1-15 M10 10v7 M14 10v7',
 copy:'M8 8h12v13H8z M16 8V3H3v13h5',
 rect:'M4 5h16v14H4z',
 ellipse:'M20 12a8 7 0 1 1-16 0 8 7 0 1 1 16 0',
 line:'M4 20L20 4',
 roundRect:'M8 5h8a4 4 0 0 1 4 4v6a4 4 0 0 1-4 4H8a4 4 0 0 1-4-4V9a4 4 0 0 1 4-4z',
 triangle:'M12 4l9 16H3z',
 diamond:'M12 3l9 9-9 9-9-9z',
 parallelogram:'M8 4h13l-5 16H3z',
 connector:'M3 3h5v5H3z M16 16h5v5h-5z M8 8l8 8',
 textbox:'M3 4h18v16H3z M8 8h8 M12 8v9 M9 17h6',
 play:'M3 4h18v14H3z M10 8l5 3-5 3z M9 21h6',
 zoomIn:'M10 3a7 7 0 1 0 0 14 7 7 0 1 0 0-14 M15 15l6 6 M7 10h6 M10 7v6',
 zoomOut:'M10 3a7 7 0 1 0 0 14 7 7 0 1 0 0-14 M15 15l6 6 M7 10h6',
 fit:'M3 9V3h6 M15 3h6v6 M21 15v6h-6 M9 21H3v-6 M7 7h10v10H7z',
 group:'M3 3h12v12H3z M9 9h12v12H9z',
 ungroup:'M3 3h7v7H3z M14 14h7v7h-7z M12 5l7 7',
 layers:'M3 8l9-5 9 5-9 5z M3 12l9 5 9-5 M3 16l9 5 9-5',
 rotate:'M6 7a8 8 0 1 1-2 9 M6 2v5H1',
 freeze:'M3 4h18v16H3z M3 10h18 M10 4v16 M15 12v6 M12 15h6',
 calc:'M6 3h12v18H6z M8 6h8 M9 11h1 M14 11h1 M9 15h1 M14 15h1 M9 18h1 M14 18h1',
 settings:'M3 6h18 M3 12h18 M3 18h18 M8 3v6 M16 9v6 M10 15v6',
 undo:'M8 4L3 9l5 5 M3 9h12a5 5 0 0 1 0 10h-3',
 redo:'M16 4l5 5-5 5 M21 9H9a5 5 0 0 0 0 10h3',
 save:'M4 3h13l3 3v15H4z M8 3v7h8V3 M8 21v-7h8v7',
 folder:'M3 6h7l2 3h9v11H3z',
 print:'M7 3h10v5H7z M5 8h14v9h-2 M7 17H5V8 M7 14h10v7H7z M16 11h1',
 check:'M4 12l5 5 11-11',
 close:'M6 6l12 12 M18 6L6 18',
 chevron:'M6 9l6 6 6-6',
 panel:'M3 4h18v16H3z M15 4v16',
 text:'M3 4h18 M12 4v16 M8 20h8',
 percent:'M5 19L19 5 M9 7a2 2 0 1 1-4 0 2 2 0 1 1 4 0 M19 17a2 2 0 1 1-4 0 2 2 0 1 1 4 0'
};
export function officeIcon(name:string):SVGSVGElement{
 const svg=document.createElementNS(NS,'svg');svg.setAttribute('viewBox','0 0 24 24');svg.setAttribute('aria-hidden','true');svg.setAttribute('focusable','false');svg.classList.add('office-icon');
 const path=document.createElementNS(NS,'path');path.setAttribute('d',paths[name]??paths.settings!);path.setAttribute('fill','none');path.setAttribute('stroke','currentColor');path.setAttribute('stroke-width','1.7');path.setAttribute('stroke-linecap','round');path.setAttribute('stroke-linejoin','round');svg.append(path);return svg;
}
const exact:Record<string,string>={
 'home':'page','open':'folder','new-menu':'add','undo':'undo','redo':'redo','save':'save','save-as':'save','export-pdf':'page','print-document':'print',
 'writer-ribbon-table':'table','writer-ribbon-pictures':'image','writer-ribbon-layout':'page','writer-ribbon-typography':'text','writer-page-edit':'page','writer-exit-page-edit':'text','writer-continuous':'text','writer-split':'split','writer-page-break':'break','writer-section-break':'break','writer-insert-table':'table','writer-table-add-row':'rowAdd','writer-table-delete-row':'rowDelete','writer-table-add-column':'colAdd','writer-table-delete-column':'colDelete','writer-table-border':'border','writer-table-delete':'remove','writer-insert-image':'image','writer-insert-inline-image':'image','writer-apply-page':'check','writer-list-restart':'number',
 'sheet-merge':'merge','sheet-unmerge':'unmerge','sheet-wrap':'wrap','sheet-border':'border','sheet-clear':'erase','sheet-fill-down':'down','sheet-fill-right':'rightArrow','sheet-add':'add','sheet-delete':'remove','sheet-move-left':'leftArrow','sheet-move-right':'rightArrow','sheet-insert-rows':'rowAdd','sheet-delete-rows':'rowDelete','sheet-insert-columns':'colAdd','sheet-delete-columns':'colDelete','sheet-freeze':'freeze','sheet-unfreeze':'freeze','sheet-grid-lines':'table','sheet-recalculate':'calc','sheet-cancel-calculation':'close','sheet-rtl':'leftArrow',
 'slide-add':'add','slide-duplicate':'copy','slide-delete':'remove','slide-earlier':'up','slide-later':'down','slide-present':'play','slide-zoom-out':'zoomOut','slide-zoom-in':'zoomIn','slide-zoom-fit':'fit','slide-group':'group','slide-ungroup':'ungroup','slide-edit-group':'group','slide-exit-group':'leftArrow','slide-object-duplicate':'copy','slide-object-delete':'remove','slide-connect-objects':'connector','slide-detach-connector':'unmerge','slide-flipH':'rightArrow','slide-flipV':'down','slide-insert-text':'textbox','slide-insert-image':'image','slide-fit-text':'fit','slide-reset-fit':'text'
};
function iconName(id:string):string{
 if(exact[id])return exact[id]!;
 if(id.startsWith('slide-insert-'))return id.slice(13);
 if(id.includes('-layer-'))return 'layers';
 if(/(?:align-|text-)(start|left)$/.test(id))return 'left';
 if(/(?:align-|text-)(center|middle)$/.test(id))return id.endsWith('middle')?'middle':'center';
 if(/(?:align-|text-)(end|right)$/.test(id))return 'right';
 if(id.endsWith('-top'))return 'top';if(id.endsWith('-bottom'))return 'bottom';
 for(const key of ['bold','italic','underline'])if(id.endsWith('-'+key))return key;
 return 'settings';
}
const iconOnly=new Set(['bold','italic','underline','left','center','right','top','bottom','middle','up','down','leftArrow','rightArrow','zoomIn','zoomOut','fit','undo','redo','close']);
export function decorateOfficeButton(button:HTMLButtonElement,name=iconName(button.id),compact=false):void{
 if(button.dataset.officeIcon)return;
 const label=button.textContent?.trim()??'';if(!label)return;
 button.dataset.officeIcon=name;button.title=label;button.setAttribute('aria-label',label);
 const span=document.createElement('span');span.className='office-command-label';span.textContent=label;
 button.replaceChildren(officeIcon(name),span);button.classList.add('office-command');
 if(compact||iconOnly.has(name)&&!['sheet-fill-down','sheet-fill-right'].includes(button.id))button.classList.add('office-command-icon');
}
/** Find an actual control or its existing field label. No duplicate IDs. */
export function ribbonNodes(root:ParentNode,...ids:string[]):(HTMLElement|undefined)[]{
 return ids.map(id=>{const e=root.querySelector<HTMLElement>('#'+id);if(!e)return;return e.matches('input,select,textarea')?(e.closest('label')??e):e;});
}
export function installRibbon(root:HTMLElement,state:RibbonState,t:Translate,tabs:RibbonTab[],canSwitch:()=>boolean=()=>true):HTMLElement{
 const ribbon=document.createElement('section');ribbon.className='office-ribbon';ribbon.setAttribute('aria-label',t('ribbonCommands'));
 const strip=document.createElement('div');strip.className='office-ribbon-tabs';strip.setAttribute('role','tablist');strip.setAttribute('aria-label',t('ribbonCommands'));ribbon.append(strip);
 const panels=document.createElement('div');panels.className='office-ribbon-panels';ribbon.append(panels);
 const buttons:HTMLButtonElement[]=[],views:HTMLElement[]=[];
 const select=(id:string,focus=false)=>{
  if(!canSwitch())return;
  const index=Math.max(0,tabs.findIndex(tab=>tab.id===id));state.ribbonTab=tabs[index]!.id;
  buttons.forEach((b,i)=>{b.setAttribute('aria-selected',String(i===index));b.tabIndex=i===index?0:-1;views[i]!.hidden=i!==index;});
  if(focus)buttons[index]?.focus();
 };
 const prefix='ribbon-'+crypto.randomUUID();
 for(const [index,tab] of tabs.entries()){
  const b=document.createElement('button');b.type='button';b.className='office-ribbon-tab';b.id=prefix+'-tab-'+tab.id;b.dataset.ribbonTab=tab.id;b.textContent=t(tab.label);b.setAttribute('role','tab');b.setAttribute('aria-controls',prefix+'-panel-'+tab.id);b.onpointerdown=e=>{if(e.pointerType==='mouse')e.preventDefault();};b.onclick=()=>select(tab.id);b.onkeydown=e=>{if(['ArrowLeft','ArrowRight','Home','End'].includes(e.key)){e.preventDefault();const direction=document.documentElement.dir==='rtl'?-1:1,step=e.key==='ArrowRight'?direction:-direction;const next=e.key==='Home'?0:e.key==='End'?tabs.length-1:(index+step+tabs.length)%tabs.length;select(tabs[next]!.id,true);}};
  strip.append(b);buttons.push(b);
  const panel=document.createElement('div');panel.className='office-ribbon-panel';panel.id=prefix+'-panel-'+tab.id;panel.dataset.ribbonPanel=tab.id;panel.setAttribute('role','tabpanel');panel.setAttribute('aria-labelledby',b.id);
  for(const spec of tab.groups){
   const rows=spec.rows.map(row=>row.filter((n):n is HTMLElement=>!!n)).filter(row=>row.length);
   if(!rows.length)continue;
   const group=document.createElement('div');group.className='office-ribbon-group';
   const tools=document.createElement('div');tools.className='office-ribbon-tools';
   for(const items of rows){
    const row=document.createElement('div');row.className='office-ribbon-row';
    for(const node of items){
     if(node instanceof HTMLButtonElement)decorateOfficeButton(node);
     if(node instanceof HTMLLabelElement){
      node.classList.add('office-ribbon-field');
      const field=node.querySelector('input,select,textarea'),label=node.querySelector('span')?.textContent?.trim()||field?.getAttribute('aria-label');
      if(label){node.title=label;field?.setAttribute('aria-label',label);}
      if(field?.matches('input[type=color]'))node.classList.add('office-ribbon-color');
     }else if(node.matches('select,input'))node.classList.add('office-ribbon-input');
     row.append(node);
    }
    tools.append(row);
   }
   const label=document.createElement('div');label.className='office-ribbon-group-title';label.textContent=t(spec.label);group.append(tools,label);panel.append(group);
  }
  views.push(panel);panels.append(panel);
 }
 root.prepend(ribbon);
 const initial=tabs.some(tab=>tab.id===state.ribbonTab)?state.ribbonTab!:tabs[0]!.id;
 // Initialization does not require an active editor or flush.
 state.ribbonTab=initial;buttons.forEach((b,i)=>{const active=tabs[i]!.id===initial;b.setAttribute('aria-selected',String(active));b.tabIndex=active?0:-1;views[i]!.hidden=!active;});
 return ribbon;
}
