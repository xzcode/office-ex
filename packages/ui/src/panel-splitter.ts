// SPDX-License-Identifier: MIT
/** Pointer-captured, keyboard-accessible splitter. Width preferences are
 * UI-only; they never modify a document, its undo history or save state. */
export function panelWidth(key:string,fallback:number,min:number,max:number):number{
 try{const raw=localStorage.getItem('office-ex.panel.'+key);if(raw!==null){const n=Number(raw);if(Number.isFinite(n)&&n>=min&&n<=max)return n;}}catch{/* optional local preference */}
 return fallback;
}
export function persistPanelWidth(key:string,value:number):void{try{localStorage.setItem('office-ex.panel.'+key,String(value));}catch{/* session width still works */}}
export function createPanelSplitter(options:{key:string;label:string;side:'start'|'end';panel:HTMLElement;width():number;setWidth(value:number):void;limits():{min:number;max:number}}):HTMLDivElement{
 const handle=document.createElement('div');handle.className='office-panel-splitter';handle.dataset.splitter=options.key;handle.tabIndex=0;handle.setAttribute('role','separator');handle.setAttribute('aria-label',options.label);handle.setAttribute('aria-orientation','vertical');
 const bounds=()=>options.limits();
 const update=(value:number,persist=false)=>{
  const {min,max}=bounds(),next=Math.round(Math.max(min,Math.min(max,value)));
  options.setWidth(next);handle.setAttribute('aria-valuemin',String(min));handle.setAttribute('aria-valuemax',String(max));handle.setAttribute('aria-valuenow',String(next));
  if(persist)persistPanelWidth(options.key,next);
 };
 const direction=()=>{const p=options.panel.getBoundingClientRect(),h=handle.getBoundingClientRect();return p.left<h.left?1:-1;};
 let drag:{id:number;x:number;width:number;direction:number}|undefined;
 handle.onpointerdown=e=>{
  if(e.button!==0||drag)return;
  e.preventDefault();e.stopPropagation();handle.focus({preventScroll:true});drag={id:e.pointerId,x:e.clientX,width:options.panel.getBoundingClientRect().width,direction:direction()};handle.setPointerCapture(e.pointerId);handle.classList.add('dragging');
 };
 handle.onpointermove=e=>{if(drag?.id!==e.pointerId)return;e.preventDefault();update(drag.width+(e.clientX-drag.x)*drag.direction);};
 const end=(e:PointerEvent)=>{if(drag?.id!==e.pointerId)return;drag=undefined;handle.classList.remove('dragging');if(handle.hasPointerCapture(e.pointerId))handle.releasePointerCapture(e.pointerId);persistPanelWidth(options.key,options.width());};
 handle.onpointerup=end;handle.onpointercancel=end;handle.onlostpointercapture=end;
 handle.onkeydown=e=>{
  const {min,max}=bounds();let value=options.width();
  if(e.key==='Home')value=min;else if(e.key==='End')value=max;else if(e.key==='ArrowLeft'||e.key==='ArrowRight')value+=(e.key==='ArrowRight'?1:-1)*direction()*(e.shiftKey?24:8);else return;
  e.preventDefault();e.stopPropagation();update(value,true);
 };
 handle.setAttribute('aria-valuenow',String(options.width()));
 return handle;
}

export function panelCollapsed(key:string):boolean{try{return localStorage.getItem('office-ex.panel.'+key+'.collapsed')==='true';}catch{return false;}}
export function persistPanelCollapsed(key:string,value:boolean):void{try{localStorage.setItem('office-ex.panel.'+key+'.collapsed',String(value));}catch{/* UI only */}}
