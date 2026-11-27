// SPDX-License-Identifier: MIT
/** Closed DrawingML chart graph: plain one-cell anchors, local column
 * references, clustered column/standard line/pie, stock axes and legend.
 * Reject unsupported chart markup, external dependencies, cached/opaque
 * values or custom appearance instead of losing it during an edit. */
import {OfficeError} from '@office-ex/contracts';
import type {DeepReadonly,SheetChart,SheetRange} from '@office-ex/core';
import {columnName,parseFormula} from '@office-ex/sheets';
import type {Element} from '@xmldom/xmldom';
import {parseXml,childElements} from './xml.js';
import type {OpcPackage} from './opc.js';
import type {ContainerOptions} from './limits.js';
const X='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',A='http://schemas.openxmlformats.org/drawingml/2006/main',C='http://schemas.openxmlformats.org/drawingml/2006/chart',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',PKG='http://schemas.openxmlformats.org/package/2006/relationships',XMLNS='http://www.w3.org/2000/xmlns/';
const DRAWING='application/vnd.openxmlformats-officedocument.drawing+xml',CHART='application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const enc=new TextEncoder();
const xml=(s:string)=>enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+s);
function need(v:unknown,m:string):asserts v{if(!v)throw new OfficeError('unsupported','XLSX chart: '+m);}
function get(e:Element,k:string){return e.hasAttribute(k)?e.getAttribute(k)??undefined:undefined;}
function attrs(e:Element,keys:readonly string[]=[]){for(let i=0;i<e.attributes.length;i++){const a=e.attributes.item(i)!;if(a.namespaceURI===XMLNS)continue;need(keys.includes(a.namespaceURI===R?'r:'+a.localName:a.namespaceURI?'?'+a.name:a.name),'Unknown chart attribute '+a.name);}}
function kids(e:Element,ns:string,names:readonly string[]):Element[]{for(let n=e.firstChild;n;n=n.nextSibling)need(n.nodeType===1||n.nodeType===3&&!(n.nodeValue??'').trim(),'Unexpected chart text');const c=childElements(e);need(c.every(x=>x.namespaceURI===ns&&names.includes(x.localName??'')),'Unsupported chart element in '+e.localName);return c;}
function one(c:Element[],n:string,required=true){const a=c.filter(x=>x.localName===n);need(a.length<2&&(!required||a.length===1),'Unexpected '+n+' cardinality');return a[0];}
function leaf(e:Element,keys:readonly string[]=[]){attrs(e,keys);for(let n=e.firstChild;n;n=n.nextSibling)need(n.nodeType===3||n.nodeType===4,'Unexpected nested chart content');return e.textContent??'';}
function empty(e:Element,keys:readonly string[]=[]){need(leaf(e,keys)==='','Unexpected chart text');}
function number(v:string|undefined,min=0,max=2147483647){need(v!==undefined&&/^\d+$/.test(v),'Invalid chart integer');const n=Number(v);need(Number.isSafeInteger(n)&&n>=min&&n<=max,'Chart integer outside budget');return n;}
function val(e:Element,value:string){empty(e,['val']);need(get(e,'val')===value,'Unsupported '+e.localName);}
function root(bytes:Uint8Array,ns:string,name:string,options:ContainerOptions){const e=parseXml(bytes,options).documentElement;need(e?.namespaceURI===ns&&e.localName===name,'Unexpected chart root');attrs(e);return e;}
function ref(e:Element,sheet:string,category=false,single=false):SheetRange{
 attrs(e);const wrapper=kids(e,C,category?['numRef','strRef']:single?['strRef']:['numRef']);
 need(wrapper.length===1,'Missing chart reference');attrs(wrapper[0]!);const f=one(kids(wrapper[0]!,C,['f']),'f')!;
 const ast=parseFormula('='+leaf(f));
 need(ast.kind==='reference'||ast.kind==='range'&&!ast.axis,'Chart requires explicit cell range');
 const a=ast.kind==='reference'?ast:ast.start,b=ast.kind==='reference'?ast:ast.end;
 need(a.sheet?.toLowerCase()===sheet.toLowerCase()&&(!b.sheet||b.sheet.toLowerCase()===sheet.toLowerCase())&&a.column===b.column&&b.row>=a.row&&b.row-a.row<10000&&(!single||a.row===b.row),'Chart source must be a local column');
 return {row:a.row,column:a.column,rowCount:b.row-a.row+1,columnCount:1};
}
function defaultLine(e:Element){attrs(e);const line=one(kids(e,A,['ln']),'ln')!;attrs(line);val(one(kids(line,A,['prstDash']),'prstDash')!,'solid');}
function axis(e:Element,type:'catAx'|'valAx',id:string,cross:string){
 attrs(e);const c=kids(e,C,['axId','scaling','axPos','majorGridlines','majorTickMark','minorTickMark','crossAx','lblOffset']);
 val(one(c,'axId')!,id);val(one(c,'crossAx')!,cross);
 const scale=one(c,'scaling')!;attrs(scale);val(one(kids(scale,C,['orientation']),'orientation')!,'minMax');
 const pos=one(c,'axPos')!;empty(pos,['val']);need(['l','b'].includes(get(pos,'val')??''),'Custom axis position');
 for(const k of ['majorTickMark','minorTickMark']){const x=one(c,k,false);if(x)val(x,'none');}
 const grid=one(c,'majorGridlines',false);if(grid){need(type==='valAx','Category gridline profile');empty(grid);}
 const label=one(c,'lblOffset',false);if(label){need(type==='catAx','Value label offset');val(label,'100');}
}
function readChart(bytes:Uint8Array,sheet:string,options:ContainerOptions):Omit<SheetChart,'anchor'>{
 const top=root(bytes,C,'chartSpace',options),chart=one(kids(top,C,['chart']),'chart')!;attrs(chart);
 const ch=kids(chart,C,['plotArea','legend','plotVisOnly','dispBlanksAs']);
 const legend=one(ch,'legend',false);if(legend){attrs(legend);val(one(kids(legend,C,['legendPos']),'legendPos')!,'r');}
 const visible=one(ch,'plotVisOnly',false);if(visible)val(visible,'1');
 const blank=one(ch,'dispBlanksAs',false);if(blank)val(blank,'gap');
 const plot=one(ch,'plotArea')!;attrs(plot);const p=kids(plot,C,['barChart','lineChart','pieChart','catAx','valAx']);
 const types=p.filter(e=>['barChart','lineChart','pieChart'].includes(e.localName??''));need(types.length===1,'Mixed chart types');
 const t=types[0]!,kind=t.localName==='barChart'?'column':t.localName==='lineChart'?'line':'pie';attrs(t);
 const c=kids(t,C,kind==='column'?['barDir','grouping','ser','gapWidth','axId']:kind==='line'?['grouping','ser','axId']:['varyColors','ser','firstSliceAng']);
 if(kind==='column'){val(one(c,'barDir')!,'col');val(one(c,'grouping')!,'clustered');const gap=one(c,'gapWidth',false);if(gap)val(gap,'150');}
 else if(kind==='line')val(one(c,'grouping')!,'standard');
 else{const vary=one(c,'varyColors',false),angle=one(c,'firstSliceAng',false);if(vary)val(vary,'1');if(angle)val(angle,'0');}
 const axes=c.filter(e=>e.localName==='axId');if(kind==='pie')need(!axes.length&&p.length===1,'Pie axes');else{
  need(axes.length===2&&p.length===3,'Missing chart axes');for(const a of axes)empty(a,['val']);const first=get(axes[0]!,'val')!,second=get(axes[1]!,'val')!;
  number(first,1);number(second,1);need(first!==second,'Duplicate chart axis');axis(one(p,'catAx')!,'catAx',first,second);axis(one(p,'valAx')!,'valAx',second,first);
 }
 const seriesNodes=c.filter(e=>e.localName==='ser');need(seriesNodes.length>0&&seriesNodes.length<=(kind==='pie'?1:16),'Chart series budget');
 let categories:SheetRange|undefined;
 const series=seriesNodes.map((s,i)=>{
  attrs(s);const nodes=kids(s,C,['idx','order','tx','spPr','marker','cat','val']);val(one(nodes,'idx')!,String(i));val(one(nodes,'order')!,String(i));
  const sp=one(nodes,'spPr',false);if(sp)defaultLine(sp);
  const marker=one(nodes,'marker',false);if(marker){need(kind==='line','Unsupported marker');attrs(marker);const m=kids(marker,C,['symbol','spPr']);val(one(m,'symbol')!,'none');const sp=one(m,'spPr',false);if(sp)defaultLine(sp);}
  const cat=ref(one(nodes,'cat')!,sheet,true),values=ref(one(nodes,'val')!,sheet);
  if(categories)need(JSON.stringify(cat)===JSON.stringify(categories),'Different chart category ranges');else categories=cat;
  need(cat.rowCount===values.rowCount,'Chart range length mismatch');const tx=one(nodes,'tx',false),label=tx?ref(tx,sheet,false,true):undefined;
  return {values,...(label?{label:{row:label.row,column:label.column}}:{})};
 });
 return {kind,categories:categories!,series};
}
export function chartRelationFile(path:string){const i=path.lastIndexOf('/');return path.slice(0,i+1)+'_rels/'+path.slice(i+1)+'.rels';}
export interface SheetChartGraph {charts:SheetChart[];parts:string[];typedParts:string[]}
export function readChartGraph(opc:OpcPackage,worksheet:string,drawing:Element,sheet:string,options:ContainerOptions,expected:Set<string>):SheetChartGraph{
 empty(drawing,['r:id']);const refs=opc.relationships.get(worksheet)??[];
 need(refs.length===1&&refs[0]!.id===drawing.getAttributeNS(R,'id')&&refs[0]!.type===R+'/drawing'&&refs[0]!.mode==='Internal'&&refs[0]!.targetPart&&!refs[0]!.fragment,'Unexpected worksheet drawing relationship');
 const path=refs[0]!.targetPart!,relFile=chartRelationFile(worksheet),drawingRel=chartRelationFile(path);
 const parts=[relFile,path,drawingRel],typedParts=[path];
 const claim=(p:string,type?:string)=>{need(!expected.has(p)&&opc.parts.has(p)&&(!type||opc.contentTypes.get(p)===type),'Missing or aliased chart part');expected.add(p);};
 claim(relFile);claim(path,DRAWING);claim(drawingRel);
 const top=root(opc.parts.get(path)!,X,'wsDr',options),anchors=kids(top,X,['oneCellAnchor']),rels=opc.relationships.get(path)??[];
 need(anchors.length>0&&anchors.length<=32&&anchors.length===rels.length,'Chart drawing count');const seen=new Set<string>();
 const charts=anchors.map(a=>{
  attrs(a);const c=kids(a,X,['from','ext','graphicFrame','clientData']),from=one(c,'from')!,ext=one(c,'ext')!,frame=one(c,'graphicFrame')!;
  attrs(from);const f=kids(from,X,['col','colOff','row','rowOff']),anchor={row:number(leaf(one(f,'row')!),0,1048575)+1,column:number(leaf(one(f,'col')!),0,16383)+1,offsetX:number(leaf(one(f,'colOff')!),0,10000000),offsetY:number(leaf(one(f,'rowOff')!),0,10000000),width:number(get(ext,'cx'),1,50000000),height:number(get(ext,'cy'),1,50000000)};
  empty(ext,['cx','cy']);empty(one(c,'clientData')!);attrs(frame);
  const fc=childElements(frame);need(fc.length===3&&fc[0]!.namespaceURI===X&&fc[0]!.localName==='nvGraphicFramePr'&&fc[1]!.namespaceURI===X&&fc[1]!.localName==='xfrm'&&fc[2]!.namespaceURI===A&&fc[2]!.localName==='graphic','Unsupported chart frame');attrs(fc[0]!);const nv=kids(fc[0]!,X,['cNvPr','cNvGraphicFramePr']);const identity=one(nv,'cNvPr')!;empty(identity,['id','name']);number(get(identity,'id'),1);empty(one(nv,'cNvGraphicFramePr')!);empty(fc[1]!);
  attrs(fc[2]!);const graphic=one(kids(fc[2]!,A,['graphicData']),'graphicData')!;attrs(graphic,['uri']);need(get(graphic,'uri')===C,'Non-chart drawing');
  const chart=one(kids(graphic,C,['chart']),'chart')!;empty(chart,['r:id']);const id=chart.getAttributeNS(R,'id'),rel=rels.find(r=>r.id===id);
  need(id&&!seen.has(id)&&rel?.type===R+'/chart'&&rel.mode==='Internal'&&rel.targetPart&&!rel.fragment,'Unexpected chart relationship');seen.add(id);
  const chartPath=rel.targetPart!;claim(chartPath,CHART);need(!(opc.relationships.get(chartPath)??[]).length,'Chart external dependency');parts.push(chartPath);typedParts.push(chartPath);
  return {...readChart(opc.parts.get(chartPath)!,sheet,options),anchor};
 });
 return {charts,parts,typedParts};
}
function esc(s:string){return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');}
function formula(sheet:string,r:DeepReadonly<SheetRange>){const a='$'+columnName(r.column)+'$'+r.row,b='$'+columnName(r.column)+'$'+(r.row+r.rowCount-1);return esc("'"+sheet.replaceAll("'","''")+"'!"+a+(a===b?'':':'+b));}
function chartXml(c:DeepReadonly<SheetChart>,sheet:string){
 const tag=c.kind==='column'?'barChart':c.kind==='line'?'lineChart':'pieChart';
 const style=c.kind==='column'?'<barDir val="col"/><grouping val="clustered"/>':c.kind==='line'?'<grouping val="standard"/>':'<varyColors val="1"/>';
 const series=c.series.map((s,i)=>'<ser><idx val="'+i+'"/><order val="'+i+'"/>'+(s.label?'<tx><strRef><f>'+formula(sheet,{...s.label,rowCount:1,columnCount:1})+'</f></strRef></tx>':'')+'<cat><strRef><f>'+formula(sheet,c.categories)+'</f></strRef></cat><val><numRef><f>'+formula(sheet,s.values)+'</f></numRef></val></ser>').join('');
 const axis=(type:string,id:number,cross:number)=>'<'+type+'><axId val="'+id+'"/><scaling><orientation val="minMax"/></scaling><axPos val="'+(type==='catAx'?'b':'l')+'"/>'+(type==='valAx'?'<majorGridlines/>':'')+'<majorTickMark val="none"/><minorTickMark val="none"/><crossAx val="'+cross+'"/>'+(type==='catAx'?'<lblOffset val="100"/>':'')+'</'+type+'>';
 return xml('<chartSpace xmlns="'+C+'"><chart><plotArea><'+tag+'>'+style+series+(c.kind==='column'?'<gapWidth val="150"/>':'')+(c.kind==='pie'?'<firstSliceAng val="0"/>':'<axId val="10"/><axId val="100"/>')+'</'+tag+'>'+(c.kind==='pie'?'':axis('catAx',10,100)+axis('valAx',100,10))+'</plotArea><legend><legendPos val="r"/></legend><plotVisOnly val="1"/><dispBlanksAs val="gap"/></chart></chartSpace>');
}
export function createChartGraph(worksheet:string,sheet:string,charts:readonly DeepReadonly<SheetChart>[],parts:Map<string,Uint8Array>,allocate:(prefix:string)=>string):{drawing:string;types:{path:string;type:string}[]}{
 if(!charts.length)return {drawing:'',types:[]};
 const path=allocate('xl/drawings/oex-drawing-'),types=[{path,type:DRAWING}],rels:string[]=[];
 const nodes=charts.map((c,i)=>{
  const chartPath=allocate('xl/charts/oex-chart-');types.push({path:chartPath,type:CHART});parts.set(chartPath,chartXml(c,sheet));
  const id='rId'+(i+1);rels.push('<Relationship Id="'+id+'" Type="'+R+'/chart" Target="/'+chartPath+'"/>');
  const a=c.anchor;return '<oneCellAnchor><from><col>'+(a.column-1)+'</col><colOff>'+a.offsetX+'</colOff><row>'+(a.row-1)+'</row><rowOff>'+a.offsetY+'</rowOff></from><ext cx="'+a.width+'" cy="'+a.height+'"/><graphicFrame><nvGraphicFramePr><cNvPr id="'+(i+1)+'" name="Chart '+(i+1)+'"/><cNvGraphicFramePr/></nvGraphicFramePr><xfrm/><a:graphic><a:graphicData uri="'+C+'"><c:chart r:id="'+id+'"/></a:graphicData></a:graphic></graphicFrame><clientData/></oneCellAnchor>';
 });
 parts.set(path,xml('<wsDr xmlns="'+X+'" xmlns:a="'+A+'" xmlns:c="'+C+'" xmlns:r="'+R+'">'+nodes.join('')+'</wsDr>'));
 parts.set(chartRelationFile(path),xml('<Relationships xmlns="'+PKG+'">'+rels.join('')+'</Relationships>'));
 parts.set(chartRelationFile(worksheet),xml('<Relationships xmlns="'+PKG+'"><Relationship Id="rIdOexDrawing" Type="'+R+'/drawing" Target="/'+path+'"/></Relationships>'));
 return {drawing:'<drawing xmlns:r="'+R+'" r:id="rIdOexDrawing"/>',types};
}
