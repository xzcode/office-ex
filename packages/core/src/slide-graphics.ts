// SPDX-License-Identifier: MIT
import type {Shape} from './model.js';
import type {DeepReadonly} from './session.js';
/** Includes every node exactly once, in source paint order. */
export function* walkSlideShapes<T extends Shape|DeepReadonly<Shape>>(shapes:readonly T[]):Generator<T>{
 for(const shape of shapes){yield shape;if(shape.type==='group')yield* walkSlideShapes(shape.children as readonly T[]);}
}

import {assert, ROTATION_UNITS_PER_DEGREE} from '@office-ex/contracts';
import type {Geometry,SlideConnection,SlideConnections} from './model.js';
type ReadShape=DeepReadonly<Shape>;
export type SlidePoint={x:number;y:number};
export function isConnectorTarget(shape:ReadShape|undefined):shape is ReadShape&{type:'rect'|'text'}{
 return !!shape&&(shape.type==='rect'||shape.type==='text')&&shape.geometry.width>0&&shape.geometry.height>0;
}
/** Rectangular preset connection sites; account for local flips and rotation. */
export function connectionPoint(shape:ReadShape,site:SlideConnection['site']):SlidePoint{
 assert(isConnectorTarget(shape),'unsupported','Unsupported connection target');
 const g=shape.geometry,a=g.rotation/ROTATION_UNITS_PER_DEGREE*Math.PI/180,c=Math.cos(a),s=Math.sin(a),p=[{x:0,y:-g.height/2},{x:-g.width/2,y:0},{x:0,y:g.height/2},{x:g.width/2,y:0}][site]!;
 const x=shape.flipH?-p.x:p.x,y=shape.flipV?-p.y:p.y;
 return {x:Math.round(g.x+g.width/2+c*x-s*y),y:Math.round(g.y+g.height/2+s*x+c*y)};
}
export function nearestSlideConnections(a:ReadShape,b:ReadShape):SlideConnections{
 assert(isConnectorTarget(a)&&isConnectorTarget(b)&&a.id!==b.id,'invalid','Select two different rectangular shapes');
 let best=Infinity,first:SlideConnection['site']=0,last:SlideConnection['site']=0;
 for(const start of [0,1,2,3] as const)for(const end of [0,1,2,3] as const){const p=connectionPoint(a,start),q=connectionPoint(b,end),d=Math.hypot(p.x-q.x,p.y-q.y);if(d<best){best=d;first=start;last=end;}}
 return {start:{shapeId:a.id,site:first},end:{shapeId:b.id,site:last}};
}
export function connectorGeometry(connections:DeepReadonly<SlideConnections>,shapes:ReadonlyMap<string,ReadShape>):{geometry:Geometry;flipH:boolean;flipV:boolean}{
 const a=shapes.get(connections.start.shapeId),b=shapes.get(connections.end.shapeId);assert(isConnectorTarget(a)&&isConnectorTarget(b),'unsupported','Missing or unsupported connection target');
 const p=connectionPoint(a,connections.start.site),q=connectionPoint(b,connections.end.site);
 return {geometry:{x:Math.min(p.x,q.x),y:Math.min(p.y,q.y),width:Math.abs(p.x-q.x),height:Math.abs(p.y-q.y),rotation:0},flipH:p.x>q.x,flipV:p.y>q.y};
}
export function resolveSlideConnector<T extends ReadShape>(shape:T,siblings:ReadonlyMap<string,ReadShape>):T{
 if(shape.type!=='line'||!shape.connections)return shape;
 return {...shape,...connectorGeometry(shape.connections,siblings)};
}
