// SPDX-License-Identifier: MIT
/** Deterministic, font-free application branding generator. */
import assert from 'node:assert/strict';
import {mkdir,readFile,writeFile} from 'node:fs/promises';
import {deflateSync} from 'node:zlib';
import {createHash} from 'node:crypto';
import {pathToFileURL} from 'node:url';
import {resolve,join} from 'node:path';
const sizes=[16,24,32,48,64,128,256,512,1024];
const crcTable=Array.from({length:256},(_,n)=>{let c=n;for(let k=0;k<8;k++)c=c&1?0xedb88320^(c>>>1):c>>>1;return c>>>0;});
const crc32=data=>{let c=0xffffffff;for(const b of data)c=crcTable[(c^b)&255]^(c>>>8);return(c^0xffffffff)>>>0;};
function chunk(name,data){const tag=Buffer.from(name),header=Buffer.alloc(8),tail=Buffer.alloc(4);header.writeUInt32BE(data.length);tag.copy(header,4);tail.writeUInt32BE(crc32(Buffer.concat([tag,data])));return Buffer.concat([header,data,tail]);}
const rgb=hex=>[1,3,5].map(i=>parseInt(hex.slice(i,i+2),16));
function inPolygon(x,y,points){let yes=false;for(let i=0,j=points.length-1;i<points.length;j=i++){const a=points[i],b=points[j];if((a[1]>y)!==(b[1]>y)&&x<(b[0]-a[0])*(y-a[1])/(b[1]-a[1])+a[0])yes=!yes;}return yes;}
function insideTile(x,y,t){const edge=t.x+t.size,bottom=t.y+t.size;if(x<t.x||x>edge||y<t.y||y>bottom)return false;const dx=Math.max(t.x+t.radius-x,x-edge+t.radius,0),dy=Math.max(t.y+t.radius-y,y-bottom+t.radius,0);return dx*dx+dy*dy<=t.radius*t.radius;}
function png(mark,size){
 const raw=Buffer.alloc(size*(size*4+1)),t=mark.tile,base=rgb(t.color),glyphs=mark.glyphs.map(g=>({...g,rgb:rgb(g.color),bounds:[Math.min(...g.points.map(p=>p[0])),Math.min(...g.points.map(p=>p[1])),Math.max(...g.points.map(p=>p[0])),Math.max(...g.points.map(p=>p[1]))]}));
 const sub=size<=256?4:2,scale=mark.size/size,total=sub*sub;
 for(let y=0;y<size;y++)for(let x=0;x<size;x++){
  let alpha=0,r=0,g=0,b=0;
  for(let j=0;j<sub;j++)for(let i=0;i<sub;i++){
   const px=(x+(i+.5)/sub)*scale,py=(y+(j+.5)/sub)*scale;
   if(!insideTile(px,py,t))continue;
   let c=base;
   for(const shape of glyphs){const [l,u,rr,dd]=shape.bounds;if(px>=l&&px<=rr&&py>=u&&py<=dd&&inPolygon(px,py,shape.points))c=shape.rgb;}
   alpha++;r+=c[0];g+=c[1];b+=c[2];
  }
  const p=y*(size*4+1)+1+x*4;
  if(alpha){raw[p]=Math.round(r/alpha);raw[p+1]=Math.round(g/alpha);raw[p+2]=Math.round(b/alpha);raw[p+3]=Math.round(alpha*255/total);}
 }
 const ihdr=Buffer.alloc(13);ihdr.writeUInt32BE(size,0);ihdr.writeUInt32BE(size,4);ihdr[8]=8;ihdr[9]=6;
 return Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]),chunk('IHDR',ihdr),chunk('IDAT',deflateSync(raw,{level:9})),chunk('IEND',Buffer.alloc(0))]);
}
function ico(images){
 const chosen=[16,24,32,48,64,128,256],header=Buffer.alloc(6+16*chosen.length);header.writeUInt16LE(1,2);header.writeUInt16LE(chosen.length,4);let offset=header.length;
 for(const [i,n]of chosen.entries()){const p=6+i*16,data=images.get(n);header[p]=header[p+1]=n===256?0:n;header.writeUInt16LE(1,p+4);header.writeUInt16LE(32,p+6);header.writeUInt32LE(data.length,p+8);header.writeUInt32LE(offset,p+12);offset+=data.length;}
 return Buffer.concat([header,...chosen.map(n=>images.get(n))]);
}
function icns(images){
 const entries=[['icp4',16],['icp5',32],['icp6',64],['ic07',128],['ic08',256],['ic09',512],['ic10',1024],['ic11',32],['ic12',64],['ic13',256],['ic14',512]];
 const parts=entries.map(([type,size])=>{const data=images.get(size),head=Buffer.alloc(8);head.write(type);head.writeUInt32BE(data.length+8,4);return Buffer.concat([head,data]);});
 const h=Buffer.alloc(8);h.write('icns');h.writeUInt32BE(8+parts.reduce((sum,p)=>sum+p.length,0),4);return Buffer.concat([h,...parts]);
}
export async function buildBranding(target='dist/desktop/branding'){
 const mark=JSON.parse(await readFile('assets/branding/mark.json','utf8'));
 assert(mark.name==='Office EX'&&mark.size===1024&&mark.glyphs.length===2,'Unexpected mark');
 await mkdir(target,{recursive:true});
 const images=new Map(sizes.map(n=>[n,png(mark,n)])),t=mark.tile;
 const svg='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1024 1024"><title>Office EX</title><rect x="'+t.x+'" y="'+t.y+'" width="'+t.size+'" height="'+t.size+'" rx="'+t.radius+'" fill="'+t.color+'"/>'+mark.glyphs.map(g=>'<polygon fill="'+g.color+'" points="'+g.points.map(p=>p.join(',')).join(' ')+'"/>').join('')+'</svg>\n';
 const outputs=new Map([['office-ex.svg',Buffer.from(svg)],...sizes.map(n=>['office-ex-'+n+'.png',images.get(n)]),['office-ex.ico',ico(images)],['office-ex.icns',icns(images)]]);
 const files=[];for(const [name,data]of outputs){await writeFile(join(target,name),data);files.push({name,bytes:data.length,sha256:createHash('sha256').update(data).digest('hex')});}
 const report={schema:1,mark:'Office EX geometric monogram',sizes,files};
 await writeFile(join(target,'manifest.json'),JSON.stringify(report,null,2)+'\n');return report;
}
if(import.meta.url===pathToFileURL(resolve(process.argv[1]??'')).href){const result=await buildBranding(process.argv[2]);console.log('Built '+result.files.length+' branding assets');}
