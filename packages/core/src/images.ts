// SPDX-License-Identifier: MIT
/** Bounded canonical payloads in immutable document snapshots. The header/marker
 * checks below are not a complete raster decoder. Chromium decodes only after
 * these checks; the container adapter never executes or expands image metadata.
 */
import {assert,assertId} from '@office-ex/contracts';
import {sha256} from './resources.js';
import type {ResourceRef} from './model.js';
const MAX_BYTES=8*1024*1024,MAX_PIXELS=40_000_000;
const B64='ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
export interface ImageInfo{mimeType:'image/png'|'image/jpeg';pixelWidth:number;pixelHeight:number}
export type ImageResource=ResourceRef&ImageInfo&{dataBase64:string};
function dimensions(width:number,height:number):void{assert(Number.isSafeInteger(width)&&Number.isSafeInteger(height)&&width>0&&height>0&&width<=16384&&height<=16384&&width*height<=MAX_PIXELS,'limit','Image pixel budget exceeded');}
export function encodeImageBytes(bytes:Uint8Array):string{
  assert(bytes instanceof Uint8Array&&bytes.length>0&&bytes.length<=MAX_BYTES,'limit','Image byte budget exceeded');
  const chunks:string[]=[];let chunk='';
  for(let i=0;i<bytes.length;i+=3){const a=bytes[i]!,b=bytes[i+1]??0,c=bytes[i+2]??0,n=a*65536+b*256+c;
    chunk+=B64[n>>>18]+B64[(n>>>12)&63]+(i+1<bytes.length?B64[(n>>>6)&63]:'=')+(i+2<bytes.length?B64[n&63]:'=');
    if(chunk.length>=16384){chunks.push(chunk);chunk='';}
  }chunks.push(chunk);return chunks.join('');
}
export function decodeImageBytes(value:string):Uint8Array{
  assert(typeof value==='string'&&value.length>0&&value.length%4===0&&value.length<=Math.ceil(MAX_BYTES/3)*4,'limit','Encoded image budget exceeded');
  const padding=value.endsWith('==')?2:value.endsWith('=')?1:0;
  const length=value.length/4*3-padding;assert(length>0&&length<=MAX_BYTES,'limit','Image byte budget exceeded');
  const result=new Uint8Array(length);
  const table=new Int16Array(128).fill(-1);for(let i=0;i<B64.length;i++)table[B64.charCodeAt(i)]=i;
  let out=0;
  for(let i=0;i<value.length;i+=4){
    const last=i+4===value.length,a=table[value.charCodeAt(i)]??-1,b=table[value.charCodeAt(i+1)]??-1;
    const c=last&&padding===2?0:table[value.charCodeAt(i+2)]??-1,d=last&&padding?0:table[value.charCodeAt(i+3)]??-1;
    assert(a>=0&&b>=0&&c>=0&&d>=0&&(!last||padding!==2||(b&15)===0)&&(!last||padding!==1||(c&3)===0),'invalid','Invalid or noncanonical image base64');
    const n=a*262144+b*4096+c*64+d;result[out++]=n>>>16;if(out<length)result[out++]=(n>>>8)&255;if(out<length)result[out++]=n&255;
  }return result;
}
const crcTable=new Uint32Array(256);for(let i=0;i<256;i++){let c=i;for(let j=0;j<8;j++)c=(c&1)?0xedb88320^(c>>>1):c>>>1;crcTable[i]=c>>>0;}
function crc(bytes:Uint8Array,start:number,end:number):number{let c=0xffffffff;for(let i=start;i<end;i++)c=crcTable[(c^bytes[i]!)&255]!^(c>>>8);return (c^0xffffffff)>>>0;}
export function inspectImage(bytes:Uint8Array):ImageInfo{
  assert(bytes instanceof Uint8Array&&bytes.length>=12&&bytes.length<=MAX_BYTES,'limit','Invalid image byte length');
  const v=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);
  if([137,80,78,71,13,10,26,10].every((b,i)=>bytes[i]===b)){
    let at=8,count=0,width=0,height=0,idat=false,ended=false;
    while(at<bytes.length){
      assert(++count<=4096&&at+12<=bytes.length,'invalid','Invalid PNG chunk envelope');
      const n=v.getUint32(at);assert(n<=bytes.length-at-12,'invalid','Truncated PNG chunk');
      const type=String.fromCharCode(...bytes.subarray(at+4,at+8)),end=at+8+n;
      assert(/^[A-Za-z]{4}$/.test(type)&&crc(bytes,at+4,end)===v.getUint32(end),'invalid','PNG chunk CRC or type mismatch');
      if(count===1){
        assert(type==='IHDR'&&n===13,'invalid','Missing PNG IHDR');width=v.getUint32(at+8);height=v.getUint32(at+12);dimensions(width,height);
        const depth=bytes[at+16]!,color=bytes[at+17]!;
        assert((color===0?[1,2,4,8,16].includes(depth):color===3?[1,2,4,8].includes(depth):[2,4,6].includes(color)&&[8,16].includes(depth))&&bytes[at+18]===0&&bytes[at+19]===0&&[0,1].includes(bytes[at+20]!),'unsupported','Unsupported PNG header');
      }else{
        assert(type!=='IHDR'&&!['acTL','fcTL','fdAT','iCCP','zTXt','iTXt','eXIf'].includes(type),'unsupported','Animated or compressed metadata PNG unsupported');
        if(type==='IDAT')idat=true;
        else if(type==='IEND'){assert(n===0&&idat&&end+4===bytes.length,'invalid','Invalid PNG ending');ended=true;}
        else assert(type==='PLTE'||type[0]===type[0]!.toLowerCase(),'unsupported','Unknown critical PNG chunk');
      }
      at=end+4;if(ended)break;
    }
    assert(ended,'invalid','PNG has no complete image');return {mimeType:'image/png',pixelWidth:width,pixelHeight:height};
  }
  if(bytes[0]===255&&bytes[1]===216){
    let at=2,width=0,height=0,scans=0,markers=0,scan=false;
    while(at<bytes.length){
      assert(++markers<=10000,'limit','JPEG marker budget exceeded');
      if(scan){
        while(at<bytes.length&&bytes[at]!==255)at++;
        assert(at<bytes.length,'invalid','Truncated JPEG scan');
      }else assert(bytes[at]===255,'invalid','Invalid JPEG marker');
      while(at<bytes.length&&bytes[at]===255)at++;
      assert(at<bytes.length,'invalid','Truncated JPEG marker');const marker=bytes[at++]!;
      if(scan&&(marker===0||marker>=208&&marker<=215))continue;
      scan=false;
      if(marker===217){assert(at===bytes.length&&width>0&&height>0&&scans>0,'invalid','Invalid JPEG ending');return {mimeType:'image/jpeg',pixelWidth:width,pixelHeight:height};}
      assert(marker!==216&&marker!==0&&at+2<=bytes.length,'invalid','Invalid JPEG segment');
      const size=v.getUint16(at);assert(size>=2&&size<=bytes.length-at,'invalid','Truncated JPEG segment');
      if([192,193,194,195,197,198,199,201,202,203,205,206,207].includes(marker)){
        assert((marker===192||marker===194)&&!width&&size>=8&&bytes[at+2]===8,'unsupported','Unsupported JPEG frame');
        height=v.getUint16(at+3);width=v.getUint16(at+5);dimensions(width,height);
        const components=bytes[at+7]!;assert((components===1||components===3)&&size===8+components*3,'unsupported','Unsupported JPEG components');
      }
      // EXIF orientation and ICC profiles are not mapped by this first profile.
      if(marker===225||marker===226)assert(false,'unsupported','JPEG EXIF/ICC metadata needs an orientation/color adapter');
      if(marker===218){assert(width>0&&++scans<=128,'invalid','Invalid JPEG scan');scan=true;}
      at+=size;
    }
    assert(false,'invalid','JPEG has no complete ending');
  }
  assert(false,'unsupported','Only PNG and JPEG image payloads are supported');
}
export function makeImageResource(bytes:Uint8Array):ImageResource{
  const info=inspectImage(bytes),digest=sha256(bytes);
  return {id:'img-'+digest,sha256:digest,byteLength:bytes.length,...info,dataBase64:encodeImageBytes(bytes)};
}
export function validateImageResource(resource:unknown):asserts resource is ImageResource{
  assert(resource!==null&&typeof resource==='object','invalid','Invalid image resource');
  const r=resource as ImageResource;assertId(r.id);
  const bytes=decodeImageBytes(r.dataBase64),info=inspectImage(bytes);
  assert(r.id==='img-'+r.sha256&&r.sha256===sha256(bytes)&&r.byteLength===bytes.length&&r.mimeType===info.mimeType&&r.pixelWidth===info.pixelWidth&&r.pixelHeight===info.pixelHeight,'invalid','Image resource digest, type or dimensions mismatch');
}
