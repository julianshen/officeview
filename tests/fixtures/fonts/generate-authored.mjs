// Author: Officeview contributors. Synthetic fixture bytes: CC0 1.0.
import { readFileSync, writeFileSync } from 'node:fs'
const src = readFileSync(new URL('../../../src/core/fonts/vendor/mtx.mjs', import.meta.url), 'utf8')
const { AHuff, setDistRange } = await import('data:text/javascript;base64,' + Buffer.from(src + '\nexport { AHuff, setDistRange };').toString('base64'))
function literals(bytes, rle = undefined) {
 const bits=rle === undefined ? [] : [rle ? 1 : 0]; const n=bytes.length
 for(let i=23;i>=0;i--) bits.push((n>>>i)&1)
 const {NUM_SYMS}=setDistRange(n); const h=new AHuff({},NUM_SYMS)
 for(const b of bytes) {
  let node=h.symbolIndex[b];const path=[]
  while(node!==1){const parent=h.tree[node].up;path.push(h.tree[parent].right===node?1:0);node=parent}
  bits.push(...path.reverse()); h.updateWeight(h.symbolIndex[b])
 }
 const out=new Uint8Array(Math.ceil(bits.length/8));bits.forEach((bit,i)=>out[i>>3]|=bit<<(7-(i&7)));return out
}
// Authored CTF with one empty glyph, no user font data. All table bytes authored here.
const tables=[['head',54],['maxp',32],['hmtx',4],['glyf',2],['loca',4]]
const size=12+tables.length*16+tables.reduce((n,t)=>n+t[1],0)
const ctf=new Uint8Array(size);const dv=new DataView(ctf.buffer);dv.setUint32(0,0x10000);dv.setUint16(4,tables.length)
let offset=12+16*tables.length
for(let i=0;i<tables.length;i++) {const [tag,length]=tables[i];const at=12+i*16;[...tag].forEach((c,j)=>ctf[at+j]=c.charCodeAt(0));dv.setUint32(at+8,offset);dv.setUint32(at+12,length)
 if(tag==='head'){dv.setUint32(offset,0x10000);dv.setUint32(offset+12,0x5f0f3cf5);dv.setUint16(offset+18,1000)}
 if(tag==='maxp'){dv.setUint32(offset,0x10000);dv.setUint16(offset+4,1)}
 if(tag==='hmtx')dv.setUint16(offset,500)
 offset+=length
}
function pack(stream0) {const streams=[stream0,new Uint8Array(),new Uint8Array()].map(b=>literals(b));const out=new Uint8Array(10+streams.reduce((n,b)=>n+b.length,0));out[0]=1;let pos=10;streams.forEach((b,i)=>{if(i>0){out[1+i*3]=(pos>>16)&255;out[2+i*3]=(pos>>8)&255;out[3+i*3]=pos&255}out.set(b,pos);pos+=b.length});return out}
writeFileSync(new URL('./authored-empty.mtx', import.meta.url),pack(ctf))
const hostile=ctf.slice(); new DataView(hostile.buffer).setUint32(24,0xffffffff)
writeFileSync(new URL('./authored-hostile-table.mtx', import.meta.url),pack(hostile))
const points=new Uint8Array(ctf.length+6);points.set(ctf);points.set([0,2,253,255,255,253,255,255],182)
writeFileSync(new URL('./authored-hostile-points.mtx', import.meta.url),pack(points))

function packRle(encoded) {
 const streams=[literals(encoded,true),literals(new Uint8Array(),false),literals(new Uint8Array(),false)]
 const out=new Uint8Array(10+streams.reduce((n,b)=>n+b.length,0));out[0]=2;let pos=10
 streams.forEach((b,i)=>{if(i>0){out[1+i*3]=(pos>>16)&255;out[2+i*3]=(pos>>8)&255;out[3+i*3]=pos&255}out.set(b,pos);pos+=b.length});return out
}
const expanded=new Uint8Array(1+128*3)
for(let i=0;i<128;i++) expanded.set([0,255,1],1+i*3)
writeFileSync(new URL('./authored-hostile-rle.mtx', import.meta.url),packRle(expanded))
writeFileSync(new URL('./authored-truncated-rle.mtx', import.meta.url),packRle(new Uint8Array([0,0])))

// CVT aliases deliberately underdeclare two-byte table sizes. Its count field
// expands each table to 131070 bytes; no third-party font bytes are present.
function cvtContainer(cvtCount, cvEntries, siblingBytes = 0) {
 const records=[['head',54],['maxp',32],['hmtx',4],...(siblingBytes ? [['pad ',siblingBytes]] : []),...Array.from({length:cvtCount},()=>['cvt ',2])]
 const start=12+records.length*16, cvtOffset=start+90+siblingBytes
 const data=new Uint8Array(cvtOffset+2+cvEntries), view=new DataView(data.buffer)
 view.setUint32(0,0x10000);view.setUint16(4,records.length);let atData=start
 records.forEach(([tag,length],i)=>{const at=12+i*16;[...tag].forEach((c,j)=>data[at+j]=c.charCodeAt(0));view.setUint32(at+8,tag==='cvt '?cvtOffset:atData);view.setUint32(at+12,length)
  if(tag==='head'){view.setUint32(atData,0x10000);view.setUint16(atData+18,1000)}
  if(tag==='maxp')view.setUint32(atData,0x10000)
  if(tag!=='cvt ')atData+=length
 });view.setUint16(cvtOffset,cvEntries);return data
}
writeFileSync(new URL('./authored-hostile-cvt-aggregate.mtx', import.meta.url),pack(cvtContainer(257,65535)))
writeFileSync(new URL('./authored-cvt-sibling.mtx', import.meta.url),pack(cvtContainer(1,256,3500)))
