// @vitest-environment node
import { expect, it } from 'vitest'
import { formatExtended, formatColor } from '../src/xlsx/cellFormat'
for (const fn of ['extended','color'] as const) it(`${fn} admits its live section-array slots before truncating to four`,()=>{
 const code='[>0]0'+';'.repeat(8192),limit=40000
 let held=0,sections=0,reservedAtPeak=0
 const push=Array.prototype.push
 Array.prototype.push = function(this:unknown[],...items:unknown[]){
  for(const item of items)if(item&&typeof item==='object'&&'body'in item&&'unsupported'in item){sections++;reservedAtPeak=held}
  return push.apply(this,items)
 }
 const allocation={allocate:(n:number)=>{if(held+n>limit)throw new Error('refused');held+=n},release:(n:number)=>{held-=n}}
 try {
  const semantics={dateSystem:'1900',unicode:{version:2,source:'test'},locale:'en-US',timeZone:'UTC',epochNowMs:0} as never
  if(fn==='extended')formatExtended(3,code,semantics,String,allocation);else formatColor(3,code,allocation)
  console.log('section bound',JSON.stringify({fn,limit,codeUnits:code.length,sections,slotBytesAlone:sections*8,reservedAtPeak,heldAtEnd:held}))
  expect(reservedAtPeak).toBeGreaterThanOrEqual(sections*8)
 }finally{Array.prototype.push=push}
})
