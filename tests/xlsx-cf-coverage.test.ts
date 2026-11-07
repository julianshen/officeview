import {expect,it,vi} from 'vitest'
import { collectCfCoverage } from '../src/components/xlsx/cfCoverage'
import type {SheetMetrics} from '../src/xlsx/view/contracts'
const metrics=(reverse=false)=>({rows:{visibleBetween:(a:number,b:number)=>({runs:[{start:Math.floor(a/20),end:Math.ceil(b/20)-1}]})},cols:{visibleBetween:(a:number,b:number)=>({runs:[{start:Math.floor(a/64),end:Math.ceil(b/64)-1}]})},order:{displayToSource:(r:number)=>reverse?99-r:r}} as unknown as SheetMetrics)
it('covers full edge and ring tile boxes in source coordinates and releases metadata',()=>{
 let held=0;const budget={reserve(n:number){held+=n},release(n:number){held-=n}}
 const c=collectCfCoverage(metrics(true),[{x:0,y:0,width:1024,height:1024}],budget)
 expect(c.rects.some(r=>r.r0<=60&&r.r1>=60&&r.c1>=15)).toBe(true)
 expect(c.contains({x:512,y:512,width:512,height:512})).toBe(true)
 expect(c.contains({x:1024,y:0,width:512,height:512})).toBe(false)
 expect(held).toBeGreaterThan(0);c.dispose();expect(held).toBe(0)
})
it('refuses overview coverage before enumerating individual source rows',()=>{
 const m=metrics();const map=vi.spyOn(m.order,'displayToSource')
 const budget={reserve(_n:number){},release(_n:number){}}
 expect(()=>collectCfCoverage(m,[{x:0,y:0,width:12800,height:18000}],budget)).toThrow(/coverage/)
 expect(map).not.toHaveBeenCalled()
})
