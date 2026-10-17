// @vitest-environment node
import {expect,it,vi} from 'vitest'
import {createCanvas} from 'canvas'
import {createCellLayoutHost} from '../src/xlsx/view/renderViewport'
import {LayoutAdmissionError} from '../src/xlsx/view/textLayout'
it('refuses a conditional format before building its large section literal when no workspace is admitted',()=>{
 const code='[>0]"'+'q'.repeat(65536)+'"0',cell={ref:'A1',row:0,col:0,styleIndex:0,value:3,style:{numFmtId:164,formatCode:code}}
 let admissions=0;const builders:any[]=[],slice=String.prototype.slice
 const spy=vi.spyOn(String.prototype,'slice').mockImplementation(function(this:string,start=0,end?:number){
  if(String(this)===code && (end??code.length)-start>32768)builders.push({units:(end??code.length)-start,admissions})
  return slice.call(this,start,end)
 })
 try {
  const host=createCellLayoutHost(createCanvas(1,1).getContext('2d') as never)
  expect(()=>host.paintedText!(cell,64,{admit:()=>{admissions++;return false},release:()=>{},hold:()=>{},cacheChanged:()=>{}})).toThrow(LayoutAdmissionError)
  console.log('conditional builders',JSON.stringify({builders,admissions}))
  expect(builders.every(b=>b.admissions>0)).toBe(true)
 }finally{spy.mockRestore()}
})
