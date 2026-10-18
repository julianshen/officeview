import {expect,it} from 'vitest'
import {writeFile} from 'node:fs/promises'
import {scanChartNumber} from '../src/xlsx/chart'
import {formatCommon} from '../src/xlsx/formula/format'
import {semanticsFor} from '../src/xlsx/view/semanticsFor'
const scan=(s:string)=>scanChartNumber(s.length,i=>s.charCodeAt(i))
const evidence:Record<string,unknown>={};const save=async()=>{void evidence;void writeFile}
// Excel-style numeric text (coordinator decision, r5): optional surrounding ASCII
// spaces, sign, digits with an optional point, optional exponent; nothing else.
const excelNumeric=(s:string)=>{const t=s.replace(/^ +| +$/g,'');if(!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t))return undefined;const n=Number(t);return Number.isFinite(n)?n:undefined}

// Adapted: radix text is not a number in Excel charts (it was F1 against a JS Number() oracle).
it('radix text (0x/0o/0b) is never a chart number, at any width',async()=>{
 const failures:any[]=[]
 for(const radix of [2,8,16]) for(let bits=1;bits<=53;bits++) for(const delta of[-1n,0n,1n]){
  const n=(1n<<BigInt(bits-1))+delta;if(n<0||n>BigInt(Number.MAX_SAFE_INTEGER))continue
  const prefix=radix===2?'0b':radix===8?'0o':'0x',s=prefix+n.toString(radix),actual=scan(s),expected=excelNumeric(s)
  if(!Object.is(actual,expected))failures.push({s,actual,expected})
 }
 evidence.radix=failures;await save();expect(failures).toEqual([])
})
it('radix digits cannot accept control characters as decimal digits',async()=>{
 const failures:any[]=[]
 for(const c of [0x10,0x11,0x12,0x13,0x14,0x15,0x16,0x17,0x18,0x19]) for(const prefix of ['0b','0o','0x']){
  const s=prefix+String.fromCharCode(c),actual=scan(s),expected=excelNumeric(s)
  if(!Object.is(actual,expected))failures.push({c,prefix,actual,expected})
 }
 evidence.controls=failures;await save();expect(failures).toEqual([])
})
it('seeded long decimal and halfway controls agree with Number',async()=>{
 let seed=11;const rnd=()=>{seed=(seed*48271)%2147483647;return seed/2147483647},failures:any[]=[]
 for(let i=0;i<3000;i++){
  const digits=Array.from({length:Math.floor(rnd()*1200)+1},()=>String(Math.floor(rnd()*10))).join(''),point=Math.floor(rnd()*(digits.length+1))
  const s=(rnd()<.5?'-':'')+digits.slice(0,point)+'.'+digits.slice(point)+'e'+Math.floor(rnd()*800-400)
  if(!Object.is(scan(s),excelNumeric(s)))failures.push({s,actual:scan(s),expected:excelNumeric(s)})
 }
 for(const numerator of [1n,(1n<<53n)-1n]){
  const d=(numerator*5n**1075n).toString(),half='0.'+'0'.repeat(1075-d.length)+d
  for(const sign of ['', '-'])for(const tail of ['', '0'.repeat(1200)+'1']){
   const text=sign+half+tail
   if(!Object.is(scan(text),excelNumeric(text)))failures.push({s:text,actual:scan(text),expected:excelNumeric(text)})
  }
 }
 evidence.decimalFailures=failures;await save();expect(failures).toEqual([])
})
it('capped grouped output preserves the prefix of the canonical formatted label',async()=>{
 const failures:any[]=[]
 for(const v of [1e299,1e298,-1e299]){
  const full=formatCommon(v,'#,##0',semanticsFor('1900')),capped=formatCommon(v,'#,##0',semanticsFor('1900'),{maxOutput:256,allocate:()=>{},release:()=>{}})
  expect(full.kind).toBe('text');expect(capped.kind).toBe('text')
  if(full.kind==='text'&&capped.kind==='text')failures.push({v,fullLength:full.text.length,capLength:capped.text.length,expected:full.text.slice(0,255),actual:capped.text.slice(0,255)})
 }
 evidence.grouped=failures;await save()
 expect(failures.every(x=>x.expected===x.actual)).toBe(true)
 expect(failures.every(x=>x.capLength<=256)).toBe(true)
})
it('ordinary numeric, text, literal and date cap controls preserve prefixes',()=>{
 for(const [value,fmt]of [[12.345,'0.00'],[1,'"hello"0'],[-1,'0'],['abc','0;0;0;'+ '@'.repeat(300)],[1,'yyyy-mm-dd'.repeat(40)]]as const){
  const full=formatCommon(value,fmt,semanticsFor('1900')),cut=formatCommon(value,fmt,semanticsFor('1900'),{maxOutput:256,allocate:()=>{},release:()=>{}})
  expect(full.kind).toBe('text');expect(cut.kind).toBe('text')
  if(full.kind==='text'&&cut.kind==='text')expect(cut.text.slice(0,255)).toBe(full.text.slice(0,255))
 }
})
