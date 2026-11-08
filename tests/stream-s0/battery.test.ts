import { expect, it } from 'vitest'
import { createHmac } from 'node:crypto'
import * as message from '../../scripts/stream-s0/mark/message'
import { codeword40, extendedCodeword } from '../../scripts/stream-s0/mark/bch'
import { basis } from '../../scripts/stream-s0/mark/basis'
import { embed } from '../../scripts/stream-s0/mark/embed'
import { median } from '../../scripts/stream-s0/mark/attacks'
import { detect, scoreCapture } from '../../scripts/stream-s0/mark/detect'
import { DEFAULT_PARAMS as p } from '../../scripts/stream-s0/mark/params'
const key = new Uint8Array(32).fill(17)
it('maps the payload through HMAC with explicit generation and first 40 bits', () => {
  const bytes = Buffer.from([0,0,0,7,0,0,23,4,0])
  const bits = codeword40(23,4)
  for(let i=0;i<40;i++) bytes[4+(i>>3)] = (bytes[4+(i>>3)]! & ~(1<<(7-i%8))) | bits[i]!<<(7-i%8)
  const digest=createHmac('sha256',key).update(bytes).digest()
  expect(message.messageFor(key,7,23,4)).toEqual(Uint8Array.from({length:40},(_,i)=>(digest[i>>3]!>>(7-i%8))&1))
  expect(message.messageFor(key,8,23,4)).not.toEqual(message.messageFor(key,7,23,4))
})
it('audit insertion redraws a colliding salt and refuses exhaustion', () => {
  const used=new Set<string>([Array.from(message.messageFor(key,7,23,4)).join('')])
  const salts=[4,5]
  const row=message.insertMessage(key,7,23,used,()=>salts.shift()!)
  expect(row.salt8).toBe(5); expect(used.size).toBe(2)
  expect(()=>message.insertMessage(key,7,23,used,()=>4)).toThrow(/exhaust/)
})
const ref={width:256,height:256,data:new Uint8ClampedArray(256*256*4).fill(128)}
for(let i=3;i<ref.data.length;i+=4)ref.data[i]=255
const chips=basis(key,256,256,p)
it('two candidates above absolute threshold never attribute even with a large margin', () => {
  const code=extendedCodeword(codeword40(1,0)), other=code.slice(); other[0]*=-1
  expect(detect(embed(ref,chips,code,p),ref,chips,[code,other],p,{abs:7,margin:0}).kind).toBe('collusion-suspected')
})
it('K5 median loses the linear XOR innocent framing score under keyed messages', () => {
  const tokens=[1,2,4,8,16], innocent=31
  const linear=[...tokens,innocent].map(t=>extendedCodeword(codeword40(t,0)))
  const capture=median(linear.slice(0,5).map(c=>embed(ref,chips,c,p)))
  expect(scoreCapture(capture,ref,chips,linear,p).scores[5]).toBeGreaterThan(7)
  const keyed=[...tokens,innocent].map(t=>extendedCodeword(message.messageFor(key,1,t,0)))
  const marked=median(keyed.slice(0,5).map(c=>embed(ref,chips,c,p)))
  expect(scoreCapture(marked,ref,chips,keyed,p).scores[5]).toBeLessThan(7)
})

import { runBattery, calibrate } from '../../scripts/stream-s0/t7-t9-battery'
it('tau chosen meets fitted 1e-6 and the whole-candidate union bound', () => {
  const fit=calibrate(Array.from({length:200},(_,i)=>({best:2+i/100,margin:i/1000})),65536)
  expect(fit.tau.abs).toBeGreaterThanOrEqual(Math.sqrt(2*Math.log(65536/1e-6)))
  expect(fit.fittedAbsTail).toBeLessThanOrEqual(1.001e-6)
  expect(fit.fittedMarginTail).toBeLessThanOrEqual(1.001e-6)
})
it('battery runs on a 2-page set with M=256 and emits all attack rows', async () => {
  const result=await runBattery([{id:'synthetic-a',kind:'docx',bitmap:ref},{id:'synthetic-b',kind:'xlsx',bitmap:ref}],{candidateCount:256,nullCount:6,tokens:1,params:[p],progress:()=>{}})
  expect(result.t7.sweep).toHaveLength(1)
  expect(result.t9.fullScoredDecisions).toBe(6)
  expect(result.t9.nulls.length).toBeGreaterThanOrEqual(6)
  expect(new Set(result.t9.nulls.map(x=>x.basisKey))).toHaveProperty('size',result.t9.nulls.length)
  expect(result.t9.survival.map(x=>x.attack)).toEqual(['default','q95','q90','q80','q70','up2','down0.5','q70-down0.5','crop5','crop25','crop25-down0.5','crop25-up2'])
  expect(result.t9.collusion).toHaveLength(8)
  expect(result.t9.collusion.every(x=>Number.isFinite(x.innocentRate))).toBe(true)
},60000)
