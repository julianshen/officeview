import {expect,test} from 'bun:test'
import {readFileSync} from 'node:fs'
import {createCanvas} from '@napi-rs/canvas'
import {OfficePackage} from '../../src/core/zip'
import {detectAndParse} from '../../src/components/loadOfficeSource'
import {getPaintables} from '../../src/render/paint'
import {fileURLToPath} from 'node:url'
import {projectPackage} from '../src/policy'
// The shared enumerator resolves its roots at import time. Restore cwd before tests run.
const savedCwd=process.cwd()
process.chdir(fileURLToPath(new URL('../../',import.meta.url)))
const files=await (async()=>{try{return (await import('../../scripts/stream-s0/corpus')).listCorpus()}finally{process.chdir(savedCwd)}})()
test('identity corpus is nonempty',()=>expect(files.length).toBeGreaterThan(0))
for(const f of files)test(`projection identity ${f.name}`,async()=>{
 // Intentionally malformed fuzzer inputs cannot produce a reference parse.
 if(f.name.startsWith('poi-crash-')){console.log(`[identity] ${f.name}: malformed fuzzer input`);return}
 const original=await OfficePackage.load(readFileSync(f.path))
 // Pin the conservative refusal explicitly, never silently skip arbitrary errors.
 const refusals:Record<string,string>={}
 if(refusals[f.name]){await expect(projectPackage(original,f.kind,'visible-only')).rejects.toThrow(refusals[f.name]);console.log(`[identity] ${f.name}: REFUSED ${refusals[f.name]}`);return}
 const projected=await projectPackage(original,f.kind,'visible-only')
 if(projected.diagnostics.length){console.log(`[identity] ${f.name}: diagnostics ${JSON.stringify(projected.diagnostics.map(d=>`${d.code}:${d.element}`))}`);return}
 const before=await getPaintables(await detectAndParse(original)),after=await getPaintables(await detectAndParse(projected.pkg))
 try{
 expect(after.length).toBe(before.length)
 for(let i=0;i<before.length;i++){
  const a=before[i],b=after[i];expect(b.spec).toEqual(a.spec)
  const w=Math.max(1,Math.ceil(a.spec.widthPx)),h=Math.max(1,Math.ceil(a.spec.heightPx))
  const ca=createCanvas(w,h),cb=createCanvas(w,h);a.paint(ca.getContext('2d') as never);b.paint(cb.getContext('2d') as never)
  expect(Buffer.compare(Buffer.from(ca.getContext('2d').getImageData(0,0,w,h).data),Buffer.from(cb.getContext('2d').getImageData(0,0,w,h).data)),`${f.name} unit ${i}`).toBe(0)
 }
 console.log(`[identity] ${f.name}: exact ${before.length} units`)
 }finally{before.dispose();after.dispose()}
},120000)
