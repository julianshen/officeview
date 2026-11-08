import JSZip from 'jszip'
import {expect,test} from 'bun:test'
import {createCanvas} from '@napi-rs/canvas'
import {OfficePackage} from '../../src/core/zip'
import {detectAndParse} from '../../src/components/loadOfficeSource'
import {getPaintables} from '../../src/render/paint'
import {hiddenFixtures} from '../../scripts/stream-s0/hidden-fixtures'
import {projectPackage} from '../src/policy'
const fixtures=await hiddenFixtures()
for(const f of fixtures)test(`T2 ${f.id}: visible control, no marker in paint or model`,async()=>{
 // The two supplied comment fixtures contain only their visible control.
 // Add actual marker-bearing comment parts so these cases are not vacuous.
 const zip=await JSZip.loadAsync(f.withHidden)
 if(f.id==='xlsx-comment')zip.file('xl/comments1.xml',`<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><commentList><comment ref="A1" authorId="0"><text><t>${f.marker}</t></text></comment></commentList></comments>`)
 if(f.id==='docx-comment')zip.file('word/comments.xml',`<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0"><w:p><w:r><w:t>${f.marker}</w:t></w:r></w:p></w:comment></w:comments>`)
 const input=await OfficePackage.load(await zip.generateAsync({type:'uint8array'}))
 expect((await Promise.all(input.paths(p=>p.endsWith('.xml')).map(p=>input.text(p)))).join('')).toContain(f.marker)
 const out=await projectPackage(input,f.kind,'visible-only'),doc=await detectAndParse(out.pkg)
 expect((await Promise.all(out.pkg.paths(p=>p.endsWith('.xml')).map(p=>out.pkg.text(p)))).join('')).not.toContain(f.marker)
 const seen=new WeakSet<object>(),leaks:string[]=[]
 function walk(v:unknown){if(typeof v==='string'){if(v.includes(f.marker))leaks.push(v)}else if(v&&typeof v==='object'&&!seen.has(v)){seen.add(v);for(const a of Object.values(v))walk(a)}}walk(doc);expect(leaks).toEqual([])
 const texts:string[]=[],units=await getPaintables(doc)
 try{for(const p of units){const ctx=createCanvas(Math.max(1,Math.ceil(p.spec.widthPx)),Math.max(1,Math.ceil(p.spec.heightPx))).getContext('2d')
 const proxy=new Proxy(ctx,{get(t,k){const v=Reflect.get(t,k,t);return typeof v==='function'?(...a:unknown[])=>{if(k==='fillText'||k==='strokeText')texts.push(String(a[0]));return Reflect.apply(v,t,a)}:v},set(t,k,v){return Reflect.set(t,k,v,t)}})
 p.paint(proxy as unknown as CanvasRenderingContext2D)}
 expect(texts.join('')).toContain(f.control);expect(texts.join('')).not.toContain(f.marker);expect(texts.some(t=>t.includes(f.marker))).toBe(false)
 console.log(`[T2] ${f.id}: control=true paint-leak=false model-exposure=0`)
 }finally{units.dispose()}
},30000)
