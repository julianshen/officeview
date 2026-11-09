import { expect, test } from 'bun:test'
import JSZip from 'jszip'
import { OfficePackage } from '../../src/core/zip'
import { admit, projectPackage } from '../src/policy'
import { hiddenFixtures } from '../../scripts/stream-s0/hidden-fixtures'

async function packed(parts: Record<string,string>) { const zip=new JSZip();for(const [p,s] of Object.entries(parts))zip.file(p,s);return zip.generateAsync({type:'uint8array'}) }
test('admission refuses non-OOXML inputs on original bytes', async()=>{
 for(const b of [new TextEncoder().encode('{\\rtf1 hi}'),new Uint8Array([1,2,3]),await packed({'foo.xml':'<a/>'}),await packed({mimetype:'application/vnd.oasis.opendocument.text','word/document.xml':'<a/>'})])expect(admit(b)).toEqual({ok:false,code:'unsupported-format'})
 for(const f of await hiddenFixtures())expect(admit(f.withHidden)).toEqual({ok:true,format:f.kind})
})
test('full identity and visible projection never mutate the input package',async()=>{
 const f=(await hiddenFixtures())[0],pkg=await OfficePackage.load(f.withHidden),before=await Promise.all(pkg.paths().map(async p=>[p,await pkg.bytes(p)] as const))
 const full=await projectPackage(pkg,f.kind,'full');expect(full.pkg).toBe(pkg);expect(full.contentPolicyId).toBe('full@1');expect(full.diagnostics).toEqual([])
 await projectPackage(pkg,f.kind,'visible-only');for(const [p,b]of before)expect(await pkg.bytes(p)).toEqual(b)
})
test.each(['hidden="maybe"','xmlns:a14="urn:extension" a14:visibility="hidden"'])('unknown visibility strips the drawable, not only its flag: %s',async flag=>{
 const f=(await hiddenFixtures()).find(f=>f.id==='pptx-hidden-shape-cNvPr')!,zip=await JSZip.loadAsync(f.withHidden)
 const p='ppt/slides/slide1.xml';zip.file(p,(await zip.file(p)!.async('string')).replace('hidden="1"',flag));const pkg=await OfficePackage.load(await zip.generateAsync({type:'uint8array'}));const out=await projectPackage(pkg,'pptx','visible-only')
 expect(await out.pkg.text(p)).not.toContain(f.marker);expect(await out.pkg.text(p)).toContain(f.control);expect(out.diagnostics.some(d=>d.code==='hidden-unrecognised')).toBe(true)
})
test('unparseable XML fails the whole projection including unused parts',async()=>{
 const f=(await hiddenFixtures())[0],zip=await JSZip.loadAsync(f.withHidden);zip.file('customXml/broken.xml','<root><wrong></root>')
 const pkg=await OfficePackage.load(await zip.generateAsync({type:'uint8array'}));await expect(projectPackage(pkg,'xlsx','visible-only')).rejects.toThrow('unsupported-format')
})
test('namespace prefixes, mixed content and unknown visible extensions survive rewriting',async()=>{
 const f=(await hiddenFixtures()).find(f=>f.kind==='docx')!,zip=await JSZip.loadAsync(f.withHidden)
 zip.file('word/extension.xml','<x:root xmlns:x="urn:x" xmlns:y="urn:y" y:attr="a&amp;b">before <y:child/> after &lt;tag&gt;<![CDATA[& raw]]></x:root>')
 const pkg=await OfficePackage.load(await zip.generateAsync({type:'uint8array'})),out=await projectPackage(pkg,'docx','visible-only')
 expect(await out.pkg.text('word/extension.xml')).toBe(await pkg.text('word/extension.xml'))
})
test('removed comment parts also remove their own relationship parts',async()=>{
 const f=(await hiddenFixtures()).find(f=>f.id==='xlsx-comment')!,zip=await JSZip.loadAsync(f.withHidden)
 zip.file('xl/comments1.xml','<comments><commentList><comment ref="A1"><text><t>secret</t></text></comment></commentList></comments>')
 zip.file('xl/_rels/comments1.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="urn:test" Target="../customXml/item.xml"/></Relationships>')
 const pkg=await OfficePackage.load(await zip.generateAsync({type:'uint8array'})),out=await projectPackage(pkg,'xlsx','visible-only')
 expect(out.pkg.paths()).not.toContain('xl/_rels/comments1.xml.rels')
})
test('customXml data-store parts never travel, even with schema-level hidden flags',async()=>{
 const f=(await hiddenFixtures()).find(f=>f.kind==='pptx')!,zip=await JSZip.loadAsync(f.withHidden)
 zip.file('customXml/item1.xml','<xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:ma="urn:ma"><xsd:element name="x" ma:hidden="true"/></xsd:schema>')
 zip.file('customXml/_rels/item1.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
 const pkg=await OfficePackage.load(await zip.generateAsync({type:'uint8array'})),out=await projectPackage(pkg,'pptx','visible-only')
 expect(out.pkg.paths().filter(p=>p.startsWith('customXml/'))).toEqual([])
 expect(out.diagnostics.some(d=>d.part.startsWith('customXml/'))).toBe(false)
})
