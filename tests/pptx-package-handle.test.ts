/** S0-4: the PPTX package handle (slideshow spec 3, 8): a lease that lives exactly as long as the document. */
import { setFlagsFromString } from 'node:v8'
import { runInNewContext } from 'node:vm'
import JSZip from 'jszip'
import { describe, expect, it, vi } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parsePptx, pptxPackageHandle } from '../src/pptx/parse'
import { readPartBounded } from '../src/pptx/appearance/read'
import { HandlePackage } from '../src/pptx/appearance/handle-package'
import { drawingPartContext } from '../src/drawing/parts'

setFlagsFromString('--expose_gc')
const gc = runInNewContext('gc') as () => void
const task = () => new Promise<void>(resolve => setTimeout(resolve, 0))
async function collect() { for (let i = 0; i < 8; i++) { await task(); gc() } await task() }

const slide = (text: string) => `<p:sld><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="t"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`

/** Three slide parts listed out of file order, with one dangling sldId that parsePptx skips. */
async function deckBytes(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('ppt/presentation.xml', '<p:presentation><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="c"/><p:sldId id="257" r:id="missing"/><p:sldId id="258" r:id="a"/><p:sldId id="259" r:id="b"/></p:sldIdLst></p:presentation>')
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships><Relationship Id="a" Type="slide" Target="slides/slide1.xml"/><Relationship Id="b" Type="slide" Target="slides/slide2.xml"/><Relationship Id="c" Type="slide" Target="slides/slide3.xml"/></Relationships>')
  for (const n of [1, 2, 3]) zip.file(`ppt/slides/slide${n}.xml`, slide(`S${n}`))
  return zip.generateAsync({ type: 'uint8array' })
}
const cacheOf = (pkg: OfficePackage) => (pkg as unknown as { cache: Map<string, unknown> }).cache
/** The drawing parser's package-keyed read state, as the handle sees it after the parse. */
const drawingState = (pkg: OfficePackage) => {
  const c = drawingPartContext(pkg)
  return { relationships: c.relationships.size, relationshipNodes: c.relationshipNodes.size, placements: c.placements.size, chart: c.chart }
}
const EMPTY_DRAWING = { relationships: 0, relationshipNodes: 0, placements: 0, chart: undefined }

describe('PPTX package handle', () => {
  it('parsePptx attaches a handle whose slideParts follow deck order', async () => {
    const pkg = await OfficePackage.load(await deckBytes())
    const doc = await parsePptx(pkg)
    const handle = pptxPackageHandle(doc)
    expect(handle).toBeDefined()
    expect(handle!.pkg).toBe(pkg)
    expect(doc.slides).toHaveLength(3)
    expect(handle!.slideParts).toEqual(['ppt/slides/slide3.xml', 'ppt/slides/slide1.xml', 'ppt/slides/slide2.xml'])
    // slideParts[i] is the part of doc.slides[i] (the dangling sldId is skipped in both).
    expect(doc.slides.map(s => s.shapes[0].textBody?.paragraphs[0].runs[0].text)).toEqual(['S3', 'S1', 'S2'])
  })

  it('after parsePptx the package read cache is empty and handle reads go through entry() only', async () => {
    const pkg = await OfficePackage.load(await deckBytes())
    const handle = pptxPackageHandle(await parsePptx(pkg))!
    expect(cacheOf(pkg).size).toBe(0)
    const entry = vi.spyOn(pkg, 'entry'), text = vi.spyOn(pkg, 'text'), bytes = vi.spyOn(pkg, 'bytes')
    const read = await readPartBounded(handle.pkg, handle.slideParts[0], 4 * 1024 * 1024) as Uint8Array
    expect(new TextDecoder().decode(read)).toContain('S3')
    expect(entry).toHaveBeenCalledWith('ppt/slides/slide3.xml')
    expect(text).not.toHaveBeenCalled()
    expect(bytes).not.toHaveBeenCalled()
    expect(cacheOf(pkg).size).toBe(0)
  })

  it('the lease keeps no drawing-parser read state for the package', async () => {
    const pkg = await OfficePackage.load(await deckBytes())
    const chart = { profile: 'desktop' } as never
    const doc = await parsePptx(pkg, { chartContext: chart })
    const handle = pptxPackageHandle(doc)!
    expect(drawingState(handle.pkg)).toEqual(EMPTY_DRAWING)
    // What stays is the audit record: counters, and diagnostics the document itself already holds.
    const owned = new Set(doc.slides.flatMap(s => s.diagnostics ?? []))
    expect(drawingPartContext(handle.pkg).diagnostics.every(d => owned.has(d))).toBe(true)
  })

  it('a document parsed through a HandlePackage keeps none of its parse products', async () => {
    const source = pptxPackageHandle(await parsePptx(await OfficePackage.load(await deckBytes())))!
    const over = HandlePackage.over(source)
    const doc = await parsePptx(over)
    const handle = pptxPackageHandle(doc)!
    expect(handle.pkg).toBe(over)
    expect(handle.slideParts).toEqual(source.slideParts)
    const caches = over as unknown as { reads: Map<string, unknown>; trees: Map<string, unknown> }
    expect(caches.reads.size).toBe(0)
    expect(caches.trees.size).toBe(0)
    expect(drawingState(over)).toEqual(EMPTY_DRAWING)
    expect(cacheOf(source.pkg).size).toBe(0)
  })

  it('a cloned document has no handle and stringifies', async () => {
    const doc = await parsePptx(await OfficePackage.load(await deckBytes()))
    expect(pptxPackageHandle(doc)).toBeDefined()
    const clone = structuredClone(doc)
    expect(() => JSON.stringify(doc)).not.toThrow()
    expect(JSON.stringify(clone)).toBe(JSON.stringify(doc))
    expect(pptxPackageHandle(clone)).toBeUndefined()
    expect(pptxPackageHandle(JSON.parse(JSON.stringify(doc)))).toBeUndefined()
  })

  it('the handle does not outlive the document', async () => {
    // Built in a helper so no local in this body pins the package or the document.
    const holder: { doc?: object } = {}
    const ref = await (async () => {
      const pkg = await OfficePackage.load(await deckBytes())
      holder.doc = await parsePptx(pkg)
      return new WeakRef(pkg)
    })()
    await collect()
    // Control: the live document's handle keeps the package.
    expect(ref.deref()).toBeDefined()
    delete holder.doc
    await collect()
    expect(ref.deref()).toBeUndefined()
  })
})
