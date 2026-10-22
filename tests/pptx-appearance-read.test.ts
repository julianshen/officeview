/** S0-4: bounded uncached part reads, HandlePackage and slide-size validation (slideshow spec 6.1, 7, 8). */
import JSZip from 'jszip'
import { describe, expect, it, vi } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { PackageError } from '../src/core/package-error'
import { parsePptx, pptxPackageHandle } from '../src/pptx/parse'
import { readPartBounded, readXmlBounded, readXmlOrderedBounded, retainedArchiveBytes } from '../src/pptx/appearance/read'
import { HANDLE_AGGREGATE_BYTES, HANDLE_PART_BYTES, HandlePackage } from '../src/pptx/appearance/handle-package'
import { SLIDE_ASPECT_MAX, SLIDE_EMU_MAX, SLIDE_EMU_MIN, validateSlideSize } from '../src/pptx/appearance/size'
import { buildZip, type BuildEntry } from './helpers/zip-builder'

const MiB = 1024 * 1024
const PRESENTATION = '<p:presentation><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="s"/></p:sldIdLst></p:presentation>'
const RELS = '<Relationships><Relationship Id="s" Type="slide" Target="slides/slide1.xml"/></Relationships>'
const SLIDE = '<p:sld><p:cSld><p:spTree/></p:cSld></p:sld>'

async function deck(extra: Record<string, string> = {}) {
  const zip = new JSZip()
  zip.file('ppt/presentation.xml', PRESENTATION)
  zip.file('ppt/_rels/presentation.xml.rels', RELS)
  zip.file('ppt/slides/slide1.xml', SLIDE)
  for (const [path, value] of Object.entries(extra)) zip.file(path, value)
  const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
  const doc = await parsePptx(pkg)
  return { pkg, doc, handle: pptxPackageHandle(doc)! }
}
const cacheOf = (pkg: OfficePackage) => (pkg as unknown as { cache: Map<string, unknown> }).cache
const limitError = (error: unknown) => error instanceof PackageError && error.code === 'resource-limit'

describe('readPartBounded', () => {
  it('readPartBounded refuses on the chunk crossing the cap even when the declared size lies', async () => {
    const big = 'a'.repeat(256 * 1024)
    // The central directory declares 64 B; the entry inflates to 256 KiB.
    const forged = buildZip([{ name: 'x.xml', data: big, uncompressedSize: 64 } satisfies BuildEntry]).bytes
    const pkg = await OfficePackage.load(forged)
    expect((pkg.entry('x.xml') as unknown as { _data: { uncompressedSize: number } })._data.uncompressedSize).toBe(64)
    const seen: number[] = []
    const error = await readPartBounded(pkg, 'x.xml', 40 * 1024, { admit: n => { seen.push(n); return true } }).catch(e => e)
    expect(limitError(error)).toBe(true)
    // Refused on the crossing chunk: admitted chunks stay at or under the cap, and the crossing chunk is never admitted.
    expect(seen.reduce((a, b) => a + b, 0)).toBeLessThanOrEqual(40 * 1024)
    expect(seen.length).toBeGreaterThan(0)
    // Honest sizes under the cap read in full, absent parts are undefined.
    const honest = await OfficePackage.load(buildZip([{ name: 'y.xml', data: '<y/>' }]).bytes)
    expect(new TextDecoder().decode(await readPartBounded(honest, 'y.xml', 4) as Uint8Array)).toBe('<y/>')
    await expect(readPartBounded(honest, 'y.xml', 3)).rejects.toSatisfy(limitError)
    expect(await readPartBounded(honest, 'nope.xml', 4)).toBeUndefined()
  })

  it('admit returning false on the third chunk aborts the read and retains no chunk', async () => {
    const pkg = await OfficePackage.load(buildZip([{ name: 'm.bin', data: 'b'.repeat(200 * 1024) }]).bytes)
    let calls = 0
    const admit = vi.fn(() => ++calls < 3)
    const error = await readPartBounded(pkg, 'm.bin', 1024 * MiB, { admit, output: 'chunks' }).catch(e => e)
    expect(limitError(error)).toBe(true)
    // Let any scheduled inflate ticks run: the stream was aborted, so admit is never asked again.
    for (let i = 0; i < 5; i++) await new Promise(resolve => setTimeout(resolve, 0))
    expect(admit).toHaveBeenCalledTimes(3)
    // Chunks output: the whole part as a chunk list, without a final concatenation.
    const chunks = await readPartBounded(pkg, 'm.bin', 1024 * MiB, { output: 'chunks' }) as Uint8Array[]
    expect(Array.isArray(chunks)).toBe(true)
    expect(chunks.length).toBeGreaterThan(1)
    expect(chunks.reduce((n, c) => n + c.length, 0)).toBe(200 * 1024)
  })

  it('an aborted signal refuses before and during the read', async () => {
    const pkg = await OfficePackage.load(buildZip([{ name: 'm.bin', data: 'c'.repeat(200 * 1024) }]).bytes)
    const before = new AbortController(); before.abort()
    await expect(readPartBounded(pkg, 'm.bin', MiB, { signal: before.signal })).rejects.toBeDefined()
    const during = new AbortController()
    let calls = 0
    await expect(readPartBounded(pkg, 'm.bin', MiB, { signal: during.signal, admit: () => { if (++calls === 2) during.abort(); return true } })).rejects.toBeDefined()
    expect(calls).toBe(2)
  })

  it('an abort during admission of the final chunk refuses in both output modes', async () => {
    const pkg = await OfficePackage.load(buildZip([{ name: 'h.txt', data: 'hello' }]).bytes)
    for (const output of ['bytes', 'chunks'] as const) {
      const controller = new AbortController()
      let calls = 0
      const result = await readPartBounded(pkg, 'h.txt', 1024, { output, signal: controller.signal, admit: () => { calls++; controller.abort(); return true } }).then(r => ({ r }), e => ({ e }))
      expect(calls, output).toBe(1)
      expect('e' in result, output).toBe(true)
    }
    // Aborted after the last chunk but before end: still refused.
    const late = new AbortController()
    const bytes = readPartBounded(pkg, 'h.txt', 1024, { signal: late.signal, admit: () => { queueMicrotask(() => late.abort()); return true } })
    await expect(bytes).rejects.toBeDefined()
  })

  it('readXmlBounded and readXmlOrderedBounded parse under the cap', async () => {
    const pkg = await OfficePackage.load(buildZip([{ name: 'p.xml', data: '<a:r><a:b/><a:c v="1"/></a:r>' }]).bytes)
    const tree = await readXmlBounded(pkg, 'p.xml', 1024)
    expect(tree).toBeDefined()
    expect(await readXmlOrderedBounded(pkg, 'p.xml', 1024)).toBeDefined()
    await expect(readXmlBounded(pkg, 'p.xml', 4)).rejects.toSatisfy(limitError)
    expect(await readXmlBounded(pkg, 'q.xml', 1024)).toBeUndefined()
  })
})

describe('HandlePackage', () => {
  it('HandlePackage reads never populate the source package cache', async () => {
    const { pkg, handle } = await deck({ 'ppt/media/a.bin': 'abc' })
    expect(cacheOf(pkg).size).toBe(0)
    const text = vi.spyOn(pkg, 'text'), bytes = vi.spyOn(pkg, 'bytes'), xml = vi.spyOn(pkg, 'xml'), ordered = vi.spyOn(pkg, 'xmlOrdered')
    const over = HandlePackage.over(handle)
    expect(await over.text('ppt/presentation.xml')).toBe(PRESENTATION)
    expect(await over.xml('ppt/presentation.xml')).toBeDefined()
    expect(await over.xmlOrdered(handle.slideParts[0])).toBeDefined()
    expect(new TextDecoder().decode(await over.bytes('ppt/media/a.bin'))).toBe('abc')
    expect(await over.boundedText('ppt/media/a.bin', 3)).toBe('abc')
    expect(await over.text('nope.xml')).toBeUndefined()
    // Existing parse helpers run over the handle unchanged (a second parse reads the same deck).
    const again = await parsePptx(over)
    expect(again.slides).toHaveLength(1)
    expect(cacheOf(pkg).size).toBe(0)
    for (const spy of [text, bytes, xml, ordered]) expect(spy).not.toHaveBeenCalled()
  })

  it('HandlePackage has, entry and paths delegate synchronously; text over the per-part or aggregate cap refuses', async () => {
    const { handle } = await deck({ 'ppt/big1.xml': 'x'.repeat(17 * MiB), 'ppt/big2.xml': 'y'.repeat(17 * MiB) })
    const pkg = handle.pkg
    const over = HandlePackage.over(handle)
    const has = vi.spyOn(pkg, 'has'), entry = vi.spyOn(pkg, 'entry'), paths = vi.spyOn(pkg, 'paths')
    expect(over.has('ppt/presentation.xml')).toBe(true)
    expect(over.has('nope')).toBe(false)
    expect(over.entry('ppt/presentation.xml')).toBe(pkg.entry('ppt/presentation.xml'))
    expect(over.paths().sort()).toEqual(pkg.paths().sort())
    expect(over.paths(p => p.endsWith('.rels'))).toEqual(['ppt/_rels/presentation.xml.rels'])
    expect(has).toHaveBeenCalledWith('ppt/presentation.xml')
    expect(entry).toHaveBeenCalledWith('ppt/presentation.xml')
    expect(paths).toHaveBeenCalled()

    // Default part cap 4 MiB; a caller's tighter bytes() limit keeps the legacy message its callers branch on.
    expect(HANDLE_PART_BYTES).toBe(4 * MiB)
    expect(HANDLE_AGGREGATE_BYTES).toBe(32 * MiB)
    await expect(over.text('ppt/big1.xml')).rejects.toSatisfy(limitError)
    await expect(over.xml('ppt/big1.xml')).rejects.toSatisfy(limitError)
    await expect(over.boundedText('ppt/presentation.xml', 4)).rejects.toSatisfy(limitError)
    await expect(over.bytes('ppt/presentation.xml', 4)).rejects.toThrow(/^Part byte limit exceeded/)
    const tiny = HandlePackage.over(handle, { partBytes: 8 })
    await expect(tiny.text('ppt/presentation.xml')).rejects.toSatisfy(limitError)

    // Aggregate: two 17 MiB parts each fit a 20 MiB part cap, but not the 32 MiB instance total.
    const wide = HandlePackage.over(handle, { partBytes: 20 * MiB })
    expect((await wide.text('ppt/big1.xml'))!.length).toBe(17 * MiB)
    expect((await wide.text('ppt/big1.xml'))!.length).toBe(17 * MiB) // a cache hit is not charged twice
    await expect(wide.text('ppt/big2.xml')).rejects.toSatisfy(limitError)
    // A fresh instance has its own total.
    expect((await HandlePackage.over(handle, { partBytes: 20 * MiB }).text('ppt/big2.xml'))!.length).toBe(17 * MiB)
    expect(cacheOf(pkg).size).toBe(0)
  }, 30_000)
})

describe('HandlePackage textBytes', () => {
  it('textBytes caps text and XML reads below a picture-sized partBytes (S0-7 enrichment reader)', async () => {
    const { handle } = await deck({ 'ppt/media/big.bin': 'z'.repeat(5 * MiB), 'ppt/big.xml': '<a>' + 'x'.repeat(5 * MiB) + '</a>' })
    const reader = HandlePackage.over(handle, { partBytes: HANDLE_AGGREGATE_BYTES, textBytes: HANDLE_PART_BYTES })
    expect((await reader.bytes('ppt/media/big.bin'))!.length).toBe(5 * MiB)
    await expect(reader.text('ppt/big.xml')).rejects.toSatisfy(limitError)
    await expect(reader.xmlOrdered('ppt/big.xml')).rejects.toSatisfy(limitError)
    // a part already read whole as bytes is still refused as text
    await expect(reader.text('ppt/media/big.bin')).rejects.toSatisfy(limitError)
    expect(await reader.text('ppt/presentation.xml')).toBe(PRESENTATION)
  }, 30_000)
})

describe('HandlePackage.dropCache', () => {
  it('drops reads and trees, and a read in flight never republishes into the dropped cache', async () => {
    const { handle } = await deck()
    const over = HandlePackage.over(handle)
    const caches = over as unknown as { reads: Map<string, unknown>; trees: Map<string, unknown> }
    await over.xml('ppt/presentation.xml')
    expect([caches.reads.size, caches.trees.size]).toEqual([1, 1])
    over.dropCache()
    expect([caches.reads.size, caches.trees.size]).toEqual([0, 0])
    // A tighter-cap read publishes on completion; dropped while in flight, it must not come back.
    const inflight = over.bytes(handle.slideParts[0], 1024)
    over.dropCache()
    expect(await inflight).toBeDefined()
    await new Promise(resolve => setTimeout(resolve, 0))
    expect(caches.reads.size).toBe(0)
  })
})

describe('retainedArchiveBytes', () => {
  it('retainedArchiveBytes equals the sum of compressed entry sizes of a builder deck', async () => {
    const layout = buildZip([
      { name: 'ppt/presentation.xml', data: PRESENTATION },
      { name: 'ppt/_rels/presentation.xml.rels', data: RELS },
      { name: 'ppt/slides/slide1.xml', data: SLIDE },
      { name: 'ppt/media/image1.bin', data: 'z'.repeat(10_000) },
      { name: 'ppt/media/stored.bin', data: 'stored-bytes', method: 0 },
    ])
    const doc = await parsePptx(await OfficePackage.load(layout.bytes))
    const expected = layout.entries.reduce((n, e) => n + e.dataLength, 0)
    expect(expected).toBeGreaterThan(0)
    expect(retainedArchiveBytes(pptxPackageHandle(doc)!)).toBe(expected)
  })

  it('retainedArchiveBytes is a payload lower bound: the backing buffer can outsize it', async () => {
    // JSZip keeps each entry's compressed bytes as a view of the whole input, so headers, the central
    // directory and a comment stay alive too. The metric counts compressed payload only (plan formula).
    const layout = buildZip([
      { name: 'ppt/presentation.xml', data: PRESENTATION },
      { name: 'ppt/_rels/presentation.xml.rels', data: RELS },
      { name: 'ppt/slides/slide1.xml', data: SLIDE },
    ], { comment: new Uint8Array(16 * 1024) })
    const handle = pptxPackageHandle(await parsePptx(await OfficePackage.load(layout.bytes)))!
    const payload = retainedArchiveBytes(handle)
    const view = (handle.pkg.entry('ppt/slides/slide1.xml') as unknown as { _data: { compressedContent: Uint8Array } })._data.compressedContent
    expect(payload).toBe(layout.entries.reduce((n, e) => n + e.dataLength, 0))
    expect(view.buffer.byteLength).toBeGreaterThanOrEqual(layout.bytes.byteLength)
    expect(view.buffer.byteLength).toBeGreaterThan(payload + 16 * 1024)
  })
})

describe('validateSlideSize', () => {
  it('validateSlideSize accepts the range ends and rejects zero, negative, NaN, Infinity, one EMU outside each end and 1:21', () => {
    expect([SLIDE_EMU_MIN, SLIDE_EMU_MAX, SLIDE_ASPECT_MAX]).toEqual([914_400, 51_206_400, 20])
    expect(validateSlideSize(SLIDE_EMU_MIN, SLIDE_EMU_MIN)).toEqual({ ok: true })
    expect(validateSlideSize(SLIDE_EMU_MAX, SLIDE_EMU_MAX)).toEqual({ ok: true })
    expect(validateSlideSize(9_144_000, 6_858_000)).toEqual({ ok: true })
    expect(validateSlideSize(SLIDE_EMU_MIN * 20, SLIDE_EMU_MIN)).toEqual({ ok: true }) // exactly 1 : 20
    expect(validateSlideSize(SLIDE_EMU_MIN, SLIDE_EMU_MIN * 20)).toEqual({ ok: true })
    const bad: Array<[number, number]> = [
      [0, 6_858_000], [9_144_000, 0], [-9_144_000, 6_858_000], [9_144_000, -1],
      [NaN, 6_858_000], [9_144_000, NaN], [Infinity, 6_858_000], [9_144_000, -Infinity],
      [SLIDE_EMU_MIN - 1, SLIDE_EMU_MIN], [SLIDE_EMU_MIN, SLIDE_EMU_MIN - 1],
      [SLIDE_EMU_MAX + 1, SLIDE_EMU_MAX], [SLIDE_EMU_MAX, SLIDE_EMU_MAX + 1],
      [SLIDE_EMU_MIN * 21, SLIDE_EMU_MIN], [SLIDE_EMU_MIN, SLIDE_EMU_MIN * 21],
    ]
    for (const [w, h] of bad) {
      const result = validateSlideSize(w, h)
      expect(result.ok, `${w} x ${h}`).toBe(false)
      if (!result.ok) {
        expect(result.diagnostic.kind).toBe('slide-size')
        expect(result.diagnostic.message.length).toBeGreaterThan(0)
      }
    }
  })
})
