import { existsSync } from 'node:fs'
import { it, expect, vi } from 'vitest'
import { buildPptxDeck } from '../src/testdata/ooxml-builders'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { openShowPackage, readShowIndex } from '../src/slideshow/read/deck'
import type { Lease, SlideShowWorkspace, ShowSlide, SlideTimingReader } from '../src/slideshow/types'
import type { PptxDocument } from '../src/pptx/types'
export function workspace(limit = 512 * 1024 * 1024): SlideShowWorkspace {
  let used = 0
  return {
    limit,
    used: () => used,
    reserveTerminal: () => undefined,
    lease: (bytes, tag) => {
      if (used + bytes > limit) return undefined
      used += bytes
      let live = true
      let n = bytes
      const l: Lease = {
        get bytes() {
          return n
        },
        tag,
        grow(b) {
          if (!live || used + b > limit) return false
          used += b
          n += b
          return true
        },
        shrink(b) {
          used -= b
          n -= b
        },
        transfer: () => l,
        release() {
          if (live) {
            used -= n
            n = 0
            live = false
          }
        },
      }
      return l
    },
  }
}
async function setup(estimate?: (s: ShowSlide) => number, w = workspace()) {
  const handle = await openShowPackage(await buildPptxDeck({ slides: [{ shapes: [], notes: [['hello']] }, { shapes: [] }] }))
  const doc: PptxDocument = { slideWidthEmu: 9144000, slideHeightEmu: 6858000, images: [], slides: [0, 1].map(index => ({ index, widthEmu: 9144000, heightEmu: 6858000, shapes: [] })) }
  const sink = createDiagnosticSink(), index = await readShowIndex(doc, handle, sink)
  const { createSlideMetadataStore } = await import('../src/slideshow/read/store')
  const timing = vi.fn((_xml: unknown) => ({ timing: undefined, builds: [] }))
  const store = createSlideMetadataStore({ doc, handle, index, workspace: w, sink, timing, resolver: () => ({ shape: () => 'slide', paragraphs: () => [] }), estimate })
  return { store, handle, w, timing, sink }
}
it('acquire reads each slide part once and hands the same XML to the timing reader', async () => {
  const { store, handle, timing, w } = await setup()
  const read = vi.spyOn(handle.pkg, 'entry')
  const [a, b] = await Promise.all([store.acquire(0), store.acquire(0)])
  expect(read.mock.calls.filter(([p]) => p === 'ppt/slides/slide1.xml')).toHaveLength(1)
  expect(timing).toHaveBeenCalledTimes(1)
  expect(timing.mock.calls[0][0]).toHaveProperty('cSld')
  expect(a.slide).toBe(b.slide)
  expect(a.slide.notes[0].text).toBe('hello')
  a.release()
  b.release()
  store.close()
  expect(w.used()).toBe(0)
})
it('evict with no handles releases the charge; a revisited slide is re-read', async () => {
  const { store, handle, w } = await setup()
  const read = vi.spyOn(handle.pkg, 'entry')
  const a = await store.acquire(0)
  a.release()
  store.evict(0)
  expect(w.used()).toBe(0);
  (await store.acquire(0)).release()
  expect(read.mock.calls.filter(([p]) => p === 'ppt/slides/slide1.xml')).toHaveLength(2)
  store.close()
})
it('evicting a slide whose handle an outgoing transition still holds keeps the charge until release; re-acquire reads a fresh entry', async () => {
  const { store, w } = await setup(() => 1000)
  const a = await store.acquire(0)
  store.evict(0)
  expect(w.used()).toBe(1000)
  const b = await store.acquire(0)
  expect(b.slide).not.toBe(a.slide)
  expect(w.used()).toBe(2000)
  a.release()
  a.release()
  expect(w.used()).toBe(1000)
  b.release()
  store.close()
  expect(w.used()).toBe(0)
})
it('a preview handle held across eviction stays valid and is charged once', async () => {
  const { store, w } = await setup(() => 1000)
  const a = await store.acquire(0), b = await store.acquire(0)
  store.evict(0)
  a.release()
  expect(b.slide.notes[0].text).toBe('hello')
  expect(w.used()).toBe(1000)
  store.close()
  expect(w.used()).toBe(1000)
  b.release()
  expect(w.used()).toBe(0)
})
it('metadata at 64 MiB is accepted; one byte more evicts a neighbour, then makes the slide static', async () => {
  const cap = 64 * 1024 * 1024
  const { store, w } = await setup(s => s.index === 0 ? cap : cap + 1)
  const a = await store.acquire(0)
  expect(w.used()).toBe(cap)
  a.release()
  const b = await store.acquire(1)
  expect(store.resident()).not.toContain(0)
  expect(b.slide.timing).toEqual({ refused: 'metadata-limit' })
  expect(w.used()).toBeLessThan(cap)
  b.release()
  store.close()
})
it('attach that would cross the 64 MiB aggregate or the workspace is refused before allocation', async () => {
  const { store, w } = await setup(() => 1024)
  const a = await store.acquire(0)
  expect(a.attach(64 * 1024 * 1024 - 1024, 'compiled')).toBe(true)
  expect(a.attach(1, 'overflow')).toBe(false)
  expect(w.used()).toBe(64 * 1024 * 1024)
  a.release()
  store.close()
  expect(w.used()).toBe(0)
  const q = await setup(() => 1024, workspace(90 * 1024 * 1024))
  const b = await q.store.acquire(0)
  const pressure = q.w.lease(q.w.limit - 1024, 'pressure')!
  expect(b.attach(1, 'overflow')).toBe(false)
  pressure.release()
  b.release()
  q.store.close()
})
it('close cancels an in-flight acquire without charging', async () => {
  const { store, w } = await setup()
  const p = store.acquire(0)
  store.close()
  await expect(p).rejects.toMatchObject({ code: 'cancelled' })
  expect(w.used()).toBe(0)
})
async function fixture(spec: Parameters<typeof buildPptxDeck>[0], edit?: (zip: import('jszip')) => Promise<void>, opts: { limit?: number; timing?: SlideTimingReader } = {}) {
  const { default: JSZip } = await import('jszip')
  let data = await buildPptxDeck(spec)
  if (edit) {
    const zip = await JSZip.loadAsync(data)
    await edit(zip)
    data = await zip.generateAsync({ type: 'uint8array' })
  }
  const handle = await openShowPackage(data), sink = createDiagnosticSink(), w = workspace(opts.limit)
  const doc: PptxDocument = { slideWidthEmu: 9144000, slideHeightEmu: 6858000, images: [], slides: spec.slides.map((_, index) => ({ index, widthEmu: 9144000, heightEmu: 6858000, shapes: [] })) }
  const index = await readShowIndex(doc, handle, sink)
  const { createSlideMetadataStore } = await import('../src/slideshow/read/store')
  const store = createSlideMetadataStore({ doc, handle, index, sink, workspace: w, timing: opts.timing, resolver: () => ({ shape: () => 'slide', paragraphs: () => [] }) })
  return { store, handle, sink, w, index, doc }
}
it('notes keep level and bullet; 65536 chars of notes are kept and 65537 truncate with a diagnostic; a 1 MiB notes part is read and 1 MiB + 1 is refused', async () => {
  const withBuilds: SlideTimingReader = () => ({ timing: undefined, builds: [{ spid: '2', mode: 'whole', ranges: [], animBg: false, body: 'b' }] })
  for (const n of [65536, 65537]) {
    const q = await fixture({ slides: [{ shapes: [], notes: [['x'.repeat(n)]] }] }, async (zip) => { const p = 'ppt/notesSlides/notesSlide1.xml'; zip.file(p, (await zip.file(p)!.async('string')).replace('<a:p>', '<a:p><a:pPr lvl="2"><a:buChar char="•"/></a:pPr>')) }, { timing: withBuilds })
    const h = await q.store.acquire(0)
    expect(h.slide.notes[0]).toMatchObject({ level: 2, bullet: '•', text: 'x'.repeat(65536) })
    expect(q.sink.list().some(d => d.code === 'notes-truncated')).toBe(n > 65536)
    if (n > 65536) {
      expect(h.slide.timing).toEqual({ refused: 'notes-limit' })
      expect(h.slide.builds).toEqual([])
    }
    h.release()
    q.store.close()
  }
  for (const extra of [0, 1]) {
    const q = await fixture({ slides: [{ shapes: [], notes: [['hello']] }] }, async (zip) => { const p = 'ppt/notesSlides/notesSlide1.xml', x = await zip.file(p)!.async('string'); zip.file(p, x.replace('</p:notes>', '<!--' + 'x'.repeat(1024 * 1024 + extra - x.length - 7) + '--></p:notes>')) })
    const h = await q.store.acquire(0)
    expect(h.slide.notes.length).toBe(extra ? 0 : 1)
    expect(h.slide.timing).toEqual(extra ? { refused: 'notes-limit' } : undefined)
    expect(q.sink.list().some(d => d.code === 'notes-limit')).toBe(!!extra)
    h.release()
    q.store.close()
  }
})
it('a 4 MiB slide part is read; 4 MiB + 1 drops that slide to static with a diagnostic', async () => {
  for (const extra of [0, 1]) {
    const q = await fixture({ slides: [{ shapes: [] }] }, async (zip) => { const p = 'ppt/slides/slide1.xml', x = await zip.file(p)!.async('string'); zip.file(p, x.replace('</p:sld>', '<!--' + 'x'.repeat(4 * 1024 * 1024 + extra - x.length - 7) + '--></p:sld>')) })
    const h = await q.store.acquire(0)
    expect(h.slide.timing).toEqual(extra ? { refused: 'read-limit' } : undefined)
    h.release()
    q.store.close()
    expect(q.w.used()).toBe(0)
  }
})
it('notes, links, media and an injected timing tree that each fit but together exceed the cap are refused as one aggregate', async () => {
  const q = await fixture({ slides: [{ shapes: [{ id: 2, paragraphs: [{ runs: [{ text: 'link' }] }] }], notes: [['note']], links: [{ shapeId: 2, target: 'https://example.com', external: true }], media: [{ shapeId: 2, kind: 'audio', ext: 'mp3', data: new Uint8Array([1]) }] }] })
  q.store.close()
  const { createSlideMetadataStore } = await import('../src/slideshow/read/store')
  const estimate = vi.fn((s: ShowSlide) => (s.notes.length + s.links.length + s.media.length + (s.timing ? 1 : 0)) * 20 * 1024 * 1024)
  const store = createSlideMetadataStore({ ...q, workspace: q.w, timing: () => ({ timing: { root: { kind: 'par', ctn: { id: 1, durMs: 100, accel: 0, decel: 0, autoRev: false, start: [], end: [] }, children: [] }, build: [] }, builds: [] }), resolver: () => ({ shape: () => 'slide', paragraphs: () => [] }), estimate })
  const h = await store.acquire(0)
  expect(estimate).toHaveReturnedWith(80 * 1024 * 1024)
  expect(h.slide.timing).toEqual({ refused: 'metadata-limit' })
  h.release()
  store.close()
})
it('hlinkClick on shape and run resolves slide jumps, show jumps and http/https/mailto; javascript:, file: and ppaction://macro are refused', async () => {
  const cases = [{ target: 'slide2.xml', action: 'ppaction://hlinksldjump' }, { action: 'ppaction://hlinkshowjump?jump=nextslide' }, ...['http://example.com', 'https://example.com', 'mailto:a@example.com', 'javascript:alert(1)', 'file:///tmp/x'].map(target => ({ target, external: true })), { action: 'ppaction://macro?name=x' }]
  const shapes = cases.map((_, i) => ({
    id: i + 2, paragraphs: [{ runs: [{ text: 'link' }] }], off: ['0', '0'] as [
      string,
      string
    ], ext: ['1000', '1000'] as [
      string,
      string
    ]
  }))
  const q = await fixture({ slides: [{ shapes, links: cases.map((x, i) => ({ shapeId: i + 2, ...x, onRun: i === 2 })) }, { shapes: [] }] })
  const h = await q.store.acquire(0)
  expect(h.slide.links.map(l => l.action.kind)).toEqual(['slide', 'jump', 'url', 'url', 'url', 'refused', 'refused', 'refused'])
  expect(h.slide.links[0].action).toEqual({ kind: 'slide', index: 1 })
  expect(h.slide.links[2].run).toEqual({ p: 0, r: 0 })
  h.release()
  q.store.close()
})
it('media records keep declared size and content type; no bytes are read', async () => {
  const q = await fixture({ slides: [{ shapes: [{ id: 2, off: ['0', '0'], ext: ['1000', '1000'] }], media: [{ shapeId: 2, kind: 'video', ext: 'mp4', data: new Uint8Array([1, 2, 3]), p14: true }] }] })
  const read = vi.spyOn(q.handle.pkg, 'bytes')
  const streams = q.handle.pkg.paths(p => p.startsWith('ppt/media/')).map(p => vi.spyOn(q.handle.pkg.entry(p) as unknown as {
    internalStream: () => unknown
  }, 'internalStream'))
  const h = await q.store.acquire(0)
  expect(h.slide.media.length).toBeGreaterThan(0)
  expect(h.slide.media[0]).toMatchObject({ declaredBytes: 3, contentType: 'video/mp4' })
  expect(read).not.toHaveBeenCalled()
  for (const spy of streams)
    expect(spy).not.toHaveBeenCalled()
  h.release()
  q.store.close()
})
it.skipIf(!existsSync('corpus/lo-tdf104788.pptx'))('tdf104788 reads 6 notes bodies and showNarration=1', async () => {
  const { readFileSync } = await import('node:fs')
  const handle = await openShowPackage(new Uint8Array(readFileSync('corpus/lo-tdf104788.pptx')))
  const sink = createDiagnosticSink(), w = workspace()
  const doc: PptxDocument = { slideWidthEmu: 9144000, slideHeightEmu: 6858000, images: [], slides: handle.slideParts.map((_, index) => ({ index, widthEmu: 9144000, heightEmu: 6858000, shapes: [] })) }
  const index = await readShowIndex(doc, handle, sink)
  expect(index.settings.showNarration).toBe(true)
  expect(index.slides).toHaveLength(6)
  const { createSlideMetadataStore } = await import('../src/slideshow/read/store')
  const store = createSlideMetadataStore({ doc, handle, index, sink, workspace: w, resolver: () => ({ shape: () => 'slide', paragraphs: () => [] }) })
  for (let i = 0; i < 6; i++) {
    const h = await store.acquire(i)
    expect(h.slide.notes.length).toBeGreaterThan(0)
    h.release()
  }
  store.close()
  expect(w.used()).toBe(0)
})
it('eviction during an acquire prevents cache publication without invalidating its eventual handle', async () => {
  const { store, w } = await setup(() => 1000)
  const p = store.acquire(0)
  store.evict(0)
  const h = await p
  expect(store.resident()).toEqual([])
  expect(w.used()).toBe(1000)
  h.release()
  expect(w.used()).toBe(0)
  store.close()
})
it('a released handle cannot regrow and invalid growth never changes accounting', async () => {
  const { store, w } = await setup(() => 1000)
  const h = await store.acquire(0)
  for (const n of [-1, NaN, Infinity, .5])
    expect(h.attach(n, 'bad')).toBe(false)
  expect(w.used()).toBe(1000)
  h.release()
  expect(h.attach(1, 'late')).toBe(false)
  store.close()
})
it('metadata construction is refused before the timing reader when live handles fill the aggregate', async () => {
  const { store, timing } = await setup(() => 64 * 1024 * 1024)
  const a = await store.acquire(0)
  await expect(store.acquire(1)).rejects.toMatchObject({ code: 'resource-limit' })
  expect(timing).toHaveBeenCalledTimes(1)
  a.release()
  store.close()
})
// ---- Fix round 1: honest metadata admission, with padding that is real markup ----
const MiB = 1024 * 1024
const shapeXml = (i: number, name = `Shape ${i}`) => `<p:sp><p:nvSpPr><p:cNvPr id="${i + 100}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${i}" y="${i}"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>Quarterly revenue grew across all regions this year.</a:t></a:r></a:p></p:txBody></p:sp>`
/** Pads a slide part to exactly `total` bytes with real shapes (about one '<' or '=' per 15 bytes); the last shape's name attribute absorbs the remainder. */
async function denseSlide(zip: import('jszip'), path: string, total: number) {
  const [head, tail] = (await zip.file(path)!.async('string')).replace('</p:spTree>', '\u0000</p:spTree>').split('\u0000')
  let body = '', i = 0
  while (head.length + body.length + tail.length + shapeXml(i).length + 600 < total)
    body += shapeXml(i++)
  const last = shapeXml(i, '')
  zip.file(path, head + body + shapeXml(i, 'x'.repeat(total - head.length - body.length - tail.length - last.length)) + tail)
}
const slide1 = 'ppt/slides/slide1.xml', slide2 = 'ppt/slides/slide2.xml', slide3 = 'ppt/slides/slide3.xml'
const staticBy = (sink: { list(): readonly { code: string }[] }) => sink.list().filter(d => d.code === 'slide-refused' || d.code === 'metadata-limit').map(d => d.code)
it('a 4 MiB slide of real markup is read on the 256 MiB mobile workspace and charges only its retained model', async () => {
  const q = await fixture({ slides: [{ shapes: [] }] }, zip => denseSlide(zip, slide1, 4 * MiB), { limit: 256 * MiB })
  const h = await q.store.acquire(0)
  expect(h.slide.timing).toBeUndefined()
  expect(staticBy(q.sink)).toEqual([])
  expect(q.w.used()).toBeLessThan(1 * MiB)
  h.release()
  q.store.close()
  expect(q.w.used()).toBe(0)
}, 60_000)
it('4 MiB + 1 of real markup is refused as read-limit', async () => {
  const q = await fixture({ slides: [{ shapes: [] }] }, zip => denseSlide(zip, slide1, 4 * MiB + 1), { limit: 256 * MiB })
  const h = await q.store.acquire(0)
  expect(h.slide.timing).toEqual({ refused: 'read-limit' })
  h.release()
  q.store.close()
  expect(q.w.used()).toBe(0)
}, 60_000)
it('a 450 KB slide of real markup stays live while a neighbour handle is held', async () => {
  const q = await fixture({ slides: [{ shapes: [] }, { shapes: [] }] }, zip => denseSlide(zip, slide2, 450 * 1000))
  const a = await q.store.acquire(0), b = await q.store.acquire(1)
  expect(b.slide.timing).toBeUndefined()
  expect(staticBy(q.sink)).toEqual([])
  expect([...q.store.resident()].sort()).toEqual([0, 1])
  a.release()
  b.release()
  q.store.close()
  expect(q.w.used()).toBe(0)
})
it('admitting a large slide does not evict unheld cached neighbours', async () => {
  const q = await fixture({ slides: [{ shapes: [] }, { shapes: [] }] }, zip => denseSlide(zip, slide2, 450 * 1000));
  (await q.store.acquire(0)).release()
  expect(q.store.resident()).toEqual([0]);
  (await q.store.acquire(1)).release()
  expect([...q.store.resident()].sort()).toEqual([0, 1])
  q.store.close()
})
it('the construction reservation grows with the timing region it will build, ahead of the timing reader, and a refusal makes the slide static as metadata-limit', async () => {
  const { slideMetadataBytes } = await import('../src/slideshow/read/store')
  const timing = vi.fn<SlideTimingReader>(() => ({ timing: undefined, builds: [] }))
  const region = '<p:timing><p:tnLst><p:par><p:cTn id="1">' + Array.from({ length: 8000 }, (_, i) => `<p:cTn id="${i + 2}" fill="hold"/>`).join('') + '</p:cTn></p:par></p:tnLst></p:timing>'
  const q = await fixture({ slides: [{ shapes: [] }, { shapes: [] }, { shapes: [] }] }, async zip => { zip.file(slide3, (await zip.file(slide3)!.async('string')).replace('</p:sld>', region + '</p:sld>')) }, { timing })
  const a = await q.store.acquire(0)
  // Leave 2 MiB of the 64 MiB aggregate: a plain slide fits, a 200 KB timing region (reserve > 3 MiB) does not.
  expect(a.attach(64 * MiB - 2 * MiB - slideMetadataBytes(a.slide), 'fill')).toBe(true)
  const plain = await q.store.acquire(1)
  expect(plain.slide.timing).toBeUndefined()
  expect(timing).toHaveBeenCalledTimes(2)
  const heavy = await q.store.acquire(2)
  expect(heavy.slide.timing).toEqual({ refused: 'metadata-limit' })
  expect(timing).toHaveBeenCalledTimes(2)
  expect(q.sink.list().some(d => d.code === 'metadata-limit' && d.slideIndex === 2)).toBe(true)
  for (const h of [a, plain, heavy])
    h.release()
  q.store.close()
  expect(q.w.used()).toBe(0)
})
it('the store charges its own estimate of notes, links, media and timing together', async () => {
  const { slideMetadataBytes } = await import('../src/slideshow/read/store')
  const big: SlideTimingReader = () => ({ timing: undefined, builds: Array.from({ length: 500 }, (_, i) => ({ spid: String(i), mode: 'whole' as const, ranges: [], animBg: false, body: 'b' + i })) })
  const spec = { slides: [{ shapes: [{ id: 2, paragraphs: [{ runs: [{ text: 'link' }] }] }], notes: [['note']], links: [{ shapeId: 2, target: 'https://example.com', external: true }], media: [{ shapeId: 2, kind: 'audio' as const, ext: 'mp3', data: new Uint8Array([1]) }] }] }
  const q = await fixture(spec, undefined, { timing: big })
  const h = await q.store.acquire(0)
  const { notes, links, media, builds } = h.slide
  expect(notes.length && links.length && media.length && builds.length).toBeTruthy()
  expect(q.w.used()).toBe(slideMetadataBytes(h.slide))
  expect(slideMetadataBytes({ ...h.slide, notes: [] })).toBeLessThan(slideMetadataBytes(h.slide))
  expect(slideMetadataBytes({ ...h.slide, links: [] })).toBeLessThan(slideMetadataBytes(h.slide))
  expect(slideMetadataBytes({ ...h.slide, media: [] })).toBeLessThan(slideMetadataBytes(h.slide))
  expect(slideMetadataBytes({ ...h.slide, builds: [] })).toBeLessThan(slideMetadataBytes(h.slide) - 500 * 40)
  h.release()
  q.store.close()
})
it('an injected timing reader that throws a programming error rejects acquire and charges nothing', async () => {
  const q = await fixture({ slides: [{ shapes: [] }] }, undefined, { timing: () => { throw new TypeError('timing bug') } })
  await expect(q.store.acquire(0)).rejects.toThrow(TypeError)
  expect(q.w.used()).toBe(0)
  q.store.close()
})
it('a malformed slide part is static as read-error, not read-limit', async () => {
  const q = await fixture({ slides: [{ shapes: [] }] }, async zip => { zip.file(slide1, (await zip.file(slide1)!.async('string')).replace('</p:sld>', '<p:broken></p:sld>')) })
  const h = await q.store.acquire(0)
  expect(h.slide.timing).toEqual({ refused: 'read-error' })
  expect(q.sink.list().some(d => d.code === 'slide-refused')).toBe(true)
  h.release()
  q.store.close()
  expect(q.w.used()).toBe(0)
})
it('without a handle the store hands out static slides and no-package is diagnosed once', async () => {
  const q = await fixture({ slides: [{ shapes: [] }, { shapes: [] }] })
  q.store.close()
  const { createSlideMetadataStore } = await import('../src/slideshow/read/store')
  const sink = createDiagnosticSink(), w = workspace()
  const noHandle = await readShowIndex(q.doc, undefined, sink)
  const store = createSlideMetadataStore({ doc: q.doc, index: noHandle, sink, workspace: w, resolver: () => ({ shape: () => 'slide', paragraphs: () => [] }) })
  const [a, b] = [await store.acquire(0), await store.acquire(1)]
  expect(a.slide.timing).toEqual({ refused: 'no-package' })
  expect(b.slide.timing).toEqual({ refused: 'no-package' })
  expect(sink.list().filter(d => d.code === 'no-package')).toHaveLength(1)
  a.release()
  b.release()
  store.close()
})
const hyperlink = '<Relationship Id="rIdL1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>'
const click = '<a:hlinkClick r:id="rIdL1"/>'
const linkShape = (id: number) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:rPr>${click}</a:rPr><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>`
it('links on a group, inside mc:AlternateContent (one branch only) and in table cells are read or explicitly refused', async () => {
  const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="50" name="g">${click}</p:cNvPr><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:grpSp>`
  const alt = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="p14">${linkShape(60)}</mc:Choice><mc:Fallback>${linkShape(60)}</mc:Fallback></mc:AlternateContent>`
  const table = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="70" name="t"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:rPr>${click}</a:rPr><a:t>x</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  const q = await fixture({ slides: [{ shapes: [] }] }, async zip => {
    zip.file(slide1, (await zip.file(slide1)!.async('string')).replace('</p:spTree>', group + alt + table + '</p:spTree>'))
    const rels = 'ppt/slides/_rels/slide1.xml.rels'
    zip.file(rels, (await zip.file(rels)!.async('string')).replace('</Relationships>', hyperlink + '</Relationships>'))
  })
  const h = await q.store.acquire(0)
  const bySpid = (id: string) => h.slide.links.filter(l => l.spid === id)
  expect(bySpid('50')).toEqual([{ spid: '50', action: { kind: 'url', href: 'https://example.com/' } }])
  expect(bySpid('60')).toEqual([{ spid: '60', run: { p: 0, r: 0 }, action: { kind: 'url', href: 'https://example.com/' } }])
  expect(bySpid('70')).toEqual([{ spid: '70', action: { kind: 'refused', reason: 'table-cell' } }])
  expect(q.sink.list().some(d => d.code === 'link-refused' && d.effect?.startsWith('70:'))).toBe(true)
  h.release()
  q.store.close()
})
it('media content types match the Default extension case-insensitively', async () => {
  const q = await fixture({ slides: [{ shapes: [{ id: 2, off: ['0', '0'], ext: ['1000', '1000'] }], media: [{ shapeId: 2, kind: 'video', ext: 'mp4', data: new Uint8Array([1, 2, 3]) }] }] }, async zip => { zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace(/<Override PartName="\/ppt\/media\/[^"]+"[^>]*>/, '<Default Extension="MP4" ContentType="video/mp4"/>')) })
  const h = await q.store.acquire(0)
  expect(h.slide.media[0]).toMatchObject({ contentType: 'video/mp4' })
  h.release()
  q.store.close()
})
