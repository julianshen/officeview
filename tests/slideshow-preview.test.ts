/**
 * S2-4: presenter next-checkpoint preview and mirror (spec 5.2 current and
 * next; plan S2-4). Timing trees come from real decks read through the real
 * metadata store; the paint tests use jsdom canvases (node-canvas) and are
 * pixel-picked, never GPU pixels. Nothing waits on wall time.
 *
 * Builder decks are "Office-shaped synthetic".
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { buildPptxDeck, mainSeqXml, timingXmlOf, type PptxDeckSpec, type PptxShapeSpec } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import type { PptxDocument, PptxSlide } from '../src/pptx/types'
import { openShowPackage, readShowIndex } from '../src/slideshow/read/deck'
import { readSlideTiming } from '../src/slideshow/read/timing'
import { shapeResolver } from '../src/slideshow/read/resolve'
import { createSlideMetadataStore } from '../src/slideshow/read/store'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { createSlideShowWorkspace, surfaceBytes, type ClosableSubOwner, type SlideShowTerminalReserve } from '../src/slideshow/workspace'
import type { Lease, ShowModel, SlideShowWorkspace } from '../src/slideshow/types'
import { renderStateAt } from '../src/slideshow/engine/behaviours'
import { createRng } from '../src/slideshow/rng'
import { paintDirect } from '../src/slideshow/render/direct'
import {
  createPresenterSurfaces, mirrorFrame, nextCheckpoint, paintPreview, previewHandle, previewTimeline, primePreview, releasePreview,
  type PreviewTarget,
} from '../src/slideshow/presenter/preview'

// paintDirect numbered so the mirror can prove it rendered nothing itself.
vi.mock('../src/slideshow/render/direct', async (importOriginal) => {
  const m = await importOriginal<typeof import('../src/slideshow/render/direct')>()
  return { ...m, paintDirect: vi.fn(m.paintDirect) }
})

// ---------------------------------------------------------------------------
// Workspace ledger (the S1-2 contracts, test-owned)
// ---------------------------------------------------------------------------
function testWorkspace(limit = 512 * 1024 * 1024): SlideShowWorkspace {
  let used = 0
  return {
    limit,
    used: () => used,
    reserveTerminal: () => undefined,
    lease: (bytes, tag) => {
      if (used + bytes > limit) return undefined
      used += bytes
      let live = true
      let held = bytes
      const lease: Lease = {
        tag,
        get bytes() { return live ? held : 0 },
        grow(more) {
          if (!live || used + more > limit) return false
          used += more
          held += more
          return true
        },
        shrink(less) { const n = Math.min(less, held); held -= n; used -= n },
        transfer: () => lease,
        release() { if (!live) return; live = false; used -= held; held = 0 },
      }
      return lease
    },
  }
}

// ---------------------------------------------------------------------------
// Decks
// ---------------------------------------------------------------------------
const box = (id: number, fill: string, ext: [string, string] = ['914400', '914400']): PptxShapeSpec => ({ id, prst: 'rect', off: ['0', '0'], ext, fill })

/** One click step: shape `spid` becomes visible after `dur` ms (1 ms reads as a counted zero-duration set). */
const click = (base: number, spid: string, dur = 1): string =>
  `<p:par><p:cTn id="${base}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${base + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${base + 2}" nodeType="clickEffect" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:set><p:cBhvr><p:cTn id="${base + 3}" dur="${dur}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>` +
  `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
  `<p:to><p:strVal val="visible"/></p:to></p:set>` +
  `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`

/** One click step whose effect is one finite anim of `dur` ms on `spid` (the settle-vs-leave fixture). */
const clickAnim = (base: number, spid: string, dur: number): string =>
  `<p:par><p:cTn id="${base}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${base + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${base + 2}" nodeType="clickEffect" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:anim><p:cBhvr><p:cTn id="${base + 3}" dur="${dur}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>` +
  `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr>` +
  `<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav></p:tavLst></p:anim>` +
  `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`

/**
 * One click step whose effect is two behaviours: a zero-duration set (shape 2
 * becomes visible) joining an ever-repeating 500 ms anim (shape 3) — the
 * settled checkpoint of that step must have the set applied and the anim at
 * its first iteration's end.
 */
const zeroJoinIndefinite = (base: number): string =>
  `<p:par><p:cTn id="${base}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${base + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${base + 2}" nodeType="clickEffect" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:set><p:cBhvr><p:cTn id="${base + 3}" dur="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>` +
  `<p:tgtEl><p:spTgt spid="2"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
  `<p:to><p:strVal val="visible"/></p:to></p:set>` +
  `<p:anim><p:cBhvr><p:cTn id="${base + 4}" dur="500" repeatCount="indefinite" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>` +
  `<p:tgtEl><p:spTgt spid="3"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr>` +
  `<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav></p:tavLst></p:anim>` +
  `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`

const audio: Uint8Array = new Uint8Array(2048).map((_, i) => i % 251)

/** Two visible slides plus a hidden one; slide 0 carries the given timing. */
async function deck(timingOfSlide0: string, withMedia = false): Promise<{ doc: PptxDocument; bytes: Uint8Array }> {
  const spec: PptxDeckSpec = {
    slides: [
      {
        shapes: [box(2, 'ff0000', ['457200', '3429000']), box(3, '0000ff')],
        timingXml: timingOfSlide0,
        ...(withMedia ? { media: [{ kind: 'audio' as const, data: audio, ext: 'mp3', shapeId: 2 }] } : {}),
      },
      { shapes: [{ id: 2, prst: 'rect', off: ['0', '0'], ext: ['9144000', '6858000'], fill: '000000' }] },
      { shapes: [box(4, '00ff00')], hidden: true },
    ],
  }
  const bytes = await buildPptxDeck(spec)
  return { doc: await parsePptx(await OfficePackage.load(bytes)), bytes }
}

/** A real model over the deck: the show index plus the metadata store. */
async function modelOf(d: Awaited<ReturnType<typeof deck>>, w: SlideShowWorkspace) {
  const sink = createDiagnosticSink()
  const handle = await openShowPackage(d.bytes)
  const index = await readShowIndex(d.doc, handle, sink)
  const store = createSlideMetadataStore({
    doc: d.doc, handle, index, workspace: w, sink, timing: readSlideTiming,
    resolver: (slide) => shapeResolver(slide, []),
  })
  return { model: { index, store } as ShowModel, store, sink }
}

const NORMAL = { showAnimation: true, reducedMotion: false }

/** A live cursor 0 ms into step 1 of the primed slide (its first click). */
function liveCursor(model: ShowModel, slide: number, at: number) {
  const tl = previewTimeline(model, slide)!
  const c = tl.next(tl.initial(), at)
  expect(c.mode).toBe('live')
  return c
}

afterEach(() => vi.clearAllMocks())

// ---------------------------------------------------------------------------
// nextCheckpoint
// ---------------------------------------------------------------------------

describe('nextCheckpoint', () => {
  it('next is the next build of this slide, else the next visible slide checkpoint 0, else end', async () => {
    const d = await deck(timingXmlOf(mainSeqXml([click(3, '2'), click(10, '3')])))
    const { model } = await modelOf(d, testWorkspace())
    expect(await primePreview(model, 0, NORMAL)).toBe(true)
    const order = [0, 1]
    const tl = previewTimeline(model, 0)!

    const p1 = nextCheckpoint(model, order, 0, tl.initial(), 0)
    if (p1 === 'end') throw new Error('precondition: a slide with a further step answers a position')
    expect(p1.slide).toBe(0)
    expect(p1.cursor).toMatchObject({ step: 1, mode: 'checkpoint' })
    // Checkpoint replay: the settle point plus the step group, all at the horizon.
    expect(p1.cursor.events).toHaveLength(3)

    // At the last step the next click leaves the slide: the next visible slide, at its settled checkpoint 0.
    const atLast = nextCheckpoint(model, order, 0, tl.checkpoint(2), 0)
    expect(atLast).toEqual({ slide: 1, cursor: { step: 0, events: [], mode: 'checkpoint' } })

    // Past the visible order there is nowhere left: the end of the show.
    expect(nextCheckpoint(model, order, 1, tl.checkpoint(2), 0)).toBe('end')

    // A slide whose timing never loaded answers the slide-advance position, still honestly.
    const d2 = await deck(timingXmlOf(mainSeqXml([click(3, '2')])))
    const { model: m2 } = await modelOf(d2, testWorkspace())
    expect(nextCheckpoint(m2, order, 0, { step: 0, events: [], mode: 'initial' }, 0)).toEqual({ slide: 1, cursor: { step: 0, events: [], mode: 'checkpoint' } })
    releasePreview(model)
  })

  it('a live cursor at the last step settles while a finite effect runs and leaves once it has ended', async () => {
    // Two clicks, each a 500 ms effect: the settled-forever reading of the
    // last step must not swallow the slide advance after the effect ends.
    const d = await deck(timingXmlOf(mainSeqXml([clickAnim(3, '2', 500), clickAnim(8, '3', 500)])))
    const { model } = await modelOf(d, testWorkspace())
    expect(await primePreview(model, 0, NORMAL)).toBe(true)
    const tl = previewTimeline(model, 0)!
    const c1 = tl.next(tl.initial(), 0)
    const c2 = tl.next(c1, 1000)
    expect(c2.mode).toBe('live')
    expect(c2.step).toBe(2)

    // The clock says the final effect is still running (ends at 1500): the next
    // click is the settle point, so the preview shows this slide settled.
    const settle = nextCheckpoint(model, [0, 1], 0, c2, 1200)
    if (settle === 'end') throw new Error('precondition: the settle answer is a position')
    expect(settle.slide).toBe(0)
    expect(settle.cursor).toBe(tl.checkpoint(2))

    // After the effect has ended (5000 > 1000 + 500): the next click leaves the slide.
    expect(nextCheckpoint(model, [0, 1], 0, c2, 5000)).toEqual({ slide: 1, cursor: { step: 0, events: [], mode: 'checkpoint' } })

    // Without a clock... there is none: `t` is required, and the settled-reading
    // semantics it used to give (always leave at the last step) are gone.
    releasePreview(model)
  })

  it('a looping show wraps to the first slide of the order', async () => {
    const bytes = await buildPptxDeck({
      presProps: '<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:showPr loop="1"/></p:presentationPr>',
      slides: [
        { shapes: [box(2, 'ff0000', ['457200', '3429000'])], timingXml: timingXmlOf(mainSeqXml([click(3, '2'), click(10, '3')])) },
        { shapes: [{ id: 2, prst: 'rect', off: ['0', '0'], ext: ['9144000', '6858000'], fill: '000000' }] },
      ],
    })
    const d = { doc: await parsePptx(await OfficePackage.load(bytes)), bytes }
    const { model } = await modelOf(d, testWorkspace())
    expect(model.index.settings.loop).toBe(true)
    expect(await primePreview(model, 0, NORMAL)).toBe(true)
    const tl0 = previewTimeline(model, 0)!
    // Past the last visible slide the loop restarts at the order's first slide,
    // at that slide's own settled checkpoint 0.
    const pos = nextCheckpoint(model, [0, 1], 1, tl0.checkpoint(2), 0)
    if (pos === 'end') throw new Error('precondition: a looping show has a next slide')
    expect(pos.slide).toBe(0)
    expect(pos.cursor).toBe(tl0.checkpoint(0))
    releasePreview(model)
  })

  it('the slide-advance cursor is the next slide\'s own settled checkpoint 0 when that slide is primed', async () => {
    const d = await deck(timingXmlOf(mainSeqXml([click(3, '2'), click(10, '3')])))
    const { model } = await modelOf(d, testWorkspace())
    expect(await primePreview(model, 0, NORMAL)).toBe(true)
    expect(await primePreview(model, 1, NORMAL)).toBe(true)
    const tl0 = previewTimeline(model, 0)!
    const tl1 = previewTimeline(model, 1)!
    const pos = nextCheckpoint(model, [0, 1], 0, tl0.checkpoint(2), 0)
    if (pos === 'end') throw new Error('precondition: the advance is a position')
    expect(pos.slide).toBe(1)
    expect(pos.cursor).toBe(tl1.checkpoint(0))
    releasePreview(model)
  })

  it('zero-duration effects applied; indefinite at first-iteration end', async () => {
    const d = await deck(timingXmlOf(mainSeqXml([zeroJoinIndefinite(3)])))
    const { model } = await modelOf(d, testWorkspace())
    expect(await primePreview(model, 0, NORMAL)).toBe(true)
    const tl = previewTimeline(model, 0)!
    const pos = nextCheckpoint(model, [0], 0, tl.initial(), 0)
    if (pos === 'end') throw new Error('precondition: a slide with a step answers a position')
    const snap = tl.stateAt(pos.cursor, 0)
    // The zero-duration set: applied and frozen at its settled end.
    const set = snap.frozen.find(e => e.node.ctn.id === 6)
    expect(set).toBeDefined()
    expect(set!.progress).toBeCloseTo(1, 9)
    // The indefinite anim: shown at the end of its first iteration, still running on the live side.
    const spin = snap.active.find(e => e.node.ctn.id === 7)
    expect(spin).toBeDefined()
    expect(spin!.iteration).toBe(0)
    expect(spin!.progress).toBeCloseTo(1, 9)
    // The render state carries the settled answer: shape 2 is visible.
    const sink = createDiagnosticSink()
    const state = renderStateAt(snap, new Map(), createRng(0), sink)
    expect(state.get('sp:2')!.visible).toBe(true)
    expect(sink.list()).toEqual([])
    releasePreview(model)
  })

  it('computing next leaves the live cursor, events and media untouched', async () => {
    const d = await deck(timingXmlOf(mainSeqXml([click(3, '2'), click(10, '3')])), true)
    const w = testWorkspace()
    const { model, store } = await modelOf(d, w)
    expect(await primePreview(model, 0, NORMAL)).toBe(true)
    const cursor = liveCursor(model, 0, 0)
    const events = cursor.events
    const before = [...events]
    const handle = previewHandle(model, 0)!
    const slideRef = handle.slide
    const media = slideRef.media
    const partPath = media[0].partPath
    const used = w.used()
    const resident = [...store.resident()]

    expect(nextCheckpoint(model, [0, 1], 0, cursor, cursor.events[cursor.events.length - 1]?.t ?? 0)).not.toBe('end')

    expect(cursor.step).toBe(1)
    expect(cursor.events).toBe(events)
    expect([...events]).toEqual(before)
    expect(handle.slide).toBe(slideRef)
    expect(handle.slide.media).toBe(media)
    expect(media[0].partPath).toBe(partPath)
    expect(partPath).toContain('.mp3')
    expect(store.resident()).toEqual(resident)
    expect(w.used()).toBe(used)
    releasePreview(model)
  })
})

// ---------------------------------------------------------------------------
// primePreview holds
// ---------------------------------------------------------------------------

describe('primePreview holds', () => {
  it('an admission refusal is diagnosed false; cancellations and programming errors rethrow', async () => {
    const d = await deck(timingXmlOf(mainSeqXml([click(3, '2')])))
    { // A 10-byte workspace: the store's own admission refuses; the answer is false + one diagnostic.
      const sink = createDiagnosticSink()
      const handle = await openShowPackage(d.bytes)
      const index = await readShowIndex(d.doc, handle, sink)
      const w = testWorkspace(10)
      const store = createSlideMetadataStore({ doc: d.doc, handle, index, workspace: w, sink, timing: readSlideTiming, resolver: s => shapeResolver(s, []) })
      const model = { index, store } as ShowModel
      expect(await primePreview(model, 0, { ...NORMAL, sink })).toBe(false)
      expect(sink.list().map(x => x.code)).toContain('preview-prime')
      expect(w.used()).toBe(0)
    }
    { // A programming error (an out-of-range slide index) is rethrown, never swallowed into `false`.
      const sink = createDiagnosticSink()
      const handle = await openShowPackage(d.bytes)
      const index = await readShowIndex(d.doc, handle, sink)
      const w = testWorkspace()
      const store = createSlideMetadataStore({ doc: d.doc, handle, index, workspace: w, sink, timing: readSlideTiming, resolver: s => shapeResolver(s, []) })
      const model = { index, store } as ShowModel
      await expect(primePreview(model, 99, { ...NORMAL, sink })).rejects.toThrow(RangeError)
      expect(sink.list().map(x => x.code)).not.toContain('preview-prime')
    }
  })

  it('a release never evicts the entry the shown slide still holds; re-acquiring it does not double-charge (N1)', async () => {
    const d = await deck(timingXmlOf(mainSeqXml([click(3, '2')])))
    const w = testWorkspace()
    const { model, store } = await modelOf(d, w)
    // The audience is showing slide 0: the session (and residency) hold its metadata too.
    const shown = await store.acquire(0)
    expect(await primePreview(model, 0, NORMAL)).toBe(true)
    const charged = w.used()
    expect(charged).toBeGreaterThan(0)
    releasePreview(model, 0)
    // Spec 7: the current slide's metadata is never dropped while it is shown.
    expect(store.resident()).toEqual([0])
    // A later acquire re-reads nothing: no duplicate charge while the old handle is live.
    const again = await store.acquire(0)
    expect(w.used()).toBe(charged)
    // The whole-show form behaves the same on dispose/close.
    releasePreview(model)
    expect(store.resident()).toEqual([0])
    shown.release()
    again.release()
    // Nothing of the preview's is pinned, and the store's own close refunds everything.
    expect(previewHandle(model, 0)).toBeUndefined()
    expect(previewTimeline(model, 0)).toBeUndefined()
    await store.close()
    expect(w.used()).toBe(0)
  })

  it('concurrent primes share one handle; a release drops the holds, the store keeps its cached entry (N1)', async () => {
    const d = await deck(timingXmlOf(mainSeqXml([click(3, '2')])))
    { // Two primes in flight for one slide: one read, and one release lets go of the preview's handles.
      const w = testWorkspace()
      const { model } = await modelOf(d, w)
      const [a, b] = await Promise.all([primePreview(model, 0, NORMAL), primePreview(model, 0, NORMAL)])
      expect([a, b]).toEqual([true, true])
      expect(previewTimeline(model, 0)).toBeDefined()
      expect(w.used()).toBeGreaterThan(0)
      releasePreview(model)
      // No handle is pinned; the store keeps its normal cached entry (spec 7)
      // and refunds it under its own pressure rules — the charge stays.
      expect(previewTimeline(model, 0)).toBeUndefined()
      expect(previewHandle(model, 0)).toBeUndefined()
      expect(w.used()).toBeGreaterThan(0)
    }
    { // releasePreview while a prime is still reading: the late handle is returned, never pinned.
      const w = testWorkspace()
      const { model } = await modelOf(d, w)
      const flight = primePreview(model, 0, NORMAL)
      releasePreview(model)
      expect(await flight).toBe(false)
      expect(previewTimeline(model, 0)).toBeUndefined()
      expect(previewHandle(model, 0)).toBeUndefined()
      // Re-priming after the release reuses the store's cached entry: no duplicate charge.
      expect(await primePreview(model, 0, NORMAL)).toBe(true)
      const held = w.used()
      expect(held).toBeGreaterThan(0)
      releasePreview(model)
      expect(w.used()).toBe(held)
      expect(previewTimeline(model, 0)).toBeUndefined()
    }
  })

  it('a per-slide release hands back that slide only', async () => {
    const d = await deck(timingXmlOf(mainSeqXml([click(3, '2')])))
    const w = testWorkspace()
    const { model } = await modelOf(d, w)
    expect(await primePreview(model, 0, NORMAL)).toBe(true)
    expect(await primePreview(model, 1, NORMAL)).toBe(true)
    releasePreview(model, 1)
    expect(previewHandle(model, 0)).toBeDefined()
    expect(previewTimeline(model, 0)).toBeDefined()
    expect(previewHandle(model, 1)).toBeUndefined()
    expect(previewTimeline(model, 1)).toBeUndefined()
    releasePreview(model)
  })
})

// ---------------------------------------------------------------------------
// Paint: mirror and preview
// ---------------------------------------------------------------------------

describe('mirror', () => {
  it('the mirror shows the in-progress animation from the same state', () => {
    // The audience side: a red "in-progress" state over black, as it was rendered.
    const source = document.createElement('canvas')
    source.width = 200
    source.height = 100
    const ctx = source.getContext('2d')!
    ctx.fillStyle = '#000000'
    ctx.fillRect(0, 0, 200, 100)
    ctx.fillStyle = '#ff0000'
    ctx.fillRect(0, 0, 50, 50)

    const mirror = document.createElement('canvas')
    mirror.width = 100
    mirror.height = 100
    mirrorFrame(source, mirror.getContext('2d')!)

    // Same state: the mirror is a copy of the rendered frame, it renders nothing itself.
    expect(vi.mocked(paintDirect)).toHaveBeenCalledTimes(0)
    const pixels = (x: number, y: number) => {
      const p = mirror.getContext('2d')!.getImageData(x, y, 1, 1).data
      return [p[0], p[1], p[2]]
    }
    // 200x100 fitted into 100x100 draws at half scale, y 25..75; the red square occupies the top-left quarter.
    expect(pixels(10, 30)).toEqual([255, 0, 0])
    expect(pixels(60, 50)).toEqual([0, 0, 0])
    expect(pixels(50, 10)).toEqual([0, 0, 0])
    expect(pixels(50, 90)).toEqual([0, 0, 0])
  })

  it('the mirrored frame is the rendered state, painted once, and the bands hold nothing stale', async () => {
    const d = await deck('')
    const state = new Map()
    const assets = { images: [], resolveFont: (f: string) => f, scale: 1, scratch: { lease: () => undefined } }
    // The audience frame is painted once, through paintDirect, from this state.
    const source = document.createElement('canvas')
    source.width = 400
    source.height = 300
    paintPreview(source.getContext('2d')!, { slide: d.doc.slides[1], state, backing: { w: 400, h: 300 } }, assets)
    expect(vi.mocked(paintDirect)).toHaveBeenCalledTimes(1)
    expect(vi.mocked(paintDirect).mock.calls.length).toBeGreaterThan(0)
    expect(vi.mocked(paintDirect).mock.calls[0]![2]).toBe(state)

    const mirror = document.createElement('canvas')
    mirror.width = 200
    mirror.height = 150
    mirrorFrame(source, mirror.getContext('2d')!)
    // No second render: the black full-slide frame carried across 2:1 -> the
    // pane is covered, and paintDirect still ran exactly once.
    expect(vi.mocked(paintDirect)).toHaveBeenCalledTimes(1)
    const px = (x: number, y: number) => [...mirror.getContext('2d')!.getImageData(x, y, 1, 1).data].slice(0, 3)
    expect(px(100, 75)).toEqual([0, 0, 0])
  })

  it('mirrorFrame clears the bands instead of keeping stale pixels', () => {
    const square = document.createElement('canvas')
    square.width = 100
    square.height = 100
    square.getContext('2d')!.fillStyle = '#ff0000'
    square.getContext('2d')!.fillRect(0, 0, 100, 100)
    const mirror = document.createElement('canvas')
    mirror.width = 100
    mirror.height = 100
    mirrorFrame(square, mirror.getContext('2d')!)
    // The first blit fills the whole mirror red; a 2:1 source then draws 100x50
    // at y 25..75 — the bands must be cleared, not left red.
    const wide = document.createElement('canvas')
    wide.width = 200
    wide.height = 100
    mirrorFrame(wide, mirror.getContext('2d')!)
    const bands = [...mirror.getContext('2d')!.getImageData(50, 5, 1, 1).data].slice(0, 3)
    expect(bands).toEqual([0, 0, 0])
  })
})

describe('watermark', () => {
  it('preview and mirror paint the watermark on top', async () => {
    const d = await deck('')
    const assets = {
      images: [],
      resolveFont: (family: string) => family,
      scale: 1,
      scratch: { lease: () => undefined },
      watermark: { text: 'OV-STEP', color: '#ff00ff', opacity: 1, fontSizePt: 320, placement: 'center' as const },
    }
    const target = (slide: PptxSlide): PreviewTarget => ({ slide, state: new Map(), backing: { w: 400, h: 300 } })
    const hasMark = (c: HTMLCanvasElement, x0: number, y0: number, x1: number, y1: number) => {
      const ctx = c.getContext('2d')!
      const px = ctx.getImageData(x0, y0, x1 - x0, y1 - y0).data
      for (let i = 0; i < px.length; i += 4) {
        if (px[i]! > 180 && px[i + 2]! > 180 && px[i + 1]! < 120) return true
      }
      return false
    }

    // The preview: a full-slide black shape, then the watermark drawn last over it.
    const preview = document.createElement('canvas')
    preview.width = 400
    preview.height = 300
    paintPreview(preview.getContext('2d')!, target(d.doc.slides[1]), assets)
    expect(hasMark(preview, 120, 100, 280, 200)).toBe(true)

    // The mirror: the audience frame was watermarked on the stage, the copy keeps it on top.
    const audience = document.createElement('canvas')
    audience.width = 400
    audience.height = 300
    paintPreview(audience.getContext('2d')!, target(d.doc.slides[1]), assets)
    const mirror = document.createElement('canvas')
    mirror.width = 200
    mirror.height = 150
    mirrorFrame(audience, mirror.getContext('2d')!)
    expect(hasMark(mirror, 60, 50, 140, 100)).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// Presenter surfaces
// ---------------------------------------------------------------------------

describe('presenter surfaces', () => {
  it('presenter surfaces lease on open, re-lease on resize, release everything on dispose, and never touch the terminal reserve', async () => {
    void (await deck(''))
    const ws = createSlideShowWorkspace('desktop', { limit: 8 * 1024 * 1024, scratch: 512 * 1024 })
    const terminal = ws.reserveTerminal(surfaceBytes(800, 600))! as SlideShowTerminalReserve
    const terminalScratch = terminal.scratch as ClosableSubOwner
    const terminalScratchBytes = terminalScratch.used()
    const base = ws.used()
    const sink = createDiagnosticSink()

    const surfaces = createPresenterSurfaces(document, ws, sink)
    expect(sink.list()).toEqual([])
    expect(ws.used()).toBeGreaterThan(base)
    expect(terminal.outputBytes).toBe(surfaceBytes(800, 600))
    expect(terminalScratch.used()).toBe(terminalScratchBytes)
    expect(surfaces.mirror.width).toBe(480)
    expect(surfaces.mirror.height).toBe(270)

    expect(surfaces.resize({ w: 400, h: 300 }, 2)).toBe(true)
    expect(surfaces.mirror.width).toBe(800)
    expect(surfaces.preview.height).toBe(600)
    expect(terminalScratch.used()).toBe(terminalScratchBytes)
    const resized = ws.used()

    // The grid rides along, leased on demand like every presenter surface.
    const g = surfaces.grid(0)
    expect(g.width).toBeGreaterThan(0)
    expect(ws.used()).toBeGreaterThan(resized)

    surfaces.dispose()
    expect(ws.used()).toBe(base)
    expect(terminal.outputBytes).toBe(surfaceBytes(800, 600))
  })

  it('a resize that steps down the DPR is diagnosed (spec 7: every step down is diagnosed)', async () => {
    // The open default fits; the resize to 800x600 at the capped DPR (3,840,000 B for both) does
    // not; 700x525 at DPR 1.75 (2,940,000 B) does — one diagnosed step down.
    const ws = createSlideShowWorkspace('desktop', { limit: 4_000_000 })
    const sink = createDiagnosticSink()
    const surfaces = createPresenterSurfaces(document, ws, sink)
    expect(surfaces.resize({ w: 400, h: 300 }, 2)).toBe(true)
    expect(surfaces.mirror.width).toBe(700)
    expect(surfaces.mirror.height).toBe(525)
    expect(surfaces.preview.width).toBe(700)
    const downs = sink.list().filter(x => x.code === 'presenter-surface' && (x.effect ?? '').startsWith('dpr-stepdown'))
    expect(downs).toHaveLength(1)
    expect(downs[0]!.effect).toContain('1.75')
    expect(sink.list()).toHaveLength(1)
    surfaces.dispose()
  })

  it('a refused presenter lease falls back to a placeholder', async () => {
    const sink = createDiagnosticSink()
    // Not even the open default fits both surfaces: the placeholder, diagnosed at construction.
    const tiny = createSlideShowWorkspace('desktop', { limit: 900_000 })
    const s2 = createPresenterSurfaces(document, tiny, sink)
    expect(sink.list().map(x => x.code)).toContain('presenter-surface')
    expect(s2.mirror.width).toBe(0)
    expect(s2.mirror.height).toBe(0)
    expect(s2.mirror.getAttribute('data-officeview-show-placeholder')).toBe('presenter-surface')
    expect(s2.mirror.getAttribute('data-officeview-show-surface')).toBe('presenter')
    expect(s2.mirror.textContent).toContain('placeholder')
    expect(s2.preview.width).toBe(0)
    expect(s2.resize({ w: 400, h: 300 }, 2)).toBe(false)
    const codes = sink.list().map(x => x.code)
    expect(codes).toContain('presenter-surface')
    // The refused attempts charged nothing and neither does the placeholder.
    expect(tiny.used()).toBe(0)
    s2.dispose()
  })

  it('dispose releases scratch leases, refuses further leases, and leaves safe zero-size canvases', () => {
    const ws = createSlideShowWorkspace('desktop')
    const sink = createDiagnosticSink()
    const surfaces = createPresenterSurfaces(document, ws, sink)
    const work = surfaces.scratch.lease(1000, 'presenter:work')
    expect(work).toBeDefined()
    expect(ws.used()).toBeGreaterThan(0)
    expect(surfaces.grid(0).width).toBe(320)

    // The replaced canvas hands its bitmap back with its lease.
    const oldMirror = surfaces.mirror
    expect(surfaces.resize({ w: 400, h: 300 }, 2)).toBe(true)
    expect(oldMirror.width).toBe(0)
    expect(oldMirror.height).toBe(0)

    surfaces.dispose()
    expect(ws.used()).toBe(0)
    expect(work!.bytes).toBe(0)
    expect(surfaces.scratch.lease(1000, 'presenter:work')).toBeUndefined()
    expect(ws.used()).toBe(0)
    expect(sink.list().map(x => x.code)).toContain('presenter-surface')
    // The panes answer safe zero-size canvases instead of throwing.
    expect(surfaces.mirror.width).toBe(0)
    expect(surfaces.preview.width).toBe(0)
    const late = surfaces.grid(0)
    expect(late.width).toBe(0)
    expect(late.getAttribute('data-officeview-show-placeholder')).toBe('presenter-surface')
    expect(ws.used()).toBe(0)
  })

  it('the grid is capped, and a refused thumbnail retries', () => {
    const sink = createDiagnosticSink()
    { // A short workspace: the refused placeholder is not cached; the thumbnail retries later.
      const tight = createSlideShowWorkspace('desktop', { limit: 1_300_000 })
      const surfaces = createPresenterSurfaces(document, tight, sink)
      const pressure = tight.lease(50_000, 'pressure')
      const refused = surfaces.grid(0)
      expect(refused.width).toBe(0)
      expect(refused.getAttribute('data-officeview-show-placeholder')).toBe('presenter-surface')
      pressure?.release()
      const retried = surfaces.grid(0)
      expect(retried.width).toBe(320)
      expect(retried).not.toBe(refused)
      surfaces.dispose()
    }
    { // Distinct thumbnails stop at the cap (48 at DPR 1), diagnosed; the ledger stays bounded.
      const ws = createSlideShowWorkspace('desktop')
      const surfaces = createPresenterSurfaces(document, ws, sink)
      for (let i = 0; i < 48; i++) expect(surfaces.grid(i).width).toBe(320)
      expect(surfaces.grid(48).width).toBe(0)
      expect(sink.list().map(x => x.code)).toContain('presenter-surface')
      expect(ws.used()).toBe(1_036_800 + 48 * 230_400)
      surfaces.dispose()
      expect(ws.used()).toBe(0)
    }
  })

  it('releaseGrid frees a cap slot, so thumbnails beyond the cap are not stuck as placeholders (N3)', () => {
    const ws = createSlideShowWorkspace('desktop')
    const surfaces = createPresenterSurfaces(document, ws)
    for (let i = 0; i < 48; i++) expect(surfaces.grid(i).width).toBe(320)
    expect(surfaces.grid(48).width).toBe(0)
    // The presenter frees the scrolled-past thumbnail (S2-5 calls releaseGrid on scroll):
    // the slot opens and the newly visible slide leases for real.
    expect(surfaces.releaseGrid(0)).toBe(true)
    expect(surfaces.grid(48).width).toBe(320)
    expect(ws.used()).toBe(1_036_800 + 48 * 230_400)
    // Freeing again is false: nothing is held for that slide now.
    expect(surfaces.releaseGrid(0)).toBe(false)
    surfaces.dispose()
    expect(ws.used()).toBe(0)
  })

  it('an invalid device pixel ratio at open is clamped and diagnosed, not thrown', () => {
    const ws = createSlideShowWorkspace('desktop')
    const sink = createDiagnosticSink()
    const doc = { createElement: document.createElement.bind(document), defaultView: { devicePixelRatio: -2 } } as unknown as Document
    const surfaces = createPresenterSurfaces(doc, ws, sink)
    expect(surfaces.mirror.width).toBe(480)
    expect(sink.list().map(x => x.effect)).toContain('dpr-clamped')
    surfaces.dispose()
    expect(ws.used()).toBe(0)
  })

  it('dispose runs the close callback exactly once, and a throwing one cannot stop the teardown (N2)', () => {
    const ws = createSlideShowWorkspace('desktop')
    let closed = 0
    const surfaces = createPresenterSurfaces(document, ws, undefined, () => { closed++ })
    const work = surfaces.scratch.lease(1000, 'presenter:work')
    surfaces.dispose()
    surfaces.dispose()
    expect(closed).toBe(1)
    expect(work!.bytes).toBe(0)

    // A throwing close callback runs last and guarded: the leases are already
    // released when it fires, the error still surfaces, and a retry after the
    // throw is a no-op — nothing leaks and nothing re-runs.
    const ws2 = createSlideShowWorkspace('desktop')
    let booms = 0
    const throwing = createPresenterSurfaces(document, ws2, undefined, () => { booms++; throw new Error('boom') })
    const held = throwing.scratch.lease(1000, 'presenter:work')
    expect(() => throwing.dispose()).toThrow('boom')
    expect(booms).toBe(1)
    expect(held!.bytes).toBe(0)
    expect(ws2.used()).toBe(0)
    expect(() => throwing.dispose()).not.toThrow()
    expect(booms).toBe(1)
  })
})

describe('PR #19 review 2', () => {
  it('primePreview schedules the expanded text timing the session schedules (by word / by letter), with its order', async () => {
    // The slide's own timing has one click; the text fan-out (as stage.textTiming would answer) has two.
    const d = await deck(timingXmlOf(mainSeqXml([click(3, '2')])))
    const { model } = await modelOf(d, testWorkspace())
    const fanned = await modelOf(await deck(timingXmlOf(mainSeqXml([click(3, '2'), click(10, '3')]))), testWorkspace())
    const h2 = await fanned.store.acquire(0)
    const tree = h2.slide.timing as import('../src/slideshow/types').TimingTree
    const orderOf = new Map<import('../src/slideshow/types').TimeNode, number>()
    const calls: Array<[number, unknown, boolean]> = []
    const textTiming = vi.fn((index: number, meta: unknown, charge: (bytes: number, tag: string) => boolean) => {
      calls.push([index, meta, charge(16, 'test:text-timing')])
      return { tree, orderOf }
    })
    expect(await primePreview(model, 0, { ...NORMAL, textTiming })).toBe(true)
    expect(textTiming).toHaveBeenCalledTimes(1)
    // The slide's metadata, charged to the preview's own handle.
    expect(calls[0]![0]).toBe(0)
    expect(calls[0]![1]).toBe(previewHandle(model, 0)!.slide)
    expect(calls[0]![2]).toBe(true)
    expect(previewTimeline(model, 0)!.steps).toBe(2)
    // No fan-out: the slide's own timing.
    const plain = await modelOf(d, testWorkspace())
    expect(await primePreview(plain.model, 0, { ...NORMAL, textTiming: () => undefined })).toBe(true)
    expect(previewTimeline(plain.model, 0)!.steps).toBe(1)
    h2.release()
  })
})

describe('PR #19 review 3', () => {
  it('releasing one slide cancels its in-flight prime: the late handle comes back, nothing is stored', async () => {
    const d = await deck(timingXmlOf(mainSeqXml([click(3, '2')])))
    const { model, store } = await modelOf(d, testWorkspace())
    let open!: () => void
    const gate = new Promise<void>((res) => { open = res })
    const released: number[] = []
    const slow = Object.create(store) as typeof store
    slow.acquire = async (i, signal) => {
      await gate
      const h = await store.acquire(i, signal)
      return { slide: h.slide, attach: (b, t) => h.attach(b, t), release: () => { released.push(i); h.release() } }
    }
    const m = { index: model.index, store: slow } as ShowModel
    const job = primePreview(m, 0, NORMAL)
    // The presenter's preview moved on before the read finished.
    releasePreview(m, 0)
    open()
    expect(await job).toBe(false)
    expect(previewTimeline(m, 0)).toBeUndefined()
    expect(previewHandle(m, 0)).toBeUndefined()
    expect(released).toEqual([0])
    // A fresh prime after the release is not handed the cancelled one.
    expect(await primePreview(m, 0, NORMAL)).toBe(true)
    expect(previewTimeline(m, 0)).toBeDefined()
    releasePreview(m)
  })
})
