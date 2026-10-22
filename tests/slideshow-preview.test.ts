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

afterEach(() => vi.restoreAllMocks())

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

    const p1 = nextCheckpoint(model, order, 0, tl.initial())
    if (p1 === 'end') throw new Error('precondition: a slide with a further step answers a position')
    expect(p1.slide).toBe(0)
    expect(p1.cursor).toMatchObject({ step: 1, mode: 'checkpoint' })
    // Checkpoint replay: the settle point plus the step group, all at the horizon.
    expect(p1.cursor.events).toHaveLength(3)

    // At the last step the next click leaves the slide: the next visible slide, at its settled checkpoint 0.
    const atLast = nextCheckpoint(model, order, 0, tl.checkpoint(2))
    expect(atLast).toEqual({ slide: 1, cursor: { step: 0, events: [], mode: 'checkpoint' } })

    // Past the visible order there is nowhere left: the end of the show.
    expect(nextCheckpoint(model, order, 1, tl.checkpoint(2))).toBe('end')

    // A slide whose timing never loaded answers the slide-advance position, still honestly.
    const d2 = await deck(timingXmlOf(mainSeqXml([click(3, '2')])))
    const { model: m2 } = await modelOf(d2, testWorkspace())
    expect(nextCheckpoint(m2, order, 0, { step: 0, events: [], mode: 'initial' })).toEqual({ slide: 1, cursor: { step: 0, events: [], mode: 'checkpoint' } })
    releasePreview(model)
  })

  it('zero-duration effects applied; indefinite at first-iteration end', async () => {
    const d = await deck(timingXmlOf(mainSeqXml([zeroJoinIndefinite(3)])))
    const { model } = await modelOf(d, testWorkspace())
    expect(await primePreview(model, 0, NORMAL)).toBe(true)
    const tl = previewTimeline(model, 0)!
    const pos = nextCheckpoint(model, [0], 0, tl.initial())
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

    expect(nextCheckpoint(model, [0, 1], 0, cursor)).not.toBe('end')

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
    // The letterbox bands match whatever the audience frame carries there (nothing here paints black).
    expect(pixels(50, 10)).toEqual([0, 0, 0])
    expect(pixels(50, 90)).toEqual([0, 0, 0])
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

  it('a refused presenter lease falls back to a lower DPR, then a placeholder', async () => {
    const sink = createDiagnosticSink()
    // The open default (480x270 at DPR 1, 518,400 B per surface) fits; the resize to
    // 800x600 at the capped DPR (3,840,000 B for both) does not, 700x525 at DPR 1.75 (2,940,000 B) does.
    const ws = createSlideShowWorkspace('desktop', { limit: 4_000_000 })
    const surfaces = createPresenterSurfaces(document, ws, sink)
    expect(surfaces.resize({ w: 400, h: 300 }, 2)).toBe(true)
    expect(surfaces.mirror.width).toBe(700)
    expect(surfaces.mirror.height).toBe(525)
    expect(surfaces.preview.width).toBe(700)
    expect(sink.list().filter(x => x.code === 'presenter-surface')).toEqual([])

    // Not even the open default fits both surfaces: the placeholder, with its diagnostic.
    const tiny = createSlideShowWorkspace('desktop', { limit: 900_000 })
    const s2 = createPresenterSurfaces(document, tiny, sink)
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
})
