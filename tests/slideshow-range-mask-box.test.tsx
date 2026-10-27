/**
 * A paragraph build's mask (blinds, wipe, …) sweeps PowerPoint's text box for the range, not the layer's padded
 * bitmap: from 1.3 em above the first baseline to 0.9 em below the last (fit to PowerPoint's blinds frames in
 * lo-tdf93868). jsdom with node-canvas: the real stage runs on the Canvas 2D compositor, with the show clock and
 * frames injected and driven. Decks are "Office-shaped synthetic".
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { runShow } from '../src/slideshow'
import { sessionDeps, type SessionController } from '../src/slideshow/player/session'
import { ManualClock } from '../src/slideshow/clock'
import type { ShowStart } from '../src/components/present'
import { buildPptxDeck, mainSeqXml, timingXmlOf, type PptxShapeSpec } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parsePptx, pptxPackageHandle } from '../src/pptx/parse'
import type { PptxDocument } from '../src/pptx/types'
import { buildRenderPlan } from '../src/slideshow/render/plan'
import { textClip, textRangeBox } from '../src/slideshow/render/scene'

const PX = 9525
// Shape at y 40 with the default 3.6 pt (4.8 px) top inset. PowerPoint's line model puts a 100% line's baseline
// 0.96 em below its top, so a lone 40 pt (53.33 px) paragraph's baseline is at 40 + 4.8 + 51.2 = 96.
const EM = 40 * 4 / 3, BASELINE = 40 + 4.8 + 0.96 * EM
const TOP = BASELINE - 1.3 * EM, BOTTOM = BASELINE + 0.9 * EM

const shape = (paragraphs: PptxShapeSpec['paragraphs']): PptxShapeSpec => ({
  id: 2, name: 'Text 2', off: [String(40 * PX), String(40 * PX)], ext: [String(880 * PX), String(400 * PX)], paragraphs,
})
const MS = 'MMMMMM MMMMMM MMMMMM'

/** A click step whose blinds(horizontal) entrance reveals paragraph `st..end` of `spid` over `dur` ms. */
const blinds = (id: number, spid: number, st: number, end: number, dur: number) => {
  const tgt = `<p:tgtEl><p:spTgt spid="${spid}"><p:txEl><p:pRg st="${st}" end="${end}"/></p:txEl></p:spTgt></p:tgtEl>`
  return `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${id + 2}" presetID="3" presetClass="entr" presetSubtype="10" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:set><p:cBhvr><p:cTn id="${id + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
    `<p:animEffect transition="in" filter="blinds(horizontal)"><p:cBhvr><p:cTn id="${id + 4}" dur="${dur}"/>${tgt}</p:cBhvr></p:animEffect>` +
    `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
}

async function deck(paragraphs: PptxShapeSpec['paragraphs'], timingXml?: string): Promise<PptxDocument> {
  return parsePptx(await OfficePackage.load(await buildPptxDeck({ slides: [{ shapes: [shape(paragraphs)], ...(timingXml ? { timingXml } : {}) }] })))
}

const measure = () => createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D

describe('textRangeBox', () => {
  it('spans 1.3 em above the first baseline to 0.9 em below the last, for the kept paragraphs only', async () => {
    const doc = await deck([{ runs: [{ text: 'Title', sz: '4000' }] }, { runs: [{ text: 'Body', sz: '2000' }] }])
    const slide = doc.slides[0]!
    const builds = [{ spid: '2', mode: 'paragraph' as const, ranges: [{ st: 0, end: 0, key: 'sp:2/p:0-0' }, { st: 1, end: 1, key: 'sp:2/p:1-1' }], animBg: false, body: 'sp:2' }]
    const plan = buildRenderPlan(slide, new Set(['sp:2/p:0-0', 'sp:2/p:1-1']), new Map(), builds)
    const range = (key: string) => {
      const layer = plan.layers.find(l => l.kind === 'target' && l.key === key && l.part === 'range')
      if (layer?.kind !== 'target') throw new Error(`no range layer ${key}`)
      return layer.nodes
    }
    const title = textRangeBox(range('sp:2/p:0-0'), slide, f => f, measure())!
    expect(title.minY).toBeCloseTo(TOP, 6)
    expect(title.maxY).toBeCloseTo(BOTTOM, 6)
    // The 20 pt (26.67 px) body line follows the 64 px title line: its baseline is 0.96 em below 40 + 4.8 + 64.
    const em = 20 * 4 / 3, baseline = 40 + 4.8 + 1.2 * EM + 0.96 * em
    const body = textRangeBox(range('sp:2/p:1-1'), slide, f => f, measure())!
    expect(body.minY).toBeCloseTo(baseline - 1.3 * em, 6)
    expect(body.maxY).toBeCloseTo(baseline + 0.9 * em, 6)
    // Horizontally: the shape's width, with the painted bounds' em/2 slack.
    expect(body.minX).toBeCloseTo(40 - em / 2, 6)
    expect(body.maxX).toBeCloseTo(920 + em / 2, 6)
  })

  it('leaves word and letter clones to their painted bounds', async () => {
    const doc = await deck([{ runs: [{ text: 'Alpha beta', sz: '4000' }] }])
    const slide = doc.slides[0]!
    const clone = { ...slide.shapes[0]! }
    textClip.set(clone, { items: [], key: 'sp:2/w:0' })
    expect(textRangeBox([clone], slide, f => f, measure())).toBeUndefined()
  })
})

// ---------------------------------------------------------------------------
// The live stage
// ---------------------------------------------------------------------------

let clock: ManualClock
beforeEach(() => {
  clock = new ManualClock()
  vi.spyOn(sessionDeps, 'clock').mockImplementation(() => clock)
  vi.spyOn(sessionDeps, 'frames').mockImplementation(() => () => {})
  vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
    return { left: 0, top: 0, right: 960, bottom: 720, width: 960, height: 720, x: 0, y: 0, toJSON() {} } as DOMRect
  })
})
afterEach(() => {
  vi.restoreAllMocks()
  document.body.replaceChildren()
})

const flush = async () => { for (let i = 0; i < 20; i++) await new Promise((r) => setTimeout(r, 0)) }

async function show(doc: PptxDocument): Promise<{ s: SessionController; host: HTMLElement; key(k: string): Promise<void>; tick(ms: number): Promise<void> }> {
  const host = document.createElement('div')
  host.tabIndex = -1
  host.setAttribute('data-officeview-show', '')
  host.style.visibility = 'visible'
  document.body.append(host)
  const ac = new AbortController()
  const s = await runShow({
    doc, handle: pptxPackageHandle(doc), host,
    token: { release() {}, released: false, owns: () => false },
    platform: { hasActivation: () => true, requestFullscreen: async () => {}, exitFullscreen: async () => {}, fullscreenElement: () => null, open: () => null },
    fullscreen: Promise.resolve('entered'), retryFullscreen: async () => 'entered',
    from: 'beginning', motion: 'auto',
    callbacks: {
      onStart: () => {}, onEnd: () => {}, onSlideChange: () => {}, onDiagnostic: () => {},
      currentSlide: () => 0, restoreFocus: () => {}, scrollToSlide: () => {},
    },
    signal: ac.signal, generation: 1, isCurrent: () => !ac.signal.aborted,
    cleanup: () => { ac.abort(); host.style.visibility = 'hidden' },
  } as ShowStart)
  await flush()
  s.tick()
  return {
    s, host,
    async key(k) { host.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true })); await flush(); s.tick() },
    async tick(ms) { clock.advance(ms); s.tick(); await flush(); s.tick() },
  }
}

describe('a paragraph build under blinds', () => {
  it('cuts its six bands from PowerPoint\'s text box', async () => {
    const doc = await deck([{ runs: [{ text: MS, sz: '4000' }] }], timingXmlOf(mainSeqXml([blinds(10, 2, 0, 0, 1000)])))
    const r = await show(doc)
    const output = () => {
      const c = r.host.querySelector('canvas') as unknown as HTMLCanvasElement
      return { k: c.width / 960, data: c.getContext('2d')!.getImageData(0, 0, c.width, c.height) }
    }
    const before = output()
    await r.key('ArrowRight')
    await r.tick(500)
    const half = output()
    await r.tick(1000)
    const full = output()
    const { k } = full
    // Per device row: how much of the paragraph's ink (full − before) the half-way frame shows.
    const ink = (img: ImageData, y: number) => { let s = 0; for (let x = 0; x < img.width; x++) { const i = (y * img.width + x) * 4; s += Math.abs(img.data[i] - before.data.data[i]) } return s }
    const band = (BOTTOM - TOP) / 6
    let checked = 0
    for (let y = 0; y < full.data.height; y++) {
      const total = ink(full.data, y)
      if (total < 255 * 4) continue
      // Slide-px span of this device row, and whether it lies wholly inside or outside the revealed band halves.
      const y0 = y / k, y1 = (y + 1) / k
      const at = (v: number) => (v - TOP) / band - Math.floor((v - TOP) / band)
      if (Math.floor((y0 - TOP) / band) !== Math.floor((y1 - TOP - 1e-9) / band)) continue
      const a = at(y0), b = at(y1 - 1e-9)
      const shown = ink(half.data, y) / total
      if (b <= 0.5) { expect(shown, `row ${y}`).toBeGreaterThan(0.9); checked++ }
      else if (a >= 0.5) { expect(shown, `row ${y}`).toBeLessThan(0.1); checked++ }
    }
    expect(checked).toBeGreaterThan(10)
    r.s.exit()
  })
})
