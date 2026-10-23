/**
 * S3-6 fix round 1: Morph through the real session and stage (spec 6.3).
 * jsdom with node-canvas: the stage runs on the Canvas 2D compositor, which
 * plays Morph too (transforms plus crossfade). The show clock and frames
 * are injected and driven. `morphFrame` is wrapped in a spy so the tests
 * see the frame spec the stage hands it. Decks are "Office-shaped synthetic".
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

vi.mock('../src/slideshow/morph/render', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/slideshow/morph/render')>()
  return { ...real, morphFrame: vi.fn(real.morphFrame) }
})

import { runShow } from '../src/slideshow'
import { sessionDeps, type SessionController } from '../src/slideshow/player/session'
import { ManualClock } from '../src/slideshow/clock'
import { morphFrame, type MorphFrameItem, type MorphFrameSpec } from '../src/slideshow/morph/render'
import type { ShowStart } from '../src/components/present'
import { buildPptxDeck, mainSeqXml, timingXmlOf, type PptxDeckSpec, type PptxShapeSpec } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parsePptx, pptxPackageHandle } from '../src/pptx/parse'
import type { PptxDocument } from '../src/pptx/types'

const px = (n: number) => String(n * 9525)
const square = (x: number, y: number, fill = 'D06040'): PptxShapeSpec => ({ id: 2, name: 'Square', off: [px(x), px(y)], ext: [px(120), px(120)], fill })
const title = (id: number, x: number, y: number, text: string): PptxShapeSpec => ({
  id, name: `Title ${id}`, off: [px(x), px(y)], ext: [px(400), px(80)], paragraphs: [{ runs: [{ text }] }],
})
const MORPH = { xml: '<p:transition spd="slow" p14:dur="1200"><p159:morph xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" option="byObject"/></p:transition>' }

/** A click step whose exit effect hides `spid` at once. */
const exitStep = (id: number, spid: number) =>
  `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 2}" presetID="1" presetClass="exit" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:set><p:cBhvr><p:cTn id="${id + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set>` +
  `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`

async function deck(spec: PptxDeckSpec): Promise<PptxDocument> {
  return parsePptx(await OfficePackage.load(await buildPptxDeck(spec)))
}

let clock: ManualClock
beforeEach(() => {
  clock = new ManualClock()
  vi.mocked(morphFrame).mockClear()
  vi.spyOn(sessionDeps, 'clock').mockImplementation(() => clock)
  vi.spyOn(sessionDeps, 'frames').mockImplementation(() => () => {})
  vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
    return { left: 0, top: 0, right: 1024, bottom: 768, width: 1024, height: 768, x: 0, y: 0, toJSON() {} } as DOMRect
  })
})
afterEach(() => {
  vi.restoreAllMocks()
  document.body.replaceChildren()
})

const flush = async () => { for (let i = 0; i < 20; i++) await new Promise((r) => setTimeout(r, 0)) }

async function show(doc: PptxDocument, motion: 'auto' | 'reduced' = 'auto'): Promise<{ s: SessionController; log: string[]; key(k: string): Promise<void>; tick(ms: number): Promise<void> }> {
  const host = document.createElement('div')
  host.tabIndex = -1
  host.setAttribute('data-officeview-show', '')
  host.style.visibility = 'visible'
  document.body.append(host)
  const ac = new AbortController()
  const log: string[] = []
  const s = await runShow({
    doc, handle: pptxPackageHandle(doc), host,
    token: { release() {}, released: false, owns: () => false },
    platform: { hasActivation: () => true, requestFullscreen: async () => {}, exitFullscreen: async () => {}, fullscreenElement: () => null, open: () => null },
    fullscreen: Promise.resolve('entered'), retryFullscreen: async () => 'entered',
    from: 'beginning', motion,
    callbacks: {
      onStart: () => {}, onEnd: () => {}, onSlideChange: () => {},
      onDiagnostic: (d: { code: string; effect?: string }) => log.push(`diag:${d.code}${d.effect ? `:${d.effect}` : ''}`),
      currentSlide: () => 0, restoreFocus: () => {}, scrollToSlide: () => {},
    },
    signal: ac.signal, generation: 1, isCurrent: () => !ac.signal.aborted,
    cleanup: () => { ac.abort(); host.style.visibility = 'hidden' },
  } as ShowStart)
  await flush()
  s.tick()
  return {
    s, log,
    async key(k) { host.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true })); await flush(); s.tick() },
    async tick(ms) { clock.advance(ms); s.tick(); await flush(); s.tick() },
  }
}

/** The spec of the last `morphFrame` call and its progress. */
function lastMorph(): { spec: MorphFrameSpec; progress: number } | undefined {
  const calls = vi.mocked(morphFrame).mock.calls
  const c = calls[calls.length - 1]
  return c && { spec: c[0], progress: c[1] }
}
const keyed = (items: readonly MorphFrameItem[], key: string) => items.filter(i => i.from?.key === key || i.to?.key === key)

describe('S3-6 morph through the session and stage', () => {
  it('Next onto a Morph slide runs the morph with both slides’ real geometry; unanimated shapes pair; the background goes under them', async () => {
    // Neither slide has any animation (the usual Morph deck).
    const doc = await deck({
      slides: [
        { shapes: [square(80, 60), title(3, 240, 300, 'Morph me')] },
        { shapes: [square(700, 380), title(3, 240, 80, 'Morph me')], transition: MORPH },
      ],
    })
    const r = await show(doc)
    expect(r.s.stage!.level).toBe('full')
    await r.key('ArrowRight')
    await r.tick(600)
    expect(r.s.state.slide).toBe(1)
    // The morph route ran (not the fade look-alike).
    expect(vi.mocked(morphFrame)).toHaveBeenCalled()
    expect(r.s.stage!.inspect().lastDraws.some(k => k.startsWith('morph'))).toBe(true)
    const m = lastMorph()!
    expect(m.progress).toBeGreaterThan(0)
    expect(m.progress).toBeLessThan(1)
    // The unanimated square pairs: one item carrying both sides, from (80,60) to (700,380).
    const sq = keyed(m.spec.pairs, 'sp:2')
    const pair = sq.find(i => i.from && i.to)!
    expect(pair).toBeDefined()
    expect(pair.from!.geometry.x).toBeCloseTo(80, 3)
    expect(pair.from!.geometry.y).toBeCloseTo(60, 3)
    expect(pair.to!.geometry.x).toBeCloseTo(700, 3)
    expect(pair.to!.geometry.y).toBeCloseTo(380, 3)
    // And the title pairs too.
    expect(keyed(m.spec.pairs, 'sp:3').some(i => i.from && i.to)).toBe(true)
    // Z order: the background pair comes first; every keyed shape paints above it.
    const bgIndex = m.spec.pairs.findIndex(i => i.background)
    expect(bgIndex).toBe(0)
    expect(m.spec.pairs.indexOf(pair)).toBeGreaterThan(bgIndex)
    // No fallback diagnostic.
    expect(r.log).not.toContain('diag:transition-approximate:morph-fallback')
    // After the 1.2 s window the slide draws normally again.
    await r.tick(1000)
    expect(r.s.stage!.inspect().lastDraws.some(k => k.startsWith('morph'))).toBe(false)
    r.s.exit()
  })

  it('the outgoing side is the slide as it was left: a shape that exited stays hidden', async () => {
    const doc = await deck({
      slides: [
        { shapes: [square(80, 60), title(4, 240, 300, 'Gone')], timingXml: timingXmlOf(mainSeqXml([exitStep(10, 4)])) },
        { shapes: [square(700, 380)], transition: MORPH },
      ],
    })
    const r = await show(doc)
    await r.key('ArrowRight') // plays the exit
    await r.tick(100)
    expect(r.s.state.slide).toBe(0)
    await r.key('ArrowRight') // Next slide: the morph
    await r.tick(300)
    expect(r.s.state.slide).toBe(1)
    const m = lastMorph()!
    expect(m).toBeDefined()
    expect(keyed(m.spec.pairs, 'sp:2').some(i => i.from && i.to)).toBe(true)
    // The exited title is not drawn as a fade-out.
    expect(keyed(m.spec.pairs, 'sp:4')).toEqual([])
    r.s.exit()
  })

  it('each pair blends by what changed: same shape under, changed text split (fill under, text crossfades), changed outline crossfades', async () => {
    const shape = (id: number, name: string, x: number, prst: string, text: string): PptxShapeSpec => ({
      id, name, prst, off: [px(x), px(40)], ext: [px(160), px(100)], fill: '3060C0', paragraphs: [{ runs: [{ text }] }],
    })
    const doc = await deck({
      slides: [
        { shapes: [shape(2, 'Same', 20, 'rect', 'Keep'), shape(3, 'Text', 220, 'rect', 'Before'), shape(4, 'Outline', 420, 'rect', 'Shape')] },
        { shapes: [shape(2, 'Same', 40, 'rect', 'Keep'), shape(3, 'Text', 240, 'rect', 'After'), shape(4, 'Outline', 440, 'ellipse', 'Shape')], transition: MORPH },
      ],
    })
    const r = await show(doc)
    await r.key('ArrowRight')
    await r.tick(300)
    const m = lastMorph()!
    expect(m).toBeDefined()
    // Same outline and content: one pair, the outgoing side held under.
    const same = keyed(m.spec.pairs, 'sp:2')
    expect(same).toHaveLength(1)
    expect(same[0]).toMatchObject({ under: true })
    expect(same[0]!.from && same[0]!.to).toBeTruthy()
    // Changed text: split into parts; fill blends under, the text crossfades.
    const text = keyed(m.spec.pairs, 'sp:3')
    const fill = text.find(i => i.to?.part === 'fill')
    const words = text.find(i => i.to?.part === 'text')
    expect(fill).toMatchObject({ under: true })
    expect(fill!.from?.part).toBe('fill')
    expect(words).toBeDefined()
    expect(words!.from?.part).toBe('text')
    expect(words!.under).toBeUndefined()
    // Changed outline: one crossfading pair.
    const outline = keyed(m.spec.pairs, 'sp:4')
    expect(outline).toHaveLength(1)
    expect(outline[0]!.from && outline[0]!.to).toBeTruthy()
    expect(outline[0]!.under).toBeUndefined()
    r.s.exit()
  })

  it('reduced motion plays the morph as a fade, never the morph route', async () => {
    const doc = await deck({
      slides: [
        { shapes: [square(80, 60)] },
        { shapes: [square(700, 380)], transition: MORPH },
      ],
    })
    const r = await show(doc, 'reduced')
    await r.key('ArrowRight')
    await r.tick(100)
    // Reduced motion won: the transition running is the fade, not a morph fallback.
    expect(r.s.stage!.transition?.effect).toBe('fade')
    expect(r.log).not.toContain('diag:transition-approximate:morph-fallback')
    await r.tick(200)
    expect(r.s.state.slide).toBe(1)
    expect(vi.mocked(morphFrame)).not.toHaveBeenCalled()
    r.s.exit()
  })
})
