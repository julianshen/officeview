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

/** A hook run inside the Morph's build (its matching): the round-2 clock test advances the show clock there. */
const morphHooks: { duringMatch?: () => void } = {}
vi.mock('../src/slideshow/morph/match', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/slideshow/morph/match')>()
  return { ...real, matchMorph: (...args: Parameters<typeof real.matchMorph>) => { morphHooks.duringMatch?.(); return real.matchMorph(...args) } }
})

/**
 * PR #20 review 3: `backendHooks.rebuild()` emulates a backend switch (a WebGL2 restore or the Canvas 2D
 * fallback): the old backend's uploads are released, as a switched-out backend releases its leases, and the
 * stage's `rebuildLayers` runs against the compositor. `backendHooks.onUpload` sees every layer the stage uploads.
 */
const backendHooks: { rebuild?: () => Promise<void>; onUpload?: (layer: { canvas: HTMLCanvasElement }) => void } = {}
vi.mock('../src/slideshow/render/backend', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/slideshow/render/backend')>()
  return {
    ...real,
    createBackendManager: (host: HTMLElement, deps: Parameters<typeof real.createBackendManager>[1]) => {
      const m = real.createBackendManager(host, deps)
      const uploaded = new Set<{ id: number }>()
      const bind = <T extends object>(t: T, k: string | symbol) => { const v = Reflect.get(t, k); return typeof v === 'function' ? v.bind(t) : v }
      const facade = new Proxy(m.current, {
        get(t, k) {
          if (k === 'upload') return (l: never, lease: never) => { backendHooks.onUpload?.(l); const r = t.upload(l, lease); if (!('refused' in r)) uploaded.add(r); return r }
          if (k === 'release') return (h: { id: number }) => { uploaded.delete(h); t.release(h) }
          return bind(t, k)
        },
      })
      backendHooks.rebuild = async () => {
        for (const h of uploaded) m.current.release(h)
        uploaded.clear()
        await deps.rebuildLayers(facade)
      }
      return new Proxy(m, { get(t, k) { return k === 'current' ? facade : bind(t, k) } })
    },
  }
})

import { runShow } from '../src/slideshow'
import { sessionDeps, type SessionController } from '../src/slideshow/player/session'
import { ManualClock } from '../src/slideshow/clock'
import { morphFrame, interpGeometry, type MorphFrameItem, type MorphFrameSpec } from '../src/slideshow/morph/render'
import type { Compositor, LayerDraw } from '../src/slideshow/types'
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
const morphOf = (option: 'byObject' | 'byWord' | 'byChar') => ({ xml: `<p:transition spd="slow" p14:dur="1200"><p159:morph xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" option="${option}"/></p:transition>` })
const MORPH = morphOf('byObject')

/** A click step whose entrance (appear) shows `spid` at once. */
const enterStep = (id: number, spid: number) =>
  `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 2}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:set><p:cBhvr><p:cTn id="${id + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
  `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`

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
  morphHooks.duringMatch = undefined
  backendHooks.onUpload = undefined
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
    // The morph route ran (not the fade look-alike), and the stage reports it as the morph (merge fix 87f0b6c).
    expect(r.s.stage!.transition).toEqual({ effect: 'morph', duration: 1200 })
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
    expect(r.s.stage!.transition).toBeUndefined()
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

  describe('PR #20 review: byWord and byChar Morph move the matched units', () => {
    /** Where the real `morphFrame` draws `item` at `progress`: the side's draw transforms. */
    const drawsOf = (item: MorphFrameItem, progress: number): LayerDraw[] => {
      let out: LayerDraw[] = []
      morphFrame({ pairs: [item] }, progress, { frame: (d: readonly LayerDraw[]) => { out = [...d] } } as unknown as Compositor)
      return out
    }
    const centre = (g: { x: number; y: number; w: number; h: number }) => [g.x + g.w / 2, g.y + g.h / 2] as const
    const apply = (m: readonly number[], x: number, y: number) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]] as const

    /** The matched unit `to` (from `from`) moves on its own path: a unit box on each side, drawn at the midpoint at p = 0.5. */
    function expectUnitMoves(m: { spec: MorphFrameSpec }, kinds: readonly string[], from: string, to: string, container: { from: number; to: number }) {
      expect(kinds).toContain(`morph:pair:${to}`)
      const item = m.spec.pairs.find(i => i.from?.key === from && i.to?.key === to)!
      expect(item).toBeDefined()
      expect(item.path).toBeDefined()
      // Tight unit boxes, not the 400 px container: the same word keeps its size (no stretch on the way).
      expect(item.path!.from.w).toBeLessThan(200)
      expect(Math.abs(item.path!.to.w - item.path!.from.w)).toBeLessThan(2)
      expect(item.from!.geometry).toEqual(item.path!.from)
      expect(item.to!.geometry).toEqual(item.path!.to)
      // Its own path: the unit moves by something other than its container's offset.
      const unitDx = item.path!.to.x - item.path!.from.x
      expect(Math.abs(unitDx - (container.to - container.from))).toBeGreaterThan(5)
      const mid = interpGeometry(item.path!.from, item.path!.to, 0.5)
      const [fx, fy] = centre(item.path!.from), [tx, ty] = centre(item.path!.to)
      expect(centre(mid)[0]).toBeCloseTo((fx + tx) / 2, 3)
      expect(centre(mid)[1]).toBeCloseTo((fy + ty) / 2, 3)
      // The real frame draws both sides with the unit's centre at the midpoint.
      const d = drawsOf(item, 0.5)
      expect(d.length).toBeGreaterThan(0)
      for (const draw of d) {
        const [cx, cy] = draw.handle === item.from!.handle ? [fx, fy] : [tx, ty]
        const [x, y] = apply(draw.state.transform, cx, cy)
        expect(x).toBeCloseTo((fx + tx) / 2, 3)
        expect(y).toBeCloseTo((fy + ty) / 2, 3)
      }
      return item
    }

    it('byWord: a word that changes place moves on its own path; the unmatched word fades', async () => {
      // Rule 2 pairs the box; LCS aligns "Alpha" (c:0-4 → c:5-9); "Beta" fades out and in.
      const doc = await deck({
        slides: [
          { shapes: [title(3, 100, 100, 'Alpha Beta')] },
          { shapes: [title(3, 140, 400, 'Beta Alpha')], transition: morphOf('byWord') },
        ],
      })
      const r = await show(doc)
      await r.key('ArrowRight')
      await r.tick(600)
      expect(r.s.stage!.transition).toEqual({ effect: 'morph', duration: 1200 })
      const m = lastMorph()!
      expect(m).toBeDefined()
      const kinds = r.s.stage!.inspect().lastDraws
      expectUnitMoves(m, kinds, 'sp:3/c:0-4', 'sp:3/c:5-9', { from: 100, to: 140 })
      // The unmatched word fades on each side.
      expect(kinds).toContain('morph:out:sp:3/c:6-9')
      expect(kinds).toContain('morph:in:sp:3/c:0-3')
      // ...moving with its container while it fades.
      const gone = m.spec.pairs.find(i => i.from?.key === 'sp:3/c:6-9' && !i.to)!
      expect(gone.path!.from.x).toBeCloseTo(100, 3)
      expect(gone.path!.to.x).toBeCloseTo(140, 3)
      expect(gone.from!.geometry).toEqual(gone.path!.from)
      // The container still pairs (its body: geometry, fill and line).
      expect(kinds).toContain('morph:pair:sp:3')
      expect(r.log).not.toContain('diag:transition-approximate:morph-fallback')
      expect(r.log).not.toContain('diag:morph-units-coarsened:sld')
      r.s.exit()
    })

    it('byChar: a letter that changes place moves on its own path', async () => {
      // "AB" → "xAB": A (c:0-0 → c:1-1) and B (c:1-1 → c:2-2) pair; x fades in.
      const doc = await deck({
        slides: [
          { shapes: [title(3, 100, 100, 'AB')] },
          { shapes: [title(3, 100, 300, 'xAB')], transition: morphOf('byChar') },
        ],
      })
      const r = await show(doc)
      await r.key('ArrowRight')
      await r.tick(600)
      const m = lastMorph()!
      expect(m).toBeDefined()
      const kinds = r.s.stage!.inspect().lastDraws
      expectUnitMoves(m, kinds, 'sp:3/c:0-0', 'sp:3/c:1-1', { from: 100, to: 100 })
      expect(kinds).toContain('morph:pair:sp:3/c:2-2')
      expect(kinds).toContain('morph:in:sp:3/c:0-0')
      r.s.exit()
    })

    it('byWord with the same text: the body pairs under, every word pairs, nothing draws twice', async () => {
      const doc = await deck({
        slides: [
          { shapes: [title(3, 100, 100, 'Alpha Beta')] },
          { shapes: [title(3, 300, 400, 'Alpha Beta')], transition: morphOf('byWord') },
        ],
      })
      const r = await show(doc)
      await r.key('ArrowRight')
      await r.tick(600)
      const m = lastMorph()!
      expect(m).toBeDefined()
      const kinds = r.s.stage!.inspect().lastDraws
      expect(kinds.filter(k => k === 'morph:pair:sp:3/c:0-4')).toHaveLength(1)
      expect(kinds.filter(k => k === 'morph:pair:sp:3/c:6-9')).toHaveLength(1)
      expect(kinds.filter(k => k === 'morph:pair:sp:3')).toHaveLength(1)
      expect(kinds.some(k => k.startsWith('morph:out:') || k.startsWith('morph:in:'))).toBe(false)
      const body = keyed(m.spec.pairs, 'sp:3')
      expect(body).toHaveLength(1)
      expect(body[0]).toMatchObject({ under: true })
      expect(body[0]!.from?.part).toBeUndefined()
      r.s.exit()
    })

    it('when the unit layers do not fit the workspace the boxes morph whole, diagnosed (not the fade look-alike)', async () => {
      // The workspace refuses every word layer (`layer:sp:3/c:…`); everything else fits.
      const real = sessionDeps.workspace
      vi.spyOn(sessionDeps, 'workspace').mockImplementation((profile) => {
        const w = real(profile)
        return new Proxy(w, {
          get(t, k) {
            if (k === 'lease') return (bytes: number, tag: string) => tag.includes('/c:') ? undefined : t.lease(bytes, tag)
            const v = Reflect.get(t, k)
            return typeof v === 'function' ? v.bind(t) : v
          },
        })
      })
      const doc = await deck({
        slides: [
          { shapes: [title(3, 100, 100, 'Alpha Beta')] },
          { shapes: [title(3, 140, 400, 'Beta Alpha')], transition: morphOf('byWord') },
        ],
      })
      const r = await show(doc)
      await r.key('ArrowRight')
      await r.tick(600)
      expect(r.s.stage!.transition).toEqual({ effect: 'morph', duration: 1200 })
      expect(r.log).toContain('diag:morph-units-coarsened:sld')
      expect(r.log).not.toContain('diag:transition-approximate:morph-fallback')
      const kinds = r.s.stage!.inspect().lastDraws
      expect(kinds).toContain('morph:pair:sp:3')
      expect(kinds.some(k => k.includes('sp:3/'))).toBe(false)
      r.s.exit()
    })

    it('past the Morph layer bound the units coarsen: the boxes morph whole, diagnosed', async () => {
      // 128 words a side split into ~130 layers each: more than MORPH_MAX_LAYERS (256) together.
      const words = Array.from({ length: 128 }, () => 'w').join(' ')
      const doc = await deck({
        slides: [
          { shapes: [title(3, 100, 100, words)] },
          { shapes: [title(3, 140, 300, words)], transition: morphOf('byWord') },
        ],
      })
      const r = await show(doc)
      await r.key('ArrowRight')
      await r.tick(600)
      expect(r.s.stage!.transition).toEqual({ effect: 'morph', duration: 1200 })
      expect(r.log).toContain('diag:morph-units-coarsened:sld')
      expect(r.log).not.toContain('diag:transition-approximate:morph-fallback')
      const kinds = r.s.stage!.inspect().lastDraws
      expect(kinds).toContain('morph:pair:sp:3')
      expect(kinds.some(k => k.startsWith('morph:pair:sp:3/'))).toBe(false)
      const m = lastMorph()!
      expect(keyed(m.spec.pairs, 'sp:3').some(i => i.from && i.to)).toBe(true)
      r.s.exit()
    })
  })

  describe('PR #20 review 2', () => {
    it('the Morph clock starts after its build: show time spent building does not eat the transition', async () => {
      const doc = await deck({
        slides: [
          { shapes: [square(80, 60), title(3, 240, 300, 'Morph me')] },
          { shapes: [square(700, 380), title(3, 240, 80, 'Morph me')], transition: MORPH },
        ],
      })
      const r = await show(doc)
      // The build takes 1 s of show time (matching, layout, rasterising).
      morphHooks.duringMatch = () => clock.advance(1000)
      await r.key('ArrowRight')
      morphHooks.duringMatch = undefined
      expect(r.s.state.slide).toBe(1)
      // The first morph frame is at the start of the transition, not 1 s into it.
      const first = vi.mocked(morphFrame).mock.calls[0]
      expect(first).toBeDefined()
      expect(first![1]).toBeLessThan(0.05)
      // It runs its full 1.2 s from there: still running at 1.1 s, done (and the slide's timeline on) after.
      await r.tick(1100)
      expect(r.s.stage!.transition).toEqual({ effect: 'morph', duration: 1200 })
      expect(lastMorph()!.progress).toBeGreaterThan(0.85)
      await r.tick(200)
      expect(r.s.stage!.transition).toBeUndefined()
      r.s.exit()
    })

    it('a matched word whose incoming box has not entered yet is not shown before its box enters', async () => {
      const doc = await deck({
        slides: [
          { shapes: [title(3, 100, 100, 'Alpha Beta')] },
          { shapes: [title(3, 140, 400, 'Beta Alpha')], transition: morphOf('byWord'), timingXml: timingXmlOf(mainSeqXml([enterStep(10, 3)])) },
        ],
      })
      const r = await show(doc)
      await r.key('ArrowRight')
      await r.tick(600)
      expect(r.s.state.slide).toBe(1)
      expect(r.s.stage!.transition).toEqual({ effect: 'morph', duration: 1200 })
      const m = lastMorph()!
      expect(m).toBeDefined()
      // Nothing of the incoming box draws (no pair, no fade-in of its words or body); the outgoing words fade out.
      expect(m.spec.pairs.some(i => i.to?.key?.startsWith('sp:3'))).toBe(false)
      const kinds = r.s.stage!.inspect().lastDraws
      expect(kinds.some(k => k.startsWith('morph:in:sp:3') || k.startsWith('morph:pair:sp:3'))).toBe(false)
      expect(kinds.some(k => k.startsWith('morph:out:sp:3'))).toBe(true)
      r.s.exit()
    })
  })

  describe('PR #20 review 3', () => {
    it('a backend rebuild during a byWord morph releases the Morph’s unit records: the workspace returns to its baseline', async () => {
      // Track the session workspace and its live leases by tag.
      const real = sessionDeps.workspace
      const live = new Map<string, number>()
      let ws: ReturnType<typeof real> | undefined
      vi.spyOn(sessionDeps, 'workspace').mockImplementation((profile) => {
        const w = real(profile)
        ws = w
        return new Proxy(w, {
          get(t, k) {
            if (k === 'lease') {
              return (bytes: number, tag: string) => {
                const l = t.lease(bytes, tag)
                if (!l) return l
                live.set(tag, (live.get(tag) ?? 0) + 1)
                const release = l.release.bind(l)
                let done = false
                l.release = () => { if (!done) { done = true; live.set(tag, live.get(tag)! - 1) } release() }
                return l
              }
            }
            const v = Reflect.get(t, k)
            return typeof v === 'function' ? v.bind(t) : v
          },
        })
      })
      const doc = await deck({
        slides: [
          { shapes: [title(3, 100, 100, 'Alpha Beta')] },
          { shapes: [title(3, 140, 400, 'Beta Alpha')], transition: morphOf('byWord') },
        ],
      })
      const r = await show(doc)
      await r.key('ArrowRight')
      await r.tick(300)
      expect(r.s.stage!.inspect().lastDraws).toContain('morph:pair:sp:3/c:5-9')
      const units = () => [...live].filter(([t, n]) => t.startsWith('morph-units') && n > 0).length
      expect(units()).toBe(1)
      // The backend switches mid-morph.
      await backendHooks.rebuild!()
      await r.tick(16)
      expect(units()).toBe(0)
      await r.tick(2000)
      r.s.exit()
      await new Promise((res) => setTimeout(res, 0))
      expect(units()).toBe(0)
      expect([...live].filter(([, n]) => n > 0)).toEqual([])
      expect(ws!.used()).toBe(0)
    })
  })

  describe('PR #20 review 4: text-build layers ride their container’s Morph path', () => {
    const box2 = (x: number, y: number): PptxShapeSpec => ({
      id: 3, name: 'Body', off: [px(x), px(y)], ext: [px(400), px(120)], paragraphs: [{ runs: [{ text: 'One' }] }, { runs: [{ text: 'Two' }] }],
    })
    /** A click step revealing paragraph `pr` of `spid` (a paragraph build). */
    const paraStep = (id: number, spid: number, pr: number) =>
      `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
      `<p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
      `<p:par><p:cTn id="${id + 2}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
      `<p:set><p:cBhvr><p:cTn id="${id + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"><p:txEl><p:pRg st="${pr}" end="${pr}"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
      `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
    const withBuild = (xml: string) => xml.replace('</p:tnLst></p:timing>', '</p:tnLst><p:bldLst><p:bldP spid="3" grpId="0" build="p"/></p:bldLst></p:timing>')
    /** A by-word entrance on `spid`, 1 ms apart. */
    const byWord = (id: number, spid: number) =>
      `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
      `<p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
      `<p:par><p:cTn id="${id + 2}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
      `<p:iterate type="wd"><p:tmAbs val="1"/></p:iterate><p:childTnLst>` +
      `<p:set><p:cBhvr><p:cTn id="${id + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
      `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`

    /** Every item drawing a layer of `prefix` sits on the container path (100,100) → (300,400): mapped to the box at p = 0.5. */
    function expectRides(m: { spec: MorphFrameSpec }, prefix: string) {
      const items = m.spec.pairs.filter(i => (i.from?.key ?? i.to?.key ?? '').startsWith(prefix))
      expect(items.length).toBeGreaterThan(0)
      for (const item of items) {
        expect(item.path, JSON.stringify(item.from?.key ?? item.to?.key)).toBeDefined()
        expect(item.path!.from.x).toBeCloseTo(100, 3)
        expect(item.path!.to.x).toBeCloseTo(300, 3)
        let draws: LayerDraw[] = []
        morphFrame({ pairs: [item] }, 0.5, { frame: (d: readonly LayerDraw[]) => { draws = d.flatMap(x => x.sum ? [x, x.sum.with] : [x]) } } as unknown as Compositor)
        expect(draws.length).toBeGreaterThan(0)
        for (const d of draws) {
          const side = d.handle === item.from?.handle ? item.from! : item.to!
          const g = side.geometry
          const [cx, cy] = [g.x + g.w / 2, g.y + g.h / 2]
          const t = d.state.transform
          // The container's centre lands mid-path: (100 + 200, 100 + 60) → (300 + 200, 400 + 60) at 0.5.
          expect(t[0] * cx + t[2] * cy + t[4]).toBeCloseTo(400, 3)
          expect(t[1] * cx + t[3] * cy + t[5]).toBeCloseTo(310, 3)
        }
      }
    }

    it('byObject with a paragraph build on the outgoing box: its paragraph layers move with the box', async () => {
      const doc = await deck({
        slides: [
          { shapes: [box2(100, 100)], timingXml: withBuild(timingXmlOf(mainSeqXml([paraStep(10, 3, 0), paraStep(20, 3, 1)]))) },
          { shapes: [box2(300, 400)], transition: MORPH },
        ],
      })
      const r = await show(doc)
      await r.key('ArrowRight')
      await r.key('ArrowRight')
      await r.tick(50)
      expect(r.s.state.slide).toBe(0)
      await r.key('ArrowRight')
      await r.tick(600)
      expect(r.s.state.slide).toBe(1)
      expect(r.s.stage!.transition).toEqual({ effect: 'morph', duration: 1200 })
      const m = lastMorph()!
      expect(r.s.stage!.inspect().lastDraws.some(k => k.includes('sp:3/p:'))).toBe(true)
      expectRides(m, 'sp:3/p:')
      r.s.exit()
    })

    it('byObject with a by-word animation split on the outgoing box: its word layers move with the box', async () => {
      const doc = await deck({
        slides: [
          { shapes: [box2(100, 100)], timingXml: timingXmlOf(mainSeqXml([byWord(10, 3)])) },
          { shapes: [box2(300, 400)], transition: MORPH },
        ],
      })
      const r = await show(doc)
      await r.key('ArrowRight')
      await r.tick(100)
      expect(r.s.state.slide).toBe(0)
      await r.key('ArrowRight')
      await r.tick(600)
      expect(r.s.state.slide).toBe(1)
      const m = lastMorph()!
      expect(r.s.stage!.inspect().lastDraws.some(k => k.includes('sp:3/c:'))).toBe(true)
      expectRides(m, 'sp:3/c:')
      r.s.exit()
    })

    it('the same paragraph build on both boxes: the outgoing built paragraphs ride; a not-yet-built incoming paragraph stays hidden', async () => {
      const timing = withBuild(timingXmlOf(mainSeqXml([paraStep(10, 3, 0), paraStep(20, 3, 1)])))
      const doc = await deck({
        slides: [
          { shapes: [box2(100, 100)], timingXml: timing },
          { shapes: [box2(300, 400)], transition: MORPH, timingXml: timing },
        ],
      })
      const r = await show(doc)
      await r.key('ArrowRight')
      await r.key('ArrowRight')
      await r.tick(50)
      await r.key('ArrowRight')
      await r.tick(600)
      expect(r.s.state.slide).toBe(1)
      const m = lastMorph()!
      // The incoming paragraphs are not built yet: none of them draws.
      expect(m.spec.pairs.some(i => i.to?.key?.startsWith('sp:3/p:'))).toBe(false)
      // The outgoing ones fade while riding the box's path.
      expectRides(m, 'sp:3/p:')
      expect(r.s.stage!.inspect().lastDraws).toContain('morph:pair:sp:3')
      r.s.exit()
    })

    it('the same visible paragraph range on both boxes pairs and rides the box (a paragraph exit not yet played)', async () => {
      const exitPara = (id: number, pr: number) => paraStep(id, 3, pr).replace('presetClass="entr"', 'presetClass="exit"').replace('<p:strVal val="visible"/>', '<p:strVal val="hidden"/>')
      const timing = withBuild(timingXmlOf(mainSeqXml([exitPara(10, 0), exitPara(20, 1)])))
      const doc = await deck({
        slides: [
          { shapes: [box2(100, 100)], timingXml: timing },
          { shapes: [box2(300, 400)], transition: MORPH, timingXml: timing },
        ],
      })
      const r = await show(doc)
      // Next with effects pending would play them: go to slide 2 instead.
      await r.key('2')
      await r.key('Enter')
      await r.tick(600)
      if (r.s.state.slide !== 1) { r.s.exit(); throw new Error(`slide ${r.s.state.slide}`) }
      const m = lastMorph()!
      const kinds = r.s.stage!.inspect().lastDraws
      expect(kinds).toContain('morph:pair:sp:3/p:0-0')
      expect(kinds).toContain('morph:pair:sp:3/p:1-1')
      const pair = m.spec.pairs.find(i => i.from?.key === 'sp:3/p:0-0' && i.to?.key === 'sp:3/p:0-0')!
      expect(pair).toMatchObject({ under: true })
      expectRides(m, 'sp:3/p:')
      r.s.exit()
    })
  })

  describe('rapid Next-Next: the outgoing side is the slide actually left', () => {
    /** A blue box (spid 4) at (360, 300); the first slide's exit hides it. */
    const blue = (): PptxShapeSpec => ({ id: 4, name: 'Blue', off: [px(360), px(300)], ext: [px(160), px(120)], fill: '2050C0' })

    /**
     * Presses Next, and Next again while the stage is still entering the slide the first one went to (its
     * `enter` waits for the slide to be drawable): the second slide change leaves a slide whose enter never
     * completed. `onUpload` sees the canvases the stage uploads while that slide is left.
     */
    async function rapidNext(r: Awaited<ReturnType<typeof show>>, onUpload?: (canvas: HTMLCanvasElement) => void) {
      const stage = r.s.stage!
      const host = document.querySelector<HTMLElement>('[data-officeview-show]')!
      const enter = stage.enter.bind(stage), leave = stage.leave.bind(stage)
      let pressed = false
      stage.leave = () => {
        if (pressed && onUpload) backendHooks.onUpload = (l) => onUpload(l.canvas)
        try { leave() } finally { backendHooks.onUpload = undefined }
      }
      stage.enter = (model, t, neighbours) => {
        const entering = enter(model, t, neighbours)
        if (model.index === 1 && !pressed) {
          pressed = true
          host.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true, cancelable: true }))
        }
        return entering
      }
      await r.key('ArrowRight')
      expect(pressed).toBe(true)
    }

    const slides = (third: PptxDeckSpec['slides'][number]['transition']): PptxDeckSpec['slides'] => [
      // The blue box exits on the first slide's only click.
      { shapes: [square(80, 60), blue()], timingXml: timingXmlOf(mainSeqXml([exitStep(10, 4)])) },
      { shapes: [square(80, 60), blue()] },
      { shapes: [square(700, 380), blue()], transition: third },
    ]

    it('the flat transition’s outgoing frame is the second slide (its blue box shown), not the first slide’s state', async () => {
      const doc = await deck({ slides: slides({ xml: '<p:transition spd="slow"><p:fade/></p:transition>' }) })
      const r = await show(doc)
      await r.key('ArrowRight') // plays the exit: the first slide's state hides spid 4
      await r.tick(100)
      expect(r.s.state.slide).toBe(0)
      const frames: HTMLCanvasElement[] = []
      await rapidNext(r, (c) => frames.push(c))
      await r.tick(100)
      expect(r.s.state.slide).toBe(2)
      expect(r.s.stage!.transition?.effect).toBe('fade')
      // The one upload made by the leave is the outgoing frame of the second slide.
      expect(frames).toHaveLength(1)
      const c = frames[0]!
      const k = c.width / 960
      const [red, green, b] = c.getContext('2d')!.getImageData(Math.round(440 * k), Math.round(360 * k), 1, 1).data
      expect([red, green, b]).toEqual([0x20, 0x50, 0xc0])
      r.s.exit()
    })

    it('the Morph’s outgoing state is the second slide’s: its blue box pairs instead of fading in', async () => {
      const doc = await deck({ slides: slides(MORPH) })
      const r = await show(doc)
      await r.key('ArrowRight')
      await r.tick(100)
      expect(r.s.state.slide).toBe(0)
      await rapidNext(r)
      await r.tick(300)
      expect(r.s.state.slide).toBe(2)
      expect(r.log).not.toContain('diag:transition-approximate:morph-fallback')
      const m = lastMorph()!
      expect(m).toBeDefined()
      // Shown on the slide left, so it is on the outgoing side and pairs.
      expect(keyed(m.spec.pairs, 'sp:4').some(i => i.from && i.to)).toBe(true)
      r.s.exit()
    })
  })
})
