/** 1d Task 5: §7.7 Stage-1 occlusion evidence (ink bounds + opaque-coverage grid). */
import { describe, expect, it } from 'bun:test'
import { createCanvas } from '@napi-rs/canvas'
import { OfficePackage } from '../../src/core/zip'
import { detectAndParse } from '../../src/components/loadOfficeSource'
import { getPaintables } from '../../src/render/paint'
import { hiddenFixtures } from '../../scripts/stream-s0/hidden-fixtures'
import { projectPackage } from '../src/policy'
import { occlusionObserver } from '../src/occlusion'
import { PAINT_DECORATION } from '../../src/core/text-recording'

/** Paint `f` as a painter-announced text decoration (the PAINT_DECORATION hook). */
function deco(c: CanvasRenderingContext2D, f: () => void) {
  const hook = (c as unknown as Record<symbol, ((on: boolean) => void) | undefined>)[PAINT_DECORATION]
  hook?.(true)
  try { f() } finally { hook?.(false) }
}

function run(paint: (ctx: CanvasRenderingContext2D) => void, w = 200, h = 100) {
  const ctx = createCanvas(w, h).getContext('2d') as unknown as CanvasRenderingContext2D
  const o = occlusionObserver(w, h)
  const wrapped = o.wrap(ctx)
  o.beforeUnit(wrapped, 0)
  paint(wrapped)
  o.afterUnit(wrapped, 0)
  return o.evidence()
}
const text = (ctx: CanvasRenderingContext2D, s = 'Hello', x = 20, y = 50) => { ctx.font = '20px sans-serif'; ctx.fillStyle = '#000'; ctx.fillText(s, x, y) }

describe('occlusion evidence', () => {
  it('text alone is servable', () => {
    expect(run((c) => text(c))).toMatchObject({ runs: 1, servable: 1, withheld: 0 })
  })
  it('text then an opaque rect over it is withheld (covered)', () => {
    const e = run((c) => { text(c); c.fillStyle = '#fff'; c.fillRect(10, 20, 120, 50) })
    expect(e).toMatchObject({ runs: 1, servable: 0, withheld: 1, reasons: { covered: 1 } })
  })
  it('an opaque rect painted BEFORE the text does not withhold it', () => {
    expect(run((c) => { c.fillStyle = '#fff'; c.fillRect(0, 0, 200, 100); text(c) })).toMatchObject({ servable: 1 })
  })
  it('the same rect at globalAlpha 0.5 leaves the text servable', () => {
    expect(run((c) => { text(c); c.globalAlpha = 0.5; c.fillStyle = '#fff'; c.fillRect(10, 20, 120, 50) })).toMatchObject({ servable: 1 })
  })
  it('a translucent fill colour does not cover', () => {
    expect(run((c) => { text(c); c.fillStyle = 'rgba(255,255,255,0.4)'; c.fillRect(10, 20, 120, 50) })).toMatchObject({ servable: 1 })
  })
  it('an opaque rect elsewhere does not withhold', () => {
    expect(run((c) => { text(c); c.fillStyle = '#fff'; c.fillRect(150, 0, 40, 100) })).toMatchObject({ servable: 1 })
  })
  it('text partly outside a rect clip is withheld (clipped)', () => {
    const e = run((c) => { c.save(); c.beginPath(); c.rect(0, 0, 40, 100); c.clip(); text(c); c.restore() })
    expect(e.reasons.clipped).toBe(1)
  })
  it('text fully inside a rect clip is servable; restore() drops the clip', () => {
    expect(run((c) => { c.save(); c.beginPath(); c.rect(0, 0, 200, 100); c.clip(); text(c); c.restore() })).toMatchObject({ servable: 1 })
    expect(run((c) => { c.save(); c.beginPath(); c.rect(0, 0, 40, 100); c.clip(); c.restore(); text(c) })).toMatchObject({ servable: 1 })
  })
  it('text under an arc clip is withheld (nonrect-clip)', () => {
    const e = run((c) => { c.beginPath(); c.arc(50, 50, 200, 0, Math.PI * 2); c.clip(); text(c) })
    expect(e.reasons['nonrect-clip']).toBe(1)
  })
  it('destination-out elsewhere is fine; over the ink it withholds (composite)', () => {
    expect(run((c) => { text(c); c.globalCompositeOperation = 'destination-out'; c.fillRect(150, 0, 40, 100) })).toMatchObject({ servable: 1 })
    expect(run((c) => { text(c); c.globalCompositeOperation = 'destination-out'; c.fillRect(10, 20, 120, 50) }).reasons.composite).toBe(1)
  })
  it('clearRect over the ink withholds (composite)', () => {
    expect(run((c) => { text(c); c.clearRect(0, 0, 200, 100) }).reasons.composite).toBe(1)
  })
  it('a filter over the ink withholds (filter)', () => {
    expect(run((c) => { text(c); c.filter = 'blur(2px)'; c.fillStyle = 'rgba(0,0,0,0.1)'; c.fillRect(10, 20, 120, 50) }).reasons.filter).toBe(1)
  })
  it('text drawn past the right edge is withheld (off-surface)', () => {
    expect(run((c) => text(c, 'Hello world, far too long', 150)).reasons['off-surface']).toBe(1)
  })
  it('fully transparent text is withheld (transparent)', () => {
    expect(run((c) => { c.globalAlpha = 0; text(c) }).reasons.transparent).toBe(1)
  })
  it('rotated text with nothing after it is servable', () => {
    expect(run((c) => { c.translate(100, 50); c.rotate(0.3); text(c, 'Hi', -10, 5) })).toMatchObject({ servable: 1 })
  })
  it('a later path fill over the ink withholds; its bounds come from the path points', () => {
    expect(run((c) => { text(c); c.beginPath(); c.moveTo(10, 20); c.lineTo(130, 20); c.lineTo(130, 70); c.closePath(); c.fill() }).reasons.covered).toBe(1)
    expect(run((c) => { text(c); c.beginPath(); c.moveTo(150, 0); c.lineTo(190, 0); c.lineTo(190, 90); c.closePath(); c.fill() })).toMatchObject({ servable: 1 })
  })
  it('gridlines stroked after the text in one path do not cover the cell between them', () => {
    const e = run((c) => {
      text(c, 'Hi', 30, 50)
      c.strokeStyle = '#d0d0d0'; c.lineWidth = 1; c.beginPath()
      c.moveTo(10.5, 0); c.lineTo(10.5, 100); c.moveTo(150.5, 0); c.lineTo(150.5, 100); c.moveTo(0, 10.5); c.lineTo(200, 10.5); c.moveTo(0, 90.5); c.lineTo(200, 90.5)
      c.stroke()
    })
    expect(e).toMatchObject({ servable: 1 })
  })
  it('strokeRect border around the text does not cover it; a line through it does', () => {
    expect(run((c) => { text(c, 'Hi', 30, 50); c.strokeRect(5, 5, 190, 90) })).toMatchObject({ servable: 1 })
    expect(run((c) => { text(c, 'Hi', 30, 50); c.lineWidth = 2; c.beginPath(); c.moveTo(0, 45); c.lineTo(200, 45); c.stroke() }).reasons.covered).toBe(1)
  })
  it('a later drawImage over the ink withholds', () => {
    const img = createCanvas(50, 50)
    expect(run((c) => { text(c); c.drawImage(img as never, 10, 20, 120, 50) }).reasons.covered).toBe(1)
  })
})

describe('occlusion evidence on the T2 hidden fixtures (visible-only)', async () => {
  const fixtures = await hiddenFixtures()
  // The control must never be withheld as covered/composited; a control that
  // overflows its (tiny) fixture page is legitimately withheld 'off-surface'.
  for (const f of fixtures) it(`${f.id}: runs > 0, control servable or geometrically cut, marker never servable`, async () => {
    const out = await projectPackage(await OfficePackage.load(f.withHidden), f.kind, 'visible-only')
    const units = await getPaintables(await detectAndParse(out.pkg))
    try {
      let runs = 0
      const servableText: string[] = [], withheld: Array<{ text: string; reason: string }> = []
      for (const p of units) {
        const w = Math.max(1, Math.ceil(p.spec.widthPx)), h = Math.max(1, Math.ceil(p.spec.heightPx))
        const ctx = createCanvas(w, h).getContext('2d') as unknown as CanvasRenderingContext2D
        const o = occlusionObserver(w, h)
        p.paint(o.wrap(ctx))
        const e = o.evidence()
        runs += e.runs
        servableText.push(...e.servableText); withheld.push(...e.withheldRuns)
      }
      expect(runs).toBeGreaterThan(0)
      const cut = withheld.filter((r) => r.text.includes(f.control))
      if (!servableText.join('').includes(f.control)) expect(cut.length).toBeGreaterThan(0)
      for (const r of cut) expect(['off-surface', 'clipped']).toContain(r.reason)
      expect(servableText.join('')).not.toContain(f.marker)
    } finally { units.dispose() }
  }, 30_000)
})

describe('the paint pass that serves a unit records evidence', () => {
  it('raster and command renders both carry occlusion evidence; the visible mark is not a run', async () => {
    const { DocumentHost } = await import('../src/render')
    const { readFileSync } = await import('node:fs')
    const host = new DocumentHost()
    const { docKey } = await host.open('occ', new Uint8Array(readFileSync(new URL('../../corpus/pypptx-test.pptx', import.meta.url))), 'visible-only')
    for (const projection of ['final-raster', 'commands'] as const) {
      const u = host.render(docKey, { unitIndex: 0, scale: 1, projection, mark: { kind: 'fixed' } })
      expect(u.evidence?.runs, projection).toBeGreaterThan(0)
      expect(u.evidence!.servableText.join('')).not.toContain('protected')
    }
  })
})

describe('runDetails (S3a Task 2)', () => {
  function details(paint: (ctx: CanvasRenderingContext2D) => void, w = 200, h = 100) {
    const ctx = createCanvas(w, h).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(w, h)
    paint(o.wrap(ctx))
    return { details: o.runDetails(), evidence: o.evidence() }
  }
  it('one entry per inked fillText, with device-px ink and the run text', () => {
    const { details: d } = details((c) => { text(c, 'Hello', 20, 50); text(c, 'World', 20, 80) })
    expect(d.map((r) => r.text)).toEqual(['Hello', 'World'])
    for (const r of d) expect(r).toMatchObject({ servable: true })
    expect(d[0]!.seq).toBeLessThan(d[1]!.seq)
    expect(d[0]!.reason).toBeUndefined()
    expect(d[0]!.ink.x).toBeGreaterThanOrEqual(19)
    expect(d[0]!.ink.y + d[0]!.ink.height).toBeLessThanOrEqual(56)
    expect(d[0]!.ink.width).toBeGreaterThan(0)
    expect(d[0]!.ink.height).toBeGreaterThan(0)
  })
  it("an opaque fillRect over the text ⇒ servable:false, reason:'covered'", () => {
    const { details: d } = details((c) => { text(c); c.fillStyle = '#fff'; c.fillRect(10, 20, 120, 50) })
    expect(d).toHaveLength(1)
    expect(d[0]).toMatchObject({ text: 'Hello', servable: false, reason: 'covered' })
  })
  it('ink is in device px (a scale-2 transform doubles it)', () => {
    const a = details((c) => text(c)).details[0]!
    const b = details((c) => { c.scale(2, 2); text(c) }, 400, 200).details[0]!
    expect(b.ink.x).toBeCloseTo(a.ink.x * 2, 0)
    expect(b.ink.width).toBeCloseTo(a.ink.width * 2, 0)
  })
  it('agrees with evidence() run for run', () => {
    const { details: d, evidence: e } = details((c) => {
      text(c, 'Kept', 20, 30); text(c, 'Gone', 20, 70); c.fillStyle = '#000'; c.fillRect(0, 50, 200, 30)
      c.globalAlpha = 0; text(c, 'Ghost', 120, 30)
    })
    expect(d).toHaveLength(e.runs)
    expect(d.filter((r) => r.servable).map((r) => r.text)).toEqual(e.servableText)
    expect(d.filter((r) => !r.servable).map((r) => ({ text: r.text, reason: r.reason }))).toEqual(e.withheldRuns)
  })
})

describe('stroke bounds', () => {
  it('a 1 px gridline 4 px below the ink does not cover it (miter allowance only at joins)', () => {
    const e = run((c) => { text(c, 'Hi', 30, 50); c.lineWidth = 1; c.beginPath(); c.moveTo(0, 54.5); c.lineTo(200, 54.5); c.stroke() })
    expect(e).toMatchObject({ servable: 1 })
  })
  it('a miter join next to the ink covers it', () => {
    const e = run((c) => { text(c, 'Hi', 30, 50); c.lineWidth = 4; c.lineJoin = 'miter'; c.miterLimit = 10; c.beginPath(); c.moveTo(0, 80); c.lineTo(35, 60); c.lineTo(70, 80); c.stroke() })
    expect(e.reasons.covered).toBe(1)
  })
})

describe('text decorations (S3a Task 3c)', () => {
  // 'Typing' has a descender, so an underline below the baseline overlaps its ink.
  const word = (c: CanvasRenderingContext2D) => text(c, 'Typing', 20, 50)
  const ink = (() => {
    const ctx = createCanvas(200, 100).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(200, 100)
    word(o.wrap(ctx))
    const i = o.runDetails()[0]!.ink
    return { ...i, x1: i.x + i.width, y1: i.y + i.height }
  })()
  /** An opaque bar across the word, `t` px thick, starting `at` px below the ink top. */
  const plainBar = (c: CanvasRenderingContext2D, at: number, t: number) => { c.fillStyle = '#000'; c.fillRect(ink.x - 2, ink.y + at, ink.width + 4, t) }
  /** The same bar announced by the painter as a text decoration. */
  const bar = (c: CanvasRenderingContext2D, at: number, t: number) => deco(c, () => plainBar(c, at, t))
  const verdict = (paint: (c: CanvasRenderingContext2D) => void, scale?: number) => {
    const ctx = createCanvas(200, 100).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(200, 100, undefined, scale)
    paint(o.wrap(ctx))
    const d = o.runDetails()
    expect(d).toHaveLength(1)
    return d[0]!
  }
  it('the fixture word is tall enough for the thresholds below to be meaningful', () => {
    expect(ink.height).toBeGreaterThan(15)
  })
  it('a 1 px underline fillRect below the baseline does not withhold the word', () => {
    // Baseline at y=50; the descender reaches below it.
    expect(ink.y1).toBeGreaterThan(52)
    expect(verdict((c) => { word(c); deco(c, () => { c.fillStyle = '#000'; c.fillRect(20, 51, ink.width, 1) }) })).toMatchObject({ servable: true })
  })
  it('a 1 px strike through the middle does not withhold the word', () => {
    expect(verdict((c) => { word(c); bar(c, ink.height / 2, 1) })).toMatchObject({ servable: true })
  })
  it('a stroke()d line path over the word (strokes are recorded as cover marks) does not withhold it', () => {
    // A single axis-aligned segment is recorded at its true width across the line: 1 px here.
    expect(verdict((c) => { word(c); deco(c, () => { c.strokeStyle = '#000'; c.lineWidth = 1; c.beginPath(); c.moveTo(ink.x, ink.y + ink.height / 2); c.lineTo(ink.x1, ink.y + ink.height / 2); c.stroke() }) })).toMatchObject({ servable: true })
  })
  it('underline + strike + overline together (≤ 25 % of ink height) still served', () => {
    expect(verdict((c) => { word(c); bar(c, 0, 1); bar(c, ink.height / 2, 1); bar(c, ink.height - 1, 1) })).toMatchObject({ servable: true })
  })
  it('repeated bars at one place merge: their union, not their sum, is measured', () => {
    expect(verdict((c) => { word(c); for (let i = 0; i < 10; i++) bar(c, ink.height / 2, 1) })).toMatchObject({ servable: true })
  })
  it('four thin bars totalling > 25 % of ink height ⇒ covered', () => {
    const t = ink.height * 0.1
    expect(verdict((c) => { word(c); for (const at of [0, 0.25, 0.5, 0.75]) bar(c, ink.height * at, t) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a bar thicker than 12 % of ink height (and > 1.5 px) ⇒ covered', () => {
    const t = Math.max(1.6, ink.height * 0.13)
    expect(verdict((c) => { word(c); bar(c, ink.height / 2, t) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('an opaque rect over the whole word ⇒ covered (unchanged)', () => {
    expect(verdict((c) => { word(c); c.fillStyle = '#fff'; c.fillRect(ink.x - 1, ink.y - 1, ink.width + 2, ink.height + 2) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a thin bar with a non-source-over composite or a filter still withholds (unchanged)', () => {
    expect(verdict((c) => { word(c); c.globalCompositeOperation = 'destination-out'; bar(c, ink.height / 2, 1) })).toMatchObject({ servable: false, reason: 'composite' })
    expect(verdict((c) => { word(c); c.clearRect(ink.x, ink.y + ink.height / 2, ink.width, 1) })).toMatchObject({ servable: false, reason: 'composite' })
    expect(verdict((c) => { word(c); c.filter = 'blur(1px)'; bar(c, ink.height / 2, 1) })).toMatchObject({ servable: false, reason: 'filter' })
  })
  it('a thin decoration does not excuse a separate thick cover', () => {
    expect(verdict((c) => { word(c); bar(c, ink.height / 2, 1); c.fillStyle = '#fff'; c.fillRect(ink.x, ink.y, ink.width / 3, ink.height) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('the 1.5 px floor scales with the render scale', () => {
    // Thicker than 12 % of the ink and than 1.5 px, but under 1.5 × 2 = 3 px.
    const t = Math.min(2.9, ink.height * 0.25)
    expect(t).toBeGreaterThan(Math.max(1.5, ink.height * 0.12))
    expect(verdict((c) => { word(c); bar(c, ink.height / 2, t) })).toMatchObject({ servable: false, reason: 'covered' })
    expect(verdict((c) => { word(c); bar(c, ink.height / 2, t) }, 2)).toMatchObject({ servable: true })
  })
  it('evidence() agrees: the decorated word is servable text', () => {
    expect(run((c) => { word(c); bar(c, ink.height / 2, 1) })).toMatchObject({ servable: 1, withheld: 0, servableText: ['Typing'] })
  })
  // Fix round 1: only painter-announced decorations are eligible.
  it('the observer\'s wrapped context defines PAINT_DECORATION', () => {
    const o = occlusionObserver(10, 10)
    const w = o.wrap(createCanvas(10, 10).getContext('2d') as unknown as CanvasRenderingContext2D) as unknown as Record<symbol, unknown>
    expect(typeof w[PAINT_DECORATION]).toBe('function')
  })
  it('a thin plain (unannounced) fillRect over the word is covered', () => {
    expect(verdict((c) => { word(c); plainBar(c, ink.height / 2, 1) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a thin plain stroke over the word is covered', () => {
    expect(verdict((c) => { word(c); c.strokeStyle = '#000'; c.lineWidth = 1; c.beginPath(); c.moveTo(ink.x, ink.y + ink.height / 2); c.lineTo(ink.x1, ink.y + ink.height / 2); c.stroke() })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('the flag ends with the hook: a plain bar after an announced one is covered', () => {
    expect(verdict((c) => { word(c); bar(c, ink.height / 2, 1); plainBar(c, 2, 1) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('review counterexample: a sparse run whose glyphs are fully painted out by plain fillRects is withheld', () => {
    const ctx = createCanvas(500, 100).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(500, 100)
    const c = o.wrap(ctx)
    c.fillStyle = '#fff'; c.fillRect(0, 0, 500, 100)
    c.font = '24px sans-serif'; c.fillStyle = '#000'; c.fillText('A' + ' '.repeat(30) + 'B', 20, 50)
    c.fillStyle = '#fff'; c.fillRect(20, 0, 16, 100); c.fillRect(237, 0, 14, 100)
    expect(o.runDetails()).toMatchObject([{ servable: false, reason: 'covered' }])
    expect(o.evidence().servableText).toEqual([])
  })

  // Ruling (a): exact stroke geometry for one straight axis-aligned segment.
  const thin = Math.max(1.5, 0.12 * ink.height)
  const line = (c: CanvasRenderingContext2D, lw: number, path: (c: CanvasRenderingContext2D) => void) => deco(c, () => { c.strokeStyle = '#000'; c.lineWidth = lw; c.beginPath(); path(c); c.stroke() })
  const mid = ink.y + ink.height / 2
  it('one axis-aligned stroke segment is measured at its true width, not lineWidth·√2', () => {
    const lw = thin * 0.9 // ≤ the threshold; lw·√2 is above it
    expect(lw * Math.SQRT2).toBeGreaterThan(thin)
    expect(verdict((c) => { word(c); line(c, lw, (p) => { p.moveTo(ink.x - 2, mid); p.lineTo(ink.x1 + 2, mid) }) })).toMatchObject({ servable: true })
    // A vertical segment through a wide word: its thin axis is x.
    expect(verdict((c) => { word(c); line(c, 1, (p) => { p.moveTo(ink.x + 10, ink.y - 2); p.lineTo(ink.x + 10, ink.y1 + 2) }) })).toMatchObject({ servable: true })
  })
  it('a true-width stroke thicker than the threshold still covers', () => {
    expect(verdict((c) => { word(c); line(c, thin * 1.1, (p) => { p.moveTo(ink.x - 2, mid); p.lineTo(ink.x1 + 2, mid) }) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('two segments (separate subpaths, no joins) keep the inflated box', () => {
    const lw = thin * 0.9
    expect(verdict((c) => { word(c); line(c, lw, (p) => { p.moveTo(ink.x - 2, mid); p.lineTo(ink.x + 20, mid); p.moveTo(ink.x + 20, mid); p.lineTo(ink.x1 + 2, mid) }) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a rotated (non-axis-aligned) segment keeps the inflated box', () => {
    const lw = thin * 0.9
    expect(verdict((c) => { word(c); c.translate(ink.x - 2, mid); c.rotate(0.001); line(c, lw, (p) => { p.moveTo(0, 0); p.lineTo(ink.width + 4, 0) }) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a curve or a multi-segment polyline keeps the inflated box', () => {
    const lw = thin * 0.9
    expect(verdict((c) => { word(c); line(c, lw, (p) => { p.moveTo(ink.x - 2, mid); p.quadraticCurveTo(ink.x + 20, mid + 0.5, ink.x1 + 2, mid) }) })).toMatchObject({ servable: false, reason: 'covered' })
    expect(verdict((c) => { word(c); c.lineJoin = 'round'; line(c, lw, (p) => { p.moveTo(ink.x - 2, mid); p.lineTo(ink.x + 20, mid); p.lineTo(ink.x1 + 2, mid) }) })).toMatchObject({ servable: false, reason: 'covered' })
  })
})

describe("text decorations measured along the mark's own thin axis (Task 3c ruling b)", () => {
  const inkOf = (s: string) => {
    const ctx = createCanvas(200, 100).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(200, 100)
    text(o.wrap(ctx), s, 20, 50)
    const i = o.runDetails()[0]!.ink
    return { ...i, x1: i.x + i.width, y1: i.y + i.height }
  }
  const verdictOf = (s: string, paint: (c: CanvasRenderingContext2D) => void) => {
    const ctx = createCanvas(200, 100).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(200, 100)
    const w = o.wrap(ctx)
    text(w, s, 20, 50); w.fillStyle = '#000'; deco(w, () => paint(w))
    return o.runDetails()[0]!
  }
  const I = inkOf('I')
  it('the I fixture is tall and narrow', () => {
    expect(I.height).toBeGreaterThan(I.width)
    expect(I.width).toBeLessThan(4)
  })
  it('a horizontal 1 px strike over a tall narrow run (I) is served', () => {
    expect(verdictOf('I', (c) => c.fillRect(I.x - 3, I.y + I.height / 2, I.width + 6, 1))).toMatchObject({ servable: true })
  })
  it('a thin vertical bar through I is covered (its union across the ink width passes 25 %)', () => {
    expect(verdictOf('I', (c) => c.fillRect(I.x + I.width / 2 - 0.5, I.y - 2, 1, I.height + 4))).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a square mark is never a decoration', () => {
    expect(verdictOf('I', (c) => c.fillRect(I.x, I.y + 2, 1, 1))).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('the two axes are tracked separately: one thin bar on each axis is served; either axis past 25 % covers', () => {
    const T = inkOf('Typing')
    const h = (c: CanvasRenderingContext2D, at: number, t: number) => c.fillRect(T.x - 2, T.y + at, T.width + 4, t)
    const v = (c: CanvasRenderingContext2D, at: number, t: number) => c.fillRect(T.x + at, T.y - 2, t, T.height + 4)
    expect(verdictOf('Typing', (c) => { h(c, T.height / 2, 1); v(c, 10, 1) })).toMatchObject({ servable: true })
    expect(verdictOf('Typing', (c) => { h(c, T.height / 2, 1); for (const at of [0, 0.25, 0.5, 0.75]) v(c, T.width * at, T.width * 0.1) })).toMatchObject({ servable: false, reason: 'covered' })
  })
})

describe('destructive composites (review 4)', () => {
  for (const op of ['copy', 'destination-in', 'source-in', 'source-out', 'destination-atop'] as const) it(`${op} far from the ink still withholds it (it clears outside its source)`, () => {
    const e = run((c) => { text(c); c.globalCompositeOperation = op; c.fillRect(170, 70, 20, 20) })
    expect(e.reasons.composite).toBe(1)
  })
  it('the same modes under a rect clip that excludes the ink leave it servable', () => {
    expect(run((c) => { text(c); c.save(); c.beginPath(); c.rect(150, 60, 50, 40); c.clip(); c.globalCompositeOperation = 'copy'; c.fillRect(170, 70, 20, 20); c.restore() })).toMatchObject({ servable: 1 })
  })
})
