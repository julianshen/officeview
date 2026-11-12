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
import { paintCellLayout } from '../../src/xlsx/render'
import { renderPages, type PageLayout, type Segment } from '../../src/docx/layout'
import type { CellBodyLayout } from '../../src/xlsx/view/textLayout'

/** The PAINT_DECORATION hook as the painters call it: 'text' opens a scope, 'line' starts its lines, false ends it. */
const phase = (c: CanvasRenderingContext2D, v: unknown) => (c as unknown as Record<symbol, ((v: unknown) => void) | undefined>)[PAINT_DECORATION]?.(v)
/** Paint `f` as the line phase of the open decoration scope. */
function deco(c: CanvasRenderingContext2D, f: () => void) { phase(c, 'line'); f() }

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
  /** Draws the word as the text of a decoration scope (its lines follow via deco/bar/line). */
  const word = (c: CanvasRenderingContext2D) => { phase(c, 'text'); text(c, 'Typing', 20, 50) }
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
    expect(verdict((c) => { word(c); bar(c, ink.height / 2, 1); phase(c, false); plainBar(c, 2, 1) })).toMatchObject({ servable: false, reason: 'covered' })
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

  // Fix round 2: a decoration exempts only the runs of its own scope.
  const other = (c: CanvasRenderingContext2D) => text(c, 'Elsewhere', 20, 90)
  const runOf = (paint: (c: CanvasRenderingContext2D) => void, t = 'Typing') => {
    const ctx = createCanvas(200, 100).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(200, 100)
    paint(o.wrap(ctx))
    return o.runDetails().find((r) => r.text === t)!
  }
  it("another scope's decoration over the word is ordinary cover", () => {
    expect(runOf((c) => { word(c); phase(c, false); phase(c, 'text'); other(c); bar(c, ink.height / 2, 1); phase(c, false) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a decoration over an unscoped word is ordinary cover', () => {
    expect(runOf((c) => { text(c, 'Typing', 20, 50); phase(c, 'text'); other(c); bar(c, ink.height / 2, 1); phase(c, false) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a line phase with no owned runs yields ordinary cover', () => {
    expect(runOf((c) => { text(c, 'Typing', 20, 50); phase(c, 'text'); bar(c, ink.height / 2, 1); phase(c, false) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a run drawn during the line phase is not owned', () => {
    expect(runOf((c) => { phase(c, 'text'); phase(c, 'line'); text(c, 'Typing', 20, 50); plainBar(c, ink.height / 2, 1); phase(c, false) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it("the scope's own decoration still exempts its run, and the other run is unaffected", () => {
    const ctx = createCanvas(200, 100).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(200, 100)
    const c = o.wrap(ctx)
    other(c); word(c); bar(c, ink.height / 2, 1); phase(c, false)
    expect(o.runDetails().map((r) => r.servable)).toEqual([true, true])
  })
  it('a scope may alternate text and line phases: a later segment\'s line exempts an earlier run only as a continuation of that run\'s own line (fix round 5)', () => {
    const both = (at: number) => {
      const ctx = createCanvas(200, 100).getContext('2d') as unknown as CanvasRenderingContext2D
      const o = occlusionObserver(200, 100)
      const c = o.wrap(ctx)
      // Own line over the left half of the word; the later segment's line carries on from where it ends.
      const mid = ink.x + ink.width / 2
      word(c); deco(c, () => { c.fillStyle = '#000'; c.fillRect(ink.x - 2, ink.y, mid - ink.x + 2, 1) })
      phase(c, 'text'); other(c); deco(c, () => c.fillRect(mid, ink.y + at, ink.x1 + 2 - mid, 1)); phase(c, false)
      return o.runDetails().map((r) => [r.text, r.servable])
    }
    expect(both(0)).toEqual([['Typing', true], ['Elsewhere', true]])
    // At another height it is the later segment's line only: ordinary cover for the earlier run.
    expect(both(ink.height / 2)).toEqual([['Typing', false], ['Elsewhere', true]])
  })
  it('malformed sequences fail closed', () => {
    const cases: Array<(c: CanvasRenderingContext2D) => void> = [
      (c) => { text(c, 'Typing', 20, 50); phase(c, 'line'); plainBar(c, ink.height / 2, 1); phase(c, false) }, // line without text
      (c) => { word(c); deco(c, () => plainBar(c, ink.height / 2, 1)); phase(c, 'bogus'); phase(c, false) }, // a bad call voids the scope's earlier marks
      (c) => { word(c); phase(c, false); phase(c, 'line'); plainBar(c, ink.height / 2, 1); phase(c, false) }, // line after false (no open scope)
      (c) => { phase(c, true); text(c, 'Typing', 20, 50); plainBar(c, ink.height / 2, 1); phase(c, false) }, // the old boolean form
      (c) => { word(c); phase(c, 'bogus'); deco(c, () => plainBar(c, ink.height / 2, 1)); phase(c, false) }, // unknown phase
    ]
    for (const paint of cases) expect(runOf(paint)).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('review counterexample through the real XLSX paintCellLayout: foreign strike decorations erasing a sparse run withhold it', () => {
    const ctx = createCanvas(500, 100).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(500, 100)
    const c = o.wrap(ctx)
    c.fillStyle = '#fff'; c.fillRect(0, 0, 500, 100)
    const extent = { width: 500, height: 100 }
    const cell = (text: string, color: string, decoration: { strike?: boolean; underline?: 'single'; sizePx: number }): CellBodyLayout =>
      ({ x: 0, y: 0, w: 500, h: 100, formattedText: text, text: { kind: 'plain', text, font: '24px sans-serif', color, tx: 0, ty: 0, width: 100, decoration } }) as CellBodyLayout
    const victim = cell('A' + ' '.repeat(30) + 'B', '#000', { underline: 'single', sizePx: 24 })
    ;(victim.text as { tx: number; ty: number }).tx = 20; (victim.text as { tx: number; ty: number }).ty = 50
    paintCellLayout(c, victim, extent, 'content')
    // Strike at sizePx 256 is a 16 px bar at y −69..−53; rotated by π/2 it is
    // a vertical bar at x 53..69 (+ translation), y 0..100.
    for (const tx of [-33, 184]) {
      c.save(); c.translate(tx, 0); c.rotate(Math.PI / 2)
      paintCellLayout(c, cell(' ', '#fff', { strike: true, sizePx: 256 }), extent, 'content')
      c.restore()
    }
    // No glyph pixel remains: everything above the baseline is blank (the
    // victim's own underline sits below it, at y 53–55).
    const px = ctx.getImageData(0, 0, 500, 51).data
    let dark = 0
    for (let i = 0; i < px.length; i += 4) if (px[i]! < 128) dark++
    expect(dark).toBe(0)
    const d = o.runDetails()
    expect(d).toHaveLength(1)
    expect(d[0]).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('the real XLSX paintCellLayout: a cell\'s own strike exempts its own text', () => {
    const ctx = createCanvas(300, 100).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(300, 100)
    const cellText = { kind: 'plain', text: 'some more', font: '24px sans-serif', color: '#000', tx: 20, ty: 50, width: 0, decoration: { strike: true, sizePx: 16 } }
    const m = (createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D); m.font = cellText.font; cellText.width = m.measureText(cellText.text).width
    paintCellLayout(o.wrap(ctx), { x: 0, y: 0, w: 300, h: 100, formattedText: cellText.text, text: cellText } as CellBodyLayout, { width: 300, height: 100 }, 'content')
    expect(o.runDetails()).toMatchObject([{ text: 'some more', servable: true }])
  })

  // Ruling (a): exact stroke geometry for one straight axis-aligned segment.
  const thin = Math.max(1.5, 0.12 * ink.height)
  const line = (c: CanvasRenderingContext2D, lw: number, path: (c: CanvasRenderingContext2D) => void) => deco(c, () => { c.strokeStyle = '#000'; c.lineWidth = lw; c.beginPath(); path(c); c.stroke() })
  const mid = ink.y + ink.height / 2
  it('one axis-aligned stroke segment is measured at its true width, not lineWidth·√2', () => {
    const lw = thin * 0.9 // ≤ the threshold; lw·√2 is above it
    expect(lw * Math.SQRT2).toBeGreaterThan(thin)
    expect(verdict((c) => { word(c); line(c, lw, (p) => { p.moveTo(ink.x - 2, mid); p.lineTo(ink.x1 + 2, mid) }) })).toMatchObject({ servable: true })
    // A vertical segment through a horizontal word is perpendicular to its baseline: ordinary cover (fix round 3).
    expect(verdict((c) => { word(c); line(c, 1, (p) => { p.moveTo(ink.x + 10, ink.y - 2); p.lineTo(ink.x + 10, ink.y1 + 2) }) })).toMatchObject({ servable: false, reason: 'covered' })
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

describe('text decorations judged across the run baseline (Task 3c ruling b, fix round 3)', () => {
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
    phase(w, 'text'); text(w, s, 20, 50); w.fillStyle = '#000'; deco(w, () => paint(w)); phase(w, false)
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
  // Fix round 3: one axis per run — perpendicular to its baseline.
  it('a thin vertical bar through horizontal text is covered even inside the same scope', () => {
    const T = inkOf('Typing')
    expect(verdictOf('Typing', (c) => c.fillRect(T.x + 10, T.y - 2, 1, T.height + 4))).toMatchObject({ servable: false, reason: 'covered' })
    expect(verdictOf('Typing', (c) => { c.fillRect(T.x - 2, T.y + T.height / 2, T.width + 4, 1); c.fillRect(T.x + 10, T.y - 2, 1, T.height + 4) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a mark not elongated along the baseline (side along < 2 × side across) is covered', () => {
    const T = inkOf('Typing')
    // 1.9 px long, 1 px thick: thin enough, but not a line along the baseline.
    expect(verdictOf('Typing', (c) => c.fillRect(T.x + 10, T.y + T.height / 2, 1.9, 1))).toMatchObject({ servable: false, reason: 'covered' })
    expect(verdictOf('Typing', (c) => c.fillRect(T.x + 10, T.y + T.height / 2, 2.1, 1))).toMatchObject({ servable: true })
  })
  it('decorated rotated (non-axis-aligned) text is withheld', () => {
    const ctx = createCanvas(200, 100).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(200, 100)
    const c = o.wrap(ctx)
    c.translate(60, 50); c.rotate(0.3)
    phase(c, 'text'); text(c, 'Hi', -10, 5)
    phase(c, 'line'); c.fillStyle = '#000'; c.fillRect(-12, 0, 30, 1); phase(c, false)
    expect(o.runDetails()).toMatchObject([{ servable: false, reason: 'covered' }])
  })
  it('vertical text (90° transform) with a thin line parallel to its baseline is served; one across it is covered', () => {
    const vert = (across: boolean) => {
      const ctx = createCanvas(200, 200).getContext('2d') as unknown as CanvasRenderingContext2D
      const o = occlusionObserver(200, 200)
      const c = o.wrap(ctx)
      c.translate(100, 20); c.rotate(Math.PI / 2)
      phase(c, 'text'); text(c, 'Typing', 0, 0)
      const i = o.runDetails()[0]!.ink
      c.setTransform(1, 0, 0, 1, 0, 0)
      phase(c, 'line'); c.fillStyle = '#000'
      if (across) c.fillRect(i.x - 2, i.y + 10, i.width + 4, 1)
      else c.fillRect(i.x + i.width / 2, i.y - 2, 1, i.height + 4)
      phase(c, false)
      expect(i.height).toBeGreaterThan(i.width)
      return o.runDetails()[0]!
    }
    expect(vert(false)).toMatchObject({ servable: true })
    expect(vert(true)).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('review probe through the real DOCX renderPages: same-scope rotated struck segments erasing a sparse run withhold it', () => {
    const ctx = createCanvas(500, 100).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(500, 100)
    const c = o.wrap(ctx)
    c.fillStyle = '#fff'; c.fillRect(0, 0, 500, 100)
    const style = (pt: number) => ({ fontFamily: 'Arial', fontSizePt: pt, bold: false, italic: false })
    const seg = (text: string, pt: number, color: string, deco: 'underline' | 'strike', e: number, f: number, rotated: boolean, widthPx: number): Segment =>
      ({ text, run: { text, color, [deco]: true }, style: style(pt), widthPx, transform: rotated ? { a: 0, b: 1, c: -1, d: 0, e, f } : { a: 1, b: 0, c: 0, d: 1, e, f } }) as Segment
    const victim = 'A' + ' '.repeat(30) + 'B'
    const page = { widthPx: 500, heightPx: 100, tables: [], images: [], lines: [{
      yPx: 0, xPx: 0, widthPx: 500, align: 'left', isParagraphEnd: true, heightPx: 100, contentWidthPx: 500,
      segs: [seg(victim, 18, '000000', 'underline', 20, 50, false, 240), seg(' ', 800 / 3, 'FFFFFF', 'strike', -52, 0, true, 100), seg(' ', 800 / 3, 'FFFFFF', 'strike', 165, 0, true, 100)],
    }] } as unknown as PageLayout
    renderPages([page], c)
    const px = ctx.getImageData(0, 0, 500, 50).data
    let dark = 0
    for (let i = 0; i < px.length; i += 4) if (px[i]! < 128) dark++
    expect(dark).toBe(0)
    expect(o.runDetails()).toMatchObject([{ text: victim, servable: false, reason: 'covered' }])
  })
})

describe('decoration thresholds are bounded by the run font size (Task 3c fix round 4)', () => {
  // R = min(ink extent across the baseline, font size in device px).
  /** One scoped run `s` in `font`, then scoped bars `[offset below the ink top, thickness]`; `pre` sets the transform. */
  const probe = (s: string, font: string, bars: (ink: { y: number; height: number }) => Array<[number, number]>, opts: { pre?: (c: CanvasRenderingContext2D) => void; scale?: number; setFont?: (c: CanvasRenderingContext2D) => void } = {}) => {
    const ctx = createCanvas(400, 600).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(400, 600, undefined, opts.scale)
    const c = o.wrap(ctx)
    opts.pre?.(c)
    phase(c, 'text'); c.font = font; opts.setFont?.(c); c.fillStyle = '#000'; c.fillText(s, 20, 250)
    const i = o.runDetails()[0]!.ink
    c.setTransform(1, 0, 0, 1, 0, 0)
    phase(c, 'line')
    for (const [at, t] of bars(i)) c.fillRect(i.x - 2, i.y + at, i.width + 4, t)
    phase(c, false)
    return { ...o.runDetails()[0]!, ink: i }
  }
  const tall = 'secret' + '̍'.repeat(40)
  it('a font size smaller than the ink extent bounds the thickness limit', () => {
    // 20 px font, ink > 150 px: the limit is 0.12 × 20 = 2.4 px, not 12 % of the ink.
    const base = probe(tall, '20px sans-serif', () => [])
    expect(base.ink.height).toBeGreaterThan(150)
    expect(probe(tall, '20px sans-serif', (i) => [[i.height - 8, 2.3]])).toMatchObject({ servable: true })
    expect(probe(tall, '20px sans-serif', (i) => [[i.height - 8, 2.5]])).toMatchObject({ servable: false, reason: 'covered' })
    expect(probe(tall, '20px sans-serif', (i) => [[i.height - 12, 10]])).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a font size smaller than the ink extent bounds the union cap', () => {
    // Cap 0.25 × 20 = 5 px: two 2.3 px bars pass, three (6.9 px) do not — though far under 25 % of the ink.
    expect(probe(tall, '20px sans-serif', (i) => [[10, 2.3], [i.height - 8, 2.3]])).toMatchObject({ servable: true })
    expect(probe(tall, '20px sans-serif', (i) => [[10, 2.3], [100, 2.3], [i.height - 8, 2.3]])).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('an ink extent smaller than the font size still bounds the limit (unchanged)', () => {
    // 'a' in a 100 px font: ink ≈ 55 px, so 12 % of the ink (< 12 px), not of the font size.
    const a = probe('a', '100px sans-serif', () => [])
    expect(a.ink.height).toBeLessThan(80)
    expect(probe('a', '100px sans-serif', (i) => [[i.height / 2, 0.11 * i.height]])).toMatchObject({ servable: true })
    expect(probe('a', '100px sans-serif', (i) => [[i.height / 2, 0.13 * i.height]])).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a run whose font size cannot be determined gets no exemption', () => {
    // The same 1 px strike is served with a readable font and covered when ctx.font does not yield a px size.
    const unreadable = (value: unknown) => (c: CanvasRenderingContext2D) => Object.defineProperty(c, 'font', { get: () => value, configurable: true })
    expect(probe('Typing', '20px sans-serif', (i) => [[i.height / 2, 1]])).toMatchObject({ servable: true })
    for (const v of ['sans-serif', '', undefined, '0px sans-serif', 'NaNpx sans-serif', '12pt sans-serif', 20])
      expect(probe('Typing', '20px sans-serif', (i) => [[i.height / 2, 1]], { setFont: unreadable(v) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a px-like token in the family name is not the font size', () => {
    // The size is the first px token of the serialised font; the family follows it.
    const font = '20px "x 9999px y", sans-serif'
    expect(probe(tall, font, (i) => [[i.height - 8, 2.3]])).toMatchObject({ servable: true })
    expect(probe(tall, font, (i) => [[i.height - 8, 2.5]])).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a scaled transform (scale 2) scales the font-size bound', () => {
    // 20 px font under scale(2): 40 device px, limit 4.8 px, cap 10 px.
    const at = (c: CanvasRenderingContext2D) => c.scale(2, 2)
    expect(probe(tall, '20px sans-serif', () => [], { pre: at }).ink.height).toBeGreaterThan(300)
    expect(probe(tall, '20px sans-serif', (i) => [[i.height - 16, 4.7]], { pre: at })).toMatchObject({ servable: true })
    expect(probe(tall, '20px sans-serif', (i) => [[i.height - 16, 4.9]], { pre: at })).toMatchObject({ servable: false, reason: 'covered' })
    expect(probe(tall, '20px sans-serif', (i) => [[10, 4.7], [i.height - 16, 4.7]], { pre: at })).toMatchObject({ servable: true })
    expect(probe(tall, '20px sans-serif', (i) => [[10, 4.7], [100, 4.7], [i.height - 16, 4.7]], { pre: at })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('only the scale across the baseline counts: scale(3, 1) leaves the bound at the font size', () => {
    const pre = (c: CanvasRenderingContext2D) => c.scale(3, 1)
    expect(probe('secret̍̍̍̍̍̍̍̍', '20px sans-serif', (i) => [[i.height - 8, 2.3]], { pre })).toMatchObject({ servable: true })
    expect(probe('secret̍̍̍̍̍̍̍̍', '20px sans-serif', (i) => [[i.height - 8, 2.5]], { pre })).toMatchObject({ servable: false, reason: 'covered' })
  })
})

describe('only a run\'s own lines and their continuations are decorations (Task 3c fix round 5)', () => {
  const W = 600, H = 600
  const style = (pt: number) => ({ fontFamily: 'sans-serif', fontSizePt: pt, bold: false, italic: false })
  const measure = (s: string, pt = 18) => { const c = createCanvas(10, 10).getContext('2d'); c.font = `${pt}pt sans-serif`; return c.measureText(s) }
  const placed = (text: string, pt: number, color: string, deco: 'underline' | 'strike', e: number, f: number, widthPx: number): Segment =>
    ({ text, run: { text, color, [deco]: true }, style: style(pt), widthPx, transform: { a: 1, b: 0, c: 0, d: 1, e, f } }) as Segment
  const render = (segs: Segment[], line: Record<string, unknown> = {}) => {
    const ctx = createCanvas(W, H).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(W, H)
    const c = o.wrap(ctx)
    c.fillStyle = '#fff'; c.fillRect(0, 0, W, H)
    renderPages([{ widthPx: W, heightPx: H, tables: [], images: [], lines: [{ yPx: 0, xPx: 0, widthPx: W, align: 'left', isParagraphEnd: true, heightPx: H, contentWidthPx: W, segs, ...line }] } as unknown as PageLayout], c)
    return { ctx, runs: o.runDetails() }
  }
  /** Black underlined 18 pt `word` at baseline 500, with two adjacent white strikes (from `pt`-sized struck spaces of the same stretch) centred on y = `mid`. */
  const arrangement = (word: string, base: string, pt: number, mid: number, lines = true) => {
    const lw = Math.max(1, pt * 0.06), b = measure(base)
    const segs = [placed(word, 18, '000000', 'underline', 40, 500, measure(word).width), ...(lines ? [mid - lw / 2, mid + lw / 2].map((y) => placed(' ', pt, 'FFFFFF', 'strike', 20, y + 0.3 * pt, b.width + 40)) : [])]
    const { ctx, runs } = render(segs)
    const y0 = Math.floor(500 - b.actualBoundingBoxAscent), h = Math.ceil(b.actualBoundingBoxAscent + b.actualBoundingBoxDescent)
    const px = ctx.getImageData(40, y0, Math.ceil(b.width), h).data
    let dark = 0
    for (let i = 0; i < px.length; i += 4) if (px[i]! < 160) dark++
    return { run: runs.find((r) => r.text === word)!, dark }
  }
  it('real DOCX renderPages: two neighbour strikes erasing superscript letters beside a capital withhold the run', () => {
    // ˢᵉᶜʳᵉᵗ are ~8 px tall at 24 px; two 2.75 px lines (each ≤ 12 % of the 23 px ink, 5.5 px together ≤ 25 %) erase all but their edge rows.
    const clean = arrangement('ˢᵉᶜʳᵉᵗÉ', 'ˢᵉᶜʳᵉᵗ', 45.8, 490, false), hit = arrangement('ˢᵉᶜʳᵉᵗÉ', 'ˢᵉᶜʳᵉᵗ', 45.8, 490)
    expect(clean.run).toMatchObject({ servable: true })
    expect(hit.dark).toBeLessThan(clean.dark * 0.4)
    expect(hit.run).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('real DOCX renderPages: two neighbour strikes erasing superscript digits under stacked combining marks withhold the run', () => {
    const word = '¹²³' + '̍'.repeat(40)
    const clean = arrangement(word, '¹²³', 47.8, 487.5, false), hit = arrangement(word, '¹²³', 47.8, 487.5)
    expect(clean.run).toMatchObject({ servable: true })
    expect(hit.dark).toBeLessThan(clean.dark * 0.5)
    expect(hit.run).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('real DOCX renderPages: an underlined word, space and word in one stretch are all served (the line continues under the neighbours)', () => {
    const seg = (text: string): Segment => ({ text, run: { text, color: '000000', underline: true }, style: style(18), widthPx: measure(text).width }) as Segment
    const { runs } = render([seg('typing'), seg(' '), seg('guide')], { xPx: 40, baselinePx: 300 })
    expect(runs.map((r) => [r.text, r.servable])).toEqual([['typing', true], ['guide', true]])
  })

  // Final review, finding 1 (residual A): a 6 pt struck word (8 px font, 1 px own line) and two white struck
  // spaces of the same stretch at ±0.45 px, each spanning the word: the three lines band to about 1.9 px.
  const reviewer = (word: string, hit: boolean, shift = 0.45) => {
    const seg = (text: string, color: string, f: number): Segment => ({ text, run: { text, color, strike: true }, style: style(6), widthPx: 40, transform: { a: 1, b: 0, c: 0, d: 1, e: 20, f } }) as Segment
    const { ctx, runs } = render([seg(word, '000000', 80), ...(hit ? [-shift, shift].map((d) => seg(' ', 'FFFFFF', 80 + d)) : [])])
    const px = ctx.getImageData(15, 65, 60, 20).data
    let dark = 0
    for (let i = 0; i < px.length; i += 4) if (px[i]! < 160) dark++
    return { run: runs.find((r) => r.text === word)!, dark, count: runs.length }
  }
  for (const word of ['secretÉ', 'ˢᵉᶜʳᵉᵗÉ']) it(`final review repro through the real DOCX renderPages: 6 pt struck ${word} under two near-collinear white strikes is covered; its control is served`, () => {
    const clean = reviewer(word, false), hit = reviewer(word, true)
    expect(clean.run).toMatchObject({ servable: true })
    expect(clean.dark).toBeGreaterThan(0)
    expect(hit.dark).toBeLessThan(clean.dark * 0.6)
    expect(hit.count).toBe(1)
    expect(hit.run).toMatchObject({ servable: false, reason: 'covered' })
    // Exactly collinear too: a foreign line lying on the own line is cover whatever its offset.
    expect(reviewer(word, true, 0).run).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('real DOCX renderPages: a struck word, space and word placed by transforms (abutting segments) are all served', () => {
    const w = (s: string) => measure(s, 18).width
    const seg = (text: string, e: number): Segment => ({ text, run: { text, color: '000000', strike: true }, style: style(18), widthPx: w(text), transform: { a: 1, b: 0, c: 0, d: 1, e, f: 300 } }) as Segment
    const { runs } = render([seg('some', 40), seg(' ', 40 + w('some')), seg('more', 40 + w('some') + w(' '))])
    expect(runs.map((r) => [r.text, r.servable])).toEqual([['some', true], ['more', true]])
  })

  // Observer level: 'Typing' at (20, 50), 20 px.
  const probe = (paint: (c: CanvasRenderingContext2D, ink: { x: number; y: number; width: number; height: number }) => void) => {
    const ctx = createCanvas(300, 100).getContext('2d') as unknown as CanvasRenderingContext2D
    const o = occlusionObserver(300, 100)
    const c = o.wrap(ctx)
    phase(c, 'text'); text(c, 'Typing', 20, 50)
    paint(c, o.runDetails()[0]!.ink)
    phase(c, false)
    return o.runDetails()[0]!
  }
  const bar = (c: CanvasRenderingContext2D, ink: { x: number; width: number }, y: number, t: number) => { c.fillStyle = '#000'; c.fillRect(ink.x - 2, y, ink.width + 4, t) }
  /** The left part of a line across the word (the run's own line), ending `cut` px into the ink… */
  const left = (c: CanvasRenderingContext2D, ink: { x: number; width: number }, y: number, t: number, cut = ink.width / 2) => { c.fillStyle = '#000'; c.fillRect(ink.x - 2, y, cut + 2, t) }
  /** …and the rest of it, starting `from` px into the ink (a neighbour's line carrying on beside the own one when from = cut). */
  const rest = (c: CanvasRenderingContext2D, ink: { x: number; width: number }, y: number, t: number, from = ink.width / 2) => { c.fillStyle = '#000'; c.fillRect(ink.x + from, y, ink.width - from + 2, t) }
  /** Another segment of the scope: its own text phase (a run elsewhere), then its line phase. */
  const neighbour = (c: CanvasRenderingContext2D, lines: () => void) => { phase(c, 'text'); text(c, 'next', 150, 90); phase(c, 'line'); lines() }
  it('a line of another segment at another position is ordinary cover', () => {
    expect(probe((c, i) => { phase(c, 'line'); bar(c, i, i.y + 3, 1); neighbour(c, () => bar(c, i, i.y + 9, 1)) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a line of another segment at the same position and thickness, carrying on beside the own line, is a continuation', () => {
    expect(probe((c, i) => { phase(c, 'line'); left(c, i, i.y + 3, 1); neighbour(c, () => rest(c, i, i.y + 3, 1)) })).toMatchObject({ servable: true })
    expect(probe((c, i) => { phase(c, 'line'); left(c, i, i.y + 3, 1); neighbour(c, () => rest(c, i, i.y + 3.005, 1)) })).toMatchObject({ servable: true })
  })
  it('a continuation must match both edges within 0.01 px (final fix: was 0.5)', () => {
    for (const d of [0.02, 0.4, 0.6]) {
      expect(probe((c, i) => { phase(c, 'line'); left(c, i, i.y + 3, 1); neighbour(c, () => rest(c, i, i.y + 3 + d, 1)) })).toMatchObject({ servable: false, reason: 'covered' })
      expect(probe((c, i) => { phase(c, 'line'); left(c, i, i.y + 3, 1); neighbour(c, () => rest(c, i, i.y + 3 - d, 1)) })).toMatchObject({ servable: false, reason: 'covered' })
      expect(probe((c, i) => { phase(c, 'line'); left(c, i, i.y + 3, 1); neighbour(c, () => rest(c, i, i.y + 3, 1 + d)) })).toMatchObject({ servable: false, reason: 'covered' })
    }
  })
  it('a collinear line of another segment lying on the own line is ordinary cover (final fix: a continuation is beside the own line)', () => {
    // Exactly the own line's position and thickness, drawn over it: whole, or reaching 1 px back over its end.
    expect(probe((c, i) => { phase(c, 'line'); bar(c, i, i.y + 3, 1); neighbour(c, () => bar(c, i, i.y + 3, 1)) })).toMatchObject({ servable: false, reason: 'covered' })
    expect(probe((c, i) => { phase(c, 'line'); left(c, i, i.y + 3, 1); neighbour(c, () => rest(c, i, i.y + 3, 1, i.width / 2 - 1)) })).toMatchObject({ servable: false, reason: 'covered' })
    expect(probe((c, i) => { phase(c, 'line'); left(c, i, i.y + 3, 1); neighbour(c, () => rest(c, i, i.y + 3, 1, i.width / 2 - 0.02)) })).toMatchObject({ servable: false, reason: 'covered' })
    // Float slack only: 0.005 px back is still beside it; a gap is too.
    expect(probe((c, i) => { phase(c, 'line'); left(c, i, i.y + 3, 1); neighbour(c, () => rest(c, i, i.y + 3, 1, i.width / 2 - 0.005)) })).toMatchObject({ servable: true })
    expect(probe((c, i) => { phase(c, 'line'); left(c, i, i.y + 3, 1); neighbour(c, () => rest(c, i, i.y + 3, 1, i.width / 2 + 1)) })).toMatchObject({ servable: true })
  })
  it('stroked lines: two abutting strokes overlap only by their cap allowances and continue; a stroke reaching back over the own line covers', () => {
    const stroke = (c: CanvasRenderingContext2D, x0: number, x1: number, y: number) => { c.strokeStyle = '#000'; c.lineWidth = 1; c.beginPath(); c.moveTo(x0, y); c.lineTo(x1, y); c.stroke() }
    const at = (back: number, clip = false, lineCap: CanvasLineCap = 'butt') => probe((c, i) => {
      const mid = i.x + i.width / 2, y = i.y + 3.5
      c.lineCap = lineCap
      phase(c, 'line'); stroke(c, i.x - 2, mid, y)
      neighbour(c, () => {
        if (clip) { c.save(); c.beginPath(); c.rect(mid - 0.3, 0, 300, 100); c.clip() }
        stroke(c, mid - back, i.x + i.width + 2, y)
        if (clip) c.restore()
      })
    })
    expect(at(0)).toMatchObject({ servable: true })
    expect(at(0.005)).toMatchObject({ servable: true })
    expect(at(0.02)).toMatchObject({ servable: false, reason: 'covered' })
    expect(at(1.2)).toMatchObject({ servable: false, reason: 'covered' })
    // A clip that cuts into the stroke's box (0.3 px before its true start, inside its 0.71 px cap allowance)
    // takes the allowance away: the whole clipped box counts, and that reaches back over the own line.
    expect(at(0, true)).toMatchObject({ servable: false, reason: 'covered' })
    // Square and round caps paint past the stroke's ends: no allowance is discounted, so abutting strokes overlap.
    for (const cap of ['square', 'round'] as const) expect(at(0, false, cap)).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a run with no own lines gets no exemption from the scope\'s other lines', () => {
    expect(probe((c, i) => { neighbour(c, () => bar(c, i, i.y + 3, 1)) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('a continuation of another scope is ordinary cover', () => {
    expect(probe((c, i) => { phase(c, 'line'); bar(c, i, i.y + 3, 1); phase(c, false); neighbour(c, () => bar(c, i, i.y + 3, 1)) })).toMatchObject({ servable: false, reason: 'covered' })
  })
  it('own lines and continuations share one union cap', () => {
    // Own lines at three heights and their continuations: still three bands (3 px ≤ 25 %); a fourth and fifth own band pass it.
    const at = [2, 6, 10]
    expect(probe((c, i) => { phase(c, 'line'); for (const y of at) left(c, i, i.y + y, 1); neighbour(c, () => { for (const y of at) rest(c, i, i.y + y, 1) }) })).toMatchObject({ servable: true })
    expect(probe((c, i) => { phase(c, 'line'); for (const y of [...at, 12, 14, 16]) left(c, i, i.y + y, 1); neighbour(c, () => { for (const y of [...at, 12, 14, 16]) rest(c, i, i.y + y, 1) }) })).toMatchObject({ servable: false, reason: 'covered' })
    expect(probe((c, i) => { phase(c, 'line'); for (const y of [...at, 12, 14, 16]) bar(c, i, i.y + y, 1) })).toMatchObject({ servable: false, reason: 'covered' })
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
