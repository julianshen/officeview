/**
 * Stage 4a Task 5 — the servable sheet snapshot (S4 §6.2): opened over the projected package, base coverage
 * (steps 4–5) under `visible-only`, saved values (step 7) and the injected session clock (R9, F2).
 *
 * Step 7, "builder verifies" (recorded on the base 5aa80083): the server's `DocumentHost.open` parses with
 * `detectAndParse(projected.pkg)`, i.e. `parseXlsx` with no options, so `valuePolicy 'engine'` (the parse
 * recomputes; `doc.valuePolicy` is undefined). Units mode keeps that parse. Tiles mode adds a second parse with
 * `{ valuePolicy: 'cached' }` (`OpenDoc.savedParse`, F1), and every snapshot is built from it.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { DocumentHost, RenderRefusal } from '../src/render'
import { baseCoverage, type BaseCoverage } from '../src/sheet/coverage'
import { clockEpochMs, type SheetClock } from '../src/sheet/clock'
import type { SheetSnapshot } from '../src/sheet/snapshot'
import { cfClockOptions } from '../../src/components/xlsx/sheetAssets'
import { cfDocumentEvaluationOf } from '../../src/xlsx/cf/formula'
import { prepareCfPaint } from '../../src/xlsx/cf/render'
import { buildSheetGeometry } from '../../src/xlsx/view/viewport'
import { installCanvasBackend } from '../../src/core/canvas-backend'
import type { XlsxCell, XlsxDocument, XlsxDrawing, XlsxSheet } from '../../src/xlsx/types'
import { augmentedFixtures } from './fixtures'
import { rejection } from './fixtures/rejection'
import {
  cfToday, cfTodayFormula, crossSheetCoveredRef, sparklineUnderShape, CF_TODAY_SERIALS, DATA_BAR_VALUES, rotatedBox, savedFormula, spillUnderShape, SPILL_TEXT,
  unfilledShapeOverDataBar, unfilledShapeOverText, unresolvedAnchor, labelBesideShape, LONG_LABEL, BOLD_REACH_LABEL,
} from './sheet-fixtures'

const hosts: DocumentHost[] = []
const fresh = () => { const h = new DocumentHost(); hosts.push(h); return h }
afterAll(() => { for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k) })

const UTC_22: SheetClock = { day: '2026-11-22', timeZone: 'UTC' }
const UTC_23: SheetClock = { day: '2026-11-23', timeZone: 'UTC' }
/** UTC+14: 12:00 on 2026-11-23 there is 22:00 UTC on 2026-11-22, so a zone that is not threaded reads the 22nd. */
const KIRITIMATI_23: SheetClock = { day: '2026-11-23', timeZone: 'Pacific/Kiritimati' }

type Policy = 'visible-only' | 'full'
async function snapshotOf(bytes: Uint8Array, policy: Policy = 'visible-only', clock: SheetClock = UTC_23, sheetIndex = 0): Promise<{ host: DocumentHost; docKey: string; snap: SheetSnapshot }> {
  const host = fresh()
  const o = await host.open('doc', bytes, policy, { sheetMode: 'tiles' })
  return { host, docKey: o.docKey, snap: await host.sheetSnapshot(o.docKey, sheetIndex, clock) }
}
const cellAt = (sheet: XlsxSheet, row: number, col: number): XlsxCell | undefined => sheet.rows.flatMap((r) => r.cells).find((c) => c.row === row && c.col === col)
const known = (c: BaseCoverage | 'off') => {
  if (c === 'off' || c.kind !== 'known') throw new Error(`coverage is not known: ${JSON.stringify(c)}`)
  return c
}
/** The cells (0-based) of rows 0..rows-1, cols 0..cols-1. */
const cellsIn = (rows: number, cols: number) => Array.from({ length: rows * cols }, (_, i) => [Math.floor(i / cols), i % cols] as const)

describe('the servable snapshot (S4 §6.2)', () => {
  test('the snapshot opens over the projected package', async () => {
    const fixtures = (await augmentedFixtures()).filter((f) => f.kind === 'xlsx')
    expect(fixtures.length).toBeGreaterThan(0)
    for (const f of fixtures) {
      const host = fresh()
      const o = await host.open(f.id, f.withHidden, 'visible-only', { sheetMode: 'tiles' })
      expect(o.sheets!.length).toBeGreaterThan(0)
      for (let i = 0; i < o.sheets!.length; i++) {
        const snap = await host.sheetSnapshot(o.docKey, i, UTC_23)
        expect(snap.sheetIndex).toBe(i)
        expect(snap.sheet.name).toBe(o.sheets![i]!)
        for (const row of snap.sheet.rows) for (const cell of row.cells) expect(String(cell.value)).not.toContain(f.marker)
      }
      // The snapshot's sheets are exactly the manifest's: one past the last is refused.
      const past = await rejection(host.sheetSnapshot(o.docKey, o.sheets!.length, UTC_23))
      expect(past).toBeInstanceOf(RenderRefusal)
      expect((past as RenderRefusal).code).toBe('bad-job')
    }
  }, 120_000)

  test('unfilled shape covers', async () => {
    const bytes = await unfilledShapeOverText()
    const { snap } = await snapshotOf(bytes)
    const c = known(snap.coverage)
    expect(c.covered(1, 1)).toBe(true)
    expect(c.count).toBe(1)
    for (const [r, col] of cellsIn(4, 4)) if (r !== 1 || col !== 1) expect(c.covered(r, col)).toBe(false)
    expect(snap.contentPolicy).toBe('visible-only')
    const b2 = cellAt(snap.sheet, 1, 1)!
    expect(b2.value).toBeNull()
    const authored = cellAt((await snapshotOf(bytes, 'full')).snap.sheet, 1, 1)!
    expect(b2.styleIndex).toBe(authored.styleIndex)
    expect(b2.style).toEqual(authored.style!)
    // The geometry the tiles paint from is the blanked one too.
    if (snap.geometry.kind !== 'ready') throw new Error('geometry not ready')
    expect(snap.geometry.cellsOf(1)!.find((x) => x.col === 1)!.value).toBeNull()
    expect(cellAt(snap.sheet, 1, 0)!.value).toBe('CELL-2-1')

    const full = (await snapshotOf(bytes, 'full')).snap
    expect(full.coverage).toBe('off')
    expect(full.contentPolicy).toBe('full')
    expect(cellAt(full.sheet, 1, 1)!.value).toBe('COVERED-B2')
  }, 60_000)

  test('covered cells draw no CF', async () => {
    const bytes = await unfilledShapeOverDataBar()
    const vis = (await snapshotOf(bytes)).snap
    expect(known(vis.coverage).count).toBe(1)
    const cf = vis.assets.conditionalFormatting
    expect(cf).toBeDefined()
    expect(cf!.barAt(1, 1)).toBeUndefined()
    // Not vacuous: the uncovered cells of the same rule still draw their bars.
    expect(cf!.barAt(0, 1)).toBeDefined()
    expect(cf!.barAt(2, 1)).toBeDefined()
    const full = (await snapshotOf(bytes, 'full')).snap
    expect(cellAt(full.sheet, 1, 1)!.value).toBe(DATA_BAR_VALUES[1]!)
    expect(full.assets.conditionalFormatting!.barAt(1, 1)).toBeDefined()
  }, 60_000)

  test('rotated box adversary', async () => {
    const { snap } = await snapshotOf(await rotatedBox())
    const c = known(snap.coverage)
    if (snap.geometry.kind !== 'ready') throw new Error('geometry not ready')
    const m = snap.geometry.metrics
    // Every cell whose grid rect meets the spec's x 40..60, y 30..50 is covered…
    for (const [r, col] of cellsIn(6, 3)) {
      const box = m.sourceCellToLogical({ row: r, col })!
      const meets = box.x < 60 && box.x + box.width > 40 && box.y < 50 && box.y + box.height > 30
      if (meets) expect(c.covered(r, col)).toBe(true)
    }
    // …and so is A1 (F19: the rotated box spans x 40..60, y -40..60), with A2 and A3.
    expect(c.covered(0, 0)).toBe(true)
    expect(c.covered(1, 0)).toBe(true)
    expect(c.covered(2, 0)).toBe(true)
    // B1 meets the unrotated anchor box (x 0..100, y 0..20) only: beyond the rotated extent, not covered.
    expect(c.covered(0, 1)).toBe(false)
    expect(c.covered(3, 0)).toBe(false)
    expect(c.count).toBe(3)
    expect(cellAt(snap.sheet, 0, 0)!.value).toBeNull()
    expect(cellAt(snap.sheet, 0, 1)!.value).toBe('ROT-1-2')
  }, 60_000)

  test('spill counts toward coverage', async () => {
    const left = (await snapshotOf(await spillUnderShape('left'))).snap
    expect(known(left.coverage).covered(2, 0)).toBe(true)
    expect(known(left.coverage).count).toBe(1)
    expect(cellAt(left.sheet, 2, 0)!.value).toBeNull()
    const right = (await snapshotOf(await spillUnderShape('right'))).snap
    expect(known(right.coverage).covered(2, 0)).toBe(false)
    expect(known(right.coverage).count).toBe(0)
    expect(cellAt(right.sheet, 2, 0)!.value).toBe(SPILL_TEXT)
    // The box itself sits in C3 either way.
    expect(known(right.coverage).covered(2, 2)).toBe(true)
  }, 60_000)

  test('measured spill: a short label beside a shape paints (fix round 1)', async () => {
    // "Q1" fits A2: the empty run B2:E2 reaches the box in D2, but the text never does.
    const { snap } = await snapshotOf(await labelBesideShape('Q1'))
    expect([known(snap.coverage).covered(1, 0), known(snap.coverage).count]).toEqual([false, 0])
    expect(cellAt(snap.sheet, 1, 0)!.value).toBe('Q1')
    expect(known(snap.coverage).covered(1, 3)).toBe(true)
  }, 60_000)

  test('measured spill: long text that reaches the shape is still blanked', async () => {
    const { snap } = await snapshotOf(await labelBesideShape(LONG_LABEL))
    expect([known(snap.coverage).covered(1, 0), known(snap.coverage).count]).toEqual([true, 1])
    expect(cellAt(snap.sheet, 1, 0)!.value).toBeNull()
  }, 60_000)

  test('measured spill: unmeasurable keeps the whole-run box', async () => {
    // Rotated text is never measured.
    const rotated = (await snapshotOf(await labelBesideShape('Q1', { rotated: true }))).snap
    expect(known(rotated.coverage).covered(1, 0)).toBe(true)
    expect(cellAt(rotated.sheet, 1, 0)!.value).toBeNull()
    // No canvas backend: nothing can be measured.
    const sheet = (await savedParseFor(await labelBesideShape('Q1'))).sheets[0]!
    const geometry = buildSheetGeometry(sheet)
    expect(known(baseCoverage(sheet, geometry)).covered(1, 0)).toBe(false)
    const previous = installCanvasBackend(undefined)
    try { expect(known(baseCoverage(sheet, geometry)).covered(1, 0)).toBe(true) } finally { installCanvasBackend(previous) }
  }, 60_000)

  test('measured spill on a CF sheet: fill-only or bold rules let a short label paint (fix round 2)', async () => {
    for (const cf of ['fill', 'bold'] as const) {
      const { snap } = await snapshotOf(await labelBesideShape('Q1', { cf }))
      expect([cf, known(snap.coverage).covered(1, 0), cellAt(snap.sheet, 1, 0)!.value]).toEqual([cf, false, 'Q1'])
      expect([cf, known(snap.coverage).covered(1, 3)]).toEqual([cf, true])
    }
  }, 60_000)

  test('measured spill on a CF sheet: a label that reaches the shape only when bold is blanked', async () => {
    // Control: without CF the regular text stays clear of the box.
    const plain = (await snapshotOf(await labelBesideShape(BOLD_REACH_LABEL))).snap
    expect([known(plain.coverage).covered(1, 0), cellAt(plain.sheet, 1, 0)!.value]).toEqual([false, BOLD_REACH_LABEL])
    // A CF rule could make it bold (worst case: bold and italic), so it is measured that way, even under a fill-only rule.
    for (const cf of ['fill', 'bold'] as const) {
      const { snap } = await snapshotOf(await labelBesideShape(BOLD_REACH_LABEL, { cf }))
      expect([cf, known(snap.coverage).covered(1, 0), cellAt(snap.sheet, 1, 0)!.value]).toEqual([cf, true, null])
    }
  }, 60_000)

  test('measured spill on a CF sheet: a dxf font size keeps the whole-run box', async () => {
    const { snap } = await snapshotOf(await labelBesideShape('Q1', { cf: 'size' }))
    expect([known(snap.coverage).covered(1, 0), cellAt(snap.sheet, 1, 0)!.value]).toEqual([true, null])
  }, 60_000)

  test('unknown coverage paints as authored', async () => {
    const host = fresh()
    const o = await host.open('unres', await unresolvedAnchor(), 'visible-only', { sheetMode: 'tiles' })
    const snap = await host.sheetSnapshot(o.docKey, 0, UTC_23)
    expect(snap.coverage).toEqual({ kind: 'unknown', reason: 'unresolved-anchors' })
    const projected = savedParseOf(host, o.docKey).sheets[0]!
    expect(snap.sheet).toBe(projected)

    // A node whose transform cannot be computed (S4 §6.2 step 4 pitfall, F16): never treated as absent.
    const rot = (await savedParseFor(await rotatedBox())).sheets[0]!
    const geometry = buildSheetGeometry(rot)
    const variants: Array<(d: XlsxDrawing) => XlsxDrawing> = [
      (d) => ({ ...d, transformValid: false }),
      (d) => ({ ...d, xEmu: Number.NaN }),
      (d) => ({ ...d, widthEmu: -9525 }),
      (d) => ({ ...d, rotationDeg: Number.POSITIVE_INFINITY }),
    ]
    for (const v of variants) expect(baseCoverage({ ...rot, drawings: rot.drawings!.map(v) }, geometry)).toEqual({ kind: 'unknown', reason: 'transform' })
    // A scaled or skewed shadow (ShadowDefinition.scaledOrSkewed, resolved as the `outerShdw-scale` issue) is unbounded.
    const skewed = rot.drawings!.map((d) => ({ ...d, drawingStyle: { issues: [{ kind: 'unsupported-effect' as const, message: 'scaled', feature: 'outerShdw-scale' }], shadow: { color: { r: 0, g: 0, b: 0, a: 1 }, blurPx: 2, offsetX: 1, offsetY: 1 } } }))
    expect(baseCoverage({ ...rot, drawings: skewed }, geometry)).toEqual({ kind: 'unknown', reason: 'unbounded-effect' })
    // A child inheriting a group's scaled shadow: the issue sits on the group's style.
    const group: XlsxDrawing = { ...rot.drawings![0]!, rotationDeg: 0, group: { off: { x: 0, y: 0 }, ext: { width: 952500, height: 190500 }, chOff: { x: 0, y: 0 }, chExt: { width: 952500, height: 190500 } }, children: [{ ...rot.drawings![0]!, rotationDeg: 0, drawingStyle: { issues: [], shadow: { color: { r: 0, g: 0, b: 0, a: 1 }, blurPx: 2, offsetX: 1, offsetY: 1 } } }], drawingStyle: { issues: [{ kind: 'unsupported-effect', message: 'scaled', feature: 'outerShdw-scale' }] } } as XlsxDrawing
    expect(baseCoverage({ ...rot, drawings: [group] }, geometry)).toEqual({ kind: 'unknown', reason: 'unbounded-effect' })
    // A geometry that could not be built cannot place a cell.
    expect(baseCoverage(rot, { kind: 'diagnostic', diagnostics: [] })).toEqual({ kind: 'unknown', reason: 'transform' })
  }, 60_000)

  test('effect padding: the painted shadow covers, unpainted effects do not', async () => {
    const rot = (await savedParseFor(await rotatedBox())).sheets[0]!
    const geometry = buildSheetGeometry(rot)
    // An unrotated 40×10 box at (10, 5), inside A1 (x 0..64, y 0..20).
    const box = (style?: XlsxDrawing['drawingStyle']): XlsxDrawing => ({ ...rot.drawings![0]!, xEmu: 10 * 9525, yEmu: 5 * 9525, widthEmu: 40 * 9525, heightEmu: 10 * 9525, rotationDeg: 0, ...(style ? { drawingStyle: style } : {}) })
    const plain = known(baseCoverage({ ...rot, drawings: [box()] }, geometry))
    expect([plain.covered(0, 0), plain.covered(1, 0), plain.covered(0, 1)]).toEqual([true, false, false])
    // A shadow offset 20 px down reaches A2; a 30 px blur alone reaches B1.
    const down = known(baseCoverage({ ...rot, drawings: [box({ issues: [], shadow: { color: { r: 0, g: 0, b: 0, a: 1 }, blurPx: 0, offsetX: 0, offsetY: 20 } })] }, geometry))
    expect(down.covered(1, 0)).toBe(true)
    const blur = known(baseCoverage({ ...rot, drawings: [box({ issues: [], shadow: { color: { r: 0, g: 0, b: 0, a: 1 }, blurPx: 30, offsetX: 0, offsetY: 0 } })] }, geometry))
    expect(blur.covered(0, 1)).toBe(true)
    // A 30 px line pads by at least its half width (15 px): it reaches B1 and A2.
    const thick = known(baseCoverage({ ...rot, drawings: [box({ issues: [], line: { width: 30, fill: { kind: 'solid', color: { r: 0, g: 0, b: 0, a: 1 } } } })] }, geometry))
    expect([thick.covered(0, 1), thick.covered(1, 0)]).toEqual([true, true])
    // Glow, reflection and the like are never painted: no extent.
    const glow = known(baseCoverage({ ...rot, drawings: [box({ issues: [{ kind: 'unsupported-effect', message: 'deferred', feature: 'glow' }] })] }, geometry))
    expect([glow.covered(0, 0), glow.covered(1, 0), glow.covered(0, 1)]).toEqual([true, false, false])
  }, 60_000)

  test('saved values, no recompute', async () => {
    const bytes = await savedFormula()
    const host = fresh()
    const o = await host.open('saved', bytes, 'visible-only', { sheetMode: 'tiles' })
    const snap = await host.sheetSnapshot(o.docKey, 0, UTC_23)
    expect(cellAt(snap.sheet, 0, 0)!.value).toBe(7)
    expect(savedParseOf(host, o.docKey).valuePolicy).toBe('cached')
    // The fixture discriminates (fullCalcOnLoad): the engine parse units mode keeps recomputes A1 to 2.
    const engine = (host as unknown as { docs: Map<string, { parsed: XlsxDocument }> }).docs.get(o.docKey)!.parsed
    expect(engine.valuePolicy).toBeUndefined()
    expect(cellAt(engine.sheets[0]!, 0, 0)!.value).toBe(2)
    // A tiles open joining an entry a units open created gets its saved-values parse too (F1).
    const joined = fresh()
    const u = await joined.open('saved', bytes, 'visible-only')
    expect(savedParseOf(joined, u.docKey, false)).toBeUndefined()
    const t = await joined.open('saved', bytes, 'visible-only', { sheetMode: 'tiles' })
    expect(t.docKey).toBe(u.docKey)
    expect(cellAt((await joined.sheetSnapshot(t.docKey, 0, UTC_23)).sheet, 0, 0)!.value).toBe(7)
    expect(joined.parses).toBe(1)
  }, 60_000)

  test('clock is injected', async () => {
    const bytes = await cfToday()
    const fills = (snap: SheetSnapshot) => CF_TODAY_SERIALS.map((_, r) => snap.assets.conditionalFormatting!.layer.styleAt(r, 0, cellAt(snap.sheet, r, 0)!.style)?.fillColor ?? '')
    const host = fresh()
    const o = await host.open('today', bytes, 'visible-only', { sheetMode: 'tiles' })
    const doc = savedParseOf(host, o.docKey)
    for (const [clock, filled] of [[UTC_22, 3], [UTC_23, 4], [KIRITIMATI_23, 4]] as const) {
      const options = cfClockOptions(cfDocumentEvaluationOf(doc), { epochNowMs: clockEpochMs(clock), timeZone: clock.timeZone })
      expect(options.dateContext.epochNowMs).toBe(clockEpochMs(clock))
      expect(options.now()).toBe(clockEpochMs(clock))
      expect(options.evaluation.timeZone).toBe(clock.timeZone)
      expect(options.dateContext.timeZone).toBe(clock.timeZone)
      const snap = await host.sheetSnapshot(o.docKey, 0, clock)
      expect(snap.clock).toEqual(clock)
      // Exactly the clock's day fills (A4 is 2026-11-22, A5 2026-11-23), whatever the wall clock reads.
      expect(fills(snap).map((f, r) => (f ? r : -1)).filter((r) => r >= 0)).toEqual([filled])
    }
    // One promise per (docKey, sheetIndex, day, timeZone).
    expect(host.sheetSnapshot(o.docKey, 0, UTC_23)).toBe(host.sheetSnapshot(o.docKey, 0, { ...UTC_23 }))
    expect(host.sheetSnapshot(o.docKey, 0, UTC_23)).not.toBe(host.sheetSnapshot(o.docKey, 0, KIRITIMATI_23))
  }, 60_000)

  test('a TODAY() formula rule still evaluates under the injected clock (F2)', async () => {
    const bytes = await cfTodayFormula()
    const host = fresh()
    const o = await host.open('todayf', bytes, 'visible-only', { sheetMode: 'tiles' })
    for (const [clock, filled] of [[UTC_22, 3], [KIRITIMATI_23, 4]] as const) {
      const snap = await host.sheetSnapshot(o.docKey, 0, clock)
      const cf = snap.assets.conditionalFormatting!
      const fills = CF_TODAY_SERIALS.map((_, r) => cf.layer.styleAt(r, 0, cellAt(snap.sheet, r, 0)!.style)?.fillColor ?? '')
      expect(fills.map((f, r) => (f ? r : -1)).filter((r) => r >= 0)).toEqual([filled])
      // The stamp carries the clock's own local day number (unix days), not the wall clock's.
      expect(cf.stamp.clockGen).toBe(Date.UTC(...(clock.day.split('-').map(Number) as [number, number, number]).map((v, i) => (i === 1 ? v - 1 : v)) as [number, number, number]) / 86_400_000)
      expect((snap.sheet.diagnostics ?? []).map((d) => d.feature)).not.toContain('cf-evaluation-context-missing')
    }
    // The control: a date context that disagrees with the evaluation drops the formulas (cfAlignClock), so no cell fills.
    const doc = savedParseOf(host, o.docKey), sheet = doc.sheets[0]!
    const epochNowMs = clockEpochMs(UTC_22)
    const dropped = await prepareCfPaint(sheet, { valuePolicy: 'cached', evaluation: { ...cfDocumentEvaluationOf(doc), timeZone: 'UTC' }, now: () => epochNowMs, dateContext: { dateSystem: '1900', timeZone: 'UTC', epochNowMs: epochNowMs + 1 }, documentSheets: doc.sheets })
    expect(CF_TODAY_SERIALS.some((_, r) => dropped!.layer.styleAt(r, 0, cellAt(sheet, r, 0)!.style)?.fillColor)).toBe(false)
    dropped!.dispose()
  }, 60_000)

  test('cross-sheet CF reads another sheet\'s covered cell as blank (fix round 1)', async () => {
    const bytes = await crossSheetCoveredRef()
    const fillOf = (snap: SheetSnapshot) => snap.assets.conditionalFormatting!.layer.styleAt(0, 0, cellAt(snap.sheet, 0, 0)!.style)?.fillColor
    // Under 'full' the rule reads S!B2 = 20 and T!A1 fills: the fixture's rule does evaluate across sheets.
    const full = await snapshotOf(bytes, 'full', UTC_23, 1)
    expect(fillOf(full.snap)).toBe('FFFF0000')
    // Under visible-only S!B2 is covered, so T's rule reads it blank, whichever sheet was snapshotted first.
    const vis = await snapshotOf(bytes, 'visible-only', UTC_23, 1)
    expect(fillOf(vis.snap)).toBeUndefined()
    const s0 = await vis.host.sheetSnapshot(vis.docKey, 0, UTC_23)
    expect(known(s0.coverage).covered(1, 1)).toBe(true)
    expect(cellAt(s0.sheet, 1, 1)!.value).toBeNull()
  }, 60_000)

  test('a sparkline located in a covered cell is absent (fix round 1)', async () => {
    const bytes = await sparklineUnderShape()
    const vis = (await snapshotOf(bytes)).snap
    expect(known(vis.coverage).covered(1, 1)).toBe(true)
    expect(vis.sheet.resolvedSparklines?.get(1)?.has(1)).toBeFalsy()
    // The uncovered control at C2 still draws.
    expect(vis.sheet.resolvedSparklines?.get(1)?.has(2)).toBe(true)
    const full = (await snapshotOf(bytes, 'full')).snap
    expect(full.sheet.resolvedSparklines?.get(1)?.has(1)).toBe(true)
  }, 60_000)

  test('drop releases the cached snapshots', async () => {
    const host = fresh()
    const o = await host.open('drop', await unfilledShapeOverText(), 'visible-only', { sheetMode: 'tiles' })
    const snap = await host.sheetSnapshot(o.docKey, 0, UTC_23)
    let released = 0
    const release = snap.release.bind(snap)
    snap.release = () => { released++; release() }
    host.drop(o.docKey)
    await Promise.resolve()
    await Promise.resolve()
    expect(released).toBe(1)
    const gone = await rejection(host.sheetSnapshot(o.docKey, 0, UTC_23))
    expect((gone as RenderRefusal).code).toBe('unknown-doc')
  }, 60_000)
})

/** The saved-values parse of a fresh tiles-mode open of `bytes`. */
async function savedParseFor(bytes: Uint8Array): Promise<XlsxDocument> {
  const host = fresh()
  return savedParseOf(host, (await host.open('p', bytes, 'visible-only', { sheetMode: 'tiles' })).docKey)
}
function savedParseOf(host: DocumentHost, docKey: string): XlsxDocument
function savedParseOf(host: DocumentHost, docKey: string, required: false): XlsxDocument | undefined
function savedParseOf(host: DocumentHost, docKey: string, required = true): XlsxDocument | undefined {
  const doc = (host as unknown as { docs: Map<string, { savedParse?: XlsxDocument }> }).docs.get(docKey)
  if (!doc) throw new Error(`no entry ${docKey}`)
  if (required && !doc.savedParse) throw new Error('no savedParse')
  return doc.savedParse
}
