// @vitest-environment node
/**
 * The upper-segment window follow-up, phase 2 (owner ruling after capture ed0fccb6, stage2-scale-upper-window): Excel 16.115
 * drew every captured 2-colour or 3-colour LOWER-segment interior cell whose step product lies a rounding error below a
 * whole number one step smaller than production's guarded truncation (`floor(f·d + 1e-11)`): A3, A4, D2, D4, E7..E10,
 * E12..E14 and the twins H8, H18, I5. Production now withholds that shape (`cf-unmeasured-guard-boundary`): an interior cell
 * where, for some channel with d != 0, floor(p + CF_TRUNCATION_GUARD) != floor(p), p = f·|d| computed as
 * `interpolateChannel` computes it. The literal num-segment withhold and the upper-segment rules keep precedence. Raw floor
 * (which would paint Excel's colour) was never pre-registered then, so nothing was painted in its place.
 * The raw-floor follow-up (stage2-scale-raw-floor, ce3f701d, docs/validation/xlsx-cf.md §14.z.18) matched `raw-floor`, which
 * is now registered and published: with the id's evidence production paints the shape at raw floor (Excel's fill) where
 * neither bound is formula-resolved. The withhold is pinned here without that evidence (`unpublished`, `withoutEvidence`), and
 * the registered paint beside it.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import * as colorScale from '../src/xlsx/cf/colorScale'
import { guardBoundaryProduct, interpolateScaleColor, scaleColorAt } from '../src/xlsx/cf/colorScale'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cfDocumentEvaluationOf, prepareCfPaint } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { VERIFIED_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { CfScalePoint } from '../src/xlsx/cf/model'
import type { CfManifest } from '../scripts/cf-refs/cases'
import { collectSegmentCells } from '../scripts/cf-refs/num-segment-candidates'
import { stage5JustifiedRefusals, stage5SheetCells } from '../scripts/cf-refs/stage5-withhold'
import { stage2GuardBoundaryShape, stage2NumSegmentShape, stage2UpperSegmentWholeProductShape } from '../scripts/cf-refs/stage2-withhold'
import { caseParseOptionsOf } from '../scripts/cf-refs/case-clock'
import { independentGuardBoundary, JUSTIFY_TRUNCATION_GUARD } from '../scripts/cf-refs/upper-whole-justify'
import type { XlsxSheet } from '../src/xlsx/types'

const REPO_ROOT = resolve(__dirname, '..')
const HEAVY = 300_000
const P = 0xff14f028, Q = 0xffe61efa // P + 210·(1, −1, 1)
const num = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'num', value }, argb })
const min = (argb: number): CfScalePoint => ({ bound: { type: 'min' }, argb })
const max = (argb: number): CfScalePoint => ({ bound: { type: 'max' }, argb })
const pctl = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'percentile', value }, argb })
const all = { hasEvidence: () => true }
/** Every capability but the raw-floor registration's: the guard-boundary shape is withheld, as before ce3f701d. */
const unpublished = { hasEvidence: (id: string) => id !== colorScale.CF_GUARD_BOUNDARY_CAP }
const withoutEvidence = <T>(fn: () => T): T => withExactCfEvidence([...VERIFIED_CAPABILITIES].filter(id => id !== colorScale.CF_GUARD_BOUNDARY_CAP), fn)
/** As recorded at ce3f701d: neither the guard-boundary nor the later literal-lower registration (a3f85309) held. */
const asRecorded = <T>(fn: () => T): T => withExactCfEvidence([...VERIFIED_CAPABILITIES].filter(id => id !== colorScale.CF_GUARD_BOUNDARY_CAP && id !== colorScale.CF_LITERAL_LOWER_CAP), fn)
/** Raw floor per ARGB channel, restated (the registered reading): c0 moved by floor(f·|d|) toward c1. */
const rawFloor = (c0: number, c1: number, f: number): number => [24, 16, 8, 0].reduce((acc, sh) => {
  const a = (c0 >>> sh) & 0xff, b = (c1 >>> sh) & 0xff, k = Math.floor(f * Math.abs(b - a))
  return acc + (b >= a ? a + k : a - k) * 2 ** sh
}, 0)
const GUARD = { kind: 'unresolved', reason: 'cf-unmeasured-guard-boundary' }
const LITERAL = { kind: 'unresolved', reason: 'cf-unmeasured-num-segment-fraction' }
const UPPER = { kind: 'unresolved', reason: 'cf-unmeasured-upper-segment-whole-product' }
const ready = (argb: number): { kind: 'ready'; value: number } => ({ kind: 'ready', value: argb >>> 0 })
const manifest = JSON.parse(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
const CASE = 'stage2-scale-upper-window'
const MOVED = ['A3', 'A4', 'D2', 'D4', 'E7', 'E8', 'E9', 'E10', 'E12', 'E13', 'E14', 'H8', 'H18', 'I5']

describe('the shape, exactly (production, src/xlsx/cf/colorScale.ts)', () => {
  it('E7, computed as production paints it: f = 0.33333333333333326, p = 69.99999999999999; guarded 70, raw 69, Excel 69', () => {
    const f = colorScale.safeFraction(0.6, 0.4, 1)!
    expect(f).toBe(0.33333333333333326)
    expect(f * 210).toBe(69.99999999999999)
    expect(interpolateScaleColor(Q, P, f) >>> 0).toBe(0xffa064b4) // production's guarded step 70; Excel drew FFA163B5
    expect(guardBoundaryProduct(Q, P, 0.6, 0.4, 1, f)).toBe(true)
  })

  it('only where the guard changes the step: just above an integer, an exact integer, 1e-9 below, k = 0, d = 0 and endpoints are not the shape', () => {
    expect(guardBoundaryProduct(P, Q, 57, 0, 210, 57 / 210)).toBe(true) // H8: 56.99999999999999
    expect(guardBoundaryProduct(P, Q, 115, 0, 210, 115 / 210)).toBe(false) // H19: 115.00000000000001
    expect(guardBoundaryProduct(P, Q, 70, 0, 210, 70 / 210)).toBe(false) // H12: 70 exactly
    expect(guardBoundaryProduct(P, Q, 57, 0, 210, (57 - 1e-9) / 210)).toBe(false) // 1e-9 below: outside the guard
    expect(guardBoundaryProduct(P, Q, 1e-12, 0, 210, 1e-14)).toBe(false) // k = 0
    expect(guardBoundaryProduct(P, P, 57, 0, 210, 57 / 210)).toBe(false) // d = 0 on every channel
    expect(guardBoundaryProduct(P, Q, 0, 0, 210, 0)).toBe(false) // endpoints
    expect(guardBoundaryProduct(P, Q, 210, 0, 210, 1)).toBe(false)
    expect(guardBoundaryProduct(P, Q, 57, 57, 210, 0)).toBe(false) // value not strictly inside
  })

  it('the alpha channel counts too (interpolateChannel truncates it with the same guard)', () => {
    const a0 = 0x0014f028, a1 = 0xd214f028 // alpha 0 -> 210, RGB constant
    expect(guardBoundaryProduct(a0, a1, 57, 0, 210, 57 / 210)).toBe(true)
  })

  it('2-colour scales: withheld in the shape without the evidence (normal and inverted thresholds), painted elsewhere; registered, the shape paints raw floor', () => {
    expect(scaleColorAt(0.6, [num(0.4, Q), num(1, P)], [0.4, 1], unpublished)).toEqual(GUARD) // E7
    expect(scaleColorAt(57, [num(0, P), num(210, Q)], [0, 210], unpublished)).toEqual(GUARD) // H8
    expect(scaleColorAt(108, [num(0, P), num(210, Q)], [0, 210], unpublished)).toEqual(GUARD) // H18
    expect(scaleColorAt(127, [num(0, P), num(210, Q)], [0, 210], unpublished)).toEqual(GUARD) // I5
    expect(scaleColorAt(1.21, [min(Q), max(P)], [1.1, 2.2], unpublished)).toEqual(GUARD) // E12
    expect(scaleColorAt(57, [num(210, Q), num(0, P)], [210, 0], unpublished)).toEqual(GUARD) // inverted: same fraction, same product
    // Registered with its evidence (raw-floor, ce3f701d): Excel's fills of ed0fccb6.
    expect(scaleColorAt(0.6, [num(0.4, Q), num(1, P)], [0.4, 1], all)).toEqual(ready(0xffa163b5)) // E7
    expect(scaleColorAt(57, [num(0, P), num(210, Q)], [0, 210], all)).toEqual(ready(0xff4cb860)) // H8
    expect(scaleColorAt(108, [num(0, P), num(210, Q)], [0, 210], all)).toEqual(ready(0xff7f8593)) // H18
    expect(scaleColorAt(127, [num(0, P), num(210, Q)], [0, 210], all)).toEqual(ready(0xff9272a6)) // I5
    expect(scaleColorAt(1.21, [min(Q), max(P)], [1.1, 2.2], all)).toEqual(ready(rawFloor(Q, P, colorScale.safeFraction(1.21, 1.1, 2.2)!))) // E12
    expect(scaleColorAt(57, [num(210, Q), num(0, P)], [210, 0], all)).toEqual(ready(0xffae56c2)) // inverted: Q moved floor(56.99999999999999) = 56 toward P
    expect(rawFloor(Q, P, colorScale.safeFraction(57, 0, 210)!) >>> 0).toBe(0xffae56c2)
    // A formula-resolved bound in the shape stays withheld (never captured in this shape).
    expect(scaleColorAt(57, [{ bound: { type: 'num', value: 0, resolved: 'formula' }, argb: P } as CfScalePoint, num(210, Q)], [0, 210], all)).toEqual(GUARD)
    expect(scaleColorAt(115, [num(0, P), num(210, Q)], [0, 210], all)).toEqual(ready(0xff877d9b)) // H19, above an integer
    expect(scaleColorAt(70, [num(0, P), num(210, Q)], [0, 210], all)).toEqual(ready(0xff5aaa6e)) // H12, exact
    expect(scaleColorAt(0, [num(0, P), num(210, Q)], [0, 210], all)).toEqual(ready(P))
    expect(scaleColorAt(57, [num(0, P), num(210, Q)], [0, 210], { ...all, rounding: 'half-up' })).toEqual(ready(0xff4db761)) // no guard: not withheld
  })

  it('3-colour lower segment: withheld without the evidence (after the midpoint-above-max re-sort), raw floor registered; the literal num-segment and upper-segment withholds keep precedence', () => {
    const stat = [min(P), pctl(50, Q), max(P)]
    expect(scaleColorAt(0.13, stat, [0.1, 0.4, 1], unpublished)).toEqual(GUARD) // A3
    expect(scaleColorAt(0.21, stat, [0.1, 0.4, 1], unpublished)).toEqual(GUARD) // A4
    expect(scaleColorAt(0.13, stat, [0.1, 0.4, 1], all)).toEqual(ready(0xff28dc3c)) // A3 registered: Excel's fill
    expect(scaleColorAt(0.21, stat, [0.1, 0.4, 1], all)).toEqual(ready(0xff60a474)) // A4 registered: Excel's fill
    expect(scaleColorAt(0.27, stat, [0.1, 0.4, 1], all)).toEqual(ready(0xff8b799f)) // A5: 119 exactly
    expect(scaleColorAt(0.2, [min(P), { bound: { type: 'percent', value: 50 }, argb: Q }, max(P)], [0, 0.7, 1.4], all)).toEqual(ready(0xff50b464)) // C12, above
    // literal: the num-segment withhold first, unless the literal-lower reading is held (a3f85309: lower-arithmetic hands the
    // cell to this shape, so it paints Excel's A3 fill with both registrations, and is withheld without either).
    expect(scaleColorAt(0.13, [num(0.1, P), num(0.4, Q), num(1, P)], [0.1, 0.4, 1], { hasEvidence: (id: string) => id !== colorScale.CF_LITERAL_LOWER_CAP })).toEqual(LITERAL)
    expect(scaleColorAt(0.13, [num(0.1, P), num(0.4, Q), num(1, P)], [0.1, 0.4, 1], unpublished)).toEqual(GUARD)
    expect(scaleColorAt(0.13, [num(0.1, P), num(0.4, Q), num(1, P)], [0.1, 0.4, 1], all)).toEqual(ready(0xff28dc3c))
    expect(scaleColorAt(0.6, stat, [0.1, 0.4, 1], all)).toEqual(UPPER) // the upper segment: its own rules
    // Midpoint above max: thresholds re-sort to 0.1, 0.4, 1 and the lower segment runs 0.1 .. 0.4 with colours 1 -> 2.
    expect(scaleColorAt(0.13, [min(P), pctl(90, Q), pctl(50, P)], [0.1, 1, 0.4], unpublished)).toEqual(GUARD)
    expect(scaleColorAt(0.13, [min(P), pctl(90, Q), pctl(50, P)], [0.1, 1, 0.4], all)).toEqual(ready(0xff28dc3c))
  })
})

async function parsed(c: CfManifest['cases'][number]): Promise<{ sheet: XlsxSheet; doc: Awaited<ReturnType<typeof parseXlsx>> }> {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved)))), c.stage === 5 ? caseParseOptionsOf(c) : {})
  return { doc, sheet: doc.sheets.find(x => x.name === c.sheet) ?? doc.sheets[0]! }
}
const fillAt = (paint: Awaited<ReturnType<typeof prepareCfPaint>>, ref: string): string | undefined => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!
  return paint!.layer.styleAt(Number(m[2]) - 1, [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1, undefined)?.fillColor?.toUpperCase()
}

describe('the capture (stage2-scale-upper-window, ed0fccb6)', () => {
  const c = manifest.cases.find(x => x.id === CASE)!
  const sidecar = JSON.parse(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/refs/stage2-scale-upper-window.json'), 'utf8')) as { observedCells: Array<{ ref: string; fillRgb: string }> }
  const excel = new Map(sidecar.observedCells.map(o => [o.ref, o.fillRgb]))

  it('production without the evidence withholds the 14 moved cells and paints every other painted cell at Excel\'s fill', async () => {
    const { doc, sheet } = await parsed(c)
    const paint = await withoutEvidence(() => prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, evaluation: cfDocumentEvaluationOf(doc) }))
    const withheld: string[] = []
    for (const [ref, want] of excel) {
      const got = fillAt(paint, ref)
      if (want === 'FFFFFFFF') { expect(got, ref).toBeUndefined(); continue }
      if (got === undefined) withheld.push(ref)
      else expect(got, ref).toBe(want)
    }
    paint!.dispose()
    for (const ref of MOVED) expect(withheld, ref).toContain(ref)
    // 56 before the ruling (27 statistical window + 2 integer as upper whole product, 27 literal as num-segment) + 14.
    expect(withheld.length).toBe(70)
  }, HEAVY)

  it('production as committed (raw-floor registered and published) paints the 14 moved cells at Excel\'s fill; the other 56 stay withheld', async () => {
    const { doc, sheet } = await parsed(c)
    const paint = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, evaluation: cfDocumentEvaluationOf(doc) })
    const withheld: string[] = []
    for (const [ref, want] of excel) {
      const got = fillAt(paint, ref)
      if (want === 'FFFFFFFF') { expect(got, ref).toBeUndefined(); continue }
      if (got === undefined) withheld.push(ref)
      else expect(got, ref).toBe(want)
    }
    paint!.dispose()
    for (const ref of MOVED) expect(withheld, ref).not.toContain(ref)
    expect(withheld.length).toBe(56)
  }, HEAVY)

  it('the independent justification names exactly the 14 without the evidence, and no cell another withhold already justifies; registered, none', async () => {
    const { sheet } = await parsed(c)
    const cells = stage5SheetCells(sheet)
    const numbers = new Map([...cells].filter(([, v]) => typeof v.value === 'number' && !v.error).map(([k, v]) => [k, v.value as number]))
    const justified = withoutEvidence(() => [...excel.keys()].filter(ref => stage2GuardBoundaryShape(c.markers, ref, cells.get(ref)?.value, numbers)))
    expect(justified.sort()).toEqual([...MOVED].sort())
    expect([...excel.keys()].filter(ref => stage2GuardBoundaryShape(c.markers, ref, cells.get(ref)?.value, numbers))).toEqual([])
    for (const ref of justified) {
      expect(stage2NumSegmentShape(c.markers, ref, cells.get(ref)?.value), ref).toBe(false)
      expect(stage2UpperSegmentWholeProductShape(c.markers, ref, cells.get(ref)?.value, numbers), ref).toBe(false)
    }
  }, HEAVY)

  it('negative mutations refuse: a changed value, a moved threshold, an above-integer, an exact-integer and a 1e-9-below product are never justified', async () => {
    const { sheet } = await parsed(c)
    const cells = stage5SheetCells(sheet)
    const numbers = new Map([...cells].filter(([, v]) => typeof v.value === 'number' && !v.error).map(([k, v]) => [k, v.value as number]))
    const shape = (markers: typeof c.markers, ref: string, value: number, vals = numbers): boolean => withoutEvidence(() => stage2GuardBoundaryShape(markers, ref, value, vals))
    expect(shape(c.markers, 'E7', 0.6)).toBe(true)
    expect(shape(c.markers, 'E7', 0.61)).toBe(false) // a changed value
    const moved = c.markers.map(m => (m.sqref === 'E7:E10' ? { ...m, scale: { ...m.scale!, cfvo: m.scale!.cfvo.map((x, i) => (i === 0 ? { ...x, val: '0.41' } : x)) } } : m))
    expect(shape(moved, 'E7', 0.6)).toBe(false) // a moved threshold
    expect(shape(c.markers, 'H19', 115)).toBe(false) // above an integer
    expect(shape(c.markers, 'H12', 70)).toBe(false) // an exact integer
    expect(shape(c.markers, 'C12', 0.2)).toBe(false) // a lower-segment cell above an integer
    expect(shape(c.markers, 'A11', 0.6)).toBe(false) // an upper-segment cell: not this rule's
    expect(independentGuardBoundary('FF14F028', 'FFE61EFA', (57 - 1e-9) / 210)).toBe(false) // 1e-9 below
    expect(independentGuardBoundary('FF14F028', 'FFE61EFA', 57 / 210)).toBe(true)
    expect(JUSTIFY_TRUNCATION_GUARD).toBe(colorScale.CF_TRUNCATION_GUARD)
  }, HEAVY)
})

const RAW_FLOOR_CASE = 'stage2-scale-raw-floor'

describe('the reach: 14 cells in stage2-scale-upper-window and the 67 evidence cells of stage2-scale-raw-floor, none anywhere else', () => {
  it('through production\'s own segment evaluations, and the Stage 2 and Stage 5 justifications agree case by case without the evidence (as recorded at ce3f701d, before the literal-lower registration); registered, nothing is justified', async () => {
    const { cells } = await collectSegmentCells(manifest, REPO_ROOT, { include: 'all' })
    const literalWithheld = (s: (typeof cells)[number]['segment']): boolean => s.colours === 3 && s.loKind === 'num' && s.hiKind === 'num' && !Number.isInteger(s.fraction * 1048576)
    const inShape = cells.filter(x => (x.segment.colours === 2 || x.segment.segment === 0) && !literalWithheld(x.segment) && guardBoundaryProduct(x.segment.c0, x.segment.c1, x.segment.value, x.segment.lo, x.segment.hi, x.segment.fraction))
    const byCase: Record<string, string[]> = {}
    for (const x of inShape) (byCase[x.caseId] ??= []).push(x.ref)
    const { s2RawFloorCells } = await import('../scripts/cf-refs/cases')
    const evidence = s2RawFloorCells().filter(x => x.role === 'evidence' && x.configuration !== 'literal').map(x => x.ref)
    expect(evidence.length).toBe(67)
    expect(Object.fromEntries(Object.entries(byCase).map(([k, v]) => [k, [...v].sort()]))).toEqual({ [CASE]: [...MOVED].sort(), [RAW_FLOOR_CASE]: [...evidence].sort() })
    for (const x of inShape) expect(hex(interpolateScaleColor(x.segment.c0, x.segment.c1, x.segment.fraction)), `${x.caseId}:${x.ref}`).not.toBe(x.excel)
    for (const cs of manifest.cases.filter(x => x.status === 'measured' && x.id !== 'stage2-scale-literal-lower' && (x.stage === 2 || x.stage === 5) && x.markers.some(m => m.ruleType === 'colorScale'))) {
      const { sheet } = await parsed(cs)
      const cellsOf = stage5SheetCells(sheet)
      const numbers = new Map([...cellsOf].filter(([, v]) => typeof v.value === 'number' && !v.error).map(([k, v]) => [k, v.value as number]))
      const justifiedNow = (): string[] => cs.expected.cells.map(e => e.ref).filter(ref => cs.stage === 5
        ? stage5JustifiedRefusals(cs, ref, undefined, cellsOf).some(j => j.refused === 'cf-unmeasured-guard-boundary')
        : !stage2NumSegmentShape(cs.markers, ref, cellsOf.get(ref)?.value) && stage2GuardBoundaryShape(cs.markers, ref, cellsOf.get(ref)?.value, numbers))
      expect(asRecorded(justifiedNow).sort(), cs.id).toEqual([...(byCase[cs.id] ?? [])].sort())
      expect(justifiedNow(), `${cs.id} registered`).toEqual([])
    }
  }, HEAVY)
})

const hex = (argb: number): string => (argb >>> 0).toString(16).toUpperCase().padStart(8, '0')
