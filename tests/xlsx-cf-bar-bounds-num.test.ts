// @vitest-environment node
/**
 * Follow-up to Task 18b (orchestrator ruling 2026-11-19): stage3-bar-bounds-num, discriminating evidence for
 * bar:bounds-num.
 *
 * stage3-bar-lengths E holds 20..80 under num 20/80, so num and min/max predict the same px (0 px apart:
 * CF_STAGE3_NONDISCRIMINATING). This case keeps every value strictly inside the num bounds (25..75 under num 20/80): no
 * cell is beyond a bound (Ruling S3-E), and the two readings put the bar edges far apart. The case has no capture yet, so
 * its family is run end to end through the real harness on a synthetic capture: the generated workbook, Excel's measured
 * geometry convention (180 pt columns, 30 pt rows, 1 px per pt) and solid bars planted under one reading by an
 * independent painter. The planted reading must come back as the one match. This checks the family, not Excel.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { CF_STAGE3_CAPABILITIES, CF_STAGE3_CASES, CF_STAGE3_NONDISCRIMINATING, CF_STAGE3_OWNERS, type CfManifest } from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import {
  computeWorksheetRangeGeometry,
  evaluateBarCapabilities,
  measureCaseBars,
  readBarSheetFacts,
  capabilityOwnerCase,
  isCapabilityVerified,
  readZipPartsSync,
  verifyCfCase,
  type DecodedPng,
} from '../scripts/cf-refs/verify'

const REPO_ROOT = resolve(__dirname, '..')
const TIMEOUT_MS = 60_000
const ID = 'stage3-bar-bounds-num'
const CAP = 'bar:bounds-num'
const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
const c = manifest.cases.find(x => x.id === ID)!
const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved))))!
const facts = readBarSheetFacts(parts)
const rows = Number(/\d+$/.exec(c.range)![0])
const colWidthsPt = [84, 180]
const geom = computeWorksheetRangeGeometry(parts, c.range, { colWidthsPt, rowHeightsPt: Array.from({ length: rows }, () => 30) })!
const half = (x: number): number => Math.floor(x + 0.5 + 1e-9)
const BLUE = [0, 0, 255]
const S = 177
const scope = [2, 3, 4, 5, 6].map(r => (facts.values.get(`B${r}`) as { v: number }).v)

/** Solid 0/100 bars in column B, left to right, under the bounds `lo`/`hi` (the planted reading). */
function synthetic(lo: number, hi: number): DecodedPng {
  const width = colWidthsPt.reduce((a, b) => a + b, 0) + 1, height = rows * 30 + 1
  const data = new Uint8ClampedArray(width * height * 4).fill(255)
  const put = (x: number, y: number) => { const j = (y * width + x) * 4; data[j] = BLUE[0]!; data[j + 1] = BLUE[1]!; data[j + 2] = BLUE[2]! }
  for (const cell of c.expected.cells.filter(x => x.matched)) {
    const row = Number(cell.ref.slice(1)) - 1
    const v = (facts.values.get(cell.ref) as { v: number }).v
    const t = Math.max(0, Math.min(1, (v - lo) / (hi - lo)))
    const len = half(t * S)
    const x0 = 84 + 2, top = row * 30 + 2, bottom = row * 30 + 27
    for (let y = top; y <= bottom; y++) for (let i = 0; i < len; i++) put(x0 + i, y)
  }
  return { width, height, data }
}

const judge = (lo: number, hi: number) => {
  const png = synthetic(lo, hi)
  const measured = measureCaseBars(c, png, geom, parts)
  expect(measured instanceof Map, JSON.stringify(measured).slice(0, 300)).toBe(true)
  return evaluateBarCapabilities(c, parts, png, geom, measured as never)
}

describe('stage3-bar-bounds-num (synthetic capture)', () => {
  it('owns bar:bounds-num (appended last, so no earlier input changes), with one paired solid 0/100 bar under num 20/80', () => {
    expect(c.status).toBe('measured')
    expect(CF_STAGE3_OWNERS[CAP]).toBe(ID)
    expect(CF_STAGE3_CAPABILITIES.find(e => e.capabilityId === CAP)!.caseId).toBe(ID)
    expect(CF_STAGE3_CASES.at(-1)!.id).toBe(ID)
    // stage3-bar-lengths keeps its column E (now a cross-check) and its capture.
    expect(CF_STAGE3_CASES.find(x => x.id === 'stage3-bar-lengths')!.markers.some(m => m.capabilityId === CAP)).toBe(true)
    expect(c.markers.map(m => m.capabilityId)).toEqual([CAP, CAP])
    for (const m of c.markers) {
      expect(m.sqref).toBe('B2:B6')
      expect(m.bar?.cfvo).toEqual(m.extension === 'x14' ? [{ type: 'num', f: '20' }, { type: 'num', f: '80' }] : [{ type: 'num', val: '20' }, { type: 'num', val: '80' }])
    }
    expect(c.markers.find(m => m.extension === 'x14')!.bar!.attrs).toMatchObject({ minLength: '0', maxLength: '100', gradient: '0' })
    expect(c.expected.cells.filter(x => x.matched).map(x => x.ref)).toEqual(['B2', 'B3', 'B4', 'B5', 'B6'])
    // Captured (6a83cbe2): the capture verifies and discriminates, so the registry entry is retired.
    expect(CF_STAGE3_NONDISCRIMINATING[CAP]).toBeUndefined()
  })

  it('isCapabilityVerified asks the owner, not the first case carrying the marker (stage3-bar-lengths, listed first)', () => {
    expect(capabilityOwnerCase(CAP, manifest)?.id).toBe(ID)
    // Everything else keeps its owner: the registry's case for every Stage 3 capability.
    for (const e of CF_STAGE3_CAPABILITIES) expect(capabilityOwnerCase(e.capabilityId, manifest)?.id, e.capabilityId).toBe(e.caseId)
  })

  it('every value lies strictly inside the num bounds (no S3-E beyond-bounds cell), and min/max differ from them', () => {
    expect(scope).toEqual([25, 36, 45, 56, 75])
    for (const v of scope) expect(v > 20 && v < 80, String(v)).toBe(true)
    expect([Math.min(...scope), Math.max(...scope)]).toEqual([25, 75])
    // No predicted edge sits on a half pixel under either reading, so neither leans on snapping.
    for (const [lo, hi] of [[20, 80], [25, 75]] as const) {
      for (const v of scope) {
        const f = ((v - lo) / (hi - lo)) * S
        expect(Math.abs((f % 1) - 0.5) > 0.05, `${v} under ${lo}/${hi}: ${f}`).toBe(true)
      }
    }
  })

  it('planted num: bar:bounds-num comes back as num, the one match', () => {
    const ev = judge(20, 80)
    expect([ev[CAP]!.status, ev[CAP]!.matched], ev[CAP]!.reason).toEqual(['discriminating', 'num'])
  }, TIMEOUT_MS)

  it('planted min/max: bar:bounds-num comes back as min-max, the one match', () => {
    const ev = judge(25, 75)
    expect([ev[CAP]!.status, ev[CAP]!.matched], ev[CAP]!.reason).toEqual(['discriminating', 'min-max'])
  }, TIMEOUT_MS)

  it('bounds outside both readings match none (refused, never fitted)', () => {
    expect(judge(0, 100)[CAP]!.status).toBe('none')
  }, TIMEOUT_MS)

  it('the two readings are far apart: a bar against no bar at 25, 15 px at 75, and at least 3 px at every other cell', () => {
    const ev = judge(20, 80)
    const hyps = ev[CAP]!.hypotheses!
    const sep = (ref: string): number => {
      const [a, b] = hyps[ref]!
      return a!.noBar || b!.noBar ? (a!.noBar && b!.noBar ? 0 : Infinity) : Math.max(Math.abs(a!.leftPx - b!.leftPx), Math.abs(a!.rightPx - b!.rightPx))
    }
    // num: 15/47/74/106/162 px of the 177 px span; min/max: none/39/71/110/177 px.
    expect(['B2', 'B3', 'B4', 'B5', 'B6'].map(sep)).toEqual([Infinity, 8, 3, 4, 15])
    // At least the 3 px the +-1 px tolerance needs (discriminateBar's minSeparationPx), at every designated cell.
    for (const ref of ['B2', 'B3', 'B4', 'B5', 'B6']) expect(sep(ref)).toBeGreaterThanOrEqual(3)
  }, TIMEOUT_MS)

  // A full case verification: CI is about 10x slower than local.
  it('the committed capture (Excel 16.114) verifies bar:bounds-num as num, discriminating', () => {
    const res = verifyCfCase(c, { rootDir: REPO_ROOT })
    expect(res.kind, JSON.stringify(res).slice(0, 300)).toBe('verified')
    if (res.kind !== 'verified') return
    expect(res.verifiedCapabilities).toEqual([CAP])
    expect(res.nonDiscriminatingCapabilities ?? []).toEqual([])
    expect(res.matchedHypotheses?.[CAP]).toBe('num')
    expect(isCapabilityVerified(CAP, manifest, { rootDir: REPO_ROOT })).toBe(true)
    // Excel's right edges, B2..B6: the num reading (min/max would be none/124/156/195/262).
    const side = JSON.parse(readFileSync(resolve(REPO_ROOT, c.captures[0]!.sidecarPath), 'utf8')) as { observedBars: Array<{ ref: string; bar?: { rightPx: number } }> }
    expect(['B2', 'B3', 'B4', 'B5', 'B6'].map(r => side.observedBars.find(o => o.ref === r)?.bar?.rightPx)).toEqual([100, 132, 159, 191, 247])
  }, 600_000)
})
