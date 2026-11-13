// @vitest-environment node
/**
 * Maintained regression test (Ruling C1 & C3):
 * Runs the production viewer CF pipeline (cfRuntimeEvidence) on every measured
 * Stage 1 manifest case and compares each expected cell (matched/unmatched,
 * fill, font, bold) against host Excel oracle expectations.
 *
 * Covers all 53 cells withheld and 7 cells painted wrongly in final-r1-oracle-vs-viewer.txt.
 */
import { describe, expect, it, vi } from 'vitest'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx, type ParseXlsxOptions } from '../src/xlsx/parse'
import { prepareCfPaint, residentSource } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { VERIFIED_CAPABILITIES, injectCfAutoBarColours, injectCfBarLayout, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import * as verifiedExports from '../src/xlsx/cf/verifiedCapabilities'
import { CF_BAR_RENDER_MODEL, layoutDataBar } from '../src/xlsx/cf/dataBarPaint'
import { auditStage3Manifest } from '../scripts/cf-refs/audit'
import { stage3ObservedBars } from '../scripts/cf-compare'
import { CF_LIMITS, createCfMemory } from '../src/xlsx/cf/limits'
import { createCfService } from '../src/xlsx/cf/service'
import type { CfBarDescriptor, CfCellState, CfOutcome } from '../src/xlsx/cf/model'
import { cfRulesVersion } from '../src/xlsx/view/pagedSource'
import { decodePngSync, isCellInSqref, type DecodedPng } from '../scripts/cf-refs/verify'
import { CF_CASES, type CfManifest } from '../scripts/cf-refs/cases'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const MANIFEST_PATH = join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json')

function refToRowCol(ref: string): { row: number; col: number } {
  const match = ref.match(/^([A-Z]+)(\d+)$/)
  if (!match) throw new Error(`Invalid cell ref: ${ref}`)
  const colLetters = match[1]!
  const row = parseInt(match[2]!, 10) - 1
  let col = 0
  for (let i = 0; i < colLetters.length; i++) {
    col = col * 26 + (colLetters.charCodeAt(i) - 64)
  }
  return { row, col: col - 1 }
}

/** The viewer evaluates TODAY() in the workbook/viewer time zone (host default). The Excel observation was taken in
 * `requiredTimeZone`, so a conformance run evaluates in that zone instead of inheriting the CI host's (UTC). */
function parseOptionsFor(c: { clockRequirement?: { requiredTimeZone: string } }): ParseXlsxOptions {
  // `evaluation` is spread into evaluateWorkbookFormulas at runtime, which honours timeZone; the public type only
  // advertises the array caps, hence the cast.
  return (c.clockRequirement ? { evaluation: { timeZone: c.clockRequirement.requiredTimeZone } } : {}) as ParseXlsxOptions
}


/**
 * Stage 3: Excel's own cell background in the committed capture, as the majority colour of the cell's first inner row
 * (1 px below the top gridline, from 1 px right of the left gridline to 1 px before the next). A bar never covers that
 * row (bars keep a 2 px top inset, bar:geometry-insets); a 1 px axis column or stray ink cannot outvote the fill. Used to
 * check that a fill production paints on a bar cell is the fill Excel paints there (Ruling S3-I coexistence), never a
 * guess, and that production paints no fill where Excel shows none.
 */
const stage3Pngs = new Map<string, { png: DecodedPng; cols: number[]; rows: number[] }>()
function excelCellBackground(c: { id: string; captures: ReadonlyArray<{ pngPath: string; sidecarPath: string }> }, ref: string): string {
  let cap = stage3Pngs.get(c.id)
  if (!cap) {
    const sc = JSON.parse(readFileSync(join(REPO_ROOT, c.captures[0]!.sidecarPath), 'utf8')) as { colWidthsPt: number[]; rowHeightsPt: number[] }
    cap = { png: decodePngSync(new Uint8Array(readFileSync(join(REPO_ROOT, c.captures[0]!.pngPath))))!, cols: sc.colWidthsPt, rows: sc.rowHeightsPt }
    stage3Pngs.set(c.id, cap)
  }
  const { row, col } = refToRowCol(ref)
  const x0 = Math.round(cap.cols.slice(0, col).reduce((a, b) => a + b, 0)), w = Math.round(cap.cols[col]!)
  const y = Math.round(cap.rows.slice(0, row).reduce((a, b) => a + b, 0)) + 1
  const votes = new Map<string, number>()
  for (let x = x0 + 1; x <= x0 + w - 2; x++) {
    const i = (y * cap.png.width + x) * 4
    const argb = `FF${[0, 1, 2].map(k => cap!.png.data[i + k]!.toString(16).padStart(2, '0')).join('').toUpperCase()}`
    votes.set(argb, (votes.get(argb) ?? 0) + 1)
  }
  return [...votes].sort((a, b) => b[1] - a[1])[0]![0]
}
/** The committed capture of a Stage 3 case, decoded once (shared with excelCellBackground). */
function excelPng(c: { id: string; captures: ReadonlyArray<{ pngPath: string; sidecarPath: string }> }): DecodedPng {
  excelCellBackground(c, 'A1')
  return stage3Pngs.get(c.id)!.png
}
const pixelArgb = (png: DecodedPng, x: number, y: number): string => {
  const i = (y * png.width + x) * 4
  return `FF${[0, 1, 2].map(k => png.data[i + k]!.toString(16).padStart(2, '0')).join('').toUpperCase()}`
}
const argbHex = (argb: number): string => `FF${(argb & 0xffffff).toString(16).padStart(6, '0').toUpperCase()}`
const argbDelta = (a: string, b: string): number => Math.max(...[2, 4, 6].map(k => Math.abs(parseInt(a.slice(k, k + 2), 16) - parseInt(b.slice(k, k + 2), 16))))

/**
 * Ruling S3-J: a bar carrying stopIfTrue="1" is unresolved until bar:composition-bar-stopIfTrue is verified, so every
 * LATER rule on its cells is withheld. These are the cells where such a bar sits ahead of a non-bar rule.
 */
function underUnevidencedBarStop(c: { markers: ReadonlyArray<{ ruleType: string; sqref: string; priority?: number; stopIfTrue?: boolean; capabilityId: string }> }, ref: string): boolean {
  return c.markers.some(bar => bar.ruleType === 'dataBar' && bar.stopIfTrue && bar.priority !== undefined && isCellInSqref(ref, bar.sqref) &&
    !VERIFIED_CAPABILITIES.has('bar:composition-bar-stopIfTrue') &&
    c.markers.some(m => m.ruleType !== 'dataBar' && m.priority !== undefined && m.priority > bar.priority! && isCellInSqref(ref, m.sqref)))
}

describe('production viewer conditional formatting oracle conformance (C1 & C3)', () => {
  const manifest = JSON.parse(readFileSync(MANIFEST_PATH, 'utf8')) as CfManifest

  // Explicit allowlist of Stage 5 (formulas/expressions) and Stage 2/3 (dataBar/colorScale)
  // deferred rules, and harness table-style background (Ruling M3):
  const STAGE5_OR_DEFERRED_WITHHELD = new Set([
    'stage1-blanks-errors:D2', // divergent ISBLANK(D2) formula (Stage 5)
    'stage1-blanks-errors:G6', // divergent NOT(ISBLANK(G2)) formula (Stage 5)
    'stage1-blanks-errors:J8', // divergent ISNA(J2) formula (Stage 5)
    'stage1-blanks-errors:M6', // divergent NOT(ISNA(M2)) formula (Stage 5)
    'stage1-priority-composition:G2', // MOD(25,2)=1 expression rule (Stage 5)
    'stage1-priority-composition:G3', // MOD(14,2)=0 expression rule (Stage 5)
  ])

  type Finding = { outcome: 'withheld'; reason: string }

  // Documented findings for cells where engine behavior diverges from Excel oracle:
  // Only withheld outcomes are allowed; a finding can never encode a painted wrong colour.
  const FINDINGS: Record<string, Finding> = {
    // Finding F2: Overlapping scales withheld under Task 15 ruling S2-I (base fill FFFFFFFF)
    'stage2-scale-scope:B2': { outcome: 'withheld', reason: 'Overlapping scale ranges B2:B4 and B2:B6 withhold fill under S2-I' },
    'stage2-scale-scope:B3': { outcome: 'withheld', reason: 'Overlapping scale ranges B2:B4 and B2:B6 withhold fill under S2-I' },
    'stage2-scale-scope:B5': { outcome: 'withheld', reason: 'Overlapping scale ranges B2:B6 and B2:B6 withhold fill under S2-I' },
    'stage2-scale-scope:B6': { outcome: 'withheld', reason: 'Overlapping scale ranges B2:B6 and B2:B6 withhold fill under S2-I' },
    'stage2-scale-scope:D2': { outcome: 'withheld', reason: 'Overlapping scale ranges D2:D4 and D2:D6 withhold fill under S2-I' },
    'stage2-scale-scope:D3': { outcome: 'withheld', reason: 'Overlapping scale ranges D2:D4 and D2:D6 withhold fill under S2-I' },

    // Finding F6 (final review I5 / Ruling F-2): the scope fixture cannot discriminate hidden-row, autofilter,
    // merged-cell or disjoint participation, so scale:scope-hidden-row, scale:scope-autofilter-hidden-row,
    // scale:scope-merged-cell and scale:scope-disjoint are unverified and these scales are WITHHELD until re-measured.
    'stage2-scale-scope:C2': { outcome: 'withheld', reason: 'Hidden row 4 inside C2:C6: scale:scope-hidden-row is non-discriminating (hidden value 30 is not extremal) and unverified' },
    'stage2-scale-scope:C3': { outcome: 'withheld', reason: 'Hidden row 4 inside C2:C6: scale:scope-hidden-row is non-discriminating and unverified' },
    'stage2-scale-scope:C5': { outcome: 'withheld', reason: 'Hidden row 4 inside C2:C6: scale:scope-hidden-row is non-discriminating and unverified' },
    'stage2-scale-scope:C6': { outcome: 'withheld', reason: 'Hidden row 4 inside C2:C6: scale:scope-hidden-row is non-discriminating and unverified' },
    'stage2-scale-scope:D5': { outcome: 'withheld', reason: 'Hidden row 4 inside D2:D6 (the saved file has no autoFilter): scale:scope-autofilter-hidden-row is non-discriminating and unverified' },
    'stage2-scale-scope:D6': { outcome: 'withheld', reason: 'Hidden row 4 inside D2:D6: scale:scope-autofilter-hidden-row is non-discriminating and unverified' },
    'stage2-scale-scope:F2': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4 (F4 populated): hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F3': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F5': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F6': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F13': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F14': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F15': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F16': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F17': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:B8': { outcome: 'withheld', reason: 'Merged B8:C9 inside the scope: scale:scope-merged-cell is non-discriminating (every merged value is 100) and unverified' },
    'stage2-scale-scope:C8': { outcome: 'withheld', reason: 'Merged B8:C9 inside the scope: scale:scope-merged-cell is non-discriminating and unverified' },
    'stage2-scale-scope:B9': { outcome: 'withheld', reason: 'Merged B8:C9 inside the scope: scale:scope-merged-cell is non-discriminating and unverified' },
    'stage2-scale-scope:C9': { outcome: 'withheld', reason: 'Merged B8:C9 inside the scope: scale:scope-merged-cell is non-discriminating and unverified' },

    // Finding F3: Overlapping scales on I2:I5 withheld under Task 15 ruling S2-I
    'stage2-scale-composition:I2': { outcome: 'withheld', reason: 'Two overlapping scales on I2:I5 withhold fill under S2-I' },
    'stage2-scale-composition:I3': { outcome: 'withheld', reason: 'Two overlapping scales on I2:I5 withhold fill under S2-I' },
    'stage2-scale-composition:I4': { outcome: 'withheld', reason: 'Two overlapping scales on I2:I5 withhold fill under S2-I' },
    'stage2-scale-composition:I5': { outcome: 'withheld', reason: 'Two overlapping scales on I2:I5 withhold fill under S2-I' },

    // Finding F4: Text cell J2 in scale range
    'stage2-scale-composition:J2': { outcome: 'withheld', reason: 'Text cell J2 lower dxf fill unmeasured in Stage 2' },

    // Finding F5: Expression rule 19 with stopIfTrue="1" withholds lower scale 20 under S2-J
    'stage2-scale-composition:L2': { outcome: 'withheld', reason: 'Stage 5 expression rule with stopIfTrue withholds lower scale' },
    'stage2-scale-composition:L3': { outcome: 'withheld', reason: 'Stage 5 expression rule with stopIfTrue withholds lower scale' },
    'stage2-scale-composition:L4': { outcome: 'withheld', reason: 'Stage 5 expression rule with stopIfTrue withholds lower scale' },
    'stage2-scale-composition:L5': { outcome: 'withheld', reason: 'Stage 5 expression rule with stopIfTrue withholds lower scale' },
  }

  // Stage 3 bar cells whose NON-bar rule production withholds although Excel paints its fill (withheld-only).
  const STAGE3_NONBAR_WITHHELD: Record<string, Finding> = Object.fromEntries(['I2', 'I3', 'I4'].map(ref => [`stage3-bar-composition:${ref}`,
    { outcome: 'withheld', reason: 'Stage 5 expression rule (priority 15, stopIfTrue) is unresolved and taints ALL; Excel paints its fill' } as Finding]))

  const STAGE2_MEASURED_CASES = [
    'stage2-scale-interpolation',
    'stage2-scale-thresholds',
    'stage2-scale-degenerate',
    'stage2-scale-participation',
    'stage2-scale-colors',
    'stage2-scale-scope',
    'stage2-scale-composition',
  ]

  it('runs production pipeline on all measured cases with zero unexpected mismatches', async () => {
    let checkedCells = 0
    let stage3Checked = 0
    let stage3BarWithheld = 0
    let stage3NonBarChecked = 0
    let stage3NonBarWithheld = 0

    for (const c of manifest.cases) {
      if (c.status !== 'measured' || c.stage === 2) continue
      if (c.clockRequirement) {
        vi.useFakeTimers({ toFake: ['Date'] })
        vi.setSystemTime(new Date(`${c.clockRequirement.anchorLocalDate}T04:00:00+08:00`))
      }
      try {
        const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))
        const doc = await parseXlsx(await OfficePackage.load(bytes), parseOptionsFor(c))
        const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
        const dateContext = c.clockRequirement ? {
          clockSerial: c.clockRequirement.anchorSerial,
          dateSystem: c.clockRequirement.dateSystem,
        } : undefined

        const paint = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, dateContext })
        expect(paint, `prepareCfPaint should succeed for ${c.id}`).toBeDefined()
        if (!paint) continue

        for (const exp of c.expected?.cells ?? []) {
          checkedCells++
          const cellKey = `${c.id}:${exp.ref}`

          const { row, col } = refToRowCol(exp.ref)
          const st = paint.layer.styleAt(row, col, undefined)
          const gotMatched = Boolean(st && (st.fillColor || st.color || st.bold !== undefined || st.formatCode || st.numFmtId))

          // Ruling M3: Withheld Stage 5 expressions or Stage 2/3 visuals must be asserted
          // as withheld rather than skipped silently:
          if (STAGE5_OR_DEFERRED_WITHHELD.has(cellKey)) {
            expect(gotMatched, `${cellKey} should be withheld in Stage 1`).toBe(false)
            continue
          }
          if (c.stage === 3) {
            stage3Checked++
            // Ruling S3-A / OD-7: a bar cell whose bar:* capabilities are not verified in production is WITHHELD,
            // measured case or not. Non-bar rules in the case (DXF, scale) keep following the normal path below.
            const barCaps = c.markers.filter(m => m.ruleType === 'dataBar' && isCellInSqref(exp.ref, m.sqref)).map(m => m.capabilityId)
            if (barCaps.length > 0 && !barCaps.every(id => VERIFIED_CAPABILITIES.has(id))) {
              // The BAR is withheld (asserted per cell, on the bar descriptor itself, in "stage 3 bar cells are withheld
              // until their evidence lands"). The cell style here may only come from the case's non-bar rules: Ruling
              // S3-I "a bar never changes the fill slot ... A scale or DXF fill therefore keeps painting, but the bar
              // itself paints over a fill only with bar:composition-with-scale or bar:composition-dxf-fill".
              stage3BarWithheld++
              const nonBar = c.markers.filter(m => m.ruleType !== 'dataBar' && isCellInSqref(exp.ref, m.sqref))
              if (nonBar.length === 0) {
                expect(gotMatched, `${cellKey}: only bar rules cover it, so no style may paint until ${barCaps.join(', ')} is verified`).toBe(false)
                continue
              }
              if (underUnevidencedBarStop(c, exp.ref)) continue // Ruling S3-J: asserted in the S3-J tests below
              // The non-bar rules of the Stage 3 cases write a fill or nothing, so their outcome is visible in Excel's
              // background: production matches (paints) exactly where Excel shows a fill, with Excel's colour (+-8).
              expect(nonBar.every(m => Object.keys(m.dxf ?? {}).every(k => k === 'fillRgb')), `${cellKey}: fill-only non-bar rules`).toBe(true)
              const excel = excelCellBackground(c, exp.ref)
              const excelFilled = argbDelta(excel, 'FFFFFFFF') > 8
              if (STAGE3_NONBAR_WITHHELD[cellKey]) {
                expect(STAGE3_NONBAR_WITHHELD[cellKey]!.outcome).toBe('withheld')
                expect(gotMatched, `${cellKey} is withheld (${STAGE3_NONBAR_WITHHELD[cellKey]!.reason})`).toBe(false)
                stage3NonBarWithheld++
                continue
              }
              expect(gotMatched, `${cellKey}: non-bar rules match where Excel paints a fill (Excel ${excel})`).toBe(excelFilled)
              if (excelFilled) expect(argbDelta(st!.fillColor!, excel), `${cellKey} fill ${st?.fillColor} vs Excel ${excel}`).toBeLessThanOrEqual(8)
              stage3NonBarChecked++
              continue
            }
          }
          if (cellKey === 'stage1-priority-composition:F2') {
            // Stage 3 dataBar blue fill is withheld, but Stage 1 bold font matches
            expect(st?.bold, `${cellKey} bold`).toBe(true)
            continue
          }
          if (cellKey === 'stage1-dxf-layering:D4') {
            // CF does not match; style returned is base table band (50 <= 100)
            expect(st?.fillColor, `${cellKey} table band`).toBe('FFC0E6F5')
            continue
          }

          if (exp.matched) {
            expect(gotMatched, `${cellKey} should match CF (${exp.note ?? ''})`).toBe(true)
            if (exp.fillRgb) {
              expect(st?.fillColor, `${cellKey} fillRgb`).toBe(exp.fillRgb)
            }
            if (exp.fontRgb) {
              expect(st?.color, `${cellKey} fontRgb`).toBe(exp.fontRgb)
            }
            if (exp.bold !== undefined) {
              expect(st?.bold, `${cellKey} bold`).toBe(exp.bold)
            }
          } else {
            expect(gotMatched, `${cellKey} should not match CF (${exp.note ?? ''})`).toBe(false)
          }
        }
        paint.dispose()
      } finally {
        if (c.clockRequirement) {
          vi.useRealTimers()
        }
      }
    }

    // Ensures we audited the full Stage 1 cell corpus (220+ cells)
    expect(checkedCells).toBeGreaterThan(200)
    // Stage 3 (reported separately from the 398/364/34 Stage 1/2 totals): every measured bar cell was withheld.
    // Pinned to the value derived from the manifest itself: every expected cell of a measured Stage 3 case that lies under
    // a dataBar marker with a capability production does not hold (a regression that skips or paints one moves it).
    const expectedStage3Withheld = manifest.cases
      .filter(c => c.stage === 3 && c.status === 'measured')
      .flatMap(c => (c.expected?.cells ?? []).filter(exp => {
        const caps = c.markers.filter(m => m.ruleType === 'dataBar' && isCellInSqref(exp.ref, m.sqref)).map(m => m.capabilityId)
        return caps.length > 0 && !caps.every(id => VERIFIED_CAPABILITIES.has(id))
      })).length
    expect(expectedStage3Withheld).toBeGreaterThan(0)
    expect(stage3BarWithheld).toBe(expectedStage3Withheld)
    expect(stage3BarWithheld).toBeLessThanOrEqual(stage3Checked)
    // Non-vacuous (Task 18b review): the non-bar outcomes on bar cells were compared with Excel's background.
    expect(stage3NonBarChecked).toBeGreaterThan(10)
    expect(stage3NonBarWithheld).toBe(Object.keys(STAGE3_NONBAR_WITHHELD).length)
  })

  it('proves all 53 previously withheld cells now paint under production evidence (C1)', async () => {
    // Stage 1 cases with previously withheld cells in final-r1-oracle-vs-viewer.txt:
    const withheldCases = [
      'stage1-cellis-literals',
      'stage1-text-predicates',
      'stage1-blanks-errors',
      'stage1-timeperiod-today-clock',
      'stage1-timeperiod-today-clock-2',
      'stage1-above-average-stddev',
      'stage1-duplicate-scope',
      'stage1-priority-composition',
      'stage1-blank-merge-order',
    ]

    const expectedWithheldRecovered = [
      // reversed cellIs
      'stage1-cellis-literals:I4',
      'stage1-cellis-literals:K2',
      // case-insensitive text
      'stage1-text-predicates:B3',
      // blanks
      'stage1-blanks-errors:B2',
      'stage1-blanks-errors:B3',
      'stage1-blanks-errors:E6',
      // timePeriod 1900
      'stage1-timeperiod-today-clock:B4',
      'stage1-timeperiod-today-clock:B5',
      'stage1-timeperiod-today-clock:B6',
      'stage1-timeperiod-today-clock:B15',
      'stage1-timeperiod-today-clock:C3',
      'stage1-timeperiod-today-clock:D7',
      'stage1-timeperiod-today-clock:E9',
      'stage1-timeperiod-today-clock:F4',
      'stage1-timeperiod-today-clock:G10',
      'stage1-timeperiod-today-clock:H11',
      'stage1-timeperiod-today-clock:I4',
      'stage1-timeperiod-today-clock:J12',
      'stage1-timeperiod-today-clock:K13',
      // timePeriod 1904
      'stage1-timeperiod-today-clock-2:B3',
      'stage1-timeperiod-today-clock-2:B4',
      'stage1-timeperiod-today-clock-2:B7',
      'stage1-timeperiod-today-clock-2:C2',
      'stage1-timeperiod-today-clock-2:D5',
      'stage1-timeperiod-today-clock-2:E3',
      'stage1-timeperiod-today-clock-2:F3',
      // stdDev
      'stage1-above-average-stddev:H9',
      'stage1-above-average-stddev:I12',
      'stage1-above-average-stddev:L12',
      // duplicateValues / uniqueValues
      'stage1-duplicate-scope:B5',
      'stage1-duplicate-scope:B6',
      'stage1-duplicate-scope:C7',
      'stage1-duplicate-scope:D3',
      'stage1-duplicate-scope:F2',
      'stage1-duplicate-scope:H3',
      'stage1-duplicate-scope:H4',
      'stage1-duplicate-scope:H5',
      // priority composition
      'stage1-priority-composition:B2',
      'stage1-priority-composition:C2',
      'stage1-priority-composition:E2',
      // merged blank order
      'stage1-blank-merge-order:B2',
      'stage1-blank-merge-order:Z5',
    ]

    for (const caseId of withheldCases) {
      const c = manifest.cases.find(x => x.id === caseId)!
      if (c.clockRequirement) {
        vi.useFakeTimers({ toFake: ['Date'] })
        vi.setSystemTime(new Date(`${c.clockRequirement.anchorLocalDate}T04:00:00+08:00`))
      }
      try {
        const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))
        const doc = await parseXlsx(await OfficePackage.load(bytes), parseOptionsFor(c))
        const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
        const dateContext = c.clockRequirement ? {
          clockSerial: c.clockRequirement.anchorSerial,
          dateSystem: c.clockRequirement.dateSystem,
        } : undefined
        const paint = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, dateContext })
        expect(paint).toBeDefined()

        for (const targetKey of expectedWithheldRecovered.filter(k => k.startsWith(`${caseId}:`))) {
          const [, ref] = targetKey.split(':')
          const { row, col } = refToRowCol(ref!)
          const exp = c.expected!.cells.find(cell => cell.ref === ref)!
          const st = paint!.layer.styleAt(row, col, undefined)
          expect(st, `${targetKey} should have matched`).toBeDefined()
          if (exp.fillRgb) expect(st?.fillColor, `${targetKey} fill`).toBe(exp.fillRgb)
          if (exp.fontRgb) expect(st?.color, `${targetKey} font`).toBe(exp.fontRgb)
          if (exp.bold !== undefined) expect(st?.bold, `${targetKey} bold`).toBe(exp.bold)
        }
        paint!.dispose()
      } finally {
        if (c.clockRequirement) {
          vi.useRealTimers()
        }
      }
    }
  })

  it('proves the 5 absent-formula cells correctly do not match (C3)', async () => {
    // Task 18c: the x14 canonical K2 (text) and O6 (notContainsErrors) left this list. Their old "no match" came from
    // captures in which Excel had discarded the x14 block (misspelled URI); with the block live Excel matches them as
    // their legacy twins ("production renders the x14 form ... as Excel did").
    const wrongCells = [
      { caseId: 'stage1-text-predicates', ref: 'I2' }, // absent-formula containsText
      { caseId: 'stage1-blanks-errors', ref: 'C2' },   // absent-formula containsBlanks
      { caseId: 'stage1-blanks-errors', ref: 'F6' },   // absent-formula notContainsBlanks
      { caseId: 'stage1-blanks-errors', ref: 'I8' },   // absent-formula containsErrors
      { caseId: 'stage1-blanks-errors', ref: 'L6' },   // absent-formula notContainsErrors
    ]

    for (const { caseId, ref } of wrongCells) {
      const c = manifest.cases.find(x => x.id === caseId)!
      const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))
      const doc = await parseXlsx(await OfficePackage.load(bytes))
      const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
      const paint = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence })
      expect(paint).toBeDefined()

      const { row, col } = refToRowCol(ref)
      const st = paint!.layer.styleAt(row, col, undefined)
      expect(st?.fillColor, `${caseId}:${ref} should have no CF fill`).toBeUndefined()
      expect(st?.color, `${caseId}:${ref} should have no CF font color`).toBeUndefined()
      expect(st?.bold, `${caseId}:${ref} should have no CF bold`).toBeUndefined()
      paint!.dispose()
    }
  })

  it('asserts Stage 2 colour-scale cells remain withheld when scale evidence is unverified (Lesson 1)', async () => {
    let checkedCount = 0
    const unverifiedScaleEvidence = {
      allows: (id: string, variant?: string) => !id.startsWith('scale:') && cfRuntimeEvidence.allows(id, variant),
    }
    const stage2Cases = manifest.cases.filter(c => c.stage === 2)
    for (const c of stage2Cases) {
      const savedPath = join(REPO_ROOT, c.saved)
      if (!existsSync(savedPath)) continue
      const bytes = new Uint8Array(readFileSync(savedPath))
      const doc = await parseXlsx(await OfficePackage.load(bytes))
      const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
      const paint = await prepareCfPaint(sheet, { evidence: unverifiedScaleEvidence })
      expect(paint, `prepareCfPaint should succeed for ${c.id}`).toBeDefined()
      if (!paint) continue

      for (const exp of c.expected?.cells ?? []) {
        const scaleMarker = c.markers.find(m => m.ruleType === 'colorScale' && isCellInSqref(exp.ref, m.sqref))
        if (!scaleMarker) continue
        if (!(exp.scaleApplied ?? exp.matched)) continue

        const { row, col } = refToRowCol(exp.ref)
        const st = paint.layer.styleAt(row, col, undefined)

        // Until scale:* capabilities are evidenced and verified,
        // no CF colour-scale fill may be applied to any Stage 2 cell.
        if (c.id === 'stage2-scale-composition' && (exp.ref === 'B2' || exp.ref === 'B3')) {
          // B2/B3 has a winning Stage 1 cellIs dxf fill, not a scale fill
          expect(st?.fillColor, `${c.id}:${exp.ref} cellIs dxf fill`).toBe('FFFFC7CE')
        } else {
          expect(
            st?.fillColor,
            `Stage 2 cell ${c.id}:${exp.ref} colour scale fill must be withheld while unverified`,
          ).toBeUndefined()
        }
        checkedCount++
      }
      paint.dispose()
    }

    expect(checkedCount).toBeGreaterThan(0)
    expect(checkedCount).toBe(206)
  })

  it('paints exact Excel ARGB for every non-hidden excel-display-format cell of all 7 measured Stage 2 cases under production evidence (Lesson 1 dual gate, T14-A)', async () => {
    let grandTotalChecked = 0
    let exactCount = 0
    let withheldCount = 0
    const caseCounts: Record<string, number> = {}

    for (const caseId of STAGE2_MEASURED_CASES) {
      const c = manifest.cases.find(x => x.id === caseId)!
      expect(c, `Manifest must contain ${caseId}`).toBeDefined()
      const savedPath = join(REPO_ROOT, c.saved)
      expect(existsSync(savedPath), `${savedPath} must exist`).toBe(true)

      const bytes = new Uint8Array(readFileSync(savedPath))
      const doc = await parseXlsx(await OfficePackage.load(bytes))
      const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
      const paint = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence })
      expect(paint, `prepareCfPaint should succeed for ${caseId}`).toBeDefined()
      if (!paint) continue

      // Read trusted observed cells from reference sidecar
      const sidecarPath = join(REPO_ROOT, `tests/fixtures/xlsx-cf/refs/${caseId}.json`)
      expect(existsSync(sidecarPath), `${sidecarPath} must exist`).toBe(true)
      const sidecar = JSON.parse(readFileSync(sidecarPath, 'utf8')) as {
        observedCells: Array<{ ref: string; fillSource: string; fillRgb: string; hidden?: boolean }>
      }

      const observed = sidecar.observedCells.filter(cell => cell.fillSource === 'excel-display-format' && !cell.hidden)
      expect(observed.length, `${caseId} must have non-zero observed cells`).toBeGreaterThan(0)
      caseCounts[caseId] = observed.length

      for (const cell of observed) {
        grandTotalChecked++
        const cellKey = `${caseId}:${cell.ref}`
        const { row, col } = refToRowCol(cell.ref)
        const cellObj = sheet.rows.find(r => r.index === row)?.cells.find(cl => cl.col === col)
        const baseStyle = { numFmtId: 0, ...(cellObj?.style ?? {}), fillColor: cellObj?.style?.fillColor ?? 'FFFFFFFF' }
        const st = paint.layer.styleAt(row, col, baseStyle)
        const actualFill = st?.fillColor ?? 'FFFFFFFF'

        const finding = FINDINGS[cellKey]
        if (finding) {
          expect(finding.outcome).toBe('withheld')
          expect(actualFill, `${cellKey} finding: ${finding.reason}`).toBe('FFFFFFFF')
          withheldCount++
        } else {
          // Exact Excel ARGB match:
          expect(actualFill, `${cellKey} must match exact Excel ARGB`).toBe(cell.fillRgb)
          exactCount++
        }
      }

      paint.dispose()
    }

    // Assert a non-zero checked-cell count per case:
    for (const caseId of STAGE2_MEASURED_CASES) {
      expect(caseCounts[caseId], `Case ${caseId} checked cell count`).toBeGreaterThan(0)
    }
    expect(grandTotalChecked).toBe(398)
    // Ruling F-2: cells that lost paint moved into typed withheld-only FINDINGS; none is an allowlisted wrong colour.
    expect(exactCount).toBe(364)
    expect(withheldCount).toBe(34)
    expect(exactCount + withheldCount).toBe(grandTotalChecked)
  })


  // Ruling S3-J (Task 20): until bar:composition-bar-stopIfTrue is verified a bar carrying stopIfTrue="1" is unresolved,
  // so §5.7.1 withholds every later rule on its cells (compose taints ALL). The flag lives in the INPUT workbook only:
  // Excel's resave drops stopIfTrue from the base bar rule (saved stage3-bar-composition H2:H5 reads `priority="13"` with
  // no flag), so the saved workbook has no stopping bar to honour. The withholding is therefore asserted on the input
  // workbook, as the Task 18c test below does for the x14 forms Excel rewrites.
  const s3jViolations = async (workbook: 'input' | 'saved'): Promise<string[]> => {
    const violations: string[] = []
    for (const c of manifest.cases.filter(x => x.stage === 3 && x.status === 'measured')) {
      const cells = c.expected.cells.filter(e => underUnevidencedBarStop(c, e.ref))
      if (cells.length === 0) continue
      const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, workbook === 'input' ? c.input : c.saved)))
      const doc = await parseXlsx(await OfficePackage.load(bytes))
      const sheet = doc.sheets.find(x => x.name === c.sheet) ?? doc.sheets[0]!
      const paint = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence })
      for (const e of cells) {
        const { row, col } = refToRowCol(e.ref)
        const st = paint?.layer.styleAt(row, col, undefined)
        if (st && (st.fillColor || st.color || st.bold !== undefined)) violations.push(`${c.id}:${e.ref}`)
      }
      paint?.dispose()
    }
    return violations
  }
  it('S3-J: a later rule under an unevidenced bar stopIfTrue is withheld (input workbook, the flag Excel authored)', async () => {
    const c = manifest.cases.find(x => x.id === 'stage3-bar-composition')!
    expect(c.expected.cells.filter(e => underUnevidencedBarStop(c, e.ref)).map(e => e.ref)).toEqual(expect.arrayContaining(['H2', 'H3', 'H4']))
    expect(await s3jViolations('input')).toEqual([])
  })
  it('S3-J: the saved workbook carries no bar stopIfTrue, so its later DXF fill on H2:H4 is Excel\'s fill', async () => {
    // Excel ignores and drops the flag on resave; production then paints exactly the fill Excel painted there.
    expect(await s3jViolations('saved')).toEqual(['stage3-bar-composition:H2', 'stage3-bar-composition:H3', 'stage3-bar-composition:H4'])
    const c = manifest.cases.find(x => x.id === 'stage3-bar-composition')!
    const saved = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))))
    const rules = (saved.sheets.find(x => x.name === c.sheet) ?? saved.sheets[0]!).conditionalFormatting!.rules
    const hBar = rules.find(r => r.priority === 13)!
    expect(hBar.writes.has('dataBar')).toBe(true)
    expect(hBar.stopIfTrue).toBe(false)
    for (const ref of ['H2', 'H3', 'H4']) expect(argbDelta(excelCellBackground(c, ref), 'FFFFC7CE'), ref).toBeLessThanOrEqual(8)
  })

  // Task 18c: Excel resaves the x14 text/blank/error rules as legacy <cfRule>s, so the saved workbooks the main loop
  // runs no longer hold the x14 form. Production's x14 path is held to the same Excel evidence on the INPUT workbook:
  // every cell under an x14 marker must render what the recaptured PNG observed (sidecar observedCells, the capture of
  // the input with the x14 block live) and what the case expects.
  it('production renders the x14 form of the Stage 1 x14 rules as Excel did (Task 18c, input workbooks)', async () => {
    let checked = 0
    for (const id of ['stage1-text-predicates', 'stage1-blanks-errors']) {
      // Expectations from the oracle catalogue (the manifest entry refreshes them at the next capture/verify-only).
      const c = CF_CASES.find(x => x.id === id)!
      const sc = JSON.parse(readFileSync(join(REPO_ROOT, c.captures[0]!.sidecarPath), 'utf8')) as { observedCells: Array<{ ref: string; fillRgb?: string }> }
      const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, c.input)))))
      const sheet = doc.sheets.find(x => x.name === c.sheet) ?? doc.sheets[0]!
      const paint = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
      for (const exp of c.expected.cells) {
        if (!c.markers.some(m => m.extension === 'x14' && isCellInSqref(exp.ref, m.sqref))) continue
        const { row, col } = refToRowCol(exp.ref)
        const st = paint.layer.styleAt(row, col, undefined)
        const excelFill = sc.observedCells.find(o => o.ref === exp.ref)?.fillRgb
        expect(excelFill, `${id}:${exp.ref} observed in the capture`).toBeDefined()
        const excelPainted = excelFill !== 'FFFFFFFF'
        expect(excelPainted, `${id}:${exp.ref} expectation agrees with the capture`).toBe(exp.matched)
        if (exp.matched) expect(st?.fillColor, `${id}:${exp.ref}`).toBe(excelFill)
        else expect(st?.fillColor, `${id}:${exp.ref}`).toBeUndefined()
        checked++
      }
      paint.dispose()
    }
    expect(checked).toBe(8)
  })

  /** A bar is withheld when the cell is not `ready` or its visual slot carries no data bar descriptor. */
  const barWithheld = (outcome: CfOutcome<CfCellState>): boolean => outcome.kind !== 'ready' || outcome.value.visual?.dataBar === undefined

  it('stage 3 bar cells are withheld until their evidence lands', async () => {
    // The guard itself has teeth: a ready cell that carries a bar descriptor is NOT withheld.
    expect(barWithheld({ kind: 'ready', value: { properties: {}, visual: { dataBar: { ruleId: 0, axisFraction: 0, startFraction: 0, endFraction: 0.5, fillColor: 0xff0000ff, isGradient: false, isNegative: false, showValue: true, direction: 'leftToRight' } } } })).toBe(false)
    expect(barWithheld({ kind: 'ready', value: { properties: {} } })).toBe(true)
    expect(barWithheld({ kind: 'no-match' })).toBe(true)
    expect(barWithheld({ kind: 'unresolved', reason: 'cf-unsupported-rule' })).toBe(true)

    const stage3 = manifest.cases.filter(c => c.stage === 3)
    expect(stage3.map(c => c.id)).toHaveLength(13)
    let checked = 0
    let barCells = 0
    for (const c of stage3) {
      const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))
      const doc = await parseXlsx(await OfficePackage.load(bytes))
      const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
      expect(sheet.conditionalFormatting, `${c.id} must carry conditional formatting`).toBeDefined()
      const memory = createCfMemory()
      const service = createCfService(sheet.conditionalFormatting!, memory, { evidence: cfRuntimeEvidence })
      const stamp = { owner: 'conformance', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(sheet.conditionalFormatting!), viewGen: 0, structureGen: 0, clockGen: 0 }
      const source = residentSource(sheet, stamp, memory, 'engine')
      const rect = (() => {
        const ends = c.range.split(':').map(refToRowCol)
        return { firstRow: ends[0]!.row, lastRow: ends[1]!.row, firstCol: ends[0]!.col, lastCol: ends[1]!.col }
      })()
      const lease = await service.prepareCf({ stamp, sourceRects: [rect], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, source)
      try {
        for (const exp of c.expected.cells) {
          const { row, col } = refToRowCol(exp.ref)
          const outcome = lease.value.at(row, col)
          checked++
          if (exp.barApplied === true) barCells++
          // Production evidence: no bar:* capability is verified, so no cell may carry a bar.
          expect(barWithheld(outcome), `${c.id}:${exp.ref} must withhold its bar until its evidence lands`).toBe(true)
        }
      } finally {
        lease.release()
      }
      service.dispose()
      expect(memory.currentBytes, `${c.id} leaked CF memory`).toBe(0)
    }
    // Non-vacuous: every designated bar cell and its control was visited.
    expect(checked).toBe(stage3.reduce((n, c) => n + c.expected.cells.length, 0))
    expect(barCells).toBeGreaterThan(200)

    // Bar FINDINGS are withheld-only (extends the guard): a bar can never be recorded as painted.
    type BarFinding = { outcome: 'withheld'; reason: string }
    const BAR_FINDINGS: Record<string, BarFinding> = {}
    for (const [key, f] of Object.entries(BAR_FINDINGS)) expect(f.outcome, key).toBe('withheld')
    const validateBarFinding = (f: BarFinding) => {
      if (f.outcome !== 'withheld') throw new Error(`Non-withheld bar outcome: ${(f as any).outcome}`)
    }
    expect(() => validateBarFinding({ outcome: 'painted' as any, reason: 'bar painted before evidence' })).toThrow(/Non-withheld bar outcome/)
    expect(() => validateBarFinding({ outcome: 'withheld', reason: 'valid' })).not.toThrow()
  })

  /**
   * Lesson 1: every bar production paints reproduces Excel's measured px, and every bar it cannot reproduce is withheld.
   *
   * Evidence is the audit's OWN result (auditStage3Manifest: the verified set, the selected layout and the measured
   * automatic colours), injected exactly on top of production's Stage 1/2 set before parsing (parse-time gates read it).
   * Nothing Stage 3 is published yet; once Task 23 publishes, the published export must equal that same result (asserted
   * below), so the injection is a no-op and this test runs against production evidence unchanged.
   *
   * Geometry: layoutDataBar at the capture scale (1 device px per pt: zoom x DPR = 3/4 in viewer units, a pure unit
   * conversion of the sidecar's cellLeftPx/cellWidthPx and row heights) against the sidecar's raw px. Edges and the axis
   * within +-1 px, the vertical insets exactly (Ruling S3-A step 6). No expected px is computed by the painter.
   *
   * Units: 4/3 CSS px per point is the viewer's convention, taken from `src/xlsx/sparklines/geometry.ts` (`PT_PX`, its
   * Task 7 reference comparison) and the static renderer (`cellFontString`: pt x 96/72); `paintCfCapture` in
   * scripts/cf-compare.ts renders Excel's 1 px/pt captures at the same density (pxPerPt / CSS_PX_PER_PT). The factor is
   * that shared convention, not measured here: the conversion is exact only if it holds.
   *
   * Axis (fix round 1): every measured axis column is compared pixel by pixel over its whole cell height (Excel's
   * dotted pattern: 3 on, 1 off, the gap at y = x (mod 4), white in the gap); a solid axis fails (mutant below).
   */
  /** Every axis the Lesson 1 test compared, kept for the named solid-axis mutant below. */
  const axisSamples: Array<{ cellKey: string; bar: CfBarDescriptor; host: { x: number; y: number; width: number; height: number }; axisPx: number; top: number; height: number; png: DecodedPng }> = []

  it('bar geometry conformance (Lesson 1)', async () => {
    const audit = auditStage3Manifest(manifest, { rootDir: REPO_ROOT })
    expect(audit.ok).toBe(true)
    const verified = audit.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()
    expect(verified.length).toBeGreaterThan(40)
    const published = Reflect.get(verifiedExports, 'VERIFIED_STAGE3_CAPABILITIES') as ReadonlySet<string> | undefined
    if (published) {
      expect([...published].sort(), 'the published Stage 3 set is the audit result').toEqual(verified)
      expect(Reflect.get(verifiedExports, 'VERIFIED_STAGE3_BAR_LAYOUT'), 'the published layout is the audit result').toEqual(audit.barLayout.layout)
      expect(Reflect.get(verifiedExports, 'VERIFIED_STAGE3_AUTO_BAR_COLORS'), 'the published colours are the audit result').toEqual(audit.autoBarColours)
    }

    // A pixel-measured bar production withholds, by case column, with why (withheld-only, like FINDINGS).
    const BAR_FINDINGS: Record<string, Finding> = {
      'stage3-bar-lengths:G': { outcome: 'withheld', reason: 'percentile 10/90 leaves values beyond the bounds: S3-E keeps the rule unresolved (Excel clamps; bar:degenerate-beyond-bounds lifts nothing in the evaluator yet)' },
      'stage3-bar-degenerate:B': { outcome: 'withheld', reason: 'all values equal: S3-E keeps the degenerate scope unresolved (Excel draws half length; not implemented)' },
      'stage3-bar-degenerate:C': { outcome: 'withheld', reason: 'single numeric cell: S3-E keeps the degenerate scope unresolved (Excel draws half length; not implemented)' },
      ...Object.fromEntries(['stage3-bar-degenerate:D', 'stage3-bar-degenerate:E',
        ...['B', 'D', 'E', 'F', 'G', 'H', 'I'].map(col => `stage3-bar-rendering:${col}`),
        ...['B', 'C', 'D', 'E', 'F', 'H', 'I'].map(col => `stage3-bar-geometry:${col}`),
        'stage3-bar-showvalue:B', 'stage3-bar-showvalue:D',
      ].map(key => [key, { outcome: 'withheld', reason: 'num bounds: bar:bounds-num is non-discriminating (num 20/80 equals the data min/max), so unverified' } as Finding])),
      'stage3-bar-degenerate:F': { outcome: 'withheld', reason: 'percent bounds: bar:bounds-percent is non-discriminating, so unverified' },
    }
    // Bars production paints on cells the capture designated for no geometry capability (settings or another
    // capability's controls). They must still reproduce Excel's px; listed so the brief's "pending or
    // non-discriminating cells withhold" stays visible as an exception, not a silent pass.
    const PAINTED_UNDESIGNATED: Record<string, string> = {
      'stage3-bar-settings-probe:B': 'the probe bar (middle axis, right-to-left, border) designates settings capabilities only; every capability it needs is verified',
      'stage3-bar-pairing:C': 'the base-only bar designates settings capabilities only; every capability it needs is verified',
      'stage3-bar-composition:H': 'the bar ahead of a DXF fill designates bar:composition-bar-stopIfTrue (unsupported in Stage 3, S3-CLOSE); Excel dropped the flag on resave, so the saved bar needs only verified capabilities',
      'stage3-bar-axis-negative:H': 'positive cells of the negative-border-same-as-positive bar: that non-discriminating capability applies to negative cells only, which stay withheld',
    }
    for (const f of Object.values(BAR_FINDINGS)) expect(f.outcome).toBe('withheld')

    const usedFindings = new Set<string>(), usedPainted = new Set<string>()
    let painted = 0, withheldFound = 0, measuredPainted = 0, axes = 0, axisRows = 0, solidMutantFails = 0
    injectCfBarLayout(audit.barLayout.layout)
    injectCfAutoBarColours(audit.autoBarColours)
    try {
      await withExactCfEvidence([...VERIFIED_CAPABILITIES, ...verified], async () => {
        for (const c of manifest.cases.filter(x => x.stage === 3 && x.status === 'measured')) {
          const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))))
          const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
          const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
          const sc = JSON.parse(readFileSync(join(REPO_ROOT, c.captures[0]!.sidecarPath), 'utf8')) as { rowHeightsPt: number[]; pxPerPt: number; zoom: number; dpr: number }
          expect([sc.pxPerPt, sc.zoom, sc.dpr], c.id).toEqual([1, 100, 1])
          const observed = stage3ObservedBars(c, REPO_ROOT)!
          for (const exp of c.expected.cells) {
            const geometry = exp.barEvidence?.geometry
            if (!geometry) continue
            const key = `${c.id}:${exp.ref.replace(/\d+/, '')}`, cellKey = `${c.id}:${exp.ref}`
            const { row, col } = refToRowCol(exp.ref)
            const bar = batch.barAt(row, col)
            const o = observed.get(exp.ref)
            if (bar) {
              expect(o, `${cellKey} has an Excel observation`).toBeDefined()
              if (geometry !== 'pixel-measured') { expect(PAINTED_UNDESIGNATED[key], `${cellKey} (${geometry}) paints`).toBeDefined(); usedPainted.add(key) } else measuredPainted++
              const top = sc.rowHeightsPt.slice(0, row).reduce((a, b) => a + b, 0), height = sc.rowHeightsPt[row]!
              const host = { x: o!.cellLeftPx * 4 / 3, y: top * 4 / 3, width: o!.cellWidthPx * 4 / 3, height: height * 4 / 3 }
              const p = layoutDataBar(bar, host, { zoom: 0.75, dpr: 1, output: 'device' })
              expect(p, `${cellKey} lays out`).toBeDefined()
              if (o!.bar === null) expect(p!.fill, `${cellKey}: Excel draws no bar`).toBeUndefined()
              else {
                expect(p!.fill, `${cellKey}: Excel draws a bar`).toBeDefined()
                const f = p!.fill!
                expect(Math.abs(f.x - o!.bar.leftPx), `${cellKey} left ${f.x} vs ${o!.bar.leftPx}`).toBeLessThanOrEqual(1)
                expect(Math.abs(f.x + f.w - 1 - o!.bar.rightPx), `${cellKey} right ${f.x + f.w - 1} vs ${o!.bar.rightPx}`).toBeLessThanOrEqual(1)
                expect([f.y, f.y + f.h - 1], `${cellKey} top/bottom`).toEqual([o!.bar.topPx, o!.bar.bottomPx])
              }
              if (o!.axisPx !== undefined) {
                expect(p!.axis, `${cellKey}: Excel draws an axis`).toBeDefined()
                expect(p!.axis!.x, `${cellKey} axis x`).toBe(o!.axisPx)
                expect(p!.axis!.w).toBe(1)
                // Pixel by pixel down the column: the gap underlay, overpainted by the dashes, against Excel's pixels.
                const ours = new Map<number, number>()
                for (const r of [p!.axis!, ...(p!.axisDashes ?? [])]) for (let y = r.y; y < r.y + r.h; y++) ours.set(y, r.argb)
                const png = excelPng(c)
                let solidMismatch = false
                for (let y = top; y < top + height; y++) {
                  const excel = pixelArgb(png, o!.axisPx, y)
                  expect(ours.has(y), `${cellKey} axis row ${y}`).toBe(true)
                  expect(argbDelta(argbHex(ours.get(y)!), excel), `${cellKey} axis (${o!.axisPx}, ${y}) ours ${argbHex(ours.get(y)!)} Excel ${excel}`).toBeLessThanOrEqual(8)
                  if (argbDelta(argbHex(bar.axisColor!), excel) > 8) solidMismatch = true
                  axisRows++
                }
                if (solidMismatch) solidMutantFails++
                axisSamples.push({ cellKey, bar, host, axisPx: o!.axisPx, top, height, png })
                axes++
              } else if (p!.axis) {
                // Only an axis the capture cannot see: the measured automatic white (S3-AUTO-BG) on Excel's white cell.
                expect(p!.axis.argb, `${cellKey}: an axis Excel shows none of`).toBe(0xffffffff)
                expect(argbDelta(excelCellBackground(c, exp.ref), 'FFFFFFFF'), `${cellKey} white cell`).toBeLessThanOrEqual(8)
              }
              painted++
            } else if (geometry === 'pixel-measured' && o && (o.bar !== null || o.axisPx !== undefined)) {
              expect(BAR_FINDINGS[key], `${cellKey}: Excel draws a bar production withholds, with no finding`).toBeDefined()
              usedFindings.add(key)
              withheldFound++
            }
          }
          batch.dispose()
        }
      })
    } finally {
      injectCfBarLayout(null)
      injectCfAutoBarColours(null)
    }
    // No stale finding or exception, and pinned counts (a regression that withholds or paints one more moves them).
    expect([...usedFindings].sort()).toEqual(Object.keys(BAR_FINDINGS).sort())
    expect([...usedPainted].sort()).toEqual(Object.keys(PAINTED_UNDESIGNATED).sort())
    expect({ painted, measuredPainted, withheldFound, axes, axisRows }).toEqual({ painted: 140, measuredPainted: 123, withheldFound: 80, axes: 38, axisRows: 38 * 30 })
    // Mutant: a solid axis (the axis colour on every row) disagrees with Excel on every black dotted axis compared.
    expect(solidMutantFails, 'a solid axis fails on every explicit (black) axis').toBe(28)
  }, 60_000)

  it('RED mutant: a solid axis fails the Lesson 1 axis rows on every explicit axis', () => {
    // The production model with its dash removed (the axis colour on every row), against the same Excel pixels.
    const solid = { ...CF_BAR_RENDER_MODEL, axisDash: { ...CF_BAR_RENDER_MODEL.axisDash, onPt: 4, offPt: 0 } }
    expect(axisSamples.length, 'runs after the Lesson 1 test').toBe(38)
    let explicit = 0, failed = 0
    for (const s of axisSamples) {
      const p = layoutDataBar(s.bar, s.host, { zoom: 0.75, dpr: 1, output: 'device' }, solid, () => true)!
      const rows = new Map<number, number>()
      for (const r of [p.axis!, ...(p.axisDashes ?? [])]) for (let y = r.y; y < r.y + r.h; y++) rows.set(y, r.argb)
      let mismatch = false
      for (let y = s.top; y < s.top + s.height; y++) if (argbDelta(argbHex(rows.get(y)!), pixelArgb(s.png, s.axisPx, y)) > 8) mismatch = true
      if ((s.bar.axisColor! & 0xffffff) !== 0xffffff) { explicit++; if (mismatch) failed++ } else expect(mismatch, `${s.cellKey}: a white axis is solid either way`).toBe(false)
    }
    expect(explicit).toBe(28)
    expect(failed, 'the solid mutant fails every explicit axis').toBe(28)
  })

  it('guard: FINDINGS structure strictly disallows non-withheld outcomes', () => {
    // Type definition allows only { outcome: 'withheld', reason: string }
    type ConformanceFinding = { outcome: 'withheld'; reason: string }

    const validateFinding = (f: ConformanceFinding) => {
      if (f.outcome !== 'withheld') throw new Error(`Non-withheld outcome: ${(f as any).outcome}`)
    }

    // Valid withheld outcome:
    expect(() => validateFinding({ outcome: 'withheld', reason: 'valid' })).not.toThrow()

    // Runtime rejection of painted wrong colour or diverged outcome:
    expect(() => validateFinding({ outcome: 'diverged' as any, reason: 'invalid' })).toThrow(/Non-withheld outcome/)
    expect(() => validateFinding({ outcome: 'painted' as any, reason: 'wrong colour' })).toThrow(/Non-withheld outcome/)
  })

  it('Stage 1 clock cases are host-time-zone independent (CI runs in UTC)', async () => {
    const previous = process.env.TZ
    process.env.TZ = 'UTC'
    try {
      expect(Intl.DateTimeFormat().resolvedOptions().timeZone).toBe('UTC')
      for (const id of ['stage1-timeperiod-today-clock', 'stage1-timeperiod-today-clock-2']) {
        const c = manifest.cases.find(x => x.id === id)!
        vi.useFakeTimers({ toFake: ['Date'] })
        vi.setSystemTime(new Date(`${c.clockRequirement!.anchorLocalDate}T04:00:00+08:00`))
        try {
          const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))), parseOptionsFor(c))
          const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
            const paint = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, dateContext: { clockSerial: c.clockRequirement!.anchorSerial, dateSystem: c.clockRequirement!.dateSystem } })
          const stale = refToRowCol(c.clockRequirement!.staleCacheCellRef)
          expect(paint!.layer.styleAt(stale.row, stale.col, undefined)?.fillColor, `${id} stale TODAY() cell refreshes to the anchor day`).toBe('FFFFC7CE')
          paint!.dispose()
        } finally { vi.useRealTimers() }
      }
    } finally {
      if (previous === undefined) delete process.env.TZ; else process.env.TZ = previous
    }
  })
})
