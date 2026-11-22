/**
 * The Stage 2 exact-ARGB conformance pass (Lesson 1 dual gate, T14-A), shared by tests/xlsx-cf-oracle-conformance.test.ts
 * and the num-segment rehearsal (tests/xlsx-cf-scale-num-segment-case.test.ts), which runs it under a proposed
 * registration. Every non-hidden excel-display-format cell of each named measured case is production's paint against
 * Excel's fill: a typed finding (STAGE2_WITHHELD_FINDINGS) must withhold, the num-segment and statistical withholds must
 * withhold where they hold, every other cell must be Excel's exact ARGB. Violations are returned, never skipped.
 */
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { cfDocumentEvaluationOf, prepareCfPaint } from '../../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../../src/xlsx/cf/predicates'
import { VERIFIED_CAPABILITIES } from '../../src/xlsx/cf/evidence'
import { STAGE2_WITHHELD_FINDINGS, stage2NumSegmentShape, stage2UpperSegmentWholeProductShape } from '../../scripts/cf-refs/stage2-withhold'
import type { CfManifest } from '../../scripts/cf-refs/cases'

export interface Stage2ConformanceResult {
  total: number
  exact: number
  withheld: number
  numSegmentWithheld: number
  upperWholeWithheld: number
  stubCompared: number
  caseCounts: Record<string, number>
  violations: string[]
}

const rc = (ref: string): { row: number; col: number } => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!
  return { row: Number(m[2]) - 1, col: [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 }
}

/** Ruling S5-STUB: the stub cells of stage2-scale-composition (L2:L5) compare once their successors are verified, or when painted. */
const STAGE2_STUB_SUCCESSORS = ['formula:dialect-base', 'formula:anchor-single', 'formula:truth-boolean', 'formula:fn-math', 'formula:composition-stop-scale', 'formula:composition-stub-successor-scale']
const defaultStubCompare = (_key: string, gotMatched: boolean): boolean => gotMatched || STAGE2_STUB_SUCCESSORS.every(id => VERIFIED_CAPABILITIES.has(id))

export async function stage2Conformance(manifest: CfManifest, rootDir: string, caseIds: readonly string[], stubCompare: (cellKey: string, gotMatched: boolean) => boolean = defaultStubCompare): Promise<Stage2ConformanceResult> {
  const out: Stage2ConformanceResult = { total: 0, exact: 0, withheld: 0, numSegmentWithheld: 0, upperWholeWithheld: 0, stubCompared: 0, caseCounts: {}, violations: [] }
  for (const caseId of caseIds) {
    const c = manifest.cases.find(x => x.id === caseId)
    if (!c || !existsSync(join(rootDir, c.saved))) { out.violations.push(`${caseId}: missing case or saved workbook`); continue }
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(rootDir, c.saved)))))
    const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
    const paint = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, evaluation: cfDocumentEvaluationOf(doc) })
    if (!paint) { out.violations.push(`${caseId}: prepareCfPaint failed`); continue }
    const sidecar = JSON.parse(readFileSync(join(rootDir, c.captures.find(x => x.appearance === 'screen')!.sidecarPath), 'utf8')) as {
      observedCells: Array<{ ref: string; fillSource: string; fillRgb: string; hidden?: boolean }>
    }
    const observed = sidecar.observedCells.filter(cell => cell.fillSource === 'excel-display-format' && !cell.hidden)
    out.caseCounts[caseId] = observed.length
    const values = new Map<string, number>()
    for (const r of sheet.rows) for (const cl of r.cells) if (typeof cl.value === 'number') values.set(`${colName(cl.col)}${r.index + 1}`, cl.value)
    for (const cell of observed) {
      out.total++
      const key = `${caseId}:${cell.ref}`
      const { row, col } = rc(cell.ref)
      const cellObj = sheet.rows.find(r => r.index === row)?.cells.find(cl => cl.col === col)
      const baseStyle = { numFmtId: 0, ...(cellObj?.style ?? {}), fillColor: cellObj?.style?.fillColor ?? 'FFFFFFFF' }
      const actual = paint.layer.styleAt(row, col, baseStyle)?.fillColor ?? 'FFFFFFFF'
      const finding = STAGE2_WITHHELD_FINDINGS[key]
      const stub = finding !== undefined && /^stage2-scale-composition:L[2-5]$/.test(key) && stubCompare(key, actual !== 'FFFFFFFF')
      if (stub) out.stubCompared++
      const withhold = (counter?: 'numSegmentWithheld' | 'upperWholeWithheld'): void => {
        if (actual !== 'FFFFFFFF') out.violations.push(`${key}: expected withheld, painted ${actual}`)
        out.withheld++
        if (counter) out[counter]++
      }
      if (finding && !stub) withhold()
      else if (stage2NumSegmentShape(c.markers, cell.ref, cellObj?.value)) withhold('numSegmentWithheld')
      else if (stage2UpperSegmentWholeProductShape(c.markers, cell.ref, cellObj?.value, values)) withhold('upperWholeWithheld')
      else if (actual !== cell.fillRgb) out.violations.push(`${key}: production ${actual} is not Excel's ${cell.fillRgb}`)
      else out.exact++
    }
    paint.dispose()
  }
  return out
}

function colName(col: number): string {
  let s = ''
  for (let c = col + 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s
  return s
}
