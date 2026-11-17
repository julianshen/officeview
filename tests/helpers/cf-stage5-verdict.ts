/**
 * Task 39 (fix rounds 2-4): the Stage 5 cell verdict shared by the conformance test and the mutant tests. Production's
 * outcome at one cell comes from the production service with production's inputs (the document's evaluation context, its
 * sheets as the cross-sheet seam); Excel's from the capture's DisplayFormat observation. Fix round 4: production must equal
 * the counterfactual run of the real compositor (`stage5CounterfactualOf`), and its resolved components Excel's.
 */
import { readFileSync, existsSync } from 'node:fs'
import { join } from 'node:path'
import { createCfService } from '../../src/xlsx/cf/service'
import { createCfMemory, CF_LIMITS } from '../../src/xlsx/cf/limits'
import { cfRuntimeEvidence } from '../../src/xlsx/cf/predicates'
import { cfDocumentEvaluationOf, cfWorkbookOfDocument, residentSource } from '../../src/xlsx/cf/render'
import { cfEvaluationContextFor } from '../../src/xlsx/cf/formula'
import { cfRulesVersion } from '../../src/xlsx/view/pagedSource'
import { stage5CellAgreement, stage5CounterfactualOf, stage5TokenOf, type Stage5Excel, type Stage5Production, type Stage5Truth } from '../../scripts/cf-refs/stage5-withhold'
import type { CfManifest } from '../../scripts/cf-refs/cases'
import type { XlsxDocument, XlsxSheet } from '../../src/xlsx/types'

type Case = CfManifest['cases'][number]
const rc = (ref: string) => { const m = /^([A-Z]+)(\d+)$/.exec(ref)!; return { row: Number(m[2]) - 1, col: [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 } }

/** Excel's DisplayFormat token at `ref` from the case's capture sidecar, or undefined unobserved. */
export function stage5ExcelAt(c: Case, rootDir: string, ref: string): Stage5Excel | undefined {
  const cap = c.captures.find(x => x.appearance === 'screen') ?? c.captures[0]
  if (!cap?.sidecarPath || !existsSync(join(rootDir, cap.sidecarPath))) return undefined
  const o = (JSON.parse(readFileSync(join(rootDir, cap.sidecarPath), 'utf8')) as { formulaObservations?: Record<string, { displayFormatFill?: string; baseFill?: string; strikethrough?: boolean }> }).formulaObservations?.[ref]
  if (!o || typeof o.displayFormatFill !== 'string') return undefined
  return { fill: o.displayFormatFill.toUpperCase() === (o.baseFill ?? 'FFFFFFFF').toUpperCase() ? 'none' : o.displayFormatFill.toUpperCase(), strike: o.strikethrough === true }
}

/** Production's token at `ref` through the production service (resident, production evidence and inputs). */
export async function stage5ProductionAt(doc: XlsxDocument, sheet: XlsxSheet, ref: string): Promise<Stage5Production> {
  const memory = createCfMemory()
  const cf = sheet.conditionalFormatting
  if (!cf) return stage5TokenOf(undefined)
  const service = createCfService(cf, memory, { evidence: cfRuntimeEvidence, evaluation: cfEvaluationContextFor(cfDocumentEvaluationOf(doc), Date.now), sheetName: sheet.name, workbook: cfWorkbookOfDocument(doc.sheets, memory, 'engine') })
  const stamp = { owner: 'stage5-verdict', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(cf), viewGen: 0, structureGen: 0, clockGen: 0 }
  try {
    const { row, col } = rc(ref)
    const lease = await service.prepareCf({ stamp, sourceRects: [{ firstRow: row, lastRow: row, firstCol: col, lastCol: col }], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, residentSource(sheet, stamp, memory, 'engine'))
    const r = lease.value.at(row, col)
    lease.release()
    return stage5TokenOf(r)
  } finally { service.dispose() }
}

/** The counterfactual at `ref` with the same production inputs (the evaluation clock read at call time). */
export async function stage5TruthAt(c: Case, doc: XlsxDocument, sheet: XlsxSheet, ref: string): Promise<Stage5Truth | undefined> {
  return (await stage5CounterfactualOf({ c, doc, sheet, refs: [ref], now: Date.now })).get(ref)
}

/** The verdict at one Stage 5 cell (undefined: production equals the counterfactual and its resolved components Excel's). */
export async function stage5SlotVerdict(c: Case, doc: XlsxDocument, sheet: XlsxSheet, ref: string, rootDir: string): Promise<string | undefined> {
  return stage5VerdictOf(c, ref, await stage5ProductionAt(doc, sheet, ref), rootDir, await stage5TruthAt(c, doc, sheet, ref))
}
/** The verdict for a given production token and counterfactual at `ref`. */
export function stage5VerdictOf(c: Case, ref: string, production: Stage5Production, rootDir: string, truth: Stage5Truth | undefined): string | undefined {
  const v = stage5CellAgreement(production, stage5ExcelAt(c, rootDir, ref), truth)
  return v.agrees ? undefined : `${c.id}:${ref}: ${v.disagreement}`
}
