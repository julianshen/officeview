// @vitest-environment node
/**
 * Task 35: every reading production registers (`CF_FORMULA_PRODUCTION_READINGS`) is the reading production implements,
 * shown on the Stage 5 fixtures themselves. Under the registered capabilities (and each listed selection in turn), the
 * fill, strikethrough and visual production shows at each designated cell of each configuration equal that reading's
 * predicted token (a `twin:` token resolved against production's own twin cell), or the cell is withheld (fail closed:
 * an unadmitted function, a whole-rule hoist refusal). Each entry is reproduced in full by at least one configuration,
 * except the named unreachable ones, and a rival reading is contradicted there, so the comparison discriminates.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cfDocumentEvaluationOf, cfEvaluationContextFor } from '../src/xlsx/cf/formula'
import { residentSource } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { createCfService } from '../src/xlsx/cf/service'
import { CF_LIMITS, createCfMemory } from '../src/xlsx/cf/limits'
import { injectCfFormulaSelections, withInjectedCfEvidence } from '../src/xlsx/cf/evidence'
import { CF_FORMULA_PRODUCTION_READINGS } from '../src/xlsx/cf/formulaReadings'
import { cfRulesVersion } from '../src/xlsx/view/pagedSource'
import type { CfManifest } from '../scripts/cf-refs/cases'
import { rc, shown } from './helpers/cf-formula-rules'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
/** Registered entries no Stage 5 configuration reproduces in full yet, with why (the cells it cannot reach withhold, never paint). */
const UNREACHABLE: Readonly<Record<string, string>> = {
  'formula:predicate-timeperiod-divergent': 'the divergent timePeriod formulas use TODAY(), and formula:fn-today is Task 38',
  'formula:truth-error': 'each configuration holds an NA() row, and NA is not on the allow-list (S5-B): those cells withhold, the 1/0 rows show error-no-match',
  'formula:fn-conditional-aggregate': 'its configuration\'s COUNTIF($A:$A,A2) column holds booleans and numeric text, a whole-rule hoist refusal (S5-HOIST); reproduced through the stage5-formula-hoist rows',
}
const REUSE_ONLY = /^formula:composition-stub-successor-/

describe('every registered formula reading is the reading production implements, on the Stage 5 fixtures', () => {
  it('each designated cell is the registered reading\'s prediction or withheld; each entry is reproduced and discriminating', async () => {
    const manifest = JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
    const registry = CF_FORMULA_PRODUCTION_READINGS
    const lists = Object.entries(registry).filter((e): e is [string, readonly string[]] => typeof e[1] !== 'string')
    const reproduced = new Set<string>(), violations: string[] = []
    let discriminating = 0
    for (const c of manifest.cases.filter(x => x.stage === 5)) {
      for (const cfg of c.formulaConfigurations ?? []) {
        if (cfg.record || cfg.role) continue
        const entry = registry[cfg.capabilityId]
        if (entry === undefined) continue
        for (const reading of typeof entry === 'string' ? [entry] : entry) {
          const predicted = cfg.hypotheses[reading] as Readonly<Record<string, string>> | undefined
          expect(predicted, `${cfg.capabilityId} ${cfg.id}: ${reading} is a hypothesis`).toBeDefined()
          const refs = cfg.refs.filter(r => !r.includes(':')) // build-time pre:/calc:/post: observations are not paint
          if (refs.length === 0) continue
          const selections = Object.fromEntries(lists.map(([k, v]) => [k, v[0]!]))
          selections[cfg.capabilityId] = reading
          const got = await withInjectedCfEvidence(Object.keys(registry), async () => {
            injectCfFormulaSelections(selections)
            try {
              const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))))
              const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
              const twins = Object.values(cfg.hypotheses).flatMap(h => Object.values(h)).filter(t => t.startsWith('twin:')).map(t => t.slice(5).replace('+strike', ''))
              const cells = [...new Set([...refs, ...twins])]
              const memory = createCfMemory()
              const service = createCfService(sheet.conditionalFormatting!, memory, { evidence: cfRuntimeEvidence, evaluation: cfEvaluationContextFor(cfDocumentEvaluationOf(doc), Date.now), sheetName: sheet.name })
              const stamp = { owner: 'readings', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(sheet.conditionalFormatting!), viewGen: 0, structureGen: 0, clockGen: 0 }
              const lease = await service.prepareCf({ stamp, sourceRects: cells.map(r => { const { row, col } = rc(r); return { firstRow: row, lastRow: row, firstCol: col, lastCol: col } }), signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, residentSource(sheet, stamp, memory, 'engine'))
              const out = new Map<string, string>()
              for (const r of cells) {
                const { row, col } = rc(r)
                const o = lease.value.at(row, col), s = shown(o)
                const visual = o.kind === 'ready' ? Object.keys(o.value.visual ?? {}).sort().join('&') : ''
                out.set(r, s.withheld ? 'withheld' : `${s.fill}${s.strike ? '+strike' : ''}${cfg.observable === 'crop' ? `|${visual}` : ''}`)
              }
              lease.release(); service.dispose()
              return out
            } finally { injectCfFormulaSelections(null) }
          })
          // The token a hypothesis predicts at `ref`, in production's terms (`twin:` resolved against production's twin).
          const want = (h: Readonly<Record<string, string>>, ref: string): string => {
            const p = h[ref]!, strike = p.endsWith('+strike'), base = p.replace('+strike', '')
            if (!base.startsWith('twin:')) return `${base}${strike ? '+strike' : ''}`
            const twin = got.get(base.slice(5))!
            return cfg.observable === 'crop' ? twin : `${twin.replace('+strike', '')}${strike ? '+strike' : ''}`
          }
          const bad = refs.filter(r => got.get(r) !== 'withheld' && got.get(r) !== want(predicted!, r))
          for (const r of bad) violations.push(`${cfg.capabilityId} ${c.id}/${cfg.id} ${reading} ${r}: production ${got.get(r)}, reading predicts ${want(predicted!, r)}`)
          if (bad.length === 0 && refs.every(r => got.get(r) !== 'withheld')) {
            reproduced.add(cfg.capabilityId)
            // Discrimination: some rival reading predicts something production does not show.
            if (Object.entries(cfg.hypotheses).some(([name, h]) => name !== reading && refs.some(r => want(h, r) !== got.get(r)))) discriminating++
            else violations.push(`${cfg.capabilityId} ${cfg.id}: no rival of ${reading} is contradicted (the comparison does not discriminate)`)
          }
        }
      }
    }
    expect(violations).toEqual([])
    const unreproduced = Object.keys(registry).filter(id => !reproduced.has(id) && !REUSE_ONLY.test(id) && UNREACHABLE[id] === undefined)
    expect(unreproduced).toEqual([])
    for (const id of Object.keys(UNREACHABLE)) expect(reproduced.has(id), `${id} is now reachable: drop it from UNREACHABLE`).toBe(false)
    expect(discriminating).toBeGreaterThan(30)
  }, 300_000)
})
