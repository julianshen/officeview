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
  'formula:composition-stop-icon': 'its continues witnesses K3/K5/K7 (a red fill plus an icon) are withheld in production, because icon:composition-dxf-fill is unpublished, so the configuration is unusable (S5-WITNESS); the compose gate is covered by tests/xlsx-cf-formula-compose.test.ts',
  'formula:fn-conditional-aggregate': 'its configuration\'s COUNTIF($A:$A,A2) column holds booleans and numeric text, a whole-rule hoist refusal (S5-HOIST); reproduced through the stage5-formula-hoist rows',
}
const REUSE_ONLY = /^formula:composition-stub-successor-/

type Config = NonNullable<CfManifest['cases'][number]['formulaConfigurations']>[number]
/** What production shows at one cell. */
export type Shown = Readonly<{ fill: string; strike: boolean; visual: string; withheld: boolean
  /** A bar or icon is drawn: a visual descriptor that is not a zero-length bar (S5-WITNESS: a bar at its lower bound is empty). */
  drawn: boolean }>
const tokenOf = (cfg: Config, s: Shown | undefined): string => !s ? 'missing' : s.withheld ? 'withheld' : `${s.fill}${s.strike ? '+strike' : ''}${cfg.observable === 'crop' ? `|${s.visual}` : ''}`
/** The cells a configuration's verdict reads: its designated cells, every twin its hypotheses name, every declared witness. */
const cellsOf = (cfg: Config): string[] => {
  const refs = cfg.refs.filter(r => !r.includes(':'))
  const twins = Object.values(cfg.hypotheses).flatMap(h => Object.values(h)).filter(t => t.startsWith('twin:')).map(t => t.slice(5).replace('+strike', ''))
  return [...new Set([...refs, ...twins, ...Object.keys(cfg.twinExpect ?? {})])]
}
/**
 * The verdict of one configuration under one reading, pure in what production showed (`got`). Returns its violations,
 * whether the reading is reproduced in full (no designated cell withheld), and whether a rival reading is contradicted.
 */
export function configurationVerdict(cfg: Config, reading: string, got: ReadonlyMap<string, Shown>): { violations: string[]; unusable: string[]; reproduced: boolean; discriminating: boolean } {
  const refs = cfg.refs.filter(r => !r.includes(':')) // build-time pre:/calc:/post: observations are not paint
  const predicted = cfg.hypotheses[reading] as Readonly<Record<string, string>>
  const violations: string[] = []
  // Ruling S5-WITNESS: every declared witness is validated against its own declaration before any twin observation is
  // used as an expectation. A crop witness is about its bar or icon (`visual`: one is drawn; `empty`: none, or a
  // zero-length bar); a fill witness about its fill or strikethrough. A withheld witness satisfies neither. A configuration
  // whose witnesses do not hold cannot be evidence: it is neither reproduced nor discriminating (`unusable`).
  const unusable: string[] = []
  for (const [ref, expected] of Object.entries(cfg.twinExpect ?? {})) {
    const w = got.get(ref)
    if (expected === 'either') continue
    const renders = w !== undefined && !w.withheld && (cfg.observable === 'crop' ? w.drawn : w.fill !== 'none' || w.strike)
    const empty = w !== undefined && !w.withheld && !renders
    if (expected === 'visual' ? !renders : !empty) unusable.push(`${cfg.id} witness ${ref}: declared ${expected}, production shows ${tokenOf(cfg, w)}`)
  }
  const witnessesHold = unusable.length === 0
  const want = (h: Readonly<Record<string, string>>, ref: string): string => {
    const p = h[ref]!, strike = p.endsWith('+strike'), base = p.replace('+strike', '')
    if (!base.startsWith('twin:')) return `${base}${strike ? '+strike' : ''}`
    const twin = tokenOf(cfg, got.get(base.slice(5)))
    return cfg.observable === 'crop' ? twin : `${twin.replace('+strike', '')}${strike ? '+strike' : ''}`
  }
  const bad = refs.filter(r => tokenOf(cfg, got.get(r)) !== 'withheld' && tokenOf(cfg, got.get(r)) !== want(predicted, r))
  for (const r of bad) violations.push(`${cfg.id} ${reading} ${r}: production ${tokenOf(cfg, got.get(r))}, reading predicts ${want(predicted, r)}`)
  const reproduced = witnessesHold && violations.length === 0 && refs.every(r => tokenOf(cfg, got.get(r)) !== 'withheld')
  const discriminating = reproduced && Object.entries(cfg.hypotheses).some(([name, h]) => name !== reading && refs.some(r => want(h, r) !== tokenOf(cfg, got.get(r))))
  if (reproduced && !discriminating) violations.push(`${cfg.id}: no rival of ${reading} is contradicted (the comparison does not discriminate)`)
  return { violations, unusable, reproduced, discriminating }
}

describe('the configuration verdict validates declared witnesses (fix round 1 #4)', () => {
  it('suppressing stop-bar E4 and its required positive twin H4 fails', () => {
    const manifest = JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
    const cfg = manifest.cases.find(x => x.id === 'stage5-composition')!.formulaConfigurations!.find(x => x.id === 'stop-bar')!
    expect(cfg.twinExpect?.['H4']).toBe('visual')
    // An in-memory world that satisfies `halts` everywhere: on rows 3/5/7 E shows the red-only twin F, elsewhere the bar-only twin H.
    const red: Shown = { fill: 'FFFFC7CE', strike: false, visual: '', withheld: false, drawn: false }, bar: Shown = { fill: 'none', strike: false, visual: 'dataBar', withheld: false, drawn: true }
    const redBar: Shown = { fill: 'FFFFC7CE', strike: false, visual: 'dataBar', withheld: false, drawn: true }, empty: Shown = { fill: 'none', strike: false, visual: '', withheld: false, drawn: false }
    // H2 holds the column minimum: a zero-length bar, declared empty.
    const zeroBar: Shown = { fill: 'none', strike: false, visual: 'dataBar', withheld: false, drawn: false }
    const world = new Map<string, Shown>()
    for (const r of [2, 3, 4, 5, 6, 7]) {
      const on = r % 2 === 1
      world.set(`E${r}`, on ? red : r === 2 ? zeroBar : bar); world.set(`F${r}`, red); world.set(`G${r}`, redBar); world.set(`H${r}`, r === 2 ? zeroBar : bar)
    }
    expect(configurationVerdict(cfg, 'halts', world)).toMatchObject({ violations: [], unusable: [], reproduced: true, discriminating: true })
    // Codex counterexample: E4 and H4 both suppressed. Production's twin alone would call it reproduced.
    const suppressed = new Map(world)
    suppressed.set('E4', empty); suppressed.set('H4', empty)
    const v = configurationVerdict(cfg, 'halts', suppressed)
    expect(v.reproduced).toBe(false)
    expect(v.discriminating).toBe(false)
    expect(v.unusable).toEqual(['stop-bar witness H4: declared visual, production shows none|'])
  })
})

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
              const cells = cellsOf(cfg)
              const memory = createCfMemory()
              const service = createCfService(sheet.conditionalFormatting!, memory, { evidence: cfRuntimeEvidence, evaluation: cfEvaluationContextFor(cfDocumentEvaluationOf(doc), Date.now), sheetName: sheet.name })
              const stamp = { owner: 'readings', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(sheet.conditionalFormatting!), viewGen: 0, structureGen: 0, clockGen: 0 }
              const lease = await service.prepareCf({ stamp, sourceRects: cells.map(r => { const { row, col } = rc(r); return { firstRow: row, lastRow: row, firstCol: col, lastCol: col } }), signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, residentSource(sheet, stamp, memory, 'engine'))
              const out = new Map<string, Shown>()
              for (const r of cells) {
                const { row, col } = rc(r)
                const o = lease.value.at(row, col), s = shown(o)
                const v = o.kind === 'ready' ? o.value.visual : undefined
                const drawn = v !== undefined && (v.iconSet !== undefined || (v.dataBar !== undefined && v.dataBar.endFraction !== v.dataBar.startFraction))
                out.set(r, { fill: s.fill, strike: s.strike, withheld: s.withheld, drawn, visual: v ? Object.keys(v).sort().join('&') : '' })
              }
              lease.release(); service.dispose()
              return out
            } finally { injectCfFormulaSelections(null) }
          })
          const verdict = configurationVerdict(cfg, reading, got)
          for (const v of verdict.violations) violations.push(`${cfg.capabilityId} ${c.id}/${v}`)
          if (verdict.reproduced) reproduced.add(cfg.capabilityId)
          if (verdict.discriminating) discriminating++
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
