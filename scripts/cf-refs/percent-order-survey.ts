/**
 * The percent-order follow-up (brief cf-percent-order, Phase 1 task 1): the survey of every captured percent / percentile
 * cfvo of a colour scale, data bar or icon set, and of every captured cell whose outcome depends on the order in which the
 * threshold's arithmetic is done.
 *
 * The two orders (pre-registered readings, recipe §15):
 * - `lerp-first` (production today): percent t = lo + (p/100)·(hi − lo) (`lerp`, colorScale.ts); percentile rank
 *   1 + (p/100)·(n − 1), then t = x_i + frac·(x_{i+1} − x_i) (`percentileOrderStatistic`, 'inc');
 * - `multiply-first`: percent t = lo + ((hi − lo)·p)/100; percentile rank 1 + (p·(n − 1))/100, then the same interpolation.
 *
 * Scope: the rule's numeric, non-error cells, each bound computed over two scopes, every stored cell (`all`) and the
 * cells of visible rows (`visible`), so a bound is order-dependent when it differs under either. Out-of-domain percent /
 * percentile (p < 0 or p > 100) is excluded: production refuses that shape outright (Ruling F-2), so no order applies.
 *
 * Per order-dependent cell: the outcome under each order (colour scale: `scaleColorAt` with production's evidence; icon
 * set: `iconBucket`), Excel's observation where the capture has one, and for a formula-resolved p the 2×2 of order ×
 * {as-literal, fifteen-digit}. A data-bar bound that differs between the orders throws (the survey has no bar model; none
 * is captured). Nothing here is registered or reachable from production.
 *
 * Run: `npx tsx scripts/cf-refs/percent-order-survey.ts` (prints the table of docs/validation/xlsx-cf-recipes-stage2.md §15.1).
 */
import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { scaleColorAt } from '../../src/xlsx/cf/colorScale'
import { iconBucket } from '../../src/xlsx/cf/iconSet'
import type { CfScalePoint } from '../../src/xlsx/cf/model'
import { caseParseOptionsOf } from './case-clock'
import type { CfManifest, CfPartMarker } from './cases'
import { isCellInSqref } from './verify'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')

/** The captures the survey was run before (ecf00906): left out, so the survey stays as run (recipe §15.1). */
export const PERCENT_ORDER_SURVEY_LATER_CASES: readonly string[] = ['stage2-scale-percent-order', 'stage5-icon-percent-order']

export type PercentOrder = 'lerp-first' | 'multiply-first'
export const PERCENT_ORDERS: readonly PercentOrder[] = ['lerp-first', 'multiply-first']

/** A percent threshold under an order (restated independently of production). */
export function surveyPercent(lo: number, hi: number, p: number, order: PercentOrder): number {
  return order === 'lerp-first' ? lo + (p / 100) * (hi - lo) : lo + ((hi - lo) * p) / 100
}

/** A PERCENTILE.INC threshold under an order (restated independently of production); undefined outside 0..100. */
export function surveyPercentile(sorted: readonly number[], p: number, order: PercentOrder): number | undefined {
  const n = sorted.length
  if (n === 0 || !(p >= 0 && p <= 100)) return undefined
  if (p === 0) return sorted[0]
  if (p === 100) return sorted[n - 1]
  const rank = order === 'lerp-first' ? 1 + (p / 100) * (n - 1) : 1 + (p * (n - 1)) / 100
  const ip = Math.floor(rank), frac = rank - ip, i = ip - 1
  if (i >= n - 1) return sorted[n - 1]
  return sorted[i]! + frac * (sorted[i + 1]! - sorted[i]!)
}

export interface SurveyBound {
  caseId: string
  stage: number
  rule: 'colorScale' | 'dataBar' | 'iconSet'
  sqref: string
  index: number
  type: 'percent' | 'percentile'
  text: string
  formula: boolean
  p: number
  /** Per scope: the threshold under each order, and whether they differ. */
  scopes: Record<'all' | 'visible', { n: number; lerp: number | undefined; multiply: number | undefined; differs: boolean }>
}

export interface SurveyCell {
  caseId: string
  sqref: string
  ref: string
  value: number
  scope: 'all' | 'visible'
  lerp: string
  multiply: string
  excel: string | undefined
  /** A formula-resolved p: the outcome under order × {as-literal, fifteen-digit}. */
  grid?: Record<string, string>
}

const colName = (c: number): string => { let s = ''; for (let x = c + 1; x > 0; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(65 + ((x - 1) % 26)) + s; return s }
type Cfvo = { type: string; val?: string; f?: string; gte?: boolean }
const cfvoOf = (k: CfPartMarker): Cfvo[] => (k.scale?.cfvo ?? k.bar?.cfvo ?? k.icon?.cfvo ?? []) as Cfvo[]

/** Excel's observation at a cell: a Stage 2 screen sidecar's DisplayFormat fill, a Stage 5 formula observation; else undefined. */
function observationOf(sidecar: Record<string, unknown>, ref: string): string | undefined {
  for (const o of (Array.isArray(sidecar.observedCells) ? sidecar.observedCells : []) as Array<Record<string, unknown>>) {
    if (String(o.ref).toUpperCase() === ref && o.fillSource === 'excel-display-format' && typeof o.fillRgb === 'string') return o.fillRgb.toUpperCase()
  }
  const rec = (sidecar.formulaObservations as Record<string, { displayFormatFill?: string }> | undefined)?.[ref]
  return typeof rec?.displayFormatFill === 'string' ? rec.displayFormatFill.toUpperCase() : undefined
}

export async function percentOrderSurveyOf(manifestPath = resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), root = REPO_ROOT): Promise<{ bounds: SurveyBound[]; excluded: SurveyBound[]; cells: SurveyCell[] }> {
  const m = JSON.parse(readFileSync(manifestPath, 'utf8')) as CfManifest
  const bounds: SurveyBound[] = []
  const excluded: SurveyBound[] = []
  const cells: SurveyCell[] = []
  for (const c of m.cases) {
    if (c.status !== 'measured' || PERCENT_ORDER_SURVEY_LATER_CASES.includes(c.id)) continue
    const vis = c.markers.filter(k => k.ruleType === 'colorScale' || k.ruleType === 'dataBar' || k.ruleType === 'iconSet')
    if (!vis.some(k => cfvoOf(k).some(v => v.type === 'percent' || v.type === 'percentile'))) continue
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(root, c.saved)))), c.stage === 5 ? caseParseOptionsOf(c) : {})
    const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
    const values = new Map<string, { value: unknown; error: boolean; hidden: boolean }>()
    for (const r of sheet.rows) for (const cl of r.cells) values.set(`${colName(cl.col)}${r.index + 1}`, { value: cl.value, error: cl.valueIsError === true, hidden: r.hidden === true })
    const screen = c.captures.find(x => x.appearance === 'screen')
    const sidecar = screen ? (JSON.parse(readFileSync(resolve(root, screen.sidecarPath), 'utf8')) as Record<string, unknown>) : {}
    for (const k of vis) {
      const cfvo = cfvoOf(k)
      const inScope = [...values].filter(([ref, x]) => isCellInSqref(ref, k.sqref) && typeof x.value === 'number' && !x.error)
      const scopeOf = (s: 'all' | 'visible') => inScope.filter(([, x]) => s === 'all' || !x.hidden)
      const resolveP = (v: Cfvo): number | undefined => {
        const t = v.f ?? v.val
        if (t !== undefined && /^-?\d+(\.\d+)?(e[+-]?\d+)?$/i.test(String(t))) return Number(t)
        const mm = /^=?\$?([A-Z]+)\$?(\d+)$/.exec(String(t ?? ''))
        const x = mm ? values.get(`${mm[1]}${mm[2]}`)?.value : undefined
        return typeof x === 'number' ? x : undefined
      }
      const thresholds = (scope: number[], order: PercentOrder, round: boolean): Array<number | undefined> => {
        const sorted = [...scope].sort((a, b) => a - b)
        const min = sorted[0]!, max = sorted.at(-1)!
        return cfvo.map(v => {
          if (v.type === 'min') return min
          if (v.type === 'max') return max
          let p = resolveP(v)
          if (p === undefined) return undefined
          if (round && (v.f !== undefined)) p = Number(p.toPrecision(15))
          if (v.type === 'percent') return surveyPercent(min, max, p, order)
          if (v.type === 'percentile') return surveyPercentile(sorted, p, order)
          return p
        })
      }
      const outcome = (v: number, t: Array<number | undefined>): string => {
        if (t.some(x => x === undefined)) return 'unresolved'
        if (k.ruleType === 'colorScale') {
          const pts: CfScalePoint[] = cfvo.map((x, i) => ({ bound: x.type === 'min' || x.type === 'max' ? { type: x.type } : { type: x.type === 'formula' ? 'num' : (x.type as 'num' | 'percent' | 'percentile'), value: 0 }, argb: parseInt(k.scale!.colors[i]!, 16) >>> 0 }) as CfScalePoint)
          const o = scaleColorAt(v, pts, t as number[])
          return o.kind === 'ready' ? hex(o.value) : `${o.kind}:${(o as { reason?: string }).reason ?? ''}`
        }
        if (k.ruleType === 'iconSet') return `bucket ${iconBucket(v, t as number[], cfvo.map(x => x.gte !== false))}`
        throw new Error(`${c.id} ${k.sqref}: a data-bar bound differs between the orders; the survey has no bar model`)
      }
      cfvo.forEach((v, index) => {
        if (v.type !== 'percent' && v.type !== 'percentile') return
        const p = resolveP(v)
        if (p === undefined || inScope.length === 0) throw new Error(`${c.id} ${k.sqref}: cfvo ${index} does not resolve`)
        const per = (s: 'all' | 'visible') => {
          const scope = scopeOf(s).map(([, x]) => x.value as number).sort((a, b) => a - b)
          if (scope.length === 0) return { n: 0, lerp: undefined, multiply: undefined, differs: false }
          const [l, mu] = PERCENT_ORDERS.map(o => (v.type === 'percent' ? surveyPercent(scope[0]!, scope.at(-1)!, p, o) : surveyPercentile(scope, p, o)))
          return { n: scope.length, lerp: l, multiply: mu, differs: !Object.is(l, mu) }
        }
        const row: SurveyBound = { caseId: c.id, stage: c.stage, rule: k.ruleType as SurveyBound['rule'], sqref: k.sqref, index, type: v.type, text: String(v.f ?? v.val), formula: v.f !== undefined, p, scopes: { all: per('all'), visible: per('visible') } }
        if (p < 0 || p > 100) { excluded.push(row); return }
        bounds.push(row)
      })
      for (const s of ['all', 'visible'] as const) {
        const scope = scopeOf(s)
        if (scope.length === 0) continue
        const nums = scope.map(([, x]) => x.value as number)
        if (cfvo.some(v => (v.type === 'percent' || v.type === 'percentile') && !((resolveP(v) ?? -1) >= 0 && (resolveP(v) ?? 101) <= 100))) continue
        const [tl, tm] = PERCENT_ORDERS.map(o => thresholds(nums, o, false))
        if (tl!.every((x, i) => Object.is(x, tm![i]))) continue
        for (const [ref, x] of scope) {
          const v = x.value as number
          const [a, b] = [outcome(v, tl!), outcome(v, tm!)]
          if (a === b) continue
          if (cells.some(e => e.caseId === c.id && e.sqref === k.sqref && e.ref === ref)) continue
          const cell: SurveyCell = { caseId: c.id, sqref: k.sqref, ref, value: v, scope: s, lerp: a, multiply: b, excel: observationOf(sidecar, ref) }
          if (cfvo.some(y => y.f !== undefined && (y.type === 'percent' || y.type === 'percentile'))) {
            cell.grid = Object.fromEntries(PERCENT_ORDERS.flatMap(o => [[`${o} as-literal`, outcome(v, thresholds(nums, o, false))], [`${o} fifteen-digit`, outcome(v, thresholds(nums, o, true))]]))
          }
          cells.push(cell)
        }
      }
    }
  }
  return { bounds, excluded, cells }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const { bounds, excluded, cells } = await percentOrderSurveyOf()
  const by = (f: (r: SurveyBound) => boolean): number => bounds.filter(f).length
  const differs = (r: SurveyBound): boolean => r.scopes.all.differs || r.scopes.visible.differs
  console.log(`bounds ${bounds.length} (percent ${by(r => r.type === 'percent')}, percentile ${by(r => r.type === 'percentile')}); excluded out-of-domain ${excluded.length}; cases ${new Set(bounds.map(r => r.caseId)).size}`)
  console.log('| Rule | Type | Bounds | Differ (either scope) |')
  console.log('|---|---|---|---|')
  for (const rule of ['colorScale', 'dataBar', 'iconSet'] as const) for (const type of ['percent', 'percentile'] as const) console.log(`| ${rule} | ${type} | ${by(r => r.rule === rule && r.type === type)} | ${by(r => r.rule === rule && r.type === type && differs(r))} |`)
  console.log('\nDiffering bounds:')
  for (const r of bounds.filter(differs)) console.log(`- ${r.caseId} ${r.sqref} cfvo ${r.index} ${r.type} ${r.text}${r.formula ? ' (formula)' : ''} p=${r.p}: all n=${r.scopes.all.n} ${r.scopes.all.lerp} / ${r.scopes.all.multiply}; visible n=${r.scopes.visible.n} ${r.scopes.visible.lerp} / ${r.scopes.visible.multiply}`)
  for (const r of excluded) console.log(`- excluded (out of domain): ${r.caseId} ${r.sqref} cfvo ${r.index} ${r.type} ${r.text}`)
  console.log('\nOrder-dependent cells:')
  for (const c of cells) console.log(`- ${c.caseId} ${c.ref}=${c.value} (${c.scope}): lerp-first ${c.lerp}, multiply-first ${c.multiply}, Excel ${c.excel ?? '(unobserved)'}${c.grid ? `; ${Object.entries(c.grid).map(([k, v]) => `${k} ${v}`).join(', ')}` : ''}`)
}
