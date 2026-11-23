/**
 * The bar all-equal follow-up (brief cf-bar-allequal, Phase 1 task 1): the census of every captured data bar whose scope's
 * numeric values are all equal (two or more numeric cells, v_min = v_max), with every single-numeric-cell scope listed
 * apart. For each such rule: its bounds (as authored; the x14 companion's cfvo and attributes win for a paired bar), lengths,
 * fill type, border, direction, axis, the common value against the resolved thresholds, Excel's measured px per cell, and
 * production's outcome today (the real service, `cellOutcomes` seam: the bar rule's own outcome, painted or its refusal).
 *
 * Excel's px:
 * - Stage 3: the sidecar's `observedBars` (the measurement the verifier re-derives from the hash-bound PNG);
 * - any other stage: the Stage 3 column scan on the committed PNG against the cell's recorded control
 *   (`measureBarGeometry`, as `s5BarVisual` does), `unmeasured` when the case records no control for the cell.
 * Lengths are inclusive px counts from the cell's leading inset (left edge + 2); `span` is the cell's inner width (w - 3).
 *
 * Nothing here is registered or reachable from production. Run: `npx tsx scripts/cf-refs/bar-all-equal-census.ts`.
 */
import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { cfEvaluationContextFor } from '../../src/xlsx/cf/formula'
import { createCfMemory, CF_LIMITS } from '../../src/xlsx/cf/limits'
import { createCfService } from '../../src/xlsx/cf/service'
import { cfRuntimeEvidence } from '../../src/xlsx/cf/predicates'
import { cfWorkbookOfDocument, residentSource } from '../../src/xlsx/cf/render'
import { cfRulesVersion } from '../../src/xlsx/view/pagedSource'
import type { CfOutcome } from '../../src/xlsx/cf/model'
import type { CfManifest, CfOracleCase } from './cases'
import { caseClockOf, caseDocumentEvaluationOf, caseParseOptionsOf } from './case-clock'
import { isCellInSqref, loadBarScreenCapture, measureBarGeometry } from './verify'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

type Cfvo = Readonly<{ type: string; val?: string; f?: string }>

export interface AllEqualCell {
  ref: string
  value: number
  /** Excel: `none`, `L..R` (absolute px, inclusive) or `unmeasured`. */
  excel: string
  /** Excel's bar length in px from the leading inset (undefined: no bar or unmeasured). */
  excelLen?: number
  cellLeftPx?: number
  cellWidthPx?: number
  /** Production today: `painted e=<endFraction>` or the bar rule's refusal reason. */
  production: string
}

export interface AllEqualRule {
  caseId: string
  stage: number
  sqref: string
  /** `distinct`: a scope whose numeric values differ (scanned for the completeness check only). */
  shape: 'all-equal' | 'single-cell' | 'distinct'
  pairing: 'paired' | 'base-only' | 'standalone'
  bounds: string
  lengths: string
  gradient: boolean
  border: boolean
  direction: string
  axis: string
  /** The common value, and where it sits against literal num bounds (`lo == hi` for equal literal bounds). */
  value: number
  position: string
  cells: AllEqualCell[]
}

const cfvoText = (c: Cfvo): string => c.type + (c.val !== undefined ? `=${c.val}` : c.f !== undefined ? `=${c.f}` : '')
const refRc = (ref: string): { row: number; col: number } => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!
  let col = 0
  for (const ch of m[1]!) col = col * 26 + (ch.charCodeAt(0) - 64)
  return { row: Number(m[2]) - 1, col: col - 1 }
}
const colLetter = (n: number): string => {
  let s = ''
  for (let c = n; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s
  return s
}
function expand(sqref: string): string[] {
  const out: string[] = []
  for (const tok of sqref.trim().split(/\s+/)) {
    const [a, b] = tok.split(':') as [string, string | undefined]
    const ma = /^([A-Z]+)(\d+)$/.exec(a)!, mb = /^([A-Z]+)(\d+)$/.exec(b ?? a)!
    const ca = refRc(`${ma[1]}1`).col + 1, cb = refRc(`${mb[1]}1`).col + 1
    for (let r = Math.min(+ma[2]!, +mb[2]!); r <= Math.max(+ma[2]!, +mb[2]!); r++) for (let c = Math.min(ca, cb); c <= Math.max(ca, cb); c++) out.push(`${colLetter(c)}${r}`)
  }
  return out
}

/** Literal num bound, else undefined. */
const literal = (c: Cfvo): number | undefined => {
  if (c.type !== 'num') return undefined
  const s = (c.val ?? c.f ?? '').trim()
  return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : undefined
}

export interface AllEqualCensus {
  /** Every all-equal and single-cell bar rule of every measured case, with Excel's px and production's outcome. */
  rules: AllEqualRule[]
  /** Bar rules scanned per measured case (every bar rule with a numeric cell in scope). */
  scanned: Record<string, number>
  /** Every scanned rule production refuses with cf-degenerate-scope at some cell, whatever its scope (completeness). */
  degenerate: Array<{ caseId: string; sqref: string; shape: AllEqualRule['shape']; bounds: string; position: string }>
}

export async function barAllEqualCensus(manifest: CfManifest, rootDir: string = REPO_ROOT): Promise<AllEqualCensus> {
  const rows: AllEqualRule[] = []
  const scanned: Record<string, number> = {}
  const degenerate: AllEqualCensus['degenerate'] = []
  for (const c of manifest.cases.filter(x => x.status === 'measured' && x.markers.some(m => m.ruleType === 'dataBar'))) {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(rootDir, c.saved)))), caseParseOptionsOf(c))
    const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
    const values = new Map<string, number>()
    for (const row of sheet.rows) for (const cell of row.cells) {
      if (typeof cell.value === 'number' && cell.valueIsError !== true) values.set(`${colLetter(cell.col + 1)}${row.index + 1}`, cell.value)
    }
    // One entry per bar rule: the legacy (base) marker, with its x14 companion's cfvo and attributes when paired.
    const bars = c.markers.filter(m => m.ruleType === 'dataBar' && (m.extension === 'legacy' || m.bar?.pairing === 'standalone'))
    const selected: AllEqualRule[] = []
    for (const m of bars) {
      const comp = m.bar?.pairing === 'paired' ? c.markers.find(x => x.ruleType === 'dataBar' && x.extension === 'x14' && x.sqref === m.sqref && x.bar?.pairing === 'paired') : undefined
      const def = comp ?? m
      const scope = expand(m.sqref).filter(r => values.has(r))
      if (scope.length === 0) continue
      const vs = scope.map(r => values.get(r)!)
      const equal = vs.every(v => v === vs[0])
      const cfvo = (def.bar?.cfvo ?? [{ type: 'min' }, { type: 'max' }]) as readonly Cfvo[]
      const attrs = (def.bar?.attrs ?? {}) as Record<string, string>
      const base = (m.bar?.attrs ?? {}) as Record<string, string>
      const pairing = (m.bar?.pairing ?? 'base-only') as AllEqualRule['pairing']
      const v = vs[0]!
      const lo = literal(cfvo[0]!), hi = literal(cfvo[1]!)
      const position = lo === undefined || hi === undefined ? 'statistical (lo = hi = v)'
        : lo === hi ? `lo == hi (${lo})` : lo > hi ? `inverted (${lo} > ${hi})`
          : v === hi ? 'at hi' : v === lo ? 'at lo' : v > lo && v < hi ? `interior f=${(v - lo) / (hi - lo)}` : 'beyond'
      selected.push({
        caseId: c.id, stage: c.stage, sqref: m.sqref, shape: !equal ? 'distinct' : scope.length === 1 ? 'single-cell' : 'all-equal', pairing,
        bounds: cfvo.map(cfvoText).join(' / '),
        lengths: pairing === 'base-only' ? `${base.minLength ?? '10'}/${base.maxLength ?? '90'}` : `${attrs.minLength ?? '10'}/${attrs.maxLength ?? '90'}`,
        gradient: pairing === 'base-only' ? true : attrs.gradient !== '0',
        border: attrs.border === '1',
        direction: attrs.direction ?? 'context',
        axis: attrs.axisPosition ?? 'automatic',
        value: v, position, cells: scope.map(ref => ({ ref, value: values.get(ref)!, excel: 'unmeasured', production: '?' })),
      })
    }
    scanned[c.id] = selected.length
    if (selected.length === 0) continue

    // Excel's px (all-equal and single-cell rules only). A Stage 3 cell the sidecar does not record (not designated) is
    // measured from the PNG against its same-content control (the case's control row offset), as any other stage's.
    const sidecar = JSON.parse(readFileSync(resolve(rootDir, c.captures[0]!.sidecarPath), 'utf8')) as { observedBars?: Array<{ ref: string; cellLeftPx: number; cellWidthPx: number; bar: { leftPx: number; rightPx: number } | null }> }
    const cap = loadBarScreenCapture(c as CfOracleCase, rootDir)
    const pair = c.expected.cells.find(e => e.matched && e.controlRef && /^[A-Z]+\d+$/.test(e.controlRef))
    const offset = c.stage === 3 && pair ? refRc(pair.controlRef!).row - refRc(pair.ref).row : undefined
    for (const r of selected.filter(x => x.shape !== 'distinct')) for (const cell of r.cells) {
      let o: { cellLeftPx: number; cellWidthPx: number; bar: { leftPx: number; rightPx: number } | null } | undefined
      if (c.stage === 3) o = sidecar.observedBars?.find(x => x.ref === cell.ref)
      if (!o && cap) {
        const { row } = refRc(cell.ref)
        const ctrl = c.expected.cells.find(e => e.ref === cell.ref)?.controlRef ?? (offset !== undefined ? `${cell.ref.replace(/\d+$/, '')}${row + 1 + offset}` : undefined)
        if (ctrl) {
          const res = measureBarGeometry(cap.png, cap.geom, [cell.ref], { controlOf: () => ctrl, axisArgb: 'FF000000', tolerancePx: 1 })
          if (!('kind' in res)) o = (res as ReadonlyMap<string, { cellLeftPx: number; cellWidthPx: number; bar: { leftPx: number; rightPx: number } | null }>).get(cell.ref)
        }
      }
      if (!o) continue
      cell.cellLeftPx = o.cellLeftPx
      cell.cellWidthPx = o.cellWidthPx
      cell.excel = o.bar ? `${o.bar.leftPx}..${o.bar.rightPx}` : 'none'
      if (o.bar) cell.excelLen = o.bar.rightPx - o.bar.leftPx + 1
    }

    // Production today: the bar rule's own outcome per cell (the service's cellOutcomes seam).
    const cf = sheet.conditionalFormatting!
    const memory = createCfMemory()
    const clock = caseClockOf(c)
    const evaluation = cfEvaluationContextFor(caseDocumentEvaluationOf(c, doc), clock?.now ?? (() => Date.now()))
    let current: Map<string, string> | undefined
    const service = createCfService(cf, memory, {
      evidence: cfRuntimeEvidence, evaluation, sheetName: sheet.name, workbook: cfWorkbookOfDocument(doc.sheets, memory, 'engine'),
      cellOutcomes: (_row, _col, rules, outcomes) => {
        for (let i = 0; i < rules.length; i++) {
          const rule = rules[i]!
          if (rule.definition.kind !== 'dataBar') continue
          const o = outcomes[i] as CfOutcome<{ endFraction: number; startFraction: number }> | undefined
          const key = rule.ranges.map(g => `${colLetter(g.firstCol + 1)}${g.firstRow + 1}:${colLetter(g.lastCol + 1)}${g.lastRow + 1}`).join(' ')
          current?.set(key, !o ? 'no outcome' : o.kind === 'unresolved' ? o.reason : o.kind === 'ready' ? `painted ${o.value.startFraction}..${o.value.endFraction}` : o.kind)
        }
      },
    })
    const stamp = { owner: 'bar-all-equal-census', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(cf), viewGen: 0, structureGen: 0, clockGen: 0 }
    try {
      const source = residentSource(sheet, stamp, memory, 'engine')
      for (const r of selected) for (const cell of r.cells) {
        const { row, col } = refRc(cell.ref)
        current = new Map()
        const lease = await service.prepareCf({ stamp, sourceRects: [{ firstRow: row, lastRow: row, firstCol: col, lastCol: col }], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, source)
        lease.release()
        const norm = (s: string) => s.replace(/([A-Z]+\d+):\1(?=\s|$)/g, '$1')
        const hit = [...current.entries()].find(([k]) => isCellInSqref(cell.ref, k) && norm(k) === norm(r.sqref)) ?? [...current.entries()].find(([k]) => isCellInSqref(cell.ref, k))
        cell.production = hit ? hit[1] : 'no bar rule reached'
        current = undefined
      }
    } finally { service.dispose() }
    for (const r of selected) {
      if (r.cells.some(x => x.production === 'cf-degenerate-scope')) degenerate.push({ caseId: r.caseId, sqref: r.sqref, shape: r.shape, bounds: r.bounds, position: r.position })
    }
    rows.push(...selected.filter(x => x.shape !== 'distinct'))
  }
  return { rules: rows, scanned, degenerate }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const manifest = JSON.parse(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
  void barAllEqualCensus(manifest).then(({ rules, scanned, degenerate }) => {
    console.log(`scanned bar rules per measured case: ${JSON.stringify(scanned)}`)
    console.log('rules production refuses with cf-degenerate-scope:')
    for (const d of degenerate) console.log(`  ${d.caseId} ${d.sqref} [${d.shape}] ${d.bounds}: ${d.position}`)
    for (const r of rules) {
      console.log(`${r.caseId} ${r.sqref} [${r.shape}] ${r.pairing} bounds ${r.bounds}; lengths ${r.lengths}; ${r.gradient ? 'gradient' : 'solid'}${r.border ? ' border' : ''}; ${r.direction}; axis ${r.axis}; v=${r.value} ${r.position}`)
      for (const c of r.cells) console.log(`  ${c.ref} cell ${c.cellLeftPx ?? '?'}+${c.cellWidthPx ?? '?'} excel ${c.excel}${c.excelLen !== undefined ? ` (len ${c.excelLen}, span ${(c.cellWidthPx ?? 0) - 3})` : ''} | production ${c.production}`)
    }
  })
}
