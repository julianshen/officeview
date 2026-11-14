/**
 * Calibration probe for `CF_LIMITS.activeUnits` (Ruling S3-I1-UNITS): prints the work units ONE `prepareCf` call spends
 * on its per-call budget for each production call shape, and the idle units-per-millisecond rate on this machine, so a
 * reader can see what the unit cap bounds in wall time. Nothing here is asserted by a test and no outcome depends on
 * the timings: the counting budget below mirrors `createCfWorkBudget`'s cooperative yields but never refuses.
 *
 * Production call shapes (docs/validation/xlsx-cf.md section 11): the text path issues at most 1,000 cells per call
 * (`CHUNK_CELLS`), the paint path at most 256 cells of one row per call (`CF_PAINT_BATCH_CELLS`).
 *
 *   npx tsx scripts/cf-refs/active-units.ts [runs]
 */
import { readFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { createCfService } from '../../src/xlsx/cf/service'
import { createCfMemory, createCfWorkBudget, CF_LIMITS, type CfWorkBudget } from '../../src/xlsx/cf/limits'
import { residentSource } from '../../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../../src/xlsx/cf/predicates'
import { cfRulesVersion } from '../../src/xlsx/view/pagedSource'
import { barRule, scaleRule } from './audit'
import type { XlsxSheet } from '../../src/xlsx/types'

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..')
type Rect = { firstRow: number; lastRow: number; firstCol: number; lastCol: number }

/** Counts units exactly as createCfWorkBudget does, and active time (elapsed between resume and the next yield, idle excluded) without ever refusing. */
function countingBudget(): CfWorkBudget & { activeMs(): number } {
  const now = () => performance.now()
  let start = now(), active = 0, units = 0, slice = 0, paused = false
  return {
    get units() { return units },
    spend(n) {
      units += n // a spend made while paused counts, as in createCfWorkBudget
      if (paused) return 'yield'
      slice += n
      const elapsed = Math.max(0, now() - start)
      if (slice >= CF_LIMITS.checkpointUnits || elapsed >= CF_LIMITS.sliceMs) { active += elapsed; paused = true; return 'yield' }
      return 'continue'
    },
    resume() { start = now(); slice = 0; paused = false },
    activeMs: () => active + (paused ? 0 : Math.max(0, now() - start)),
  }
}

type Source = ReturnType<typeof residentSource>
async function prepare(sheet: XlsxSheet, rect: Rect, budget: () => CfWorkBudget, shared?: Source) {
  const memory = createCfMemory()
  // The paint path's statistics budget (render.ts): a whole-scan budget, not the per-call one.
  const service = createCfService(sheet.conditionalFormatting!, memory, { evidence: cfRuntimeEvidence,
    statisticsBudget: () => createCfWorkBudget({ scan: true, maxUnits: 10_000_000 }), budget })
  const stamp = { owner: 'active-units', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(sheet.conditionalFormatting!), viewGen: 0, structureGen: 0, clockGen: 0 }
  const source = shared ?? residentSource(sheet, stamp, memory, 'engine')
  let outcome = 'ready', bars = 0
  try {
    const lease = await service.prepareCf({ stamp, sourceRects: [rect], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, source)
    for (let r = rect.firstRow; r <= rect.lastRow; r++) for (let c = rect.firstCol; c <= rect.lastCol; c++) {
      const o = lease.value.at(r, c)
      if (o.kind === 'ready' && o.value.visual?.dataBar) bars++
    }
    lease.release()
  } catch (e) { outcome = `refused (${(e as { code?: string }).code ?? (e as Error).message})` }
  service.dispose()
  return { outcome, bars }
}

async function fixture(id: string) {
  const m = JSON.parse(readFileSync(join(ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8'))
  const c = m.cases.find((x: { id: string }) => x.id === id)
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(ROOT, c.saved)))))
  const sheet = (doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!) as unknown as XlsxSheet
  const [a, b] = c.range.split(':')
  const rc = (ref: string) => { const mm = /^([A-Z]+)(\d+)$/.exec(ref)!; return { col: [...mm[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1, row: Number(mm[2]) - 1 } }
  const s = rc(a), e = rc(b)
  return { sheet, rect: { firstRow: s.row, lastRow: e.row, firstCol: s.col, lastCol: e.col }, cells: (e.row - s.row + 1) * (e.col - s.col + 1) }
}

const populated = (n: number, k: number) => async () => {
  const rows = Array.from({ length: n }, (_, r) => ({ index: r, cells: [{ ref: `A${r + 1}`, row: r, col: 0, styleIndex: 0, value: (r % 2) * 10 + 5, hasCachedValue: true }] }))
  const sheet = { name: 'P', sheetId: 'p', rows, cols: [], mergeRefs: [], mergeRanges: [], conditionalFormatting: { diagnostics: [], rules: [barRule(0, { firstRow: 0, lastRow: n - 1, firstCol: 0, lastCol: 0 })] } } as unknown as XlsxSheet
  return { sheet, rect: { firstRow: 0, lastRow: k - 1, firstCol: 0, lastCol: 0 }, cells: k }
}

const scenarios: Array<{ name: string; build: () => Promise<{ sheet: XlsxSheet; rect: Rect; cells: number }> }> = [
  { name: 'stage3-bar-lengths-default A1:V28 (largest fixture)', build: () => fixture('stage3-bar-lengths-default') },
  { name: 'stage3-bar-auto-bounds', build: () => fixture('stage3-bar-auto-bounds') },
  { name: 'sparse 1M-row declared A:A, 100 bars + 100 scales, 40-cell request', build: async () => {
    const whole = { firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0 }
    const rules = [...Array.from({ length: 100 }, (_, i) => barRule(i, whole)), ...Array.from({ length: 100 }, (_, i) => scaleRule(100 + i, whole))]
    const sheet = { name: 'P', sheetId: 'p', rows: [{ index: 0, cells: Array.from({ length: 100 }, (_, row) => ({ ref: `A${row + 1}`, row, col: 0, styleIndex: 0, value: (row % 2) * 10 + 5, hasCachedValue: true })) }], cols: [], mergeRefs: [], mergeRanges: [], conditionalFormatting: { diagnostics: [], rules } } as unknown as XlsxSheet
    return { sheet, rect: { firstRow: 0, lastRow: 39, firstCol: 0, lastCol: 0 }, cells: 40 }
  } },
  { name: 'populated 100,000-row A:A, 1 bar, 1,000-cell call (text-path maximum)', build: populated(100_000, 1000) },
  { name: 'populated 100,000-row A:A, 1 bar, 256-cell call (paint-path maximum)', build: populated(100_000, 256) },
  { name: 'populated 100,000-row A:A, 1 bar, 5,000-cell call (not a production size)', build: populated(100_000, 5000) },
]

const runs = Math.max(1, Number(process.argv[2] ?? 5))
let worstUnits = 0
for (const s of scenarios) {
  const { sheet, rect, cells } = await s.build()
  const samples: Array<{ units: number; ms: number; outcome: string; bars: number }> = []
  for (let i = 0; i < runs; i++) {
    const b = countingBudget()
    const r = await prepare(sheet, rect, () => b)
    samples.push({ units: b.units, ms: b.activeMs(), ...r })
  }
  // Wall time of the per-call work alone: a source whose bounds are already cached, so the separately budgeted
  // whole-sheet bounds scan (its own CF_LIMITS.workUnits budget) does not land between two per-call spends.
  const stamp = { owner: 'active-units', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(sheet.conditionalFormatting!), viewGen: 0, structureGen: 0, clockGen: 0 }
  const warm = residentSource(sheet, stamp, createCfMemory(), 'engine')
  await prepare(sheet, rect, () => countingBudget(), warm)
  const warmMs: number[] = []
  for (let i = 0; i < runs; i++) {
    const b = countingBudget()
    await prepare(sheet, rect, () => b, warm)
    if (b.units !== samples[0]!.units) throw new Error(`${s.name}: warm unit count differs`)
    warmMs.push(b.activeMs())
  }
  warmMs.sort((a, b) => a - b)
  const warmMedian = warmMs[Math.floor(warmMs.length / 2)]!
  const units = samples[0]!.units
  if (samples.some(x => x.units !== units)) throw new Error(`${s.name}: unit count is not deterministic`)
  const ms = samples.map(x => x.ms).sort((a, b) => a - b)
  const median = ms[Math.floor(ms.length / 2)]!
  const real = await prepare(sheet, rect, () => createCfWorkBudget())
  if (!s.name.includes('not a production')) worstUnits = Math.max(worstUnits, units)
  console.log(`${s.name}: cells=${cells} units=${units} activeMs median=${median.toFixed(2)} max=${ms.at(-1)!.toFixed(2)} warm-bounds activeMs median=${warmMedian.toFixed(2)} units/ms(warm median)=${(units / warmMedian).toFixed(0)} cap=${CF_LIMITS.activeUnits} margin=${(CF_LIMITS.activeUnits / units).toFixed(1)}x default budget: ${real.outcome}, bars=${real.bars}`)
}
console.log(`worst production call: ${worstUnits} units; activeUnits=${CF_LIMITS.activeUnits} (${(CF_LIMITS.activeUnits / worstUnits).toFixed(1)}x)`)
