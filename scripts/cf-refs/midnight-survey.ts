/**
 * The across-midnight follow-up (brief cf-midnight, Phase 1 task 1): the survey of every path that can paint or cache CF
 * once the local day differs from the day a sheet was loaded or evaluated on, and of every case cell that crossing midnight
 * reaches.
 *
 * Paths: each entry names a file and a source needle; the survey resolves the line it is on today (so a moved line never
 * goes stale) and fails if the needle is gone.
 *
 * Cells, per case with a workbook: the document is parsed as the harness parses it (`caseParseOptionsOf`), then
 * - `clockCells`: the range cells covered by a clock-reading rule (`timePeriod`, or a formula calling TODAY()/NOW());
 * - `taintedCells`: the cells whose value may move with the clock (the load record's taint, `formula/loadClock.ts`);
 * - `readerCells`: the range cells of a rule whose scope or formula references hold a tainted cell (rule-level, an
 *   over-approximation computed here from the scope and the formulas' reference atoms, independent of the service);
 * - `withheldNextDay`: the cells at which production's service withholds a rule across midnight one day after the load,
 *   unregistered (the withhold), and `changedNextDay` the cells whose prepared state differs between the two days.
 * The checks (pinned by tests/xlsx-cf-midnight-survey.test.ts): every clock cell is withheld the next day, every withheld
 * cell is a clock cell or a reader cell, and no cell outside both changes.
 *
 * Run: `npx tsx scripts/cf-refs/midnight-survey.ts` (prints the path table and the per-case counts).
 */
import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { documentLoadClockOf } from '../../src/xlsx/formula/loadClock'
import { createCfService } from '../../src/xlsx/cf/service'
import { cfWorkbookOfDocument, residentSource } from '../../src/xlsx/cf/render'
import { cfAlignClock, cfEvaluationContextFor } from '../../src/xlsx/cf/formula'
import { CF_LIMITS, createCfMemory } from '../../src/xlsx/cf/limits'
import { cfRuntimeEvidence } from '../../src/xlsx/cf/predicates'
import { CF_ACROSS_MIDNIGHT_REASON, cfResidentMidnight } from '../../src/xlsx/cf/midnight'
import { cfDefinitionFormulas, cfRulesUseClock, type CfRule, type SourceRect } from '../../src/xlsx/cf/model'
import { caseDocumentEvaluationOf, caseParseOptionsOf } from './case-clock'
import type { CfManifest } from './cases'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const DAY = 86_400_000

export interface MidnightPath { readonly id: string; readonly file: string; readonly needle: string; readonly role: string }
/** Every path that paints, caches or recalculates across midnight, and the seam it holds. */
export const MIDNIGHT_PATHS: readonly MidnightPath[] = [
  { id: 'resident-tick', file: 'src/components/xlsx/InteractiveSheetDoc.tsx', needle: 'setClockGen(localDaySerial())', role: 'the resident viewer\'s local-midnight tick (advances clockGen; re-keys cfLiveKey and the remote stamps)' },
  { id: 'tick-reach', file: 'src/components/xlsx/InteractiveSheetDoc.tsx', needle: 'const clockFree = cfRules ?', role: 'which sheets the tick reaches: clock-reading rules, or clock-free rules over clock-tainted cells' },
  { id: 'remote-stamp', file: 'src/components/xlsx/InteractiveSheetDoc.tsx', needle: 'structureGen, clockGen: tickGen,', role: 'remote text/find/a11y preparations stamp the applied day' },
  { id: 'pivot-recompute', file: 'src/components/xlsx/InteractiveSheetDoc.tsx', needle: 'recomputeDependents(doc, changed', role: 'a post-load recalculation (pivot commit) that can re-read the clock' },
  { id: 'recompute-taint', file: 'src/xlsx/formula/workbook.ts', needle: 'if (onlyKeys && volatileKeys.size > 0) markDocumentClockUnknown(doc)', role: 'a recalculation after load that read the clock taints every cell (fail closed)' },
  { id: 'load-record', file: 'src/xlsx/parse.ts', needle: 'recordDocumentLoadClock(doc,', role: 'the parse records its one clock sample, zone and the clock taint' },
  { id: 'resident-paint', file: 'src/xlsx/cf/render.ts', needle: 'const midnight = cfResidentMidnight(sheet)', role: 'prepareCfPaint: static paint (render/paint.ts), sheet assets and print (sheetAssets.ts), the viewport paint' },
  { id: 'static-paint', file: 'src/render/paint.ts', needle: 'cf.prepareCfPaint(sheet,', role: 'export / static paint (through prepareCfPaint)' },
  { id: 'assets-print', file: 'src/components/xlsx/sheetAssets.ts', needle: 'cf?.prepareCfPaint(sheet,', role: 'interactive assets and print units (through prepareCfPaint)' },
  { id: 'viewport-paint', file: 'src/components/xlsx/InteractiveSheetDoc.tsx', needle: 'cf?.prepareCfPaint(sheet,{store:source.store(activeSheetId!)', role: 'the scrolled viewport paint (through prepareCfPaint)' },
  { id: 'resident-text', file: 'src/xlsx/cf/text.ts', needle: 'const midnight = cfResidentMidnight(sheet)', role: 'prepareCfTextScope: copy, find, accessibility on resident sheets' },
  { id: 'paged-clock', file: 'src/xlsx/view/pagedSource.ts', needle: 'let liveClockGen: number | undefined', role: 'the paged route\'s own live clock (one sample per preparation)' },
  { id: 'paged-context', file: 'src/xlsx/view/pagedSource.ts', needle: 'midnightWorkbook: this.workbook,', role: 'the paged route\'s across-midnight context (first judged clock, clockFormulas taint)' },
  { id: 'paged-stale-day', file: 'src/xlsx/view/pagedSource.ts', needle: 'Conditional formatting request names a different day', role: 'a request naming another day than the sampled one is refused (cf-stale-generation)' },
  { id: 'paged-clock-free', file: 'src/xlsx/view/pagedSource.ts', needle: 'liveRules?.clockFree === true && !pagedWorkbookClockTainted(this.workbook)', role: 'the worker reply claims clock-free only for an untainted workbook' },
  { id: 'paged-service', file: 'src/xlsx/view/pagedSource.ts', needle: 'midnight: serviceModule.cfPagedMidnight(options.midnightWorkbook)', role: 'prepareCfRemote hands the context to the service' },
  { id: 'worker-cache', file: 'src/worker/viewerRemoteCf.ts', needle: '|${stamp.structureGen}|${stamp.clockGen}`', role: 'the worker CF cache key names the day (a new day is a new entry)' },
  { id: 'worker-routes', file: 'src/worker/viewerSession.ts', needle: 'structureGen: 0, clockGen: 0 },', role: 'worker-side routes (find, copy, export) stamp clockGen 0 and adopt the live day' },
  { id: 'service-gate', file: 'src/xlsx/cf/service.ts', needle: 'const midnight = options.midnight', role: 'the gate: withhold, or the registered reading' },
  { id: 'harness-oracle', file: 'scripts/cf-compare.ts', needle: 'const oracle = createCfService(sheet.conditionalFormatting, memory,', role: 'the cf-compare oracle: no context, the case clock pinned (same day by construction)' },
  { id: 'harness-pins', file: 'scripts/cf-refs/case-clock.ts', needle: 'export function caseParseOptionsOf', role: 'validation pins parse and rules to one instant, so the load day is the evaluation day' },
]
export function resolveMidnightPaths(root = REPO_ROOT): Array<MidnightPath & { line: number }> {
  return MIDNIGHT_PATHS.map(p => {
    const lines = readFileSync(resolve(root, p.file), 'utf8').split('\n')
    const line = lines.findIndex(l => l.includes(p.needle)) + 1
    if (line === 0) throw new Error(`midnight survey: ${p.id}: needle not found in ${p.file}: ${p.needle}`)
    return { ...p, line }
  })
}

export interface MidnightCaseSurvey {
  readonly id: string
  readonly clockCells: readonly string[]
  readonly taintedCells: readonly string[] | 'all'
  readonly readerCells: readonly string[]
  readonly withheldNextDay: readonly string[]
  readonly changedNextDay: readonly string[]
}
const refOf = (row: number, col: number): string => {
  let c = col + 1, s = ''
  while (c > 0) { const r = (c - 1) % 26; s = String.fromCharCode(65 + r) + s; c = Math.floor((c - 1) / 26) }
  return `${s}${row + 1}`
}
const rangeRect = (a1: string): SourceRect => {
  const [from, to = from] = a1.split(':')
  const at = (r: string) => { const m = /^([A-Z]+)(\d+)$/.exec(r)!; return { row: Number(m[2]) - 1, col: [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 } }
  const a = at(from!), b = at(to!)
  return { firstRow: a.row, lastRow: b.row, firstCol: a.col, lastCol: b.col }
}
const inside = (r: SourceRect, row: number, col: number): boolean => row >= r.firstRow && row <= r.lastRow && col >= r.firstCol && col <= r.lastCol
const overlaps = (a: SourceRect, b: SourceRect): boolean => a.firstRow <= b.lastRow && b.firstRow <= a.lastRow && a.firstCol <= b.lastCol && b.firstCol <= a.lastCol
/** A rule's read set, over-approximated: its scope, and each formula atom spread over the scope's extent (relative axes shift). */
function readRects(rule: CfRule): Array<{ sheet?: string; rect: SourceRect }> {
  const out: Array<{ sheet?: string; rect: SourceRect }> = rule.ranges.map(rect => ({ rect }))
  const rows = Math.max(...rule.ranges.map(r => r.lastRow)) - Math.min(...rule.ranges.map(r => r.firstRow))
  const cols = Math.max(...rule.ranges.map(r => r.lastCol)) - Math.min(...rule.ranges.map(r => r.firstCol))
  for (const f of cfDefinitionFormulas(rule.definition)) for (const a of f.atoms) {
    out.push({ ...(a.sheet !== undefined ? { sheet: a.sheet } : {}), rect: { firstRow: Math.max(0, a.firstRow - (a.absRow ? 0 : rows)), lastRow: a.lastRow + (a.absRow ? 0 : rows), firstCol: Math.max(0, a.firstCol - (a.absCol ? 0 : cols)), lastCol: a.lastCol + (a.absCol ? 0 : cols) } })
  }
  return out
}

/** The per-case survey. Only cases with a workbook on disk are surveyed. */
export async function midnightCaseSurveyOf(manifestPath = resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), root = REPO_ROOT, ids?: readonly string[]): Promise<MidnightCaseSurvey[]> {
  const manifest = JSON.parse(readFileSync(manifestPath, 'utf8')) as CfManifest
  const out: MidnightCaseSurvey[] = []
  for (const c of manifest.cases) {
    if (ids && !ids.includes(c.id)) continue
    if (c.status === 'unsupported') continue
    let bytes: Uint8Array
    try { bytes = new Uint8Array(readFileSync(resolve(root, c.saved))) } catch { continue }
    const doc = await parseXlsx(await OfficePackage.load(bytes), caseParseOptionsOf(c))
    const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
    const cf = sheet.conditionalFormatting
    if (!cf?.rules.length) continue
    const record = documentLoadClockOf(doc)!
    const range = rangeRect(c.range)
    const cells: Array<{ row: number; col: number }> = []
    for (let row = range.firstRow; row <= range.lastRow; row++) for (let col = range.firstCol; col <= range.lastCol; col++) cells.push({ row, col })
    const taintHas = (sheetName: string | undefined, rect: SourceRect): boolean => {
      if (record.taint === 'all') return true
      const set = record.taint.get((sheetName ?? sheet.name).toLowerCase())
      if (!set) return false
      for (const id of set) if (inside(rect, Math.floor(id / 16384), id % 16384)) return true
      return false
    }
    const clock = new Set<string>(), readers = new Set<string>()
    for (const rule of cf.rules) {
      const covered = cells.filter(x => rule.ranges.some(r => inside(r, x.row, x.col))).map(x => refOf(x.row, x.col))
      if (cfRulesUseClock({ rules: [rule] })) for (const r of covered) clock.add(r)
      if (readRects(rule).some(e => taintHas(e.sheet, e.rect))) for (const r of covered) readers.add(r)
    }
    const policy = caseDocumentEvaluationOf(c, doc)
    const memory = createCfMemory()
    const workbook = cfWorkbookOfDocument(doc.sheets, memory, 'engine')
    /** Production's service at `instant` with the sheet's across-midnight context: each cell's state and whether a rule
     * covering it was withheld across midnight (read through the validation seam `cellOutcomes`). */
    const at = async (instant: number) => {
      const { evaluation, dateContext } = cfAlignClock(cfEvaluationContextFor(policy, () => instant), undefined, () => undefined)
      const withheld = new Set<string>()
      const midnight = cfResidentMidnight(sheet)
      const service = createCfService(cf, memory, { evidence: cfRuntimeEvidence, ...(evaluation ? { evaluation } : {}), ...(dateContext ? { dateContext } : {}), sheetName: sheet.name, workbook,
        ...(midnight ? { midnight } : {}),
        cellOutcomes: (row, col, _rules, outcomes) => { if (outcomes.some(o => o.kind === 'unresolved' && o.reason === CF_ACROSS_MIDNIGHT_REASON)) withheld.add(refOf(row, col)) } })
      const stamp = { owner: 'midnight-survey', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
      const source = residentSource(sheet, stamp, memory, 'engine')
      const m = new Map<string, { withheld: boolean; style: string }>()
      try {
        // One row at a time, or one cell when a row exceeds the admitted batch; a cell production refuses even alone is
        // recorded as refused (the same on both days).
        const prepare = async (rect: SourceRect): Promise<boolean> => {
          let lease
          try { lease = await service.prepareCf({ stamp, sourceRects: [rect], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, source) } catch { return false }
          try {
            for (let row = rect.firstRow; row <= rect.lastRow; row++) for (let col = rect.firstCol; col <= rect.lastCol; col++) {
              const ref = refOf(row, col)
              m.set(ref, { withheld: withheld.has(ref), style: JSON.stringify(lease.value.at(row, col)) })
            }
          } finally { lease.release() }
          return true
        }
        for (let row = range.firstRow; row <= range.lastRow; row++) {
          if (await prepare({ firstRow: row, lastRow: row, firstCol: range.firstCol, lastCol: range.lastCol })) continue
          for (let col = range.firstCol; col <= range.lastCol; col++) {
            if (!(await prepare({ firstRow: row, lastRow: row, firstCol: col, lastCol: col }))) m.set(refOf(row, col), { withheld: false, style: 'refused' })
          }
        }
      } finally { service.dispose() }
      return m
    }
    const day0 = await at(record.epochMs), day1 = await at(record.epochMs + DAY)
    const sortRefs = (refs: Iterable<string>): string[] => [...refs].sort((a, b) => (Number(/\d+/.exec(a)![0]) - Number(/\d+/.exec(b)![0])) || a.localeCompare(b))
    out.push({
      id: c.id,
      clockCells: sortRefs(clock),
      taintedCells: record.taint === 'all' ? 'all' : sortRefs([...(record.taint.get(sheet.name.toLowerCase()) ?? [])].map(id => refOf(Math.floor(id / 16384), id % 16384))),
      readerCells: sortRefs(readers),
      withheldNextDay: sortRefs([...day1].filter(([r, v]) => v.withheld && !day0.get(r)!.withheld).map(([r]) => r)),
      changedNextDay: sortRefs([...day1].filter(([r, v]) => v.style !== day0.get(r)!.style).map(([r]) => r)),
    })
  }
  return out
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  for (const p of resolveMidnightPaths()) console.log(`${p.id.padEnd(18)} ${p.file}:${p.line}  ${p.role}`)
  const cases = await midnightCaseSurveyOf()
  for (const s of cases) if (s.clockCells.length || s.readerCells.length || s.withheldNextDay.length || s.changedNextDay.length || s.taintedCells === 'all' || s.taintedCells.length) {
    console.log(`${s.id}: clock ${s.clockCells.length}, tainted ${s.taintedCells === 'all' ? 'all' : s.taintedCells.length}, readers ${s.readerCells.length}, withheld next day ${s.withheldNextDay.length}, changed next day ${s.changedNextDay.length}`)
  }
  console.log(`${cases.length} cases surveyed`)
}
