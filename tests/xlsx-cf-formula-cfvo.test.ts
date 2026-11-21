// @vitest-environment node
/**
 * Task 37 (Rulings S5-H, S5-I, S5-J): formula cfvo bounds on colour scales, data bars and icon sets, the standalone x14
 * colour-scale arm and `activePresent`. Every expectation is a literal twin rule production itself resolves, or a stated
 * number; nothing here records or predicts an Excel observation. S5-I and S5-J are tested under each outcome the Task 31
 * capture can record (injected), and once under the outcome actually recorded, which today is "unrecorded".
 */
import { describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { parseXml } from '../src/core/xml'
import { OfficePackage } from '../src/core/zip'
import { tokenizeBytes } from '../src/core/xml-stream'
import { parseXlsx } from '../src/xlsx/parse'
import * as formulaModule from '../src/xlsx/cf/formula'
import { cfDocumentEvaluationOf, cfEvaluationContextFor } from '../src/xlsx/cf/formula'
import { cfWorkbookOfDocument, residentSource } from '../src/xlsx/cf/render'
import { cfvoBoundAnchors, createCfService } from '../src/xlsx/cf/service'
import { createCfStreamEvents } from '../src/xlsx/cf/capture'
import { CF_LIMITS, createCfMemory, createCfWorkBudget } from '../src/xlsx/cf/limits'
import { VERIFIED_CAPABILITIES, injectCfFormulaSelections, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { CF_FORMULA_PRODUCTION_READINGS, cfFormulaReading } from '../src/xlsx/cf/formulaReadings'
import { MEASURED_CF_OFFICE_THEME, parseCfSheet, type CfStylesSource } from '../src/xlsx/cf/parse'
import { cfRulesVersion } from '../src/xlsx/view/pagedSource'
import { CF_STAGE5_CAPABILITIES, CF_STAGE5_UNSUPPORTED, type CfManifest } from '../scripts/cf-refs/cases'
import { auditStage5Manifest, stage5Publication, type Stage5AuditReport } from '../scripts/cf-refs/audit'
import { resolveScaleThresholds } from '../src/xlsx/cf/colorScale'
import { resolveBarThresholds } from '../src/xlsx/cf/dataBar'
import { resolveIconThresholds } from '../src/xlsx/cf/iconSet'
import { buildZip } from './helpers/zip-builder'
import { createPagedWorkbookSource, type PagedXlsxSheetAdapter } from '../src/xlsx/view/pagedSource'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import type { CfReadResult } from '../src/xlsx/view/contracts'
import type { CfCellState, CfDefinition, CfOutcome, CfSheetRules, CfStamp } from '../src/xlsx/cf/model'
import type { CfWorkbookSheet, CfWorkbookSource } from '../src/xlsx/cf/dependencies'
import type { XlsxSheet } from '../src/xlsx/types'
import { GREEN, UTC, rc, sheetOf, shown, stylesFor } from './helpers/cf-formula-rules'
import { EVALUATION, fakeSource, formulaOf, formulaRule } from './helpers/cf-formula'
import { prepareCfFormulaBatch } from '../src/xlsx/cf/dependencies'
import { createCfStatistics } from '../src/xlsx/cf/statistics'

// A pass-through spy on the one evaluator (Tasks 34-36): every formula call goes through `evaluateCfFormula`.
vi.mock('../src/xlsx/cf/formula', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/cf/formula')>()
  return { ...actual, evaluateCfFormula: vi.fn(actual.evaluateCfFormula) }
})
const evaluateSpy = formulaModule.evaluateCfFormula as unknown as ReturnType<typeof vi.fn>

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const XM = 'http://schemas.microsoft.com/office/excel/2006/main'
const CF_EXT = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}'
const DB_EXT = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}'
const styles: CfStylesSource = { colors: { theme: MEASURED_CF_OFFICE_THEME }, dxfNodes: [] }

/** The formula capabilities a same-sheet cfvo bound of these tests needs (with production's readings registered). */
const CFVO_CAPS: readonly string[] = [
  'formula:dialect-base', 'formula:dialect-x14', 'formula:anchor-single', 'formula:anchor-mixed-abs',
  'formula:fn-aggregate', 'formula:fn-math', 'formula:cfvo-reference', 'formula:cfvo-expression',
  'scale:thresholds-formula', 'bar:bounds-formula', 'icon:thresholds-formula',
  // The literal num twins of the bar tests need it (CF_STAGE3_NONDISCRIMINATING); a formula-resolved bar threshold does not
  // (fix round 1, ruling 2: it is authorised through bar:bounds-formula).
  'bar:bounds-num',
]
/** Production's evidence plus `extra`: the exact set a test runs under (parse and service alike). */
const capsWith = (...extra: readonly string[]): string[] => [...VERIFIED_CAPABILITIES, ...CFVO_CAPS, ...extra]

const worksheet = (inner: string): string => `<worksheet xmlns="${NS}" xmlns:x14="${X14}" xmlns:xm="${XM}">${inner}</worksheet>`
function parseRules(inner: string, caps: Iterable<string>): CfSheetRules {
  return withExactCfEvidence(caps, () => parseCfSheet(parseXml(worksheet(inner)), styles, createCfMemory(), { sheetRightToLeft: false }).value)
}
async function streamRules(inner: string, caps: Iterable<string>): Promise<CfSheetRules> {
  return withExactCfEvidence(caps, async () => {
    const memory = createCfMemory(), handle = createCfStreamEvents(styles, memory, { sheetRightToLeft: false })
    const bytes = new TextEncoder().encode(worksheet(inner))
    await tokenizeBytes((async function* () { yield bytes })(), handle.events)
    return handle.finish().value
  })
}
const COLORS = '<color rgb="FFF8696B"/><color rgb="FF63BE7B"/>'
const scale = (sqref: string, priority: number, cfvos: string): string =>
  `<conditionalFormatting sqref="${sqref}"><cfRule type="colorScale" priority="${priority}"><colorScale>${cfvos}${COLORS}</colorScale></cfRule></conditionalFormatting>`
const icon = (sqref: string, priority: number, cfvos: string): string =>
  `<conditionalFormatting sqref="${sqref}"><cfRule type="iconSet" priority="${priority}"><iconSet iconSet="3Arrows">${cfvos}</iconSet></cfRule></conditionalFormatting>`
/** A paired bar (Excel's form, the stage5-cfvo-formulas shape): base cfvo plus an x14 companion carrying its own cfvo. */
const pairedBar = (sqref: string, priority: number, guid: string, base: string, x14: string) => ({
  base: `<conditionalFormatting sqref="${sqref}"><cfRule type="dataBar" priority="${priority}"><dataBar>${base}<color rgb="FF0000FF"/></dataBar>`
    + `<extLst><ext uri="${DB_EXT}"><x14:id>${guid}</x14:id></ext></extLst></cfRule></conditionalFormatting>`,
  x14: `<x14:conditionalFormatting><x14:cfRule type="dataBar" id="${guid}"><x14:dataBar minLength="0" maxLength="100" gradient="0">${x14}`
    + `<x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting>`,
})
const ext = (rules: string): string => `<extLst><ext uri="${CF_EXT}"><x14:conditionalFormattings>${rules}</x14:conditionalFormattings></ext></extLst>`
const col = (letter: string, values: readonly (number | string | { err: string })[], from = 2): Record<string, number | string | { err: string }> =>
  Object.fromEntries(values.map((v, i) => [`${letter}${from + i}`, v]))
const rows = (letter: string, from: number, to: number): string[] => Array.from({ length: to - from + 1 }, (_, i) => `${letter}${from + i}`)
const DATA = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
/** A bar's data stays within its bounds (Ruling S3-E: a cell beyond them withholds the whole rule). */
const BAR_DATA = [0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50]

type Prepared = Map<string, CfOutcome<CfCellState>>
interface PrepareOptions { selections?: Readonly<Record<string, string>>; workbook?: CfWorkbookSource; sheetName?: string; evaluation?: boolean; stamp?: Partial<CfStamp> }
/** Prepare `refs` through the real service (`createCfService` + resident source) under exactly `caps`. */
async function prepareSheet(sheet: XlsxSheet, refs: readonly string[], caps: Iterable<string>, options: PrepareOptions = {}): Promise<{ out: Prepared; at: (ref: string) => CfOutcome<CfCellState>; pending?: readonly string[] }> {
  const set = new Set(caps)
  const memory = createCfMemory()
  const evaluation = options.evaluation === false ? undefined : cfEvaluationContextFor(UTC, () => Date.UTC(2026, 10, 16))
  const service = createCfService(sheet.conditionalFormatting!, memory, { evidence: { allows: cap => set.has(cap) }, ...(evaluation ? { evaluation } : {}), sheetName: options.sheetName ?? sheet.name, ...(options.workbook ? { workbook: options.workbook } : {}) })
  const stamp: CfStamp = { owner: 't', sheetId: sheet.sheetId ?? 'sheet1', workbookValueGen: 0, cfGen: cfRulesVersion(sheet.conditionalFormatting!), viewGen: 0, structureGen: 0, clockGen: 0, ...options.stamp }
  injectCfFormulaSelections(options.selections ?? null)
  try {
    const rects = refs.map(r => { const { row, col: c } = rc(r); return { firstRow: row, lastRow: row, firstCol: c, lastCol: c } })
    const source = residentSource(sheet, stamp, memory, 'engine')
    const lease = await service.prepareCf({ stamp, sourceRects: rects, signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, source)
    const out: Prepared = new Map()
    for (const r of refs) { const { row, col: c } = rc(r); out.set(r, lease.value.at(row, c)) }
    const pending = lease.value.pendingSheetIds
    return { out, at: (ref: string) => { const { row, col: c } = rc(ref); return lease.value.at(row, c) }, ...(pending ? { pending } : {}) }
  } finally {
    injectCfFormulaSelections(null)
  }
}
/** The visual one cell shows, with the rule id dropped (a twin is another rule): or why it shows none. */
function visualOf(o: CfOutcome<CfCellState> | undefined): string {
  if (!o) return 'missing'
  if (o.kind !== 'ready') return `unresolved:${o.kind === 'unresolved' ? o.reason : o.kind}`
  if ((o.value.unresolvedWrites?.size ?? 0) > 0) return 'withheld'
  const v = o.value.visual
  if (!v) return 'none'
  const strip = (x: object | undefined): string | undefined => {
    if (!x) return undefined
    const { ruleId: _ruleId, ...rest } = x as Record<string, unknown>
    void _ruleId
    return JSON.stringify(rest)
  }
  return JSON.stringify({ scale: strip(v.colorScale), bar: strip(v.dataBar), icon: strip(v.iconSet) })
}
const twinsAgree = (out: Prepared, pairs: ReadonlyArray<readonly [string, string]>): void => {
  for (const [a, b] of pairs) {
    expect(visualOf(out.get(a)), `${a} vs twin ${b}`).toBe(visualOf(out.get(b)))
    expect(visualOf(out.get(a)), a).toMatch(/^\{/)
  }
}
/** Withheld by name: an unresolved cell, or a ready cell whose writes are tainted (composition keeps the slot). */
const WITHHELD = /^(unresolved:.*|withheld)$/
const definitionOf = (rules: CfSheetRules, priority: number): CfDefinition => rules.rules.find(r => r.priority === priority)!.definition

/* ------------------------------------------------------------------ */
/* S5-H: formula bounds resolve once per measured scope.                */
/* ------------------------------------------------------------------ */

describe('a formula bound resolves once per measured scope (S5-H, round-1 #10)', () => {
  it('scale formula =$F$1 (F1 = 50) over {0..100} resolves its min threshold to 50, evaluated once per preparation', async () => {
    const xml = scale('B2:B12', 1, '<cfvo type="formula" val="$F$1"/><cfvo type="max"/>') + scale('C2:C12', 2, '<cfvo type="num" val="50"/><cfvo type="max"/>')
    const caps = capsWith()
    const sheet = sheetOf(parseRules(xml, caps), { F1: 50, ...col('B', DATA), ...col('C', DATA) })
    evaluateSpy.mockClear()
    const { out } = await prepareSheet(sheet, [...rows('B', 2, 12), ...rows('C', 2, 12)], caps)
    // One evaluation per preparation, whatever the number of cells the scale covers (S5-H: memoised per rule).
    expect(evaluateSpy).toHaveBeenCalledTimes(1)
    twinsAgree(out, rows('B', 2, 12).map(r => [r, `C${r.slice(1)}`] as const))
    // 80 sits at (80 - 50) / (100 - 50) = 0.6 of the scale.
    const b10 = out.get('B10')!
    expect(b10.kind === 'ready' && b10.value.visual?.colorScale?.fraction).toBeCloseTo(0.6, 12)
  })

  it('percent $F$1 with F1 = 25 resolves to lerp(min, max, 0.25)', async () => {
    const xml = scale('B2:B12', 1, '<cfvo type="percent" val="$F$1"/><cfvo type="max"/>') + scale('C2:C12', 2, '<cfvo type="percent" val="25"/><cfvo type="max"/>')
      + scale('D2:D12', 3, '<cfvo type="num" val="25"/><cfvo type="max"/>')
    const caps = capsWith()
    const sheet = sheetOf(parseRules(xml, caps), { F1: 25, ...col('B', DATA), ...col('C', DATA), ...col('D', DATA) })
    const { out } = await prepareSheet(sheet, [...rows('B', 2, 12), ...rows('C', 2, 12), ...rows('D', 2, 12)], caps)
    // lerp(0, 100, 0.25) = 25: the percent twin and the num-25 twin agree with it.
    twinsAgree(out, rows('B', 2, 12).flatMap(r => [[r, `C${r.slice(1)}`], [r, `D${r.slice(1)}`]] as const))
  })

  it('a bar formula max and an icon formula threshold resolve the same way', async () => {
    const formulaBar = pairedBar('B2:B12', 1, '{00000000-0000-4000-8000-000000000001}', '<cfvo type="min"/><cfvo type="formula" val="$F$1"/>', '<x14:cfvo type="min"/><x14:cfvo type="formula"><xm:f>$F$1</xm:f></x14:cfvo>')
    const twinBar = pairedBar('C2:C12', 2, '{00000000-0000-4000-8000-000000000002}', '<cfvo type="min"/><cfvo type="num" val="50"/>', '<x14:cfvo type="min"/><x14:cfvo type="num"><xm:f>50</xm:f></x14:cfvo>')
    const xml = formulaBar.base + twinBar.base
      + icon('D2:D12', 3, '<cfvo type="percent" val="0"/><cfvo type="formula" val="$F$1"/><cfvo type="num" val="80"/>')
      + icon('E2:E12', 4, '<cfvo type="percent" val="0"/><cfvo type="num" val="50"/><cfvo type="num" val="80"/>')
      + ext(formulaBar.x14 + twinBar.x14)
    const caps = capsWith()
    const sheet = sheetOf(parseRules(xml, caps), { F1: 50, ...col('B', BAR_DATA), ...col('C', BAR_DATA), ...col('D', DATA), ...col('E', DATA) })
    evaluateSpy.mockClear()
    const { out } = await prepareSheet(sheet, [...rows('B', 2, 12), ...rows('C', 2, 12), ...rows('D', 2, 12), ...rows('E', 2, 12)], caps)
    // One evaluation per rule per preparation: the bar's max and the icon's threshold.
    expect(evaluateSpy).toHaveBeenCalledTimes(2)
    twinsAgree(out, rows('B', 2, 12).map(r => [r, `C${r.slice(1)}`] as const))
    twinsAgree(out, rows('D', 2, 12).map(r => [r, `E${r.slice(1)}`] as const))
  })

  it('a rect that excludes the anchor resolves the same thresholds as one that includes it', async () => {
    // =B3 written on B2 is read from A1 (anchor-a1, the registered reading of formula:cfvo-relative): A2 = 7.
    const xml = scale('B2:B12', 1, '<cfvo type="formula" val="B3"/><cfvo type="max"/>') + scale('C2:C12', 2, '<cfvo type="num" val="7"/><cfvo type="max"/>')
    const caps = capsWith('formula:cfvo-relative')
    const sheet = sheetOf(parseRules(xml, caps), { A2: 7, A3: 99, ...col('B', DATA), ...col('C', DATA) })
    const far = await prepareSheet(sheet, ['B10', 'C10'], caps)
    const near = await prepareSheet(sheet, ['B2', 'B10', 'C10'], caps)
    expect(visualOf(far.out.get('B10'))).toBe(visualOf(far.out.get('C10')))
    expect(visualOf(near.out.get('B10'))).toBe(visualOf(far.out.get('B10')))
  })

  describe('a multi-range rule with any formula bound', () => {
    const scopeCaps = capsWith('formula:cfvo-scope-multi-range', 'formula:cfvo-relative', 'scale:scope-disjoint')
    const designated = [...rows('B', 2, 5), ...rows('D', 2, 5), ...rows('F', 2, 5)]
    const twinsOf = (perRange: boolean) => (r: string): string => {
      const map: Record<string, string> = perRange ? { B: 'P', D: 'S', F: 'T' } : { B: 'P', D: 'Q', F: 'R' }
      return `${map[r[0]!]!}${r.slice(1)}`
    }
    const scaleTwins = ['P', 'Q', 'R', 'S', 'T'].flatMap(c => rows(c, 2, 5))
    it('is unresolved without formula:cfvo-scope-multi-range', async () => {
      // Published since stage5-cfvo-relative-a1 (anchor-a1): removed here to test its gate.
      const caps = capsWith('formula:cfvo-relative', 'scale:scope-disjoint').filter(c => c !== 'formula:cfvo-scope-multi-range')
      const doc = await loadFixture('stage5-cfvo-scope', caps)
      expect(doc.sheets[0]!.conditionalFormatting!.rules.find(r => r.priority === 1)!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
      // Parsed with it, but the service lacks it: still unresolved.
      const parsed = await loadFixture('stage5-cfvo-scope', scopeCaps)
      const { out } = await prepareSheet(parsed.sheets[0]!, designated, caps, { selections: { 'formula:cfvo-scope-multi-range': 'rule-wide' } })
      for (const r of designated) expect(visualOf(out.get(r)), r).toMatch(WITHHELD)
    })
    for (const reading of ['rule-wide', 'per-range'] as const) {
      it(`under ${reading}, memoisation is keyed ${reading === 'rule-wide' ? 'rule-wide' : 'per range'} and the disjoint fixture's cells follow it`, async () => {
        const doc = await loadFixture('stage5-cfvo-scope', scopeCaps)
        const sheet = doc.sheets[0]!
        evaluateSpy.mockClear()
        const scaleOnly = { ...sheet, conditionalFormatting: { ...sheet.conditionalFormatting!, rules: sheet.conditionalFormatting!.rules.filter(r => r.priority === 1 || r.priority >= 4 && r.definition.kind === 'colorScale') } }
        const { out } = await prepareSheet(scaleOnly, [...designated, ...scaleTwins], scopeCaps, { selections: { 'formula:cfvo-scope-multi-range': reading } })
        // Rule-wide: one value at B2. Per-range: one value per range, at each range's own anchor.
        expect(evaluateSpy).toHaveBeenCalledTimes(reading === 'rule-wide' ? 1 : 3)
        twinsAgree(out, designated.map(r => [r, twinsOf(reading === 'per-range')(r)] as const))
      })
    }

    // The cfvo-scope follow-ups: Excel reads a relative bound written on the first range's top-left as if evaluated at A1,
    // so =MAX(B2:B7) reads A1:A6 (stage5-cfvo-scope-v2, 5dc2fe9e). Production implements it (cfvoBoundAnchors) and registers
    // it since the confirmation case stage5-cfvo-relative-a1 (f9d5be94) matched it: the expectations are literal twins.
    describe('anchor-a1 (registered: stage5-cfvo-relative-a1)', () => {
      const v2 = (JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest).cases.find(c => c.id === 'stage5-cfvo-scope-v2')!
      it('cfvoBoundAnchors: anchor-a1 is the one scope at A1; rule-wide and per-range are unchanged', () => {
        const ranges = [{ firstRow: 1, lastRow: 6, firstCol: 1, lastCol: 1 }, { firstRow: 1, lastRow: 6, firstCol: 3, lastCol: 3 }]
        expect(cfvoBoundAnchors(ranges, 'anchor-a1')).toEqual([{ row: 0, col: 0 }])
        expect(cfvoBoundAnchors(ranges, 'rule-wide')).toEqual([{ row: 1, col: 1 }])
        expect(cfvoBoundAnchors(ranges, undefined)).toEqual([{ row: 1, col: 1 }])
        expect(cfvoBoundAnchors(ranges, 'per-range')).toEqual([{ row: 1, col: 1 }, { row: 1, col: 3 }])
      })
      it('a bound task targeted at A1 reads the stored window shifted from the first range\'s top-left; off the sheet it needs an edge reading', async () => {
        const cells = { A1: -35, A2: 15, A3: -30, A4: 5, A5: -20, A6: 10, ...col('B', [-40, -25, -5, 5, 20, 35]), ...col('D', [-40, -15, 2, 10, 30, 75]) } as Record<string, number>
        const at = async (text: string, edge: 'ref-error' | undefined) => {
          const rule = formulaRule(text, 'B2:B7 D2:D7')
          const memory = createCfMemory(), signal = new AbortController().signal, source = fakeSource(cells)
          const batch = await prepareCfFormulaBatch([{ rule, formula: formulaOf(rule), targets: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }] }], {
            sheetName: 'Sheet1', evaluation: EVALUATION, edge, memory, signal, budget: createCfWorkBudget({ signal, scan: true, maxUnits: 10_000_000 }),
            statistics: createCfStatistics(memory), hoist: true, source, sheetId: 's1', stamp: source.stamp('s1'),
          })
          try { return [batch.evaluate(0, 0, 0), batch.evaluate(0, 1, 1)] as const } finally { batch.dispose() }
        }
        // =MAX(B2:B7) written on B2, evaluated at A1: A1:A6, max 15 (not B2:B7, 35). B2 is not a target: no evaluation there.
        const [a1, b2] = await at('MAX(B2:B7)', 'ref-error')
        expect(a1).toEqual({ kind: 'ready', value: { kind: 'number', value: 15 } })
        expect(b2).toEqual({ kind: 'no-match' })
        // =MAX(A2:A7) written on B2 reads the column left of A from A1: unresolved without an edge reading, #REF! under ref-error.
        expect((await at('MAX(A2:A7)', undefined))[0]).toMatchObject({ kind: 'unresolved' })
        expect((await at('MAX(A2:A7)', 'ref-error'))[0]).toMatchObject({ kind: 'ready', value: { kind: 'error' } })
      })
      it('production selects anchor-a1 with its own evidence: the v2 icon set follows its anchor-a1 twins; its multi-range scales paint (scale:scope-disjoint published)', async () => {
        expect(CF_FORMULA_PRODUCTION_READINGS['formula:cfvo-scope-multi-range']).toEqual(['rule-wide', 'per-range', 'anchor-a1'])
        const caps = [...VERIFIED_CAPABILITIES]
        expect(caps.includes('scale:scope-disjoint')).toBe(true)
        expect(cfFormulaReading('formula:cfvo-scope-multi-range', c => caps.includes(c))).toBe('anchor-a1')
        expect(cfFormulaReading('formula:cfvo-relative', c => caps.includes(c))).toBe('anchor-a1')
        const cfgOf = (id: string) => v2.formulaConfigurations!.find(x => x.id === id)!
        const icon = cfgOf('scope2-icon')
        const pairs = icon.refs.map(r => [r, icon.hypotheses['anchor-a1']![r]!.replace('twin:', '')] as const)
        const scaleRefs = cfgOf('scope2-scale').refs
        const doc = await loadFixture('stage5-cfvo-scope-v2', caps)
        const { out } = await prepareSheet(doc.sheets[0]!, [...pairs.flat(), ...scaleRefs], caps)
        twinsAgree(out, pairs)
        // Since stage2-scale-scope-disjoint (union) the multi-range scales paint; they follow their anchor-a1 twins.
        const scale = cfgOf('scope2-scale')
        const { out: scaleOut } = await prepareSheet(doc.sheets[0]!, [...scale.refs, ...scale.refs.map(r => scale.hypotheses['anchor-a1']![r]!.replace('twin:', ''))], caps)
        for (const r of scaleRefs) expect(visualOf(scaleOut.get(r)), r).not.toMatch(WITHHELD)
        twinsAgree(scaleOut, scale.refs.map(r => [r, scale.hypotheses['anchor-a1']![r]!.replace('twin:', '')] as const))
      })
      it('a multi-range icon set with an absolute formula threshold paints like its literal twin under production\'s evidence', async () => {
        const caps = [...VERIFIED_CAPABILITIES]
        const xml = icon('B2:B4 D2:D4', 1, '<cfvo type="percent" val="0"/><cfvo type="formula" val="$F$1"/><cfvo type="num" val="50"/>')
          + icon('H2:H4 J2:J4', 2, '<cfvo type="percent" val="0"/><cfvo type="num" val="20"/><cfvo type="num" val="50"/>')
        const values = [10, 30, 60]
        const sheet = sheetOf(parseRules(xml, caps), { F1: 20, ...col('B', values), ...col('D', values), ...col('H', values), ...col('J', values) })
        const refs = [...rows('B', 2, 4), ...rows('D', 2, 4)]
        const { out } = await prepareSheet(sheet, [...refs, ...rows('H', 2, 4), ...rows('J', 2, 4)], caps)
        twinsAgree(out, refs.map(r => [r, `${r[0] === 'B' ? 'H' : 'J'}${r.slice(1)}`] as const))
      })
    })
  })
})

/* ------------------------------------------------------------------ */
/* Non-numeric, relative and cross-sheet results.                       */
/* ------------------------------------------------------------------ */

describe('non-numeric or relative cfvo results withhold', () => {
  const xml = scale('L2:L12', 1, '<cfvo type="formula" val="$F$3"/><cfvo type="max"/>') + scale('M2:M12', 2, '<cfvo type="formula" val="$F$4"/><cfvo type="max"/>')
    + scale('N2:N12', 3, '<cfvo type="formula" val="$F$5"/><cfvo type="max"/>')
    + scale('Q2:Q12', 4, '<cfvo type="num" val="0"/><cfvo type="max"/>') + scale('R2:R12', 5, '<cfvo type="min"/><cfvo type="max"/>')
  const cells = { F3: 'abc', F5: { err: '#N/A' }, ...Object.assign({}, ...['L', 'M', 'N', 'Q', 'R', 'V'].map(c => col(c, [5, 15, 30, 45, 60, 75, 90, 100]))) }
  const refsOf = (c: string): string[] => rows(c, 2, 9)
  const all = ['L', 'M', 'N', 'Q', 'R', 'V'].flatMap(refsOf)
  it('text, blank or error is unresolved until formula:cfvo-nonnumeric', async () => {
    const caps = capsWith()
    const { out } = await prepareSheet(sheetOf(parseRules(xml, caps), cells), all, caps)
    for (const c of ['L', 'M', 'N']) for (const r of refsOf(c)) expect(visualOf(out.get(r)), r).toMatch(WITHHELD)
    // Selected but unevidenced: still unresolved.
    const selected = await prepareSheet(sheetOf(parseRules(xml, caps), cells), all, caps, { selections: { 'formula:cfvo-nonnumeric': 'zero' } })
    for (const r of refsOf('L')) expect(visualOf(selected.out.get(r)), r).toMatch(WITHHELD)
  })
  it('under the withheld reading the bound is unresolved and keeps its slot taint: a lower matching fill never paints (fix round 1, #1)', async () => {
    const caps = capsWith('formula:cfvo-nonnumeric')
    const greenBelow = scale('L2:L9', 1, '<cfvo type="formula" val="$F$3"/><cfvo type="max"/>')
      + '<conditionalFormatting sqref="L2:L9"><cfRule type="cellIs" dxfId="1" priority="2" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>'
    const rules = withExactCfEvidence(caps, () => parseCfSheet(parseXml(worksheet(greenBelow)), stylesFor(), createCfMemory(), { sheetRightToLeft: false }).value)
    const { out } = await prepareSheet(sheetOf(rules, cells), refsOf('L'), caps, { selections: { 'formula:cfvo-nonnumeric': 'withheld' } })
    for (const r of refsOf('L')) {
      expect(visualOf(out.get(r)), r).toMatch(WITHHELD)
      expect(shown(out.get(r)).fill, r).not.toBe(GREEN)
    }
  })
  for (const [reading, twin] of [['zero', 'Q'], ['skipped', 'R']] as const) {
    it(`under the ${reading} reading each non-numeric bound shows twin ${twin}`, async () => {
      const caps = capsWith('formula:cfvo-nonnumeric')
      const { out } = await prepareSheet(sheetOf(parseRules(xml, caps), cells), all, caps, { selections: { 'formula:cfvo-nonnumeric': reading } })
      for (const c of ['L', 'M', 'N']) for (const r of refsOf(c)) {
        const want = visualOf(out.get(`${twin}${r.slice(1)}`))
        expect(visualOf(out.get(r)), `${r} (${reading})`).toBe(want)
      }
    })
  }
  it('a boolean result is unresolved under every reading', async () => {
    const boolXml = scale('L2:L9', 1, '<cfvo type="formula" val="$F$6"/><cfvo type="max"/>')
    const caps = capsWith('formula:cfvo-nonnumeric')
    for (const reading of ['withheld', 'zero', 'skipped']) {
      const { out } = await prepareSheet(sheetOf(parseRules(boolXml, caps), { F6: true, ...col('L', [1, 2, 3, 4]) }), ['L2', 'L3'], caps, { selections: { 'formula:cfvo-nonnumeric': reading } })
      expect(visualOf(out.get('L3')), reading).toMatch(WITHHELD)
    }
  })
  it('a relative =K3 cfvo is unresolved until formula:cfvo-relative, and then evaluates read from A1 (A2 = 10)', async () => {
    const relXml = scale('K2:K9', 1, '<cfvo type="formula" val="K3"/><cfvo type="max"/>') + scale('U2:U9', 2, '<cfvo type="num" val="10"/><cfvo type="max"/>')
    const relCells = { ...col('A', [10, 15, 30, 45, 60, 75, 90, 100]), ...col('K', [5, 15, 30, 45, 60, 75, 90, 100]), ...col('U', [5, 15, 30, 45, 60, 75, 90, 100]) }
    // Published since stage5-cfvo-relative-a1: removed here to test its gate.
    const without = capsWith().filter(c => c !== 'formula:cfvo-relative')
    const off = await prepareSheet(sheetOf(parseRules(relXml, without), relCells), rows('K', 2, 9), without)
    for (const r of rows('K', 2, 9)) expect(visualOf(off.out.get(r)), r).toMatch(WITHHELD)
    const caps = capsWith('formula:cfvo-relative')
    const { out } = await prepareSheet(sheetOf(parseRules(relXml, caps), relCells), [...rows('K', 2, 9), ...rows('U', 2, 9)], caps)
    twinsAgree(out, rows('K', 2, 9).map(r => [r, `U${r.slice(1)}`] as const))
  })

  describe('a cross-sheet cfvo follows Task 36', () => {
    const crossXml = scale('B2:B12', 1, '<cfvo type="formula" val="Sheet2!$A$1"/><cfvo type="max"/>') + scale('C2:C12', 2, '<cfvo type="num" val="50"/><cfvo type="max"/>')
    const crossCaps = capsWith('formula:cross-sheet-reference')
    const host = (): XlsxSheet => sheetOf(parseRules(crossXml, crossCaps), { ...col('B', DATA), ...col('C', DATA) })
    const other = (value: number): XlsxSheet => ({ ...sheetOf({ rules: [], diagnostics: [] }, { A1: value }), name: 'Sheet2', sheetId: 'sheet2', conditionalFormatting: undefined } as XlsxSheet)
    it('reads the other sheet through the workbook seam and carries its dependency', async () => {
      const sheet = host()
      expect(sheet.conditionalFormatting!.rules.find(r => r.priority === 1)!.definition).toMatchObject({ kind: 'colorScale', dependencySheets: ['Sheet2'] })
      const memory = createCfMemory()
      const { out } = await prepareSheet(sheet, [...rows('B', 2, 12), ...rows('C', 2, 12)], crossCaps, { workbook: cfWorkbookOfDocument([sheet, other(50)], memory, 'engine') })
      twinsAgree(out, rows('B', 2, 12).map(r => [r, `C${r.slice(1)}`] as const))
    })
    it('without the workbook seam it withholds', async () => {
      const { out } = await prepareSheet(host(), ['B10'], crossCaps)
      expect(visualOf(out.get('B10'))).toMatch(WITHHELD)
    })
    it('a pending referenced sheet makes the cells pending, naming it', async () => {
      const sheet = host()
      const memory = createCfMemory()
      const ready = cfWorkbookOfDocument([sheet, other(50)], memory, 'engine')
      const pendingSheet = (id: string): CfWorkbookSheet | undefined => { const s = ready.sheet(id); return s && s.sheetId === 'sheet2' ? { ...s, state: () => 'pending' } : s }
      const { out, pending } = await prepareSheet(sheet, ['B10'], crossCaps, { workbook: { sheetIdOf: ready.sheetIdOf, sheet: pendingSheet } })
      expect(out.get('B10')).toMatchObject({ kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: ['sheet2'] })
      expect(pending).toEqual(['sheet2'])
    })
    it('a commit on the referenced sheet makes the prepared cells stale', async () => {
      const sheet = host()
      const memory = createCfMemory()
      let gen = 0
      const base = cfWorkbookOfDocument([sheet, other(50)], memory, 'engine')
      const moving: CfWorkbookSource = { sheetIdOf: base.sheetIdOf, sheet: id => { const s = base.sheet(id); return s && s.sheetId === 'sheet2' ? { ...s, valueGen: () => gen } : s } }
      const { at } = await prepareSheet(sheet, ['B10'], crossCaps, { workbook: moving })
      expect(visualOf(at('B10'))).toMatch(/^\{/)
      gen++
      expect(at('B10')).toMatchObject({ kind: 'unresolved', reason: 'cf-stale-generation' })
    })
  })
})

/* ------------------------------------------------------------------ */
/* num-val references (Stage 2's verified id plus formula:cfvo-reference). */
/* ------------------------------------------------------------------ */

describe('num-val references lift only with both capabilities', () => {
  const xml = scale('B2:B12', 1, '<cfvo type="num" val="$A$1"/><cfvo type="max"/>')
  const all = capsWith()
  it('resolves with scale:thresholds-num-val-reference and formula:cfvo-reference', async () => {
    expect(VERIFIED_CAPABILITIES.has('scale:thresholds-num-val-reference')).toBe(true)
    const d = definitionOf(parseRules(xml, all), 1)
    expect(d.kind).toBe('colorScale')
    expect(d.kind === 'colorScale' && d.points[0]!.bound).toMatchObject({ type: 'formula', cfvo: 'num' })
    const twin = scale('C2:C12', 2, '<cfvo type="num" val="40"/><cfvo type="max"/>')
    const sheet = sheetOf(parseRules(xml + twin, all), { A1: 40, ...col('B', DATA), ...col('C', DATA) })
    const { out } = await prepareSheet(sheet, [...rows('B', 2, 12), ...rows('C', 2, 12)], all)
    twinsAgree(out, rows('B', 2, 12).map(r => [r, `C${r.slice(1)}`] as const))
  })
  for (const missing of ['scale:thresholds-num-val-reference', 'formula:cfvo-reference']) {
    it(`without ${missing} it is cf-unsupported-rule, as on main`, () => {
      const d = definitionOf(parseRules(xml, all.filter(c => c !== missing)), 1)
      expect(d).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    })
  }
})

/* ------------------------------------------------------------------ */
/* S5-I and S5-J: data-driven from the recorded outcome.                */
/* ------------------------------------------------------------------ */

const manifest = (): CfManifest => JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
/** The sheet XML Excel saved for a case, only when the case carries an Excel observation (Task 31). */
async function observedSheetXml(caseId: string): Promise<string | undefined> {
  const c = manifest().cases.find(x => x.id === caseId)
  if (!c || c.expected.provenance !== 'excel-observed') return undefined
  return (await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, c.saved))))).text('xl/worksheets/sheet1.xml')
}
type S5IOutcome = { dialect: 'unrecorded' } | { dialect: 'x14' } | { dialect: 'base'; reason: string }
/** S5-I as Task 31 recorded it: a refusal in CF_STAGE5_UNSUPPORTED, else what Excel saved, else unrecorded. */
async function recordedS5I(): Promise<S5IOutcome> {
  const reason = CF_STAGE5_UNSUPPORTED['scale:dialect-standalone']
  if (reason !== undefined) return { dialect: 'base', reason }
  const xml = await observedSheetXml('stage5-x14-scale')
  if (xml === undefined) return { dialect: 'unrecorded' }
  return /<x14:colorScale\b/.test(xml) ? { dialect: 'x14' } : { dialect: 'base', reason: 'Excel saved the scale in the base dialect' }
}
type S5JOutcome = { kind: 'unrecorded' } | { kind: 'retained' } | { kind: 'stripped'; reason: string }
async function recordedS5J(visual: 'scale' | 'bar' | 'icon'): Promise<S5JOutcome> {
  const reason = CF_STAGE5_UNSUPPORTED[`formula:active-present-${visual}`]
  if (reason !== undefined) return { kind: 'stripped', reason }
  const xml = await observedSheetXml(`stage5-active-present-${visual}`)
  if (xml === undefined) return { kind: 'unrecorded' }
  return /activePresent="(1|true)"/.test(xml) ? { kind: 'retained' } : { kind: 'stripped', reason: 'Excel stripped activePresent on resave' }
}
async function loadFixture(name: string, caps: Iterable<string>) {
  const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, `tests/fixtures/xlsx-cf/${name}.xlsx`)))
  return withExactCfEvidence(caps, async () => parseXlsx(await OfficePackage.load(bytes)))
}
const X14_SCALE_CAPS = capsWith('formula:cross-sheet-reference', 'scale:dialect-standalone')

/** What production does under each S5-I outcome (the stage5-x14-scale seed workbook: B is x14, C its base twin). */
async function assertS5I(outcome: S5IOutcome): Promise<void> {
  if (outcome.dialect === 'x14') {
    // Excel kept the x14 form: with scale:dialect-standalone the x14 scale resolves to the same points as its base twin.
    const doc = await loadFixture('stage5-x14-scale', X14_SCALE_CAPS)
    const sheet = doc.sheets.find(s => s.name === 'Sheet1')!
    const memory = createCfMemory()
    const { out } = await prepareSheet(sheet, [...rows('B', 2, 9), ...rows('C', 2, 9)], X14_SCALE_CAPS, { workbook: cfWorkbookOfDocument(doc.sheets, memory, 'engine') })
    twinsAgree(out, rows('B', 2, 9).map(r => [r, `C${r.slice(1)}`] as const))
    return
  }
  // Base form recorded, or nothing recorded yet: production keeps the x14 shape cf-unsupported-rule.
  if (outcome.dialect === 'base') expect(outcome.reason.length, 'the recorded refusal carries its reason').toBeGreaterThan(0)
  // Task 39: a base (or unrecorded) outcome would never publish scale:dialect-standalone, so production's evidence here is
  // the published set without it (the recorded outcome is x14, and production does hold it).
  const doc = await loadFixture('stage5-x14-scale', [...VERIFIED_CAPABILITIES].filter(id => id !== 'scale:dialect-standalone'))
  const rule = doc.sheets.find(s => s.name === 'Sheet1')!.conditionalFormatting!.rules.find(r => r.priority === 1)!
  expect(rule.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
}

describe('x14 colour scales follow the recorded dialect (S5-I)', () => {
  it('with scale:dialect-standalone injected (Excel kept x14), the x14 scale resolves to the same points as its base twin', async () => {
    await assertS5I({ dialect: 'x14' })
  })
  it('if Task 31 recorded the base form, the x14 shape stays cf-unsupported-rule with the recorded reason', async () => {
    await assertS5I({ dialect: 'base', reason: 'test-injected: Excel saved the base dialect' })
  })
  it('under the outcome actually recorded', async () => {
    await assertS5I(await recordedS5I())
  })
  it('without scale:dialect-standalone the x14 scale is cf-unsupported-rule even with every other capability', async () => {
    const caps = X14_SCALE_CAPS.filter(c => c !== 'scale:dialect-standalone')
    const doc = await loadFixture('stage5-x14-scale', caps)
    expect(doc.sheets[0]!.conditionalFormatting!.rules.find(r => r.priority === 1)!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
  })
  it('a literal x14 scale parses identically on the resident and streamed paths', async () => {
    const x14 = `<x14:conditionalFormatting><x14:cfRule type="colorScale" priority="1" id="{00000000-0000-4000-8000-0000000000AA}"><x14:colorScale>`
      + `<x14:cfvo type="num"><xm:f>20</xm:f></x14:cfvo><x14:cfvo type="formula"><xm:f>$F$1</xm:f></x14:cfvo><x14:color rgb="FFF8696B"/><x14:color rgb="FF63BE7B"/></x14:colorScale></x14:cfRule><xm:sqref>B2:B9</xm:sqref></x14:conditionalFormatting>`
    const inner = `<sheetData/>${ext(x14)}`
    const resident = parseRules(inner, X14_SCALE_CAPS)
    const streamed = await streamRules(inner, X14_SCALE_CAPS)
    expect(resident.rules[0]!.definition.kind).toBe('colorScale')
    expect(streamed.rules.map(r => r.definition)).toEqual(resident.rules.map(r => r.definition))
  })
})

const ACTIVE_CAPS = (visual: 'scale' | 'bar' | 'icon'): string[] => capsWith(
  'formula:truth-boolean', `formula:active-present-${visual}`,
  ...(visual === 'scale' ? ['scale:dialect-standalone'] : visual === 'bar' ? ['bar:pairing-standalone'] : []),
)
/** The active rows of each seed workbook (A > 0) and the row span. */
const ACTIVE: Record<'scale' | 'bar' | 'icon', { rows: [number, number]; on: readonly number[] }> = {
  scale: { rows: [2, 5], on: [2, 4] }, icon: { rows: [2, 5], on: [2, 4] }, bar: { rows: [2, 7], on: [3, 5, 7] },
}

/** What production does under each S5-J outcome, for one visual kind. */
async function assertS5J(visual: 'scale' | 'bar' | 'icon', outcome: S5JOutcome): Promise<void> {
  const name = `stage5-active-present-${visual}`
  if (outcome.kind === 'retained') {
    const caps = ACTIVE_CAPS(visual)
    const doc = await loadFixture(name, caps)
    const sheet = doc.sheets[0]!
    const [from, to] = ACTIVE[visual].rows
    const { out } = await prepareSheet(sheet, [...rows('B', from, to), ...rows('C', from, to)], caps)
    for (let r = from; r <= to; r++) {
      if (ACTIVE[visual].on.includes(r)) {
        expect(visualOf(out.get(`B${r}`)), `B${r} active`).toBe(visualOf(out.get(`C${r}`)))
        expect(visualOf(out.get(`C${r}`)), `C${r}`).toMatch(/^\{|none/)
      } else {
        // An activation false at a cell withholds that cell's visual only: no visual, nothing tainted.
        expect(visualOf(out.get(`B${r}`)), `B${r} inactive`).toBe('none')
      }
    }
    // A missing or extra rule-level <xm:f> stays cf-invalid-rule.
    const xml = (await (await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, `tests/fixtures/xlsx-cf/${name}.xlsx`))))).text('xl/worksheets/sheet1.xml'))!
    const body = xml.slice(xml.indexOf('<sheetData'), xml.lastIndexOf('</worksheet>')).replace(/<sheetData>.*<\/sheetData>/s, '<sheetData/>')
    for (const variant of [body.replace('<xm:f>$A2&gt;0</xm:f>', ''), body.replace('<xm:f>$A2&gt;0</xm:f>', '<xm:f>$A2&gt;0</xm:f><xm:f>$A2&gt;1</xm:f>')]) {
      const rules = parseRules(variant, caps)
      expect(rules.rules.find(r => r.priority === 1)!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-rule' })
    }
    return
  }
  if (outcome.kind === 'stripped') {
    expect(outcome.reason.length).toBeGreaterThan(0)
    expect(VERIFIED_CAPABILITIES.has(`formula:active-present-${visual}`)).toBe(false)
  }
  // Stripped, or nothing recorded yet: every activePresent="1" visual stays cf-unsupported-rule under production evidence.
  const doc = await loadFixture(name, VERIFIED_CAPABILITIES)
  expect(doc.sheets[0]!.conditionalFormatting!.rules.find(r => r.priority === 1)!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
}

describe('activePresent follows its recorded outcome (S5-J)', () => {
  for (const visual of ['scale', 'bar', 'icon'] as const) {
    it(`${visual}: if Excel stripped it, the visual stays cf-unsupported-rule with the recorded reason`, async () => {
      await assertS5J(visual, { kind: 'stripped', reason: 'test-injected: Excel stripped activePresent' })
    })
    it(`${visual}: if Excel retained it, an activation false at a cell withholds that cell's visual only`, async () => {
      await assertS5J(visual, { kind: 'retained' })
    })
    it(`${visual}: under the outcome actually recorded`, async () => {
      await assertS5J(visual, await recordedS5J(visual))
    })
  }
  it('without its activation capability the visual is cf-unsupported-rule, before any structural check', () => {
    const caps = ACTIVE_CAPS('icon').filter(c => c !== 'formula:active-present-icon')
    const xml = `<sheetData/>${ext(`<x14:conditionalFormatting><x14:cfRule type="iconSet" priority="1" activePresent="1" id="{00000000-0000-4000-8000-0000000000AB}"><x14:iconSet iconSet="3Arrows"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>B2:B5</xm:sqref></x14:conditionalFormatting>`)}`
    expect(parseRules(xml, caps).rules[0]!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
  })
  it('with the activation compiled but its capability missing from the service, the visual withholds', async () => {
    expect(cfFormulaReading('formula:active-present-icon', () => true)).toBe('activation-gates')
    const caps = ACTIVE_CAPS('icon')
    const doc = await loadFixture('stage5-active-present-icon', caps)
    const { out } = await prepareSheet(doc.sheets[0]!, ['B2'], caps.filter(c => c !== 'formula:active-present-icon'))
    expect(visualOf(out.get('B2'))).toMatch(WITHHELD)
  })
})

describe('fix round 1: formula bar provenance, activation first, publication', () => {
  it('a formula-resolved bar threshold is authorised by bar:bounds-formula, not bar:bounds-num; literal num keeps its gate (ruling 2)', async () => {
    const formulaBar = pairedBar('B2:B12', 1, '{00000000-0000-4000-8000-000000000021}', '<cfvo type="min"/><cfvo type="formula" val="$F$1"/>', '<x14:cfvo type="min"/><x14:cfvo type="formula"><xm:f>$F$1</xm:f></x14:cfvo>')
    const twinBar = pairedBar('C2:C12', 2, '{00000000-0000-4000-8000-000000000022}', '<cfvo type="min"/><cfvo type="num" val="50"/>', '<x14:cfvo type="min"/><x14:cfvo type="num"><xm:f>50</xm:f></x14:cfvo>')
    const xml = formulaBar.base + twinBar.base + ext(formulaBar.x14 + twinBar.x14)
    const cells = { F1: 50, ...col('B', BAR_DATA), ...col('C', BAR_DATA) }
    const noNum = capsWith().filter(c => c !== 'bar:bounds-num')
    const { out } = await prepareSheet(sheetOf(parseRules(xml, noNum), cells), [...rows('B', 2, 12), ...rows('C', 2, 12)], noNum)
    for (const r of rows('B', 3, 12)) {
      const o = out.get(r)!
      expect(o.kind === 'ready' && o.value.visual?.dataBar !== undefined, `${r} formula bar paints`).toBe(true)
      expect(visualOf(out.get(`C${r.slice(1)}`)), `C${r.slice(1)} literal num twin`).toMatch(WITHHELD)
    }
    // With both, the formula bar equals its literal twin.
    const both = await prepareSheet(sheetOf(parseRules(xml, capsWith()), cells), [...rows('B', 2, 12), ...rows('C', 2, 12)], capsWith())
    twinsAgree(both.out, rows('B', 2, 12).map(r => [r, `C${r.slice(1)}`] as const))
    // Without bar:bounds-formula in the service, the formula bar withholds.
    const noFormula = await prepareSheet(sheetOf(parseRules(xml, capsWith()), cells), ['B5'], capsWith().filter(c => c !== 'bar:bounds-formula'))
    expect(visualOf(noFormula.out.get('B5'))).toMatch(WITHHELD)
    // Stage 3 refusals stay: data beyond the formula bound withholds the whole rule.
    const beyond = await prepareSheet(sheetOf(parseRules(xml, noNum), { ...cells, B12: 70 }), ['B5'], noNum)
    expect(visualOf(beyond.out.get('B5'))).toMatch(WITHHELD)
  })
  it('a confirmed false activation is no-match even when the scope statistics refuse (ruling 3)', async () => {
    const caps = ACTIVE_CAPS('scale')
    const x14 = `<x14:conditionalFormatting><x14:cfRule type="colorScale" priority="1" activePresent="1" id="{00000000-0000-4000-8000-0000000000F1}"><xm:f>$A2&gt;0</xm:f><x14:colorScale><x14:cfvo type="min"/><x14:cfvo type="max"/><x14:color rgb="FFF8696B"/><x14:color rgb="FF63BE7B"/></x14:colorScale></x14:cfRule><xm:sqref>B2:B5</xm:sqref></x14:conditionalFormatting>`
    const rules = parseRules(`<sheetData/>${ext(x14)}`, caps)
    const clean = await prepareSheet(sheetOf(rules, { A2: -1, A3: 1, B2: 10, B3: 20, B4: 30, B5: 40 }), ['B2', 'B3'], caps)
    expect(visualOf(clean.out.get('B2'))).toBe('none')
    expect(visualOf(clean.out.get('B3'))).toMatch(/^\{/)
    // codex's repro: another scope cell #N/A makes the scope refuse (participation-error-presence); the inactive cell stays clean.
    const na = await prepareSheet(sheetOf(rules, { A2: -1, A3: 1, B2: 10, B3: 20, B4: { err: '#N/A' }, B5: 40 }), ['B2', 'B3'], caps.filter(c => c !== 'scale:participation-error-presence'))
    expect(visualOf(na.out.get('B2'))).toBe('none')
    expect(visualOf(na.out.get('B3'))).toMatch(WITHHELD)
  })
  it('a synthetic verified report does not publish bar:bounds-formula or formula:active-present-bar (finding 4)', () => {
    const base = auditStage5Manifest({ cases: [] } as never, { rootDir: REPO_ROOT })
    const ids = ['bar:bounds-formula', 'formula:active-present-bar', 'icon:thresholds-formula']
    const report: Stage5AuditReport = { ...base, ok: true, stageComplete: true, blocked: [], blockedCapabilities: [],
      perCapability: ids.map(capabilityId => ({ capabilityId, caseId: 'stage5-cfvo-formulas', disposition: 'verified' as const })),
      matchedReadings: { 'bar:bounds-formula': 'evaluated', 'formula:active-present-bar': 'activation-gates', 'icon:thresholds-formula': 'evaluated' } }
    const pub = stage5Publication(report)
    expect(pub.published).toEqual(['icon:thresholds-formula'])
    for (const id of ids.slice(0, 2)) expect(pub.withheld[id], id).toMatch(/fixture-unconfirmable/)
  })
})

/* ------------------------------------------------------------------ */
/* Numerics and slots.                                                  */
/* ------------------------------------------------------------------ */

describe('formula bounds stay finite and visuals keep their slots', () => {
  it('bounds at +-1e308 resolve finite or unresolved, never NaN', async () => {
    const caps = capsWith()
    for (const [f1, data] of [[1e308, [-1e308, 0, 1e308]], [-1e308, [-1e308, 0, 1e308]], [1.7976931348623157e308, [-1.7976931348623157e308, 1.7976931348623157e308]]] as const) {
      const xml = scale('B2:B4', 1, '<cfvo type="formula" val="$F$1"/><cfvo type="max"/>') + scale('C2:C4', 2, '<cfvo type="min"/><cfvo type="formula" val="$F$1"/>')
        + scale('D2:D4', 3, '<cfvo type="percent" val="$F$2"/><cfvo type="max"/>')
      const sheet = sheetOf(parseRules(xml, caps), { F1: f1, F2: 50, ...col('B', data), ...col('C', data), ...col('D', data) })
      const { out } = await prepareSheet(sheet, [...rows('B', 2, 1 + data.length), ...rows('C', 2, 1 + data.length), ...rows('D', 2, 1 + data.length)], caps)
      for (const [ref, o] of out) {
        if (o.kind !== 'ready') { expect(o.kind, ref).toBe('unresolved'); continue }
        const f = o.value.visual?.colorScale?.fraction
        if (f !== undefined) expect(Number.isFinite(f), `${ref} ${f}`).toBe(true)
        const argb = o.value.visual?.colorScale?.argb
        if (argb !== undefined) expect(Number.isFinite(argb), ref).toBe(true)
      }
    }
  })
  it('a formula-bounded bar and icon on one cell follow the Stage 3-4 slot gates unchanged', async () => {
    const formulaBar = pairedBar('B2:B12', 1, '{00000000-0000-4000-8000-000000000011}', '<cfvo type="min"/><cfvo type="formula" val="$F$1"/>', '<x14:cfvo type="min"/><x14:cfvo type="formula"><xm:f>$F$1</xm:f></x14:cfvo>')
    const twinBar = pairedBar('C2:C12', 3, '{00000000-0000-4000-8000-000000000012}', '<cfvo type="min"/><cfvo type="num" val="50"/>', '<x14:cfvo type="min"/><x14:cfvo type="num"><xm:f>50</xm:f></x14:cfvo>')
    const xml = formulaBar.base + icon('B2:B12', 2, '<cfvo type="percent" val="0"/><cfvo type="formula" val="$F$1"/><cfvo type="num" val="80"/>')
      + twinBar.base + icon('C2:C12', 4, '<cfvo type="percent" val="0"/><cfvo type="num" val="50"/><cfvo type="num" val="80"/>')
      + ext(formulaBar.x14 + twinBar.x14)
    const caps = capsWith()
    const sheet = sheetOf(parseRules(xml, caps), { F1: 50, ...col('B', BAR_DATA), ...col('C', BAR_DATA) })
    const { out } = await prepareSheet(sheet, [...rows('B', 2, 12), ...rows('C', 2, 12)], caps)
    twinsAgree(out, rows('B', 2, 12).map(r => [r, `C${r.slice(1)}`] as const))
    // Both slots are drawn on each cell (the Stage 3-4 slot gates are unchanged by a formula bound).
    for (const r of rows('B', 3, 12)) {
      const o = out.get(r)!
      expect(o.kind === 'ready' && o.value.visual?.dataBar !== undefined && o.value.visual?.iconSet !== undefined, r).toBe(true)
    }
  })
})

/* ------------------------------------------------------------------ */
/* Catalog pins.                                                        */
/* ------------------------------------------------------------------ */

describe('every cfvo, dialect and activation capability production queries is a catalog id with a catalog reading', () => {
  const catalog = new Set(CF_STAGE5_CAPABILITIES.map(e => e.capabilityId))
  it('the ids Task 37 gates on are catalog ids', () => {
    const ids = ['scale:thresholds-formula', 'bar:bounds-formula', 'icon:thresholds-formula', 'formula:cfvo-reference', 'formula:cfvo-expression',
      'formula:cfvo-relative', 'formula:cfvo-nonnumeric', 'formula:cfvo-scope-multi-range', 'scale:dialect-standalone',
      'formula:active-present-scale', 'formula:active-present-bar', 'formula:active-present-icon']
    for (const id of ids) {
      expect(catalog.has(id), id).toBe(true)
      expect(CF_FORMULA_PRODUCTION_READINGS[id], `${id} has a production reading`).toBeDefined()
    }
    // The literal Stage 5 visual ids in the sources are catalog ids too.
    const src = ['service.ts', 'parse.ts', 'colorScale.ts', 'dataBar.ts', 'iconSet.ts', 'formulaReadings.ts'].map(f => readFileSync(join(REPO_ROOT, 'src/xlsx/cf', f), 'utf8')).join('\n')
    for (const m of src.matchAll(/'(scale:dialect-standalone|scale:thresholds-formula|bar:bounds-formula|icon:thresholds-formula)'/g)) expect(catalog.has(m[1]!), m[1]).toBe(true)
  })
  it('the evaluation context is required: without it a formula-bounded scale withholds', async () => {
    const caps = capsWith()
    const sheet = sheetOf(parseRules(scale('B2:B12', 1, '<cfvo type="formula" val="$F$1"/><cfvo type="max"/>'), caps), { F1: 50, ...col('B', DATA) })
    const { out } = await prepareSheet(sheet, ['B10'], caps, { evaluation: false })
    expect(visualOf(out.get('B10'))).toBe('withheld')
  })
})

describe('a formula bound never reaches a resolver unsubstituted', () => {
  it('each resolver refuses a formula bound (cf-formula-unevaluated)', () => {
    const rules = parseRules(scale('B2:B4', 1, '<cfvo type="formula" val="$F$1"/><cfvo type="max"/>')
      + icon('C2:C4', 2, '<cfvo type="percent" val="0"/><cfvo type="formula" val="$F$1"/><cfvo type="num" val="80"/>'), capsWith())
    const stats = { numericCount: 3, min: 0, max: 10, sortedAsc: [0, 5, 10] } as never
    const all = { hasEvidence: () => true }
    const sc = definitionOf(rules, 1), ic = definitionOf(rules, 2)
    expect(resolveScaleThresholds(sc, stats, all)).toEqual({ kind: 'unresolved', reason: 'cf-formula-unevaluated' })
    expect(ic.kind === 'iconSet' && resolveIconThresholds(ic, stats, all)).toEqual({ kind: 'unresolved', reason: 'cf-formula-unevaluated' })
    const f = sc.kind === 'colorScale' ? sc.points[0]!.bound : undefined
    expect(resolveBarThresholds({ bounds: [{ type: 'min' }, f] } as never, stats, all)).toEqual({ kind: 'unresolved', reason: 'cf-formula-unevaluated' })
  })
})

/* ------------------------------------------------------------------ */
/* Every view agrees: the resident leg and the paged worker leg.        */
/* ------------------------------------------------------------------ */

describe('the resident and paged worker legs agree on formula bounds (Lesson 13)', () => {
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const rowsXml = (cells: Readonly<Record<string, number>>): string => {
    const byRow = new Map<number, string[]>()
    for (const [ref, v] of Object.entries(cells)) { const { row } = rc(ref); byRow.set(row, [...(byRow.get(row) ?? []), `<c r="${ref}"><v>${v}</v></c>`]) }
    return [...byRow].sort((a, b) => a[0] - b[0]).map(([r, cs]) => `<row r="${r + 1}">${cs.join('')}</row>`).join('')
  }
  const book = (sheets: ReadonlyArray<{ name: string; cells: Readonly<Record<string, number>>; cf?: string }>): Uint8Array => buildZip([
    { name: '[Content_Types].xml', data: `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>` },
    { name: '_rels/.rels', data: `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
    { name: 'xl/workbook.xml', data: `<?xml version="1.0"?><workbook xmlns="${NS}" xmlns:r="${R}"><sheets>${sheets.map((x, i) => `<sheet name="${x.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>` },
    { name: 'xl/_rels/workbook.xml.rels', data: `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rIdS" Type="${R}/styles" Target="styles.xml"/></Relationships>` },
    { name: 'xl/styles.xml', data: `<?xml version="1.0"?><styleSheet xmlns="${NS}"><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>` },
    ...sheets.map((x, i) => ({ name: `xl/worksheets/sheet${i + 1}.xml`, data: `<?xml version="1.0"?><worksheet xmlns="${NS}" xmlns:x14="${X14}" xmlns:xm="${XM}"><sheetData>${rowsXml(x.cells)}</sheetData>${x.cf ?? ''}</worksheet>` })),
  ]).bytes
  it('a formula-bounded scale and an x14 scale bounded by a never-activated Sheet2 paint the same on both legs', async () => {
    const caps = X14_SCALE_CAPS
    const x14 = `<x14:conditionalFormatting><x14:cfRule type="colorScale" priority="2" id="{00000000-0000-4000-8000-0000000000E1}"><x14:colorScale><x14:cfvo type="num"><xm:f>Sheet2!$A$1</xm:f></x14:cfvo><x14:cfvo type="max"/><x14:color rgb="FFF8696B"/><x14:color rgb="FF63BE7B"/></x14:colorScale></x14:cfRule><xm:sqref>C2:C12</xm:sqref></x14:conditionalFormatting>`
    const cf = scale('B2:B12', 1, '<cfvo type="formula" val="$F$1"/><cfvo type="max"/>') + ext(x14)
    const bytes = book([{ name: 'Sheet1', cells: { F1: 50, ...(col('B', DATA) as Record<string, number>), ...(col('C', DATA) as Record<string, number>) }, cf }, { name: 'Sheet2', cells: { A1: 30 } }])
    const refs = [...rows('B', 2, 12), ...rows('C', 2, 12)]
    // Resident leg.
    const doc = await withExactCfEvidence(caps, async () => parseXlsx(await OfficePackage.load(bytes)))
    const host = doc.sheets.find(x => x.name === 'Sheet1')!
    const resident = await prepareSheet(host, refs, caps, { workbook: cfWorkbookOfDocument(doc.sheets, createCfMemory(), 'engine') })
    // Paged worker leg (Sheet2 is never activated: its store is created by the same path a cross-sheet read uses).
    const session = await withExactCfEvidence(caps, async () => { const x = await openImportSession({ kind: 'blob', blob: new Blob([bytes.buffer as ArrayBuffer]) }, { docId: 'cfvo-legs', profile: 'D' }); await x.done; return x })
    const source = createPagedWorkbookSource(session)
    const adapter = source.store(source.sheets[0]!.sheetId) as PagedXlsxSheetAdapter
    const read: CfReadResult = await withExactCfEvidence(caps, () => adapter.readConditionalFormatting({
      rects: refs.map(ref => { const { row, col: c } = rc(ref); return { r0: row, r1: row, c0: c, c1: c } }),
      stamp: { owner: adapter.stamp.docId, sheetId: adapter.stamp.sheetId, workbookValueGen: adapter.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: adapter.structureGen, clockGen: 0 },
      grantBytes: 1 << 22,
    }))
    expect(read.kind).toBe("ready")
    for (const ref of refs) {
      const { row, col: c } = rc(ref)
      const wire = read.states.find(x => x.row === row && x.col === c)
      const o = resident.out.get(ref)!
      const mine = o.kind === 'ready' ? o.value.visual?.colorScale : undefined
      expect(mine, `${ref} resident paints`).toBeDefined()
      expect(wire?.unresolvedReason, ref).toBeUndefined()
      expect({ argb: wire?.scale?.argb, fraction: wire?.scale?.fraction }, ref).toEqual({ argb: mine!.argb, fraction: mine!.fraction })
    }
  })
})

void cfDocumentEvaluationOf
void residentSource
