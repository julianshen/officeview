// @vitest-environment node
/**
 * Stage 5, Task 32 (round-2 #3): the evaluation context is built once per document and the clock once per preparation, and
 * every preparation route forwards it before any formula evaluates. A route that is given none leaves formula rules
 * `cf-evaluation-context-missing`; Stage 1-4 rules are unchanged either way.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { getChildren, parseXml, type XmlNode } from '../src/core/xml'
import { createCfMemory, createCfWorkBudget } from '../src/xlsx/cf/limits'
import { withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { parseCfSheet, type CfStylesSource } from '../src/xlsx/cf/parse'
import {
  cfDocumentEvaluationOf, cfDocumentEvaluationOfPaged, cfEvaluationContextFor, type CfDocumentEvaluation,
} from '../src/xlsx/cf/formula'
import * as renderModule from '../src/xlsx/cf/render'
import * as textModule from '../src/xlsx/cf/text'
import * as serviceModule from '../src/xlsx/cf/service'
import * as composeModule from '../src/xlsx/cf/compose'
import { workbookValuePolicy } from '../src/xlsx/formula/value-policy'
import { prepareCfRemote, PagedXlsxSheetAdapter } from '../src/xlsx/view/pagedSource'
import type { CfLease, CfSheetRules, CfStamp, CfValuePage, CfValueSource } from '../src/xlsx/cf/model'
import type { CfReadRequest } from '../src/xlsx/view/contracts'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'
import type { FormatBudget } from '../src/xlsx/view/formatWorkspace'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { xlsxBytes, xlsxEntries } from './helpers/xlsx-package'
import { buildZip } from './helpers/zip-builder'

vi.mock('../src/xlsx/cf/service', async importOriginal => {
  const original = await importOriginal<typeof import('../src/xlsx/cf/service')>()
  return { ...original, createCfService: vi.fn(original.createCfService) }
})
vi.mock('../src/xlsx/cf/compose', async importOriginal => {
  const original = await importOriginal<typeof import('../src/xlsx/cf/compose')>()
  return { ...original, composeCellAsync: vi.fn(original.composeCellAsync) }
})
const createService = vi.mocked(serviceModule.createCfService)
const compose = vi.mocked(composeModule.composeCellAsync)
beforeEach(() => { createService.mockClear(); compose.mockClear() })
/** Every rule outcome the service handed the compositor, by rule priority (the cell state keeps only the taint). */
const outcomesByPriority = (): Map<number, unknown[]> => {
  const out = new Map<number, unknown[]>()
  for (const [rules, outcomes] of compose.mock.calls) rules.forEach((rule, i) => out.set(rule.priority, [...(out.get(rule.priority) ?? []), outcomes[i]]))
  return out
}
const shape = (outcome: { kind: string; value?: { properties: object; unresolvedWrites?: ReadonlySet<string> } }) =>
  outcome.kind === 'ready' ? { properties: outcome.value!.properties, tainted: [...(outcome.value!.unresolvedWrites ?? [])].sort() } : outcome

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const DXF_FILL = '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>'
const PARSE_SET = ['formula:dialect-base', 'formula:anchor-single', 'formula:anchor-mixed-abs'] as const
const TAIPEI: CfDocumentEvaluation = Object.freeze({ calcMode: 'auto', valuePolicy: 'engine', dateSystem: '1900', timeZone: 'Asia/Taipei' })

const stylesFor = (): CfStylesSource => {
  const root = parseXml(`<styleSheet xmlns="${NS}"><dxfs count="1">${DXF_FILL}</dxfs></styleSheet>`)
  return { colors: {}, dxfNodes: getChildren(getChildren(root, 'dxfs')[0] as XmlNode, 'dxf') }
}
const CF_XML = `<conditionalFormatting sqref="A1:A3"><cfRule type="cellIs" priority="2" dxfId="0" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>`
  + `<conditionalFormatting sqref="A1:A3"><cfRule type="expression" priority="1" dxfId="0"><formula>$B1&gt;0</formula></cfRule></conditionalFormatting>`
  + `<conditionalFormatting sqref="C1:C3"><cfRule type="cellIs" priority="3" dxfId="0" operator="greaterThan"><formula>$F$1</formula></cfRule></conditionalFormatting>`
function formulaRules(): CfSheetRules {
  return withExactCfEvidence(PARSE_SET, () => parseCfSheet(parseXml(`<worksheet xmlns="${NS}">${CF_XML}</worksheet>`), stylesFor(), createCfMemory()).value)
}
function residentSheet(rules: CfSheetRules): XlsxSheet {
  const cell = (row: number, col: number, value: number) => ({ ref: `R${row}C${col}`, row, col, styleIndex: 0, style: { numFmtId: 0 }, value, hasCachedValue: true })
  return {
    name: 'S', sheetId: 's', rows: [{ index: 0, cells: [cell(0, 0, 5), cell(0, 2, 5)] }, { index: 1, cells: [cell(1, 0, 0)] }], cols: [], merges: [], mergeRanges: [],
    conditionalFormatting: rules, semantics: { dateSystem: '1900', timeZone: 'Asia/Taipei', locale: 'en-US', epochNowMs: 0, unicode: { version: 1, source: 'workbook-default' } },
  } as unknown as XlsxSheet
}
const stamp = (rules: CfSheetRules): CfStamp => ({ owner: 'o', sheetId: 's', workbookValueGen: 1, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 0, ...(rules ? {} : {}) })
function fakeSource(stampValue: CfStamp): CfValueSource {
  return {
    bounds: async () => ({ firstRow: 0, lastRow: 2, firstCol: 0, lastCol: 2 }),
    stamp: () => stampValue,
    pages: async function* () {
      yield { value: { rows: Uint32Array.from([0, 0]), cols: Uint16Array.from([0, 2]), values: [{ kind: 'number', value: 5 }, { kind: 'number', value: 5 }] }, bytes: 0, release() {} } as unknown as CfLease<CfValuePage>
    },
  } as unknown as CfValueSource
}
const budget = (): FormatBudget => ({ reserve: () => {}, release: () => {}, retain: (_targets: readonly object[], _bytes: number, commit: () => void) => commit(), settle: () => {} })

describe('the document policy is built once and resolves its defaults', () => {
  it('equal on repeated calls; frozen; an undefined valuePolicy is engine and an absent calc is auto', () => {
    const doc = { semantics: { dateSystem: '1904', timeZone: 'Asia/Taipei' } } as unknown as XlsxDocument
    const a = cfDocumentEvaluationOf(doc), b = cfDocumentEvaluationOf(doc)
    expect(a).toEqual(b)
    expect(Object.isFrozen(a)).toBe(true)
    expect(a).toEqual({ calcMode: 'auto', valuePolicy: 'engine', dateSystem: '1904', timeZone: 'Asia/Taipei' })
    expect(cfDocumentEvaluationOf({} as XlsxDocument)).toEqual({ calcMode: 'auto', valuePolicy: 'engine', dateSystem: '1900', timeZone: '' })
  })
  it('follows the workbook value policy and calculation mode exactly', () => {
    for (const valuePolicy of [undefined, 'engine', 'cached'] as const) {
      const doc = { valuePolicy } as XlsxDocument
      expect(cfDocumentEvaluationOf(doc).valuePolicy).toBe(workbookValuePolicy(doc))
    }
    for (const calcMode of ['auto', 'manual', 'autoNoTable'] as const) {
      expect(cfDocumentEvaluationOf({ calc: { calcMode, fullCalcOnLoad: false, iterate: false, iterateCount: 100, iterateDelta: 0.001 } } as XlsxDocument).calcMode).toBe(calcMode)
    }
  })
  it('the paged twin is always cached, reads the recorded calc mode and the workbook date system', () => {
    expect(cfDocumentEvaluationOfPaged({ calcMode: 'manual', formatting: { dateSystem: '1904' } })).toMatchObject({ calcMode: 'manual', valuePolicy: 'cached', dateSystem: '1904' })
    expect(cfDocumentEvaluationOfPaged({})).toMatchObject({ calcMode: 'auto', valuePolicy: 'cached', dateSystem: '1900' })
    expect(cfDocumentEvaluationOfPaged({}).timeZone).not.toBe('')
  })
})

describe('the clock is taken once per preparation from the injected now', () => {
  it('two contexts from one policy take two snapshots; the local-midnight tick advances clockGen', () => {
    const t0 = Date.UTC(2026, 10, 16, 15, 59, 59), t1 = Date.UTC(2026, 10, 16, 16, 0, 0)
    const now = vi.fn().mockReturnValueOnce(t0).mockReturnValueOnce(t1)
    const first = cfEvaluationContextFor(TAIPEI, now), second = cfEvaluationContextFor(TAIPEI, now)
    expect(now).toHaveBeenCalledTimes(2)
    expect([first.clock.epochNowMs, second.clock.epochNowMs]).toEqual([t0, t1])
    expect(second.clock.clockGen).toBe(first.clock.clockGen + 1)
    expect(first).toMatchObject(TAIPEI)
    expect(Object.isFrozen(first) && Object.isFrozen(first.clock)).toBe(true)
    // An unresolvable zone still advances by UTC day instead of freezing at zero.
    expect(cfEvaluationContextFor({ ...TAIPEI, timeZone: '' }, () => Date.UTC(2026, 10, 16)).clock.clockGen).toBe(Math.floor(Date.UTC(2026, 10, 16) / 86_400_000))
  })
})

describe('every route forwards the policy and takes one clock snapshot per preparation', () => {
  it('prepareCfPaint', async () => {
    const rules = formulaRules()
    const sheet = residentSheet(rules)
    const now = vi.fn().mockReturnValueOnce(1000).mockReturnValueOnce(2000)
    await renderModule.prepareCfPaint(sheet, { memory: createCfMemory(), evaluation: TAIPEI, now })?.then(batch => batch?.dispose())
    await renderModule.prepareCfPaint(sheet, { memory: createCfMemory(), evaluation: TAIPEI, now })?.then(batch => batch?.dispose())
    expect(now).toHaveBeenCalledTimes(2)
    const seen = createService.mock.calls.map(call => call[2]?.evaluation)
    expect(seen.length).toBeGreaterThanOrEqual(2)
    expect(seen[0]).toMatchObject({ ...TAIPEI, clock: { epochNowMs: 1000 } })
    expect(seen.at(-1)).toMatchObject({ clock: { epochNowMs: 2000 } })
    createService.mockClear()
    await renderModule.prepareCfPaint(sheet, { memory: createCfMemory() })?.then(batch => batch?.dispose())
    expect(createService.mock.calls.every(call => call[2]?.evaluation === undefined)).toBe(true)
  })
  it('prepareCfTextScope', async () => {
    const rules = formulaRules()
    const sheet = residentSheet(rules)
    const now = vi.fn().mockReturnValue(4242)
    const run = (evaluation?: CfDocumentEvaluation) => textModule.prepareCfTextScope(sheet, {
      rects: [{ r0: 0, r1: 2, c0: 0, c1: 2 }], signal: new AbortController().signal, locale: 'en-US', budget: budget(), ...(evaluation ? { evaluation, now } : {}),
    })
    await run(TAIPEI)
    expect(now).toHaveBeenCalledTimes(1)
    expect(createService.mock.calls[0]![2]?.evaluation).toMatchObject({ ...TAIPEI, clock: { epochNowMs: 4242 } })
    createService.mockClear()
    await run()
    expect(createService.mock.calls[0]![2]?.evaluation).toBeUndefined()
  })
  it('prepareCfRemote (the worker provider)', async () => {
    const rules = formulaRules()
    const st = { ...stamp(rules), cfGen: (await import('../src/xlsx/view/pagedSource')).cfRulesVersion(rules) }
    const request: CfReadRequest = { rects: [{ r0: 0, r1: 2, c0: 0, c1: 2 }], stamp: { owner: st.owner, sheetId: st.sheetId, workbookValueGen: st.workbookValueGen, cfGen: st.cfGen, viewGen: st.viewGen, structureGen: st.structureGen, clockGen: st.clockGen }, grantBytes: 1 << 24 }
    const now = vi.fn().mockReturnValue(777)
    await prepareCfRemote(request, 's', fakeSource(st), rules, createCfMemory(), { budget: () => createCfWorkBudget({ now: () => 0 }), evaluation: TAIPEI, now })
    expect(now).toHaveBeenCalledTimes(1)
    expect(createService.mock.calls[0]![2]?.evaluation).toMatchObject({ ...TAIPEI, clock: { epochNowMs: 777 } })
    createService.mockClear()
    await prepareCfRemote(request, 's', fakeSource(st), rules, createCfMemory(), { budget: () => createCfWorkBudget({ now: () => 0 }) })
    expect(createService.mock.calls[0]![2]?.evaluation).toBeUndefined()
  })
  it('the worker adapter derives the policy from its PagedWorkbook: cached, with the recorded calc mode', async () => {
    const rules = formulaRules()
    const fakeSheet = { sheetId: 's', name: 's', state: 'complete', importedThroughRow: 1048575, metadata: {}, tables: undefined, arrayCells: [], drawingAnchors: { records: [], unresolved: 0 }, mergeRects: [], invalidMerges: 0, used: undefined, conditionalFormatting: rules }
    const workbook = { calcMode: 'manual', formatting: { dateSystem: '1904' } }
    const adapter = new PagedXlsxSheetAdapter({ docId: 'doc', sheetId: 's', valueGen: 1 }, workbook as never, fakeSheet as never, undefined, {}, false)
    const st = { ...stamp(rules), cfGen: (await import('../src/xlsx/view/pagedSource')).cfRulesVersion(rules), owner: 'owner' }
    try {
      await adapter.readConditionalFormatting({ rects: [{ r0: 0, r1: 2, c0: 0, c1: 2 }], stamp: st, grantBytes: 1 << 24 }, createCfMemory(), { budget: () => createCfWorkBudget({ now: () => 0 }) }).catch(() => undefined)
      expect(createService.mock.calls.length).toBeGreaterThan(0)
      expect(createService.mock.calls[0]![2]?.evaluation).toMatchObject({ calcMode: 'manual', valuePolicy: 'cached', dateSystem: '1904' })
    } finally { await adapter.release().catch(() => {}) }
    // A sheet with no formula rule never loads the formula module or carries a context.
    createService.mockClear()
    const plain: CfSheetRules = { diagnostics: [], rules: [rules.rules.find(r => r.definition.kind === 'cellIs' && r.priority === 2)!] }
    const plainAdapter = new PagedXlsxSheetAdapter({ docId: 'doc', sheetId: 's', valueGen: 1 }, workbook as never, { ...fakeSheet, conditionalFormatting: plain } as never, undefined, {}, false)
    try {
      await plainAdapter.readConditionalFormatting({ rects: [{ r0: 0, r1: 2, c0: 0, c1: 2 }], stamp: { ...st, cfGen: (await import('../src/xlsx/view/pagedSource')).cfRulesVersion(plain) }, grantBytes: 1 << 24 }, createCfMemory(), { budget: () => createCfWorkBudget({ now: () => 0 }) }).catch(() => undefined)
      expect(createService.mock.calls.every(call => call[2]?.evaluation === undefined)).toBe(true)
    } finally { await plainAdapter.release().catch(() => {}) }
  })
  it('the paged import records calcMode beside fullCalcOnLoad (absent means auto)', async () => {
    const open = async (calcPr: string) => {
      const entries = xlsxEntries({}).map(entry => entry.name === 'xl/workbook.xml' ? { ...entry, data: String(entry.data).replace('</workbook>', `${calcPr}</workbook>`) } : entry)
      const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(buildZip(entries).bytes).buffer]) }, { docId: `calc-${calcPr.length}`, profile: 'D' })
      try { await session.done; return session.workbook.calcMode } finally { await session.release() }
    }
    expect(await open('<calcPr calcMode="manual"/>')).toBe('manual')
    expect(await open('<calcPr calcMode="autoNoTable"/>')).toBe('autoNoTable')
    expect(await open('<calcPr calcMode="nonsense"/>')).toBe('auto')
    expect(await open('')).toBe('auto')
    void xlsxBytes
  })
})

describe('a deeply nested formula does not make the paged view refuse what the resident view holds', () => {
  it('a 150-term chain (depth 150) hashes, sizes and prepares on the worker path', async () => {
    const chain = Array.from({ length: 150 }, () => '1').join('+')
    const rules = withExactCfEvidence(PARSE_SET, () => parseCfSheet(parseXml(`<worksheet xmlns="${NS}"><conditionalFormatting sqref="A1:A3"><cfRule type="expression" priority="1" dxfId="0"><formula>${chain}&gt;0</formula></cfRule></conditionalFormatting></worksheet>`), stylesFor(), createCfMemory()).value)
    expect(rules.rules[0]!.definition.kind).toBe('expression')
    const { cfRulesVersion } = await import('../src/xlsx/view/pagedSource')
    const st = { ...stamp(rules), cfGen: cfRulesVersion(rules) }
    // The version is a function of the rule's text and ranges, not of the AST's depth.
    expect(cfRulesVersion(rules)).toBe(cfRulesVersion(withExactCfEvidence(PARSE_SET, () => parseCfSheet(parseXml(`<worksheet xmlns="${NS}"><conditionalFormatting sqref="A1:A3"><cfRule type="expression" priority="1" dxfId="0"><formula>${chain}&gt;0</formula></cfRule></conditionalFormatting></worksheet>`), stylesFor(), createCfMemory()).value)))
    const out = await prepareCfRemote({ rects: [{ r0: 0, r1: 2, c0: 0, c1: 2 }], stamp: { owner: st.owner, sheetId: st.sheetId, workbookValueGen: st.workbookValueGen, cfGen: st.cfGen, viewGen: st.viewGen, structureGen: st.structureGen, clockGen: st.clockGen }, grantBytes: 1 << 24 }, 's', fakeSource(st), rules, createCfMemory(), { budget: () => createCfWorkBudget({ now: () => 0 }) })
    expect(out.kind).toBe('ready')
    expect(out.diagnostics.every(d => d.feature !== 'cf-scope-too-large')).toBe(true)
  })
})

describe('a route given no policy leaves formula rules withheld by name; Stage 1-4 rules are unchanged', () => {
  async function prepare(evaluation: ReturnType<typeof cfEvaluationContextFor> | undefined) {
    const rules = formulaRules()
    const st = stamp(rules)
    const service = serviceModule.createCfService(rules, createCfMemory(), { ...(evaluation ? { evaluation } : {}) })
    const lease = await service.prepareCf({ stamp: st, sourceRects: [{ firstRow: 0, lastRow: 2, firstCol: 0, lastCol: 2 }], signal: new AbortController().signal, grantBytes: 1 << 22 }, fakeSource(st))
    return { lease, service, at: (row: number, col: number) => lease.value.at(row, col) }
  }
  it('without a context: the expression and the formula operand are cf-evaluation-context-missing', async () => {
    const { lease, service, at } = await prepare(undefined)
    const by = outcomesByPriority()
    expect(by.get(1)).toContainEqual({ kind: 'unresolved', reason: 'cf-evaluation-context-missing' })
    expect(by.get(3)).toContainEqual({ kind: 'unresolved', reason: 'cf-evaluation-context-missing' })
    // The cell keeps only the taint: nothing paints, and the rule's writes are withheld from the rules below it.
    expect(shape(at(0, 0) as never)).toEqual({ properties: {}, tainted: ['fill'] })
    expect(shape(at(0, 2) as never)).toEqual({ properties: {}, tainted: ['fill'] })
    lease.release(); service.dispose()
  })
  it('with a context the formula is still withheld until the evaluator lands, never evaluated as a literal', async () => {
    const { lease, service, at } = await prepare(cfEvaluationContextFor(TAIPEI, () => 0))
    const by = outcomesByPriority()
    expect(by.get(1)).toContainEqual({ kind: 'unresolved', reason: 'cf-formula-unevaluated' })
    expect(by.get(3)).toContainEqual({ kind: 'unresolved', reason: 'cf-formula-unevaluated' })
    expect(shape(at(0, 0) as never)).toEqual({ properties: {}, tainted: ['fill'] })
    lease.release(); service.dispose()
  })
  it('a literal rule alone decides identically with and without a context', async () => {
    const plain = withExactCfEvidence([], () => parseCfSheet(parseXml(`<worksheet xmlns="${NS}">${CF_XML.split('</conditionalFormatting>')[0]}</conditionalFormatting></worksheet>`), stylesFor(), createCfMemory()).value)
    const run = async (evaluation?: ReturnType<typeof cfEvaluationContextFor>) => {
      const st = stamp(plain)
      const service = serviceModule.createCfService(plain, createCfMemory(), { ...(evaluation ? { evaluation } : {}) })
      const lease = await service.prepareCf({ stamp: st, sourceRects: [{ firstRow: 0, lastRow: 2, firstCol: 0, lastCol: 2 }], signal: new AbortController().signal, grantBytes: 1 << 22 }, fakeSource(st))
      const out = [lease.value.at(0, 0), lease.value.at(1, 0)].map(o => shape(o as never))
      lease.release(); service.dispose()
      return out
    }
    const without = await run(), withContext = await run(cfEvaluationContextFor(TAIPEI, () => 0))
    expect(withContext).toEqual(without)
    expect(without[0]).toEqual({ properties: expect.objectContaining({ 'fill.pattern': 'solid' }), tainted: [] })
  })
})

describe('eager sites import no CF value and reach the builder only through the loaded lazy module', () => {
  const read = (path: string) => readFileSync(new URL(`../${path}`, import.meta.url), 'utf8')
  const staticCfValueImports = (source: string): string[] => source.split('\n').filter(line => /^\s*(import|export)\b/.test(line) && !/^\s*(import|export) type\b/.test(line) && /from\s+['"][^'"]*\/cf\/[^'"]*['"]/.test(line))
  it.each([
    ['src/render/paint.ts', /cf\.cfDocumentEvaluationOf\(doc\)/],
    ['src/components/xlsx/sheetAssets.ts', /cf\.cfDocumentEvaluationOf\(doc\)/],
    ['src/components/xlsx/InteractiveSheetDoc.tsx', /cf\.cfDocumentEvaluationOf\(doc\)/],
    ['src/components/xlsx/InteractiveSheetDoc.tsx', /text\.cfDocumentEvaluationOf\(doc\)/],
  ])('%s', (path, receiver) => {
    const source = read(path)
    expect(staticCfValueImports(source)).toEqual([])
    expect(source).toMatch(receiver)
    expect(source).not.toMatch(/import\s*\{[^}]*cfDocumentEvaluationOf[^}]*\}\s*from/)
  })
  it('the paged provider, the loader and the import load the formula module only dynamically (or not at all)', () => {
    for (const path of ['src/xlsx/view/pagedSource.ts', 'src/xlsx/cfLoader.ts', 'src/xlsx/storage/pagedImport.ts', 'src/worker/viewerSession.ts']) {
      expect(staticCfValueImports(read(path)), path).toEqual([])
    }
    expect(read('src/xlsx/view/pagedSource.ts')).toMatch(/await import\('\.\.\/cf\/formula'\)/)
    expect(read('src/xlsx/cfLoader.ts')).not.toMatch(/cf\/formula/)
  })
})
