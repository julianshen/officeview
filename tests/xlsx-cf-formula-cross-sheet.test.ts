// @vitest-environment node
/**
 * Stage 5, Task 36 (Ruling S5-K, Review Focus 4): CF formulas reading other sheets of the same workbook. Values and the
 * generation they are published under come from one overlay-aware store per sheet; `dependencyGen` (the sum of the
 * referenced sheets' value generations) rides every stamp, so `at()` goes stale synchronously at a referenced commit on
 * the resident, store and worker legs alike. No Stage 5 capability is verified (no Excel capture has run), so every paint
 * assertion here is under injected exact evidence.
 */
import { describe, expect, it } from 'vitest'
import { buildZip } from './helpers/zip-builder'
import { CF_LIMITS, createCfMemory } from '../src/xlsx/cf/limits'
import { cfEvaluationContextFor, compileCfFormula, isCfFormulaRefusal, cfTreeReadsMatchView } from '../src/xlsx/cf/formula'
import { createCfService } from '../src/xlsx/cf/service'
import { cfWorkbookOfDocument, cfWorkbookOfStores, prepareCfPaint, residentSource, storeCfSource } from '../src/xlsx/cf/render'
import {
  CF_REFERENCED_SHEETS_LIMIT, cfDependencyGenOf, cfDependencyIdentities, prepareCfFormulaBatch, withCfDependencyGen,
  type CfWorkbookSheet, type CfWorkbookSource,
} from '../src/xlsx/cf/dependencies'
import { injectCfFormulaSelections, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { cfRulesVersion, createPagedWorkbookSource, prepareCfRemote, PagedXlsxSheetAdapter } from '../src/xlsx/view/pagedSource'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { measurePayloadBytes } from '../src/xlsx/view/payloadBytes'
import { parseFormula } from '../src/xlsx/formula/parser'
import type { CfCellState, CfOutcome, CfSheetRules, CfStamp, CfValueSource, SourceRect } from '../src/xlsx/cf/model'
import type { CfReadResult, SourceStore, ValueOverlay } from '../src/xlsx/view/contracts'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'
import { FORMULA_CAPS, GREEN, NS, SELECTIONS, UTC, parseRules, rc, sheetOf, shown } from './helpers/cf-formula-rules'

const CROSS = ['formula:cross-sheet-reference', 'formula:cross-sheet-quoted-name', 'formula:cellis-operand-cross-sheet']
const CAPS: readonly string[] = [...FORMULA_CAPS, ...CROSS]
const expr = (sqref: string, priority: number, formula: string, dxf = 1): string =>
  `<conditionalFormatting sqref="${sqref}"><cfRule type="expression" priority="${priority}" dxfId="${dxf}"><formula>${formula}</formula></cfRule></conditionalFormatting>`
const named = (name: string, sheetId: string, cells: Parameters<typeof sheetOf>[1], rules?: CfSheetRules): XlsxSheet =>
  ({ ...sheetOf(rules as CfSheetRules, cells), name, sheetId, ...(rules ? {} : { conditionalFormatting: undefined }) }) as XlsxSheet
const docOf = (sheets: XlsxSheet[]): XlsxDocument => ({ sheets }) as unknown as XlsxDocument
const rectOf = (ref: string): SourceRect => { const { row, col } = rc(ref); return { firstRow: row, lastRow: row, firstCol: col, lastCol: col } }
const fillOf = (o: CfOutcome<CfCellState> | undefined): string => (o?.kind === 'unresolved' ? o.reason : shown(o).withheld ? 'withheld' : shown(o).fill)
const evaluation = () => cfEvaluationContextFor(UTC, () => Date.UTC(2026, 10, 16))
const evidence = (caps: Iterable<string>) => { const set = new Set(caps); return { allows: (cap: string) => set.has(cap) } }
/** A one-cell value overlay (a committed pivot output), the commit both resident and worker stores take. */
function overlayOf(cells: Readonly<Record<string, number>>): ValueOverlay {
  const at = new Map(Object.entries(cells).map(([ref, v]) => { const { row, col } = rc(ref); return [`${row},${col}`, v] as const }))
  const rects = Object.keys(cells).map(ref => { const { row, col } = rc(ref); return { r0: row, r1: row, c0: col, c1: col } })
  return { rects, protectedRects: rects, merges: [], cell: (row, col) => { const v = at.get(`${row},${col}`); return v === undefined ? undefined : { value: v } } }
}
const commit = (store: SourceStore, overlay: ValueOverlay | undefined): void => (store as unknown as { setValueOverlay(o: ValueOverlay | undefined): void }).setValueOverlay(overlay)

/** Sheet1 `=Sheet2!A2>0` over A2:A3 painting green; Sheet2 A2 = 5, A3 = -1. */
function crossSheetDoc(formula = 'Sheet2!A2&gt;0', sheet2: Parameters<typeof sheetOf>[1] = { A2: 5, A3: -1 }) {
  const rules = parseRules(expr('A2:A3', 1, formula), CAPS)
  const s1 = named('Sheet1', 'sheet1', { A2: 1, A3: 1 }, rules)
  const s2 = named('Sheet2', 'sheet2', sheet2)
  return { rules, s1, s2, doc: docOf([s1, s2]) }
}
/** The store leg: the service over `storeCfSource` of the host's canonical store, the workbook over the same stores. */
async function prepareStoreLeg(rules: CfSheetRules, stores: { sheets: readonly { sheetId: string; name: string }[]; store(id: string): SourceStore }, refs: readonly string[], caps = CAPS) {
  const memory = createCfMemory()
  const workbook = cfWorkbookOfStores(stores.sheets, id => stores.store(id), memory)
  const host = stores.store('sheet1')
  const service = createCfService(rules, memory, { evidence: evidence(caps), evaluation: evaluation(), sheetName: 'Sheet1', workbook })
  const stamp: CfStamp = { owner: host.stamp.docId, sheetId: 'sheet1', workbookValueGen: host.stamp.valueGen, cfGen: 1, viewGen: 0, structureGen: host.structureGen ?? 0, clockGen: 0,
    dependencyGen: cfDependencyGenOf(rules, 'Sheet1', workbook) }
  injectCfFormulaSelections(SELECTIONS)
  try {
    const lease = await service.prepareCf({ stamp, sourceRects: refs.map(rectOf), signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, storeCfSource(host, stamp, memory))
    return { lease, service, workbook, stamp }
  } finally { injectCfFormulaSelections(null) }
}
const at = (lease: { value: { at(r: number, c: number): CfOutcome<CfCellState> } }, ref: string) => { const { row, col } = rc(ref); return lease.value.at(row, col) }

/* ------------------------------------------------------------------ */
/* A paged workbook (the worker leg): two sheets, CF on Sheet1.         */
/* ------------------------------------------------------------------ */
const CT = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`
const RELS = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const DXFS = `<dxfs count="2"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf><dxf><fill><patternFill patternType="solid"><bgColor rgb="${GREEN}"/></patternFill></fill></dxf></dxfs>`
const rowsXml = (cells: Readonly<Record<string, number>>): string => {
  const rows = new Map<number, string[]>()
  for (const [ref, v] of Object.entries(cells)) { const { row } = rc(ref); rows.set(row, [...(rows.get(row) ?? []), `<c r="${ref}"><v>${v}</v></c>`]) }
  return [...rows].sort((a, b) => a[0] - b[0]).map(([r, cs]) => `<row r="${r + 1}">${cs.join('')}</row>`).join('')
}
function pagedBytes(sheets: ReadonlyArray<{ name: string; cells: Readonly<Record<string, number>>; cf?: string }>): Uint8Array {
  const wb = `<?xml version="1.0"?><workbook xmlns="${NS}" xmlns:r="${R}"><sheets>${sheets.map((s, i) => `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`
  const wbRels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rIdS" Type="${R}/styles" Target="styles.xml"/></Relationships>`
  return buildZip([
    { name: '[Content_Types].xml', data: CT }, { name: '_rels/.rels', data: RELS }, { name: 'xl/workbook.xml', data: wb }, { name: 'xl/_rels/workbook.xml.rels', data: wbRels },
    { name: 'xl/styles.xml', data: `<?xml version="1.0"?><styleSheet xmlns="${NS}"><cellXfs count="1"><xf numFmtId="0"/></cellXfs>${DXFS}</styleSheet>` },
    ...sheets.map((s, i) => ({ name: `xl/worksheets/sheet${i + 1}.xml`, data: `<?xml version="1.0"?><worksheet xmlns="${NS}"><sheetData>${rowsXml(s.cells)}</sheetData>${s.cf ?? ''}</worksheet>` })),
  ]).bytes
}
async function openPaged(sheets: Parameters<typeof pagedBytes>[0]) {
  const bytes = pagedBytes(sheets)
  const session = await withExactCfEvidence(CAPS, async () => {
    const s = await openImportSession({ kind: 'blob', blob: new Blob([bytes.buffer as ArrayBuffer]) }, { docId: 'paged-cross', profile: 'D' })
    await s.done
    return s
  })
  return { session, source: createPagedWorkbookSource(session) }
}
async function readWorker(store: SourceStore, refs: readonly string[], grantBytes = 1 << 22): Promise<CfReadResult> {
  const adapter = store as PagedXlsxSheetAdapter
  injectCfFormulaSelections(SELECTIONS)
  try {
    return await withExactCfEvidence(CAPS, () => adapter.readConditionalFormatting({
      rects: refs.map(ref => { const { row, col } = rc(ref); return { r0: row, r1: row, c0: col, c1: col } }),
      stamp: { owner: adapter.stamp.docId, sheetId: adapter.stamp.sheetId, workbookValueGen: adapter.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: adapter.structureGen, clockGen: 0 },
      grantBytes,
    }))
  } finally { injectCfFormulaSelections(null) }
}
const wireFill = (result: CfReadResult, ref: string): string => {
  const { row, col } = rc(ref)
  const s = result.states.find(x => x.row === row && x.col === col)
  return !s ? 'none' : s.unresolvedReason ?? (s.unresolvedWrites?.length ? 'withheld' : String(s.properties['fill.bg'] ?? 'none'))
}

describe('a referenced-sheet commit makes the old batch stale synchronously (Review Focus 4)', () => {
  it('store leg: green at g1; Sheet2 A2 = -1 commits; the same tick at() is cf-stale-generation; g2 paints base; a late g1 never publishes', async () => {
    const { rules, doc } = crossSheetDoc()
    const wb = createObjectWorkbookSource(doc, 'resident-cross')
    const g1 = await prepareStoreLeg(rules, wb, ['A2', 'A3'])
    expect(fillOf(at(g1.lease, 'A2'))).toBe(GREEN)
    expect(fillOf(at(g1.lease, 'A3'))).toBe('none')
    commit(wb.store('sheet2'), overlayOf({ A2: -1 }))
    // Synchronously, before any await: the old batch never paints after the commit.
    expect(at(g1.lease, 'A2')).toEqual({ kind: 'unresolved', reason: 'cf-stale-generation' })
    expect(at(g1.lease, 'A3')).toEqual({ kind: 'unresolved', reason: 'cf-stale-generation' })
    g1.lease.release(); g1.service.dispose()
    const g2 = await prepareStoreLeg(rules, wb, ['A2'])
    expect(fillOf(at(g2.lease, 'A2'))).toBe('none')
    g2.lease.release(); g2.service.dispose()
    // A preparation started at g2 whose referenced sheet commits before it adopts: cancelled, never published.
    const memory = createCfMemory()
    const workbook = cfWorkbookOfStores(wb.sheets, id => wb.store(id), memory)
    const host = wb.store('sheet1')
    const service = createCfService(rules, memory, { evidence: evidence(CAPS), evaluation: evaluation(), sheetName: 'Sheet1', workbook })
    const stamp: CfStamp = { owner: host.stamp.docId, sheetId: 'sheet1', workbookValueGen: host.stamp.valueGen, cfGen: 1, viewGen: 0, structureGen: host.structureGen ?? 0, clockGen: 0, dependencyGen: cfDependencyGenOf(rules, 'Sheet1', workbook) }
    injectCfFormulaSelections(SELECTIONS)
    const late = service.prepareCf({ stamp, sourceRects: [rectOf('A2')], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, storeCfSource(host, stamp, memory))
    commit(wb.store('sheet2'), undefined) // the release: A2 is 5 again
    await expect(late).rejects.toMatchObject({ code: 'cancelled' })
    injectCfFormulaSelections(null); service.dispose()
  })

  it('resident leg (the interactive CfPaintBatch): its layer stops painting at the referenced commit, in the same tick', async () => {
    const { s1, doc } = crossSheetDoc()
    const wb = createObjectWorkbookSource(doc, 'resident-batch')
    injectCfFormulaSelections(SELECTIONS)
    const batch = await withExactCfEvidence(CAPS, () => prepareCfPaint(s1, { store: wb.store('sheet1'), workbook: wb, valuePolicy: 'cached', evaluation: UTC, paintRects: [rectOf('A2'), rectOf('A3')], now: () => Date.UTC(2026, 10, 16) }))
    injectCfFormulaSelections(null)
    expect(batch).toBeDefined()
    expect(batch!.stamp.dependencyGen ?? 0).toBe(0)
    expect(batch!.layer.styleAt(1, 0, undefined)?.fillColor).toBe(GREEN)
    commit(wb.store('sheet2'), overlayOf({ A2: -1 }))
    expect(batch!.layer.styleAt(1, 0, undefined)?.fillColor).toBeUndefined()
    batch!.dispose()
  })

  it('worker leg: the paged adapter reads Sheet2 through its own adapter; the reply lists it; a commit makes the service stale', async () => {
    const { source } = await openPaged([
      { name: 'Sheet1', cells: { A2: 1, A3: 1 }, cf: expr('A2:A3', 1, 'Sheet2!A2&gt;0') },
      { name: 'Sheet2', cells: { A2: 5, A3: -1 } },
    ])
    const s1 = source.store(source.sheets[0]!.sheetId)
    const first = await readWorker(s1, ['A2', 'A3'])
    expect(first.kind).toBe('ready')
    expect(wireFill(first, 'A2')).toBe(GREEN)
    expect(wireFill(first, 'A3')).toBe('none')
    expect(first.dependencies).toEqual([{ sheetId: source.sheets[1]!.sheetId, valueGen: 0 }])
    commit(source.store(source.sheets[1]!.sheetId), overlayOf({ A2: -1 }))
    const second = await readWorker(s1, ['A2'])
    expect(wireFill(second, 'A2')).toBe('none')
    expect(second.dependencies).toEqual([{ sheetId: source.sheets[1]!.sheetId, valueGen: 1 }])
  })
})

describe('cross-sheet values and generation come from one overlay-aware store (round-1 #6)', () => {
  it('saved 5, overlay -5, release: the rule matches, then not, then again; value and stamp from the same store', async () => {
    const { rules, doc } = crossSheetDoc()
    const wb = createObjectWorkbookSource(doc, 'overlay-cross')
    const s2 = wb.store('sheet2')
    const run = async (): Promise<{ fill: string; dependencyGen: number | undefined }> => {
      const leg = await prepareStoreLeg(rules, wb, ['A2'])
      const out = { fill: fillOf(at(leg.lease, 'A2')), dependencyGen: leg.stamp.dependencyGen }
      leg.lease.release(); leg.service.dispose()
      return out
    }
    expect(await run()).toEqual({ fill: GREEN, dependencyGen: 0 })
    commit(s2, overlayOf({ A2: -5 }))
    expect(await run()).toEqual({ fill: 'none', dependencyGen: s2.stamp.valueGen })
    commit(s2, undefined)
    expect(await run()).toEqual({ fill: GREEN, dependencyGen: s2.stamp.valueGen })
    expect(s2.stamp.valueGen).toBe(2)
  })

  it('worker: the overlay on the never-activated Sheet2 adapter is what the formula reads; activation reuses that adapter', async () => {
    const { source } = await openPaged([
      { name: 'Sheet1', cells: { A2: 1 }, cf: expr('A2', 1, 'Sheet2!A2&gt;0') },
      { name: 'Sheet2', cells: { A2: 5 } },
    ])
    const s1 = source.store(source.sheets[0]!.sheetId)
    expect(wireFill(await readWorker(s1, ['A2']), 'A2')).toBe(GREEN)
    // Activation after the cross-sheet read returns the very adapter that read was served from (one adapter per sheet).
    const s2 = source.store(source.sheets[1]!.sheetId)
    commit(s2, overlayOf({ A2: -5 }))
    expect(wireFill(await readWorker(s1, ['A2']), 'A2')).toBe('none')
    commit(s2, undefined)
    const released = await readWorker(s1, ['A2'])
    expect(wireFill(released, 'A2')).toBe(GREEN)
    expect(released.dependencies).toEqual([{ sheetId: s2.stamp.sheetId, valueGen: 2 }])
    expect(source.store(source.sheets[1]!.sheetId)).toBe(s2)
  })

  it('a commit between the snapshot build and adoption cancels the preparation (cf-stale-generation on the held batch)', async () => {
    const { rules, doc } = crossSheetDoc()
    const wb = createObjectWorkbookSource(doc, 'race-cross')
    const memory = createCfMemory()
    const stores = cfWorkbookOfStores(wb.sheets, id => wb.store(id), memory)
    // Sheet2's values are served once, then the commit lands before the preparation adopts.
    const real = stores.sheet('sheet2')!
    let served = false
    const racing: CfWorkbookSheet = { ...real, values: { ...real.values, pages: (id, rects, budget, signal) => {
      const inner = real.values.pages(id, rects, budget, signal)
      return (async function* () { for await (const page of inner) yield page; if (!served) { served = true; commit(wb.store('sheet2'), overlayOf({ A2: -1 })) } })()
    } } }
    const workbook: CfWorkbookSource = { sheetIdOf: n => stores.sheetIdOf(n), sheet: id => (id === 'sheet2' ? racing : stores.sheet(id)) }
    const host = wb.store('sheet1')
    const service = createCfService(rules, memory, { evidence: evidence(CAPS), evaluation: evaluation(), sheetName: 'Sheet1', workbook })
    const stamp: CfStamp = { owner: host.stamp.docId, sheetId: 'sheet1', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0, dependencyGen: 0 }
    injectCfFormulaSelections(SELECTIONS)
    try {
      await expect(service.prepareCf({ stamp, sourceRects: [rectOf('A2')], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, storeCfSource(host, stamp, memory)))
        .rejects.toMatchObject({ code: 'cancelled' })
      expect(served).toBe(true)
      expect(service.at(stamp, 1, 0)).toEqual({ kind: 'unresolved', reason: 'cf-pending-data' })
    } finally { injectCfFormulaSelections(null); service.dispose() }
  })
})

describe('dependencyGen strictly increases on any referenced commit and ignores unrelated sheets', () => {
  it('Sheet2 and Sheet3 referenced: a Sheet3 commit moves it, a Sheet4 commit does not; a host commit moves workbookValueGen', () => {
    const rules = parseRules(expr('A2', 1, 'Sheet2!A2+Sheet3!A2&gt;0'), CAPS)
    const doc = docOf([named('Sheet1', 'sheet1', { A2: 1 }, rules), named('Sheet2', 'sheet2', { A2: 1 }), named('Sheet3', 'sheet3', { A2: 1 }), named('Sheet4', 'sheet4', { A2: 1 })])
    const wb = createObjectWorkbookSource(doc, 'gen-cross')
    const memory = createCfMemory()
    const workbook = cfWorkbookOfStores(wb.sheets, id => wb.store(id), memory)
    const host = wb.store('sheet1')
    const base: CfStamp = { owner: 'gen-cross', sheetId: 'sheet1', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
    const live = withCfDependencyGen(storeCfSource(host, base, memory), rules, 'Sheet1', workbook)
    const seen: number[] = [live.stamp('sheet1').dependencyGen!]
    commit(wb.store('sheet3'), overlayOf({ A2: 2 }))
    seen.push(live.stamp('sheet1').dependencyGen!)
    commit(wb.store('sheet4'), overlayOf({ A2: 2 }))
    seen.push(live.stamp('sheet1').dependencyGen!)
    commit(wb.store('sheet2'), overlayOf({ A2: 2 }))
    seen.push(live.stamp('sheet1').dependencyGen!)
    expect(seen).toEqual([0, 1, 1, 2])
    expect(cfDependencyIdentities(rules, 'Sheet1', workbook)).toEqual([{ sheetId: 'sheet2', valueGen: 1 }, { sheetId: 'sheet3', valueGen: 1 }])
    // The host's own commit moves its own generation, not dependencyGen.
    const before = live.stamp('sheet1')
    commit(host, overlayOf({ A2: 3 }))
    expect(live.stamp('sheet1').workbookValueGen).toBe(before.workbookValueGen + 1)
    expect(live.stamp('sheet1').dependencyGen).toBe(before.dependencyGen)
  })

  it('wrapping a store without a cross-sheet rule returns it untouched (no dependencyGen field, no workbook read)', () => {
    const rules = parseRules(expr('A2', 1, 'A2&gt;0'), CAPS)
    let reads = 0
    const workbook: CfWorkbookSource = { sheetIdOf: () => { reads++; return undefined }, sheet: () => { reads++; return undefined } }
    const host: CfValueSource = residentSource(named('Sheet1', 'sheet1', { A2: 1 }, rules), { owner: 'o', sheetId: 'sheet1', workbookValueGen: 3, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }, createCfMemory(), 'cached')
    expect(withCfDependencyGen(host, rules, 'Sheet1', workbook)).toBe(host)
    expect('dependencyGen' in host.stamp('sheet1')).toBe(false)
    expect(cfDependencyGenOf(rules, 'Sheet1', workbook)).toBe(0)
    expect(reads).toBe(0)
  })
})

/** A workbook sheet whose import state the test drives. */
function drivenSheet(name: string, sheetId: string, sheet: XlsxSheet, state: { now: 'ready' | 'pending' | 'failed' }): CfWorkbookSheet {
  const values = residentSource(sheet, { owner: 'd', sheetId, workbookValueGen: 0, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }, createCfMemory(), 'cached')
  return { sheetId, name, values: { ...values, bounds: (id, b, s) => (state.now === 'ready' ? values.bounds(id, b, s) : Promise.resolve(undefined)) }, state: () => state.now, valueGen: () => 0 }
}
async function prepareWith(rules: CfSheetRules, s1: XlsxSheet, workbook: CfWorkbookSource | undefined, refs: readonly string[], caps = CAPS, name: string | null = 'Sheet1') {
  const sheetName = name ?? undefined
  const memory = createCfMemory()
  const service = createCfService(rules, memory, { evidence: evidence(caps), evaluation: evaluation(), ...(sheetName !== undefined ? { sheetName } : {}), ...(workbook ? { workbook } : {}) })
  const stamp: CfStamp = { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0, dependencyGen: cfDependencyGenOf(rules, sheetName, workbook) }
  injectCfFormulaSelections(SELECTIONS)
  try {
    const lease = await service.prepareCf({ stamp, sourceRects: refs.map(rectOf), signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, residentSource(s1, stamp, memory, 'cached'))
    const out = new Map(refs.map(ref => [ref, at(lease, ref)] as const))
    const pending = lease.value.pendingSheetIds
    lease.release(); service.dispose()
    return { out, pending }
  } finally { injectCfFormulaSelections(null) }
}

describe('unimported referenced pages are pending, then repaint once', () => {
  it('a Sheet2 still importing: cells reading it are cf-pending-data naming Sheet2, no permanent diagnostic; it ends incomplete: cf-read-failed', async () => {
    const { rules, s1, s2 } = crossSheetDoc()
    const state = { now: 'pending' as 'ready' | 'pending' | 'failed' }
    const sheet2 = drivenSheet('Sheet2', 'sheet2', s2, state)
    const workbook: CfWorkbookSource = { sheetIdOf: n => (n.toLowerCase() === 'sheet2' ? 'sheet2' : undefined), sheet: id => (id === 'sheet2' ? sheet2 : undefined) }
    const pending = await prepareWith(rules, s1, workbook, ['A2', 'A3'])
    expect(pending.out.get('A2')).toEqual({ kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: ['sheet2'] })
    expect(pending.pending).toEqual(['sheet2'])
    expect(rules.diagnostics.filter(d => d.feature.includes('pending'))).toEqual([])
    state.now = 'ready'
    const ready = await prepareWith(rules, s1, workbook, ['A2', 'A3'])
    expect(fillOf(ready.out.get('A2'))).toBe(GREEN)
    expect(ready.pending).toBeUndefined()
    state.now = 'failed'
    const failed = await prepareWith(rules, s1, workbook, ['A2'])
    expect(failed.out.get('A2')?.kind).toBe('ready')
    expect(shown(failed.out.get('A2')).withheld).toBe(true)
    expect(failed.pending).toBeUndefined()
  })

  it('worker: a pending Sheet2 makes the reply pending on it, then ready once it completes', async () => {
    const { rules, s1, s2 } = crossSheetDoc()
    const state = { now: 'pending' as 'ready' | 'pending' | 'failed' }
    const sheet2 = drivenSheet('Sheet2', 'sheet2', s2, state)
    const workbook: CfWorkbookSource = { sheetIdOf: n => (n.toLowerCase() === 'sheet2' ? 'sheet2' : undefined), sheet: id => (id === 'sheet2' ? sheet2 : undefined) }
    const host = residentSource(s1, { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: cfRulesVersion(rules), viewGen: 0, structureGen: 0, clockGen: 0 }, createCfMemory(), 'cached')
    const read = async () => {
      injectCfFormulaSelections(SELECTIONS)
      try {
        return await withExactCfEvidence(CAPS, () => prepareCfRemote({ rects: [{ r0: 1, r1: 2, c0: 0, c1: 0 }], stamp: { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }, grantBytes: 1 << 22 },
          'sheet1', host, rules, createCfMemory(), { evaluation: UTC, workbook, sheetName: 'Sheet1' }))
      } finally { injectCfFormulaSelections(null) }
    }
    const first = await read()
    expect(first).toMatchObject({ kind: 'pending', reason: 'cf-pending-data', pendingSheetIds: ['sheet2'], states: [] })
    state.now = 'ready'
    const second = await read()
    expect(second.kind).toBe('ready')
    expect(wireFill(second, 'A2')).toBe(GREEN)
  })
})

describe('cross-sheet edge cases refuse identically', () => {
  it('an unknown sheet is unresolved (no workbook name); a quoted name resolves; no workbook or no host name withholds', async () => {
    const rules = parseRules(expr('A2', 1, 'Gone!A2&gt;0') + expr('B2', 2, "'My Sheet'!A2&gt;0"), CAPS)
    const s1 = named('Sheet1', 'sheet1', { A2: 1, B2: 1 }, rules)
    const doc = [s1, named('My Sheet', 'sheet3', { A2: 4 })]
    const workbook = cfWorkbookOfDocument(doc, createCfMemory(), 'cached')
    const got = await prepareWith(rules, s1, workbook, ['A2', 'B2'])
    expect(shown(got.out.get('A2')).withheld).toBe(true)
    expect(fillOf(got.out.get('B2'))).toBe(GREEN)
    for (const [wb, name] of [[undefined, 'Sheet1'], [workbook, null]] as const) {
      const refused = await prepareWith(rules, s1, wb, ['B2'], CAPS, name)
      expect(shown(refused.out.get('B2')).withheld).toBe(true)
    }
  })

  it('#REF! only under the missing-sheet reading (batch level): the call reads a #REF! literal, never a guessed value', async () => {
    const rules = parseRules(expr('A2', 1, 'ISERROR(Gone!A2)'), CAPS)
    const s1 = named('Sheet1', 'sheet1', { A2: 1 }, rules)
    const memory = createCfMemory()
    const workbook = cfWorkbookOfDocument([s1], memory, 'cached')
    const stamp: CfStamp = { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
    const formula = (rules.rules[0]!.definition as { formula: import('../src/xlsx/cf/model').CfFormula }).formula
    const run = async (missingSheet?: 'ref-error') => {
      const signal = new AbortController().signal
      const { createCfWorkBudget } = await import('../src/xlsx/cf/limits')
      const batch = await prepareCfFormulaBatch([{ rule: rules.rules[0]!, formula }], { source: residentSource(s1, stamp, memory, 'cached'), sheetId: 'sheet1', sheetName: 'Sheet1', stamp,
        evaluation: evaluation(), edge: undefined, memory, signal, budget: createCfWorkBudget({ signal, scan: true, maxUnits: 1_000_000 }), workbook, ...(missingSheet ? { missingSheet } : {}) })
      try { return batch.evaluate(0, 1, 0) } finally { batch.dispose() }
    }
    expect(await run()).toEqual({ kind: 'unresolved', reason: 'cf-formula-dependency' })
    expect(await run('ref-error')).toMatchObject({ kind: 'ready', value: { kind: 'boolean', value: true } })
  })

  it('more than 16 referenced sheets refuses (cf-resource-limit): one formula at compile, the sheet\'s union in the batch; an unknown generation refuses by name', async () => {
    const names = Array.from({ length: CF_REFERENCED_SHEETS_LIMIT + 2 }, (_, i) => `S${i + 2}`)
    const sum = (list: readonly string[]) => list.map(n => `${n}!A1`).join('+')
    // One formula naming 17 sheets is refused at compile.
    const one = parseRules(expr('A2', 1, `${sum(names.slice(0, CF_REFERENCED_SHEETS_LIMIT + 1))}&gt;0`), CAPS)
    expect(one.rules[0]!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    // Two formulas of 9 sheets each: the union is 18, so the rule that crosses 16 is refused and the first evaluates.
    const rules = parseRules(expr('A2', 1, `${sum(names.slice(0, 9))}&gt;0`) + expr('A2', 2, `${sum(names.slice(9, 18))}&gt;0`), CAPS)
    const s1 = named('Sheet1', 'sheet1', { A2: 1 }, rules)
    const sheets = [s1, ...names.map((n, i) => named(n, `x${i}`, { A1: 1 }))]
    const memory = createCfMemory()
    const formulaOf = (r: CfSheetRules, i: number) => (r.rules[i]!.definition as { formula: import('../src/xlsx/cf/model').CfFormula }).formula
    const stamp: CfStamp = { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
    const { createCfWorkBudget } = await import('../src/xlsx/cf/limits')
    const batchOf = async (r: CfSheetRules, workbook: CfWorkbookSource) => {
      const signal = new AbortController().signal
      return prepareCfFormulaBatch(r.rules.map((rule, i) => ({ rule, formula: formulaOf(r, i) })), { source: residentSource(s1, stamp, memory, 'cached'), sheetId: 'sheet1', sheetName: 'Sheet1', stamp,
        evaluation: evaluation(), edge: undefined, memory, signal, budget: createCfWorkBudget({ signal, scan: true, maxUnits: 1_000_000 }), workbook })
    }
    const union = await batchOf(rules, cfWorkbookOfDocument(sheets, memory, 'cached'))
    expect(union.outcome(0)).toEqual({ kind: 'ready', value: undefined })
    expect(union.outcome(1)).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    union.dispose()
    const two = parseRules(expr('A2', 1, 'S2!A1&gt;0'), CAPS)
    const doc = cfWorkbookOfDocument([s1, sheets[1]!], memory, 'cached')
    const unknownGen: CfWorkbookSource = { sheetIdOf: n => doc.sheetIdOf(n), sheet: id => { const s = doc.sheet(id); return s && { ...s, valueGen: () => undefined } } }
    const batch = await batchOf(two, unknownGen)
    expect(batch.outcome(0)).toEqual({ kind: 'unresolved', reason: 'cf-dependency-generation-unknown' })
    batch.dispose()
  })

  it('the host qualified by its own name is the host; a cross-sheet reference matches only a cross-sheet view area (guard regression)', async () => {
    const rules = parseRules(expr('A2', 1, 'Sheet1!B2&gt;0'), CAPS)
    const s1 = named('Sheet1', 'sheet1', { A2: 1, B2: 3 }, rules)
    expect(fillOf((await prepareWith(rules, s1, undefined, ['A2'])).out.get('A2'))).toBe(GREEN)
    const tree = (src: string) => { const ast = parseFormula(src); if ('error' in (ast as object) && (ast as { type?: string }).type === undefined) throw new Error('parse'); return ast as import('../src/xlsx/formula/types').AstNode }
    const b2 = { firstRow: 1, lastRow: 1, firstCol: 1, lastCol: 1 }
    expect(cfTreeReadsMatchView(tree('Sheet1!B2>0'), [{ rect: b2 }], 'Sheet1')).toBe(true)
    expect(cfTreeReadsMatchView(tree('sheet1!B2>0'), [b2], 'Sheet1')).toBe(true)
    expect(cfTreeReadsMatchView(tree('Sheet2!B2>0'), [{ sheet: 'Sheet2', rect: b2 }], 'Sheet1')).toBe(true)
    // Fail closed on genuine mismatches: another sheet against a host area, the host against another sheet's area.
    expect(cfTreeReadsMatchView(tree('Sheet2!B2>0'), [{ rect: b2 }], 'Sheet1')).toBe(false)
    expect(cfTreeReadsMatchView(tree('B2>0'), [{ sheet: 'Sheet2', rect: b2 }], 'Sheet1')).toBe(false)
    expect(cfTreeReadsMatchView(tree('Sheet3!B2>0'), [{ sheet: 'Sheet2', rect: b2 }], 'Sheet1')).toBe(false)
  })

  it('two sheets\' A1 never share a snapshot key: Sheet2!A1 and Sheet3!A1 are read separately', async () => {
    const rules = parseRules(expr('A2', 1, 'Sheet2!A1-Sheet3!A1=1'), CAPS)
    const s1 = named('Sheet1', 'sheet1', { A2: 1 }, rules)
    const wb = cfWorkbookOfDocument([s1, named('Sheet2', 'sheet2', { A1: 3 }), named('Sheet3', 'sheet3', { A1: 2 })], createCfMemory(), 'cached')
    expect(fillOf((await prepareWith(rules, s1, wb, ['A2'])).out.get('A2'))).toBe(GREEN)
    const swapped = cfWorkbookOfDocument([s1, named('Sheet2', 'sheet2', { A1: 2 }), named('Sheet3', 'sheet3', { A1: 3 })], createCfMemory(), 'cached')
    expect(fillOf((await prepareWith(rules, s1, swapped, ['A2'])).out.get('A2'))).toBe('none')
  })
})

describe('dependency bytes are measurer-derived', () => {
  it('16 dependencies are priced by measurePayloadBytes; at grant - 1 the reply is cf-grant-required with the exact size', async () => {
    const names = Array.from({ length: CF_REFERENCED_SHEETS_LIMIT }, (_, i) => `S${i + 2}`)
    const rules = parseRules(expr('A2', 1, `${names.map(n => `${n}!A1`).join('+')}&gt;0`), CAPS)
    const s1 = named('Sheet1', 'sheet1', { A2: 1 }, rules)
    const sheets = [s1, ...names.map((n, i) => named(n, `x${i}`, { A1: 1 }))]
    const workbook = cfWorkbookOfDocument(sheets, createCfMemory(), 'cached')
    const host = residentSource(s1, { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: cfRulesVersion(rules), viewGen: 0, structureGen: 0, clockGen: 0 }, createCfMemory(), 'cached')
    const read = async (grantBytes: number) => {
      injectCfFormulaSelections(SELECTIONS)
      try {
        return await withExactCfEvidence(CAPS, () => prepareCfRemote({ rects: [{ r0: 1, r1: 1, c0: 0, c1: 0 }], stamp: { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }, grantBytes },
          'sheet1', host, rules, createCfMemory(), { evaluation: UTC, workbook, sheetName: 'Sheet1' }))
      } finally { injectCfFormulaSelections(null) }
    }
    const full = await read(1 << 22)
    expect(full.kind).toBe('ready')
    expect(full.dependencies).toHaveLength(16)
    const { dependencies, ...without } = full
    const depBytes = measurePayloadBytes(full)! - measurePayloadBytes(without)!
    expect(depBytes).toBe(measurePayloadBytes({ dependencies })! - measurePayloadBytes({})!)
    expect(depBytes).toBeGreaterThan(16 * 2 * 8)
    // The exact grant: the smallest that returns the full reply; one byte less omits the cross-sheet part (its states and
    // the dependency list together, fix round 1), never a partial dependency list.
    const fullAt = async (g: number) => { const r = await read(g); return r.kind === 'ready' && r.dependencies?.length === 16 }
    let lo = 1, hi = 1 << 22
    while (lo < hi) { const mid = (lo + hi) >>> 1; if (await fullAt(mid)) hi = mid; else lo = mid + 1 }
    const short = await read(lo - 1)
    expect(short.kind).toBe('ready')
    expect(short.diagnostics.some(d => d.feature === 'cf-grant-required')).toBe(true)
    expect(short.states).toEqual([])
    expect(short.dependencies).toBeUndefined()
  })
})

describe('a no-cross-sheet workbook does zero cross-sheet work', () => {
  it('no CfWorkbookSource read, dependencyGen stays 0, and the reply carries no dependencies key', async () => {
    const rules = parseRules(expr('A2', 1, 'A2&gt;0'), CAPS)
    const s1 = named('Sheet1', 'sheet1', { A2: 1 }, rules)
    let reads = 0
    const spy: CfWorkbookSource = { sheetIdOf: () => { reads++; return undefined }, sheet: () => { reads++; return undefined } }
    const local = await prepareWith(rules, s1, spy, ['A2'])
    expect(fillOf(local.out.get('A2'))).toBe(GREEN)
    const host = residentSource(s1, { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: cfRulesVersion(rules), viewGen: 0, structureGen: 0, clockGen: 0 }, createCfMemory(), 'cached')
    injectCfFormulaSelections(SELECTIONS)
    const reply = await withExactCfEvidence(CAPS, () => prepareCfRemote({ rects: [{ r0: 1, r1: 1, c0: 0, c1: 0 }], stamp: { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }, grantBytes: 1 << 22 },
      'sheet1', host, rules, createCfMemory(), { evaluation: UTC, workbook: spy, sheetName: 'Sheet1' })).finally(() => injectCfFormulaSelections(null))
    expect(reply.kind).toBe('ready')
    expect('dependencies' in reply).toBe(false)
    expect(reads).toBe(0)
    // The worker adapter builds no workbook seam at all for such a sheet.
    const { source } = await openPaged([{ name: 'Sheet1', cells: { A2: 1 }, cf: expr('A2', 1, 'A2&gt;0') }, { name: 'Sheet2', cells: { A2: 1 } }])
    const workerReply = await readWorker(source.store(source.sheets[0]!.sheetId), ['A2'])
    expect('dependencies' in workerReply).toBe(false)
    expect(wireFill(workerReply, 'A2')).toBe(GREEN)
    void compileCfFormula; void isCfFormulaRefusal
  })
})

describe('cross-sheet hoists, worker races and the snapshot layer (Task 36)', () => {
  it('an aggregate over another sheet\'s range is never hoisted from the host: SUM(Sheet2!$A$1:$A$3) reads Sheet2', async () => {
    const caps = [...CAPS, 'formula:hoist-aggregate', 'formula:hoist-countif-frequency']
    const rules = parseRules(expr('B1', 1, 'SUM(Sheet2!$A$1:$A$3)&gt;5'), caps)
    const s1 = named('Sheet1', 'sheet1', { A1: 0, A2: 0, A3: 0, B1: 1 }, rules)
    const wb = cfWorkbookOfDocument([s1, named('Sheet2', 'sheet2', { A1: 3, A2: 3, A3: 3 })], createCfMemory(), 'cached')
    expect(fillOf((await prepareWith(rules, s1, wb, ['B1'], caps)).out.get('B1'))).toBe(GREEN)
  })

  it('worker: a referenced sheet that commits while the preparation reads it returns cf-stale-generation, never a published reply', async () => {
    const { rules, s1, s2 } = crossSheetDoc()
    let gen = 0
    const base = cfWorkbookOfDocument([s1, s2], createCfMemory(), 'cached').sheet('sheet2')!
    const racing: CfWorkbookSheet = { ...base, valueGen: () => gen, values: { ...base.values, pages: (id, rects, budget, signal) => {
      const inner = base.values.pages(id, rects, budget, signal)
      return (async function* () { for await (const page of inner) yield page; gen = 1 })()
    } } }
    const workbook: CfWorkbookSource = { sheetIdOf: n => (n.toLowerCase() === 'sheet2' ? 'sheet2' : undefined), sheet: id => (id === 'sheet2' ? racing : undefined) }
    const host = residentSource(s1, { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: cfRulesVersion(rules), viewGen: 0, structureGen: 0, clockGen: 0 }, createCfMemory(), 'cached')
    injectCfFormulaSelections(SELECTIONS)
    const reply = await withExactCfEvidence(CAPS, () => prepareCfRemote({ rects: [{ r0: 1, r1: 1, c0: 0, c1: 0 }], stamp: { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }, grantBytes: 1 << 22 },
      'sheet1', host, rules, createCfMemory(), { evaluation: UTC, workbook, sheetName: 'Sheet1' })).finally(() => injectCfFormulaSelections(null))
    expect(reply).toMatchObject({ kind: 'unresolved', reason: 'cf-stale-generation', states: [] })
    expect(reply.dependencies).toBeUndefined()
  })

  it('the snapshot layer refuses on its own: no host name, unknown generation, pending and failed referenced sheets', async () => {
    const { prepareCfDependencies } = await import('../src/xlsx/cf/dependencies')
    const { createCfWorkBudget } = await import('../src/xlsx/cf/limits')
    const { s1, s2 } = crossSheetDoc()
    const host = residentSource(s1, { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }, createCfMemory(), 'cached')
    const atoms = [{ sheet: 'Sheet2', rect: rectOf('A2') }]
    const sheetWith = (over: Partial<CfWorkbookSheet>): CfWorkbookSource => {
      const base = cfWorkbookOfDocument([s1, s2], createCfMemory(), 'cached')
      return { sheetIdOf: n => base.sheetIdOf(n), sheet: id => { const s = base.sheet(id); return s && { ...s, ...over } } }
    }
    const snap = async (workbook: CfWorkbookSource | undefined, sheetName: string | undefined) => {
      const signal = new AbortController().signal
      const out = await prepareCfDependencies(atoms, host, workbook, createCfWorkBudget({ signal, scan: true, maxUnits: 100_000 }), createCfMemory(), signal, { sheetId: 'sheet1', ...(sheetName ? { sheetName } : {}) })
      if (out.kind === 'ready') out.value.value.dispose()
      return out.kind === 'ready' ? 'ready' : out
    }
    expect(await snap(sheetWith({}), 'Sheet1')).toBe('ready')
    expect(await snap(sheetWith({}), undefined)).toEqual({ kind: 'unresolved', reason: 'cf-formula-dependency' })
    expect(await snap(undefined, 'Sheet1')).toEqual({ kind: 'unresolved', reason: 'cf-formula-dependency' })
    expect(await snap(sheetWith({ valueGen: () => undefined }), 'Sheet1')).toEqual({ kind: 'unresolved', reason: 'cf-dependency-generation-unknown' })
    expect(await snap(sheetWith({ state: () => 'pending' }), 'Sheet1')).toEqual({ kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: ['sheet2'] })
    expect(await snap(sheetWith({ state: () => 'failed' }), 'Sheet1')).toEqual({ kind: 'unresolved', reason: 'cf-read-failed' })
  })
})

/* ------------------------------------------------------------------ */
/* Fix round 1 (codex review of Task 36).                              */
/* ------------------------------------------------------------------ */
describe('Task 36 fix round 1', () => {
  const remoteRead = async (rules: CfSheetRules, s1: XlsxSheet, workbook: CfWorkbookSource, refs: readonly string[], grantBytes = 1 << 22, host?: CfValueSource) => {
    const h = host ?? residentSource(s1, { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: cfRulesVersion(rules), viewGen: 0, structureGen: 0, clockGen: 0 }, createCfMemory(), 'cached')
    injectCfFormulaSelections(SELECTIONS)
    try {
      return await withExactCfEvidence(CAPS, () => prepareCfRemote({ rects: refs.map(ref => { const { row, col } = rc(ref); return { r0: row, r1: row, c0: col, c1: col } }),
        stamp: { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }, grantBytes }, 'sheet1', h, rules, createCfMemory(), { evaluation: UTC, workbook, sheetName: 'Sheet1' }))
    } finally { injectCfFormulaSelections(null) }
  }

  it('#3: sheets literally named "missing" and "unknown" are ordinary sheets, never the internal sentinels', async () => {
    for (const name of ['missing', 'unknown']) {
      const rules = parseRules(expr('A2', 1, `ISERROR(${name}!A1)`), CAPS)
      const s1 = named('Sheet1', 'sheet1', { A2: 1 }, rules)
      const wb = cfWorkbookOfDocument([s1, named(name, 'sheet2', { A1: 7 })], createCfMemory(), 'cached')
      expect(fillOf((await prepareWith(rules, s1, wb, ['A2'])).out.get('A2')), name).toBe('none')
      // The same formula over a sheet whose A1 is an error does match: the reference was read, not replaced.
      const err = cfWorkbookOfDocument([s1, named(name, 'sheet2', { A1: { err: '#DIV/0!' } })], createCfMemory(), 'cached')
      expect(fillOf((await prepareWith(rules, s1, err, ['A2'])).out.get('A2')), name).toBe(GREEN)
    }
  })

  it('#4: print (the asset path without a host store) and export agree on Sheet2!A1>0', async () => {
    const rules = parseRules(expr('A1', 1, 'Sheet2!A1&gt;0'), CAPS)
    const s1 = named('Sheet1', 'sheet1', { A1: 1 }, rules), s2 = named('Sheet2', 'sheet2', { A1: 5 })
    const doc = { ...docOf([s1, s2]), valuePolicy: 'cached' } as XlsxDocument
    const { prepareSheetAssets } = await import('../src/components/xlsx/sheetAssets')
    injectCfFormulaSelections(SELECTIONS)
    try {
      await withExactCfEvidence(CAPS, async () => {
        const lease = await prepareSheetAssets(doc, { sheetIndex: 0 })
        const exported = await prepareCfPaint(s1, { valuePolicy: 'cached', evaluation: UTC, documentSheets: doc.sheets })
        expect(exported!.layer.styleAt(0, 0, undefined)?.fillColor).toBe(GREEN)
        expect(lease.assets.conditionalFormatting?.layer.styleAt(0, 0, undefined)?.fillColor).toBe(GREEN)
        exported!.dispose(); lease.dispose()
      })
    } finally { injectCfFormulaSelections(null) }
  })

  it('#5: one persistent service: prepared while Sheet2 is pending, prepared again once it is ready, paints', async () => {
    const { rules, s1, s2 } = crossSheetDoc()
    const state = { now: 'pending' as 'ready' | 'pending' | 'failed' }
    const sheet2 = drivenSheet('Sheet2', 'sheet2', s2, state)
    const workbook: CfWorkbookSource = { sheetIdOf: n => (n.toLowerCase() === 'sheet2' ? 'sheet2' : undefined), sheet: id => (id === 'sheet2' ? sheet2 : undefined) }
    const memory = createCfMemory()
    const service = createCfService(rules, memory, { evidence: evidence(CAPS), evaluation: evaluation(), sheetName: 'Sheet1', workbook })
    const stamp: CfStamp = { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
    const once = async () => {
      injectCfFormulaSelections(SELECTIONS)
      try {
        const lease = await service.prepareCf({ stamp, sourceRects: [rectOf('A2')], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, residentSource(s1, stamp, memory, 'cached'))
        const out = at(lease, 'A2'); lease.release(); return out
      } finally { injectCfFormulaSelections(null) }
    }
    expect((await once()).kind).toBe('unresolved')
    state.now = 'ready'
    expect(fillOf(await once())).toBe(GREEN)
    service.dispose()
  })

  it('#7: the 16-sheet cap bounds resolution, stamps and the reply: 17 rules over 17 sheets list 16 dependencies', async () => {
    const names = Array.from({ length: CF_REFERENCED_SHEETS_LIMIT + 1 }, (_, i) => `S${i + 2}`)
    const rules = parseRules(names.map((n, i) => expr('A2', i + 1, `${n}!A1&gt;0`)).join(''), CAPS)
    const s1 = named('Sheet1', 'sheet1', { A2: 1 }, rules)
    const base = cfWorkbookOfDocument([s1, ...names.map((n, i) => named(n, `x${i}`, { A1: 1 }))], createCfMemory(), 'cached')
    const resolved = new Set<string>()
    const workbook: CfWorkbookSource = { sheetIdOf: n => base.sheetIdOf(n), sheet: id => { resolved.add(id); return base.sheet(id) } }
    expect(cfDependencyIdentities(rules, 'Sheet1', workbook)).toHaveLength(CF_REFERENCED_SHEETS_LIMIT)
    const reply = await remoteRead(rules, s1, workbook, ['A2'])
    expect(reply.kind).toBe('ready')
    expect(reply.dependencies).toHaveLength(CF_REFERENCED_SHEETS_LIMIT)
    expect(reply.dependencies!.map(d => d.sheetId)).not.toContain(`x${CF_REFERENCED_SHEETS_LIMIT}`)
    expect(resolved.has(`x${CF_REFERENCED_SHEETS_LIMIT}`)).toBe(false)
  })

  it('#8: SUM(Sheet2!$A$1:$A$1000)>0 is hoisted through Sheet2\'s own store, invalidated by its generation, under the scan budget', async () => {
    const caps = [...CAPS, 'formula:hoist-aggregate', 'formula:hoist-countif-frequency']
    const rules = parseRules(expr('B1', 1, 'SUM(Sheet2!$A$1:$A$1000)&gt;0'), caps)
    const s1 = named('Sheet1', 'sheet1', { B1: 1 }, rules)
    // 200 populated cells: per cell this SUM would exceed the 128-unit cap, so only a hoist can paint it.
    const cells: Record<string, number> = { A500: 3 }
    for (let r = 1; r <= 200; r++) cells[`A${r}`] = r === 1 ? -1 : 0
    const doc = docOf([s1, named('Sheet2', 'sheet2', cells)])
    const wb = createObjectWorkbookSource(doc, 'hoist-cross')
    const run = async (c: readonly string[] = caps) => { const leg = await prepareStoreLeg(rules, wb, ['B1'], c); const out = fillOf(at(leg.lease, 'B1')); return { out, leg } }
    const unhoisted = await run(CAPS)
    expect(unhoisted.out).toBe('withheld')
    unhoisted.leg.lease.release(); unhoisted.leg.service.dispose()
    const first = await run()
    expect(first.out).toBe(GREEN)
    commit(wb.store('sheet2'), overlayOf({ A500: -3 }))
    expect(at(first.leg.lease, 'B1')).toEqual({ kind: 'unresolved', reason: 'cf-stale-generation' })
    first.leg.lease.release(); first.leg.service.dispose()
    const second = await run()
    expect(second.out).toBe('none')
    second.leg.lease.release(); second.leg.service.dispose()
    // The dependency scan pays from the scan budget: a starved one refuses the hoist (never paints).
    const { createCfWorkBudget } = await import('../src/xlsx/cf/limits')
    const memory = createCfMemory()
    const workbook = cfWorkbookOfStores(wb.sheets, id => wb.store(id), memory)
    const host = wb.store('sheet1')
    const service = createCfService(rules, memory, { evidence: evidence(caps), evaluation: evaluation(), sheetName: 'Sheet1', workbook, statisticsBudget: () => createCfWorkBudget({ scan: true, maxUnits: 1 }) })
    const stamp: CfStamp = { owner: host.stamp.docId, sheetId: 'sheet1', workbookValueGen: host.stamp.valueGen, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0, dependencyGen: cfDependencyGenOf(rules, 'Sheet1', workbook) }
    injectCfFormulaSelections(SELECTIONS)
    try {
      const lease = await service.prepareCf({ stamp, sourceRects: [rectOf('B1')], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, storeCfSource(host, stamp, memory))
      expect(shown(at(lease, 'B1')).withheld || at(lease, 'B1').kind === 'unresolved').toBe(true)
      lease.release()
    } finally { injectCfFormulaSelections(null); service.dispose() }
  })

  it('(a) at grant - 1 the cross-sheet states and dependencies are omitted, base cells return, and admitReply keeps the reply', async () => {
    const names = Array.from({ length: CF_REFERENCED_SHEETS_LIMIT }, (_, i) => `S${i + 2}`)
    const plain = '<conditionalFormatting sqref="B2"><cfRule type="cellIs" priority="2" dxfId="0" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>'
    const rules = parseRules(expr('A2', 1, `${names.map(n => `${n}!A1`).join('+')}&gt;0`) + plain, CAPS)
    const s1 = named('Sheet1', 'sheet1', { A2: 1, B2: 1 }, rules)
    const workbook = cfWorkbookOfDocument([s1, ...names.map((n, i) => named(n, `x${i}`, { A1: 1 }))], createCfMemory(), 'cached')
    const full = await remoteRead(rules, s1, workbook, ['A2', 'B2'])
    expect(full.kind).toBe('ready')
    expect(wireFill(full, 'A2')).toBe(GREEN)
    expect(wireFill(full, 'B2')).toBe('FFFFC7CE')
    expect(full.dependencies).toHaveLength(16)
    const fullAt = async (g: number) => { const r = await remoteRead(rules, s1, workbook, ['A2', 'B2'], g); return r.kind === 'ready' && r.dependencies !== undefined }
    let lo = 1, hi = 1 << 22
    while (lo < hi) { const mid = (lo + hi) >>> 1; if (await fullAt(mid)) hi = mid; else lo = mid + 1 }
    const { CF_REPLY_ENVELOPE_BYTES } = await import('../src/xlsx/view/contracts')
    expect(lo).toBe(measurePayloadBytes(full, () => true)! + CF_REPLY_ENVELOPE_BYTES)
    const tight = await remoteRead(rules, s1, workbook, ['A2', 'B2'], lo - 1)
    expect(tight.kind).toBe('ready')
    expect(tight.dependencies).toBeUndefined()
    expect(wireFill(tight, 'A2')).toBe('none')
    expect(wireFill(tight, 'B2')).toBe('FFFFC7CE')
    expect(tight.diagnostics.some(d => d.feature === 'cf-grant-required')).toBe(true)
    expect(measurePayloadBytes(tight, () => true)! + CF_REPLY_ENVELOPE_BYTES).toBeLessThanOrEqual(lo - 1)
    const { admitReply } = await import('../src/worker/viewerReplyLimit')
    for (const [grant, result] of [[lo - 1, tight], [lo, full]] as const) {
      const owner = { docId: 'o', sheetId: 'sheet1', valueGen: 0, fontGen: 0, viewGen: 0 }
      const cmd = { kind: 'read', requestId: 'g', owner, query: { op: 'conditionalFormatting', rects: [], stamp: result.stamp, grantBytes: grant } } as never
      const reply = { commandKind: 'read', requestId: 'g', owner, diagnostics: [], kind: 'complete', result: { kind: 'read', result: { op: 'conditionalFormatting', result } } } as never
      const out = admitReply(cmd, reply, grant, () => true, false) as { kind: string; replyBytes?: number }
      expect(out.kind, `grant ${grant}`).toBe('complete')
      expect(out.replyBytes!).toBeLessThanOrEqual(grant)
    }
  })

  it('(b) a dependency commit queued as a microtask at snapshot adoption is caught by the worker liveAfter recheck', async () => {
    const { rules, s1, s2 } = crossSheetDoc()
    let gen = 0, armed = false
    const base = cfWorkbookOfDocument([s1, s2], createCfMemory(), 'cached').sheet('sheet2')!
    const sheet2: CfWorkbookSheet = { ...base, valueGen: () => { if (armed) { armed = false; queueMicrotask(() => { gen = 1 }) } return gen } }
    const workbook: CfWorkbookSource = { sheetIdOf: n => (n.toLowerCase() === 'sheet2' ? 'sheet2' : undefined), sheet: id => (id === 'sheet2' ? sheet2 : undefined) }
    const inner = residentSource(s1, { owner: 'o', sheetId: 'sheet1', workbookValueGen: 0, cfGen: cfRulesVersion(rules), viewGen: 0, structureGen: 0, clockGen: 0 }, createCfMemory(), 'cached')
    // The host's own page loop is the service's last read: the next stamp read is the adoption check.
    const host: CfValueSource = { ...inner, pages: (id, rects, budget, signal) => {
      const it = inner.pages(id, rects, budget, signal)
      return (async function* () { for await (const p of it) yield p; armed = true })()
    } }
    const reply = await remoteRead(rules, s1, workbook, ['A2'], 1 << 22, host)
    expect(gen).toBe(1)
    expect(reply).toMatchObject({ kind: 'unresolved', reason: 'cf-stale-generation', states: [] })
  })
})

describe('Task 36 fix round 2', () => {
  it('#1: a hoisted COUNTIF/SUMIF over Sheet2 never reads its unqualified criterion from Sheet2', async () => {
    const caps = [...CAPS, 'formula:hoist-aggregate', 'formula:hoist-countif-frequency']
    for (const formula of ['COUNTIF(Sheet2!$A$1:$A$2,$B$1)&gt;0', 'SUMIF(Sheet2!$A$1:$A$2,$B$1)&gt;0']) {
      const rules = parseRules(expr('C1', 1, formula), caps)
      const s1 = named('Sheet1', 'sheet1', { B1: 9, C1: 1 }, rules)
      const wb = cfWorkbookOfDocument([s1, named('Sheet2', 'sheet2', { A1: 1, A2: 2, B1: 1 })], createCfMemory(), 'cached')
      expect(fillOf((await prepareWith(rules, s1, wb, ['C1'], CAPS)).out.get('C1')), `${formula} unhoisted`).toBe('none')
      expect(fillOf((await prepareWith(rules, s1, wb, ['C1'], caps)).out.get('C1')), `${formula} hoisted`).toBe('none')
    }
  })

  it('#7 residual: the resident live key reads only the admitted, deduplicated 16 referenced stores', async () => {
    const { cfResidentDependencyKey } = await import('../src/xlsx/cf/render')
    // Round 3 (#2): the helper lives in the lazy CF runtime, never in the eager loader.
    expect('cfResidentDependencyKey' in (await import('../src/xlsx/cfLoader'))).toBe(false)
    const names = Array.from({ length: CF_REFERENCED_SHEETS_LIMIT + 1 }, (_, i) => `S${i + 2}`)
    const rules = parseRules([...names, 'S2', 'sheet1'].map((n, i) => expr('A2', i + 1, `${n}!A1&gt;0`)).join(''), CAPS)
    const sheets = [{ sheetId: 'h', name: 'Sheet1' }, ...names.map((n, i) => ({ sheetId: `x${i}`, name: n }))]
    let reads = 0
    const key = cfResidentDependencyKey(rules, 'Sheet1', sheets, id => { reads++; return { stamp: { valueGen: id === 'x0' ? 5 : 0 }, completeness: 'complete' } })
    expect(reads).toBe(CF_REFERENCED_SHEETS_LIMIT)
    expect(key.startsWith('5complete')).toBe(true)
  })
})
