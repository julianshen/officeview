/**
 * Task 38 cross-view harness (Ruling S5-P): one workbook, every leg. Each leg prepares the SAME sheet through its own
 * production route (static `prepareCfPaint`, resident store `prepareCfPaint`, the text scope `prepareCfTextScope`, the
 * worker's `readConditionalFormatting` and the forced-paged wire scope `prepareCfWireScope`) and reports what a viewer would
 * show per cell: the fill, strikethrough and number format the cell resolves to, and the accessible descriptor. Nothing here
 * decides an expectation: the tests compare legs with each other and with a stated literal twin.
 */
import { createCanvas } from 'canvas'
import { readFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { prepareCfPaint } from '../../src/xlsx/cf/render'
import { prepareCfTextScope, prepareCfWireScope, type CfFrameHostSource } from '../../src/xlsx/cf/text'
import { cfDocumentEvaluationOf, cfEvaluationContextFor, type CfDocumentEvaluation } from '../../src/xlsx/cf/formula'
import { createCfMemory } from '../../src/xlsx/cf/limits'
import { withInjectedCfEvidence, injectCfFormulaSelections } from '../../src/xlsx/cf/evidence'
import { CF_FORMULA_PRODUCTION_READINGS } from '../../src/xlsx/cf/formulaReadings'
import { cfRulesVersion, createPagedWorkbookSource } from '../../src/xlsx/view/pagedSource'
import { openImportSession } from '../../src/xlsx/storage/pagedImport'
import { createObjectWorkbookSource } from '../../src/xlsx/view/source'
import type { CfReadResult } from '../../src/xlsx/view/contracts'
import type { XlsxCellStyle, XlsxDocument, XlsxSheet } from '../../src/xlsx/types'
import type { ValuePolicy } from '../../src/xlsx/view/contracts'
import { xlsxBytes } from './xlsx-package'

export const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..')
/** 2026-11-16 12:00 UTC: the injected clock of every leg. */
export const NOW = Date.UTC(2026, 10, 16, 12)
/** The document policy of the synthetic workbooks (the tests never read the ambient zone). */
export const UTC_AUTO: CfDocumentEvaluation = Object.freeze({ calcMode: 'auto', valuePolicy: 'cached', dateSystem: '1900', timeZone: 'UTC' })

/** Every Stage 5 reading production registers, as additive evidence (what each leg is compared under: real paint). */
export const REGISTERED = Object.keys(CF_FORMULA_PRODUCTION_READINGS)
export function firstSelections(extra: Readonly<Record<string, string>> = {}): Record<string, string> {
  const out: Record<string, string> = {}
  for (const [k, v] of Object.entries(CF_FORMULA_PRODUCTION_READINGS)) if (typeof v !== 'string') out[k] = v[0]!
  return { ...out, ...extra }
}
/** Runs `fn` under additive evidence (production's plus `caps`) and the given reading selections. */
export async function withPaint<T>(fn: () => Promise<T>, caps: readonly string[] = REGISTERED, selections: Readonly<Record<string, string>> = firstSelections()): Promise<T> {
  return withInjectedCfEvidence(caps, async () => {
    injectCfFormulaSelections(selections)
    try { return await fn() } finally { injectCfFormulaSelections(null) }
  })
}

/* ------------------------------ synthetic workbooks ------------------------------ */

export const RED = 'FFFFC7CE', GREEN = 'FFC6EFCE'
/** dxf 0 red fill, 1 green fill, 2 red fill with number format 0.00. */
export const DXFS = `<dxfs count="3"><dxf><fill><patternFill patternType="solid"><bgColor rgb="${RED}"/></patternFill></fill></dxf>`
  + `<dxf><fill><patternFill patternType="solid"><bgColor rgb="${GREEN}"/></patternFill></fill></dxf>`
  + `<dxf><numFmt numFmtId="164" formatCode="0.00"/><fill><patternFill patternType="solid"><bgColor rgb="${RED}"/></patternFill></fill></dxf></dxfs>`
/** xf 0 General, xf 1 the built-in date format (numFmtId 14): the cell kinds a `timePeriod` rule reads as dates. */
const STYLES = `<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs>${DXFS}`

export type CellSpec = number | string | boolean | { f: string; v?: number } | { date: number }
/** `rows`: A1 notation to a value; a `{ f }` is a formula (cached `v` only when stated: an uncached formula otherwise). */
export function sheetData(cells: Readonly<Record<string, CellSpec>>, rowHt?: number): string {
  const byRow = new Map<number, Array<[string, CellSpec]>>()
  for (const [ref, v] of Object.entries(cells)) { const r = Number(/\d+/.exec(ref)![0]); byRow.set(r, [...(byRow.get(r) ?? []), [ref, v]]) }
  const colIndex = (ref: string): number => [...ref.replace(/\d+/g, '')].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)
  return [...byRow].sort((a, b) => a[0] - b[0]).map(([r, list]) => `<row r="${r}"${rowHt ? ` ht="${rowHt}" customHeight="1"` : ''}>${list.sort((a, b) => colIndex(a[0]) - colIndex(b[0])).map(([ref, v]) => {
    if (typeof v === 'number') return `<c r="${ref}"><v>${v}</v></c>`
    if (typeof v === 'boolean') return `<c r="${ref}" t="b"><v>${v ? 1 : 0}</v></c>`
    if (typeof v === 'string') return `<c r="${ref}" t="inlineStr"><is><t>${v}</t></is></c>`
    if ('date' in v) return `<c r="${ref}" s="1"><v>${v.date}</v></c>`
    return `<c r="${ref}"><f>${v.f}</f>${v.v === undefined ? '' : `<v>${v.v}</v>`}</c>`
  }).join('')}</row>`).join('')
}
export const cfBlock = (sqref: string, rules: string): string => `<conditionalFormatting sqref="${sqref}">${rules}</conditionalFormatting>`
export const exprRule = (formula: string, priority: number, dxfId: number): string => `<cfRule type="expression" priority="${priority}" dxfId="${dxfId}"><formula>${formula}</formula></cfRule>`
export const cellIsRule = (op: string, formula: string, priority: number, dxfId: number): string => `<cfRule type="cellIs" operator="${op}" priority="${priority}" dxfId="${dxfId}"><formula>${formula}</formula></cfRule>`
/** `rowHt`: every row's authored height in pt (an icon's frame is decided only on an exact row height, S4-ROW-EXACT). */
export const workbookBytes = (cells: Readonly<Record<string, CellSpec>>, cf: string, rowHt?: number): Uint8Array => xlsxBytes({ sheetData: sheetData(cells, rowHt), afterSheetData: cf, styles: STYLES })

/* ------------------------------ loading ------------------------------ */

export interface Loaded { doc: XlsxDocument; sheet: XlsxSheet; bytes: Uint8Array; sheetId: string }
export async function load(bytes: Uint8Array, valuePolicy: ValuePolicy = 'cached', sheetName?: string): Promise<Loaded> {
  const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy })
  const sheet = doc.sheets.find(s => s.name === sheetName) ?? doc.sheets[0]!
  return { doc, sheet, bytes, sheetId: sheet.sheetId ?? sheet.name }
}
export const fixtureBytes = (id: string): Uint8Array => new Uint8Array(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf', `${id}.xlsx`)))
/** The authored input of a fixture (what Excel opens before any calculate): `stage5-calc-manual` keeps its stale saved values here, while
 * the saved file is Excel's resave with them recomputed. */
export const inputBytes = (id: string): Uint8Array => new Uint8Array(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/inputs', `${id}.xlsx`)))

/* ------------------------------ observations ------------------------------ */

export interface Obs { fill?: string; strike?: boolean; fmt?: string; desc?: string }
export const rcOf = (ref: string): { row: number; col: number } => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!
  return { row: Number(m[2]) - 1, col: [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 }
}
const normFill = (c: string | undefined): string | undefined => (c === undefined ? undefined : c.replace('#', '').toUpperCase())
const obsOfStyle = (s: XlsxCellStyle | undefined, desc?: { key: string; text: string }): Obs => {
  const o: Obs = {}
  if (s?.fillColor !== undefined) o.fill = normFill(s.fillColor)
  if (s?.strike) o.strike = true
  if (s?.formatCode !== undefined && s.formatCode !== 'General') o.fmt = s.formatCode
  if (desc) o.desc = `${desc.key}|${desc.text}`
  return o
}
export interface LegOptions {
  /** The document policy the route is given; `false` withholds it (a route called without the context). */
  evaluation?: CfDocumentEvaluation | false
  /** Overrides the route's value policy (default: the route's own default, as a direct caller would get it). */
  valuePolicy?: ValuePolicy
  now?: () => number
  /** An explicit predicate date context (a caller's own clock). */
  dateContext?: import('../../src/xlsx/cf/predicates').CfDateContext
  /** Hosts for bar and icon announcements (the same on every scope leg). */
  hosts?: CfFrameHostSource
  /** The other sheets of the document: absent, a cross-sheet rule has no workbook. */
  workbook?: boolean
}
const policyOf = (loaded: Loaded, o: LegOptions): CfDocumentEvaluation | undefined => o.evaluation === false ? undefined : o.evaluation ?? cfDocumentEvaluationOf(loaded.doc)
const now = (o: LegOptions): (() => number) => o.now ?? (() => NOW)
export const frameHosts = (width = 200, base?: XlsxCellStyle): CfFrameHostSource => ({ cell: (row, col) => ({ box: { x: col * width, y: row * 20, width, height: 20 }, base }) })

export async function staticLeg(l: Loaded, refs: readonly string[], o: LegOptions = {}): Promise<Map<string, Obs>> {
  const batch = await prepareCfPaint(l.sheet, { memory: createCfMemory(), ...(o.valuePolicy ? { valuePolicy: o.valuePolicy } : {}), ...(policyOf(l, o) ? { evaluation: policyOf(l, o)! } : {}), now: now(o), ...(o.dateContext ? { dateContext: o.dateContext } : {}), ...(o.workbook === false ? {} : { documentSheets: l.doc.sheets }) })
  const out = new Map<string, Obs>()
  try { for (const r of refs) { const { row, col } = rcOf(r); out.set(r, obsOfStyle(batch?.layer.styleAt(row, col, undefined))) } } finally { batch?.dispose() }
  return out
}
export async function residentLeg(l: Loaded, refs: readonly string[], o: LegOptions = {}): Promise<Map<string, Obs>> {
  const source = createObjectWorkbookSource(l.doc, `resident-${Math.random()}`)
  const batch = await prepareCfPaint(l.sheet, { memory: createCfMemory(), store: source.store(l.sheetId), ...(o.workbook === false ? {} : { workbook: source }), ...(o.valuePolicy ? { valuePolicy: o.valuePolicy } : {}), ...(policyOf(l, o) ? { evaluation: policyOf(l, o)! } : {}), now: now(o), ...(o.dateContext ? { dateContext: o.dateContext } : {}) })
  const out = new Map<string, Obs>()
  try { for (const r of refs) { const { row, col } = rcOf(r); out.set(r, obsOfStyle(batch?.layer.styleAt(row, col, undefined))) } } finally { batch?.dispose(); await source.release() }
  return out
}
const BUDGET = { reserve() {}, release() {} }
const rectOf = (refs: readonly string[]) => refs.map(r => { const { row, col } = rcOf(r); return { r0: row, r1: row, c0: col, c1: col } })
/** The text scope (copy, find, a11y) over the static inputs (`resident: false`) or the resident store. */
export async function textLeg(l: Loaded, refs: readonly string[], o: LegOptions & { resident?: boolean } = {}): Promise<Map<string, Obs>> {
  const source = o.resident ? createObjectWorkbookSource(l.doc, `text-${Math.random()}`) : undefined
  const scope = await prepareCfTextScope(l.sheet, {
    rects: rectOf(refs), signal: new AbortController().signal, budget: BUDGET,
    ...(source ? { store: source.store(l.sheetId), ...(o.workbook === false ? {} : { workbook: source }) } : o.workbook === false ? {} : { documentSheets: l.doc.sheets }),
    ...(o.valuePolicy ? { valuePolicy: o.valuePolicy } : {}), ...(policyOf(l, o) ? { evaluation: policyOf(l, o)! } : {}), now: now(o), ...(o.dateContext ? { dateContext: o.dateContext } : {}), ...(o.hosts ? { hosts: o.hosts, layoutContext: () => createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D } : {}),
  } as never)
  const out = new Map<string, Obs>()
  try { for (const r of refs) { const { row, col } = rcOf(r); out.set(r, scope ? obsOfStyle(scope.effectiveStyle(row, col, undefined), scope.describe(row, col)) : {}) } } finally { scope?.dispose(); await source?.release() }
  return out
}
/** The worker route over the same bytes: the wire states and the forced-paged scope that reads them. */
export async function workerLeg(l: Loaded, refs: readonly string[], o: LegOptions = {}): Promise<{ wire: CfReadResult; states: Map<string, { properties: Record<string, unknown>; unresolvedReason?: string; unresolvedWrites?: readonly string[] }>; scope: Map<string, Obs> }> {
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(l.bytes)]) }, { docId: `worker-${Math.random()}`, profile: 'D' })
  const source = createPagedWorkbookSource(session)
  try {
    await session.done
    const sheetId = source.sheets.find(s => s.name === l.sheet.name)?.sheetId ?? source.sheets[0]!.sheetId
    const store = source.store(sheetId)
    const rects = rectOf(refs)
    const read = (rs: readonly { r0: number; r1: number; c0: number; c1: number }[]) => store.readConditionalFormatting!({ rects: rs, stamp: { owner: 'views', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }, grantBytes: 8 * 1024 * 1024 }, createCfMemory(), { signal: new AbortController().signal, now: now(o), ...(o.dateContext ? { dateContext: o.dateContext } : {}) } as never)
    const wire = await read(rects)
    const states = new Map<string, { properties: Record<string, unknown>; unresolvedReason?: string; unresolvedWrites?: readonly string[] }>()
    if (wire.kind === 'ready') for (const s of wire.states) { states.set(refOfRc(s.row, s.col), { properties: s.properties, ...(s.unresolvedReason ? { unresolvedReason: s.unresolvedReason } : {}), ...(s.unresolvedWrites ? { unresolvedWrites: s.unresolvedWrites } : {}) }) }
    const scope = await prepareCfWireScope(async rs => ({ result: structuredClone(await read(rs as never)), release() {} }), rects, { budget: BUDGET, ...(o.hosts ? { hosts: o.hosts, rowPtOf: () => 15, rowUniformOf: () => true } : {}) } as never)
    const out = new Map<string, Obs>()
    try { for (const r of refs) { const { row, col } = rcOf(r); out.set(r, scope ? obsOfStyle(scope.effectiveStyle(row, col, undefined), scope.describe(row, col)) : {}) } } finally { scope?.dispose() }
    return { wire, states, scope: out }
  } finally { await source.release() }
}
export function refOfRc(row: number, col: number): string {
  let c = col + 1, s = ''
  while (c > 0) { const r = (c - 1) % 26; s = String.fromCharCode(65 + r) + s; c = Math.floor((c - 1) / 26) }
  return `${s}${row + 1}`
}
void cfEvaluationContextFor; void cfRulesVersion

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', XM = 'http://schemas.microsoft.com/office/excel/2006/main'
const BAR_GUID = '{00000000-0000-4000-8000-0000000000a1}'
/** B: a data bar whose max is `$F$1`; C: a colour scale whose max is `$F$1`; D: a 3Arrows icon set whose middle threshold is `$F$1`. */
export const formulaBoundsBytes = (): Uint8Array => {
  const data = [0, 5, 10, 20, 30, 40, 50]
  const cells: Record<string, number> = { F1: 50 }
  data.forEach((v, i) => { cells[`B${i + 2}`] = v; cells[`C${i + 2}`] = v; cells[`D${i + 2}`] = v * 2 })
  const bar = `<conditionalFormatting sqref="B2:B8"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="formula" val="$F$1"/><color rgb="FF0000FF"/></dataBar>`
    + `<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${X14}"><x14:id>${BAR_GUID}</x14:id></ext></extLst></cfRule></conditionalFormatting>`
  const scale = `<conditionalFormatting sqref="C2:C8"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="formula" val="$F$1"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>`
  const icon = `<conditionalFormatting sqref="D2:D8"><cfRule type="iconSet" priority="3"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="formula" val="$F$1"/><cfvo type="num" val="80"/></iconSet></cfRule></conditionalFormatting>`
  const x14 = `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${X14}" xmlns:xm="${XM}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="${BAR_GUID}">`
    + `<x14:dataBar minLength="0" maxLength="100" gradient="0"><x14:cfvo type="min"/><x14:cfvo type="formula"><xm:f>$F$1</xm:f></x14:cfvo><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar>`
    + `</x14:cfRule><xm:sqref>B2:B8</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`
  return workbookBytes(cells, bar + scale + icon + x14, 15)
}
