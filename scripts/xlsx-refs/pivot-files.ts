/**
 * File access shared by the pivot oracle scripts (build-engine.ts --interaction, classify-stage2.ts,
 * measure-interaction.ts): zip parts by name, a sheet's XML by sheet name, and a pivot parsed and decoded with the
 * viewer's own loaders (so indices and shapes are the engine's).
 */
import { readFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { loadPivotWorkbook } from '../../src/xlsx/pivots/load'
import { decodeRecords, RECORD_LIMITS, type RecordStore } from '../../src/xlsx/pivots/engine/records'
import type { PivotCacheDefinition, PivotTableDefinition } from '../../src/xlsx/pivots/model'
import type { XlsxSheet } from '../../src/xlsx/types'

export const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
export const PIVOT_DIR = join(ROOT, 'tests/fixtures/xlsx-pivots')
/** A fixture named relative to tests/fixtures/xlsx-pivots without its extension ('engine/baseline'). */
export const fixturePath = (rel: string) => join(PIVOT_DIR, `${rel}.xlsx`)

export const PIVOT_RE = /^xl\/pivotTables\/pivotTable\d+\.xml$/
export const CACHE_RE = /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/
export const RECORDS_RE = /^xl\/pivotCache\/pivotCacheRecords\d+\.xml$/
export const CHART_RE = /^xl\/charts\/chart\d+\.xml$/

export async function loadZip(file: string): Promise<JSZip> { return JSZip.loadAsync(readFileSync(file)) }
export const partNames = (zip: JSZip, re: RegExp) => Object.keys(zip.files).filter(n => re.test(n)).sort((a, b) => a.localeCompare(b, 'en', { numeric: true }))
export async function partText(zip: JSZip, name: string | undefined): Promise<string | undefined> {
  return name ? zip.file(name)?.async('string') : undefined
}

const attr = (tag: string, name: string) => new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1]
/** A sheet's part name, found by its name through workbook.xml and its relationships (sheet order is not file order). */
export async function sheetPart(zip: JSZip, sheet: string): Promise<string | undefined> {
  const workbook = await partText(zip, 'xl/workbook.xml'), rels = await partText(zip, 'xl/_rels/workbook.xml.rels')
  if (!workbook || !rels) return undefined
  const decoded = (v: string) => v.replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&apos;/g, "'")
  const tag = [...workbook.matchAll(/<sheet\b[^>]*>/g)].map(m => m[0]).find(t => decoded(attr(t, 'name') ?? '') === sheet)
  const id = tag && attr(tag, 'r:id')
  const rel = id && [...rels.matchAll(/<Relationship\b[^>]*>/g)].map(m => m[0]).find(t => attr(t, 'Id') === id)
  const target = rel && attr(rel, 'Target')
  if (!target) return undefined
  return target.startsWith('/') ? target.slice(1) : `xl/${target}`
}
export async function sheetXml(zip: JSZip, sheet: string): Promise<string | undefined> { return partText(zip, await sheetPart(zip, sheet)) }

/** Shared strings in order (rich text runs joined). */
export async function sharedStrings(zip: JSZip): Promise<string[]> {
  const xml = await partText(zip, 'xl/sharedStrings.xml')
  if (!xml) return []
  return [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m => [...m[1]!.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(t => unescape(t[1]!)).join(''))
}
const unescape = (v: string) => v.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&')

export interface SheetCell { ref: string; row: number; col: number; formula?: string; value?: string; type?: string }
/** Every `<c>` of a sheet with its formula and cached value (shared strings resolved). */
export function sheetCells(xml: string, strings: readonly string[]): SheetCell[] {
  const out: SheetCell[] = []
  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const ref = attr(m[1]!, 'r')
    if (!ref) continue
    const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref) ?? []
    const col = [...letters!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
    const type = attr(m[1]!, 't'), body = m[2] ?? ''
    const f = /<f\b[^>]*>([\s\S]*?)<\/f>/.exec(body)?.[1]
    const v = /<v>([\s\S]*?)<\/v>/.exec(body)?.[1] ?? /<is>[\s\S]*?<t\b[^>]*>([\s\S]*?)<\/t>/.exec(body)?.[1]
    const value = v === undefined ? undefined : type === 's' ? strings[Number(v)] : unescape(v)
    out.push({ ref, row: Number(digits) - 1, col, ...(f !== undefined ? { formula: unescape(f) } : {}), ...(value !== undefined ? { value } : {}), ...(type ? { type } : {}) })
  }
  return out
}

export interface ColWidth { min: number; max: number; width: number; customWidth: boolean; bestFit: boolean }
export function colWidths(xml: string): ColWidth[] {
  return [...xml.matchAll(/<col\b[^>]*>/g)].map(m => ({
    min: Number(attr(m[0], 'min')), max: Number(attr(m[0], 'max')), width: Number(attr(m[0], 'width')),
    customWidth: attr(m[0], 'customWidth') === '1', bestFit: attr(m[0], 'bestFit') === '1',
  }))
}
/** The sheet's default column width (`sheetFormatPr`), when written. */
export const defaultColWidth = (xml: string) => {
  const tag = /<sheetFormatPr\b[^>]*>/.exec(xml)?.[0]
  const v = tag && (attr(tag, 'defaultColWidth') ?? attr(tag, 'baseColWidth'))
  return v === undefined || v === null ? undefined : Number(v)
}

/** The first pivot of a workbook, parsed with the viewer's loaders, and its records decoded when present. */
export interface LoadedPivot { def: PivotTableDefinition; cache: PivotCacheDefinition; sheet: XlsxSheet; store?: RecordStore; date1904: boolean; storeError?: string }
export async function loadPivot(file: string): Promise<LoadedPivot> {
  const pkg = await OfficePackage.load(new Uint8Array(readFileSync(file)))
  const doc = await parseXlsx(pkg)
  const regions = new Map(doc.sheets.filter(sheet => sheet.pivotRegions?.length).map(sheet => [sheet.name, sheet.pivotRegions!]))
  const book = await loadPivotWorkbook(pkg, regions)
  const def = book.tables[0]
  if (!def) throw new Error(`${file}: no pivot table`)
  const cache = book.caches.get(def.cacheId)
  if (!cache) throw new Error(`${file}: no cache for ${def.name}`)
  const date1904 = doc.sheets[0]?.semantics?.dateSystem === '1904'
  const sheet = doc.sheets.find(s => s.name === def.sheet)!
  if (!cache.recordsPartPath) return { def, cache, sheet, date1904 }
  const bytes = await pkg.bytes(cache.recordsPartPath)
  async function* once(b: Uint8Array): AsyncIterable<Uint8Array> { yield b }
  const result = await decodeRecords(once(bytes!), cache, RECORD_LIMITS.D, date1904)
  return result.ok ? { def, cache, sheet, store: result.store, date1904 } : { def, cache, sheet, date1904, storeError: result.diagnostic.message }
}
