/**
 * Shared loaders for the Stage 2 pivot engine tests: an xlsx file (bytes) → each pivot's definition, cache, sheet and,
 * when its records part exists, the decoded record store. Used by the axes, values and faithfulness tests.
 */
import { readFileSync } from 'node:fs'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { loadPivotWorkbook } from '../../src/xlsx/pivots/load'
import { decodeRecords, RECORD_LIMITS, type RecordStore } from '../../src/xlsx/pivots/engine/records'
import type { PivotCellValue } from '../../src/xlsx/pivots/engine/place'
import type { PivotCacheDefinition, PivotTableDefinition } from '../../src/xlsx/pivots/model'
import type { XlsxSheet } from '../../src/xlsx/types'

/** A pivot as parsed, before any record decode (the §7.6 refusals have no records part to decode). */
export interface PivotParts { file: string; sheet: XlsxSheet; def: PivotTableDefinition; cache: PivotCacheDefinition | undefined; date1904: boolean; pkg: OfficePackage }
export interface Loaded { file: string; sheet: XlsxSheet; def: PivotTableDefinition; cache: PivotCacheDefinition; store: RecordStore; date1904: boolean }

export async function* once(bytes: Uint8Array): AsyncIterable<Uint8Array> { yield bytes }

/** Every pivot of a workbook, parsed but not decoded. */
export async function loadPivotParts(file: string, bytes: Uint8Array = new Uint8Array(readFileSync(file))): Promise<PivotParts[]> {
  const pkg = await OfficePackage.load(bytes)
  const doc = await parseXlsx(pkg)
  const regions = new Map(doc.sheets.filter(sheet => sheet.pivotRegions?.length).map(sheet => [sheet.name, sheet.pivotRegions!]))
  const workbook = await loadPivotWorkbook(pkg, regions)
  const date1904 = doc.sheets[0]?.semantics?.dateSystem === '1904'
  return workbook.tables.map(def => ({ file, sheet: doc.sheets.find(sheet => sheet.name === def.sheet)!, def, cache: workbook.caches.get(def.cacheId), date1904, pkg }))
}

/** Decode a pivot's records part (throws on a decode failure: the callers' fixtures all decode). */
export async function decodeParts(parts: PivotParts): Promise<Loaded> {
  const { file, sheet, def, cache, date1904, pkg } = parts
  if (!cache?.recordsPartPath) throw new Error(`${file}: ${def.name} has no records part`)
  const bytes = await pkg.bytes(cache.recordsPartPath)
  const result = await decodeRecords(once(bytes!), cache, RECORD_LIMITS.D, date1904)
  if (!result.ok) throw new Error(`${file}: ${result.diagnostic.message}`)
  return { file, sheet, def, cache, store: result.store, date1904 }
}

/** Every pivot with a records part, decoded. `missingRecords: 'throw'` fails on a pivot without one; 'skip' leaves it out. */
export async function loadAll(file: string, missingRecords: 'throw' | 'skip' = 'throw'): Promise<Loaded[]> {
  const out: Loaded[] = []
  for (const parts of await loadPivotParts(file)) {
    if (!parts.cache?.recordsPartPath && missingRecords === 'skip') continue
    out.push(await decodeParts(parts))
  }
  return out
}

/** A saved sheet cell as a layout value ('empty' for no value). */
export function savedCell(sheet: Pick<XlsxSheet, 'rows'>, row: number, col: number): PivotCellValue {
  const cell = sheet.rows.find(r => r.index === row)?.cells.find(c => c.col === col)
  if (!cell || cell.value === null || cell.value === '') return { kind: 'empty' }
  if (cell.valueIsError) return { kind: 'error', value: String(cell.value) }
  if (typeof cell.value === 'number') return { kind: 'number', value: cell.value }
  return { kind: 'text', value: String(cell.value) }
}
