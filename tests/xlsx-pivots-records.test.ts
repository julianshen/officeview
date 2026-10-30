// @vitest-environment node
import { readFileSync, existsSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { bytesRangeSource, openPart, readZipInventory } from '../src/core/zip-stream'
import { tokenizeBytes } from '../src/core/xml-stream'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { loadPivotWorkbook } from '../src/xlsx/pivots/load'
import type { PivotCacheDefinition, PivotCacheField } from '../src/xlsx/pivots/model'
import { decodeRecords, isoToSerial, RECORD_LIMITS, recordEvents, VALUE_TAG } from '../src/xlsx/pivots/engine/records'

function field(name: string, inline: boolean, databaseField = true): PivotCacheField {
  return { name, serverField: false, uniqueList: true, sqlType: 0, hierarchy: 0, level: 0, databaseField, memberPropertyField: false,
    ...(databaseField ? { sharedItems: { containsSemiMixedTypes: false, containsNonDate: false, containsDate: false, containsString: !inline,
      containsBlank: false, containsMixedTypes: false, containsNumber: inline, containsInteger: false, longText: false, inline,
      items: inline ? [] : [{ type: 'string', value: 'first' }, { type: 'string', value: 'second' }] } } : {}) }
}
function cache(fields: PivotCacheField[], recordCount?: number): PivotCacheDefinition {
  return { cacheId: 1, partPath: 'xl/pivotCache/pivotCacheDefinition1.xml', source: { type: 'worksheet', connectionId: 0 }, fields,
    invalid: false, saveData: true, refreshOnLoad: false, optimizeMemory: false, enableRefresh: true, backgroundQuery: false,
    createdVersion: 0, refreshedVersion: 0, minRefreshableVersion: 0, upgradeOnRefresh: false, supportSubquery: false, supportAdvancedDrill: false, olap: false, unsupported: [], recordCount }
}
async function* chunks(text: string, size: number): AsyncIterable<Uint8Array> {
  const bytes = new TextEncoder().encode(text)
  for (let i = 0; i < bytes.length; i += size) yield bytes.subarray(i, i + size)
}
async function decode(text: string, fields: PivotCacheField[], limits = { maxBytes: 1e6, maxRecords: 1000, maxInlineStrings: 1000 }) {
  return decodeRecords(chunks(text, 17), cache(fields), limits, false)
}
async function corpusRecords(file: string) {
  const raw = new Uint8Array(readFileSync(resolve(process.cwd(), 'corpus', file)))
  const pkg = await OfficePackage.load(raw), doc = await parseXlsx(pkg)
  const regions = new Map(doc.sheets.filter(sheet => sheet.pivotRegions?.length).map(sheet => [sheet.name, sheet.pivotRegions!]))
  const workbook = await loadPivotWorkbook(pkg, regions), definition = [...workbook.caches.values()][0]!
  const archive = await JSZip.loadAsync(raw), recordXml = await archive.file(definition.recordsPartPath!)!.async('string')
  return { definition, result: await decodeRecords(chunks(recordXml, 127), definition, RECORD_LIMITS.D, false) }
}

describe('pivot cache record store', () => {
  it('decodes shared and inline columns and omits calculated fields', async () => {
    const result = await decode('<pivotCacheRecords count="2"><r><x v="1"/><n v="12.5"/><s v="alpha"/></r><r><x v="0"/><m/><s v="beta"/></r></pivotCacheRecords>',
      [field('Shared', false), field('Calculated', true, false), field('Amount', true), field('Label', true)])
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.store.count).toBe(2)
    expect(result.store.columns[0]).toEqual({ kind: 'shared', index: new Uint32Array([1, 0]) })
    expect(result.store.columns[1]).toBeUndefined()
    const amount = result.store.columns[2]
    expect(amount).toMatchObject({ kind: 'inline', tag: new Uint8Array([VALUE_TAG.number, VALUE_TAG.missing]), num: new Float64Array([12.5, 0]) })
    expect(result.store.columns[3]).toMatchObject({ kind: 'inline', tag: new Uint8Array([VALUE_TAG.string, VALUE_TAG.string]), text: ['alpha', 'beta'], textIndex: new Uint32Array([0, 1]) })
  })

  it('is independent of one-byte chunk boundaries', async () => {
    const xml = '<pivotCacheRecords count="1"><r><s v="A&#x1F642;"/><n v="1.25"/></r></pivotCacheRecords>', fields = [field('S', true), field('N', true)]
    const a = await decodeRecords(chunks(xml, 1), cache(fields), { maxBytes: 1000, maxRecords: 10, maxInlineStrings: 10 }, false)
    const b = await decodeRecords(chunks(xml, 4096), cache(fields), { maxBytes: 1000, maxRecords: 10, maxInlineStrings: 10 }, false)
    expect(a).toEqual(b)
  })

  it('returns typed parse diagnostics for invalid shared indices, row widths and value kinds', async () => {
    const shared = field('K', false); shared.sharedItems!.items = [{ type: 'string', value: 'only' }]
    for (const xml of [
      '<pivotCacheRecords><r><x v="1"/></r></pivotCacheRecords>',
      '<pivotCacheRecords><r/></pivotCacheRecords>',
      '<pivotCacheRecords><r><s v="inline"/></r></pivotCacheRecords>',
    ]) {
      const result = await decode(xml, [shared])
      expect(result).toMatchObject({ ok: false, diagnostic: { kind: 'parse', feature: 'pivot-records' } })
    }
  })

  it('stops reading at each cap and returns the typed resource-limit diagnostic', async () => {
    const inline = field('S', true)
    const testLimit = async (limits: { maxBytes: number; maxRecords: number; maxInlineStrings: number }) => {
      let finished = false
      async function* input(): AsyncIterable<Uint8Array> {
        yield new TextEncoder().encode('<pivotCacheRecords><r><s v="a"/></r><r><s v="b"/></r>')
        yield new TextEncoder().encode('<r><s v="c"/></r></pivotCacheRecords>')
        finished = true
      }
      const result = await decodeRecords(input(), cache([inline]), limits, false)
      expect(result).toMatchObject({ ok: false, diagnostic: { kind: 'unsupported', feature: 'pivot-records-limit' } })
      expect(finished).toBe(false)
    }
    await testLimit({ maxBytes: 1, maxRecords: 100, maxInlineStrings: 100 })
    await testLimit({ maxBytes: 10000, maxRecords: 1, maxInlineStrings: 100 })
    await testLimit({ maxBytes: 10000, maxRecords: 100, maxInlineStrings: 1 })
  })

  it('pins the D and M caps', () => {
    expect(RECORD_LIMITS).toEqual({ D: { maxBytes: 64 * 1024 * 1024, maxRecords: 4_000_000, maxInlineStrings: 1_048_576 }, M: { maxBytes: 16 * 1024 * 1024, maxRecords: 1_000_000, maxInlineStrings: 262_144 } })
  })

  it('converts ISO dates to the workbook 1900 and 1904 serial systems', () => {
    expect(isoToSerial('1899-12-30T10:00:00', false)).toBeCloseTo(10 / 24)
    expect(isoToSerial('1899-12-31T00:00:00', false)).toBe(1)
    expect(isoToSerial('1900-01-01T00:00:00', false)).toBe(1)
    expect(isoToSerial('1900-02-28T00:00:00', false)).toBe(59)
    expect(isoToSerial('1900-03-01T12:00:00', false)).toBe(61.5)
    expect(isoToSerial('1904-01-01T06:00:00', true)).toBe(0.25)
    expect(isoToSerial('bad', false)).toBeNaN()
    expect(isoToSerial('1899-12-29T23:59:59', false)).toBeNaN()
  })

  it('fails closed on date values before the 1900 epoch', async () => {
    const result = await decode('<pivotCacheRecords><r><d v="1899-12-29T23:59:59"/></r></pivotCacheRecords>', [field('Date', true)])
    expect(result).toMatchObject({ ok: false, diagnostic: { kind: 'parse', feature: 'pivot-records', message: 'record date is invalid' } })
  })

  it.each([
    ['bad number', '<n v="nope"/>', 'record number is invalid'],
    ['bad boolean', '<b v="maybe"/>', 'record boolean is invalid'],
    ['bad date', '<d v="2023-02-29T00:00:00"/>', 'record date is invalid'],
    ['missing value', '<n/>', 'record value is missing its v attribute'],
    ['unknown value kind', '<z v="x"/>', 'unknown record value element z'],
  ])('returns a specific parse diagnostic for %s', async (_label, child, message) => {
    const result = await decode(`<pivotCacheRecords><r>${child}</r></pivotCacheRecords>`, [field('V', true)])
    expect(result).toMatchObject({ ok: false, diagnostic: { kind: 'parse', feature: 'pivot-records', message } })
  })

  it('stores inline date values as serials under both workbook date systems', async () => {
    const dateField = field('Date', true)
    const a = await decodeRecords(chunks('<pivotCacheRecords><r><d v="1900-03-01T12:00:00"/></r></pivotCacheRecords>', 8), cache([dateField]), { maxBytes: 1000, maxRecords: 10, maxInlineStrings: 10 }, false)
    const b = await decodeRecords(chunks('<pivotCacheRecords><r><d v="1904-01-01T06:00:00"/></r></pivotCacheRecords>', 8), cache([dateField]), { maxBytes: 1000, maxRecords: 10, maxInlineStrings: 10 }, true)
    expect(a.ok && a.store.columns[0]).toMatchObject({ kind: 'inline', tag: new Uint8Array([VALUE_TAG.date]), num: new Float64Array([61.5]) })
    expect(b.ok && b.store.columns[0]).toMatchObject({ kind: 'inline', tag: new Uint8Array([VALUE_TAG.date]), num: new Float64Array([0.25]) })
  })

  it('reports retained byte size from typed arrays and the stable text estimate', async () => {
    const result = await decode('<pivotCacheRecords><r><n v="4"/><s v="abc"/></r></pivotCacheRecords>', [field('N', true), field('S', true)])
    expect(result.ok).toBe(true)
    if (!result.ok) return
    const columns = result.store.columns.filter((column): column is NonNullable<typeof column> => column !== undefined)
    const typedBytes = columns.reduce((sum, col) => sum + (col.kind === 'shared' ? col.index.byteLength : col.tag.byteLength + col.num.byteLength + col.textIndex.byteLength), 0)
    expect(result.store.bytes).toBe(typedBytes + 48 + 2 * 3)
  })

  it('applies the inline-string cap per field rather than across the whole store', async () => {
    const xml = '<pivotCacheRecords><r><s v="a"/><s v="x"/></r><r><s v="b"/><s v="y"/></r></pivotCacheRecords>'
    const result = await decodeRecords(chunks(xml, 31), cache([field('A', true), field('B', true)]), { maxBytes: 10000, maxRecords: 10, maxInlineStrings: 2 }, false)
    expect(result.ok).toBe(true)
  })

  it('matches a resident records stream and worker openPart stream', async () => {
    const fixture = resolve(process.cwd(), 'tests/fixtures/xlsx-pivots/pivot-layout-compact.xlsx')
    if (!existsSync(fixture)) return
    const raw = new Uint8Array(readFileSync(fixture)), archive = await JSZip.loadAsync(raw)
    const pkg = await OfficePackage.load(raw), doc = await parseXlsx(pkg)
    const regions = new Map(doc.sheets.filter(sheet => sheet.pivotRegions?.length).map(sheet => [sheet.name, sheet.pivotRegions!]))
    const workbook = await loadPivotWorkbook(pkg, regions), definition = [...workbook.caches.values()][0]
    expect(definition?.recordsPartPath).toBeTruthy()
    if (!definition?.recordsPartPath) return
    const xml = await archive.file(definition.recordsPartPath)!.async('uint8array')
    const expected = await decodeRecords(chunks(new TextDecoder().decode(xml), 4096), definition, RECORD_LIMITS.D, false)
    const inventory = await readZipInventory(bytesRangeSource(raw), { maxEntries: 10000, maxCentralDirectoryBytes: 16 * 1024 * 1024, maxNameBytes: 4096, maxPartExpandedBytes: 128 * 1024 * 1024, maxTotalExpandedBytes: 256 * 1024 * 1024, maxRatio: 1000 })
    const entry = inventory.find(definition.recordsPartPath)!
    const events = recordEvents(definition, RECORD_LIMITS.D, false)
    await tokenizeBytes(openPart(bytesRangeSource(raw), entry, { signal: new AbortController().signal, budget: { charge() {} } }), events.events)
    expect(events.finish()).toEqual(expected)
  })

  it.each([
    ['POI sample', 'poi-ExcelPivotTableSample.xlsx', [0, 1, 0]],
    ['two data fields', 'lo-calc-two-data-fields.xlsx', [0, 0.7034962745892106]],
    ['shared number groups', 'lo-calc-shared-numgroup.xlsx', [0, 0, 55707, 72900, 87902, 0, 124479, 141672, 158865, 176058]],
  ])('decodes the %s corpus cache in database-field order', async (_label, file, firstValues) => {
    const { definition, result } = await corpusRecords(file)
    expect(result.ok, result.ok ? '' : result.diagnostic.message).toBe(true)
    if (!result.ok) return
    expect(result.store.count).toBe(definition.recordCount)
    const first = result.store.columns.filter((column): column is NonNullable<typeof column> => column !== undefined).map(column => column.kind === 'shared' ? column.index[0] : column.tag[0] === VALUE_TAG.number ? column.num[0] : column.text[column.textIndex[0]!])
    expect(first).toEqual(firstValues)
  })

  it('leaves calculated fields without record columns in calcfields corpus', async () => {
    const { definition, result } = await corpusRecords('lo-calc-calcfields.xlsx')
    expect(definition.fields).toHaveLength(8)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.store.count).toBe(definition.recordCount)
    expect(result.store.columns).toHaveLength(8)
    expect(result.store.columns.filter(column => column !== undefined)).toHaveLength(3)
    expect(result.store.columns.map(Boolean)).toEqual(definition.fields.map(field => field.databaseField))
  })
})
