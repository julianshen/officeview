import { describe, expect, it, vi } from 'vitest'
import { parseXml } from '../src/core/xml'
import { tokenizeBytes } from '../src/core/xml-stream'
import { createCfMemory } from '../src/xlsx/cf/limits'
import { createCfStreamEvents } from '../src/xlsx/cf/capture'
import { parseCfSheet } from '../src/xlsx/cf/parse'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { xlsxBytes, xlsxEntries, SS_NS } from './helpers/xlsx-package'
import { buildZip } from './helpers/zip-builder'
import * as formatting from '../src/xlsx/storage/pagedFormatting'
import * as capturing from '../src/xlsx/cf/capture'
const styles = {
  colors: {
    theme: {
      accent1: '4F81BD',
      accent2: 'C0504D',
    },
  },
  dxfNodes: [],
}
const base = '<conditionalFormatting sqref="B2:B4"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>2</formula></cfRule></conditionalFormatting>'
const wrap = (s: string) => `<worksheet xmlns="${SS_NS}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${s}</worksheet>`
async function stream(xml: string, split = 1e9) {
  const memory = createCfMemory(), handle = createCfStreamEvents(styles, memory, {})
  const bytes = new TextEncoder().encode(xml)
  await tokenizeBytes((async function* () { yield bytes.subarray(0, split); yield bytes.subarray(split) })(), handle.events)
  const lease = handle.finish()
  return { lease, memory, handle }
}
describe('paged CF capture', () => {
  it('resident and streamed CF are deep equal at every chunk boundary', async () => {
    const twoColorScale = '<conditionalFormatting sqref="D2:D10"><cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>'
    const threeColorScale = '<conditionalFormatting sqref="E2:E10"><cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>'
    const themeScale = '<conditionalFormatting sqref="F2:F10"><cfRule type="colorScale" priority="5"><colorScale><cfvo type="min"/><cfvo type="max"/><color theme="4" tint="-0.4"/><color theme="5" tint="0.4"/></colorScale></cfRule></conditionalFormatting>'
    const x14Scale = '<x14:conditionalFormatting><x14:cfRule type="colorScale" priority="6" id="{01234567-89AB-CDEF-0123-456789ABCDE0}"><x14:colorScale><x14:cfvo type="min"/><x14:cfvo type="max"/><x14:color rgb="FFFF0000"/><x14:color rgb="FF00FF00"/></x14:colorScale></x14:cfRule><xm:sqref>G1:G10</xm:sqref></x14:conditionalFormatting>'
    const xml = wrap(
      base +
        twoColorScale +
        threeColorScale +
        themeScale +
        '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression" priority="2" id="{01234567-89AB-CDEF-0123-456789ABCDEF}"><xm:f>A1=&quot;茶&quot;</xm:f></x14:cfRule><xm:sqref>C1</xm:sqref></x14:conditionalFormatting>' +
        x14Scale +
        '</x14:conditionalFormattings></ext></extLst>',
    )
    const expected = parseCfSheet(parseXml(xml), styles, createCfMemory()).value
    for (let i = 0; i <= new TextEncoder().encode(xml).length; i++) {
      const result = await stream(xml, i)
      expect(result.lease.value).toEqual(expected)
      result.lease.release(); expect(result.memory.currentBytes).toBe(0)
    }
  })
  it.each(['formula', 'color', 'dxf'])('ten thousand singleton duplicates do not build ten thousand nodes: %s', async tag => {
    const xml = wrap(`<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1">${(`<${tag}>1</${tag}>`).repeat(10_000)}</cfRule></conditionalFormatting>`)
    const expected = parseCfSheet(parseXml(xml), styles, createCfMemory()).value
    const result = await stream(xml)
    if (tag === 'formula') expect(result.lease.value.rules[0]?.definition.kind).toBe('unresolved')
    else expect(result.lease.value.unresolvedSheet).toBeDefined()
    expect(result.lease.value).toEqual(expected)
    expect(result.handle.stats.nodes).toBeLessThan(20)
    expect(result.handle.stats.textUnits).toBeLessThan(20)
    result.lease.release()
  })
  it('overflow before sqref is known withholds sheet CF', async () => {
    const result = await stream(wrap(`<conditionalFormatting><cfRule>${'<formula>1</formula>'.repeat(10_000)}</cfRule></conditionalFormatting>`))
    expect(result.lease.value.unresolvedSheet).toBeDefined(); expect(result.lease.value.rules).toEqual([])
    result.lease.release()
  })
  it('cancelled capture releases partial construction without publishing a result', () => {
    const memory = createCfMemory(), h = createCfStreamEvents(styles, memory, {})
    h.events.open('worksheet', ['xmlns'], [SS_NS], 1, false)
    h.events.open('conditionalFormatting', ['sqref'], ['A1'], 1, false)
    expect(memory.currentBytes).toBeGreaterThan(0)
    h.dispose(); expect(memory.currentBytes).toBe(0); expect(() => h.finish()).toThrow()
  })
  it.each(['beforeSheetData', 'afterSheetData'] as const)('late CF reconciles DXFs before sheet readiness: %s', async position => {
    const bytes = xlsxBytes({ [position]: base.replace('priority="1"', 'priority="1" dxfId="0"'), styles: '<cellXfs><xf/></cellXfs><dxfs><dxf><font><b/></font></dxf></dxfs>' })
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer]) }, { docId: `cf-${position}`, profile: 'D' })
    try {
      await session.done
      const sheet = session.workbook.sheets[0]!
      expect(sheet.state).toBe('complete')
      expect(sheet.metadata.hasConditionalFormatting).toBe(true)
      expect(sheet.conditionalFormatting?.rules).toHaveLength(1)
      expect(sheet.conditionalFormatting?.rules[0]?.dxf).toBeDefined()
      expect(sheet.cfOwner?.bytes).toBeGreaterThan(0)
    } finally { await session.release() }
    expect(session.workbook.sheets[0]!.cfOwner).toBeUndefined()
  })
})

it('loads raw DXFs only once for multiple CF sheets and never for a plain sheet', async () => {
  const loader = vi.spyOn(formatting, 'pagedCfStyles')
  const plain = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(xlsxBytes({})).buffer]) }, { docId: 'plain-cf', profile: 'D' })
  await plain.done; await plain.release(); expect(loader).not.toHaveBeenCalled()
  const entries = xlsxEntries({ afterSheetData: base })
  const workbook = entries.find(e => e.name === 'xl/workbook.xml')!
  workbook.data = String(workbook.data).replace('</sheets>', '<sheet name="T" sheetId="2" r:id="rId4"/></sheets>')
  const rels = entries.find(e => e.name === 'xl/_rels/workbook.xml.rels')!
  rels.data = String(rels.data).replace('</Relationships>', '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>')
  entries.push({ name: 'xl/worksheets/sheet2.xml', data: wrap(base) })
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(buildZip(entries).bytes).buffer]) }, { docId: 'two-cf', profile: 'D' })
  try {
    await session.done
    expect(session.workbook.sheets.map(s => s.conditionalFormatting?.rules.length)).toEqual([1, 1])
    expect(loader).toHaveBeenCalledTimes(1)
  } finally { await session.release(); loader.mockRestore() }
})
it('cancel during the paged CF pass drains the partial collector and publishes nothing', async () => {
  const controller = new AbortController()
  const original = capturing.createCfStreamEvents
  let memory: ReturnType<typeof createCfMemory> | undefined
  const spy = vi.spyOn(capturing, 'createCfStreamEvents').mockImplementation((styles, account, meta) => {
    memory = account
    const h = original(styles, account, meta), open = h.events.open
    h.events.open = (...args) => { open(...args); if (args[0] === 'cfRule') controller.abort() }
    return h
  })
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(xlsxBytes({ afterSheetData: base })).buffer]) }, { docId: 'cancel-cf', profile: 'D', signal: controller.signal })
  try {
    expect((await session.done).kind).toBe('cancelled')
    expect(session.workbook.sheets[0]?.conditionalFormatting).toBeUndefined()
    expect(memory?.currentBytes).toBe(0)
  } finally { await session.release(); spy.mockRestore() }
})
it('keeps excluded AlternateContent and wrong-namespace paths out of effective CF', async () => {
  for (const inner of [
    `<AlternateContent><Choice>${base}</Choice></AlternateContent>`,
    '<extLst><ext xmlns="urn:wrong" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><xm:sqref>A1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>',
  ]) {
    const result = await stream(wrap(inner))
    expect(result.lease.value.rules).toEqual([])
    // Diagnostics may describe the malformed wrapper at different granularity.
    expect(parseCfSheet(parseXml(wrap(inner)), styles, createCfMemory()).value.rules).toEqual([])
    result.lease.release()
  }
})
it('refuses CF before constructing unadmitted raw DXFs, preserving the base import', async () => {
  const { ConstructionMemory } = await import('../src/xlsx/storage/constructionMemory')
  const { PackageError } = await import('../src/core/package-error')
  const original = ConstructionMemory.prototype.share
  const allocator = vi.spyOn(ConstructionMemory.prototype, 'share').mockImplementation(function (this: InstanceType<typeof ConstructionMemory>, name, options) {
    const share = original.call(this, name, options)
    if (name === 'CF raw DXF XML') return { ...share, need() { throw new PackageError('resource-limit', 'test admission refusal') } }
    return share
  })
  const loader = vi.spyOn(formatting, 'pagedCfStyles')
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(xlsxBytes({ afterSheetData: base })).buffer]) }, { docId: 'refuse-cf', profile: 'D' })
  try {
    expect((await session.done).kind).toBe('complete')
    expect(loader).not.toHaveBeenCalled()
    expect(session.workbook.sheets[0]?.conditionalFormatting?.unresolvedSheet).toBeDefined()
  } finally { await session.release(); allocator.mockRestore(); loader.mockRestore() }
})
it('publishes late x14 definitions with the same normalized model as resident parsing', async () => {
  const extension = '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:conditionalFormatting><x14:cfRule type="expression" priority="1" id="{01234567-89AB-CDEF-0123-456789ABCDEF}"><xm:f>A1&gt;2</xm:f><x14:dxf><x14:font><x14:b/></x14:font></x14:dxf></x14:cfRule><xm:sqref>A1:A4</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>'
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(xlsxBytes({ afterSheetData: extension })).buffer]) }, { docId: 'x14-cf', profile: 'D' })
  try {
    await session.done
    const sheet = session.workbook.sheets[0]!
    const expected = parseCfSheet(parseXml(wrap(extension)), styles, createCfMemory(), { sheetId: sheet.sheetId, partPath: sheet.partPath })
    expect(sheet.metadata.hasConditionalFormatting).toBe(true)
    expect(sheet.conditionalFormatting).toEqual(expected.value)
    expected.release()
  } finally { await session.release() }
})
it('worker admission precedes capture growth and retained ownership lasts until explicit release', async () => {
  const { ConstructionMemory } = await import('../src/xlsx/storage/constructionMemory')
  const { MemoryPool } = await import('../src/xlsx/view/tileCache')
  const construction = new ConstructionMemory(new MemoryPool({ limitBytes: 1024 * 1024 }), 'cf-byte-test')
  const memory = capturing.createPagedCfMemory(construction)
  const h = capturing.createCfStreamEvents(styles, memory, {})
  const bytes = new TextEncoder().encode(wrap(base))
  await tokenizeBytes((async function* () { yield bytes })(), h.events)
  expect(memory.constructionBytes).toBeGreaterThan(0)
  const lease = h.finish()
  expect(memory.constructionBytes).toBe(0)
  expect(memory.retainedBytes).toBe(lease.bytes)
  h.dispose(); expect(memory.retainedBytes).toBe(lease.bytes)
  expect(construction.transientBytes).toBeGreaterThanOrEqual(lease.bytes)
  memory.close(); expect(memory.retainedBytes).toBeGreaterThan(0)
  lease.release(); expect(memory.currentBytes).toBe(0)
  expect(construction.transientBytes).toBe(0)
  construction.releaseAll()
})
it('resident duplicate preflight reads only a bounded overflow witness, without copying the child list', () => {
  let reads = 0
  const formulas = new Proxy(Array.from({ length: 10_000 }, () => ({ '@attrs': {}, '#text': '1' })), {
    get(target, key, receiver) { if (typeof key === 'string' && /^\d+$/.test(key)) reads++; return Reflect.get(target, key, receiver) },
  })
  const worksheet = { '@attrs': {}, conditionalFormatting: { '@attrs': { sqref: 'A1' }, cfRule: { '@attrs': { type: 'expression', priority: '1' }, formula: formulas } } }
  const lease = parseCfSheet(worksheet, styles, createCfMemory())
  expect(lease.value.rules[0]?.definition.kind).toBe('unresolved')
  expect(reads).toBeLessThan(16)
  lease.release()
})

describe('regression bot P2: paged CF dateContext', () => {
  it('keeps local timeZone and epochNowMs in default dateContext', () => {
    const fs = require('node:fs')
    const src = fs.readFileSync('src/xlsx/view/pagedSource.ts', 'utf8')
    expect(src).toMatch(/timeZone: Intl\.DateTimeFormat\(\)\.resolvedOptions\(\)\.timeZone \?\? 'UTC'/)
    expect(src).toMatch(/epochNowMs: Date\.now\(\)/)
  })
})
