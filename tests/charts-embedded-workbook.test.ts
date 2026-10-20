/**
 * A3b: the bounded embedded-workbook reader (spec 6.4). Only workbook.xml,
 * its rels, sharedStrings, styles (numFmts/cellXfs) and the referenced
 * sheets are read, through the bounded streaming tokenizer; only cells on the
 * chart's lines are kept, and a sheet stops after its last referenced row.
 * Formulas are never evaluated. Every allocation is charged to the document
 * chart workspace. Fixtures are Office-shaped synthetic.
 */
import { describe, expect, it, vi } from 'vitest'
import type JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { XmlTokenizer } from '../src/core/xml-stream'
import { chartParseContext, documentChartContext } from '../src/charts/context'
import { readChartPart } from '../src/charts/read'
import { parseXlsxChart } from '../src/charts/parse'
import { pointList } from '../src/charts/points'
import { releaseChartModel } from '../src/charts/own'
import type { ChartModel } from '../src/charts/model'
import type { ChartProfile } from '../src/charts/workspace'
import { buildChartDocx, cell, embeddedWorkbook, officeChartWorkbook, officeChartXml, row, type ChartEmbedSpec } from './helpers/chart-fixtures'

const streams = vi.hoisted(() => ({ names: [] as string[] }))
vi.mock('../src/charts/read', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/charts/read')>()
  return { ...actual, streamEntry: (file: JSZip.JSZipObject, ...rest: Parameters<typeof actual.streamEntry> extends [unknown, ...infer R] ? R : never) => { streams.names.push(file.name); return actual.streamEntry(file, ...rest) } }
})
const { applyDataPrecedence } = await import('../src/charts/dataPrecedence')
const { embeddedWorkbookStats, EMBEDDED_SST_COUNT, EMBEDDED_SST_BYTES, EMBEDDED_CELL_LIMIT } = await import('../src/charts/embeddedWorkbook')

const NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
/** A cacheless line chart: one series per values formula (and an optional categories formula). */
const chartXml = (vals: string[], cat?: string) => `<c:chartSpace ${NS}><c:chart><c:plotArea><c:lineChart><c:grouping val="standard"/>${vals.map((f, i) => `<c:ser><c:idx val="${i}"/>${cat ? `<c:cat><c:strRef><c:f>${cat}</c:f></c:strRef></c:cat>` : ''}<c:val><c:numRef><c:f>${f}</c:f></c:numRef></c:val></c:ser>`).join('')}<c:axId val="1"/><c:axId val="2"/></c:lineChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/><c:axPos val="b"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/><c:axPos val="l"/></c:valAx></c:plotArea></c:chart><c:externalData r:id="rId1"/></c:chartSpace>`
const points = (chart: ChartModel, series = 0) => pointList(chart.groups[0].series[series].values).map(p => p.state === 'value' ? p.v : p.state === 'error' ? p.code : p.state)

/** Parse the DOCX's chart parts and apply precedence (concurrently when asked) in one document workspace. */
async function precedence(specs: ChartEmbedSpec[], options: { profile?: ChartProfile; concurrent?: boolean; signal?: AbortSignal } = {}) {
  const fixture = buildChartDocx(specs)
  const pkg = await OfficePackage.load(fixture.bytes)
  const doc = documentChartContext(options.profile ?? 'desktop', options.signal)
  // The peak of the whole workspace, and of the cell reader's own charges (tags `embedded-cells:*`).
  let peak = 0, readerPeak = 0
  const owned = new Map<string, number>(), readerUsed = () => [...owned].reduce((sum, [tag, bytes]) => sum + (tag.startsWith('embedded-cells:') ? bytes : 0), 0)
  const reserve = doc.workspace.reserve.bind(doc.workspace), release = doc.workspace.release.bind(doc.workspace)
  doc.workspace.reserve = (bytes, tag) => { const ok = reserve(bytes, tag); if (ok) owned.set(tag, (owned.get(tag) ?? 0) + bytes); peak = Math.max(peak, doc.workspace.used); readerPeak = Math.max(readerPeak, readerUsed()); return ok }
  doc.workspace.release = tag => { owned.delete(tag); release(tag) }
  streams.names.length = 0
  const one = async (path: string) => {
    const context = chartParseContext(doc, pkg, path)
    const read = await readChartPart(pkg, path, context)
    if (read?.kind !== 'tree') throw new Error('chart part unreadable')
    const model = parseXlsxChart(read.root, undefined, context)!
    read.release()
    return applyDataPrecedence(model, read.root, context)
  }
  const charts: ChartModel[] = []
  if (options.concurrent) charts.push(...await Promise.all(fixture.chartPaths.map(one)))
  else for (const path of fixture.chartPaths) charts.push(await one(path))
  return { charts, workspace: doc.workspace, peak: () => peak, readerPeak: () => readerPeak, release: () => { for (const chart of charts) releaseChartModel(chart) } }
}

describe('embedded workbook reader (spec 6.4)', () => {
  it('formula cells without a saved value are unavailable points; with a saved value they use it', async () => {
    const data = [
      row(2, cell('B2', 5, { f: 'A2*5' })),
      row(3, cell('B3', undefined, { f: 'A3*5' })),
      row(4, cell('B4', '#DIV/0!', { t: 'e', f: '1/0' })),
      row(5, cell('B5', ' 12 ', { t: 'str', f: '" 12 "' })),
      row(6, cell('B6', '0x10', { t: 'str' })),
      row(7, '<c r="B7"><f t="shared" si="0"/></c>'),
      row(8, cell('B8', '#N/A', { t: 'e' })),
      row(9, cell('B9', 1, { t: 'b' })),
    ].join('')
    const { charts, workspace, release } = await precedence([{ chartXml: chartXml(['Sheet1!$B$2:$B$9']), workbook: embeddedWorkbook({ sheets: [{ name: 'Sheet1', data }] }) }])
    expect(charts[0].unsupported).toBeUndefined()
    expect(points(charts[0])).toEqual([5, 'unavailable', '#DIV/0!', 12, 'blank', 'unavailable', '#N/A', 'blank'])
    expect(charts[0].diagnostics).toContainEqual(expect.objectContaining({ code: 'chart-formula-unsaved' }))
    release()
    expect(workspace.used).toBe(0)
  })

  it('only referenced sheets and ranges are read; reading stops after the last referenced row', async () => {
    const big = (rows: number, from: number) => Array.from({ length: rows }, (_, i) => row(from + i, cell(`A${from + i}`, `n${i}`, { t: 'str' }), cell(`B${from + i}`, i), cell(`C${from + i}`, i * 2))).join('')
    const workbook = (trailing: number) => embeddedWorkbook({
      sharedStrings: ['Category', 'Series 1', 'A', 'B'],
      sheets: [
        { name: 'Sheet1', data: row(1, cell('A1', 0, { t: 's' }), cell('B1', 1, { t: 's' }), cell('C1', 'x', { t: 'str' })) + row(2, cell('A2', 2, { t: 's' }), cell('B2', 4.3), cell('C2', 9)) + row(3, cell('A3', 3, { t: 's' }), cell('B3', 2.5), cell('C3', 9)) + big(trailing, 4) },
        { name: 'Unused', data: big(2000, 1) },
      ],
    })
    const run = async (trailing: number) => {
      const out = await precedence([{ chartXml: officeChartXml(), workbook: workbook(trailing) }])
      const stats = embeddedWorkbookStats()
      return { ...out, stats, names: [...streams.names] }
    }
    const small = await run(1000), large = await run(10000)
    for (const r of [small, large]) {
      expect(r.charts[0].unsupported).toBeUndefined()
      expect(points(r.charts[0])).toEqual([4.3, 2.5])
      expect(r.charts[0].categories).toEqual(['A', 'B'])
      // The unreferenced sheet is never inflated; the referenced one is.
      expect(r.names).toContain('xl/worksheets/sheet1.xml')
      expect(r.names).not.toContain('xl/worksheets/sheet2.xml')
      // Allocation counters: exactly the five cells on the chart's lines are kept (B1 name, A2:A3, B2:B3; never column C).
      expect(r.stats.cells).toBe(5)
      expect(r.stats.strings).toBe(3)
      r.release()
      expect(r.workspace.used).toBe(0)
    }
    // Reading stops after row 3: the rows visited, the inflated bytes and the reader's peak do not grow with the
    // trailing rows (only the admitted embedding itself, A3a's charge, is larger).
    expect(large.stats.rows).toBe(3)
    expect(small.stats.rows).toBe(3)
    expect(large.stats.parts['xl/worksheets/sheet1.xml']).toBeLessThan(256 * 1024)
    expect(large.readerPeak()).toBe(small.readerPeak())
  })

  it('names resolved with bounds; a name pointing at an external workbook is unavailable', async () => {
    const workbook = officeChartWorkbook([4, 6], ['A', 'B'], {
      definedNames: [
        { name: 'Vals', formula: 'Sheet1!$B$2:$B$3' },
        { name: 'Local', formula: 'Sheet1!$B$3', localSheetIndex: 0 },
        { name: 'Local', formula: 'Sheet1!$B$2' },
        { name: 'Ext', formula: '[1]Sheet1!$B$2:$B$3' },
        { name: 'ExtBook', formula: "'[Book2.xlsx]Sheet1'!$B$2:$B$3" },
        { name: 'Chain', formula: 'Vals' },
        { name: 'Multi', formula: 'Sheet1!$B$2,Sheet1!$B$3' },
        { name: 'Long', formula: `Sheet1!$B$2${' '.repeat(9000)}` },
        { name: 'Cats', formula: '[1]Sheet1!$A$2:$A$3' },
      ],
    })
    const { charts, workspace, release } = await precedence([
      { chartXml: chartXml(['Vals']), workbook },
      { chartXml: chartXml(['Sheet1!Local']), workbook },
      { chartXml: chartXml(['Vals'], 'Cats'), workbook },
      ...['Ext', 'ExtBook', 'Chain', 'Multi', 'Long', 'Missing', '[1]Sheet1!$B$2:$B$3'].map(name => ({ chartXml: chartXml([name]), workbook })),
    ])
    expect(points(charts[0])).toEqual([4, 6])
    // Sheet scope wins over workbook scope.
    expect(points(charts[1])).toEqual([6])
    // An external categories name: the optional role is unavailable, the chart stays with default categories.
    expect(charts[2].unsupported).toBeUndefined()
    expect(charts[2].categories).toEqual([])
    expect(charts[2].diagnostics).toContainEqual(expect.objectContaining({ code: 'chart-data-external-name' }))
    const codes = charts.slice(3).map(chart => [chart.unsupported, chart.diagnostics?.find(d => d.code.startsWith('chart-data-'))?.code])
    expect(codes).toEqual([
      ["This chart's data is not available", 'chart-data-external-name'],
      ["This chart's data is not available", 'chart-data-external-name'],
      ["This chart's data is not available", 'chart-data-name-unresolved'],
      ["This chart's data is not available", 'chart-data-name-unresolved'],
      ["This chart's data is not available", 'chart-data-name-unresolved'],
      ["This chart's data is not available", 'chart-data-name-missing'],
      ["This chart's data is not available", 'chart-data-external-reference'],
    ])
    release()
    expect(workspace.used).toBe(0)
  })

  it('shared strings over the byte or count budget refuse and refund', async () => {
    const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    const data = row(2, cell('B2', 0, { t: 's' }))
    // Count: one string over the cap. Each is empty: at 32 B per string the count cap is exactly the desktop byte
    // cap, so only empty strings reach the count first.
    const count = `<sst xmlns="${S}">${'<si><t></t></si>'.repeat(EMBEDDED_SST_COUNT + 1)}</sst>`
    // Bytes (mobile, 4 MiB at 2 B per unit plus 32 B per string): one string over the cap.
    const bytes = `<sst xmlns="${S}"><si><t>1</t></si><si><t>${'x'.repeat(EMBEDDED_SST_BYTES.mobile / 2)}</t></si></sst>`
    for (const [label, sst, profile, code] of [['count', count, 'desktop', 'embedded-sst-count'], ['bytes', bytes, 'mobile', 'embedded-sst-limit']] as const) {
      const { charts, workspace, release } = await precedence([{ chartXml: chartXml(['Sheet1!$B$2']), workbook: embeddedWorkbook({ sst, sheets: [{ name: 'Sheet1', data }] }) }], { profile })
      expect(charts[0].refused, label).toBe(true)
      expect(charts[0].diagnostics, label).toContainEqual(expect.objectContaining({ code }))
      release()
      expect(workspace.used, label).toBe(0)
    }
    // Just under both caps reads.
    const under = `<sst xmlns="${S}"><si><t>7</t></si>${'<si><t>1</t></si>'.repeat(1000)}</sst>`
    const ok = await precedence([{ chartXml: chartXml(['Sheet1!$B$2']), workbook: embeddedWorkbook({ sst: under, sheets: [{ name: 'Sheet1', data }] }) }])
    expect(points(ok.charts[0])).toEqual([7])
    ok.release()
  })

  it('cells per chart over the budget refuse and refund', async () => {
    // 64 series of 4,096 values plus 4,096 shared categories: 266,240 cells on the chart's lines, over 262,144.
    const col = (i: number): string => (i >= 26 ? String.fromCharCode(64 + Math.floor(i / 26)) : '') + String.fromCharCode(65 + (i % 26))
    const vals = Array.from({ length: 64 }, (_, i) => `Sheet1!$${col(i + 1)}$1:$${col(i + 1)}$4096`)
    expect(64 * 4096 + 4096).toBeGreaterThan(EMBEDDED_CELL_LIMIT)
    const { charts, workspace, release } = await precedence([{ chartXml: chartXml(vals, 'Sheet1!$A$1:$A$4096'), workbook: officeChartWorkbook() }])
    expect(charts[0].refused).toBe(true)
    expect(charts[0].diagnostics).toContainEqual(expect.objectContaining({ code: 'embedded-cells-limit' }))
    release()
    expect(workspace.used).toBe(0)
  })

  it('50 distinct embeddings and repeated references stay within the workspace', async () => {
    const distinct = Array.from({ length: 50 }, (_, i) => ({ workbook: officeChartWorkbook([i, i + 1]) }))
    const one = await precedence(distinct.slice(0, 1))
    const single = one.peak()
    one.release()
    const fifty = await precedence(distinct)
    expect(fifty.charts.map(chart => chart.unsupported)).toEqual(Array(50).fill(undefined))
    expect(fifty.charts.map(chart => points(chart))).toEqual(distinct.map((_, i) => [i, i + 1]))
    // One embedded read at a time, its cells discarded once filled: the peak is one read plus the retained models.
    expect(fifty.peak()).toBeLessThan(single + 50 * 16 * 1024)
    expect(fifty.peak()).toBeLessThanOrEqual(fifty.workspace.limit)
    fifty.release()
    expect(fifty.workspace.used).toBe(0)

    // Repeated references: 20 charts, one embedding part. Concurrent requests share one read.
    const target = '../embeddings/Shared.xlsx'
    const repeated: ChartEmbedSpec[] = Array.from({ length: 20 }, (_, i) => ({ target, ...(i === 0 ? { workbook: officeChartWorkbook([3, 4]) } : {}) }))
    const shared = await precedence(repeated, { concurrent: true })
    expect(shared.charts.map(chart => points(chart))).toEqual(Array(20).fill([3, 4]))
    expect(streams.names.filter(name => name === 'word/embeddings/Shared.xlsx')).toHaveLength(1)
    expect(streams.names.filter(name => name === 'xl/worksheets/sheet1.xml')).toHaveLength(1)
    shared.release()
    expect(shared.workspace.used).toBe(0)
    // Sequential repeated references read again (nothing is retained between consumers) and still refund.
    const sequential = await precedence(repeated.slice(0, 3))
    expect(sequential.charts.map(chart => points(chart))).toEqual(Array(3).fill([3, 4]))
    sequential.release()
    expect(sequential.workspace.used).toBe(0)
  })

  it('cancellation stops publication; charges are refunded once the read settles', async () => {
    const abort = new AbortController()
    abort.abort()
    const { charts, workspace, release } = await precedence([{ workbook: officeChartWorkbook() }], { signal: abort.signal })
    expect(charts[0].unsupported).toBeDefined()
    expect(charts[0].diagnostics).toContainEqual(expect.objectContaining({ code: 'chart-cancelled' }))
    release()
    expect(workspace.used).toBe(0)
  })

  it('cancellation mid-read: a gated embedded load returns cancelled at once; its charges are held until it settles', async () => {
    const fixture = buildChartDocx([{ chartXml: officeChartXml(), workbook: officeChartWorkbook() }])
    const pkg = await OfficePackage.load(fixture.bytes)
    const abort = new AbortController()
    const doc = documentChartContext('desktop', abort.signal)
    const context = chartParseContext(doc, pkg, fixture.chartPaths[0])
    const read = await readChartPart(pkg, fixture.chartPaths[0], context)
    if (read?.kind !== 'tree') throw new Error('unreadable')
    const model = parseXlsxChart(read.root, undefined, context)!
    read.release()
    const modelBytes = doc.workspace.used
    // Gate the embedding's load (the outer package is already loaded): it starts, then waits.
    let open!: () => void, started!: () => void
    const gate = new Promise<void>(resolve => { open = resolve })
    const loading = new Promise<void>(resolve => { started = resolve })
    const real = OfficePackage.load.bind(OfficePackage)
    const spy = vi.spyOn(OfficePackage, 'load').mockImplementation(async data => { started(); await gate; return real(data) })
    try {
      let settled = false
      const pending = applyDataPrecedence(model, read.root, context).then(chart => { settled = true; return chart })
      await loading
      const held = doc.workspace.used
      expect(held).toBeGreaterThan(modelBytes)
      abort.abort()
      const chart = await pending
      // Publication stopped at once, while the load is still gated: the charges stay held.
      expect(settled).toBe(true)
      expect(chart.diagnostics).toContainEqual(expect.objectContaining({ code: 'chart-cancelled' }))
      expect(doc.workspace.used).toBeGreaterThan(modelBytes)
      open()
      // Once the load settles, the read stops and refunds; only the (cancelled) model's own charge remains.
      await vi.waitFor(() => { releaseChartModel(chart); expect(doc.workspace.used).toBe(0) })
    } finally { spy.mockRestore(); open() }
  })

  it('a corrupt embedding or worksheet is unavailable, never a throw', async () => {
    const corrupt = await precedence([{ chartXml: officeChartXml(), workbook: new Uint8Array([0x50, 0x4b, 3, 4, 1, 2, 3]) }])
    expect(corrupt.charts[0].unsupported).toBe("This chart's data is not available")
    corrupt.release()
    const badSheet = await precedence([{ chartXml: officeChartXml(), workbook: embeddedWorkbook({ sheets: [{ name: 'Sheet1', raw: '<worksheet><sheetData><row r="1"><c r="A1"></row></sheetData></worksheet>' }] }) }])
    expect(badSheet.charts[0].unsupported).toBe("This chart's data is not available")
    expect(badSheet.charts[0].diagnostics).toContainEqual(expect.objectContaining({ code: 'chart-data-corrupt' }))
    badSheet.release()
    expect(badSheet.workspace.used).toBe(0)
  })

  it('tokenizer state retained across chunks stays within what is admitted; oversized names are refused', async () => {
    type Probe = { stack: string[]; pending: string; carry: string; names: string[]; values: string[] }
    const units = (t: Probe) => [...t.stack, t.pending, t.carry, ...t.names, ...t.values].reduce((n, s) => n + (s?.length ?? 0), 0)
    const orig = XmlTokenizer.prototype.write
    const run = async (sheet: string) => {
      let writes = 0, worst = -Infinity
      let used = () => 0
      XmlTokenizer.prototype.write = function (this: XmlTokenizer, text: string) {
        try { return orig.call(this, text) } finally { writes++; worst = Math.max(worst, 2 * units(this as unknown as Probe) - used()) }
      }
      try {
        const fixture = buildChartDocx([{ chartXml: officeChartXml(), workbook: embeddedWorkbook({ sheets: [{ name: 'Sheet1', raw: sheet }] }) }])
        const pkg = await OfficePackage.load(fixture.bytes)
        const doc = documentChartContext('desktop')
        used = () => doc.workspace.used
        const context = chartParseContext(doc, pkg, fixture.chartPaths[0])
        const read = await readChartPart(pkg, fixture.chartPaths[0], context)
        if (read?.kind !== 'tree') throw new Error('unreadable')
        const model = parseXlsxChart(read.root, undefined, context)!
        read.release()
        const before = doc.workspace.used
        const chart = await applyDataPrecedence(model, read.root, context)
        releaseChartModel(chart)
        expect(doc.workspace.used).toBe(0)
        return { chart, writes, worst, before }
      } finally { XmlTokenizer.prototype.write = orig }
    }
    const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    const names = Array.from({ length: 8 }, (_, i) => String.fromCharCode(65 + i).repeat(15000))
    const nested = await run(`<worksheet ${S}><extLst>${names.map(n => `<${n}>`).join('')}${[...names].reverse().map(n => `</${n}>`).join('')}</extLst><sheetData/></worksheet>`)
    expect(nested.chart.diagnostics).toContainEqual(expect.objectContaining({ code: 'chart-data-corrupt' }))
    expect(nested.worst).toBeLessThanOrEqual(0)
    // A valid sheet spread over many chunks, with long ignored extension text before the data.
    const spread = await run(`<worksheet ${S}><sheetPr><tabColor rgb="FF00FF00"/></sheetPr>${'<x:ext xmlns:x="u"><x:t>' + 'z'.repeat(100000) + '</x:t></x:ext>'}<sheetData>${row(1, cell('B1', 'Name', { t: 'str' }))}${row(2, cell('A2', 'A', { t: 'str' }), cell('B2', 1))}${row(3, cell('A3', 'B', { t: 'str' }), cell('B3', 2))}</sheetData></worksheet>`)
    expect(spread.chart.unsupported).toBeUndefined()
    expect(spread.writes).toBeGreaterThan(4)
    expect(spread.worst).toBeLessThanOrEqual(0)
  })
})
