/**
 * A3a: bounded embedded-workbook access for DOCX/PPTX chart data (spec 6.4):
 * resolve `c:externalData`, admit the embedding (inflated cap, central
 * directory inventory) through an uncached read, one read at a time per
 * document, cancellation that stops publication while charges stay held until
 * the producer settles. Fixtures are Office-shaped synthetic.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { chartParseContext, documentChartContext } from '../src/charts/context'
import { createChartWorkspace } from '../src/charts/workspace'
import { buildChartDocx, buildChartPptx, embeddedXlsx, type ChartEmbedSpec } from './helpers/chart-fixtures'
import { buildZip, type BuildEntry } from './helpers/zip-builder'
import type JSZip from 'jszip'

/**
 * Gate for an entry's inflate (`target`: the embedding, or the chart's `.rels`); only target streams count in `calls`.
 * - `collect`: the whole entry is inflated ahead into the mock, the first chunk is forwarded, the rest wait for
 *   `gate.open()`. It models a producer whose completion is pending, not a paused inflate.
 * - `pause`: the real stream delivers its first chunk, then JSZip's worker chain is paused (nothing more is read or
 *   inflated) until `gate.open()`.
 */
const gate = vi.hoisted(() => {
  const state = {
    calls: 0,
    chunks: 0,
    mode: 'collect' as 'collect' | 'pause',
    target: 'embedding' as 'embedding' | 'rels',
    pending: undefined as Promise<void> | undefined,
    open: () => {},
    close(mode: 'collect' | 'pause' = 'collect', target: 'embedding' | 'rels' = 'embedding') { state.mode = mode; state.target = target; state.pending = new Promise<void>(resolve => { state.open = resolve }) },
    reset() { state.calls = 0; state.chunks = 0; state.mode = 'collect'; state.target = 'embedding'; state.pending = undefined; state.open = () => {} },
  }
  return state
})
vi.mock('../src/charts/read', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/charts/read')>()
  return {
    ...actual,
    streamEntry: async (file: JSZip.JSZipObject, onChunk: (chunk: Uint8Array) => void) => {
      const isTarget = file.name.endsWith('.rels') === (gate.target === 'rels')
      if (!isTarget) return actual.streamEntry(file, onChunk)
      gate.calls++
      const pending = gate.pending
      if (!pending) return actual.streamEntry(file, onChunk)
      if (gate.mode === 'pause') {
        let first = true
        return actual.streamEntry(file, chunk => { gate.chunks++; onChunk(chunk) }, () => { if (first) { first = false; return pending } })
      }
      const chunks: Uint8Array[] = []
      await actual.streamEntry(file, chunk => { chunks.push(chunk) })
      onChunk(chunks[0])
      await pending
      for (const chunk of chunks.slice(1)) onChunk(chunk)
    },
  }
})

const { openEmbeddedWorkbook, resolveExternalData, EMBEDDED_RELS_LIMIT, EMBEDDED_RELS_CROSSING, EMBEDDED_WORKBOOK_LIMIT, EMBEDDED_ENTRY_LIMIT, EMBEDDED_DECLARED_LIMIT } = await import('../src/charts/embeddedZip')

const MiB = 1024 * 1024
afterEach(() => gate.reset())

/** Deterministic incompressible bytes. */
function noise(length: number, seed = 1): Uint8Array {
  const out = new Uint8Array(length)
  let x = seed >>> 0 || 1
  for (let i = 0; i < length; i++) { x ^= x << 13; x ^= x >>> 17; x ^= x << 5; out[i] = x & 0xff }
  return out
}
/** An embedded workbook padded with a stored noise part, so its inflate spans several chunks. */
async function paddedWorkbook(padding = 96 * 1024): Promise<Uint8Array> {
  const JSZipModule = (await import('jszip')).default
  const zip = await JSZipModule.loadAsync(await embeddedXlsx())
  zip.file('xl/media/padding.bin', noise(padding), { compression: 'STORE' })
  return zip.generateAsync({ type: 'uint8array' })
}
async function open(fixture: { bytes: Uint8Array; chartPaths: string[] }, options: { profile?: 'desktop' | 'mobile'; signal?: AbortSignal; chart?: number; doc?: ReturnType<typeof documentChartContext> } = {}) {
  const pkg = await OfficePackage.load(fixture.bytes)
  const doc = options.doc ?? documentChartContext(options.profile ?? 'desktop', options.signal)
  const context = chartParseContext(doc, pkg, fixture.chartPaths[options.chart ?? 0])
  return { pkg, doc, context, result: openEmbeddedWorkbook(context, 'rId1') }
}
const docx = (spec: ChartEmbedSpec) => buildChartDocx([spec])
const cacheKeys = (pkg: OfficePackage) => [...(pkg as unknown as { cache: Map<string, unknown> }).cache.keys()]

describe('embedded workbook access (A3a)', () => {
  it('resolves an internal c:externalData target and admits the embedding (DOCX and PPTX)', async () => {
    const workbook = await embeddedXlsx()
    for (const fixture of [buildChartDocx([{ workbook }]), buildChartPptx([{ workbook }])]) {
      const { doc, result } = await open(fixture)
      const read = await result
      expect(read.kind).toBe('workbook')
      if (read.kind !== 'workbook') continue
      expect(read.path).toBe(fixture.embeddingPaths[0])
      expect(read.bytes).toEqual(workbook)
      expect(read.inventory.find('xl/workbook.xml')).toBeDefined()
      // Charged before allocation, and still charged while the caller holds it.
      expect(doc.workspace.used).toBeGreaterThanOrEqual(workbook.length)
      expect(doc.workspace.used).toBe(read.charged)
      read.release()
      expect(doc.workspace.used).toBe(0)
      expect(() => read.bytes).toThrow()
      read.release()
      expect(doc.workspace.used).toBe(0)
    }
  })

  it('external targets and package-escaping paths are unavailable', async () => {
    const workbook = await embeddedXlsx()
    const cases: Array<[ChartEmbedSpec, string]> = [
      [{ target: 'file:///C:/Users/me/Book1.xlsx', external: true, workbook }, 'chart-data-external'],
      [{ target: 'https://example.com/book.xlsx', external: true }, 'chart-data-external'],
      [{ target: '../../../outside.xlsx', workbook }, 'chart-data-outside-package'],
      [{ target: '/../outside.xlsx', workbook }, 'chart-data-outside-package'],
      [{ target: '%2E%2E/%2E%2E/%2E%2E/outside.xlsx', workbook }, 'chart-data-outside-package'],
      [{ target: '..\\..\\outside.xlsx', workbook }, 'chart-data-outside-package'],
      [{ target: 'file:///C:/Book1.xlsx', workbook }, 'chart-data-outside-package'],
      [{ target: '../embeddings/absent.xlsx' }, 'chart-data-missing'],
    ]
    for (const [spec, code] of cases) {
      for (const fixture of [buildChartDocx([spec]), buildChartPptx([spec])]) {
        const { doc, context, result } = await open(fixture)
        const read = await result
        expect(read, `${spec.target}`).toMatchObject({ kind: 'unavailable', diagnostic: { code } })
        if (read.kind === 'unavailable') expect(read.diagnostic.reason).toMatch(/\w/)
        expect(doc.workspace.used).toBe(0)
        expect(await resolveExternalData(context, 'rId1')).toMatchObject({ kind: 'unavailable', diagnostic: { code } })
      }
    }
    // None of them inflated anything.
    expect(gate.calls).toBe(0)
    // An unknown relationship id, and a relationship-less chart part.
    const fixture = docx({ workbook }), pkg = await OfficePackage.load(fixture.bytes)
    const context = chartParseContext(documentChartContext('desktop'), pkg, fixture.chartPaths[0])
    expect(await resolveExternalData(context, 'rId9')).toMatchObject({ kind: 'unavailable', diagnostic: { code: 'chart-data-missing' } })
    expect(await resolveExternalData({ ...context, partPath: 'word/charts/none.xml' }, 'rId1')).toMatchObject({ kind: 'unavailable', diagnostic: { code: 'chart-data-missing' } })
    // Inside the package, `..` that stays inside is fine.
    expect(await resolveExternalData(context, 'rId1')).toEqual({ kind: 'internal', path: 'word/embeddings/Microsoft_Excel_Worksheet.xlsx' })
  })

  it('admission: inflated cap 8 and 4 MiB, at most 256 entries, declared-size sum checked before load; a forged declared size is caught by the inflated counter', async () => {
    expect(EMBEDDED_WORKBOOK_LIMIT).toEqual({ desktop: 8 * MiB, mobile: 4 * MiB })
    expect(EMBEDDED_ENTRY_LIMIT).toBe(256)
    expect(EMBEDDED_DECLARED_LIMIT).toEqual({ desktop: 32 * MiB, mobile: 16 * MiB })
    /** Peak workspace use across a read. */
    const peakOf = (doc: ReturnType<typeof documentChartContext>) => {
      let peak = 0
      const reserve = doc.workspace.reserve.bind(doc.workspace)
      doc.workspace.reserve = (bytes, tag) => { const ok = reserve(bytes, tag); peak = Math.max(peak, doc.workspace.used); return ok }
      return () => peak
    }
    const run = async (spec: ChartEmbedSpec, profile: 'desktop' | 'mobile') => {
      const pkg = await OfficePackage.load(docx(spec).bytes)
      const doc = documentChartContext(profile)
      const peak = peakOf(doc)
      const read = await openEmbeddedWorkbook(chartParseContext(doc, pkg, 'word/charts/chart1.xml'), 'rId1')
      const used = doc.workspace.used
      if (read.kind === 'workbook') read.release()
      return { read, used, peak: peak(), after: doc.workspace.used }
    }

    // Inflated cap, counted on actual bytes (desktop 8 MiB, mobile 4 MiB).
    const over8 = noise(8 * MiB + 1)
    const desktop = await run({ workbook: over8 }, 'desktop')
    expect(desktop.read).toMatchObject({ kind: 'refused', diagnostic: { code: 'embedded-workbook-limit' }, limit: 8 * MiB })
    expect(desktop.after).toBe(0)
    expect(desktop.peak).toBeLessThanOrEqual(8 * MiB + 256 * 1024)
    const over4 = noise(4 * MiB + 1)
    expect((await run({ workbook: over4 }, 'mobile')).read).toMatchObject({ kind: 'refused', diagnostic: { code: 'embedded-workbook-limit' }, limit: 4 * MiB })
    // The same bytes pass the desktop cap: the inflate is admitted; the inventory then rejects the non-zip.
    expect((await run({ workbook: over4 }, 'desktop')).read).toMatchObject({ kind: 'unavailable', diagnostic: { code: 'chart-data-corrupt' } })

    // A forged (small) declared size: the counter, not the header, stops the inflate.
    const forged = await run({ workbook: over8, entry: { uncompressedSize: 1024 } }, 'desktop')
    expect(forged.read).toMatchObject({ kind: 'refused', diagnostic: { code: 'embedded-workbook-limit' } })
    expect(forged.peak).toBeLessThanOrEqual(8 * MiB + 256 * 1024)
    expect(forged.after).toBe(0)

    // The central-directory inventory, before any load: entry count.
    const inner = (count: number, extra: Partial<BuildEntry> = {}) => buildZip(Array.from({ length: count }, (_, i) => ({ name: `xl/p${i}.xml`, data: `<x>${i}</x>`, ...extra }))).bytes
    expect((await run({ workbook: inner(257) }, 'desktop')).read).toMatchObject({ kind: 'refused', diagnostic: { code: 'embedded-workbook-entries' }, limit: 256 })
    const at256 = await run({ workbook: inner(256) }, 'desktop')
    expect(at256.read.kind).toBe('workbook')
    if (at256.read.kind === 'workbook') expect(at256.read.inventory.entries).toHaveLength(256)
    expect(at256.after).toBe(0)

    // Declared-size sum (forged large in the inner directory; nothing inflates them).
    const declared = (each: number) => buildZip([
      { name: 'xl/workbook.xml', data: '<workbook/>', uncompressedSize: each },
      { name: 'xl/sharedStrings.xml', data: '<sst/>', uncompressedSize: each },
    ]).bytes
    expect((await run({ workbook: declared(17 * MiB) }, 'desktop')).read).toMatchObject({ kind: 'refused', diagnostic: { code: 'embedded-workbook-declared' }, limit: 32 * MiB })
    expect((await run({ workbook: declared(9 * MiB) }, 'mobile')).read).toMatchObject({ kind: 'refused', diagnostic: { code: 'embedded-workbook-declared' }, limit: 16 * MiB })
    expect((await run({ workbook: declared(9 * MiB) }, 'desktop')).read.kind).toBe('workbook')

    // Workspace refusal: nothing allocated past it, everything refunded.
    const pkg = await OfficePackage.load(docx({ workbook: await embeddedXlsx() }).bytes)
    const tight = { profile: 'desktop' as const, workspace: createChartWorkspace('desktop', { reserve: (bytes: number) => bytes < 1024, release: () => {} }) }
    const refused = await openEmbeddedWorkbook(chartParseContext(tight, pkg, 'word/charts/chart1.xml'), 'rId1')
    expect(refused).toMatchObject({ kind: 'refused', diagnostic: { code: 'chart-refused' } })
    expect(tight.workspace.used).toBe(0)
  })

  it('embedding bytes are not retained in the outer package cache', async () => {
    const workbook = await embeddedXlsx()
    const fixture = docx({ workbook })
    const pkg = await OfficePackage.load(fixture.bytes)
    const doc = documentChartContext('desktop')
    const bytesSpy = vi.spyOn(pkg, 'bytes'), textSpy = vi.spyOn(pkg, 'text')
    for (let round = 0; round < 2; round++) {
      const read = await openEmbeddedWorkbook(chartParseContext(doc, pkg, fixture.chartPaths[0]), 'rId1')
      expect(read.kind).toBe('workbook')
      if (read.kind !== 'workbook') return
      expect(read.bytes).toEqual(workbook)
      read.release()
    }
    const embedding = fixture.embeddingPaths[0]!
    expect(cacheKeys(pkg).filter(key => key.includes(embedding))).toEqual([])
    expect(bytesSpy.mock.calls.some(([path]) => path === embedding)).toBe(false)
    expect(textSpy.mock.calls.some(([path]) => path === embedding)).toBe(false)
    expect(doc.workspace.used).toBe(0)
  })

  it('one embedded read at a time per document', async () => {
    const workbook = await paddedWorkbook()
    const fixture = buildChartDocx([{ workbook }, { workbook }])
    const pkg = await OfficePackage.load(fixture.bytes)
    const doc = documentChartContext('desktop')
    gate.close()
    const first = openEmbeddedWorkbook(chartParseContext(doc, pkg, fixture.chartPaths[0]), 'rId1')
    const second = openEmbeddedWorkbook(chartParseContext(doc, pkg, fixture.chartPaths[1]), 'rId1')
    await vi.waitFor(() => expect(gate.calls).toBe(1))
    // Another document is not serialized behind this one.
    const other = await open(fixture)
    await vi.waitFor(() => expect(gate.calls).toBe(2))
    gate.open()
    const a = await first
    expect(a.kind).toBe('workbook')
    // The first read holds the document's slot until its caller releases it.
    await new Promise(resolve => setTimeout(resolve, 20))
    expect(gate.calls).toBe(2)
    if (a.kind === 'workbook') a.release()
    const b = await second
    expect(gate.calls).toBe(3)
    expect(b.kind).toBe('workbook')
    if (b.kind === 'workbook') { expect(b.path).toBe(fixture.embeddingPaths[1]); b.release() }
    const c = await other.result
    if (c.kind === 'workbook') c.release()
    expect(doc.workspace.used).toBe(0)
    expect(other.doc.workspace.used).toBe(0)
  })

  it('cancellation stops publication immediately; charges are held until the gated inflate settles (collect-ahead producer)', async () => {
    const workbook = await paddedWorkbook()
    const fixture = buildChartDocx([{ workbook }, { workbook }, { workbook }])
    const controller = new AbortController()
    const pkg = await OfficePackage.load(fixture.bytes)
    const doc = documentChartContext('desktop', controller.signal)
    gate.close()
    let settled = false
    const first = openEmbeddedWorkbook(chartParseContext(doc, pkg, fixture.chartPaths[0]), 'rId1').finally(() => { settled = true })
    const queued = openEmbeddedWorkbook(chartParseContext(doc, pkg, fixture.chartPaths[1]), 'rId1')
    await vi.waitFor(() => expect(gate.calls).toBe(1))
    const held = doc.workspace.used
    expect(held).toBeGreaterThan(0)
    controller.abort()
    // Publication stops at once, with the inflate still gated.
    expect(await first).toMatchObject({ kind: 'cancelled' })
    expect(settled).toBe(true)
    expect(await queued).toMatchObject({ kind: 'cancelled' })
    // The producer has not settled: its charges stay held.
    await new Promise(resolve => setTimeout(resolve, 20))
    expect(doc.workspace.used).toBe(held)
    gate.open()
    await vi.waitFor(() => expect(doc.workspace.used).toBe(0))
    // The queued read never started; a read under the aborted signal never starts.
    expect(gate.calls).toBe(1)
    expect(await openEmbeddedWorkbook(chartParseContext(doc, pkg, fixture.chartPaths[2]), 'rId1')).toMatchObject({ kind: 'cancelled' })
    expect(gate.calls).toBe(1)
    expect(doc.workspace.used).toBe(0)

  })

  it('M1: a producer paused mid-inflate holds only what it has produced; cancellation refunds once it settles', async () => {
    const workbook = await paddedWorkbook(256 * 1024)
    const fixture = docx({ workbook })
    const controller = new AbortController()
    const pkg = await OfficePackage.load(fixture.bytes)
    const doc = documentChartContext('desktop', controller.signal)
    gate.close('pause')
    const pending = openEmbeddedWorkbook(chartParseContext(doc, pkg, fixture.chartPaths[0]), 'rId1')
    await vi.waitFor(() => expect(gate.chunks).toBe(1))
    await new Promise(resolve => setTimeout(resolve, 30))
    // Really paused: no further chunk is produced, and only the crossing plus one chunk are charged.
    expect(gate.chunks).toBe(1)
    const held = doc.workspace.used
    expect(held).toBeGreaterThan(0)
    expect(held).toBeLessThanOrEqual(64 * 1024 + 64 * 1024)
    controller.abort()
    expect(await pending).toMatchObject({ kind: 'cancelled' })
    expect(doc.workspace.used).toBe(held)
    gate.open()
    await vi.waitFor(() => expect(doc.workspace.used).toBe(0))
    // The next chunk threw at the source: the inflate ended without producing the rest.
    expect(gate.chunks).toBeLessThanOrEqual(2)
  })

  it('I1: the chart .rels is read bounded, admitted and uncached', async () => {
    const relsPath = 'word/charts/_rels/chart1.xml.rels'
    const relsXml = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`
    const comment = `<!--${'x'.repeat(1024 * 1024)}-->`
    const external = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="https://example.com/book.xlsx" TargetMode="External"/>'
    const load = (xml: string) => OfficePackage.load(buildZip([{ name: relsPath, data: xml }]).bytes)
    const parentOf = (accept: (bytes: number) => boolean) => { const calls: number[] = []; return { calls, reserve: (bytes: number) => { calls.push(bytes); return accept(bytes) }, release: () => {} } }

    // A rejecting parent: refused before anything is decoded; nothing charged, nothing cached.
    const pkg = await load(relsXml(comment + external))
    const parent = parentOf(() => false)
    const rejecting = { profile: 'desktop' as const, workspace: createChartWorkspace('desktop', parent) }
    const read = await openEmbeddedWorkbook(chartParseContext(rejecting, pkg, 'word/charts/chart1.xml'), 'rId1')
    expect(read).toMatchObject({ kind: 'refused', diagnostic: { code: 'chart-refused' } })
    expect(parent.calls[0]).toBe(EMBEDDED_RELS_CROSSING)
    expect(rejecting.workspace.used).toBe(0)
    expect(cacheKeys(pkg)).toEqual([])
    expect(await resolveExternalData(chartParseContext(rejecting, pkg, 'word/charts/chart1.xml'), 'rId1')).toMatchObject({ kind: 'refused' })

    // An admitting workspace: the 1 MiB rels part passes the comment stripper but not the byte cap.
    const big = await load(relsXml(comment + external))
    const doc = documentChartContext('desktop')
    let peak = 0
    const reserve = doc.workspace.reserve.bind(doc.workspace)
    doc.workspace.reserve = (bytes, tag) => { const ok = reserve(bytes, tag); peak = Math.max(peak, doc.workspace.used); return ok }
    expect(await openEmbeddedWorkbook(chartParseContext(doc, big, 'word/charts/chart1.xml'), 'rId1')).toMatchObject({ kind: 'refused', diagnostic: { code: 'chart-rels-limit' }, limit: EMBEDDED_RELS_LIMIT })
    expect(peak).toBeGreaterThan(0)
    expect(peak).toBeLessThanOrEqual(EMBEDDED_RELS_CROSSING + 3 * 64 * 1024)
    expect(doc.workspace.used).toBe(0)
    expect(cacheKeys(big)).toEqual([])

    // A comment under the cap is stripped, not a token: the relationship after it is found (and still uncached).
    const commented = await load(relsXml(`<!--${'y'.repeat(100 * 1024)}-->` + external))
    expect(await openEmbeddedWorkbook(chartParseContext(documentChartContext('desktop'), commented, 'word/charts/chart1.xml'), 'rId1')).toMatchObject({ kind: 'unavailable', diagnostic: { code: 'chart-data-external' } })
    expect(cacheKeys(commented)).toEqual([])

    // Only the requested id is kept; an oversized attribute is a token refusal; malformed XML is diagnosed.
    const many = await load(relsXml(Array.from({ length: 200 }, (_, i) => `<Relationship Id="rX${i}" Target="../embeddings/x${i}.xlsx"/>`).join('') + external))
    expect(await resolveExternalData(chartParseContext(documentChartContext('desktop'), many, 'word/charts/chart1.xml'), 'rId1')).toMatchObject({ kind: 'unavailable', diagnostic: { code: 'chart-data-external' } })
    const longTarget = await load(relsXml(`<Relationship Id="rId1" Target="${'a/'.repeat(20 * 1024)}x.xlsx"/>`))
    expect(await resolveExternalData(chartParseContext(documentChartContext('desktop'), longTarget, 'word/charts/chart1.xml'), 'rId1')).toMatchObject({ kind: 'refused', diagnostic: { code: 'chart-rels-limit' } })
    const broken = await load('<Relationships><Relationship Id="rId1" Target="x.xlsx"></Relationships>')
    expect(await resolveExternalData(chartParseContext(documentChartContext('desktop'), broken, 'word/charts/chart1.xml'), 'rId1')).toMatchObject({ kind: 'unavailable', diagnostic: { code: 'chart-data-rels-malformed' } })
  })

  it('I2: an abort during relationship resolution returns cancelled at once; its charges are held until it settles', async () => {
    const workbook = await embeddedXlsx()
    for (const spec of [{ workbook }, { target: 'https://example.com/book.xlsx', external: true }] as ChartEmbedSpec[]) {
      gate.reset()
      const fixture = docx(spec)
      const controller = new AbortController()
      const pkg = await OfficePackage.load(fixture.bytes)
      const doc = documentChartContext('desktop', controller.signal)
      gate.close('pause', 'rels')
      let settled = false
      const pending = openEmbeddedWorkbook(chartParseContext(doc, pkg, fixture.chartPaths[0]), 'rId1').finally(() => { settled = true })
      await vi.waitFor(() => expect(gate.chunks).toBe(1))
      const held = doc.workspace.used
      expect(held).toBeGreaterThanOrEqual(EMBEDDED_RELS_CROSSING)
      controller.abort()
      await new Promise(resolve => setTimeout(resolve, 0))
      expect(settled).toBe(true)
      expect(await pending).toMatchObject({ kind: 'cancelled' })
      expect(doc.workspace.used).toBe(held)
      gate.open()
      await vi.waitFor(() => expect(doc.workspace.used).toBe(0))
      // Never published late, never cached, and the embedding was never inflated.
      expect(await pending).toMatchObject({ kind: 'cancelled' })
      expect(cacheKeys(pkg)).toEqual([])
    }
  })

  it('a corrupt embedded zip is unavailable with a diagnostic', async () => {
    const cases: ChartEmbedSpec[] = [
      { workbook: noise(4096) },
      { workbook: new Uint8Array(0) },
      // Truncated: the end of the central directory is cut off.
      { workbook: (await embeddedXlsx()).subarray(0, 200) },
      // The outer entry's compressed stream is not deflate.
      { entry: { raw: noise(2048, 7), method: 8, uncompressedSize: 4096 } },
      // An inner directory that names an escaping part.
      { workbook: buildZip([{ name: '../evil.xml', data: '<x/>' }]).bytes },
    ]
    for (const spec of cases) {
      const { doc, result } = await open(docx(spec))
      const read = await result
      expect(read).toMatchObject({ kind: 'unavailable', diagnostic: { code: 'chart-data-corrupt' } })
      if (read.kind === 'unavailable') expect(read.diagnostic.reason).toMatch(/\w/)
      expect(doc.workspace.used).toBe(0)
    }
  })
})
