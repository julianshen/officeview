/**
 * S1.1b entry-point wiring (design note §4.5; RED tests 15b, 16, 17, 18, the
 * entry parts of 15a, 19, 20, 23 and the small-document parity of 24).
 */
import { readFileSync, readdirSync, existsSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, render, waitFor } from '@testing-library/react'
import { createElement } from 'react'
import JSZip from 'jszip'
import { loadOfficeFile, OfficeFile, useOfficeFile } from '../src/components/OfficeFile'
import { detectAndParse, loadOfficeSource, type LoadSourceOptions } from '../src/components/loadOfficeSource'
import { OfficePackage } from '../src/core/zip'
import { readByteStream, PROTECTION_HEADER, type Progress } from '../src/core/stream'
import { sniffPackage } from '../src/core/package-sniff'
import { bytesRangeSource, readZipInventory } from '../src/core/zip-stream'
import { PackageError } from '../src/core/package-error'
import { ImportRefusal } from '../src/xlsx/storage/importSession'
import type { SpooledSource } from '../src/core/spool'
import { buildDocx, buildPptx, buildXlsx } from '../src/testdata/ooxml-builders'
import { buildZip } from './helpers/zip-builder'
import { fakeBackend, fakeOpfsDirectory } from './helpers/fake-storage'
import { metaRows, valueCells, xlsxBytes, xlsxEntries } from './helpers/xlsx-package'

const KiB = 1024
const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')

function chunked(bytes: Uint8Array, size: number, log?: { pulls: number }): ReadableStream<Uint8Array> {
  let at = 0
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      if (log) log.pulls++
      if (at >= bytes.length) { controller.close(); return }
      controller.enqueue(bytes.slice(at, at + size))
      at += size
    },
  }, { highWaterMark: 0 })
}

async function refusal(p: Promise<unknown>): Promise<ImportRefusal> {
  try { await p } catch (e) { if (e instanceof ImportRefusal) return e; throw e }
  throw new Error('expected an ImportRefusal')
}

/** A spreadsheet bigger than the 64 KiB sniff prefix, so sequential input is spooled. */
const MID_XLSX = () => {
  // pseudo-random strings: still above 64 KiB once deflated
  let x = 12345
  const word = () => Array.from({ length: 20 }, () => { x = (x * 1103515245 + 12345) >>> 0; return String.fromCharCode(97 + (x >>> 27)) }).join('')
  return xlsxBytes({ sheetData: valueCells(500), sharedStrings: Array.from({ length: 16_000 }, word) })
}
/** 125k metadata rows: paged on both profiles (the §4.3 counterexample). */
const BIG_XLSX = () => xlsxBytes({ sheetData: metaRows(125_000) })

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})

describe('sniffPackage hint (RED test 15b)', () => {
  it('classifies only from a complete [Content_Types].xml inside the prefix', async () => {
    expect(sniffPackage(await buildXlsx([{ name: 'S', rows: [] }]))).toBe('xlsx')
    expect(sniffPackage(await buildDocx([{ runs: [{ text: 'x' }] }]))).toBe('docx')
    expect(sniffPackage(await buildPptx([{ prst: 'rect' }]))).toBe('pptx')
    expect(sniffPackage(new TextEncoder().encode('{\\rtf1 hi}'))).toBe('not-zip')
  })

  it('a first deflated _rels/.rels with a data descriptor, and a stored content-types part under a descriptor, still give xlsx', () => {
    const entries = xlsxEntries({})
    const rels = entries.find((e) => e.name === '_rels/.rels')!
    const ct = entries.find((e) => e.name === '[Content_Types].xml')!
    const rest = entries.filter((e) => e !== rels && e !== ct)
    expect(sniffPackage(buildZip([{ ...rels, method: 8, dataDescriptor: true }, { ...ct, method: 8 }, ...rest]).bytes)).toBe('xlsx')
    expect(sniffPackage(buildZip([{ ...ct, method: 0, dataDescriptor: true }, ...rest]).bytes)).toBe('xlsx')
  })

  it('adversarial (a): a 100,000-character generic part before the content types is undetermined', () => {
    const entries = xlsxEntries({})
    const custom = { name: 'docProps/custom.xml', data: `<Properties>${'p'.repeat(100_000)}</Properties>`, method: 0, dataDescriptor: false }
    expect(sniffPackage(buildZip([custom, ...entries]).bytes.subarray(0, 64 * KiB))).toBe('undetermined')
  })

  it('adversarial (b): a DOCX whose first part is xl/printerSettings is still docx (part names never classify)', async () => {
    const docx = await JSZip.loadAsync(await buildDocx([{ runs: [{ text: 'printer' }] }]))
    const out = new JSZip()
    out.file('xl/printerSettings/printerSettings1.bin', new Uint8Array(64))
    for (const [name, file] of Object.entries(docx.files)) if (!file.dir) out.file(name, await file.async('uint8array'))
    const bytes = await out.generateAsync({ type: 'uint8array' })
    expect(sniffPackage(bytes)).toBe('docx')
    const doc = await loadOfficeFile(chunked(bytes, 4 * KiB))
    expect('sections' in doc).toBe(true)
  })
})

describe('classification and routing (RED test 15b)', () => {
  it('remote DOCX and PPTX above the memory cap with no storage backend load exactly as today, no spool created', async () => {
    const opfs = fakeBackend('opfs', { available: false })
    const options: LoadSourceOptions = { storage: { memoryLimitBytes: KiB, opfs }, limits: { spool: { memoryLimitBytes: KiB, hardMemoryCapBytes: 2 * KiB } } }
    const docx = await buildDocx([{ runs: [{ text: 'big remote '.repeat(20_000) }] }])
    expect(docx.length).toBeGreaterThan(64 * KiB)
    const doc = await loadOfficeSource(chunked(docx, 8 * KiB), options)
    expect(doc).toEqual(await detectAndParse(await OfficePackage.load(docx)))
    const pptx = await buildPptx([{ prst: 'rect', text: 'p'.repeat(100_000) } as never])
    const slides = await loadOfficeSource(chunked(pptx, 8 * KiB), options)
    expect('slides' in slides).toBe(true)
    expect(opfs.log.created).toEqual([])
  })

  it('adversarial (a) is spooled and admitted as a spreadsheet, never loaded through the whole-buffer legacy path', async () => {
    const entries = xlsxEntries({ sheetData: metaRows(125_000) }).map((e) => ({ method: 0, ...e }))
    const bytes = buildZip([{ name: 'docProps/custom.xml', data: `<Properties>${'p'.repeat(100_000)}</Properties>`, method: 0, dataDescriptor: false }, ...entries]).bytes
    const load = vi.spyOn(OfficePackage, 'load')
    const spooled: SpooledSource[] = []
    const err = await refusal(loadOfficeSource(chunked(bytes, 32 * KiB), { onSpooled: (s) => spooled.push(s) }))
    expect(err.reason).toBe('paged-not-yet')
    expect(spooled).toHaveLength(1)
    expect(load).not.toHaveBeenCalled()

    // no storage and over the hard memory cap: the typed "could not determine" refusal
    const none = await refusal(loadOfficeSource(chunked(bytes, 32 * KiB), {
      storage: { memoryLimitBytes: 256 * KiB, hardMemoryCapBytes: 512 * KiB },
    }))
    expect(none.reason).toBe('undetermined-format')
    expect(none.code).toBe('storage-error')
    expect(none.message).toMatch(/could not be determined within the first 65536 bytes/)
    expect(load).not.toHaveBeenCalled()
  })

  it('a disguised spreadsheet (duplicate [Content_Types].xml, local-first DOCX) follows the pinned chain and never reaches parseXlsx', async () => {
    const docxCt = (await JSZip.loadAsync(await buildDocx([{ runs: [{ text: 'd' }] }]))).file('[Content_Types].xml')!.async('string')
    const entries = xlsxEntries({})
    const bytes = buildZip([{ name: '[Content_Types].xml', data: await docxCt }, ...entries]).bytes
    expect(sniffPackage(bytes)).toBe('docx')
    const xml = vi.spyOn(OfficePackage.prototype, 'xml')
    const err = await refusal(loadOfficeFile(chunked(bytes, 4 * KiB)))
    expect(err.code).toBe('parse-error')
    expect(err.message).toMatch(/duplicate part name \[Content_Types\]\.xml/)
    // the legacy view saw a spreadsheet (later duplicate wins) but parseXlsx never read the workbook
    expect(xml.mock.calls.map((c) => c[0])).toContain('[Content_Types].xml')
    expect(xml.mock.calls.map((c) => c[0])).not.toContain('xl/workbook.xml')
  })

  it('the in-repo reader rejects duplicate part names, not only the content types', async () => {
    const bytes = buildZip([...xlsxEntries({}), { name: 'xl/styles.xml', data: '<styleSheet/>' }]).bytes
    await expect(readZipInventory(bytesRangeSource(bytes), { maxEntries: 99, maxCentralDirectoryBytes: 1e6, maxNameBytes: 999, maxPartExpandedBytes: 1e9, maxTotalExpandedBytes: 1e9, maxRatio: 200 }))
      .rejects.toMatchObject({ code: 'parse-error' })
  })

  it('a seekable spreadsheet above the object caps is refused without reading the whole Blob (RED test 20)', async () => {
    const bytes = BIG_XLSX()
    const blob = new Blob([bytes as BlobPart])
    const sizes: number[] = []
    const original = Blob.prototype.arrayBuffer
    vi.spyOn(Blob.prototype, 'arrayBuffer').mockImplementation(function (this: Blob) {
      sizes.push(this.size)
      return original.call(this)
    })
    const load = vi.spyOn(OfficePackage, 'load')
    const err = await refusal(loadOfficeFile(blob))
    expect(err.reason).toBe('paged-not-yet')
    expect(err.message).toMatch(/rows|admission total/)
    expect(load).not.toHaveBeenCalled()
    expect(sizes.length).toBeGreaterThan(0)
    expect(Math.max(...sizes)).toBeLessThanOrEqual(Math.max(262_144, 65_557))
    expect(sizes).not.toContain(bytes.length)
  })

  it('a refused Blob above the paged caps names the cap (typed resource-limit, no blank document)', async () => {
    const err = await refusal(loadOfficeSource(new Blob([MID_XLSX() as BlobPart]), { limits: { paged: { compressedBytes: 1000 } } }))
    expect(err.code).toBe('resource-limit')
    expect(err.reason).toBe('limit')
    expect(err.message).toMatch(/package size \d+ exceeds the 1000 limit/)
  })
})

describe('files the old path opened still open (orchestrator decisions)', () => {
  it('decision 4: a highly compressed small spreadsheet opens, with a ratio diagnostic only', async () => {
    vi.spyOn(Date, 'now').mockReturnValue(Date.UTC(2026, 9, 11, 12))
    const bytes = xlsxBytes({ sheetData: valueCells(10), afterSheetData: `<extLst><ext uri="x">${'a'.repeat(1_500_000)}</ext></extLst>` })
    const legacy = await detectAndParse(await OfficePackage.load(bytes)) as { sheets: unknown[] }
    for (const source of [bytes, new Blob([bytes as BlobPart])]) {
      const doc = await loadOfficeFile(source) as { sheets: unknown[]; diagnostics?: Array<{ feature: string }> }
      expect(doc.sheets).toEqual(legacy.sheets)
      expect(doc.diagnostics?.some((d) => d.feature === 'high-compression-ratio')).toBe(true)
    }
  })
})

describe('decision 5: mc:AlternateContent-wrapped data in a small file', () => {
  const mc = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
  const bytes = () => buildZip(xlsxEntries({}).map((e) => e.name === 'xl/worksheets/sheet1.xml'
    ? { ...e, data: `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ${mc}><cols><col min="1" max="1" width="20" customWidth="1"/></cols><mc:AlternateContent><mc:Choice Requires="x"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></mc:Choice></mc:AlternateContent></worksheet>` }
    : e)).bytes

  it('renders as today and carries a visible diagnostic', async () => {
    vi.spyOn(Date, 'now').mockReturnValue(Date.UTC(2026, 9, 11, 12))
    const b = bytes()
    const legacy = await detectAndParse(await OfficePackage.load(b)) as { sheets: unknown[] }
    const doc = await loadOfficeFile(b) as { sheets: unknown[]; diagnostics?: Array<{ feature: string; message: string }> }
    expect(doc.sheets).toEqual(legacy.sheets)
    expect(doc.diagnostics?.find((d) => d.feature === 'alternate-content-data')?.message).toMatch(/xl\/worksheets\/sheet1.xml/)
    let seen: ReturnType<typeof useOfficeFile> | undefined
    function Probe() { seen = useOfficeFile(b); return null }
    render(createElement(Probe))
    await waitFor(() => expect(seen?.status).toBe('ready'), { timeout: 30_000 })
    expect(seen?.status === 'ready' && seen.notices?.[0]).toMatch(/mc:AlternateContent/)
    const { container } = render(createElement(OfficeFile, { data: b }))
    await waitFor(() => expect(container.querySelector('[role="status"]')?.textContent).toMatch(/not shown/), { timeout: 30_000 })
  })
})

describe('formula expansion through the entry point (RED test 19)', () => {
  it('uncached SEQUENCE(200000,1) is evaluated with the caps, hits array-resource-limit and is rerouted before publishing', async () => {
    const bytes = xlsxBytes({ sheetData: '<row r="1"><c r="A1"><f>SEQUENCE(200000,1)</f></c></row>' })
    const started = Date.now()
    const err = await refusal(loadOfficeFile(bytes))
    expect(err.reason).toBe('paged-not-yet')
    expect(err.message).toMatch(/array-resource-limit/)
    expect(Date.now() - started).toBeLessThan(10_000)
  })

  it('the same formula with a cached value stays object (engine policy) and loads', async () => {
    const bytes = xlsxBytes({ sheetData: '<row r="1"><c r="A1"><f>SEQUENCE(200000,1)</f><v>1</v></c></row>' })
    const doc = await loadOfficeFile(bytes)
    expect('sheets' in doc && doc.sheets[0].rows[0].cells[0].value).toBe(1)
  })

  it('when a G2 counter trips the legacy parser is not invoked', async () => {
    const load = vi.spyOn(OfficePackage, 'load')
    const err = await refusal(loadOfficeFile(new Blob([xlsxBytes({ sharedStrings: Array.from({ length: 60_000 }, (_, i) => `s${i}`) }) as BlobPart])))
    expect(err.message).toMatch(/unique shared strings 50001 > 50000/)
    expect(load).not.toHaveBeenCalled()
  })
})

describe('spool through the entry point (RED tests 15a, 16, 17, 18)', () => {
  it('a mid-size remote spreadsheet is spooled, admitted as an object and the spool is released after loading', async () => {
    vi.spyOn(Date, 'now').mockReturnValue(Date.UTC(2026, 9, 11, 12)) // the parse records a clock sample
    const opfs = fakeBackend('opfs')
    const bytes = MID_XLSX()
    const doc = await loadOfficeSource(chunked(bytes, 16 * KiB), { storage: { memoryLimitBytes: 32 * KiB, opfs } })
    expect('sheets' in doc).toBe(true)
    expect(opfs.log.created).toHaveLength(1)
    expect(opfs.log.live.size).toBe(0)
    expect(doc).toEqual(await detectAndParse(await OfficePackage.load(bytes)))
  })

  it('backpressure: while an append is pending the source is not read further (at most one chunk ahead)', async () => {
    const holds: Array<() => void> = []
    let accepted = 0
    const opfs = fakeBackend('opfs', { gate: () => new Promise<void>((r) => holds.push(() => { accepted++; r() })) })
    const log = { pulls: 0 }
    const bytes = MID_XLSX()
    const done = loadOfficeSource(chunked(bytes, 8 * KiB, log), { storage: { memoryLimitBytes: 8 * KiB, opfs } })
    for (let i = 0; i < 200 && holds.length === 0; i++) await new Promise((r) => setTimeout(r, 1))
    let maxAhead = 0
    while (true) {
      await new Promise((r) => setTimeout(r, 1))
      maxAhead = Math.max(maxAhead, log.pulls - accepted)
      const next = holds.shift()
      if (!next) break
      next()
    }
    await done
    // pulls include the prefix chunks (64 KiB / 8 KiB = 8) appended in one go and the final close pull
    expect(maxAhead).toBeLessThanOrEqual(8 + 1)
    expect(opfs.log.live.size).toBe(0)
  })

  it('quota denial mid-write is a typed quota refusal and the partial spool is removed', async () => {
    const opfs = fakeBackend('opfs', { failOnWrite: { at: 3, error: new DOMException('full', 'QuotaExceededError') } })
    const err = await refusal(loadOfficeSource(chunked(MID_XLSX(), 8 * KiB), { storage: { memoryLimitBytes: 8 * KiB, opfs } }))
    expect(err.code).toBe('quota')
    expect(opfs.log.live.size).toBe(0)
  })

  it('a download error mid-spool removes the spool', async () => {
    const opfs = fakeBackend('opfs')
    const bytes = MID_XLSX()
    let at = 0
    const failing = new ReadableStream<Uint8Array>({
      pull(c) {
        if (at > 80 * KiB) { c.error(new Error('connection reset')); return }
        c.enqueue(bytes.slice(at, at + 8 * KiB))
        at += 8 * KiB
      },
    }, { highWaterMark: 0 })
    await expect(loadOfficeSource(failing, { storage: { memoryLimitBytes: 8 * KiB, opfs } })).rejects.toThrow('connection reset')
    expect(opfs.log.created).toHaveLength(1)
    expect(opfs.log.live.size).toBe(0)
  })

  it('cancellation during download cancels the reader, aborts and removes the spool', async () => {
    const opfs = fakeBackend('opfs')
    const bytes = MID_XLSX()
    const ac = new AbortController()
    let cancelled = false
    let at = 0
    const stream = new ReadableStream<Uint8Array>({
      async pull(c) {
        if (at >= 80 * KiB) ac.abort()
        await new Promise((r) => setTimeout(r, 1))
        c.enqueue(bytes.slice(at, at + 8 * KiB))
        at += 8 * KiB
      },
      cancel() { cancelled = true },
    }, { highWaterMark: 0 })
    const err = await loadOfficeSource(stream, { signal: ac.signal, storage: { memoryLimitBytes: 8 * KiB, opfs } }).catch((e: unknown) => e)
    expect((err as PackageError).code).toBe('cancelled')
    expect(cancelled).toBe(true)
    expect(opfs.log.live.size).toBe(0)
  })

  it('a refusal after finish() releases the spool', async () => {
    const opfs = fakeBackend('opfs')
    await refusal(loadOfficeSource(chunked(BIG_XLSX(), 64 * KiB), { storage: { memoryLimitBytes: 64 * KiB, opfs } }))
    expect(opfs.log.created).toHaveLength(1)
    expect(opfs.log.live.size).toBe(0)
  })

  it('headers, credentials and the Request are forwarded exactly as today on the spool path; protection fires before the first byte', async () => {
    const bytes = MID_XLSX()
    const calls: Array<[unknown, RequestInit | undefined]> = []
    const order: string[] = []
    vi.stubGlobal('fetch', async (input: unknown, init?: RequestInit) => {
      calls.push([input, init])
      return new Response(chunked(bytes, 16 * KiB), { headers: { 'content-length': String(bytes.length), [PROTECTION_HEADER]: 'no-copy' } })
    })
    const doc = await loadOfficeFile({ url: 'https://files.example.com/a.xlsx', headers: { Authorization: 'Bearer t' }, credentials: 'include' }, {
      onProtection: (p) => order.push(`protection:${p.allowCopy}`),
      onProgress: (p) => order.push(`progress:${p.loaded}`),
    })
    expect('sheets' in doc).toBe(true)
    expect(calls[0][0]).toBe('https://files.example.com/a.xlsx')
    expect(calls[0][1]).toEqual({ headers: { Authorization: 'Bearer t' }, credentials: 'include' })
    expect(order[0]).toBe('protection:false')
    expect(order[1]).toBe('progress:0')

    const req = new Request('https://files.example.com/b.xlsx', { headers: { Authorization: 'Bearer r' } })
    await loadOfficeFile(req)
    expect(calls[1][0]).toBe(req)
    expect(calls[1][1]).toBeUndefined()
  })

  it('progress throttling on the spool path is the same as readByteStream', async () => {
    const bytes = MID_XLSX()
    const expected: Progress[] = []
    await readByteStream(chunked(bytes, 4 * KiB), (p) => expected.push({ ...p }), bytes.length)
    const seen: Progress[] = []
    vi.stubGlobal('fetch', async () => new Response(chunked(bytes, 4 * KiB), { headers: { 'content-length': String(bytes.length) } }))
    await loadOfficeFile({ url: 'https://x.example/a.xlsx' }, { onProgress: (p) => seen.push({ ...p }) })
    expect(seen).toEqual(expected)
  })
})

describe('<OfficeFile> ownership (RED tests 15a, 18, 23)', () => {
  function stubOpfs() {
    const dir = fakeOpfsDirectory()
    Object.defineProperty(navigator, 'storage', { configurable: true, value: { getDirectory: async () => dir } })
    vi.stubGlobal('FileSystemFileHandle', class { createWritable() {} })
    return { dir, restore: () => { delete (navigator as unknown as { storage?: unknown }).storage } }
  }

  it('a replaced source cancels the stale download, drops its late chunks and removes its spool; only the live owner publishes', async () => {
    const { dir, restore } = stubOpfs()
    try {
      const stale = MID_XLSX()
      let staleCancelled = false
      let controller!: ReadableStreamDefaultController<Uint8Array>
      const slow = new ReadableStream<Uint8Array>({ start(c) { controller = c }, cancel() { staleCancelled = true } })
      const big = { 'content-length': String(64 * 1024 * 1024) }
      vi.stubGlobal('fetch', async (url: string) => (url.endsWith('stale.xlsx')
        ? new Response(slow, { headers: big })
        : new Response(chunked(MID_XLSX(), 16 * KiB), { headers: big })))
      const states: string[] = []
      function Probe({ src }: { src: { url: string } }) {
        const state = useOfficeFile(src)
        states.push(state.status === 'ready' ? `ready:${src.url}` : state.status)
        return null
      }
      const a = { url: 'https://x.example/stale.xlsx' }
      const { rerender } = render(createElement(Probe, { src: a }))
      await act(async () => { controller.enqueue(stale.slice(0, 80 * KiB)) })
      await waitFor(() => expect(dir.files.size).toBe(1))
      const b = { url: 'https://x.example/live.xlsx' }
      rerender(createElement(Probe, { src: b }))
      await act(async () => { try { controller.enqueue(stale.slice(80 * KiB)); controller.close() } catch { /* cancelled */ } })
      await waitFor(() => expect(states[states.length - 1]).toBe('ready:https://x.example/live.xlsx'), { timeout: 30_000 })
      expect(staleCancelled).toBe(true)
      expect(states).not.toContain('ready:https://x.example/stale.xlsx')
      await waitFor(() => expect(dir.files.size).toBe(0))
    } finally {
      restore()
    }
  })

  it('a refused spreadsheet surfaces a typed error state and no document', async () => {
    let seen: ReturnType<typeof useOfficeFile> | undefined
    function Probe({ data }: { data: Blob }) {
      seen = useOfficeFile(data)
      return null
    }
    render(createElement(Probe, { data: new Blob([BIG_XLSX() as BlobPart]) }))
    await waitFor(() => expect(seen?.status).toBe('error'), { timeout: 30_000 })
    expect(seen).toMatchObject({ status: 'error', code: 'unsupported', reason: 'paged-not-yet' })
    const { container } = render(createElement(OfficeFile, { data: new Blob([BIG_XLSX() as BlobPart]) }))
    await waitFor(() => expect(container.querySelector('[role="alert"]')).not.toBeNull(), { timeout: 30_000 })
    expect(container.querySelectorAll('canvas')).toHaveLength(0)
  })
})

describe('small-document parity (RED test 24)', () => {
  const files: Array<[string, () => Uint8Array]> = []
  for (const dir of ['corpus', 'tests/fixtures/xlsx-viewer']) {
    const full = join(ROOT, dir)
    if (!existsSync(full)) continue
    for (const name of readdirSync(full)) {
      if (/\.(xlsx|docx|pptx)$/.test(name)) files.push([`${dir}/${name}`, () => new Uint8Array(readFileSync(join(full, name)))])
    }
  }
  files.push(['generated xlsx', () => xlsxBytes({ sheetData: valueCells(300), sharedStrings: ['a', 'b'] })])

  /** Volatile functions read the clock; freeze it so two parses compare equal. */
  const frozen = () => vi.spyOn(Date, 'now').mockReturnValue(Date.UTC(2026, 9, 11, 12))

  for (const [name, read] of files) {
    it(`${name} loads exactly as the legacy path through bytes, Blob and a chunked stream`, async () => {
      frozen()
      const bytes = read()
      const legacy = await OfficePackage.load(bytes).then((pkg) => detectAndParse(pkg)).then((d) => ({ ok: true as const, d }), (e: unknown) => ({ ok: false as const, e: String(e) }))
      for (const source of [() => bytes, () => new Blob([bytes as BlobPart]), () => chunked(bytes, 4 * KiB)]) {
        const got = await loadOfficeFile(source()).then((d) => ({ ok: true as const, d }), (e: unknown) => ({ ok: false as const, e: String(e) }))
        if (!legacy.ok) {
          expect(got.ok).toBe(false)
          continue
        }
        expect(got.ok).toBe(true)
        if (got.ok) expect(got.d).toEqual(legacy.d)
      }
    }, 60_000)
  }
})
