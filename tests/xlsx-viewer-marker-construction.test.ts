import { describe, expect, it } from 'vitest'
import { XmlTokenizer, tokenizeBytes } from '../src/core/xml-stream'
import { anchorRecordEvents } from '../src/xlsx/storage/sheetRecords'
import { anchorRecord, type AnchorParts } from '../src/xlsx/anchorRecords'
import { PackageError } from '../src/core/package-error'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, memoryPageBackend } from '../src/xlsx/storage/sourcePages'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { MemoryPool, WORKER_BUDGETS } from '../src/xlsx/view/tileCache'
import { DocumentSession } from '../src/worker/viewerSession'
import { buildZip } from './helpers/zip-builder'
import { xlsxEntries } from './helpers/xlsx-package'

const MiB = 1048576
const giant = '0'.repeat(2 * MiB) + '1'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const part = 'xl/drawings/drawing1.xml'
function drawingFixture(drawing: string, sharedStrings?: string[], method?: number) {
  return buildZip(xlsxEntries({ sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>', sharedStrings,
    afterSheetData: `<drawing xmlns:r="${R}" r:id="rD"/>`, extra: [
      { name: 'xl/worksheets/_rels/sheet1.xml.rels', data: `<Relationships><Relationship Id="rD" Type="${R}/drawing" Target="../drawings/drawing1.xml"/></Relationships>` },
      { name: part, data: drawing, method },
    ],
  }))
}
function fixture(col: string, sharedStrings?: string[]) {
  return drawingFixture(`<wsDr><oneCellAnchor><from><col>${col}</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from><ext cx="609600" cy="476250"/></oneCellAnchor></wsDr>`, sharedStrings).bytes
}
async function run(col: string, mode: 'provisional' | 'bound' | 'document' = 'provisional', profile: 'D' | 'M' = 'D', budget = MiB, sharedStrings?: string[]) {
  const bytes = fixture(col, sharedStrings)
  const store = new SourcePageStore({ docId: 'marker', backend: memoryPageBackend(), memoryBudgetBytes: mode === 'document' ? WORKER_BUDGETS[profile] : budget })
  const errors: Array<{ code?: string; reason?: string }> = []
  const conversions: Array<{ chars: number; transient: number }> = []
  const nativeNumber = globalThis.Number
  const write = XmlTokenizer.prototype.write
  let session: Awaited<ReturnType<typeof openImportSession>> | undefined
  let document: DocumentSession | undefined
  let source: ReturnType<typeof createPagedWorkbookSource> | undefined
  let bound = false
  globalThis.Number = new Proxy(nativeNumber, { apply(target, receiver, args) {
    if (typeof args[0] === 'string' && args[0].length === giant.length) conversions.push({ chars: args[0].length, transient: store.construction.transientBytes })
    return Reflect.apply(target, receiver, args)
  } })
  XmlTokenizer.prototype.write = function(chunk) {
    try { return write.call(this, chunk) } catch (error) {
      errors.push(error as { code?: string; reason?: string })
      throw error
    }
  }
  try {
    session = await openImportSession({ kind: 'blob', blob: new Blob([bytes as BlobPart]) }, { docId: 'marker', profile, store, onCommit: () => {
      if (mode === 'provisional' || bound) return
      bound = true
      if (mode === 'document') { source = createPagedWorkbookSource(session!); document = new DocumentSession('marker', source, { profile }) }
      else store.bindMemory(new MemoryPool({ limitBytes: budget }))
    } })
    const outcome = await session.done
    source ??= createPagedWorkbookSource(session)
    const sheet = source.store('1'), options = { signal: new AbortController().signal, stamp: sheet.stamp }
    const cell = await sheet.readCell({ row: 0, col: 0 }, options)
    const structural = await sheet.queryStructural({ r0: 0, r1: 1048575, c0: 0, c1: 16383 }, options)
    expect(store.construction.transientBytes).toBe(0)
    return { outcome, cell, structural, coverage: source.coverageDiagnostics?.('1'), anchors: session.workbook.sheets[0].drawingAnchors, diagnostics: session.diagnostics, errors, conversions }
  } finally {
    globalThis.Number = nativeNumber
    XmlTokenizer.prototype.write = write
    await source?.release()
    await session?.release()
    document?.memory.close()
    expect(store.construction.transientBytes).toBe(0)
    expect(store.construction.persistentBytes).toBe(0)
  }
}

describe('streamed drawing marker construction ownership', () => {
  it.each(['provisional', 'bound'] as const)('ordinary marker under the real 1 MiB %s budget preserves A1 and placement', async (mode) => {
    const r = await run('1', mode)
    expect(r.outcome.kind).toBe('complete')
    expect(r.cell).toMatchObject({ kind: 'cell', cell: { value: 1 } })
    expect(r.structural.completeness).toBe('complete')
    expect(r.anchors.records[0]).toMatchObject({ from: { col: 1 } })
    expect(r.errors).toEqual([])
  })
  it.each(['provisional', 'bound'] as const)('refuses giant before Number under the real 1 MiB %s budget while A1 remains readable', async (mode) => {
    const r = await run(giant, mode)
    expect(r.errors).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'resource-limit', reason: 'memory-budget' })]))
    expect(r.conversions).toEqual([])
    expect(r.outcome.kind).toBe('complete')
    expect(r.cell).toMatchObject({ kind: 'cell', cell: { value: 1 } })
    expect(r.structural.completeness).toBe('incomplete')
    expect(r.structural.missing.length).toBeGreaterThan(0)
    expect(r.coverage).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'coverage-unresolvable' })]))
    expect(r.anchors).toMatchObject({ records: [], unresolved: 1 })
    expect(r.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'import-memory-budget' })]))
  })
  it('the same supported budget really refuses pending SST ownership', async () => {
    const r = await run('1', 'provisional', 'D', MiB, Array.from({ length: 600 }, (_, i) => 'x'.repeat(1000) + i))
    expect(r.outcome).toMatchObject({ kind: 'incomplete', failure: { code: 'resource-limit' } })
    expect(r.errors).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'resource-limit', reason: 'memory-budget' })]))
  })
  it.each(['D', 'M'] as const)('admitted giant owns its full field through actual Number after default DocumentSession %s binding', async (profile) => {
    const r = await run(giant, 'document', profile)
    expect(r.outcome.kind).toBe('complete')
    expect(r.conversions).toHaveLength(2)
    for (const c of r.conversions) expect(c.transient).toBeGreaterThan(2 * c.chars)
    expect(r.anchors.records[0]).toMatchObject({ from: { col: 1 } })
    expect(r.structural.completeness).toBe('complete')
  })
})

const startField = '<wsDr><oneCellAnchor><from><col>'
const endField = '</col></from></oneCellAnchor></wsDr>'
function consumer(limit = Infinity, onAnchor: (name: string, parts: AnchorParts) => void = () => {}) {
  let bytes = 0
  const requests: number[] = []
  const events = anchorRecordEvents(onAnchor, () => {}, { need(next) {
    requests.push(next)
    if (next > limit) throw new PackageError('resource-limit', 'test admission', undefined, 'memory-budget')
    bytes = next
  } })
  return { events, tokenizer: new XmlTokenizer(events), requests, get bytes() { return bytes } }
}

describe('marker pre-retention and lifecycle', () => {
  it('consumer and anchor holder construction admits exact fit and refuses one-under before publication', () => {
    expect(() => consumer(255)).toThrow(expect.objectContaining({ reason: 'memory-budget' }))
    const oracle = consumer()
    oracle.tokenizer.write('<wsDr><oneCellAnchor/>')
    const peak = Math.max(...oracle.requests)
    let calls = 0
    const exact = consumer(peak, () => { calls++ })
    exact.tokenizer.write('<wsDr><oneCellAnchor/></wsDr>')
    expect(calls).toBe(1)
    const refused = consumer(peak - 1, () => { calls++ })
    refused.tokenizer.write('<wsDr>')
    expect(() => refused.events.open('oneCellAnchor', [], [], 0, true)).toThrow(expect.objectContaining({ reason: 'memory-budget' }))
    refused.tokenizer.write('</wsDr>')
    expect(calls).toBe(1)
    oracle.events.dispose(); exact.events.dispose(); refused.events.dispose()
  })
  it('concat admits the old and replacement text together, and a one-under refusal retains only the old text', () => {
    const oracle = consumer()
    oracle.tokenizer.write(startField)
    oracle.events.text('000', true)
    const old = oracle.bytes
    oracle.events.text('1', true)
    const peak = Math.max(...oracle.requests)
    expect(peak).toBeGreaterThanOrEqual(old + 48 + 2 * 4)
    let raw = ''
    const exact = consumer(peak, (_, parts) => { raw = parts.from!.col })
    exact.tokenizer.write(startField)
    exact.events.text('000', true)
    exact.events.text('1', true)
    exact.tokenizer.write(endField)
    expect(raw).toBe('0001')
    const refused = consumer(peak - 1, (_, parts) => { raw = parts.from!.col })
    refused.tokenizer.write(startField)
    refused.events.text('000', true)
    expect(() => refused.events.text('1', true)).toThrow(expect.objectContaining({ reason: 'memory-budget' }))
    refused.tokenizer.write(endField)
    expect(raw).toBe('000')
    oracle.events.dispose(); exact.events.dispose(); refused.events.dispose()
  })
  it('attributes and their map holder are admitted before copying; exact succeeds and one-under never retains them', () => {
    const prefix = '<wsDr><absoluteAnchor>'
    const suffix = '</absoluteAnchor></wsDr>'
    const value = '0'.repeat(1000) + '1'
    const tag = `<pos x="${value}" y="2"/>`
    const oracle = consumer()
    oracle.tokenizer.write(prefix)
    oracle.tokenizer.write(tag)
    const peak = Math.max(...oracle.requests)
    let parts: AnchorParts | undefined
    const exact = consumer(peak, (_, p) => { parts = p })
    exact.tokenizer.write(prefix + tag + suffix)
    expect(parts!.pos!.x).toBe(value)
    const refused = consumer(peak - 1, (_, p) => { parts = p })
    refused.tokenizer.write(prefix)
    expect(() => refused.events.open('pos', ['x', 'y'], [value, '2'], 2, true)).toThrow(expect.objectContaining({ reason: 'memory-budget' }))
    refused.tokenizer.write(suffix)
    expect(parts!.pos).toBeUndefined()
    oracle.events.dispose(); exact.events.dispose(); refused.events.dispose()
  })
  it('all closed from/to fields and ext/pos attributes remain jointly owned through real Number conversion', () => {
    const value = '0'.repeat(10_000) + '1'
    const marker = `<col>${value}</col><row>${value}</row><colOff>${value}</colOff><rowOff>${value}</rowOff>`
    const xml = `<wsDr><twoCellAnchor><from>${marker}</from><to>${marker}</to><ext cx="${value}" cy="${value}"/><pos x="${value}" y="${value}"/></twoCellAnchor></wsDr>`
    const store = new SourcePageStore({ docId: 'concurrent', backend: memoryPageBackend() })
    const share = store.construction.share('marker', { untilReleased: true })
    const nativeNumber = globalThis.Number
    const seen: number[] = []
    globalThis.Number = new Proxy(nativeNumber, { apply(target, receiver, args) {
      if (args[0] === value) seen.push(share.bytes)
      return Reflect.apply(target, receiver, args)
    } })
    const events = anchorRecordEvents((name, parts) => {
      expect(anchorRecord(name, 'drawing#0', parts)).toMatchObject({ kind: 'twoCell', from: { col: 1 }, to: { row: 1 } })
    }, () => {}, { need: (bytes) => share.need(bytes) })
    try {
      const parser = new XmlTokenizer(events)
      for (let at = 0; at < xml.length; at += 997) parser.write(xml.slice(at, at + 997))
      parser.end()
      expect(seen).toHaveLength(16)
      for (const bytes of seen) expect(bytes).toBeGreaterThan(12 * 2 * value.length)
    } finally { globalThis.Number = nativeNumber; events.dispose(); share.release() }
    expect(store.construction.transientBytes).toBe(0)
  })
  it('multiple/empty/self-closing/foreign/AlternateContent anchors preserve lexical Number semantics and source identity', () => {
    const records: ReturnType<typeof anchorRecord>[] = []
    let unreadable = 0, live = 0
    const events = anchorRecordEvents((name, parts) => { records.push(anchorRecord(name, `drawing#${records.length}`, parts)); expect(live).toBeGreaterThan(0) }, () => { unreadable++ }, { need: (bytes) => { live = bytes } })
    const parser = new XmlTokenizer(events)
    parser.write('<wsDr><oneCellAnchor/><oneCellAnchor><from><col> 01e0 </col><row> 0x2 </row><colOff>-1.5e2</colOff><rowOff> +.5 </rowOff></from><ext cx=" 6.096e5 " cy="0"/></oneCellAnchor><oneCellAnchor><from><col></col><row/><colOff>0</colOff><rowOff>0</rowOff></from></oneCellAnchor><AlternateContent><Choice/></AlternateContent></wsDr>')
    parser.end()
    expect(records).toEqual([undefined, { id: 'drawing#1', kind: 'oneCell', from: { col: 1, row: 2, colOffEmu: -150, rowOffEmu: 0.5 }, cx: 609600, cy: 0 }, undefined])
    expect(unreadable).toBe(1)
    const baseline = live
    events.dispose(); events.dispose()
    expect(() => events.text('late', true)).toThrow(expect.objectContaining({ code: 'cancelled' }))
    expect(live).toBe(baseline)
    const foreign = anchorRecordEvents(() => { throw new Error('foreign anchor') }, () => { unreadable++ })
    new XmlTokenizer(foreign).write('<foreign><oneCellAnchor/></foreign>')
    foreign.dispose()
    expect(unreadable).toBe(2)
  })
  it.each(['callback', 'token', 'envelope'] as const)('%s failure releases explicit marker ownership after dropping fields', async (failure) => {
    const store = new SourcePageStore({ docId: failure, backend: memoryPageBackend() })
    const share = store.construction.share('marker', { untilReleased: true })
    const events = anchorRecordEvents(() => { if (failure === 'callback') throw new Error('callback failure') }, () => {}, { need: (bytes) => share.need(bytes) })
    async function* input() {
      yield new TextEncoder().encode(startField + '0001')
      if (failure === 'envelope') throw new PackageError('resource-limit', 'part envelope')
      yield new TextEncoder().encode(failure === 'token' ? '&bad' : endField)
    }
    try { await expect(tokenizeBytes(input(), events)).rejects.toThrow() }
    finally { events.dispose(); share.release(); share.release() }
    expect(store.construction.transientBytes).toBe(0)
  })
  it('close while input is withheld keeps fields charged until parser disposal and prohibits revival', async () => {
    const store = new SourcePageStore({ docId: 'withheld', backend: memoryPageBackend() })
    const share = store.construction.share('marker', { untilReleased: true })
    const events = anchorRecordEvents(() => {}, () => {}, { need: (bytes) => share.need(bytes) })
    let release!: () => void, reached!: () => void
    const gate = new Promise<void>((resolve) => { release = resolve })
    const held = new Promise<void>((resolve) => { reached = resolve })
    async function* input() {
      yield new TextEncoder().encode(startField + '0'.repeat(100_000))
      reached(); await gate
      yield new TextEncoder().encode('1' + endField)
    }
    const parsing = tokenizeBytes(input(), events).finally(() => { events.dispose(); share.release() })
    await held
    const before = store.construction.transientBytes
    expect(before).toBeGreaterThan(200_000)
    await store.close()
    expect(store.construction.transientBytes).toBe(before)
    release()
    await expect(parsing).rejects.toMatchObject({ code: 'cancelled' })
    expect(store.construction.transientBytes).toBe(0)
    expect(() => events.text('late', true)).toThrow(expect.objectContaining({ code: 'cancelled' }))
    share.release()
    expect(store.construction.transientBytes).toBe(0)
  })
})


describe('real drawing import teardown and metadata refusal', () => {
  it('a roomy bound pool admits the giant with its ownership live through Number', async () => {
    const r = await run(giant, 'bound', 'D', 16 * MiB)
    expect(r.conversions).toHaveLength(2)
    for (const c of r.conversions) expect(c.transient).toBeGreaterThan(2 * c.chars)
    expect(r.anchors.records[0]).toMatchObject({ from: { col: 1 } })
  })
  it.each(['cancel', 'close'] as const)('%s during a real stored-ZIP drawing read keeps ownership until held input returns', async (mode) => {
    const zip = drawingFixture(startField + giant + endField, undefined, 0)
    const layout = zip.entries.at(-1)!
    let release!: () => void, reached!: () => void
    const gate = new Promise<void>((resolve) => { release = resolve })
    const held = new Promise<void>((resolve) => { reached = resolve })
    class GatedBlob extends Blob {
      override slice(start?: number, end?: number, contentType?: string): Blob {
        const slice = super.slice(start, end, contentType)
        if (start === layout.dataStart + 262144) {
          const read = slice.arrayBuffer.bind(slice)
          slice.arrayBuffer = async () => { reached(); await gate; return read() }
        }
        return slice
      }
    }
    const store = new SourcePageStore({ docId: mode, backend: memoryPageBackend(), memoryBudgetBytes: 16 * MiB })
    const session = await openImportSession({ kind: 'blob', blob: new GatedBlob([zip.bytes as BlobPart]) }, { docId: mode, profile: 'D', store })
    try {
      await held
      const before = store.construction.transientBytes
      expect(before).toBeGreaterThan(400_000)
      const cancelled = mode === 'cancel' ? session.cancel() : store.close()
      if (mode === 'close') await cancelled
      expect(store.construction.transientBytes).toBe(before)
      release()
      await cancelled
      expect(await session.done).toEqual({ kind: 'cancelled' })
      expect(store.construction.transientBytes).toBe(0)
    } finally { release(); await session.release(); await session.release() }
    expect(store.construction.persistentBytes).toBe(0)
  })
  it.each(['token', 'envelope', 'foreign'] as const)('real %s failure leaves FULL metadata incomplete and A1 readable', async (failure) => {
    const xml = failure === 'foreign' ? '<foreign><oneCellAnchor/></foreign>' : startField + '0'.repeat(100_000) + (failure === 'token' ? '&bad' : endField)
    const zip = drawingFixture(xml)
    const store = new SourcePageStore({ docId: failure, backend: memoryPageBackend(), memoryBudgetBytes: 16 * MiB })
    const session = await openImportSession({ kind: 'blob', blob: new Blob([zip.bytes as BlobPart]) }, { docId: failure, profile: 'D', store,
      ...(failure === 'envelope' ? { limits: { paged: { drawingPartEnvelopeBytes: 80_000 } } } : {}),
    })
    try {
      expect((await session.done).kind).toBe('complete')
      const sheet = createPagedWorkbookSource(session).store('1'), opts = { signal: new AbortController().signal, stamp: sheet.stamp }
      expect(await sheet.readCell({ row: 0, col: 0 }, opts)).toMatchObject({ kind: 'cell', cell: { value: 1 } })
      expect(await sheet.queryStructural({ r0: 0, r1: 1048575, c0: 0, c1: 16383 }, opts)).toMatchObject({ completeness: 'incomplete' })
      expect(session.workbook.sheets[0].drawingAnchors.unresolved).toBeGreaterThan(0)
      expect(store.construction.transientBytes).toBe(0)
    } finally { await session.release() }
  })
})
