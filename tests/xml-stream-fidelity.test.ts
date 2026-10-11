/** S1.1 end-to-end string fidelity: ZIP -> pull inflate -> UTF-8 -> tokenizer
 * -> ST_Xstring, against f4StringFor and the current parser (RED test 11). */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { createDeflateRaw } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import { bytesRangeSource, openPart, readZipInventory, type ZipLimits } from '../src/core/zip-stream'
import { crc32Update } from '../src/core/crc32'
import { XstringDecoder, localName, tokenizeBytes, type XmlEvents } from '../src/core/xml-stream'
import { getChildren, parseXml, textOf, type XmlNode } from '../src/core/xml'
import { decodeXlsxString } from '../src/xlsx/strings'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { encodeXlsxString, f4StringFor, F4_LONG_BLOB_INDEX, FIXTURE_SEEDS } from './helpers/xlsx-viewer-fixtures'
import { buildZip, concatBytes } from './helpers/zip-builder'
import { sharedStringsConsumer } from './helpers/xml-recorder'

const LARGE = process.env.IV_LARGE === '1'
const MiB = 1024 * 1024
const LIMITS: ZipLimits = {
  maxEntries: 64, maxCentralDirectoryBytes: MiB, maxNameBytes: 256,
  maxPartExpandedBytes: 512 * MiB, maxTotalExpandedBytes: 1024 * MiB, maxRatio: 200,
}
const enc = new TextEncoder()

/** Same layout the fixture generator streams for xl/sharedStrings.xml. */
function sstChunks(indices: number[]): Uint8Array[] {
  const parts: Uint8Array[] = [enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${indices.length}" uniqueCount="${indices.length}">`)]
  let batch = ''
  for (const i of indices) {
    batch += `<si><t>${encodeXlsxString(f4StringFor(FIXTURE_SEEDS.F4, i))}</t></si>`
    if (batch.length > 1 << 20) { parts.push(enc.encode(batch)); batch = '' }
  }
  parts.push(enc.encode(batch + '</sst>'))
  return parts
}

async function deflateChunks(chunks: Uint8Array[]): Promise<{ raw: Uint8Array; crc: number; size: number }> {
  const d = createDeflateRaw({ level: 6 })
  const out: Uint8Array[] = []
  d.on('data', (c: Uint8Array) => out.push(c))
  let crc = 0
  let size = 0
  for (const c of chunks) {
    crc = crc32Update(crc, c)
    size += c.length
    if (!d.write(c)) await new Promise((r) => d.once('drain', r))
  }
  const finished = new Promise((r) => d.once('end', r))
  d.end()
  await finished
  return { raw: concatBytes(...out), crc, size }
}

async function streamSharedStrings(bytes: Uint8Array, outputBufferBytes: number, maxTextSegmentChars?: number): Promise<string[]> {
  const src = bytesRangeSource(bytes)
  const inv = await readZipInventory(src, LIMITS)
  const entry = inv.find('xl/sharedStrings.xml')!
  const consumer = sharedStringsConsumer()
  const signal = new AbortController().signal
  await tokenizeBytes(openPart(src, entry, { signal, budget: { charge() {} }, outputBufferBytes }), consumer.handler, { signal, limits: maxTextSegmentChars === undefined ? undefined : { maxTextSegmentChars } })
  return consumer.strings
}

/** Current path oracle: parseXml + parse.ts sharedStrings() logic. */
function legacySharedStrings(xml: string): string[] {
  const root = parseXml(xml)
  return getChildren(root, 'si').map((si) => {
    let s = ''
    const direct = si['t']
    if (direct !== undefined) s += decodeXlsxString(textOf(typeof direct === 'object' ? direct : { '#text': String(direct ?? '') } as XmlNode))
    for (const r of getChildren(si, 'r')) s += decodeXlsxString(textOf(getChildren(r, 't')[0]))
    return s
  })
}

describe('F4-style sharedStrings fidelity', () => {
  it(`matches f4StringFor and the current parser (${LARGE ? '1,000,000 strings' : '3,000 strings plus the 100,000-character blob'})`, async () => {
    const count = LARGE ? 1_000_000 : 3000
    const indices = Array.from({ length: count }, (_, i) => i)
    if (!LARGE) indices.push(F4_LONG_BLOB_INDEX)
    const chunks = sstChunks(indices)
    const { raw, crc, size } = await deflateChunks(chunks)
    const layout = buildZip([
      { name: '[Content_Types].xml', data: '<Types/>', method: 0 },
      { name: 'xl/sharedStrings.xml', raw, crc, uncompressedSize: size },
    ])
    const expected = indices.map((i) => f4StringFor(FIXTURE_SEEDS.F4, i))
    // families covered: lone surrogates, astral, accents, long and the blob
    expect(expected.some((s) => /[\uD800-\uDBFF](?![\uDC00-\uDFFF])/.test(s))).toBe(true)
    expect(expected.some((s) => /(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(s))).toBe(true)
    expect(expected.some((s) => s.includes('😀'))).toBe(true)
    const strategies: Array<[number, number | undefined]> = LARGE
      ? [[65_536, undefined], [16_384, 1000]]
      : [[65_536, undefined], [16_384, 1000], [7, 7]]
    for (const [out, seg] of strategies) {
      const got = await streamSharedStrings(layout.bytes, out, seg)
      expect(got.length).toBe(expected.length)
      let mismatches = 0
      for (let i = 0; i < expected.length; i++) if (got[i] !== expected[i]) mismatches++
      expect(mismatches).toBe(0)
    }
    if (!LARGE) {
      const xml = new TextDecoder().decode(concatBytes(...chunks))
      expect(legacySharedStrings(xml)).toEqual(expected)
    }
  }, LARGE ? 1_800_000 : 60_000)

  it('ST_Xstring corner cases survive the whole pipeline: CR, escaped underscore, C0 controls, lone surrogates', async () => {
    const strings = ['a\rb', 'x_x0041_y', '\u0001\u001f', '\uD800', '\uDE00z', 'tab\tnl\n', 'é😀']
    const xml = `<sst>${strings.map((s) => `<si><t>${encodeXlsxString(s)}</t></si>`).join('')}</sst>`
    const layout = buildZip([{ name: 'xl/sharedStrings.xml', data: xml }])
    for (const out of [65_536, 1]) expect(await streamSharedStrings(layout.bytes, out)).toEqual(strings)
    expect(legacySharedStrings(xml)).toEqual(strings)
  })
})

describe('generator fixture parity with the current parser', () => {
  it('f4-small.xlsx inline strings stream to the same cell text parseXlsx produces', async () => {
    const bytes = new Uint8Array(readFileSync(join(process.cwd(), 'tests/fixtures/xlsx-viewer/f4-small.xlsx')))
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const legacy = new Map<string, unknown>()
    for (const row of doc.sheets[0].rows) for (const cell of row.cells) legacy.set(cell.ref, cell.value)

    const streamed = new Map<string, string>()
    let ref = ''
    let decoder: XstringDecoder | null = null
    let value = ''
    const stack: string[] = []
    const events: XmlEvents = {
      open(qname, names, values, count, selfClosing) {
        const name = localName(qname)
        if (name === 'c') {
          ref = ''
          for (let i = 0; i < count; i++) if (names[i] === 'r') ref = values[i]
          value = ''
        }
        if (selfClosing) return
        stack.push(name)
        if (name === 't' && stack.includes('is')) decoder = new XstringDecoder()
      },
      close() {
        const name = stack.pop()
        if (name === 't' && decoder !== null) { value += decoder.push('', true); decoder = null }
        if (name === 'c') streamed.set(ref, value)
      },
      text(t) { if (decoder !== null) value += decoder.push(t) },
    }
    const src = bytesRangeSource(bytes)
    const inv = await readZipInventory(src, LIMITS)
    const signal = new AbortController().signal
    await tokenizeBytes(openPart(src, inv.find('xl/worksheets/sheet1.xml')!, { signal, budget: { charge() {} }, outputBufferBytes: 4096 }), events, { signal })
    expect(streamed.size).toBe(legacy.size)
    expect(streamed.size).toBeGreaterThan(400)
    for (const [r, v] of streamed) expect(v, r).toBe(legacy.get(r))
  })
})
