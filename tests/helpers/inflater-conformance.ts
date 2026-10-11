/** Strict termination probe over the §2.4.2 table (design note RED test 5a).
 * An inflater (or engine adapter) may back the ZIP reader only if every row
 * gives the Chrome-155 answer: valid streams accepted, a final block followed
 * by junk, a non-final block alone and a truncated stream all rejected. */
import { deflateRawSync } from 'node:zlib'
import { openPart, readZipInventory, bytesRangeSource, type InflaterFactory } from '../../src/core/zip-stream'
import { buildZip, concatBytes, storedBlock } from './zip-builder'

const enc = new TextEncoder()

export interface ProbeRow { name: string; expectAccept: boolean; accepted: boolean }
export interface ProbeResult { pass: boolean; rows: ProbeRow[] }

export async function probeTermination(factory: InflaterFactory): Promise<ProbeResult> {
  const hello = new Uint8Array(deflateRawSync(enc.encode('hello')))
  const cases: Array<{ name: string; raw: Uint8Array; data: string; expectAccept: boolean }> = [
    { name: 'final stored block + 3 junk bytes', raw: concatBytes(storedBlock(enc.encode('ABC'), true), enc.encode('xyz')), data: 'ABC', expectAccept: false },
    { name: 'deflate("hello") + GARBAGE', raw: concatBytes(hello, enc.encode('GARBAGE')), data: 'hello', expectAccept: false },
    { name: 'non-final stored block only', raw: storedBlock(enc.encode('ABC'), false), data: 'ABC', expectAccept: false },
    { name: 'truncated deflate', raw: hello.subarray(0, 3), data: 'hello', expectAccept: false },
    { name: 'valid', raw: hello, data: 'hello', expectAccept: true },
  ]
  const rows: ProbeRow[] = []
  for (const c of cases) {
    const layout = buildZip([{ name: 'p.xml', data: c.data, raw: c.raw }])
    const src = bytesRangeSource(layout.bytes)
    let accepted = true
    try {
      const inv = await readZipInventory(src, { maxEntries: 8, maxCentralDirectoryBytes: 1 << 20, maxNameBytes: 256, maxPartExpandedBytes: 1 << 20, maxTotalExpandedBytes: 1 << 20, maxRatio: 1000 })
      const entry = inv.find('p.xml')!
      for await (const chunk of openPart(src, entry, { signal: new AbortController().signal, budget: { charge() {} } }, factory)) void chunk
    } catch {
      accepted = false
    }
    rows.push({ name: c.name, expectAccept: c.expectAccept, accepted })
  }
  return { pass: rows.every((r) => r.accepted === r.expectAccept), rows }
}
