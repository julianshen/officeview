import { describe, test, expect } from 'vitest'
import { readSource, readByteStream, type Progress } from '../src/core/stream'
import { loadOfficeFile } from '../src/components/OfficeFile'
import { buildDocx, buildXlsx, buildPptx } from '../src/testdata/ooxml-builders'

/** A ReadableStream that emits the source in fixed-size chunks. */
function chunkedStream(bytes: Uint8Array, chunkSize: number): ReadableStream<Uint8Array> {
  let at = 0
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      if (at >= bytes.length) {
        controller.close()
        return
      }
      controller.enqueue(bytes.subarray(at, at + chunkSize))
      at += chunkSize
    },
  })
}

const probe = (fn: unknown) => fn as (p: Progress) => void

describe('readByteStream', () => {
  test('reassembles chunked bytes exactly', async () => {
    const source = new Uint8Array(1000)
    for (let i = 0; i < source.length; i++) source[i] = i % 251
    const out = await readByteStream(chunkedStream(source, 137))
    expect(out.length).toBe(1000)
    expect(Array.from(out)).toEqual(Array.from(source))
  })

  test('reports progress: per chunk without a total, whole percent with one', async () => {
    const source = new Uint8Array(1000)
    const seen: Progress[] = []
    const spy = probe((p: Progress) => seen.push({ ...p }))
    await readByteStream(chunkedStream(source, 100), spy)
    // no total: initial (loaded 0) + one per chunk + final = 12
    expect(seen.length).toBe(12)
    expect(seen[0].loaded).toBe(0)
    expect(seen[seen.length - 1].loaded).toBe(1000)
    expect(seen.every((p) => p.total === undefined)).toBe(true)

    // with a total: whole-percent steps, monotonic, capped at 100%
    const withTotal: Progress[] = []
    await readByteStream(chunkedStream(source, 300), (p) => withTotal.push({ ...p }), 1000)
    const last = withTotal[withTotal.length - 1]
    expect(last.total).toBe(1000)
    expect(last.loaded).toBe(1000)
  })

  test('handles an empty stream', async () => {
    const out = await readByteStream(chunkedStream(new Uint8Array(0), 10))
    expect(out.length).toBe(0)
  })
})

describe('readSource', () => {
  test('accepts a Node Buffer cross-realm (jsdom makes instanceof lie)', async () => {
    // A Buffer produced by Node's fs is created in Node's realm, so under
    // vitest's jsdom environment `buf instanceof Uint8Array` is false — which
    // is why readSource must use ArrayBuffer.isView instead of instanceof
    // chains. The bytes are written to a temp file rather than read from
    // corpus/, so this suite needs no fetched fixture.
    const { writeFileSync, readFileSync, unlinkSync } = await import('fs')
    const { tmpdir } = await import('os')
    const { join } = await import('path')
    const path = join(tmpdir(), `officeview-bytes-${process.pid}.bin`)
    const payload = [0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 250, 251, 252]
    writeFileSync(path, Buffer.from(payload))
    let buf: Buffer
    try {
      buf = readFileSync(path)
    } finally {
      unlinkSync(path)
    }
    // guard the premise: if this environment ever made fs Buffers plain
    // Uint8Arrays, the test would pass trivially and prove nothing
    expect(buf instanceof Uint8Array).toBe(false)
    const out = await readSource(buf as never)
    expect(Array.from(out)).toEqual(payload)
    expect(out.length).toBe(payload.length)
  })

  test('accepts a Response with content-length progress', async () => {
    const bytes = new Uint8Array(500)
    const body = chunkedStream(bytes, 200)
    const response = new Response(body, { headers: { 'content-length': '500' } })
    const seen: Progress[] = []
    const out = await readSource(response, probe((pp: Progress) => seen.push({ ...pp })))
    expect(out.length).toBe(500)
    expect(seen[0].total).toBe(500)
    expect(seen[seen.length - 1].loaded).toBe(500)
  })

  test('accepts a Blob', async () => {
    const bytes = new Uint8Array(300)
    const out = await readSource(new Blob([bytes]))
    expect(out.length).toBe(300)
  })

  test('rejects nonsense loudly', async () => {
    await expect(readSource('not bytes' as never)).rejects.toThrow('Unsupported byte source')
  })
})

describe('loadOfficeFile with streaming sources', () => {
  test('parses a docx delivered through a chunked stream', async () => {
    const bytes = await buildDocx([{ runs: [{ text: 'streamed doc' }] }])
    const doc = await loadOfficeFile(chunkedStream(bytes, 64))
    expect('sections' in doc).toBe(true)
    expect((doc as { sections: Array<{ paragraphs: Array<{ runs: Array<{ text: string }> }> }> }).sections[0].paragraphs[0].runs[0].text).toBe('streamed doc')
  })

  test('parses an xlsx and a pptx the same way', async () => {
    const xlsx = await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', v: 42 }] }] }])
    expect('sheets' in (await loadOfficeFile(chunkedStream(xlsx, 50)))).toBe(true)
    const pptx = await buildPptx([{ prst: 'rect' }])
    expect('slides' in (await loadOfficeFile(chunkedStream(pptx, 128)))).toBe(true)
  })

  test('a truncated stream surfaces the zip error rather than hanging', async () => {
    const bytes = await buildDocx([{ runs: [{ text: 'will be cut' }] }])
    const cut = bytes.subarray(0, Math.floor(bytes.length * 0.4))
    await expect(loadOfficeFile(chunkedStream(cut, 100))).rejects.toThrow()
  })
})
