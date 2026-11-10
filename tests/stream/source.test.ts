import { afterEach, expect, it, vi } from 'vitest'
import { openByteStream, readSource, type HttpSource } from '../../src/core/stream'
import { isStreamSource, StreamSourceMisuseError } from '../../src/stream/source'

afterEach(() => { vi.restoreAllMocks() })

const stream = { kind: 'stream', url: 'wss://example.test/s' } as const

it('readSource and openByteStream refuse a StreamSource without fetching', async () => {
  const fetchSpy = vi.spyOn(globalThis, 'fetch')
  await expect(readSource(stream as never)).rejects.toBeInstanceOf(StreamSourceMisuseError)
  await expect(openByteStream(stream as never)).rejects.toBeInstanceOf(StreamSourceMisuseError)
  expect(fetchSpy).not.toHaveBeenCalled()
})

it('a Promise resolving to a StreamSource is refused too', async () => {
  const fetchSpy = vi.spyOn(globalThis, 'fetch')
  await expect(readSource(Promise.resolve(stream) as never)).rejects.toBeInstanceOf(StreamSourceMisuseError)
  await expect(openByteStream(Promise.resolve(stream) as never)).rejects.toBeInstanceOf(StreamSourceMisuseError)
  expect(fetchSpy).not.toHaveBeenCalled()
})

it('an HttpSource without kind still downloads', async () => {
  const fetchSpy = vi.spyOn(globalThis, 'fetch').mockResolvedValue(new Response(new Uint8Array([1, 2, 3])))
  const bytes = await readSource({ url: 'https://example.test/a.docx' })
  expect([...bytes]).toEqual([1, 2, 3])
  expect(fetchSpy).toHaveBeenCalledTimes(1)
})

it('isStreamSource', () => {
  expect(isStreamSource({ kind: 'stream', url: 'wss://a' })).toBe(true)
  expect(isStreamSource({ kind: 'stream', url: 'wss://a', token: 't' })).toBe(true)
  expect(isStreamSource({ url: 'wss://a' })).toBe(false)
  expect(isStreamSource({ kind: 'stream' })).toBe(false)
  expect(isStreamSource({ kind: 'other', url: 'wss://a' })).toBe(false)
  expect(isStreamSource(null)).toBe(false)
  expect(isStreamSource('wss://a')).toBe(false)
})

it('the error names itself and explains the fix', () => {
  const e = new StreamSourceMisuseError()
  expect(e.name).toBe('StreamSourceMisuseError')
  expect(e.message).toBe('A StreamSource cannot be read as bytes; pass it to <OfficeFile data> to open a stream session.')
})

it('HttpSource cannot carry kind', () => {
  // @ts-expect-error — HttpSource.kind is `never` (§6.1)
  const h: HttpSource = { kind: 'stream', url: 'wss://a' }
  expect(h.url).toBe('wss://a')
})
