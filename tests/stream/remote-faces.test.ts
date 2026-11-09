/** 1d Task 3: scale buckets and pinned-face hashing (§7.4 rules 3–4, 6). */
import { describe, expect, it } from 'vitest'
import { bucketFor } from '../../src/stream/remote/remote-doc'
import { resolveFaces } from '../../src/stream/remote/faces'

async function hash96(bytes: Uint8Array) {
  const d = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes as BufferSource))
  return [...d.subarray(0, 12)].map((b) => b.toString(16).padStart(2, '0')).join('')
}

describe('bucketFor', () => {
  it('picks the smallest bucket ≥ zoom × dpr, else the largest', () => {
    expect(bucketFor(1, 1, [1, 2])).toBe(1)
    expect(bucketFor(1, 2, [1, 2])).toBe(2)
    expect(bucketFor(1.25, 1, [1, 2])).toBe(2)
    expect(bucketFor(3, 2, [1, 2])).toBe(2)
    expect(bucketFor(0.5, 1, [1, 2])).toBe(1)
  })
})

describe('resolveFaces', () => {
  const good = new Uint8Array([1, 2, 3, 4]), bad = new Uint8Array([1, 2, 3, 5])
  const pack = { version: 1, faces: [
    { id: 'carlito-regular', file: 'Carlito-Regular.ttf', family: 'Carlito', weight: 400, style: 'normal' },
    { id: 'carlito-bold', file: 'Carlito-Bold.ttf', family: 'Carlito', weight: 700, style: 'normal' },
  ] }
  function server(files: Record<string, Uint8Array>) {
    return (async (url: string) => {
      const name = String(url).split('/').pop()!
      if (name === 'pack.json') return new Response(JSON.stringify(pack))
      const f = files[name]
      return f ? new Response(f as BodyInit) : new Response('', { status: 404 })
    }) as typeof fetch
  }
  it('without a baseUrl every required face is missing', async () => {
    expect(await resolveFaces([{ id: 'carlito-regular', hash: 'x'.repeat(24) }], undefined)).toEqual(['carlito-regular'])
  })
  it('registers a face whose bytes hash to the manifest hash; a tampered face is missing and never registered', async () => {
    const registered: Array<[string, number, string]> = []
    const required = [{ id: 'carlito-regular', hash: await hash96(good) }, { id: 'carlito-bold', hash: await hash96(good) }]
    const missing = await resolveFaces(required, 'https://cdn/fontpack', {
      fetch: server({ 'Carlito-Regular.ttf': good, 'Carlito-Bold.ttf': bad }),
      register: async (family, _bytes, d) => { registered.push([family, d.weight, d.style]) },
    })
    expect(missing).toEqual(['carlito-bold'])
    expect(registered).toEqual([['Carlito', 400, 'normal']])
  })
  it('a required id that pack.json does not list is missing', async () => {
    const missing = await resolveFaces([{ id: 'liberation-sans-regular', hash: await hash96(good) }], 'https://cdn/fontpack', { fetch: server({}), register: async () => {} })
    expect(missing).toEqual(['liberation-sans-regular'])
  })
  it('a fetch failure makes the face missing, not an error', async () => {
    const missing = await resolveFaces([{ id: 'carlito-regular', hash: await hash96(good) }], 'https://cdn/fontpack', { fetch: (async () => { throw new Error('offline') }) as typeof fetch, register: async () => {} })
    expect(missing).toEqual(['carlito-regular'])
  })
})
