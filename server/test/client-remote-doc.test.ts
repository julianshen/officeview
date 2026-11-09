/** 1d Task 3: RemoteOfficeDoc paint paths against the real 1c server (napi backend). */
import { afterAll, describe, expect, it } from 'bun:test'
import { loadImage } from '@napi-rs/canvas'
import { replayUnit } from '../../src/stream/replay'
import type { Manifest } from '../../src/stream/manifest'
import { StreamSessionError } from '../../src/stream/remote/errors'
import { createRemoteDoc, type SessionLike } from '../../src/stream/remote/remote-doc'
import { openRemoteDocument } from '../../src/stream/remote/open'
import type { StreamNotice, UnitPayload } from '../../src/stream/remote/session'
import { serverBackend } from '../src/backend'
import { ORIGIN, disposePool, start } from './harness'

afterAll(() => disposePool())
const backend = serverBackend()
const client = (s: { url: string }) => new WebSocket(s.url, { headers: { Origin: ORIGIN } } as never)
const pixels = (c: { width: number; height: number; getContext(t: '2d'): CanvasRenderingContext2D | null }) => Buffer.from(c.getContext('2d')!.getImageData(0, 0, c.width, c.height).data)

describe('RemoteOfficeDoc', () => {
  it('protected unit: the canvas holds exactly the server PNG at the bucket', async () => {
    const h = await start()
    const doc = await openRemoteDocument({ kind: 'stream', url: h.url }, { client, backend })
    const [u] = await doc.units()
    const canvas = backend.createCanvas(1, 1)
    await doc.paint(0, 1, 2, canvas)
    expect(canvas.width).toBe(u!.widthPx * 2)
    // Same bytes the session receives, decoded independently.
    const direct = await (doc as unknown as { session: SessionLike }).session.unit(0, 2)
    if (direct.kind !== 'raster') throw new Error('raster')
    const ref = backend.createCanvas(direct.w, direct.h)
    ref.getContext('2d')!.drawImage(await loadImage(Buffer.from(direct.png)) as never, 0, 0)
    expect(pixels(canvas).equals(pixels(ref))).toBe(true)
    doc.close(); h.stop()
  }, 30_000)
  it('relaxed unit with every face present: pixels equal a direct replay of the same payload', async () => {
    const h = await start({ policy: { projection: 'relaxed' } })
    const doc = await openRemoteDocument({ kind: 'stream', url: h.url }, { client, backend, faces: async () => [] })
    const canvas = backend.createCanvas(1, 1)
    await doc.paint(0, 1, 1, canvas)
    const payload = await (doc as unknown as { session: SessionLike }).session.unit(0, 1)
    if (payload.kind !== 'commands') throw new Error('commands')
    const ref = backend.createCanvas(canvas.width, canvas.height)
    await replayUnit(ref.getContext('2d')!, payload, backend)
    expect(pixels(canvas).equals(pixels(ref))).toBe(true)
    doc.close(); h.stop()
  }, 30_000)
})

function fakeSession(payloads: UnitPayload[]): SessionLike & { requests: number; closes: number } {
  const manifest = { units: { count: 1, list: [{ w: 4, h: 4 }] }, recorder: { scalesBuckets: [1, 2] }, can: { search: false, copy: false } } as unknown as Manifest
  const s = {
    manifest, requests: 0, closes: 0,
    async unit() { const p = payloads[Math.min(s.requests, payloads.length - 1)]!; s.requests++; return p },
    close() { s.closes++ },
  }
  return s
}

describe('RemoteOfficeDoc failure paths', () => {
  it('a corrupted command unit is re-requested once, then rejects with a unit-failed notice', async () => {
    const notices: StreamNotice[] = []
    const s = fakeSession([{ kind: 'commands', ops: new Uint8Array([0xff, 0xff, 0xff]), assets: new Map() }])
    const doc = createRemoteDoc(s, { backend, onNotice: (n) => notices.push(n) })
    const canvas = backend.createCanvas(4, 4)
    canvas.getContext('2d')!.fillStyle = '#123456'; canvas.getContext('2d')!.fillRect(0, 0, 4, 4)
    const before = pixels(canvas)
    const e = await doc.paint(0, 1, 1, canvas).then(() => undefined, (x) => x)
    expect((e as StreamSessionError).code).toBe('unit-failed')
    expect(s.requests).toBe(2)
    expect(notices).toEqual([{ code: 'unit-failed', unit: 0 }])
    expect(pixels(canvas).equals(before)).toBe(true)    // never half-painted
  })
  it('search and copy reject; text-unavailable is emitted once', async () => {
    const notices: StreamNotice[] = []
    const doc = createRemoteDoc(fakeSession([]), { backend, onNotice: (n) => notices.push(n) })
    await expect(doc.search('x')).rejects.toBeInstanceOf(StreamSessionError)
    await expect(doc.copy({})).rejects.toBeInstanceOf(StreamSessionError)
    expect(notices.filter((n) => n.code === 'text-unavailable').length).toBe(1)
  })
  it('close() closes the session once; later paint rejects', async () => {
    const s = fakeSession([])
    const doc = createRemoteDoc(s, { backend })
    doc.close(); doc.close()
    expect(s.closes).toBe(1)
    await expect(doc.paint(0, 1, 1, backend.createCanvas(4, 4))).rejects.toBeInstanceOf(StreamSessionError)
  })
})
