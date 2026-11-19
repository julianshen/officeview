/** Piece 5 Task 8: the shared encode cache (plan D12, D23) with a fake `EncodeRun` and an injected clock. */
import { afterAll, describe, expect, it } from 'bun:test'
import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { MediaRefusal } from '../../../src/stream/media-types'
import { createEncodeCache, MEDIA_IDLE_TTL_MS, MEDIA_INFLIGHT_RESERVE, MEDIA_SERVER_CACHE, MEDIA_SESSION_LOOKAHEAD, type EncodeCache, type EncodeRun } from '../../src/media/cache'
import type { EncodeKey } from '../../src/media/params'
import { createMediaTemp } from '../../src/media/temp'
import type { EncodeResult } from '../../src/media/transcode'

const root = mkdtempSync(join(tmpdir(), 'ov-cache-'))
const temp = await createMediaTemp(root)
afterAll(async () => { await temp.remove(); rmSync(root, { recursive: true, force: true }) })

const MiB = 1 << 20
const never = () => new AbortController().signal
const key = (docId: string, sourceHash: string): EncodeKey => ({ docId, sourceHash, paramsSha: 'p'.repeat(64), ffmpegVersion: '8.1.1' })
const okResult = (bytes: number): EncodeResult & { ok: true } => ({ ok: true, bytes, durationMs: 2000, contentType: 'audio/mp4', paramsSha: 'p'.repeat(64), ffmpegVersion: '8.1.1' })

/** A run whose output is a tiny real file reported as `bytes`; `finish()` settles it, `fail(reason)` refuses. */
function fakeRun(bytes: number) {
  let settle!: (r: 'ok' | MediaRefusal) => void
  const gate = new Promise<'ok' | MediaRefusal>((ok) => { settle = ok })
  const state = { calls: 0, aborted: false }
  const run: EncodeRun = async (signal) => {
    state.calls++
    const out = temp.reserve('m4a')
    writeFileSync(out.path, 'x')
    const r = await Promise.race([gate, new Promise<'abort'>((ok) => signal.addEventListener('abort', () => ok('abort'), { once: true }))])
    if (r === 'abort') { state.aborted = true; await out.dispose(); throw signal.reason }
    if (r !== 'ok') { await out.dispose(); return { ok: false, reason: r } }
    return { ok: true, path: out.path, result: okResult(bytes), file: out }
  }
  return { run, state, finish: () => settle('ok'), fail: (r: MediaRefusal) => settle(r) }
}
const ahead = { mode: 'ahead' as const, signal: never(), cursor: 0, order: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11] }
const request = { ...ahead, mode: 'request' as const }
function cacheWith(o: Partial<Parameters<typeof createEncodeCache>[0]> = {}) {
  let t = 1_000_000
  const c = createEncodeCache({ temp, sessionBytes: MEDIA_SESSION_LOOKAHEAD, serverBytes: MEDIA_SERVER_CACHE, now: () => t, ...o })
  return { c, advance: (ms: number) => { t += ms } }
}
/** A done entry for (sid, k): acquired, finished and awaited. */
async function fill(c: EncodeCache, sid: string, k: EncodeKey, bytes: number, unit = 0, mode = ahead) {
  const f = fakeRun(bytes)
  const a = c.acquire(sid, unit, k, f.run, mode)
  if ('wait' in a) throw new Error('unexpected wait')
  f.finish()
  const r = await a.done
  if (!r.ok) throw new Error(`fill: ${r.reason}`)
  return r.media
}
const done = async (a: ReturnType<EncodeCache['acquire']>) => { if ('wait' in a) throw new Error('unexpected wait'); return await a.done }

describe('the shared encode cache (D12, D23)', () => {
  it('a key referenced by another session is not evicted under server pressure', async () => {
    const { c } = cacheWith({ serverBytes: 64 * MiB })
    await fill(c, 's1', key('d', 'a'), 30 * MiB) // s1 keeps its reference
    await fill(c, 's2', key('d', 'b'), 30 * MiB)
    c.delivered('s2', key('d', 'b')) // unreferenced now
    // A request needing 32 MiB evicts only the unreferenced entry.
    const f = fakeRun(10 * MiB)
    const a = c.acquire('s3', 0, key('d', 'c'), f.run, request)
    f.finish()
    expect((await done(a)).ok).toBe(true)
    expect(c.peek(key('d', 'a'))).toBeDefined()
    expect(c.peek(key('d', 'b'))).toBeUndefined()
    await c.close()
  })

  it('unreferenced entries evict least-recently-used first, and survive their sessions\' end', async () => {
    const { c, advance } = cacheWith({ serverBytes: 100 * MiB })
    // Each is released as it is delivered: its idle count (and LRU order) starts there.
    await fill(c, 's1', key('d', 'old'), 30 * MiB); c.delivered('s1', key('d', 'old')); advance(10)
    await fill(c, 's1', key('d', 'mid'), 30 * MiB); c.delivered('s1', key('d', 'mid')); advance(10)
    await fill(c, 's1', key('d', 'new'), 30 * MiB); advance(10)
    c.dropSession('s1') // releases 'new' last
    expect(c.stats()).toMatchObject({ entries: 3, referenced: 0 }) // survive the session's end
    // 'old' is touched again by a hit from s2 and released: it becomes the most recently used.
    expect((await done(c.acquire('s2', 0, key('d', 'old'), fakeRun(1).run, ahead))).ok).toBe(true)
    c.delivered('s2', key('d', 'old')); advance(10)
    const f = fakeRun(1 * MiB)
    const a = c.acquire('s3', 0, key('d', 'x'), f.run, request) // 90 + 32 > 100: one eviction
    f.finish(); await done(a)
    expect(c.peek(key('d', 'mid'))).toBeUndefined()
    expect(c.peek(key('d', 'old'))).toBeDefined()
    expect(c.peek(key('d', 'new'))).toBeDefined()
    await c.close()
  })

  it('an unreferenced entry is gone after MEDIA_IDLE_TTL_MS on the injected clock; a referenced entry is not, however long it idles; releasing the reference starts the idle count', async () => {
    const { c, advance } = cacheWith()
    const free = await fill(c, 's1', key('d', 'free'), 1 * MiB)
    c.delivered('s1', key('d', 'free'))
    await fill(c, 's1', key('d', 'held'), 1 * MiB)
    advance(MEDIA_IDLE_TTL_MS - 1)
    await c.sweep()
    expect(c.peek(key('d', 'free'))).toBeDefined()
    advance(1)
    await c.sweep()
    expect(c.peek(key('d', 'free'))).toBeUndefined()
    expect('path' in free && existsSync(free.path)).toBe(false)
    advance(100 * MEDIA_IDLE_TTL_MS)
    await c.sweep()
    expect(c.peek(key('d', 'held'))).toBeDefined() // referenced: never expires
    c.delivered('s1', key('d', 'held')) // the idle count starts here
    advance(MEDIA_IDLE_TTL_MS - 1)
    await c.sweep()
    expect(c.peek(key('d', 'held'))).toBeDefined()
    advance(1)
    await c.sweep()
    expect(c.peek(key('d', 'held'))).toBeUndefined()
    expect(c.stats()).toMatchObject({ entries: 0, bytes: 0 })
    await c.close()
    expect(temp.live()).toBe(0)
  })

  it('dropDocument deletes every entry of that docId at once, referenced or not, and no other document\'s', async () => {
    const { c } = cacheWith()
    const a1 = await fill(c, 's1', key('A', '1'), 1 * MiB)
    await fill(c, 's1', key('A', '2'), 1 * MiB); c.delivered('s1', key('A', '2'))
    await fill(c, 's2', key('B', '1'), 1 * MiB)
    const f = fakeRun(1 * MiB)
    const inflight = c.acquire('s3', 0, key('A', '3'), f.run, ahead)
    await c.dropDocument('A')
    expect(c.peek(key('A', '1'))).toBeUndefined()
    expect(c.peek(key('A', '2'))).toBeUndefined()
    expect(c.inflight(key('A', '3'))).toBe(false)
    expect(f.state.aborted).toBe(true)
    expect(await done(inflight)).toEqual({ ok: false, reason: 'media-unavailable' })
    expect('path' in a1 && existsSync(a1.path)).toBe(false)
    expect(c.peek(key('B', '1'))).toBeDefined()
    expect(c.referencesOf('s1')).toBe(0)
    expect(c.referencesOf('s2')).toBe(1)
    await c.close()
  })

  it('a delivered key is released for that session and stays cached for another', async () => {
    const { c } = cacheWith()
    await fill(c, 's1', key('d', 'k'), 1 * MiB)
    expect((await done(c.acquire('s2', 0, key('d', 'k'), fakeRun(1).run, ahead))).ok).toBe(true)
    c.delivered('s1', key('d', 'k'))
    expect(c.referencesOf('s1')).toBe(0)
    expect(c.referencesOf('s2')).toBe(1)
    expect(c.stats()).toMatchObject({ entries: 1, referenced: 1 })
    c.delivered('s2', key('d', 'k'))
    expect(c.stats()).toMatchObject({ entries: 1, referenced: 0 })
    expect(c.peek(key('d', 'k'))).toBeDefined()
    await c.close()
  })

  it('a shared entry counts in full against each referencing session\'s look-ahead', async () => {
    const { c } = cacheWith({ sessionBytes: 64 * MiB })
    await fill(c, 's1', key('d', 'big'), 40 * MiB)
    // s2 references it too, and is then charged 40 MiB: another 32 MiB reservation no longer fits ahead.
    expect((await done(c.acquire('s2', 0, key('d', 'big'), fakeRun(1).run, ahead))).ok).toBe(true)
    expect(c.acquire('s2', 1, key('d', 'other'), fakeRun(1).run, ahead)).toEqual({ wait: true })
    expect(c.acquire('s1', 1, key('d', 'other2'), fakeRun(1).run, ahead)).toEqual({ wait: true })
    // A third session with no references has its whole look-ahead.
    const f = fakeRun(1 * MiB)
    const a = c.acquire('s3', 0, key('d', 'other'), f.run, ahead)
    expect('origin' in a).toBe(true)
    f.finish(); await done(a)
    await c.close()
  })

  it('server cache wholly referenced: ahead gets wait, a request gets media-busy', async () => {
    const { c } = cacheWith({ serverBytes: 64 * MiB })
    await fill(c, 's1', key('d', 'a'), 32 * MiB)
    await fill(c, 's2', key('d', 'b'), 32 * MiB)
    const f = fakeRun(1)
    expect(c.acquire('s3', 0, key('d', 'c'), f.run, ahead)).toEqual({ wait: true })
    expect(await done(c.acquire('s3', 0, key('d', 'c'), f.run, request))).toEqual({ ok: false, reason: 'media-busy' })
    expect(f.state.calls).toBe(0)
    expect(c.referencesOf('s3')).toBe(0)
    // A release wakes a waiting encode-ahead.
    const woke = c.room()
    c.delivered('s1', key('d', 'a'))
    await woke
    await c.close()
  })

  it('an encode in flight reserves 32 MiB and settles to its real size', async () => {
    const { c } = cacheWith()
    const f = fakeRun(3 * MiB)
    const a = c.acquire('s1', 0, key('d', 'k'), f.run, ahead)
    expect(MEDIA_INFLIGHT_RESERVE).toBe(32 * MiB)
    expect(c.stats()).toEqual({ entries: 0, bytes: 32 * MiB, referenced: 1, inflight: 1 })
    f.finish(); await done(a)
    expect(c.stats()).toEqual({ entries: 1, bytes: 3 * MiB, referenced: 1, inflight: 0 })
    await c.close()
  })

  it('dropSession releases references and aborts only the encodes left with no waiter', async () => {
    const { c } = cacheWith()
    const shared = fakeRun(1 * MiB), alone = fakeRun(1 * MiB)
    const a1 = c.acquire('s1', 0, key('d', 'shared'), shared.run, ahead)
    const a2 = c.acquire('s2', 0, key('d', 'shared'), fakeRun(1).run, ahead)
    const a3 = c.acquire('s1', 1, key('d', 'alone'), alone.run, ahead)
    await fill(c, 's1', key('d', 'done'), 1 * MiB)
    c.dropSession('s1')
    expect(await done(a1)).toEqual({ ok: false, reason: 'media-unavailable' })
    expect(await done(a3)).toEqual({ ok: false, reason: 'media-unavailable' })
    await Bun.sleep(0)
    expect(alone.state.aborted).toBe(true)
    expect(shared.state.aborted).toBe(false)
    expect(c.inflight(key('d', 'shared'))).toBe(true)
    expect(c.referencesOf('s1')).toBe(0)
    expect(c.peek(key('d', 'done'))).toBeDefined() // left cached
    shared.finish()
    expect((await done(a2)).ok).toBe(true)
    await c.close()
  })

  it('hit, joined and miss are reported for finished, in-flight and absent keys', async () => {
    const { c } = cacheWith()
    await fill(c, 's0', key('d', 'fin'), 1 * MiB)
    const f = fakeRun(1 * MiB)
    const miss = c.acquire('s1', 0, key('d', 'k'), f.run, ahead)
    const joined = c.acquire('s2', 0, key('d', 'k'), fakeRun(1).run, ahead)
    const hit = c.acquire('s3', 0, key('d', 'fin'), fakeRun(1).run, ahead)
    expect('origin' in miss && miss.origin).toBe('miss')
    expect('origin' in joined && joined.origin).toBe('joined')
    expect('origin' in hit && hit.origin).toBe('hit')
    f.finish()
    const [m, j] = await Promise.all([done(miss), done(joined)])
    expect(m.ok && j.ok && m.media === j.media).toBe(true)
    expect(f.state.calls).toBe(1)
    await c.close()
    expect(temp.live()).toBe(0)
  })

  it('a media-timeout reaches every joiner and is not cached; media-too-large is cached as a refusal (no reference)', async () => {
    const { c } = cacheWith()
    const t = fakeRun(1)
    const a = c.acquire('s1', 0, key('d', 't'), t.run, ahead), b = c.acquire('s2', 0, key('d', 't'), fakeRun(1).run, request)
    t.fail('media-timeout')
    expect(await done(a)).toEqual({ ok: false, reason: 'media-timeout' })
    expect(await done(b)).toEqual({ ok: false, reason: 'media-timeout' })
    expect(c.peek(key('d', 't'))).toBeUndefined()
    expect(c.referencesOf('s1')).toBe(0)
    const big = fakeRun(1)
    const x = c.acquire('s1', 0, key('d', 'big'), big.run, ahead)
    big.fail('media-too-large')
    expect(await done(x)).toEqual({ ok: false, reason: 'media-too-large' })
    expect(c.peek(key('d', 'big'))).toEqual({ refused: 'media-too-large', bytes: 0 })
    const again = fakeRun(1)
    const y = c.acquire('s2', 0, key('d', 'big'), again.run, request)
    expect('origin' in y && y.origin).toBe('hit')
    expect(await done(y)).toEqual({ ok: false, reason: 'media-too-large' })
    expect(again.state.calls).toBe(0)
    expect(c.referencesOf('s1') + c.referencesOf('s2')).toBe(0)
    await c.close()
  })
})
