/**
 * Piece 5 Task 13 (spec §4.2, §6; plan D1, D12, D21, D23; Review Focus 5): clean-up across the session lifecycle.
 * Every test ends on the invariants of `settled` (battery.ts), read through the harness's `MediaServices` while the
 * server is still up: no admission slot or place held, no reference left for the test's sessions, nothing in flight,
 * every file in the media temp dir a cache entry, and no child process.
 *
 * Two items carried from earlier reviews are pinned here: a revoked queue's pending admission ticket (Task 8), and a
 * render worker terminated mid-probe (Task 3).
 */
import { afterAll, describe, expect, it } from 'bun:test'
import { chmodSync, existsSync, linkSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, renameSync, rmSync, symlinkSync, writeFileSync } from 'node:fs'
import { spawn } from 'node:child_process'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { parseMediaShowOk, parseMediaStart } from '../../../src/stream/media-wire'
import { mediaDeck, narrationDeck } from '../../../src/testdata/show-decks'
import { createWorkerProbes, type ProbeRunReply } from '../../src/media/probe-host'
import { createMediaTemp } from '../../src/media/temp'
import { Pool } from '../../src/pool'
import { SHOW_BUILD_BYTES } from '../../src/show/build'
import { disposePool, start, type HarnessMedia } from '../harness'
import { alive, childPids, groupOf, planOf, processes, running, zombieChildren, session, settled, sidOf, source, spawnLog, until, withPart, type AuditClient } from './battery'
import { ffmpegForTests } from './ffmpeg-gate'

const info = await ffmpegForTests()
const scratch = mkdtempSync(join(tmpdir(), 'ov-t13-life-'))
afterAll(() => { disposePool(); rmSync(scratch, { recursive: true, force: true }) })

/** ~150 s of noise: a ~2.4 MB encode (three chunks), about a second of ffmpeg. */
const noise = (name = 'n150.m4a', seed = 1) => source(info!, scratch, name, `anoisesrc=c=pink:r=48000:a=0.5:seed=${seed}`, 150)
/** ~600 s of noise: an encode of several seconds (mid-encode tests). */
const longNoise = (name = 'n600.m4a', seed = 2) => source(info!, scratch, name, `anoisesrc=c=pink:r=48000:a=0.5:seed=${seed}`, 600, '32k')

const box = { id: 2, name: 'Box', prst: 'rect', off: ['3619500', '2476500'] as [string, string], ext: ['1905000', '1905000'] as [string, string], fill: '336699' }
const narration = (id: number, spid: number) => `<p:audio isNarration="1"><p:cMediaNode vol="80000"><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cMediaNode></p:audio>`
/** One narration per slide, slide i's part replaced by `parts[i]` (distinct bytes: distinct encodes). */
async function slidesDeck(parts: Uint8Array[]): Promise<Uint8Array> {
  const deck = await mediaDeck({ slides: parts.map(() => ({ boxes: [box], media: [{ spid: 3, file: 'a.m4a' as const }], timing: narration(3, 3) })) })
  const zip = await JSZip.loadAsync(deck)
  const names = Object.keys(zip.files).filter((f) => /^ppt\/media\/showMedia\d+\.m4a$/.test(f)).sort((a, b) => Number(/(\d+)\.m4a$/.exec(a)![1]) - Number(/(\d+)\.m4a$/.exec(b)![1]))
  if (names.length !== parts.length) throw new Error(`expected ${parts.length} parts, got ${names.join(',')}`)
  names.forEach((n, i) => zip.file(n, parts[i]!))
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
}

/** A wrapper around the real ffmpeg that first starts a sleeping descendant in the run's process group. */
function groupedFfmpeg(): string {
  const bin = mkdtempSync(join(scratch, 'ffbin-'))
  const wrapper = join(bin, 'ffmpeg')
  writeFileSync(wrapper, `#!/bin/sh\n/bin/sleep 61 </dev/null >/dev/null 2>&1 &\nexec '${info!.ffmpeg}' "$@"\n`)
  chmodSync(wrapper, 0o755)
  return wrapper
}
/** `f()` with `OV_FFMPEG` set to `ffmpeg` and `OV_FFPROBE` to the real one (the harness's media services read both at
 * start; ffprobe is otherwise looked for beside ffmpeg). */
async function withFfmpeg<T>(ffmpeg: string, f: () => Promise<T>): Promise<T> {
  const was = { OV_FFMPEG: process.env.OV_FFMPEG, OV_FFPROBE: process.env.OV_FFPROBE }
  process.env.OV_FFMPEG = ffmpeg
  process.env.OV_FFPROBE = info!.ffprobe
  try { return await f() } finally {
    for (const [k, v] of Object.entries(was)) { if (v === undefined) delete process.env[k]; else process.env[k] = v }
  }
}
const harness = (doc: Uint8Array, media: HarnessMedia = {}, extra: Parameters<typeof start>[0] = {}) => start({ doc, media, ...extra })
/** The plan of every unit in `units` (encode-ahead readies only uses a sent plan names), then `media-show`. */
async function play(c: AuditClient, units: number[] = [0], from = 0) {
  const recs = []
  for (const u of units) recs.push(...(await planOf(c, u)).media)
  const show = parseMediaShowOk(await c.mediaShow(from))
  return { recs, queued: show.queued }
}
/** A transport that holds every kind-9 send after the first under backpressure while `held()`. */
function heldTransport(held: () => boolean) {
  let sent = 0
  return {
    sent: () => sent,
    transport: {
      sendBinary: (ws: unknown, f: Uint8Array) => { const n = (ws as { sendBinary(f: Uint8Array): number }).sendBinary(f); if (f[5] === 9) sent++; return n },
      bufferedAmount: (ws: unknown) => (sent >= 1 && held() ? 64 << 20 : (ws as { getBufferedAmount(): number }).getBufferedAmount()),
    },
  }
}
async function closeAll(...cs: AuditClient[]) { for (const c of cs) { c.close(); await c.closed } }

describe.skipIf(!info)('media lifecycle (Task 13)', () => {
  it('close mid-delivery: ledgers empty, no pid, no file', async () => {
    const t = heldTransport(() => true)
    const h = await harness(await withPart(await narrationDeck(), await noise()), {}, { transport: t.transport })
    const c = await session(h)
    try {
      const { recs } = await play(c)
      c.send({ reqId: 4000, type: 'media-request', token: c.manifest!.session.token, mediaId: recs[0]!.mediaId })
      await until(() => t.sent() === 1, 60_000, 'the first chunk')
      await Bun.sleep(100) // parked in the high-water wait
      expect(h.stream.stats().media.deliveries).toBe(1)
      expect(h.media!.cache.referencesOf(sidOf(c))).toBeGreaterThan(0)
      const sid = sidOf(c)
      await closeAll(c)
      await until(() => h.stream.stats().media.deliveries === 0, 2000, 'the delivery to stop')
      expect(t.sent()).toBe(1)
      // The only session closed: its document closed, and its entries went with it (D12).
      await until(() => h.media!.cache.stats().entries === 0, 2000, 'the entries to go')
      await settled(h, [sid])
      expect(readdirSync(h.media!.temp.dir)).toEqual([])
    } finally { c.close(); await h.stop() }
  }, 120_000)

  it('close mid-encode kills the group when no other session waits on it, and leaves it running when one does', async () => {
    const deck = await withPart(await narrationDeck(), await longNoise())
    // Nobody else waits: a holder keeps the document open (its close would delete the entries anyway), A encodes ahead.
    // ffmpeg runs under a wrapper that leaves a sleeping descendant in the run's group (fix round 1): only a group
    // kill takes it too.
    {
      const h = await withFfmpeg(groupedFfmpeg(), () => harness(deck))
      const log = spawnLog() // after start: the startup ffmpeg probe runs on this thread too
      const holder = await session(h), a = await session(h)
      try {
        expect((await play(a)).queued).toBe(1)
        await until(() => log.ffmpeg().length === 1, 30_000, 'the encode to spawn')
        const pid = log.ffmpeg()[0]!.pid
        await until(() => groupOf(pid).length >= 2, 5000, 'the encode and its descendant')
        const members = groupOf(pid)
        expect(h.media!.cache.stats().inflight).toBe(1)
        const sa = sidOf(a)
        await closeAll(a)
        await until(() => groupOf(pid).length === 0, 3000, 'the encode group to die')
        expect(members.filter(running)).toEqual([])
        expect(h.media!.cache.stats()).toMatchObject({ entries: 0, inflight: 0 })
        await settled(h, [sa, sidOf(holder)])
      } finally { log.stop(); holder.close(); a.close(); await h.stop() }
    }
    // B waits on the same encode (joined): A's close leaves it running, and B gets the bytes.
    {
      const h = await harness(deck)
      const log = spawnLog() // after start: the startup ffmpeg probe runs on this thread too
      const a = await session(h), b = await session(h)
      try {
        await play(a)
        await until(() => log.ffmpeg().length === 1, 30_000, 'the encode to spawn')
        const pid = log.ffmpeg()[0]!.pid
        const { recs } = await play(b)
        b.send({ reqId: 4100, type: 'media-request', token: b.manifest!.session.token, mediaId: recs[0]!.mediaId })
        await until(() => h.media!.cache.inspect().entries[0]?.waiters === 2, 5000, 'B to join')
        const sa = sidOf(a)
        await closeAll(a)
        await until(() => h.media!.cache.referencesOf(sa) === 0, 2000, "A's references to go")
        expect(alive(pid)).toBe(true)
        expect(h.media!.cache.stats().inflight).toBe(1)
        const r = await b.mediaRequest(recs[0]!.mediaId, { reqId: 4100, send: false, ms: 60_000 })
        expect(r.end).toBeUndefined()
        expect(r.chunks).toHaveLength(parseMediaStart(r.start).chunks)
        expect(log.ffmpeg()).toHaveLength(1)
        expect(h.stream.audit.mediaDeliveries(sidOf(b)).map((x) => x.cacheState)).toEqual(['joined'])
        await settled(h, [sa, sidOf(b)])
      } finally { log.stop(); a.close(); b.close(); await h.stop() }
    }
  }, 180_000)

  it('media-end cancels queued encodes and releases references', async () => {
    const h = await harness(await slidesDeck([await longNoise(), await longNoise('n600b.m4a', 3), await noise()]))
    const log = spawnLog() // after start: the startup ffmpeg probe runs on this thread too
    const holder = await session(h), c = await session(h)
    try {
      expect((await play(c, [0, 1, 2])).queued).toBe(3)
      await until(() => log.ffmpeg().length === 1, 30_000, 'the first encode to spawn')
      const pid = log.ffmpeg()[0]!.pid
      expect(h.media!.cache.referencesOf(sidOf(c))).toBeGreaterThan(0)
      c.mediaEnd()
      await until(() => !alive(pid), 3000, 'the running encode to die')
      await until(() => h.media!.cache.referencesOf(sidOf(c)) === 0, 2000, 'the references to go')
      await Bun.sleep(500)
      // The two queued encodes never started.
      expect(log.ffmpeg()).toHaveLength(1)
      expect(h.media!.cache.stats()).toMatchObject({ entries: 0, inflight: 0 })
      await settled(h, [sidOf(c), sidOf(holder)])
    } finally { log.stop(); holder.close(); c.close(); await h.stop() }
  }, 120_000)

  it("session expiry releases the session's references; the cache entries remain and a later session of the same document starts no encode", async () => {
    // A token whose renewal is never acknowledged dies at its exp (4004); the holder acknowledges, so it keeps the
    // document open.
    const h = await harness(await narrationDeck(), {}, { tokenTtlSec: 3, renewLeadSec: 2 })
    const log = spawnLog() // after start: the startup ffmpeg probe runs on this thread too
    const a = await session(h), holder = await session(h)
    let later: AuditClient | undefined
    try {
      expect((await play(a)).queued).toBe(1)
      await until(() => h.media!.cache.stats().entries === 1 && h.media!.cache.referencesOf(sidOf(a)) > 0, 30_000, 'encode-ahead')
      const renew = await holder.next((m) => 'text' in m && m.text.type === 'renew', 10_000) as { text: Record<string, any> }
      holder.send({ reqId: 0, type: 'renew-ack', jti: renew.text.jti })
      expect((await a.closed).code).toBe(4004)
      await until(() => h.media!.cache.referencesOf(sidOf(a)) === 0, 2000, "A's references to go")
      expect(h.media!.cache.stats().entries).toBe(1)
      later = await session(h)
      const { recs } = await play(later)
      const r = await later.mediaRequest(recs[0]!.mediaId, { ms: 60_000 })
      expect(r.end).toBeUndefined()
      expect(log.ffmpeg()).toHaveLength(1)
      expect(h.stream.audit.mediaDeliveries(sidOf(later)).map((x) => x.cacheState)).toEqual(['hit'])
      await settled(h, [sidOf(a), sidOf(later)])
    } finally { log.stop(); a.close(); holder.close(); later?.close(); await h.stop() }
  }, 60_000)

  it('an unreferenced entry is deleted after the idle TTL, on the injected clock; a later session of that document encodes again', async () => {
    let now = 1_000_000
    const h = await harness(await narrationDeck(), { testOnly: { now: () => now, idleTtlMs: 60_000 } })
    const log = spawnLog() // after start: the startup ffmpeg probe runs on this thread too
    const holder = await session(h), a = await session(h)
    let later: AuditClient | undefined
    try {
      const { recs } = await play(a)
      expect((await a.mediaRequest(recs[0]!.mediaId, { ms: 60_000 })).end).toBeUndefined()
      await until(() => h.media!.cache.referencesOf(sidOf(a)) === 0, 2000, 'the delivery to release its reference')
      const path = h.media!.cache.inspect().entries[0]!.path!
      expect(existsSync(path)).toBe(true)
      now += 59_999
      await h.media!.cache.sweep()
      expect(h.media!.cache.stats().entries).toBe(1) // not yet
      now += 1
      await h.media!.cache.sweep()
      expect(h.media!.cache.stats().entries).toBe(0)
      expect(existsSync(path)).toBe(false)
      await settled(h, [sidOf(a)])
      later = await session(h)
      const p = await play(later)
      expect((await later.mediaRequest(p.recs[0]!.mediaId, { ms: 60_000 })).end).toBeUndefined()
      expect(log.ffmpeg()).toHaveLength(2)
      expect(h.stream.audit.mediaDeliveries(sidOf(later)).map((x) => x.cacheState)).toEqual(['miss'])
      await settled(h, [sidOf(a), sidOf(later)])
    } finally { log.stop(); holder.close(); a.close(); later?.close(); await h.stop() }
  }, 60_000)

  it('document revoke deletes its entries at once and kills its in-flight encode, even when referenced', async () => {
    // Slide 0's short narration encodes ahead and stays referenced (undelivered); slide 1's long one is then in flight.
    const h = await harness(await slidesDeck([await noise(), await longNoise()]))
    const log = spawnLog() // after start: the startup ffmpeg probe runs on this thread too
    const c = await session(h)
    try {
      await play(c, [0, 1])
      await until(() => log.ffmpeg().length === 2, 60_000, 'the second encode to spawn')
      const done = h.media!.cache.inspect().entries.find((e) => e.state === 'done')!
      expect(done.refs).toBe(1)
      const pid = log.ffmpeg()[1]!.pid
      expect(alive(pid)).toBe(true)
      await h.media!.cache.dropDocument(c.manifest!.docRef)
      // At once: the referenced entry and its file, and the flight (dropDocument resolves once it has settled).
      expect(h.media!.cache.stats()).toMatchObject({ entries: 0, inflight: 0 })
      expect(existsSync(done.path!)).toBe(false)
      expect(alive(pid)).toBe(false)
      await settled(h, [sidOf(c)])
    } finally { log.stop(); c.close(); await h.stop() }
  }, 120_000)

  it('server shutdown removes the per-process temp dir and every cache entry', async () => {
    // Slide 0's narration is delivered; slide 1's long one is still encoding ahead (a slot, a flight, a reference).
    const h = await harness(await slidesDeck([await noise(), await longNoise()]))
    const c = await session(h)
    try {
      const { recs } = await play(c, [0, 1])
      expect((await c.mediaRequest(recs[0]!.mediaId, { ms: 60_000 })).end).toBeUndefined()
      await until(() => h.media!.cache.stats().inflight === 1 && h.media!.admission.running === 1, 30_000, 'the second encode')
      expect(h.media!.cache.stats().entries).toBe(1)
      expect(h.media!.cache.referencesOf(sidOf(c))).toBeGreaterThan(0)
      const dir = h.media!.temp.dir
      // As main.ts does on shutdown.
      await h.media!.close()
      expect(h.media!.temp.live()).toBe(0)
      expect(h.media!.cache.stats().entries).toBe(0)
      expect(existsSync(dir)).toBe(false)
      expect(childPids()).toEqual([])
      // And nothing of the session is left held (fix round 1): admission, references, flights.
      expect([h.media!.admission.running, h.media!.admission.queued]).toEqual([0, 0])
      expect(h.media!.cache.referencesOf(sidOf(c))).toBe(0)
      expect(h.media!.cache.stats().inflight).toBe(0)
    } finally { c.close(); await h.stop() }
  }, 60_000)

  it('a revoked queue gives back its pending admission ticket at once, before end() (Task 8 carry)', async () => {
    const h = await harness(await narrationDeck(), { testOnly: { slots: 1 } })
    const c = await session(h)
    const held = h.media!.admission.slot('holder', 'request', new AbortController().signal)
    if ('refused' in held) throw new Error('no slot to hold')
    const lease = await held
    try {
      // The one slot is held elsewhere: encode-ahead waits in a queue place (its ticket).
      expect((await play(c)).queued).toBe(1)
      await until(() => h.media!.admission.queued === 1, 5000, 'the ticket')
      expect(h.media!.admission.snapshot().queue).toEqual([sidOf(c)])
      await h.media!.cache.dropDocument(c.manifest!.docRef)
      // The socket is still open (no end()): the revoked queue gives its place back at once.
      await until(() => h.media!.admission.queued === 0, 1000, 'the ticket to be given back')
      expect(c.ws.readyState).toBe(WebSocket.OPEN)
      lease.release()
      await Bun.sleep(100)
      expect(h.media!.admission.running).toBe(0) // the freed slot went to nobody
      await settled(h, [sidOf(c)])
    } finally { lease.release(); c.close(); await h.stop() }
  }, 60_000)
})

describe.skipIf(!info)('a render worker terminated mid-job (Task 3 carry; fix rounds 1–2)', () => {
  /** A script that never ends, as a group of three: the leader (`/bin/sh <script> <args>`, so its command line keeps
   * its arguments) and two sleeping descendants in its group. A leader-only kill leaves the descendants. */
  function sleeper(name: string): string {
    const bin = mkdtempSync(join(scratch, 'bin-'))
    const path = join(bin, name)
    writeFileSync(path, '#!/bin/sh\n/bin/sleep 31 </dev/null >/dev/null 2>&1 &\n/bin/sleep 30 </dev/null >/dev/null 2>&1\n')
    chmodSync(path, 0o755)
    return path
  }
  const job = (docRef: string) => ({ docRef, contentPolicy: 'visible-only' as const, unitIndex: 0, scale: 1 as const, motion: 'full' as const, mark: { kind: 'fixed' as const }, strictMarks: false, media: true })
  /** A pool over its own media dir, with a stand-in ffprobe that never answers, the given media deadline (test-only)
   * and `o`. */
  async function hungPool(deadlineMs = 1500, o: Partial<ConstructorParameters<typeof Pool>[0]> = {}) {
    const root = mkdtempSync(join(scratch, 'root-'))
    const temp = await createMediaTemp(root)
    const pool = new Pool({ workers: 1, media: { info: { ...info!, ffprobe: sleeper('ffprobe') }, dir: temp.dir }, testHooks: true, mediaDeadlineMs: deadlineMs, ...o })
    await pool.open('hung', await narrationDeck(), 'visible-only')
    return { pool, temp, done: async () => { pool.dispose(); await temp.remove(); rmSync(root, { recursive: true, force: true }) } }
  }
  /** The probe's group, once all three are up: our direct child whose input is a staged file in the media dir. */
  async function probeGroup(dir: string) {
    let leader: ReturnType<typeof processes>[number] | undefined
    await until(() => { leader = processes().find((p) => p.command.includes(`file:${dir}/`) && p.ppid === process.pid && !p.stat.startsWith('Z')); return !!leader && groupOf(leader.pid).length >= 3 }, 5000, 'the stand-in probe group')
    return leader!
  }

  it('a worker retired mid-probe leaves no probe group (leader and descendants), no zombie and no staged file', async () => {
    const { pool, temp, done } = await hungPool()
    try {
      const r = pool.showPlan('s1', job('hung'), SHOW_BUILD_BYTES)
      const leader = await probeGroup(temp.dir)
      expect(leader.ppid).toBe(process.pid)
      expect(leader.pgid).toBe(leader.pid) // its own group (detached)
      const members = groupOf(leader.pid)
      const out = await r
      expect('refused' in out && out.refused).toBe('show-plan-timeout')
      expect(pool.stats().retired).toBe(1)
      // Gone, zombies included: the leader is reaped by the thread that spawned it, which outlives the worker.
      await until(() => groupOf(leader.pid).length === 0 && zombieChildren().length === 0 && readdirSync(temp.dir).length === 0, 3000, 'the probe group to be killed and reaped').catch(() => {})
      expect(members.filter(running)).toEqual([])
      expect(groupOf(leader.pid)).toEqual([])
      expect(zombieChildren()).toEqual([])
      expect(childPids()).toEqual([])
      expect(readdirSync(temp.dir)).toEqual([])
      expect(pool.stats().probesRunning).toBe(0)
    } finally { await done() }
  }, 30_000)

  it('twenty retirements mid-probe: no zombie and no probe handle left', async () => {
    const { pool, temp, done } = await hungPool(300)
    try {
      for (let i = 0; i < 20; i++) {
        const out = await pool.showPlan(`s${i}`, job('hung'), SHOW_BUILD_BYTES)
        expect('refused' in out && out.refused).toBe('show-plan-timeout')
      }
      expect(pool.stats().retired).toBe(20)
      await until(() => zombieChildren().length === 0 && childPids().length === 0 && pool.stats().probesRunning === 0, 5000, 'every probe reaped').catch(() => {})
      expect(zombieChildren()).toEqual([])
      expect(childPids()).toEqual([])
      expect(pool.stats().probesRunning).toBe(0)
      expect(readdirSync(temp.dir)).toEqual([])
    } finally { await done() }
  }, 60_000)

  it("retirement kills only its own worker's runs: unrelated child groups of this process survive", async () => {
    // Our children too, each in its own group, their command lines kept alive through the check: one names the media
    // dir under another worker's prefix, one mentions the retiring probe's own staged input.
    const { pool, temp, done } = await hungPool()
    const script = sleeper('bystander')
    const alt = join(temp.dir, 'w000000000000-0123456789abcdef0123456789abcdef.bin')
    const others: ReturnType<typeof spawn>[] = []
    try {
      others.push(spawn(script, [alt], { detached: true, stdio: 'ignore' }))
      const r = pool.showPlan('s1', job('hung'), SHOW_BUILD_BYTES)
      const leader = await probeGroup(temp.dir)
      const staged = /file:(\S+)/.exec(leader.command)![1]!
      others.push(spawn(script, [`--note=${staged}`], { detached: true, stdio: 'ignore' }))
      await until(() => others.every((o) => groupOf(o.pid!).length >= 3), 5000, 'the unrelated groups')
      const bystanders = others.flatMap((o) => groupOf(o.pid!))
      const lines = () => others.map((o) => processes().find((p) => p.pid === o.pid)?.command ?? '')
      expect(lines()[0]).toContain(alt)
      expect(lines()[1]).toContain(staged)
      expect(others.every((o) => processes().find((p) => p.pid === o.pid)?.ppid === process.pid)).toBe(true)
      expect('refused' in (await r)).toBe(true)
      await until(() => groupOf(leader.pid).length === 0, 3000, 'the probe group to go').catch(() => {})
      expect(groupOf(leader.pid)).toEqual([])
      expect(bystanders.every(running)).toBe(true)
      // Still carrying their arguments after the retirement.
      expect(lines()[0]).toContain(alt)
      expect(lines()[1]).toContain(staged)
    } finally {
      for (const o of others) { try { process.kill(-o.pid!, 'SIGKILL') } catch { /* gone */ } }
      await done()
    }
  }, 30_000)

  describe('staged-file identity (fix round 3)', () => {
    const prefix = 'w0123456789ab-'
    const name = (p = prefix, n = 0) => `${p}${String(n).padStart(32, '0')}.bin`
    async function probeDir() {
      const root = mkdtempSync(join(scratch, 'root-'))
      const temp = await createMediaTemp(root)
      const fx = new Uint8Array(readFileSync(join(import.meta.dir, '../fixtures/media/a.m4a')))
      return { root, temp, fx, done: async () => { await temp.remove(); rmSync(root, { recursive: true, force: true }) } }
    }
    const cases: Array<[string, (dir: string, root: string, fx: Uint8Array) => string]> = [
      ["a valid-prefix symlink to another worker's staged file", (dir, _r, fx) => { const other = join(dir, name('wffffffffffff-')); writeFileSync(other, fx, { mode: 0o600 }); const p = join(dir, name()); symlinkSync(other, p); return p }],
      ['a valid-prefix symlink to a file outside the media dir', (dir, root, fx) => { const out = join(root, 'outside.m4a'); writeFileSync(out, fx, { mode: 0o600 }); const p = join(dir, name()); symlinkSync(out, p); return p }],
      ['a valid-prefix directory', (dir) => { const p = join(dir, name()); mkdirSync(p); return p }],
      ['a valid-prefix hard link (nlink 2)', (dir, root, fx) => { const out = join(root, 'linked.m4a'); writeFileSync(out, fx, { mode: 0o600 }); const p = join(dir, name()); linkSync(out, p); return p }],
    ]
    for (const [what, make] of cases) {
      it(`${what} is refused and spawns nothing`, async () => {
        const { root, temp, fx, done } = await probeDir()
        const log = spawnLog()
        const posted: ProbeRunReply[] = []
        try {
          const path = make(temp.dir, root, fx)
          const w = createWorkerProbes(info!, temp.dir, prefix, (r) => posted.push(r))
          w.request({ kind: 'probe-run', id: 1, path, demuxer: 'mov', wallMs: 1000 })
          await Bun.sleep(50)
          expect(log.spawns).toHaveLength(0)
          expect(posted).toEqual([{ type: 'probe-result', id: 1, ok: false, error: expect.stringMatching(/staged file/) }])
          expect(w.running).toBe(0)
        } finally { log.stop(); await done() }
      }, 30_000)
    }
    it('a staged file replaced while its probe runs fails the result closed', async () => {
      const { temp, fx, done } = await probeDir()
      const posted: ProbeRunReply[] = []
      try {
        const path = join(temp.dir, name())
        writeFileSync(path, fx, { mode: 0o600 })
        // The stand-in never answers: the 300 ms wall cap ends it. Meanwhile the entry is swapped (a new inode).
        const w = createWorkerProbes({ ...info!, ffprobe: sleeper('ffprobe') }, temp.dir, prefix, (r) => posted.push(r))
        w.request({ kind: 'probe-run', id: 1, path, demuxer: 'mov', wallMs: 300 })
        expect(w.running).toBe(1)
        const swap = join(temp.dir, 'swap.tmp')
        writeFileSync(swap, fx, { mode: 0o600 })
        renameSync(swap, path)
        await until(() => posted.length === 1, 5000, 'the result')
        expect(posted[0]).toEqual({ type: 'probe-result', id: 1, ok: false, error: expect.stringMatching(/changed/) })
      } finally { await done() }
    }, 30_000)
  })

  it('a retired probe host is released as soon as it has no run: at once, or when its last run settles (fix round 3)', async () => {
    const root = mkdtempSync(join(scratch, 'root-'))
    const temp = await createMediaTemp(root)
    const prefix = 'w0123456789ab-'
    const path = join(temp.dir, `${prefix}${'0'.repeat(32)}.bin`)
    writeFileSync(path, new Uint8Array(readFileSync(join(import.meta.dir, '../fixtures/media/a.m4a'))), { mode: 0o600 })
    try {
      let idleGone = 0, busyGone = 0
      const idle = createWorkerProbes(info!, temp.dir, prefix, () => {}, () => { idleGone++ })
      idle.retire()
      expect(idleGone).toBe(1)
      const busy = createWorkerProbes({ ...info!, ffprobe: sleeper('ffprobe') }, temp.dir, prefix, () => {}, () => { busyGone++ })
      busy.request({ kind: 'probe-run', id: 1, path, demuxer: 'mov', wallMs: 2000 })
      busy.retire()
      expect(busy.running).toBe(1) // killed, not yet reaped
      expect(busyGone).toBe(0)
      await until(() => busy.running === 0, 3000, 'the run to settle')
      expect(busyGone).toBe(1)
      busy.retire()
      expect(busyGone).toBe(1) // once
    } finally { await temp.remove(); rmSync(root, { recursive: true, force: true }) }
  }, 30_000)

  it('ten worker replacements with no stats() call leave no more probe hosts than live workers (fix round 3)', async () => {
    const { pool, done } = await hungPool(300)
    try {
      for (let i = 0; i < 10; i++) {
        const out = await pool.showPlan(`s${i}`, job('hung'), SHOW_BUILD_BYTES)
        expect('refused' in out && out.refused).toBe('show-plan-timeout')
      }
      await until(() => pool.probeHostsForTests() <= 1, 3000, 'the retired hosts to go').catch(() => {})
      expect(pool.probeHostsForTests()).toBeLessThanOrEqual(1)
    } finally { await done() }
  }, 60_000)

  it('a probe request after its worker is retired spawns nothing; a result for a retired worker is dropped', async () => {
    const root = mkdtempSync(join(scratch, 'root-'))
    const temp = await createMediaTemp(root)
    const prefix = 'w0123456789ab-'
    const staged = join(temp.dir, `${prefix}${'0'.repeat(32)}.bin`)
    writeFileSync(staged, new Uint8Array(readFileSync(join(import.meta.dir, '../fixtures/media/a.m4a'))), { mode: 0o600 })
    const log = spawnLog()
    const posted: ProbeRunReply[] = []
    try {
      // Refused before any spawn: not this worker's file, another dir, a bad demuxer or wall cap.
      const live = createWorkerProbes({ ...info!, ffprobe: sleeper('ffprobe') }, temp.dir, prefix, (r) => posted.push(r))
      live.request({ kind: 'probe-run', id: 1, path: join(temp.dir, `w999999999999-${'0'.repeat(32)}.bin`), demuxer: 'mov', wallMs: 1000 })
      live.request({ kind: 'probe-run', id: 2, path: join(root, `${prefix}${'0'.repeat(32)}.bin`), demuxer: 'mov', wallMs: 1000 })
      live.request({ kind: 'probe-run', id: 3, path: staged, demuxer: 'hls' as never, wallMs: 1000 })
      live.request({ kind: 'probe-run', id: 4, path: staged, demuxer: 'mov', wallMs: 60_000 })
      expect(posted.map((r) => [r.id, r.ok])).toEqual([[1, false], [2, false], [3, false], [4, false]])
      expect(log.spawns).toEqual([])
      // A live run, then retirement: it is killed and reaped, and its result is never posted.
      live.request({ kind: 'probe-run', id: 5, path: staged, demuxer: 'mov', wallMs: 2000 })
      expect(log.spawns).toHaveLength(1)
      expect(live.running).toBe(1)
      live.retire()
      await until(() => live.running === 0, 3000, 'the run to be reaped')
      await Bun.sleep(50)
      expect(posted.map((r) => r.id)).toEqual([1, 2, 3, 4])
      expect(zombieChildren()).toEqual([])
      // After retirement: refused silently, nothing spawned.
      live.request({ kind: 'probe-run', id: 6, path: staged, demuxer: 'mov', wallMs: 1000 })
      expect(log.spawns).toHaveLength(1)
      expect(live.running).toBe(0)
      expect(posted).toHaveLength(4)
      expect(childPids()).toEqual([])
    } finally { log.stop(); await temp.remove(); rmSync(root, { recursive: true, force: true }) }
  }, 30_000)

  it('a worker retired after mediaPart wrote its file leaves no file and no child', async () => {
    const root = mkdtempSync(join(scratch, 'root-'))
    const temp = await createMediaTemp(root)
    // Real ffmpeg: the refs probe completes; the worker then wedges once the part file is written.
    const pool = new Pool({ workers: 1, media: { info: info!, dir: temp.dir }, testHooks: true, mediaDeadlineMs: 300, mediaWedge: 'after-part' })
    try {
      await pool.open('part', await narrationDeck(), 'visible-only')
      const refs = await pool.mediaRefs('s1', 'part', 'visible-only', 0)
      const partPath = refs.refs[0]!.partPath
      expect(partPath).toMatch(/^\/?ppt\/media\//)
      await expect(pool.mediaPart('s1', 'part', 'visible-only', partPath, temp.dir)).rejects.toThrow(/deadline/)
      expect(pool.stats().retired).toBe(1)
      expect(readdirSync(temp.dir)).toEqual([])
      expect(childPids()).toEqual([])
    } finally { pool.dispose(); await temp.remove(); rmSync(root, { recursive: true, force: true }) }
  }, 30_000)
})
