/**
 * Piece 5 Task 8, fix round 2: the queue + cache + admission invariants under seeded random interleavings.
 *
 * Several sessions on two documents run start, cursor, pipelined requests, delivered, end, abort (an end not
 * awaited), dropDocument, a TTL sweep and encode completion (ok, timeout, unreadable, too-large) in random order,
 * over an in-memory temp, a fake transcode held until completed, an injected clock and a small admission class.
 * After every step the invariants below hold; at quiescence every request has settled and, once every session ended,
 * nothing is referenced, in flight or admitted. A failing schedule prints its seed and a shrunk op trace.
 *
 * `OV_FUZZ_SCHEDULES` (default 3000) sets the number of schedules; `OV_FUZZ_SEED` replays one.
 */
import { describe, expect, it } from 'bun:test'
import { createMediaAdmission } from '../../src/media/admission'
import { createEncodeCache, MEDIA_IDLE_TTL_MS } from '../../src/media/cache'
import type { FfmpegInfo } from '../../src/media/ffmpeg'
import type { MediaItem } from '../../src/media/jobs'
import { encodeKeyId, encodeKeyOf } from '../../src/media/params'
import type { ProbedMedia } from '../../src/media/probe'
import { createSessionMediaQueue, type MediaRequestResult, type SessionMediaQueue } from '../../src/media/queue'
import type { MediaFile, MediaTemp } from '../../src/media/temp'
import type { transcode } from '../../src/media/transcode'

const MiB = 1 << 20
const INFO: FfmpegInfo = { ffmpeg: '/x', ffprobe: '/x', version: '8.1.1', platform: 'darwin' }
const SESSION_BYTES = 70 * MiB, SERVER_BYTES = 100 * MiB
const SIZES: Record<string, number> = { p0: 10 * MiB, p1: 20 * MiB, p2: 30 * MiB, p3: 15 * MiB, p4: 25 * MiB }
/** Three slides per document; `#` marks a second use (another mediaId) of the same part. */
const DECK = [['p0', 'p1'], ['p2', 'p0#2'], ['p3', 'p4', 'p1#2']]
const DOCS = ['D0', 'D1']
const SESSIONS = 4

function rng(seed: number) {
  let a = seed >>> 0
  return () => { a = (a + 0x6d2b79f5) >>> 0; let t = a; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296 }
}

type Outcome = 'ok' | 'media-timeout' | 'media-unreadable' | 'media-too-large'
type Op =
  | { t: 'start'; s: number; u: number } | { t: 'cursor'; s: number; u: number } | { t: 'request'; s: number; i: number }
  | { t: 'deliver'; s: number; i: number } | { t: 'end'; s: number } | { t: 'abort'; s: number } | { t: 'drop'; d: number }
  | { t: 'sweep' } | { t: 'complete'; i: number; outcome: Outcome }
  /** Pipelined requests of one session in the same turn. */
  | { t: 'burst'; s: number; is: number[] }
  /** Two ops in the same turn, no flush between (cross-session interleavings). */
  | { t: 'together'; a: Op; b: Op }

function genOps(seed: number, n: number): Op[] {
  const r = rng(seed)
  const pick = <T>(xs: readonly T[]) => xs[Math.floor(r() * xs.length)]!
  const one = (): Op => {
    const x = r(), s = Math.floor(r() * SESSIONS)
    if (x < 0.1) return { t: 'burst', s, is: Array.from({ length: 2 + Math.floor(r() * 2) }, () => Math.floor(r() * 7)) }
    if (x < 0.2) return { t: 'together', a: plain(), b: plain() }
    return plain()
  }
  const plain = (): Op => { const o: Op[] = []; gen1(o); return o[0]! }
  const gen1 = (ops: Op[]) => {
    const x = r(), s = Math.floor(r() * SESSIONS)
    if (x < 0.12) ops.push({ t: 'start', s, u: Math.floor(r() * DECK.length) })
    else if (x < 0.22) ops.push({ t: 'cursor', s, u: Math.floor(r() * DECK.length) })
    else if (x < 0.45) ops.push({ t: 'request', s, i: Math.floor(r() * 7) })
    else if (x < 0.55) ops.push({ t: 'deliver', s, i: Math.floor(r() * 4) })
    else if (x < 0.58) ops.push({ t: 'end', s })
    else if (x < 0.61) ops.push({ t: 'abort', s })
    else if (x < 0.63) ops.push({ t: 'drop', d: Math.floor(r() * DOCS.length) })
    else if (x < 0.66) ops.push({ t: 'sweep' })
    else ops.push({ t: 'complete', i: Math.floor(r() * 4), outcome: pick(['ok', 'ok', 'ok', 'media-timeout', 'media-unreadable', 'media-too-large'] as const) })
  }
  return Array.from({ length: n }, one)
}

const flush = () => new Promise<void>((ok) => setImmediate(ok))
/** Totals over every schedule, so a run shows it exercised the paths it claims to. */
const totals = { encodes: 0, answers: new Map<string, number>() }

/** One schedule. Resolves with the first violated invariant (undefined: all held). */
async function runSchedule(ops: readonly Op[]): Promise<string | undefined> {
  // In-memory temp: a path exists while its handle is undisposed.
  const files = new Set<string>()
  let nextFile = 0
  const handle = (path: string): MediaFile => { files.add(path); return { path, dispose: async () => { files.delete(path) } } }
  const temp: MediaTemp = {
    dir: '/mem', create: async () => handle(`/mem/c${nextFile++}`), adopt: (p) => handle(p), reserve: () => handle(`/mem/o${nextFile++}.m4a`),
    live: () => files.size, bytes: () => 0, remove: async () => { files.clear() },
  }
  let t = 1_000_000
  const cache = createEncodeCache({ temp, sessionBytes: SESSION_BYTES, serverBytes: SERVER_BYTES, now: () => t })
  const inner = createMediaAdmission({ slots: 1, queue: 1 })
  // A queued place may be given up only when its session ends or its document is revoked (ruling, fix round 2).
  const admission: typeof inner = {
    slot(sid, mode, signal) {
      const r = inner.slot(sid, mode, signal)
      if ('refused' in r) return r
      void r.catch(() => {
        const s = sess.find((x) => x.sid === sid)
        const placed = placedBy.get(r)
        if (placed && s && !s.ended && !revokedAt.has(s.docId)) violations.push(`${sid}: a queued admission place was given up while the session was live`)
      })
      // Whether the waiter holds a queue place now (it may be promoted later: re-checked on every step).
      placedBy.set(r, inner.snapshot().queue.includes(sid))
      sidOf.set(r, sid)
      pendingTickets.add(r)
      void r.then(() => pendingTickets.delete(r), () => pendingTickets.delete(r))
      return r
    },
    upgrade: (t) => inner.upgrade(t),
    snapshot: () => inner.snapshot(),
    get running() { return inner.running },
    get queued() { return inner.queued },
  }
  const placedBy = new WeakMap<Promise<unknown>, boolean>()
  const sidOf = new WeakMap<Promise<unknown>, string>()
  const pendingTickets = new Set<Promise<unknown>>()
  const violations: string[] = []
  const lastOutcome = new Map<string, Outcome | 'abort'>()
  const calls: Array<{ docId: string; settle(o: Outcome | 'abort'): void; done: boolean }> = []
  const revokedAt = new Map<string, number>()
  let step = 0
  const startedAfterRevoke: string[] = []
  const fake: typeof transcode = async (_i, job, signal) => {
    const entry = job.entry
    let settle!: (o: Outcome | 'abort') => void
    const gate = new Promise<Outcome | 'abort'>((ok) => { settle = ok })
    const call = { docId: (job as { docId?: string }).docId ?? '', settle, done: false }
    calls.push(call)
    totals.encodes++
    const onAbort = () => settle('abort')
    if (signal.aborted) settle('abort'); else signal.addEventListener('abort', onAbort, { once: true })
    const o = await gate
    call.done = true
    lastOutcome.set(encodeKeyId(encodeKeyOf(call.docId, entry, INFO)), o)
    signal.removeEventListener('abort', onAbort)
    if (o === 'abort' || signal.aborted) { files.delete(job.out); throw signal.reason ?? new Error('abort') }
    if (o !== 'ok') { files.delete(job.out); return { ok: false, reason: o } }
    files.add(job.out)
    return { ok: true, bytes: SIZES[entry.partPath.split('/').pop()!]!, durationMs: 2000, contentType: 'audio/mp4', paramsSha: 'p', ffmpegVersion: INFO.version }
  }

  interface Sess { sid: string; docId: string; q?: SessionMediaQueue; answered: Array<{ key: MediaItem['key']; path: string }>; pending: Array<{ p: Promise<MediaRequestResult>; settled: boolean }>; ended: boolean; started?: boolean }
  const items = (sid: string, docId: string, u: number): MediaItem[] => DECK[u]!.map((name) => {
    const base = name.split('#')[0]!
    const e: ProbedMedia = { k: 0, partPath: `/ppt/media/${base}`, kind: 'audio', hasAudio: true, durationMs: 2000, sourceHash: `h-${base}`, demuxer: 'mov' }
    return {
      mediaId: `${sid}~${u}~${name}`, unit: u, ref: { kind: 'audio', k: 0 }, key: encodeKeyOf(docId, e, INFO), partPath: e.partPath,
      // The encode job carries its document so the fake can tell which document an encode is for.
      job: (out: string, src: string) => Object.assign({ entry: e, src, out }, { docId }),
    }
  })
  const sess: Sess[] = Array.from({ length: SESSIONS }, (_, i) => {
    const docId = DOCS[i % DOCS.length]!
    const s: Sess = { sid: `s${i}`, docId, answered: [], pending: [], ended: false }
    s.q = createSessionMediaQueue({
      sid: s.sid, units: DECK.length, info: INFO, temp, admission, cache, transcode: fake,
      part: async () => ({ path: '/mem/part', dispose: async () => {} }),
      items: async (u) => items(s.sid, docId, u), quiesceMs: 50,
    })
    return s
  })
  const allIds = (s: Sess) => DECK.flatMap((names, u) => names.map((n) => `${s.sid}~${u}~${n}`))

  const check = (where: string): string | undefined => {
    const snap = cache.inspect()
    for (const [sid, v] of Object.entries(snap.sessions)) if (v.bytes > SESSION_BYTES) return `${where}: ${sid} references ${v.bytes} > ${SESSION_BYTES}`
    if (snap.bytes > SERVER_BYTES) return `${where}: server bytes ${snap.bytes} > ${SERVER_BYTES}`
    for (const e of snap.entries) {
      if (e.refused && e.refused !== 'media-too-large' && e.refused !== 'media-unreadable') return `${where}: cached ${e.refused}`
      if (e.state === 'inflight' && e.waiters === 0) return `${where}: in-flight entry ${e.id.slice(0, 8)} has no waiter`
      if (e.path && !files.has(e.path)) return `${where}: entry ${e.id.slice(0, 8)} file unlinked while cached`
      if (revokedAt.has(e.docId)) return `${where}: entry for revoked ${e.docId}`
    }
    for (const s of sess) {
      if (s.ended || revokedAt.has(s.docId)) continue
      for (const a of s.answered) {
        const id = encodeKeyId(a.key)
        const e = snap.entries.find((x) => x.id === id)
        if (!e || e.path !== a.path || !files.has(a.path)) return `${where}: ${s.sid}'s undelivered answer ${id.slice(0, 8)} was evicted or unlinked`
        const pins = snap.sessions[s.sid]?.refs.find((r) => r.id === id)?.pins ?? 0
        const want = s.answered.filter((x) => encodeKeyId(x.key) === id).length
        if (pins < want) return `${where}: ${s.sid} holds ${pins} pins for ${want} undelivered answers`
      }
    }
    if (startedAfterRevoke.length) return `${where}: an encode started for revoked ${startedAfterRevoke[0]}`
    return undefined
  }

  const watchCalls = () => { for (const c of calls) if (!('seen' in c)) { (c as { seen?: number }).seen = step; if (revokedAt.has(c.docId)) startedAfterRevoke.push(c.docId) } }

  try {
    const request = (s: Sess, i: number) => {
      const id = allIds(s)[i % allIds(s).length]!
      const rec = { p: s.q!.request(id), settled: false }
      s.pending.push(rec)
      void rec.p.then((r) => {
        rec.settled = true
        const k = r.ok ? 'ok' : r.reason
        totals.answers.set(k, (totals.answers.get(k) ?? 0) + 1)
        if (r.ok) { if (!s.ended) s.answered.push({ key: items(s.sid, s.docId, Number(id.split('~')[1])).find((x) => x.mediaId === id)!.key, path: r.path }); return }
        // media-unavailable comes only from an ended session or a revoked document (the fake never fails otherwise).
        if (r.reason === 'media-unavailable' && !s.ended && !revokedAt.has(s.docId)) violations.push(`${s.sid}: ${id} answered media-unavailable with its session live and its document unrevoked`)
      })
    }
    /** One op; nothing awaited (a `together` runs two in the same turn). */
    const exec = (op: Op): Promise<unknown> | void => {
      const s = 's' in op ? sess[op.s]! : undefined
      if (s && s.ended) return
      switch (op.t) {
        case 'start': s!.started = true; void s!.q!.start(op.u); return
        case 'cursor': s!.q!.cursor(op.u); return
        case 'request': request(s!, op.i); return
        case 'burst': for (const i of op.is) request(s!, i); return
        case 'deliver': {
          if (!s!.answered.length) return
          const [a] = s!.answered.splice(op.i % s!.answered.length, 1)
          cache.delivered(s!.sid, a!.key)
          return
        }
        case 'end': s!.ended = true; s!.answered = []; return s!.q!.end()
        case 'abort': s!.ended = true; s!.answered = []; void s!.q!.end(); return
        case 'drop': { const d = DOCS[op.d]!; revokedAt.set(d, step); for (const x of sess) if (x.docId === d) x.answered = []; void cache.dropDocument(d); return }
        case 'sweep': t += MEDIA_IDLE_TTL_MS; return cache.sweep()
        case 'complete': { const open = calls.filter((c) => !c.done); if (open.length) open[op.i % open.length]!.settle(op.outcome); return }
        case 'together': { const x = exec(op.a), y = exec(op.b); return Promise.all([x, y]) }
      }
    }
    for (const op of ops) {
      step++
      await exec(op)
      { const q = inner.snapshot().queue; for (const r of pendingTickets) if (!placedBy.get(r) && q.includes(sidOf.get(r)!)) placedBy.set(r, true) }
      await flush(); await flush()
      watchCalls()
      const bad = violations[0] ?? check(`step ${step} (${op.t})`)
      if (bad) return bad
    }
    // Quiescence: complete every encode ok and deliver every answer until nothing is pending.
    for (let k = 0; k < 200; k++) {
      const open = calls.filter((c) => !c.done)
      for (const c of open) c.settle('ok')
      for (const s of sess) for (const a of s.answered.splice(0)) cache.delivered(s.sid, a.key)
      await flush(); await flush()
      watchCalls()
      const bad = violations[0] ?? check(`quiescence ${k}`)
      if (bad) return bad
      if (!calls.some((c) => !c.done) && sess.every((s) => s.pending.every((p) => p.settled))) break
    }
    // Liveness: a live, started session with room under both budgets has had every key of its deck encoded (by itself
    // or another session; the entry may since have expired): encode-ahead never stalls on a key.
    {
      const snap = cache.inspect()
      const held = snap.entries.reduce((n, e) => n + (e.refs || e.state === 'inflight' ? e.bytes : 0), 0)
      for (const s of sess) {
        if (s.ended || !s.started || revokedAt.has(s.docId)) continue
        if ((snap.sessions[s.sid]?.bytes ?? 0) + 32 * MiB > SESSION_BYTES || held + 32 * MiB > SERVER_BYTES) continue
        for (let u = 0; u < DECK.length; u++) for (const it of items(s.sid, s.docId, u)) {
          const id = encodeKeyId(it.key)
          if (!cache.peek(it.key) && !lastOutcome.has(id)) return `quiescence: ${s.sid}'s encode-ahead never encoded ${it.mediaId} (last: ${lastOutcome.get(id) ?? 'never tried'})`
        }
      }
    }
    for (const s of sess) for (const p of s.pending) if (!p.settled) {
      if (process.env.OV_FUZZ_DEBUG) console.log(JSON.stringify({ snap: cache.inspect(), running: admission.running, queued: admission.queued, qrun: sess.map((x) => x.q!.running), calls: calls.map((c) => [c.docId, c.done]) }, null, 1))
      return `quiescence: a request of ${s.sid} never settled`
    }
    for (const s of sess) if (!s.ended) { s.ended = true; await s.q!.end() }
    for (let k = 0; k < 5; k++) await flush()
    const snap = cache.inspect()
    if (Object.keys(snap.sessions).length) return `after every end: references remain (${Object.keys(snap.sessions).join(',')})`
    if (snap.entries.some((e) => e.state === 'inflight')) return 'after every end: an encode is still in flight'
    if (admission.running || admission.queued) return `after every end: admission running ${admission.running} queued ${admission.queued}`
    if (calls.some((c) => !c.done)) return 'after every end: a transcode is still running'
    return undefined
  } finally {
    for (const s of sess) if (!s.ended) { s.ended = true; void s.q!.end() }
    for (const c of calls) if (!c.done) c.settle('abort')
    await cache.close()
  }
}

/** Removes ops one at a time while the schedule still fails. */
async function shrink(ops: Op[]): Promise<Op[]> {
  let cur = ops
  for (let i = 0; i < cur.length;) {
    const cand = cur.slice(0, i).concat(cur.slice(i + 1))
    if (await runSchedule(cand)) cur = cand; else i++
  }
  return cur
}

describe('queue + cache + admission under random interleavings (fix round 2)', () => {
  it('every invariant holds after every step, across seeded schedules', async () => {
    const one = process.env.OV_FUZZ_SEED
    const seeds = one ? [Number(one)] : Array.from({ length: Number(process.env.OV_FUZZ_SCHEDULES ?? 3000) }, (_, i) => i + 1)
    const failures: string[] = []
    for (const seed of seeds) {
      const ops = genOps(seed, 40)
      const bad = await runSchedule(ops)
      if (!bad) continue
      const min = await shrink(ops)
      failures.push(`seed ${seed}: ${await runSchedule(min)}\n  trace: ${JSON.stringify(min)}`)
      if (failures.length >= 5) break
    }
    console.log(`[interleave] ${seeds.length} schedules · ${totals.encodes} encodes · answers ${JSON.stringify(Object.fromEntries(totals.answers))}`)
    if (failures.length) console.log(failures.join('\n'))
    expect(failures).toEqual([])
  }, 600_000)
})
