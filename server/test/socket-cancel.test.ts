/**
 * PR #78 round 5 (after round 4's show-cancel.test.ts): every per-socket pool job is cancelled when its socket
 * drops — unit renders (and their Layer-1 encodes), notes images and text jobs. Queued for a page that has gone,
 * they would keep the shared workers busy and delay every other viewer; their replies can reach no one.
 *
 * Each test warms the document on a one-worker pool, asks for a dozen jobs of one kind at once, terminates the
 * socket (the server sees 1006, as for a browser's 1001: the session is detached for a resume, so the document
 * stays registered), then checks that no job of that kind was posted to a worker after the drop.
 */
import { afterAll, expect, test } from 'bun:test'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import type { AuthDecision } from '../src/hooks'
import { Pool } from '../src/pool'
import { Client, corpus, start } from './harness'

// One worker (every job queues behind the last); generous deadlines, so a loaded machine is not a timeout.
const pool = new Pool({ workers: 1, showDeadlineMs: 60_000 })
afterAll(() => pool.dispose())

/** Six slides with notes parts (units 4 and 5 carry notes text). */
const DECK = 'lo-tdf104788.pptx'

/** Every worker request the pool posts, by type, and whether the socket had dropped by then. */
function spy(p: Pool) {
  const posts: Array<{ type: string; afterClose: boolean }> = []
  let closed = false
  const inner = p as unknown as { post: (slot: unknown, msg: { type: string }, ms?: number) => Promise<unknown> }
  const post = inner.post.bind(p)
  inner.post = (slot, msg, ms) => { posts.push({ type: msg.type, afterClose: closed }); return post(slot, msg, ms) }
  return { posts, close() { closed = true }, restore() { inner.post = post } }
}

/** Until the server has processed the socket's close (its socket table is empty). */
const serverSaw = async (h: { stream: { stats(): { sockets: number } } }) => {
  for (let i = 0; i < 2000 && h.stream.stats().sockets > 0; i++) await Bun.sleep(5)
  expect(h.stream.stats().sockets).toBe(0)
}

const idle = async (p: Pool) => {
  for (let i = 0; i < 4800; i++) {
    const st = p.stats()
    if (st.running === 0 && st.queued === 0 && st.workers.every((w) => !w.busy)) return
    await Bun.sleep(25)
  }
  throw new Error('pool never went idle')
}

/**
 * Opens a session, warms it with `warm`, sends `burst` (a dozen jobs of one kind), drops the socket once the first
 * of them reaches the worker, and returns what was posted before and after the drop.
 */
async function dropped(o: { decision?: AuthDecision; warm: (c: Client) => Promise<void>; burst: (c: Client, token: string) => void; kinds: readonly string[]; started?: readonly string[] }) {
  const store = AuditStore.open(':memory:')
  const decision = o.decision ?? { viewerId: 'alice', policy: {} }
  const h = await start({ doc: corpus(DECK), pool, audit: new AuditTable({ store }), authorize: async () => decision })
  const c = new Client(h.url)
  const m = await c.hello(); await c.ready()
  await o.warm(c)
  await idle(pool)
  const s = spy(pool)
  try {
    o.burst(c, m.session.token)
    // The drop comes once the burst's first job is on the worker (`started`: a slower job sent ahead of it).
    const first = o.started ?? o.kinds
    for (let i = 0; i < 2000 && !s.posts.some((x) => first.includes(x.type)); i++) await Bun.sleep(5)
    expect(s.posts.some((x) => first.includes(x.type))).toBe(true)
    ;(c.ws as unknown as { terminate(): void }).terminate()
    await c.closed
    await serverSaw(h)
    s.close()
    await idle(pool)
    return {
      before: s.posts.filter((x) => !x.afterClose && o.kinds.includes(x.type)).length,
      after: s.posts.filter((x) => x.afterClose && o.kinds.includes(x.type)).map((x) => x.type),
    }
  } finally {
    s.restore()
    h.stop()
  }
}

test("a dropped socket's queued unit renders (and their Layer-1 encodes) are cancelled", async () => {
  const r = await dropped({
    // Warm: the document parsed on the worker (one unit rendered and delivered).
    warm: async (c) => { c.unit(10, 0, 1); await c.unitMessages(10) },
    // Twelve distinct renders (slides 1..5 at scale 1 and 2, slide 0 at scale 2, then slide 1 at 1 again): none is
    // in the render cache.
    burst: (c, token) => { let reqId = 100; for (const [u, sc] of [[1, 1], [2, 1], [3, 1], [4, 1], [5, 1], [0, 2], [1, 2], [2, 2], [3, 2], [4, 2], [5, 2]] as const) c.send({ reqId: reqId++, type: 'unit', token, unitIndex: u, scale: sc }) },
    kinds: ['render', 'encode'],
  })
  expect(r.after).toEqual([])
  expect(r.before).toBeLessThan(11)
}, 180_000)

test("a dropped socket's queued notes images are cancelled", async () => {
  const r = await dropped({
    decision: { viewerId: 'alice', role: 'presenter', policy: {} },
    warm: async (c) => { expect((await c.notes(0, 1, { ms: 45_000 })).end).toBeUndefined() },
    burst: (c, token) => { let reqId = 200; for (const sc of [1, 2] as const) for (let u = 0; u < 6; u++) if (u || sc !== 1) c.send({ reqId: reqId++, type: 'notes', token, unit: u, scale: sc }) },
    kinds: ['notes', 'encode'],
  })
  expect(r.after).toEqual([])
  expect(r.before).toBeLessThan(11)
}, 180_000)

test("a dropped socket's queued text jobs are cancelled", async () => {
  const r = await dropped({
    decision: { viewerId: 'alice', policy: { textA11y: true } },
    warm: async (c) => { await c.text({ reqId: 300, type: 'a11y', unit: 0 }, { ms: 45_000 }) },
    // A text job is quick: a scale-2 render goes first and holds the one worker, then twelve a11y reads (every
    // slide, twice) queue behind it; the socket drops while the render runs.
    burst: (c, token) => { c.send({ reqId: 309, type: 'unit', token, unitIndex: 5, scale: 2 }); let reqId = 310; for (let k = 0; k < 2; k++) for (let u = 0; u < 6; u++) c.send({ reqId: reqId++, type: 'a11y', token, unit: u }) },
    kinds: ['text'],
    started: ['render'],
  })
  expect(r.after).toEqual([])
  expect(r.before).toBeLessThan(12)
}, 180_000)
