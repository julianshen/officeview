/**
 * PR #78 round 4: a socket's show work is cancelled when the socket closes. Show builds and their Layer-1 encodes
 * queued for a session whose page has gone must not keep the workers busy: one viewer's backlog would delay every
 * other viewer's first paint on the same pool.
 */
import { afterAll, expect, test } from 'bun:test'
import { morphDeck } from '../../src/testdata/show-decks'
import { Pool } from '../src/pool'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import { Client, start } from './harness'

// One worker (every job queues behind the last); a generous show deadline, so a loaded machine is not a timeout.
const SHOW_WORK = new Set(['show-plan', 'show-track', 'show-morph', 'mark-layer'])
const pool = new Pool({ workers: 1, showDeadlineMs: 60_000 })
afterAll(() => pool.dispose())

/** Every worker request the pool posts, by type, and whether the client had closed by then. */
function spy(p: Pool) {
  const posts: Array<{ type: string; afterClose: boolean }> = []
  let closed = false
  const inner = p as unknown as { post: (slot: unknown, msg: { type: string }, ms?: number) => Promise<unknown> }
  const post = inner.post.bind(p)
  inner.post = (slot, msg, ms) => { posts.push({ type: msg.type, afterClose: closed }); return post(slot, msg, ms) }
  return { posts, close() { closed = true } }
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

test("a dropped socket's queued show builds and Layer-1 encodes are cancelled; nothing more is posted for it", async () => {
  const h = await start({ doc: await morphDeck(), pool, audit: new AuditTable({ store: AuditStore.open(':memory:') }), policy: { contentPolicy: 'full' } })
  const c = new Client(h.url)
  const m = await c.hello(); await c.ready()
  // Warm: the document is parsed on the worker and one pair is built.
  const warm = await c.showMorph(0, 1, 0, 1, 'full', { ms: 45_000 })
  expect(warm.end).toBeUndefined()
  expect(warm.start).toBeDefined()
  await idle(pool)
  const s = spy(pool)
  // Twelve pairs asked at once (the hold test's shape), then the page goes.
  for (let i = 0; i < 12; i++) c.send({ reqId: 2000 + i, token: m.session.token, type: 'show-morph', from: 0, to: 1, fromStep: 0, scale: 1, motion: 'full' })
  for (let i = 0; i < 200 && !s.posts.some((x) => x.type === 'show-morph'); i++) await Bun.sleep(5)
  expect(s.posts.some((x) => x.type === 'show-morph')).toBe(true)
  // A page that goes away drops its socket without a close frame (the server sees 1006, as for a browser's
  // 1001): the session is detached for a resume, so its document stays registered and the backlog could run.
  ;(c.ws as unknown as { terminate(): void }).terminate()
  await c.closed
  // From the moment the server has seen the drop (a job it had already handed to the worker meanwhile is work
  // already running, which finishes).
  await serverSaw(h)
  s.close()
  await idle(pool)
  // What was on a worker when the socket closed finishes; no show build or Layer-1 encode starts for the closed
  // socket (the document's own release, `drop`, is not show work).
  expect(s.posts.filter((x) => x.afterClose && SHOW_WORK.has(x.type)).map((x) => x.type)).toEqual([])
  // Non-vacuous: the twelve were really queued behind the first (one worker).
  expect(s.posts.filter((x) => !x.afterClose && x.type === 'show-morph').length).toBeLessThan(12)
  h.stop()
}, 180_000)
