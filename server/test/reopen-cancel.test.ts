/**
 * PR #78 round 5 (codex round-4 review): the requesting socket's signal is checked again after every await before
 * a job is posted. A document dropped from the worker (`dropParsed`, a recycle or a replacement) is reopened inside
 * the job's worker turn; a socket that closes during that reopen must not still get its build. For show-plan,
 * unit render, notes and text: abort across the reopen, then no job of that kind is posted, the request settles
 * as cancelled, every reservation is returned, and a fresh request still runs.
 */
import { afterEach, expect, test } from 'bun:test'
import { morphDeck } from '../../src/testdata/show-decks'
import { Pool } from '../src/pool'

let pool: Pool | undefined
afterEach(() => { pool?.dispose(); pool = undefined })

const DOC = 'reopen'
const mark = { kind: 'off' } as const
type Kind = 'show-plan' | 'render' | 'notes' | 'text'
const run: Record<Kind, (p: Pool, signal?: AbortSignal) => Promise<unknown>> = {
  'show-plan': (p, signal) => p.showPlan('sid', { docRef: DOC, contentPolicy: 'full', unitIndex: 1, scale: 1, motion: 'full', mark, strictMarks: false }, 1024, signal),
  render: (p, signal) => p.render('sid', { docRef: DOC, contentPolicy: 'full', unitIndex: 0, scale: 1, projection: 'final-raster', mark }, 1024, signal),
  notes: (p, signal) => p.notes('sid', { docRef: DOC, contentPolicy: 'full', unitIndex: 0, scale: 1, mark, draw: false }, 1024, signal),
  text: (p, signal) => p.text('sid', { docRef: DOC, contentPolicy: 'full', op: 'a11y', unit: 0 }, undefined, signal),
}

for (const kind of Object.keys(run) as Kind[]) {
  test(`${kind}: a socket closed while its document reopens gets no ${kind} job; the reservation returns; a fresh request runs`, async () => {
    const p = pool = new Pool({ workers: 1, showDeadlineMs: 60_000 })
    await p.open(DOC, await morphDeck(), 'full')
    p.dropParsed(DOC, 'full')
    const ac = new AbortController()
    const posts: Array<{ type: string; aborted: boolean }> = []
    const inner = p as unknown as { post: (slot: unknown, msg: { type: string }, ms?: number) => Promise<unknown> }
    const post = inner.post.bind(p)
    let reached!: () => void, release!: () => void
    const opened = new Promise<void>((r) => { reached = r }), gate = new Promise<void>((r) => { release = r })
    // The reopen's reply is held at its await: the socket closes then.
    inner.post = async (slot, msg, ms) => {
      posts.push({ type: msg.type, aborted: ac.signal.aborted })
      const reply = await post(slot, msg, ms)
      if (msg.type === 'open') { reached(); await gate }
      return reply
    }
    const pending = run[kind](p, ac.signal)
    await opened
    ac.abort()
    release()
    const outcome = await pending.then(() => 'ran', (e: Error) => e.name)
    expect({ outcome, posts: posts.map((x) => x.type) }).toEqual({ outcome: 'ShowCancelled', posts: ['open'] })
    expect(p.stats()).toMatchObject({ reserved: 0, running: 0, queued: 0 })
    inner.post = post
    // A fresh request (another socket) still runs.
    expect(await run[kind](p, new AbortController().signal).then(() => 'ran', (e: Error) => e.message)).toBe('ran')
    expect(p.stats()).toMatchObject({ reserved: 0, running: 0, queued: 0 })
  }, 120_000)
}
