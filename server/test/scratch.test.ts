/** Final-review fixes: §14 scratch lease handoff is atomic, close wins, waiters are bounded. */
import { expect, it } from 'bun:test'
import { LeaseClosed, LeaseFull, ScratchLease } from '../src/scratch'

it('a release hands its slot to the head waiter: a racing acquire cannot exceed the limit', async () => {
  const l = new ScratchLease(2)
  const a = await l.acquire(), b = await l.acquire()
  const c = l.acquire()
  a()
  const d = l.acquire()
  const [rc] = await Promise.all([c])
  expect(l.inFlight).toBe(2)
  rc(); b()
  const rd = await d
  expect(l.inFlight).toBe(1)
  rd()
  expect(l.inFlight).toBe(0)
})

it('release then close before the woken waiter runs: the waiter rejects and the slot is undone', async () => {
  const l = new ScratchLease(1)
  const a = await l.acquire()
  const w = l.acquire()
  const settled = Promise.allSettled([w])
  a(); l.close()
  const [r] = await settled
  expect(r.status === 'rejected' && r.reason instanceof LeaseClosed).toBe(true)
  expect(l.inFlight).toBe(0)
})

it('waiters beyond maxWaiters are refused with LeaseFull', async () => {
  const l = new ScratchLease(1, 2)
  await l.acquire()
  const w1 = l.acquire(), w2 = l.acquire()
  await expect(l.acquire()).rejects.toBeInstanceOf(LeaseFull)
  const settled = Promise.allSettled([w1, w2])
  l.close()
  expect((await settled).map((r) => r.status === 'rejected' && r.reason instanceof LeaseClosed)).toEqual([true, true])
})
