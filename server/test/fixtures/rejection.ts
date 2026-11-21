/**
 * `await rejection(p)`: the reason `p` rejects with; a fulfilment fails the test. Use it, never `expect(p).rejects` (or
 * `.resolves`), anywhere under server/test (`no-expect-rejects.test.ts` enforces it).
 *
 * Bun 1.4.5 (measured; 1.3.10 delivered the Worker messages): while `expect(p).rejects/.resolves` waits on a pending
 * `p` and a Worker is alive anywhere in the process (an earlier test file's counts, e.g. harness `disposePool()`,
 * which does not wait for its workers to close), the process does not dispatch Worker `message` events or
 * `setImmediate` callbacks (timers, Worker `close`, subprocess exits, fetch, `MessageChannel` and `process.nextTick`
 * still run). So a `p` that needs a worker's reply never settles there (a pool test hangs to its timeout, or settles
 * some other way, e.g. a render deadline, and asserts the wrong outcome), and a `p` that is pure in-process JSZip work
 * (JSZip schedules its async steps with `setImmediate`) advances only when something else wakes the loop. Measured on
 * macOS: a 50-step `setImmediate` chain took 25 s; policy.test.ts's broken-XML projection took 2.4-3.9 s after
 * client-session or notes-evidence (its predecessors in the failing CI run, which timed it out at 5 s) against ~1 ms
 * alone. A plain `await` dispatches them.
 */
export async function rejection(p: Promise<unknown>): Promise<unknown> {
  let value: unknown
  try { value = await p } catch (reason) { return reason }
  throw new Error(`expected a rejection, got a fulfilment: ${String(value)}`)
}

/** The message of the `Error` `p` rejects with (`expect(p).rejects.toThrow(re)` becomes `expect(await rejectionMessage(p)).toMatch(re)`). */
export async function rejectionMessage(p: Promise<unknown>): Promise<string> {
  const reason = await rejection(p)
  if (!(reason instanceof Error)) throw new Error(`expected an Error rejection, got: ${String(reason)}`)
  return reason.message
}
