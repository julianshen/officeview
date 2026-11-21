/**
 * `await rejection(p)`: the reason `p` rejects with; a fulfilment fails the test. Use it, not `expect(p).rejects` (or
 * `.resolves`), for a promise that settles on a real Worker's message.
 *
 * Bun 1.4.5 (measured; 1.3.10 does not): while `expect(p).rejects/.resolves` waits on a pending `p`, the process does
 * not dispatch Worker `message` events (timers, Worker `close`, subprocess exits and fetch still run), so a `p` that
 * needs a worker's reply never settles there: a pool test hangs to its timeout, or settles some other way (a render
 * deadline) and asserts the wrong outcome. A plain `await` dispatches them.
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
