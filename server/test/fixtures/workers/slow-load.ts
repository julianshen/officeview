/**
 * A worker whose load takes a while (a 250 ms load stands in for a large module graph), then echoes.
 *
 * The handler is installed at once and every message waits for the load: Bun 1.4.5 drops a message that arrives while
 * a worker module is still in a top-level `await` (before its `onmessage` exists), where 1.3.10 queued it. No message
 * is answered before the load ends, so the first answer is still the load signal the start gate waits for.
 */
declare const self: Worker
const load = Bun.sleep(250)
self.onmessage = async (e: MessageEvent) => { await load; postMessage(e.data) }
export {}
