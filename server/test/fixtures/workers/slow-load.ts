/** A worker whose load takes a while (a top-level await stands in for a large module graph), then echoes. */
declare const self: Worker
await Bun.sleep(250)
self.onmessage = (e: MessageEvent) => { postMessage(e.data) }
export {}
