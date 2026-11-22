/**
 * The test harness shares one render pool across the stream servers it starts (`pool()`); production wires one
 * server per pool (`main.ts`). A stopped harness server must not reach into the shared pool afterwards: `stop()`
 * drops its sockets (1001), so their sessions detach and are reaped after the grace, and the reaper's last
 * `releaseDoc` calls `pool.forget(docRef)`. Unscoped, that forget removed the bytes a LATER server had registered
 * for the same document, and that server's next render failed "document not registered" (refused policy-refused).
 * In CI this hit print-wire tests about 30 s (the default grace) after the file's first sessions.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import { Client, corpus, disposePool, start } from './harness'

afterAll(async () => { await disposePool() })

describe('harness isolation: a stopped server leaves the shared pool alone', () => {
  test("a stopped server's detached session, reaped later, does not forget a later server's document", async () => {
    // Server A: a session on the default deck, then stop() while it streams (the socket drops; the session detaches).
    const opts = () => ({ doc: corpus('lo-tdf104788.pptx'), audit: new AuditTable({ store: AuditStore.open(':memory:') }) })
    const a = await start({ ...opts(), detachGraceMs: 200 }), ca = new Client(a.url)
    await ca.hello(); await ca.ready()
    await a.stop()
    // Server B, on the same shared pool and the same document (same docRef), opens it before A's reaper runs.
    const b = await start(opts()), cb = new Client(b.url)
    await cb.hello(); await cb.ready()
    await Bun.sleep(600)
    // A's grace has passed: its reaper tore its session down. B's document is still registered.
    expect((await cb.print(1, 0, 0)).end).toEqual({ reqId: 1, type: 'print-done', pages: 1 })
    cb.close(); await b.stop()
  }, 30_000)
})
