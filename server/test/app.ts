/**
 * The in-repo Playwright WS harness (§16 T2): the 1c stream server in demo
 * mode over `corpus/`, plus TEST-ONLY controls on a separate port (never part
 * of `server/src` — the product server has no HTTP routes).
 *
 *   bun --preload ./scripts/stream-s0/napi-preload.ts server/test/app.ts
 *
 * Ports (OV_STREAM_PORT, default 4181):
 *   P      stream server (token ttl 300 s)
 *   P + 2  stream server with ttl 8 s / renew lead 5 s (the expiry spec)
 *   P + 1  controls: POST /drop|/revoke|/oversize?sid=&srv=short, GET /sessions,
 *          GET /fontpack/* (CORS), GET /__stream-harness (identity)
 */
import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { demoHooks } from '../src/demo-hooks'
import { deriveKeys } from '../src/keys'
import { Pool } from '../src/pool'
import { createStreamServer } from '../src/ws'

const P = Number(process.env.OV_STREAM_PORT ?? 4181)
const ORIGINS = [`http://127.0.0.1:${process.env.IV_HARNESS_PORT ?? 4173}`, `http://localhost:${process.env.IV_HARNESS_PORT ?? 4173}`]
const root = resolve(import.meta.dir, '../..')
const keys = await deriveKeys(crypto.getRandomValues(new Uint8Array(32)))
const pool = new Pool({ workers: 2 })
const hooks = demoHooks({ docDir: resolve(root, 'corpus'), demo: true })

const servers = {
  normal: createStreamServer({ keys, originAllowlist: ORIGINS, pool, ...hooks }),
  short: createStreamServer({ keys, originAllowlist: ORIGINS, pool, ...hooks, tokenTtlSec: 8, renewLeadSec: 5 }),
}
Bun.serve({ port: P, fetch: servers.normal.fetch as never, websocket: servers.normal.websocket as never })
Bun.serve({ port: P + 2, fetch: servers.short.fetch as never, websocket: servers.short.websocket as never })

type BunWS = { terminate?: () => void; close(code?: number): void; sendBinary(b: Uint8Array): void }
const cors = { 'access-control-allow-origin': '*' }
Bun.serve({
  port: P + 1,
  async fetch(req) {
    const url = new URL(req.url)
    const srv = servers[url.searchParams.get('srv') === 'short' ? 'short' : 'normal']
    const sid = url.searchParams.get('sid') ?? ''
    if (url.pathname === '/__stream-harness') return Response.json({ ok: true, root }, { headers: cors })
    if (url.pathname === '/sessions') return Response.json(srv.table.sids().map((s) => ({ sid: s, status: srv.table.get(s)!.status })), { headers: cors })
    if (url.pathname.startsWith('/fontpack/')) {
      const name = url.pathname.slice('/fontpack/'.length)
      if (!/^[\w.-]+\.(ttf|json)$/.test(name)) return new Response('no', { status: 404, headers: cors })
      try { return new Response(await readFile(resolve(root, 'fontpack', name)), { headers: cors }) } catch { return new Response('no', { status: 404, headers: cors }) }
    }
    if (req.method !== 'POST') return new Response('no', { status: 404, headers: cors })
    const ws = srv.table.get(sid)?.activeSocket as BunWS | undefined
    if (url.pathname === '/revoke') return Response.json({ ok: srv.revoke(sid) }, { headers: cors })
    if (!ws) return Response.json({ ok: false }, { status: 404, headers: cors })
    // An abnormal drop (no close handshake) — the client sees 1006.
    if (url.pathname === '/drop') { if (ws.terminate) ws.terminate(); else ws.close(1001); return Response.json({ ok: true }, { headers: cors }) }
    if (url.pathname === '/oversize') { ws.sendBinary(new Uint8Array((1 << 20) + 1)); return Response.json({ ok: true }, { headers: cors }) }
    return new Response('no', { status: 404, headers: cors })
  },
})
console.log(`stream harness: ws://127.0.0.1:${P}/stream, short-ttl ws://127.0.0.1:${P + 2}/stream, controls http://127.0.0.1:${P + 1}`)
