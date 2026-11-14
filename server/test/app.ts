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
 *   ?copy=full|cap3|off on the stream URL sets the session's copy policy.
 *   ?a11y=unit|focused|off sets copy so that a11y resolves to that mode (R7):
 *          unit = copy 'full'; focused = copy capped 3 / session 5; off = copy and a11y off.
 *   ?print=on|off|cap2|session3 sets the session's print policy (Stage 3b):
 *          on = true; off = false; cap2 = {maxPages: 2}; session3 = {maxPages: 5, maxSessionPages: 3}.
 *   ?role=presenter makes the viewer a presenter (Stage 3c); absent = the demo's role (audience).
 *   ?notes=off sets the session's notes policy to false (Stage 3c); absent = the demo default (on).
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
const demo = demoHooks({ docDir: resolve(root, 'corpus'), demo: true })
/** TEST-ONLY `?copy=full|cap3|off` (Stage 3a): the session's copy policy; absent = the demo default. */
const COPY = { full: 'full', cap3: { maxParagraphs: 3 }, off: false } as const
/** TEST-ONLY `?a11y=unit|focused|off` (Stage 3a Task 8); applied after `?copy=`. */
const A11Y = { unit: { copy: 'full' }, focused: { copy: { maxParagraphs: 3, maxSessionParagraphs: 5 } }, off: { copy: false, textA11y: false } } as const
/** TEST-ONLY `?print=on|off|cap2|session3` (Stage 3b Task 8): the session's print policy; absent = the demo default. */
const PRINT = { on: true, off: false, cap2: { maxPages: 2 }, session3: { maxPages: 5, maxSessionPages: 3 } } as const
/** TEST-ONLY `?role=presenter` and `?notes=off` (Stage 3c Task 7). */
const ROLE = { presenter: 'presenter', audience: 'audience' } as const
const NOTES = { off: false } as const
const pick = <T extends object>(table: T, v: string | null): T[keyof T] | undefined => (v !== null && v in table ? table[v as keyof T] : undefined)
const hooks = {
  ...demo,
  authorize: async (req: Parameters<typeof demo.authorize>[0]) => {
    const d = await demo.authorize(req)
    if (!d) return d
    const copy = pick(COPY, req.url.searchParams.get('copy')), a11y = pick(A11Y, req.url.searchParams.get('a11y'))
    const print = pick(PRINT, req.url.searchParams.get('print'))
    const role = pick(ROLE, req.url.searchParams.get('role')), notes = pick(NOTES, req.url.searchParams.get('notes'))
    return {
      ...d, ...(role !== undefined ? { role } : {}),
      policy: { ...d.policy, ...(copy !== undefined ? { copy } : {}), ...a11y, ...(print !== undefined ? { print } : {}), ...(notes !== undefined ? { notes } : {}) },
    }
  },
}

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
