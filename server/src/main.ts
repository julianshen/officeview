/**
 * Demo stream server (Stage 1c). WS-only: `ws://host:port/stream?doc=<name>`.
 *
 *   bun run stream:demo                       # --demo, corpus/, ephemeral key
 *   OV_STREAM_KEY=<base64 32 B> OV_AUDIT_DB=/var/lib/ov/audit.db bun --preload ./scripts/stream-s0/napi-preload.ts \
 *     server/src/main.ts --doc-dir ./docs-to-serve --origin https://app.example
 */
import { AuditTable } from './audit'
import { AuditStore } from './audit-store'
import { loadConfig } from './config'
import { demoHooks } from './demo-hooks'
import { deriveKeys } from './keys'
import { Pool } from './pool'
import { createStreamServer } from './ws'

export async function startServer(argv: readonly string[], env: Record<string, string | undefined> = process.env) {
  const config = loadConfig(env, argv)
  if (config.ephemeral) console.warn('EPHEMERAL KEY — sessions die with this process')
  const origins = config.originAllowlist.length ? config.originAllowlist : config.demo ? ['http://localhost:5173'] : []
  if (!origins.length) throw new Error('no allowed origin: pass --origin or OV_STREAM_ORIGINS')
  const keys = await deriveKeys(config.keyring.keys.get(config.keyring.current)!, config.keyring.current)
  const store = AuditStore.open(config.auditDb), audit = new AuditTable({ store })
  // §9.3 retention: expire at startup, then hourly.
  const expire = () => audit.expire(Math.floor(Date.now() / 1000))
  expire()
  const expiry = setInterval(() => { try { expire() } catch (e) { console.error('[audit] expiry failed', e) } }, 3600_000); expiry.unref()
  const pool = new Pool()
  const stream = createStreamServer({ keys, originAllowlist: origins, pool, audit, ...demoHooks(config) })
  const server = Bun.serve({ port: config.port, fetch: stream.fetch as never, websocket: stream.websocket as never })
  return { server, stream, pool, origins, stop() { server.stop(true); pool.dispose(); clearInterval(expiry); store.close() } }
}

if (import.meta.main) {
  const s = await startServer(process.argv.slice(2))
  console.log(`[stream] ws://localhost:${s.server.port}/stream?doc=<file> · origins ${s.origins.join(', ')}`)
}
