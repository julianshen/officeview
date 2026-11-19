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
import { createMediaServices, MEDIA_ENABLED_DEFAULT } from './media/services'
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
  // Piece 5: the ffmpeg startup probe and the process's one media temp dir, admission class and encode cache. Fails
  // closed (no `info`: media unavailable). Off (D18, until Task 11) it probes nothing and touches no disk, and the
  // pool gets no media config, so its workers never probe.
  const media = await createMediaServices({ env: config.media, enabled: MEDIA_ENABLED_DEFAULT })
  const pool = new Pool(media.info ? { media: { info: media.info, dir: media.temp.dir } } : {})
  const stream = createStreamServer({ keys, originAllowlist: origins, pool, audit, animation: config.animation, media: { enabled: MEDIA_ENABLED_DEFAULT, ...(media.info ? { info: media.info } : {}), services: media }, ...demoHooks(config) })
  const server = Bun.serve({ port: config.port, fetch: stream.fetch as never, websocket: stream.websocket as never })
  return {
    server, stream, pool, origins, media,
    /** Stops serving at once; the returned promise settles once the media queues, cache and temp dir are gone. */
    stop(): Promise<void> { server.stop(true); pool.dispose(); clearInterval(expiry); store.close(); return media.close() },
  }
}

if (import.meta.main) {
  const s = await startServer(process.argv.slice(2))
  console.log(`[stream] ws://localhost:${s.server.port}/stream?doc=<file> · origins ${s.origins.join(', ')}`)
}
