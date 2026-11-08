/**
 * Demo stream server (Stage 1c). WS-only: `ws://host:port/stream?doc=<name>`.
 *
 *   bun run stream:demo                       # --demo, corpus/, ephemeral key
 *   OV_STREAM_KEY=<base64 32 B> bun --preload ./scripts/stream-s0/napi-preload.ts \
 *     server/src/main.ts --doc-dir ./docs-to-serve --origin https://app.example
 */
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
  const pool = new Pool()
  const stream = createStreamServer({ keys: await deriveKeys(config.masterKey), originAllowlist: origins, pool, ...demoHooks(config) })
  const server = Bun.serve({ port: config.port, fetch: stream.fetch as never, websocket: stream.websocket as never })
  return { server, stream, pool, origins, stop() { server.stop(true); pool.dispose() } }
}

if (import.meta.main) {
  const s = await startServer(process.argv.slice(2))
  console.log(`[stream] ws://localhost:${s.server.port}/stream?doc=<file> · origins ${s.origins.join(', ')}`)
}
