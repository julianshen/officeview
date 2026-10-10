/**
 * Vite plugin for the dev-only harness server: the identity marker (so a
 * runner can tell its own server from a stale or foreign one) and, when the
 * manual-device mode configured it, the same-origin result collector.
 */
import type { Plugin } from 'vite'
import { createCollector, createCollectorMiddleware, type CollectorConfig } from './collector'
import { HARNESS_ID_PATH, gitIdentity } from './identity'

export interface HarnessPluginOptions {
  root: string
  /** null: identity only (automated runs, Playwright's webServer). */
  collect: CollectorConfig | null
}

export function harnessPlugin(options: HarnessPluginOptions): Plugin {
  return {
    name: 'viewer-harness',
    apply: 'serve',
    configureServer(server) {
      const identity = gitIdentity(options.root)
      // Registered directly (not returned from the hook) so it runs before vite's
      // SPA fallback, which would answer 200 + index.html for this path.
      server.middlewares.use((req, res, next) => {
        if (new URL(req.url ?? '/', 'http://harness.invalid').pathname !== HARNESS_ID_PATH) {
          next()
          return
        }
        res.writeHead(200, { 'content-type': 'application/json', 'cache-control': 'no-store' })
        res.end(JSON.stringify(identity))
      })
      if (options.collect) server.middlewares.use(createCollectorMiddleware(createCollector(options.collect)))
    },
  }
}
