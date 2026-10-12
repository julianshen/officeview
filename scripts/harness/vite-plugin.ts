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

const ISOLATED_PATHS = ['/src/debug/viewer-perf.html', '/src/debug/viewerPerfWorker.ts']
export function isCrossOriginIsolatedPath(url: string): boolean {
  return ISOLATED_PATHS.includes(new URL(url, 'http://harness.invalid').pathname)
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
      // The S10.2 perf page and its worker are cross-origin isolated so Chrome's
      // performance.measureUserAgentSpecificMemory() is available there. Only
      // those two paths are affected; every other harness page is unchanged.
      server.middlewares.use((req, res, next) => {
        if (isCrossOriginIsolatedPath(req.url ?? '/')) {
          res.setHeader('Cross-Origin-Opener-Policy', 'same-origin')
          res.setHeader('Cross-Origin-Embedder-Policy', 'require-corp')
        }
        next()
      })
      if (options.collect) server.middlewares.use(createCollectorMiddleware(createCollector(options.collect)))
    },
  }
}
