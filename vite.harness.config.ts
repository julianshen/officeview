/**
 * Vite config for the dev-only interactive-viewer harness server. It extends
 * the library config untouched (so the library build and `bun run dev` are
 * unaffected) and adds the harness identity marker plus, when the manual-device
 * CLI sets the HARNESS_* variables, the same-origin result collector.
 *
 *   bunx vite --config vite.harness.config.ts --port 4173 --strictPort --host 127.0.0.1
 */
import { mergeConfig, type Plugin, type ViteDevServer } from 'vite'
import { fileURLToPath } from 'node:url'
import { join } from 'node:path'
import baseConfig from './vite.config'
import { harnessPlugin } from './scripts/harness/vite-plugin'
import type { CollectorConfig } from './scripts/harness/collector'
import type { FixtureId, HarnessMode, HarnessProfile } from './src/debug/harness/types'

const env = process.env
const collect: CollectorConfig | null = env.HARNESS_COLLECT_TOKEN
  ? {
      token: env.HARNESS_COLLECT_TOKEN,
      outputDir: env.HARNESS_OUTPUT ?? '',
      expect: {
        profile: (env.HARNESS_PROFILE ?? 'desktop') as HarnessProfile,
        mode: (env.HARNESS_MODE ?? 'smoke') as HarnessMode,
        requestedFixtureId: (env.HARNESS_FIXTURE ?? 'F1') as FixtureId,
      },
    }
  : null
if (collect && !collect.outputDir) throw new Error('HARNESS_COLLECT_TOKEN requires HARNESS_OUTPUT')

/**
 * Lazy chunks the library reaches only through dynamic import(). A page
 * cannot import them statically without defeating the laziness the
 * paged-toolbar spec holds open, so once viewer-e2e.tsx (a page that renders
 * paged documents) is served, they are warmed into the module graph, which is
 * what lets the locked-down server serve them. Nothing is warmed for a bare
 * server, so the LAN lockdown is unchanged until such a page loads.
 */
function lazyPagedChunks(): Plugin {
  let server: ViteDevServer | undefined
  return {
    name: 'officeview-harness-lazy-paged-chunks',
    apply: 'serve',
    configureServer(s) { server = s },
    transform(_code, id) {
      if (!server || !id.split('?')[0].endsWith('/src/debug/viewer-e2e.tsx')) return
      for (const url of [
        '/src/components/paged/PagedChrome.tsx',
        '/src/components/zoomRouting.ts',
        '/src/components/paged/printDocument.ts',
      ]) void server.warmupRequest(url)
    },
  }
}

export default mergeConfig(baseConfig, {
  // The S8.4 kernel scenarios reuse the S0.2 fixture builders, whose module
  // imports node:crypto for digests the harness never computes.
  resolve: { alias: { 'node:crypto': fileURLToPath(new URL('./src/debug/harness/nodeCryptoShim.ts', import.meta.url)) } },
  plugins: [harnessPlugin({ root: fileURLToPath(new URL('.', import.meta.url)), collect }), lazyPagedChunks()],
  // LAN-mode lockdown: the harness page only needs its own directory (plus the
  // node_modules its react/jszip imports resolve to). Everything else in the
  // checkout — package.json, library sources, other configs — is denied, even
  // though the server binds all interfaces in manual mode. src/styles holds the
  // stylesheet the pages import (its Tailwind/scoping pipeline is inherited
  // from vite.config.ts through mergeConfig; never registered twice).
  server: {
    fs: {
      strict: true,
      allow: [join(fileURLToPath(new URL('.', import.meta.url)), 'src', 'debug'), join(fileURLToPath(new URL('.', import.meta.url)), 'src', 'styles'), join(fileURLToPath(new URL('.', import.meta.url)), 'node_modules')],
    },
  },
})
