/**
 * Vite config for the dev-only interactive-viewer harness server. It extends
 * the library config untouched (so the library build and `bun run dev` are
 * unaffected) and adds the harness identity marker plus, when the manual-device
 * CLI sets the HARNESS_* variables, the same-origin result collector.
 *
 *   bunx vite --config vite.harness.config.ts --port 4173 --strictPort --host 127.0.0.1
 */
import { mergeConfig } from 'vite'
import { fileURLToPath } from 'node:url'
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

export default mergeConfig(baseConfig, {
  plugins: [harnessPlugin({ root: fileURLToPath(new URL('.', import.meta.url)), collect })],
})
