/**
 * Playwright configuration for the interactive-viewer real-browser harness.
 *
 * The webServer is the vite harness server (vite.harness.config.ts), which
 * transforms and serves the dev-only harness page at
 * /src/debug/interactive-viewer.html and answers /__harness-id. The page is
 * excluded from the library build (tsconfig.build.json) — this config only
 * ever runs it in a real browser engine.
 *
 * Projects: chromium always; webkit only when its engine is actually
 * installed. WebKit here is Playwright emulation and results are labeled
 * `webkit-playwright` — never Safari or physical-device evidence (spec §12.2).
 */
import { defineConfig, devices } from '@playwright/test'
import { existsSync, readdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

// The identity marker, not the page: vite answers 200 + index.html for any path,
// so only /__harness-id (JSON, checked again by the spec) proves which server this is.
const HARNESS_ID_URL = 'http://127.0.0.1:4173/__harness-id'
function msPlaywrightDir(): string {
  const override = process.env.PLAYWRIGHT_BROWSERS_PATH
  if (override) return override
  return process.platform === 'darwin'
    ? join(process.env.HOME ?? tmpdir(), 'Library', 'Caches', 'ms-playwright')
    : join(process.env.HOME ?? tmpdir(), '.cache', 'ms-playwright')
}

function engineInstalled(prefix: string): boolean {
  try {
    return readdirSync(msPlaywrightDir()).some((entry) => entry.startsWith(prefix))
  } catch {
    return false
  }
}

const projects = [
  { name: 'chromium', use: { ...devices['Desktop Chrome'] } },
]
if (engineInstalled('webkit-')) {
  projects.push({ name: 'webkit', use: { ...devices['Desktop Safari'] } })
} else {
  console.warn('[playwright.viewer.config] WebKit is not installed (bunx playwright install webkit): running chromium only')
}

export default defineConfig({
  testDir: 'tests/browser',
  timeout: 120_000,
  expect: { timeout: 15_000 },
  fullyParallel: false,
  workers: 1,
  retries: 0,
  reporter: [['list']],
  use: {
    baseURL: 'http://127.0.0.1:4173',
    trace: 'retain-on-failure',
  },
  projects,
  webServer: {
    command: 'bunx vite --config vite.harness.config.ts --port 4173 --strictPort --host 127.0.0.1',
    url: HARNESS_ID_URL,
    // A reused server is accepted only if the spec's identity check passes
    // (same commit and checkout); a stale or foreign server fails that check.
    reuseExistingServer: !process.env.CI,
    timeout: 60_000,
  },
})
