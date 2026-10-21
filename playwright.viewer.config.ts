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
 *
 * `slideshow` (slideshow plan S1-2): Chromium with SwiftShader WebGL2, running
 * only tests/browser/slideshow-*.spec.ts. The viewer projects ignore those
 * specs, so their own test set is unchanged. CI runs `--project slideshow`.
 */
import { defineConfig, devices, type Project } from '@playwright/test'
import { existsSync, readdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

// The identity marker, not the page: vite answers 200 + index.html for any path,
// so only /__harness-id (JSON, checked again by the spec) proves which server this is.
// IV_HARNESS_PORT moves the server off 4173 when another checkout holds it
// (interactive-viewer.spec.ts reads the same variable for its identity check).
const PORT = Number(process.env.IV_HARNESS_PORT ?? 4173)
const HARNESS_ID_URL = `http://127.0.0.1:${PORT}/__harness-id`
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

// `--project slideshow` alone (the CI job) is Chromium-only by design, so the
// missing-WebKit notice would be noise there.
function slideshowOnly(): boolean {
  const selected: string[] = []
  const argv = process.argv
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--project' && argv[i + 1]) selected.push(argv[++i])
    else if (argv[i].startsWith('--project=')) selected.push(argv[i].slice('--project='.length))
  }
  const only = selected.length > 0 && selected.every((name) => name === 'slideshow')
  // Workers re-load this file with their own argv; the env carries the decision to them.
  if (only) process.env.OV_PW_SLIDESHOW_ONLY = '1'
  return only
}

const SLIDESHOW_SPECS = /slideshow-.*\.spec\.ts/
const projects: Project[] = [
  { name: 'chromium', testIgnore: SLIDESHOW_SPECS, use: { ...devices['Desktop Chrome'] } },
]
if (engineInstalled('webkit-')) {
  projects.push({ name: 'webkit', testIgnore: SLIDESHOW_SPECS, use: { ...devices['Desktop Safari'] } })
} else if (process.env.OV_PW_SLIDESHOW_ONLY !== '1' && !slideshowOnly()) {
  console.warn('[playwright.viewer.config] WebKit is not installed (bunx playwright install webkit): running chromium only')
}
projects.push({
  name: 'slideshow',
  testMatch: SLIDESHOW_SPECS,
  use: {
    ...devices['Desktop Chrome'],
    launchOptions: { args: ['--use-gl=angle', '--use-angle=swiftshader', '--enable-unsafe-swiftshader'] },
  },
})

export default defineConfig({
  testDir: 'tests/browser',
  timeout: 120_000,
  expect: { timeout: 15_000 },
  fullyParallel: false,
  workers: 1,
  retries: 0,
  reporter: [['list']],
  use: {
    baseURL: `http://127.0.0.1:${PORT}`,
    trace: 'retain-on-failure',
  },
  projects,
  webServer: {
    command: `bunx vite --config vite.harness.config.ts --port ${PORT} --strictPort --host 127.0.0.1`,
    url: HARNESS_ID_URL,
    // A reused server is accepted only if the spec's identity check passes
    // (same commit and checkout); a stale or foreign server fails that check.
    reuseExistingServer: !process.env.CI,
    timeout: 60_000,
  },
})
