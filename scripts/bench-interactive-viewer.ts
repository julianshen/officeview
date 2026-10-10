/**
 * Interactive-viewer real-browser bench CLI (S0.3 scaffold).
 *
 *   bun scripts/bench-interactive-viewer.ts --browser chromium|webkit|safari-manual
 *     --profile desktop|mobile --fixture F1 [--repeat N] [--warmup W] [--output DIR]
 *     [--mode smoke] [--baseline FILE] [--serve] [--port 4173]
 *
 * chromium|webkit: drive the dev-only harness page through Playwright in one
 * fresh browser context, run W warmup + N measured input->present round-trips
 * (pooled into ONE result), validate the result against the schema, verify its
 * integrity hash, and write JSON + CSV artifacts to --output. WebKit here is
 * Playwright emulation and is labeled `webkit-playwright` — NEVER Safari or
 * physical-device evidence (spec §12.2).
 *
 * safari-manual + --serve: serve the same harness page on the LAN from the vite
 * harness server, print a tokenized URL for a real device, and ingest result
 * JSON POSTed back from the device on the SAME origin (see
 * scripts/harness/collector.ts). Device results keep the device's own browser
 * labeling; physical-device verification remains the operator's responsibility.
 *
 * This scaffold measures the harness frame only. It claims no viewer SLA. The
 * logic lives in scripts/harness/* (unit-tested); this file only orchestrates.
 */
import { chromium, webkit } from '@playwright/test'
import { randomBytes } from 'node:crypto'
import { mkdirSync, readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { verifyHarnessResult, validateHarnessResult } from '../src/debug/harness/result'
import type { HarnessResult } from '../src/debug/harness/types'
import { HELP_TEXT, parseArgs, type BenchArgs } from './harness/args'
import { writeArtifacts, type CollectedRun } from './harness/artifacts'
import { checkBaselineComparable, compareResults, parseBaseline } from './harness/baseline'
import { commitLabel, gitIdentity } from './harness/identity'
import { detectLanIp } from './harness/lan'
import { retagResult } from './harness/retag'
import { ensureHarnessServer } from './harness/vite-server'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const HARNESS_PAGE_PATH = '/src/debug/interactive-viewer.html'
const ROUND_TRIP_TIMEOUT_MS = 30_000

const VIEWPORTS = {
  desktop: { width: 1440, height: 900 },
  mobile: { width: 1024, height: 768 },
} as const

function withTimeout<T>(promise: Promise<T>, ms: number, what: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined
  const timeout = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error(`${what} timed out after ${ms} ms`)), ms)
  })
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer))
}

function pageQuery(args: BenchArgs, label: string, extra: Record<string, string> = {}): string {
  const params = new URLSearchParams({
    fixture: args.fixture,
    profile: args.profile,
    mode: args.mode,
    cache: 'cold',
    commit: label,
    ...extra,
  })
  const { deviceModel, osVersion, thermalState, batteryState } = args.operator
  if (deviceModel) params.set('deviceModel', deviceModel)
  if (osVersion) params.set('osVersion', osVersion)
  if (thermalState) params.set('thermal', thermalState)
  if (batteryState) params.set('battery', batteryState)
  return params.toString()
}

type HarnessWindow = {
  __viewerHarness?: {
    booted: boolean
    recordRoundTrip(label: string, warmup?: boolean): Promise<void>
    exportResult(): Promise<unknown>
  }
}

async function collectAutomated(args: BenchArgs, label: string): Promise<CollectedRun> {
  const engine = args.browser === 'chromium' ? chromium : webkit
  const browser = await engine.launch({ headless: true })
  const browserVersion = browser.version()
  try {
    const context = await browser.newContext({
      viewport: VIEWPORTS[args.profile],
      deviceScaleFactor: 2,
      isMobile: args.profile === 'mobile',
      hasTouch: args.profile === 'mobile',
    })
    const page = await context.newPage()
    page.setDefaultTimeout(ROUND_TRIP_TIMEOUT_MS)
    await page.goto(`http://127.0.0.1:${args.port}${HARNESS_PAGE_PATH}?${pageQuery(args, label)}`)
    await page.waitForFunction(() => (window as unknown as HarnessWindow).__viewerHarness?.booted === true, undefined, { timeout: 30_000 })

    // All round-trips in one evaluate: no driver latency between input and
    // present, and each round-trip has its own in-page timeout (harness.ts).
    const iterations = args.warmup + args.repeat
    await withTimeout(
      page.evaluate(
        async ({ warmup, repeat }) => {
          const harness = (window as unknown as HarnessWindow).__viewerHarness!
          for (let i = 0; i < warmup + repeat; i++) await harness.recordRoundTrip(`rt-${i}`, i < warmup)
        },
        { warmup: args.warmup, repeat: args.repeat },
      ),
      ROUND_TRIP_TIMEOUT_MS + iterations * 1000,
      'input-to-present round-trips',
    )
    const exported = (await withTimeout(
      page.evaluate(() => (window as unknown as HarnessWindow).__viewerHarness!.exportResult()),
      ROUND_TRIP_TIMEOUT_MS,
      'exportResult',
    )) as HarnessResult
    await context.close()

    const tagged = await retagResult(exported, {
      engine: args.browser === 'webkit' ? 'webkit' : 'chromium',
      browserVersion,
      hostPlatform: process.platform,
      commit: label,
      operator: args.operator,
    })
    const validation = validateHarnessResult(tagged)
    if (!validation.ok) throw new Error(`result failed schema validation: ${validation.errors.join('; ')}`)
    const verification = await verifyHarnessResult(tagged)
    if (!verification.ok) throw new Error(`result failed hash verification: ${verification.errors.join('; ')}`)
    console.log(
      `  collected 1 result: ${args.fixture} ${args.profile} ${tagged.environment.browser.name} ` +
        `${args.repeat} measured + ${args.warmup} warmup round-trips, ${tagged.samples.length} sample series, ${Object.keys(tagged.metrics).length} metrics, evidence ${tagged.environment.evidence}`,
    )
    return {
      browser: tagged.environment.browser.name,
      browserVersion,
      requestedFixtureId: args.fixture,
      profile: args.profile,
      mode: args.mode,
      commit: label,
      repeat: args.repeat,
      warmup: args.warmup,
      results: [tagged],
    }
  } finally {
    await browser.close()
  }
}

/** Serve the harness page for a real device and ingest its posted result. */
async function serveForManualDevice(args: BenchArgs, label: string, identity: ReturnType<typeof gitIdentity>): Promise<void> {
  const { ip, loopbackFallback } = detectLanIp()
  const token = randomBytes(16).toString('hex')
  const output = resolve(args.output)
  mkdirSync(output, { recursive: true })
  const server = await ensureHarnessServer({
    root: REPO_ROOT,
    port: args.port,
    host: '0.0.0.0',
    identity,
    allowReuse: false,
    env: { HARNESS_COLLECT_TOKEN: token, HARNESS_OUTPUT: output, HARNESS_PROFILE: args.profile, HARNESS_MODE: args.mode, HARNESS_FIXTURE: args.fixture },
  })
  try {
    const url = `http://${ip}:${args.port}${HARNESS_PAGE_PATH}?${pageQuery(args, label, { token })}`
    console.log(`Serving the harness page for manual device collection (same origin, one port):`)
    console.log(`  WARNING: the vite dev server is bound to ALL interfaces (0.0.0.0) and serves this repository's`)
    console.log(`           sources unauthenticated to anyone who can reach port ${args.port}. Only the result POST`)
    console.log(`           is protected (session token). Use a trusted network and stop the server when done.`)
    console.log(`  WARNING: the device URL travels over plain HTTP, so the session token is sniffable on this`)
    console.log(`           network. Anyone holding it can POST results until you stop the server.`)
    if (loopbackFallback) {
      console.log(`  WARNING: no LAN interface was found; the URL below uses 127.0.0.1 and NO OTHER DEVICE can reach it.`)
    }
    console.log(`  device URL (open it in Safari on the device; contains the session token):`)
    console.log(`    ${url}`)
    console.log(`  artifacts are written to: ${output}`)
    console.log(`  note: on this insecure http origin the page uses its pure-JS SHA-256 (crypto.subtle is unavailable).`)
    console.log(`On the device: load the page, press "Export result" — the page posts the result back to this server.`)
    console.log(`Results keep the device's own browser labeling; WebKit emulation is never relabeled Safari.`)
    console.log(`Press Ctrl-C to stop.`)
    await new Promise<void>((done) => {
      process.once('SIGINT', () => done())
      process.once('SIGTERM', () => done())
    })
  } finally {
    server.stop()
  }
}

async function main(): Promise<void> {
  const parsed = parseArgs(process.argv.slice(2))
  if ('help' in parsed) {
    console.log(HELP_TEXT)
    return
  }
  const args = parsed
  const identity = gitIdentity(REPO_ROOT)
  const label = commitLabel(identity)
  console.log(`interactive-viewer bench (S0.3 scaffold — measures the harness frame, claims no viewer SLA)`)
  console.log(`  browser=${args.browser} profile=${args.profile} fixture=${args.fixture} repeat=${args.repeat} warmup=${args.warmup} mode=${args.mode} port=${args.port} commit=${label}`)
  if (args.browser === 'safari-manual') {
    await serveForManualDevice(args, label, identity)
    return
  }

  // Validate the baseline BEFORE launching anything: a bad baseline must never
  // cost a run its results.
  let baseline: HarnessResult | null = null
  if (args.baseline) {
    const results = await parseBaseline(readFileSync(args.baseline, 'utf8'))
    baseline = results[0]
    const reasons = checkBaselineComparable(baseline, {
      browserLabel: args.browser === 'webkit' ? 'webkit-playwright' : 'chromium',
      profile: args.profile,
      requestedFixtureId: args.fixture,
      mode: args.mode,
    })
    if (reasons.length > 0) throw new Error(`baseline ${args.baseline} is not comparable:\n  ${reasons.join('\n  ')}`)
  }

  const server = await ensureHarnessServer({ root: REPO_ROOT, port: args.port, host: '127.0.0.1', identity, allowReuse: true })
  try {
    const run = await collectAutomated(args, label)
    if (baseline) {
      try {
        run.baselineComparison = compareResults(baseline, run.results[0])
      } catch (error) {
        run.baselineError = error instanceof Error ? error.message : String(error)
      }
    }
    const { json, csv } = writeArtifacts(run, resolve(args.output))
    console.log(`artifacts:`)
    console.log(`  JSON: ${json}`)
    console.log(`  CSV:  ${csv}`)
    if (run.baselineComparison) {
      console.log(`baseline comparison (p50 deltas, informational only):`)
      for (const warning of run.baselineComparison.warnings) console.log(`  warning: ${warning}`)
      for (const row of run.baselineComparison.rows) {
        if (row.delta === null) console.log(`  ${row.metric}: not comparable`)
        else console.log(`  ${row.metric}: ${row.delta >= 0 ? '+' : ''}${row.delta.toFixed(3)} ms vs baseline p50 ${row.baselineP50}`)
      }
    }
    if (run.baselineError) console.log(`baseline comparison failed (artifacts were still written): ${run.baselineError}`)
    console.log(`note: ${args.browser === 'webkit' ? 'webkit-playwright emulation — NOT Safari or physical-device evidence' : 'scaffold measurement — no viewer SLA is claimed'}`)
  } finally {
    server.stop()
  }
  console.log(`  vite dev server ${server.reused ? 'reused (identity verified)' : 'started and stopped'}`)
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : String(error))
  process.exit(1)
})
