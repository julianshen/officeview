/**
 * Interactive-viewer real-browser bench CLI (S0.3 scaffold).
 *
 *   bun scripts/bench-interactive-viewer.ts --browser chromium|webkit|safari-manual
 *     --profile desktop|mobile --fixture F1 [--repeat N] [--output DIR]
 *     [--mode smoke|release] [--baseline FILE] [--serve] [--port 4173]
 *
 * chromium|webkit: drive the dev-only harness page through Playwright in a real
 * browser engine, collect `--repeat` result JSON documents, validate each
 * against the result schema, verify its integrity hash, and write JSON + CSV
 * artifacts to --output. WebKit here is Playwright emulation and is labeled
 * `webkit-playwright` — NEVER Safari or physical-device evidence (spec §12.2).
 *
 * safari-manual + --serve: serve the same harness page over HTTP on the
 * LAN/local interface, print the URL for a real device, and ingest result JSON
 * POSTed back from the device (schema-validated and hash-verified). Device
 * results keep the device's own browser labeling; physical-device verification
 * remains the operator's responsibility.
 *
 * This scaffold measures the harness frame only. It claims no viewer SLA.
 */
import { chromium, webkit } from '@playwright/test'
import {
  hashResult,
  validateHarnessResult,
  verifyHarnessResult,
  type HarnessResult,
} from '../src/debug/harness/result'
import type { FixtureId, HarnessEnvironment } from '../src/debug/harness/types'
import { existsSync, mkdirSync, readdirSync, writeFileSync } from 'node:fs'
import { execSync } from 'node:child_process'
import { networkInterfaces } from 'node:os'
import { join } from 'node:path'

const REPO_ROOT = join(import.meta.dir, '..')
const HARNESS_PAGE_PATH = '/src/debug/interactive-viewer.html'
const FIXTURE_IDS: readonly FixtureId[] = ['F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8']

const VIEWPORTS = {
  desktop: { width: 1440, height: 900 },
  mobile: { width: 1024, height: 768 },
} as const

interface BenchArgs {
  browser: 'chromium' | 'webkit' | 'safari-manual'
  profile: 'desktop' | 'mobile'
  fixture: FixtureId
  repeat: number
  output: string
  mode: 'smoke' | 'release'
  baseline: string | null
  serve: boolean
  port: number
  warmup: number
}

function parseArgs(argv: string[]): BenchArgs {
  const args: BenchArgs = {
    browser: 'chromium',
    profile: 'desktop',
    fixture: 'F1',
    repeat: 1,
    output: '.tmp/viewer-bench',
    mode: 'smoke',
    baseline: null,
    serve: false,
    port: 4173,
    warmup: 0,
  }
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]
    const match = /^--([^=]+)(?:=(.*))?$/.exec(arg)
    if (!match) throw new Error(`unexpected argument ${arg}; see --help`)
    const [, name, inlineValue] = match
    let value: string | undefined = inlineValue
    if (value === undefined && !(name === 'serve' || name === 'help' || name === 'h')) {
      value = argv[++i]
    }
    if (value === undefined && name !== 'serve') throw new Error(`--${name} requires a value`)
    switch (name) {
      case 'browser':
        if (value !== 'chromium' && value !== 'webkit' && value !== 'safari-manual') {
          throw new Error(`--browser must be chromium|webkit|safari-manual, got ${value}`)
        }
        args.browser = value
        break
      case 'profile':
        if (value !== 'desktop' && value !== 'mobile') throw new Error(`--profile must be desktop|mobile, got ${value}`)
        args.profile = value
        break
      case 'fixture':
        if (!FIXTURE_IDS.includes(value as FixtureId)) throw new Error(`--fixture must be F1..F8, got ${value}`)
        args.fixture = value as FixtureId
        break
      case 'repeat': {
        const repeat = Number(value)
        if (!Number.isInteger(repeat) || repeat < 1) throw new Error(`--repeat must be a positive integer, got ${value}`)
        args.repeat = repeat
        break
      }
      case 'output':
        args.output = value
        break
      case 'mode':
        if (value !== 'smoke' && value !== 'release') throw new Error(`--mode must be smoke|release, got ${value}`)
        args.mode = value
        break
      case 'baseline':
        args.baseline = value
        break
      case 'serve':
        args.serve = value === undefined || value === 'true' || value === ''
        break
      case 'port': {
        const port = Number(value)
        if (!Number.isInteger(port) || port < 1 || port > 65535) throw new Error(`--port must be 1..65535, got ${value}`)
        args.port = port
        break
      }
      case 'warmup': {
        const warmup = Number(value)
        if (!Number.isInteger(warmup) || warmup < 0) throw new Error(`--warmup must be a non-negative integer, got ${value}`)
        args.warmup = warmup
        break
      }
      case 'help':
      case 'h':
        printHelp()
        process.exit(0)
      default:
        throw new Error(`unknown flag --${name}; see --help`)
    }
  }
  if (args.warmup >= args.repeat) throw new Error(`--warmup (${args.warmup}) must be smaller than --repeat (${args.repeat})`)
  if (args.serve && args.browser !== 'safari-manual') throw new Error('--serve applies to --browser safari-manual only')
  return args
}

function printHelp(): void {
  console.log(`interactive-viewer bench CLI (S0.3 scaffold — no viewer SLA is claimed)

  bun scripts/bench-interactive-viewer.ts --browser chromium|webkit|safari-manual
    --profile desktop|mobile --fixture F1 [--repeat N] [--output DIR]
    [--mode smoke|release] [--baseline FILE] [--warmup N] [--serve] [--port 4173]

  chromium|webkit  automated real-browser collection via Playwright; writes
                    JSON + CSV artifacts to --output.
  safari-manual    with --serve: serve the harness page on the LAN/local
                    interface, print the device URL, and ingest/validate/
                    hash-verify result JSON POSTed back from the device.
                    Labeled by the device itself; never relabeled Safari by
                    emulation.`)
}

function buildCommit(): string {
  try {
    return execSync('git rev-parse HEAD', { cwd: REPO_ROOT }).toString().trim()
  } catch {
    return 'unknown'
  }
}

async function waitForUrl(url: string, timeoutMs = 30_000): Promise<void> {
  const startedAt = Date.now()
  for (;;) {
    try {
      const response = await fetch(url)
      if (response.ok) return
    } catch {
      // not up yet
    }
    if (Date.now() - startedAt > timeoutMs) throw new Error(`timed out waiting for ${url}`)
    await new Promise((resolve) => setTimeout(resolve, 250))
  }
}

async function withViteServer(port: number, host: string, run: () => Promise<void>): Promise<{ reused: boolean }> {
  const url = `http://127.0.0.1:${port}${HARNESS_PAGE_PATH}`
  let reused = true
  let child: ReturnType<typeof Bun.spawn> | null = null
  try {
    await fetch(url)
  } catch {
    reused = false
    child = Bun.spawn(['bunx', 'vite', '--port', String(port), '--strictPort', '--host', host], {
      cwd: REPO_ROOT,
      stdout: 'ignore',
      stderr: 'ignore',
    })
    await waitForUrl(url)
  }
  try {
    await run()
  } finally {
    child?.kill()
  }
  return { reused }
}

/** Re-tag a result with driver-known environment facts and a fresh integrity hash. */
async function retagResult(result: HarnessResult, env: Partial<HarnessEnvironment>): Promise<HarnessResult> {
  const { hash: _previousHash, ...rest } = result
  const retagged = { ...rest, environment: { ...rest.environment, ...env } }
  const hash = await hashResult(retagged)
  return { ...retagged, hash }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

interface RunMetadata {
  browser: string
  browserVersion: string
  fixture: FixtureId
  profile: 'desktop' | 'mobile'
  mode: 'smoke' | 'release'
  commit: string
  baselineComparison?: Array<{ metric: string; baselineP50: number | null; p50: number | null; delta: number | null }>
}

interface CollectedRun extends RunMetadata {
  results: HarnessResult[]
}

function runMetadata(run: CollectedRun): RunMetadata {
  const { results: _results, ...metadata } = run
  return metadata
}

async function collectAutomated(args: BenchArgs): Promise<CollectedRun> {
  const engine = args.browser === 'chromium' ? chromium : webkit
  const browser = await engine.launch({ headless: true })
  const browserVersion = browser.version()
  const commit = buildCommit()
  const results: HarnessResult[] = []
  try {
    for (let run = 0; run < args.repeat; run++) {
      const context = await browser.newContext({
        viewport: VIEWPORTS[args.profile],
        deviceScaleFactor: 2,
        isMobile: args.profile === 'mobile',
        hasTouch: args.profile === 'mobile',
      })
      const page = await context.newPage()
      const pageUrl =
        `http://127.0.0.1:${args.port}${HARNESS_PAGE_PATH}` +
        `?fixture=${args.fixture}&profile=${args.profile}&mode=${args.mode}&cache=cold` +
        `&commit=${encodeURIComponent(commit)}&warmup=${args.warmup > 0 ? 1 : 0}`
      await page.goto(pageUrl)
      await page.waitForFunction(() => (window as { __viewerHarness?: { booted: boolean } }).__viewerHarness?.booted === true, undefined, { timeout: 30_000 })
      for (let iteration = 0; iteration < args.repeat; iteration++) {
        await page.evaluate(
          ({ label, warmup }) =>
            (window as { __viewerHarness?: { recordRoundTrip(label: string, warmup?: boolean): Promise<void> } }).__viewerHarness!.recordRoundTrip(label, warmup),
          { label: `run-${run}-${iteration}`, warmup: iteration < args.warmup },
        )
      }
      const exported = (await page.evaluate(
        () => (window as { __viewerHarness?: { exportResult(): Promise<HarnessResult> } }).__viewerHarness!.exportResult(),
      )) as HarnessResult
      await context.close()
      const driverLabel = args.browser === 'webkit' ? 'webkit-playwright' : 'chromium'
      const tagged = await retagResult(exported, {
        browser: {
          name: driverLabel,
          version: browser.version(),
          engine: args.browser === 'webkit' ? 'webkit' : 'blink',
          userAgent: exported.environment.browser.userAgent,
          automation: 'playwright',
        },
      })
      const validation = validateHarnessResult(tagged)
      if (!validation.ok) throw new Error(`result ${run} failed schema validation: ${validation.errors.join('; ')}`)
      const verification = await verifyHarnessResult(tagged)
      if (!verification.ok) throw new Error(`result ${run} failed hash verification: ${verification.errors.join('; ')}`)
      results.push(tagged)
      console.log(
        `  collected result ${run + 1}/${args.repeat}: ${args.fixture} ${args.profile} ${driverLabel} ` +
          `${tagged.samples.length} sample series, ${Object.keys(tagged.metrics).length} metrics`,
      )
    }
  } finally {
    await browser.close()
  }
  return {
    browser: args.browser === 'webkit' ? 'webkit-playwright' : 'chromium',
    browserVersion,
    fixture: args.fixture,
    profile: args.profile,
    mode: args.mode,
    commit,
    results,
    baselineComparison: args.baseline ? await compareBaseline(args.baseline, results[0]) : undefined,
  }
}

async function compareBaseline(baselinePath: string, result: HarnessResult) {
  const raw = JSON.parse(await Bun.file(baselinePath).text()) as unknown
  const results = Array.isArray(raw) ? raw : isRecord(raw) && Array.isArray(raw.results) ? raw.results : null
  if (!results || results.length === 0) throw new Error(`baseline ${baselinePath} contains no results`)
  const baseline = results[0] as HarnessResult
  const rows: Array<{ metric: string; baselineP50: number | null; p50: number | null; delta: number | null }> = []
  for (const [name, metric] of Object.entries(result.metrics)) {
    const base = baseline.metrics[name]
    const baselineP50 = base?.status === 'ok' || base?.status === 'insufficient-samples' ? base.p50 : null
    const p50 = metric.status === 'ok' || metric.status === 'insufficient-samples' ? metric.p50 : null
    rows.push({ metric: name, baselineP50, p50, delta: baselineP50 !== null && p50 !== null ? p50 - baselineP50 : null })
  }
  return rows
}

function writeArtifacts(run: CollectedRun, outputDir: string): { json: string; csv: string } {
  mkdirSync(outputDir, { recursive: true })
  const stamp = new Date().toISOString().replace(/[:.]/g, '-')
  const jsonPath = join(outputDir, `viewer-${run.browser}-${run.profile}-${run.fixture}-${run.mode}-${stamp}.json`)
  const csvPath = jsonPath.replace(/\.json$/, '.csv')
  writeFileSync(jsonPath, `${JSON.stringify({ schemaVersion: 1, run: runMetadata(run), results: run.results }, null, 2)}\n`)
  const csvLines = [
    'record_type,result_index,fixture_id,profile,mode,browser,sample_name,unit,value,warmup_excluded,metric_status,p50,p95,max,samples',
  ]
  run.results.forEach((result, index) => {
    for (const sample of result.samples) {
      for (const value of sample.values) {
        csvLines.push(
          ['sample', index, result.environment.fixtureId, result.environment.profile, result.environment.mode, result.environment.browser.name, sample.name, unitCsv(sample.unit), value, sample.warmupExcluded, '', '', '', ''].join(','),
        )
      }
    }
    for (const [name, metric] of Object.entries(result.metrics)) {
      csvLines.push(
        ['metric', index, result.environment.fixtureId, result.environment.profile, result.environment.mode, result.environment.browser.name, name, '', '', '', metric.status, metricP(metric, 'p50'), metricP(metric, 'p95'), metricP(metric, 'max'), metricP(metric, 'samples')].join(','),
      )
    }
  })
  writeFileSync(csvPath, `${csvLines.join('\n')}\n`)
  return { json: jsonPath, csv: csvPath }
}

function unitCsv(unit: string): string {
  return unit.includes(',') ? `"${unit}"` : unit
}

function metricP(metric: HarnessMetric, field: 'p50' | 'p95' | 'max' | 'samples'): string {
  if (metric.status === 'unavailable') return ''
  if (field === 'p95' && metric.status === 'insufficient-samples') return ''
  return String(metric[field])
}

function detectLanIp(): string {
  for (const list of Object.values(networkInterfaces())) {
    for (const address of list ?? []) {
      if (address.family === 'IPv4' && !address.internal) return address.address
    }
  }
  return '127.0.0.1'
}

/** Serve the harness page for a real device and ingest its posted result. */
async function serveForManualDevice(args: BenchArgs): Promise<void> {
  const lanIp = detectLanIp()
  const pageUrl =
    `http://${lanIp}:${args.port}${HARNESS_PAGE_PATH}` +
    `?fixture=${args.fixture}&profile=${args.profile}&mode=${args.mode}&cache=cold` +
    `&commit=${encodeURIComponent(buildCommit())}&warmup=0`
  const collectorPort = args.port + 1
  const collectorBase = `http://${lanIp}:${collectorPort}`
  const ingestUrl = `${collectorBase}/result`

  const vite = Bun.spawn(['bunx', 'vite', '--port', String(args.port), '--strictPort', '--host'], {
    cwd: REPO_ROOT,
    stdout: 'ignore',
    stderr: 'ignore',
  })
  let server: ReturnType<typeof Bun.serve> | null = null
  try {
    await waitForUrl(`http://127.0.0.1:${args.port}${HARNESS_PAGE_PATH}`)
    mkdirSync(args.output, { recursive: true })
    server = Bun.serve({
      hostname: '0.0.0.0',
      port: collectorPort,
      async fetch(request) {
        const path = new URL(request.url).pathname
        if (request.method === 'POST' && path === '/result') {
          let payload: unknown
          try {
            payload = await request.json()
          } catch {
            return Response.json({ ok: false, errors: ['request body is not valid JSON'] }, { status: 400 })
          }
          const validation = validateHarnessResult(payload)
          if (!validation.ok) {
            return Response.json({ ok: false, errors: validation.errors }, { status: 400 })
          }
          const verification = await verifyHarnessResult(payload)
          if (!verification.ok) {
            return Response.json({ ok: false, errors: verification.errors }, { status: 400 })
          }
          const device = (payload as HarnessResult).environment.browser.name
          const stamp = new Date().toISOString().replace(/[:.]/g, '-')
          const artifact = join(args.output, `viewer-manual-${device}-${args.fixture}-${stamp}.json`)
          writeFileSync(
            artifact,
            `${JSON.stringify(
              {
                schemaVersion: 1,
                run: {
                  browser: device,
                  collection: 'manual-device-post',
                  note: 'device-reported labeling; physical-device verification is the operator responsibility (spec 12.2)',
                },
                results: [payload],
              },
              null,
              2,
            )}\n`,
          )
          console.log(`  ingested device result -> ${artifact}`)
          return Response.json({ ok: true, artifact })
        }
        return Response.json({ ok: true, service: 'officeview viewer harness collector', page: pageUrl })
      },
    })
    console.log(`Serving the harness page for manual device collection:`)
    console.log(`  device URL (open in Safari on the device): ${pageUrl}&collect=${encodeURIComponent(ingestUrl)}`)
    console.log(`  collector endpoint (LAN):                  ${ingestUrl}`)
    console.log(`  artifacts are written to:                  ${args.output}`)
    console.log(`On the device: load the page, press "Export result" — the page posts the result here.`)
    console.log(`Results keep the device's own browser labeling; WebKit emulation is never relabeled Safari.`)
    console.log(`Press Ctrl-C to stop.`)
    await new Promise<void>((resolve) => {
      process.on('SIGINT', () => resolve())
    })
  } finally {
    server?.stop(true)
    vite.kill()
  }
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2))
  console.log(`interactive-viewer bench (S0.3 scaffold — measures the harness frame, claims no viewer SLA)`)
  console.log(`  browser=${args.browser} profile=${args.profile} fixture=${args.fixture} repeat=${args.repeat} mode=${args.mode} port=${args.port}`)
  if (args.browser === 'safari-manual') {
    if (!args.serve) throw new Error('--browser safari-manual requires --serve')
    await serveForManualDevice(args)
    return
  }
  const { reused } = await withViteServer(args.port, '127.0.0.1', async () => {
    const run = await collectAutomated(args)
    const { json, csv } = writeArtifacts(run, args.output)
    console.log(`artifacts:`)
    console.log(`  JSON: ${json}`)
    console.log(`  CSV:  ${csv}`)
    if (run.baselineComparison) {
      console.log(`baseline comparison (p50 deltas, informational only):`)
      for (const row of run.baselineComparison) {
        if (row.delta === null) console.log(`  ${row.metric}: not comparable`)
        else console.log(`  ${row.metric}: ${row.delta >= 0 ? '+' : ''}${row.delta.toFixed(3)} ms vs baseline p50 ${row.baselineP50}`)
      }
    }
    console.log(`note: ${args.browser === 'webkit' ? 'webkit-playwright emulation — NOT Safari or physical-device evidence' : 'scaffold measurement — no viewer SLA is claimed'}`)
    return {}
  })
  console.log(`  vite dev server ${reused ? 'reused' : 'started and stopped'}`)
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : String(error))
  process.exit(1)
})
