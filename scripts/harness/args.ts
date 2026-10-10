/**
 * Argument parsing for the bench CLI. Pure: it throws instead of exiting and
 * touches neither the filesystem nor the network, so it is unit-tested.
 */
import { FIXTURE_IDS, MAX_SAMPLE_VALUES } from '../../src/debug/harness/result'
import type { FixtureId, HarnessMode, HarnessProfile, OperatorMetadata } from '../../src/debug/harness/types'

export interface BenchArgs {
  browser: 'chromium' | 'webkit' | 'safari-manual'
  profile: HarnessProfile
  fixture: FixtureId
  /** Measured input→present round-trips, pooled into ONE result. */
  repeat: number
  /** Extra discarded round-trips run first; their count is reported. */
  warmup: number
  output: string
  mode: HarnessMode
  baseline: string | null
  serve: boolean
  port: number
  operator: OperatorMetadata
}

/**
 * Flips to true when S0.2 delivers the real F1..F8 fixtures. Until then every
 * fixture is the 120x10 placeholder and `--mode release` would label a
 * placeholder measurement as release evidence, so it is refused.
 */
export const REAL_FIXTURES_AVAILABLE = false

export const HELP_TEXT = `interactive-viewer bench CLI (S0.3 scaffold — no viewer SLA is claimed)

  bun scripts/bench-interactive-viewer.ts --browser chromium|webkit|safari-manual
    --profile desktop|mobile --fixture F1 [--repeat N] [--warmup W] [--output DIR]
    [--mode smoke] [--baseline FILE] [--serve] [--port 4173]
    [--device-model M] [--os-version V] [--thermal S] [--battery B]

  --repeat N   N measured input-to-present round-trips in ONE page load, pooled
               into ONE result (max ${MAX_SAMPLE_VALUES}). p95 needs >= 30 samples; the three-run
               gate prints insufficient-samples, which is the honest outcome.
  --warmup W   W extra round-trips run first and discarded; the count is recorded
               in each series' warmupExcluded. W does not count towards --repeat.
  --mode       smoke only for now: release is refused while the F1..F8 fixtures
               are placeholders (S0.2).
  chromium|webkit  automated real-browser collection via Playwright; writes
               JSON + CSV artifacts to --output. WebKit is labeled
               webkit-playwright, never Safari.
  safari-manual    with --serve: serve the harness page on the LAN, print the
               tokenized device URL, and ingest/validate/hash-verify result JSON
               POSTed back from the device on the same origin. Operator flags
               record device model, OS version, thermal and battery state.
  --baseline FILE  informational p50 deltas; the file is validated and checked
               for comparability BEFORE any browser launches.`

const OPERATOR_FLAGS: Record<string, keyof OperatorMetadata> = {
  'device-model': 'deviceModel',
  'os-version': 'osVersion',
  thermal: 'thermalState',
  battery: 'batteryState',
}

const KNOWN_FLAGS = ['browser', 'profile', 'fixture', 'repeat', 'warmup', 'output', 'mode', 'baseline', 'serve', 'port']

export function parseArgs(argv: string[]): BenchArgs | { help: true } {
  const args: BenchArgs = {
    browser: 'chromium',
    profile: 'desktop',
    fixture: 'F1',
    repeat: 1,
    warmup: 0,
    output: '.tmp/viewer-bench',
    mode: 'smoke',
    baseline: null,
    serve: false,
    port: 4173,
    operator: { deviceModel: null, osVersion: null, thermalState: null, batteryState: null },
  }
  let requestedMode: string = args.mode
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]
    if (arg === '-h' || arg === '--help') return { help: true }
    const match = /^--([^=]+)(?:=(.*))?$/.exec(arg)
    if (!match) throw new Error(`unexpected argument ${arg}; see --help`)
    const [, name, inlineValue] = match
    if (!KNOWN_FLAGS.includes(name) && !(name in OPERATOR_FLAGS)) throw new Error(`unknown flag --${name}; see --help`)
    let value = inlineValue
    if (name !== 'serve' && value === undefined) {
      const next = argv[i + 1]
      if (next === undefined || next.startsWith('--')) throw new Error(`--${name} requires a value`)
      value = next
      i++
    }
    const text = value ?? ''
    switch (name) {
      case 'browser':
        if (text !== 'chromium' && text !== 'webkit' && text !== 'safari-manual') throw new Error(`--browser must be chromium|webkit|safari-manual, got ${text}`)
        args.browser = text
        break
      case 'profile':
        if (text !== 'desktop' && text !== 'mobile') throw new Error(`--profile must be desktop|mobile, got ${text}`)
        args.profile = text
        break
      case 'fixture':
        if (!FIXTURE_IDS.includes(text as FixtureId)) throw new Error(`--fixture must be one of ${FIXTURE_IDS.join(', ')}, got ${text}`)
        args.fixture = text as FixtureId
        break
      case 'repeat': {
        const repeat = Number(text)
        if (!Number.isInteger(repeat) || repeat < 1 || repeat > MAX_SAMPLE_VALUES) throw new Error(`--repeat must be an integer in 1..${MAX_SAMPLE_VALUES}, got ${text}`)
        args.repeat = repeat
        break
      }
      case 'warmup': {
        const warmup = Number(text)
        if (!Number.isInteger(warmup) || warmup < 0 || warmup > 10_000) throw new Error(`--warmup must be an integer in 0..10000, got ${text}`)
        args.warmup = warmup
        break
      }
      case 'output':
        args.output = text
        break
      case 'mode':
        if (text !== 'smoke' && text !== 'release') throw new Error(`--mode must be smoke|release, got ${text}`)
        requestedMode = text
        break
      case 'baseline':
        args.baseline = text
        break
      case 'serve':
        if (value !== undefined && value !== 'true' && value !== 'false' && value !== '') throw new Error(`--serve takes no value (or true|false), got ${value}`)
        args.serve = value !== 'false'
        break
      case 'port': {
        const port = Number(text)
        if (!Number.isInteger(port) || port < 1 || port > 65535) throw new Error(`--port must be 1..65535, got ${text}`)
        args.port = port
        break
      }
      default:
        args.operator[OPERATOR_FLAGS[name]] = text === '' ? null : text
    }
  }
  if (requestedMode === 'release' && !REAL_FIXTURES_AVAILABLE) {
    throw new Error('--mode release is refused: every fixture is still the 120x10 placeholder (real F1..F8 arrive with S0.2), and a placeholder measurement must not be labeled release evidence')
  }
  args.mode = requestedMode as HarnessMode
  if (args.browser === 'safari-manual' && !args.serve) throw new Error('--browser safari-manual requires --serve')
  if (args.serve && args.browser !== 'safari-manual') throw new Error('--serve applies to --browser safari-manual only')
  if (args.baseline !== null && args.browser === 'safari-manual') throw new Error('--baseline applies to automated browsers only, not safari-manual')
  return args
}
