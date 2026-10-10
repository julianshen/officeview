// @vitest-environment node
/**
 * Bench CLI building blocks (findings 9, 12, 13, 14, 15): argument parsing,
 * artifact/CSV writing, baseline validation and comparison, driver re-tagging,
 * LAN address detection and harness identity checks. All pure or temp-dir
 * based; no browser involved.
 */
import { describe, expect, it } from 'vitest'
import { mkdtempSync, readFileSync, readdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { MAX_SAMPLE_VALUES, validateHarnessResult, verifyHarnessResult } from '../src/debug/harness/result'
import { HELP_TEXT, parseArgs, type BenchArgs } from '../scripts/harness/args'
import { csvField, metricCell, writeArtifacts, type CollectedRun } from '../scripts/harness/artifacts'
import { checkBaselineComparable, compareResults, parseBaseline } from '../scripts/harness/baseline'
import { retagResult } from '../scripts/harness/retag'
import { detectLanIp } from '../scripts/harness/lan'
import { checkIdentity, commitLabel, HARNESS_SERVICE, identityDigest } from '../scripts/harness/identity'
import { buildResult, clone, makeSample, range, rehash } from './helpers/harnessResult'

function args(argv: string[]): BenchArgs {
  const parsed = parseArgs(argv)
  if ('help' in parsed) throw new Error('unexpected help')
  return parsed
}

describe('parseArgs', () => {
  it('applies documented defaults', () => {
    expect(args([])).toEqual({
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
    })
  })

  it('parses every flag in both "--flag value" and "--flag=value" forms', () => {
    const parsed = args(['--browser', 'webkit', '--profile=mobile', '--fixture', 'F3', '--repeat=40', '--warmup', '5', '--output=out', '--baseline', 'b.json', '--port=5000', '--device-model', 'iPad13,4', '--os-version=17.5', '--thermal', 'nominal', '--battery=80%'])
    expect(parsed).toMatchObject({ browser: 'webkit', profile: 'mobile', fixture: 'F3', repeat: 40, warmup: 5, output: 'out', baseline: 'b.json', port: 5000 })
    expect(parsed.operator).toEqual({ deviceModel: 'iPad13,4', osVersion: '17.5', thermalState: 'nominal', batteryState: '80%' })
  })

  it('safari-manual requires --serve, and --serve requires safari-manual', () => {
    expect(() => args(['--browser', 'safari-manual'])).toThrow(/--serve/)
    expect(() => args(['--serve'])).toThrow(/safari-manual/)
    expect(args(['--browser', 'safari-manual', '--serve'])).toMatchObject({ browser: 'safari-manual', serve: true })
    expect(args(['--browser=safari-manual', '--serve=true']).serve).toBe(true)
  })

  it('rejects unknown flags, stray arguments and missing values', () => {
    expect(() => args(['--bogus'])).toThrow(/unknown flag --bogus/)
    expect(() => args(['stray'])).toThrow(/unexpected argument/)
    expect(() => args(['--port'])).toThrow(/requires a value/)
    expect(() => args(['--port', '--serve'])).toThrow(/requires a value/)
  })

  it('validates enum and numeric flags', () => {
    expect(() => args(['--browser', 'firefox'])).toThrow(/--browser/)
    expect(() => args(['--profile', 'tablet'])).toThrow(/--profile/)
    expect(() => args(['--fixture', 'F9'])).toThrow(/--fixture/)
    expect(() => args(['--mode', 'debug'])).toThrow(/--mode/)
    for (const bad of ['0', '-1', '1.5', 'abc', String(MAX_SAMPLE_VALUES + 1)]) expect(() => args(['--repeat', bad]), bad).toThrow(/--repeat/)
    expect(args(['--repeat', String(MAX_SAMPLE_VALUES)]).repeat).toBe(MAX_SAMPLE_VALUES)
    for (const bad of ['-1', '1.5', 'x']) expect(() => args(['--warmup', bad]), bad).toThrow(/--warmup/)
    for (const bad of ['0', '65536', '1.5', 'x']) expect(() => args(['--port', bad]), bad).toThrow(/--port/)
    expect(args(['--port', '65535']).port).toBe(65535)
  })

  it('--warmup is extra runs: it need not be smaller than --repeat', () => {
    expect(args(['--repeat', '3', '--warmup', '10'])).toMatchObject({ repeat: 3, warmup: 10 })
  })

  it('refuses --mode release while only placeholder fixtures exist', () => {
    expect(() => args(['--mode', 'release'])).toThrow(/placeholder/)
    expect(() => args(['--mode=release', '--browser', 'safari-manual', '--serve'])).toThrow(/placeholder/)
  })

  it('--baseline applies to automated browsers only', () => {
    expect(() => args(['--browser', 'safari-manual', '--serve', '--baseline', 'x.json'])).toThrow(/baseline/)
  })

  it('returns help for --help / -h and the help text documents repeat semantics', () => {
    expect(parseArgs(['--help'])).toEqual({ help: true })
    expect(parseArgs(['-h'])).toEqual({ help: true })
    expect(HELP_TEXT).toMatch(/--repeat/)
    expect(HELP_TEXT).toMatch(/pooled/i)
  })
})

describe('csvField', () => {
  it('quotes commas, quotes and newlines and doubles embedded quotes', () => {
    expect(csvField('plain')).toBe('plain')
    expect(csvField('a,b')).toBe('"a,b"')
    expect(csvField('say "hi"')).toBe('"say ""hi"""')
    expect(csvField('line\nbreak')).toBe('"line\nbreak"')
    expect(csvField('cr\rx')).toBe('"cr\rx"')
  })
  it('neutralizes spreadsheet formula injection in text', () => {
    for (const lead of ['=', '+', '-', '@']) expect(csvField(`${lead}SUM(A1)`)).toBe(`'${lead}SUM(A1)`)
  })
  it('renders numbers, booleans and null plainly', () => {
    expect(csvField(1.5)).toBe('1.5')
    expect(csvField(-3)).toBe('-3')
    expect(csvField(true)).toBe('true')
    expect(csvField(null)).toBe('')
  })
})

describe('metricCell', () => {
  it('never prints a p95 for insufficient samples and prints nothing for unavailable', () => {
    expect(metricCell({ status: 'unavailable', reason: 'x' }, 'p50')).toBe('')
    expect(metricCell({ status: 'insufficient-samples', p50: 1, max: 2, samples: 3, reason: 'r' }, 'p95')).toBe('')
    expect(metricCell({ status: 'insufficient-samples', p50: 1, max: 2, samples: 3, reason: 'r' }, 'p50')).toBe('1')
    expect(metricCell({ status: 'none-observed', samples: 0, reason: 'r' }, 'samples')).toBe('0')
    expect(metricCell({ status: 'none-observed', samples: 0, reason: 'r' }, 'p50')).toBe('')
    expect(metricCell({ status: 'ok', p50: 1, p95: 2, max: 3, samples: 30 }, 'p95')).toBe('2')
  })
})

describe('writeArtifacts', () => {
  async function run(): Promise<CollectedRun> {
    const result = await buildResult([{ ...makeSample(range(30)), warmupExcluded: 4 }, makeSample([1, 2], 'operation', 'second')])
    return {
      browser: 'chromium',
      browserVersion: '1.0',
      requestedFixtureId: 'F1',
      profile: 'desktop',
      mode: 'smoke',
      commit: 'deadbeef',
      repeat: 30,
      warmup: 4,
      results: [result],
    }
  }

  it('writes compact-enough JSON and a CSV with the excluded warmup count, exclusively', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'artifacts-'))
    const collected = await run()
    const now = () => new Date('2026-10-10T00:00:00.000Z')
    const { json, csv } = writeArtifacts(collected, dir, now, () => 'aaaa1111')
    expect(readdirSync(dir).sort()).toEqual([
      'viewer-chromium-desktop-F1-smoke-2026-10-10T00-00-00-000Z-aaaa1111.csv',
      'viewer-chromium-desktop-F1-smoke-2026-10-10T00-00-00-000Z-aaaa1111.json',
    ])
    const parsed = JSON.parse(readFileSync(json, 'utf8'))
    expect(parsed.run).toMatchObject({ browser: 'chromium', repeat: 30, warmup: 4, commit: 'deadbeef' })
    expect(parsed.results).toHaveLength(1)
    const lines = readFileSync(csv, 'utf8').trim().split('\n')
    expect(lines[0]).toBe('record_type,result_index,requested_fixture,fixture_id,evidence,profile,mode,browser,sample_name,unit,value,warmup_excluded,metric_status,p50,p95,max,samples')
    const sampleRows = lines.filter((line) => line.startsWith('sample,0,F1,F1,device,desktop,smoke,chromium,op,ms,'))
    expect(sampleRows).toHaveLength(30)
    expect(sampleRows[0].split(',')[11]).toBe('4')
    expect(lines.some((line) => line.startsWith('metric,0,F1,F1,device,desktop,smoke,chromium,op,,,,ok,15,29,30,30'))).toBe(true)
    // same name+stamp+id again must not overwrite
    expect(() => writeArtifacts(collected, dir, now, () => 'aaaa1111')).toThrow(/EEXIST|exist/i)
    expect(readFileSync(json, 'utf8')).toBe(JSON.stringify({ schemaVersion: 1, run: parsed.run, results: parsed.results }, null, 2) + '\n')
  })

  it('escapes hostile text in CSV cells (names with commas, quotes, formulas)', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'artifacts-'))
    const collected = await run()
    const hostile = clone(collected.results[0]) as any
    hostile.samples[0].name = '=cmd|"x",y'
    hostile.metrics = { '=cmd|"x",y': hostile.metrics.op }
    const { csv } = writeArtifacts({ ...collected, results: [hostile] }, dir)
    const text = readFileSync(csv, 'utf8')
    expect(text).toContain(`"'=cmd|""x"",y"`)
    expect(text.split('\n').every((line) => line === '' || !line.startsWith('='))).toBe(true)
  })
})

describe('baseline handling', () => {
  it('parses an array or a {results:[]} document of verified results', async () => {
    const result = await buildResult()
    expect(await parseBaseline(JSON.stringify([result]))).toHaveLength(1)
    expect(await parseBaseline(JSON.stringify({ schemaVersion: 1, run: {}, results: [result] }))).toHaveLength(1)
  })

  it('refuses malformed, empty, tampered or dishonest baselines with a clear reason', async () => {
    const result = await buildResult()
    await expect(parseBaseline('{nope')).rejects.toThrow(/not valid JSON/)
    await expect(parseBaseline('[]')).rejects.toThrow(/no results/)
    await expect(parseBaseline('{"results":"x"}')).rejects.toThrow(/no results/)
    const tampered = clone(result) as any
    tampered.environment.note = 'edited'
    await expect(parseBaseline(JSON.stringify([tampered]))).rejects.toThrow(/hash mismatch/)
    const forged = await rehash(result, (d) => (d.metrics.op.p50 = 1))
    await expect(parseBaseline(JSON.stringify([forged]))).rejects.toThrow(/derived/)
  })

  it('flags profile, browser, fixture and mode mismatches before launching', async () => {
    const result = await buildResult()
    const expected = { browserLabel: 'chromium', profile: 'desktop', requestedFixtureId: 'F1', mode: 'smoke' }
    expect(checkBaselineComparable(result, expected)).toEqual([])
    const reasons = checkBaselineComparable(result, { browserLabel: 'webkit-playwright', profile: 'mobile', requestedFixtureId: 'F2', mode: 'release' })
    expect(reasons).toHaveLength(4)
    expect(reasons.join('\n')).toMatch(/browser.*profile.*|profile/s)
  })

  it('compares p50 per metric and warns about environment differences without dropping rows', async () => {
    const baseline = await buildResult([makeSample([10, 20, 30])])
    const current = await buildResult([makeSample([20, 30, 40])], { dpr: 3, buildCommit: 'cafebabe' })
    const comparison = compareResults(baseline, current)
    expect(comparison.rows.find((row) => row.metric === 'op')).toEqual({ metric: 'op', baselineP50: 20, p50: 30, delta: 10 })
    expect(comparison.warnings.join('\n')).toMatch(/dpr/i)
    expect(comparison.warnings.join('\n')).toMatch(/commit/i)
  })

  it('marks metrics without a numeric p50 on either side as not comparable', async () => {
    const baseline = await buildResult([])
    const current = await buildResult([makeSample([1])])
    const rows = compareResults(baseline, current).rows
    expect(rows.find((row) => row.metric === 'op')).toMatchObject({ baselineP50: null, p50: 1, delta: null })
  })
})

describe('retagResult (driver facts)', () => {
  const operator = { deviceModel: null, osVersion: null, thermalState: null, batteryState: null }

  it('labels Playwright WebKit webkit-playwright / emulated-webkit and still passes ingest validation', async () => {
    const page = await buildResult([makeSample(range(30))], {
      browser: { name: 'webkit', version: '26.0', engine: 'webkit', userAgent: 'UA', automation: 'webdriver' },
    })
    const tagged = await retagResult(page, { engine: 'webkit', browserVersion: '26.0', hostPlatform: 'darwin', commit: 'deadbeef', operator })
    expect(tagged.environment.browser).toMatchObject({ name: 'webkit-playwright', engine: 'webkit', automation: 'playwright', userAgent: 'UA' })
    expect(tagged.environment.evidence).toBe('emulated-webkit')
    await expect(verifyHarnessResult(tagged)).resolves.toEqual({ ok: true, errors: [] })
  })

  it('never produces a safari label, and labels Chromium automation-playwright', async () => {
    const page = await buildResult()
    const tagged = await retagResult(page, { engine: 'chromium', browserVersion: '147.0', hostPlatform: 'darwin', commit: 'deadbeef', operator })
    expect(tagged.environment.browser).toMatchObject({ name: 'chromium', engine: 'blink', automation: 'playwright' })
    expect(JSON.stringify(tagged)).not.toMatch(/"safari"/)
    expect(tagged.environment.cache.note).toMatch(/fresh Playwright/)
    expect(validateHarnessResult(tagged).ok).toBe(true)
  })

  it('takes the OS from the driver host and records operator metadata', async () => {
    const page = await buildResult([], { os: { platform: 'unknown', platformNote: 'x', userAgent: 'UA' } })
    const tagged = await retagResult(page, {
      engine: 'chromium',
      browserVersion: '1',
      hostPlatform: 'darwin',
      commit: 'cafebabe-dirty',
      operator: { deviceModel: 'MacBookPro', osVersion: '26.0', thermalState: 'nominal', batteryState: 'AC' },
    })
    expect(tagged.environment.os).toMatchObject({ platform: 'macOS', userAgent: 'UA' })
    expect(tagged.environment.os.platformNote).toMatch(/driver host/)
    expect(tagged.environment.operator).toEqual({ deviceModel: 'MacBookPro', osVersion: '26.0', thermalState: 'nominal', batteryState: 'AC' })
    expect(tagged.environment.buildCommit).toBe('cafebabe-dirty')
    await expect(verifyHarnessResult(tagged)).resolves.toEqual({ ok: true, errors: [] })
  })

  it('keeps the placeholder label for placeholder content', async () => {
    const page = await buildResult([], { source: { adapter: 'xlsx-placeholder', bytes: 1, rows: 1, cols: 1 }, fixtureId: null })
    const tagged = await retagResult(page, { engine: 'webkit', browserVersion: '26', hostPlatform: 'linux', commit: 'deadbeef', operator })
    expect(tagged.environment.evidence).toBe('scaffold-placeholder')
    expect(tagged.environment.browser.name).toBe('webkit-playwright')
  })
})

describe('detectLanIp', () => {
  it('returns the first external IPv4 address', () => {
    const interfaces = {
      lo0: [{ address: '127.0.0.1', family: 'IPv4', internal: true }],
      en0: [
        { address: 'fe80::1', family: 'IPv6', internal: false },
        { address: '192.168.1.20', family: 'IPv4', internal: false },
      ],
    } as never
    expect(detectLanIp(interfaces)).toEqual({ ip: '192.168.1.20', loopbackFallback: false })
  })
  it('reports the loopback fallback explicitly when no LAN interface exists', () => {
    expect(detectLanIp({ lo0: [{ address: '127.0.0.1', family: 'IPv4', internal: true }] } as never)).toEqual({ ip: '127.0.0.1', loopbackFallback: true })
    expect(detectLanIp({})).toEqual({ ip: '127.0.0.1', loopbackFallback: true })
  })
})

describe('harness identity', () => {
  const expected = { commit: 'a'.repeat(40), root: '/repo' }
  const marker = { service: HARNESS_SERVICE, commit: expected.commit, dirty: false, salt: 's'.repeat(32), rootDigest: identityDigest('/repo', expected.commit, 's'.repeat(32)) }
  it('accepts only the harness service for this commit and checkout', () => {
    expect(checkIdentity(marker, expected)).toBeNull()
  })
  it('refuses a foreign server, an SPA fallback page, a stale commit, another checkout and a missing salt', () => {
    expect(checkIdentity('<!doctype html><html>', expected)).toMatch(/not.*harness/i)
    expect(checkIdentity({ service: 'something-else' }, expected)).toMatch(/not.*harness/i)
    expect(checkIdentity({ ...marker, commit: 'b'.repeat(40) }, expected)).toMatch(/commit/)
    expect(checkIdentity({ ...marker, rootDigest: '0'.repeat(64) }, expected)).toMatch(/checkout/)
    const { salt: _dropped, ...unsalted } = marker
    void _dropped
    expect(checkIdentity(unsalted, expected)).toMatch(/salt/)
  })
  it('labels dirty trees', () => {
    expect(commitLabel({ commit: 'abc1234', dirty: false })).toBe('abc1234')
    expect(commitLabel({ commit: 'abc1234', dirty: true })).toBe('abc1234-dirty')
    expect(commitLabel({ commit: 'unknown', dirty: true })).toBe('unknown')
  })
})
