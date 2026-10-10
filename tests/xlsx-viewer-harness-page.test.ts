/**
 * Harness page logic (S0.3): marker pairs, warmup, long-task support, export
 * aliasing, browser classification and environment gathering. Everything is
 * driven through injected clocks/rAF/observers — deterministic, NO timing or
 * SLA assertions.
 */
import { describe, expect, it } from 'vitest'
import { classifyBrowser, gatherEnvironment, parseHarnessParams, type BrowserContext } from '../src/debug/harness/environment'
import { createViewerHarness, type HarnessDeps } from '../src/debug/harness/harness'
import { COLLECT_PATH, postResult } from '../src/debug/harness/collect'
import { verifyHarnessResult } from '../src/debug/harness/result'
import { fixtureLoader } from '../src/debug/harness/fixture'
import { makeEnvironment } from './helpers/harnessResult'

/** rAF stub whose frames the test advances explicitly. */
function manualRaf() {
  const queue: Array<() => void> = []
  return {
    raf: (callback: () => void) => void queue.push(callback),
    /** Run exactly the callbacks queued so far (one animation frame). */
    frame() {
      const running = queue.splice(0, queue.length)
      for (const callback of running) callback()
    },
    pending: () => queue.length,
  }
}

/** rAF that fires on a macrotask so awaiting a round-trip resolves on its own. */
function autoRaf(callback: () => void): void {
  setTimeout(callback, 0)
}

/** Clock that advances by `step` on every read: input→present is exact. */
function steppingClock(step: number) {
  let t = 0
  return () => (t += step)
}

function makeHarness(overrides: Partial<HarnessDeps> = {}) {
  return createViewerHarness({
    clock: steppingClock(5),
    raf: autoRaf,
    environment: () => makeEnvironment(),
    timeoutMs: 2000,
    ...overrides,
  })
}

class FakeObserver {
  static supportedEntryTypes: string[] = ['longtask']
  static instances: FakeObserver[] = []
  static throwOnObserve = false
  disconnected = false
  pending: Array<{ duration: number }> = []
  constructor(private readonly callback: (list: { getEntries(): Array<{ duration: number }> }) => void) {
    FakeObserver.instances.push(this)
  }
  observe(): void {
    if (FakeObserver.throwOnObserve) throw new Error('unsupported entry type')
  }
  disconnect(): void {
    this.disconnected = true
  }
  takeRecords(): Array<{ duration: number }> {
    const records = this.pending
    this.pending = []
    return records
  }
  /** Deliver entries the way the browser would: through the callback. */
  deliver(...durations: number[]): void {
    this.callback({ getEntries: () => durations.map((duration) => ({ duration })) })
  }
}
function fakeObserverClass(supported: string[] | undefined, throwOnObserve = false) {
  FakeObserver.supportedEntryTypes = supported as string[]
  FakeObserver.instances = []
  FakeObserver.throwOnObserve = throwOnObserve
  return FakeObserver as unknown as typeof PerformanceObserver
}

describe('markInput / markPresent', () => {
  it('records the exact input→present delta through a double rAF', async () => {
    const frames = manualRaf()
    const harness = makeHarness({ raf: frames.raf })
    harness.markInput('click') // clock -> 5
    const settled = harness.markPresent('click') // present read after two frames -> 10
    expect(harness.sampleCount('input-to-present')).toBe(0)
    frames.frame()
    expect(harness.sampleCount('input-to-present')).toBe(0) // first frame only schedules the second
    expect(frames.pending()).toBe(1)
    frames.frame()
    await settled
    const result = await harness.exportResult()
    expect(result.samples.find((s) => s.name === 'input-to-present')).toEqual({
      name: 'input-to-present',
      unit: 'ms',
      kind: 'operation',
      values: [5],
      warmupExcluded: 0,
    })
  })

  it('throws for a label that was never marked instead of silently doing nothing', () => {
    const harness = makeHarness()
    expect(() => harness.markPresent('never-marked')).toThrow(/never-marked/)
    expect(harness.sampleCount('input-to-present')).toBe(0)
  })

  it('rejects a second concurrent use of a label that is still pending', async () => {
    const frames = manualRaf()
    const harness = makeHarness({ raf: frames.raf })
    harness.markInput('dup')
    expect(() => harness.markInput('dup')).toThrow(/pending/)
    expect(() => harness.markInput('dup')).toThrow(/pending/)
    // The first pairing still completes once its frames run.
    const settled = harness.markPresent('dup')
    frames.frame()
    frames.frame()
    await settled
    expect(harness.sampleCount('input-to-present')).toBe(1)
    // After completion the label is reusable.
    harness.markInput('dup')
    const again = harness.markPresent('dup')
    frames.frame()
    frames.frame()
    await again
    expect(harness.sampleCount('input-to-present')).toBe(2)
  })

  it('rejects concurrent recordRoundTrip calls sharing one label', async () => {
    const frames = manualRaf()
    const harness = makeHarness({ raf: frames.raf })
    const first = harness.recordRoundTrip('shared')
    await expect(harness.recordRoundTrip('shared')).rejects.toThrow(/pending/)
    frames.frame()
    frames.frame()
    await first
    expect(harness.sampleCount('input-to-present')).toBe(1)
  })

  it('rejects instead of hanging when no animation frame ever fires', async () => {
    const harness = makeHarness({ raf: () => undefined, timeoutMs: 20 })
    harness.markInput('hidden-tab')
    await expect(harness.markPresent('hidden-tab')).rejects.toThrow(/timed out/)
    await expect(harness.recordRoundTrip('hidden-tab-2')).rejects.toThrow(/timed out/)
  })
})

describe('warmup', () => {
  it('a warmup round-trip resolves, is discarded, and is counted in warmupExcluded', async () => {
    const harness = makeHarness()
    await harness.recordRoundTrip('w1', true)
    await harness.recordRoundTrip('w2', true)
    for (const label of ['m1', 'm2', 'm3']) await harness.recordRoundTrip(label)
    expect(harness.sampleCount('input-to-present')).toBe(3)
    expect(harness.warmupCount('input-to-present')).toBe(2)
    const result = await harness.exportResult()
    const series = result.samples.find((s) => s.name === 'input-to-present')
    expect(series?.values).toEqual([5, 5, 5])
    expect(series?.warmupExcluded).toBe(2)
    await expect(verifyHarnessResult(result)).resolves.toEqual({ ok: true, errors: [] })
  })

  it('reports warmup even when every iteration was warmup (empty measured set stays unavailable)', async () => {
    const harness = makeHarness()
    await harness.recordRoundTrip('w1', true)
    const result = await harness.exportResult()
    expect(result.samples.find((s) => s.name === 'input-to-present')).toMatchObject({ values: [], warmupExcluded: 1 })
    expect(result.metrics['input-to-present'].status).toBe('unavailable')
  })
})

describe('sample bookkeeping', () => {
  it('rejects a later sample that changes the unit or kind of an existing series', () => {
    const harness = makeHarness()
    harness.recordOneShot('render', 'ms', 'operation', 1)
    expect(() => harness.recordOneShot('render', 'bytes', 'operation', 2)).toThrow(/unit/)
    expect(() => harness.recordOneShot('render', 'ms', 'frame', 2)).toThrow(/kind/)
    harness.recordOneShot('render', 'ms', 'operation', 3)
    expect(harness.sampleCount('render')).toBe(2)
  })

  it('setFixture releases the previous fixture before charging (no double charge)', () => {
    const harness = makeHarness()
    const grid = [['a', 'b'], ['c', 'd']]
    harness.setFixture({ adapter: 'xlsx', bytes: 1000, rows: 2, cols: 2 }, grid)
    harness.setFixture({ adapter: 'xlsx', bytes: 400, rows: 2, cols: 2 }, grid)
    expect(harness.ledger.retained('main')).toBe(400 + 2 * 2 * 8)
    expect(harness.ledger.accounting().main.peak).toBe(1000 + 2 * 2 * 8)
  })
})

describe('exportResult', () => {
  it('returns snapshots: a later sample cannot mutate an exported result or break its hash', async () => {
    const harness = makeHarness()
    await harness.recordRoundTrip('a')
    const first = await harness.exportResult()
    const valuesBefore = [...first.samples[0].values]
    await harness.recordRoundTrip('b')
    expect(first.samples[0].values).toEqual(valuesBefore)
    await expect(verifyHarnessResult(first)).resolves.toEqual({ ok: true, errors: [] })
    const second = await harness.exportResult()
    expect(second.samples.find((s) => s.name === 'input-to-present')?.values).toHaveLength(2)
    await expect(verifyHarnessResult(second)).resolves.toEqual({ ok: true, errors: [] })
  })

  it('marks viewer-kernel metrics unavailable with reasons, never zero', async () => {
    const result = await makeHarness().exportResult()
    for (const name of ['viewport-frame', 'cold-import', 'progressive-first-viewport']) {
      expect(result.metrics[name]).toEqual({ status: 'unavailable', reason: expect.stringContaining(name) })
    }
  })
})

describe('long-task collection', () => {
  it('reports unavailable when the engine does not list longtask in supportedEntryTypes (WebKit)', async () => {
    const harness = makeHarness({ performanceObserver: fakeObserverClass(['paint', 'mark']) })
    expect(harness.startLongTaskCollection()).toBe(false)
    const result = await harness.exportResult()
    expect(result.metrics['longtask-duration']).toEqual({ status: 'unavailable', reason: expect.stringContaining('not supported') })
    expect(result.samples.some((s) => s.name === 'longtask-duration')).toBe(false)
  })

  it('treats a missing PerformanceObserver or supportedEntryTypes as unsupported', () => {
    expect(makeHarness({ performanceObserver: undefined }).startLongTaskCollection()).toBe(false)
    expect(makeHarness({ performanceObserver: fakeObserverClass(undefined) }).startLongTaskCollection()).toBe(false)
  })

  it('treats an observe() that throws as unsupported', () => {
    expect(makeHarness({ performanceObserver: fakeObserverClass(['longtask'], true) }).startLongTaskCollection()).toBe(false)
  })

  it('zero long tasks in a supporting engine is a measured empty set, not unavailable', async () => {
    const harness = makeHarness({ performanceObserver: fakeObserverClass(['longtask']) })
    expect(harness.startLongTaskCollection()).toBe(true)
    const result = await harness.exportResult()
    expect(result.metrics['longtask-duration']).toMatchObject({ status: 'none-observed', samples: 0 })
    expect(result.samples.find((s) => s.name === 'longtask-duration')).toMatchObject({ kind: 'event', values: [] })
    await expect(verifyHarnessResult(result)).resolves.toEqual({ ok: true, errors: [] })
  })

  it('never started collection is unavailable with that reason', async () => {
    const result = await makeHarness({ performanceObserver: fakeObserverClass(['longtask']) }).exportResult()
    expect(result.metrics['longtask-duration']).toEqual({ status: 'unavailable', reason: expect.stringContaining('not started') })
  })

  it('export does not stop collection and flushes undelivered records', async () => {
    const ctor = fakeObserverClass(['longtask'])
    const harness = makeHarness({ performanceObserver: ctor })
    harness.startLongTaskCollection()
    const observer = FakeObserver.instances[0]
    observer.deliver(60)
    observer.pending.push({ duration: 70 }) // buffered but not yet delivered
    const first = await harness.exportResult()
    expect(first.samples.find((s) => s.name === 'longtask-duration')?.values).toEqual([60, 70])
    expect(observer.disconnected).toBe(false)
    observer.deliver(80)
    const second = await harness.exportResult()
    expect(second.samples.find((s) => s.name === 'longtask-duration')?.values).toEqual([60, 70, 80])
    expect(first.samples.find((s) => s.name === 'longtask-duration')?.values).toEqual([60, 70])
  })

  it('stop then start resumes collection', async () => {
    const ctor = fakeObserverClass(['longtask'])
    const harness = makeHarness({ performanceObserver: ctor })
    harness.startLongTaskCollection()
    FakeObserver.instances[0].deliver(55)
    harness.stopLongTaskCollection()
    expect(FakeObserver.instances[0].disconnected).toBe(true)
    expect(harness.startLongTaskCollection()).toBe(true)
    FakeObserver.instances[1].deliver(65)
    const result = await harness.exportResult()
    expect(result.samples.find((s) => s.name === 'longtask-duration')?.values).toEqual([55, 65])
  })
})

describe('classifyBrowser (UA fixtures)', () => {
  const UA = {
    safariMac: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Safari/605.1.15',
    safariIpad: 'Mozilla/5.0 (iPad; CPU OS 17_5 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Mobile/15E148 Safari/604.1',
    chromeIos: 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_5 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) CriOS/126.0.6478.153 Mobile/15E148 Safari/604.1',
    firefoxIos: 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_5 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) FxiOS/127.0 Mobile/15E148 Safari/605.1.15',
    edgeIos: 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_5 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 EdgiOS/126.2592.88 Mobile/15E148 Safari/605.1.15',
    playwrightWebkit: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 15_4) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/26.0 Safari/605.1.15',
    chromeDesktop: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
    headlessChrome: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) HeadlessChrome/126.0.0.0 Safari/537.36',
    edgeDesktop: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0',
    firefoxDesktop: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:127.0) Gecko/20100101 Firefox/127.0',
  }

  it('labels a real, non-automated Safari as safari (desktop and iPad)', () => {
    expect(classifyBrowser(UA.safariMac, false)).toEqual({ name: 'safari', engine: 'webkit', version: '17.5' })
    expect(classifyBrowser(UA.safariIpad, false)).toEqual({ name: 'safari', engine: 'webkit', version: '17.5' })
  })

  it('never calls iOS Chrome/Firefox/Edge safari (WebKit engine, generic label)', () => {
    expect(classifyBrowser(UA.chromeIos, false)).toEqual({ name: 'webkit', engine: 'webkit', version: '126.0.6478.153' })
    expect(classifyBrowser(UA.firefoxIos, false)).toEqual({ name: 'webkit', engine: 'webkit', version: '127.0' })
    expect(classifyBrowser(UA.edgeIos, false)).toMatchObject({ name: 'webkit', engine: 'webkit' })
  })

  it('never calls an automated (navigator.webdriver) WebKit safari', () => {
    expect(classifyBrowser(UA.playwrightWebkit, true)).toMatchObject({ name: 'webkit', engine: 'webkit' })
  })

  it('classifies Blink and Gecko browsers', () => {
    expect(classifyBrowser(UA.chromeDesktop, false)).toEqual({ name: 'chromium', engine: 'blink', version: '126.0.0.0' })
    expect(classifyBrowser(UA.headlessChrome, true)).toMatchObject({ name: 'chromium', engine: 'blink' })
    expect(classifyBrowser(UA.edgeDesktop, false)).toMatchObject({ name: 'chromium', engine: 'blink' })
    expect(classifyBrowser(UA.firefoxDesktop, false)).toEqual({ name: 'firefox', engine: 'gecko', version: '127.0' })
  })

  it('labels unrecognized agents unknown', () => {
    expect(classifyBrowser('curl/8.0', false)).toEqual({ name: 'unknown', engine: 'unknown', version: 'unknown' })
  })
})

describe('parseHarnessParams', () => {
  it('parses defaults and explicit values', () => {
    expect(parseHarnessParams('')).toMatchObject({ fixture: 'F1', profile: 'desktop', mode: 'smoke', cache: 'cold', token: null, commit: 'unknown' })
    expect(parseHarnessParams('?fixture=F3&profile=mobile&mode=release&cache=warm&token=abc&commit=deadbeef&deviceModel=iPad13,4&osVersion=17.5&thermal=nominal&battery=80%25')).toEqual({
      fixture: 'F3',
      profile: 'mobile',
      mode: 'release',
      cache: 'warm',
      token: 'abc',
      commit: 'deadbeef',
      operator: { deviceModel: 'iPad13,4', osVersion: '17.5', thermalState: 'nominal', batteryState: '80%' },
    })
  })

  it('refuses invalid values instead of silently falling back to a default', () => {
    for (const bad of ['?fixture=F9', '?profile=tablet', '?mode=debug', '?cache=lukewarm']) {
      expect(() => parseHarnessParams(bad), bad).toThrow(/must be/)
    }
  })

  it('ignores a collect URL entirely (same-origin posting only)', () => {
    const params = parseHarnessParams('?collect=https://evil.example/steal')
    expect(JSON.stringify(params)).not.toContain('evil')
  })
})

describe('gatherEnvironment', () => {
  const baseContext: BrowserContext = {
    userAgent: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Safari/605.1.15',
    webdriver: false,
    uaDataPlatform: null,
    navigatorPlatform: 'MacIntel',
    maxTouchPoints: 0,
    hardwareConcurrency: 8,
    deviceMemory: null,
    dpr: 2,
    innerWidth: 1440,
    innerHeight: 900,
    fontsStatus: 'loaded',
    gpu: null,
  }
  const placeholder = { adapter: 'xlsx-placeholder', bytes: 5000, rows: 120, cols: 10 }
  const params = (search = '') => parseHarnessParams(search)

  it('labels placeholder content scaffold-placeholder with no spec fixture id', () => {
    const env = gatherEnvironment(baseContext, params('?fixture=F2'), placeholder)
    expect(env).toMatchObject({ requestedFixtureId: 'F2', fixtureId: null, evidence: 'scaffold-placeholder' })
    expect(env.source).toEqual(placeholder)
  })

  it('carries the fixture id only for real fixture content', () => {
    const env = gatherEnvironment(baseContext, params('?fixture=F2'), { adapter: 'xlsx', bytes: 1, rows: null, cols: null })
    expect(env).toMatchObject({ requestedFixtureId: 'F2', fixtureId: 'F2' })
  })

  it('does not claim a fresh browser context (the page cannot verify it)', () => {
    const env = gatherEnvironment(baseContext, params(), placeholder)
    expect(env.cache.state).toBe('cold')
    expect(env.cache.note).not.toMatch(/fresh/i)
    expect(env.cache.note).toMatch(/unverified|cannot verify|requested/i)
  })

  it('classifies a non-automated Safari as safari/none and records the profile conformance', () => {
    const env = gatherEnvironment(baseContext, params(), placeholder)
    expect(env.browser).toMatchObject({ name: 'safari', engine: 'webkit', automation: 'none' })
    expect(env.profileConformance).toEqual({ conforms: true, deviations: [] })
    const small = gatherEnvironment({ ...baseContext, innerWidth: 800, innerHeight: 600, dpr: 1 }, params(), placeholder)
    expect(small.profileConformance.conforms).toBe(false)
    expect(small.profileConformance.deviations).toHaveLength(2)
  })

  it('page-level WebKit under automation is webkit/webdriver, never safari', () => {
    const env = gatherEnvironment({ ...baseContext, webdriver: true }, params(), placeholder)
    expect(env.browser).toMatchObject({ name: 'webkit', automation: 'webdriver' })
  })

  it('derives the OS from userAgentData, then navigator.platform, then the UA, with a reason when unknown', () => {
    expect(gatherEnvironment({ ...baseContext, uaDataPlatform: 'macOS' }, params(), placeholder).os).toMatchObject({ platform: 'macOS', platformNote: null })
    const viaNavigator = gatherEnvironment(baseContext, params(), placeholder).os
    expect(viaNavigator.platform).toBe('macOS')
    expect(viaNavigator.platformNote).toMatch(/navigator\.platform/)
    const ipadDesktopMode = gatherEnvironment({ ...baseContext, maxTouchPoints: 5 }, params(), placeholder).os
    expect(ipadDesktopMode.platform).toBe('iPadOS')
    const unknown = gatherEnvironment({ ...baseContext, userAgent: 'x', navigatorPlatform: null }, params(), placeholder).os
    expect(unknown.platform).toBe('unknown')
    expect(unknown.platformNote).toMatch(/.+/)
  })

  it('records operator metadata from the URL', () => {
    const env = gatherEnvironment(baseContext, params('?deviceModel=iPad13,4&osVersion=17.5&thermal=nominal&battery=80%25'), placeholder)
    expect(env.operator).toEqual({ deviceModel: 'iPad13,4', osVersion: '17.5', thermalState: 'nominal', batteryState: '80%' })
  })

  it('always produces an environment the ingest validator accepts', async () => {
    const harness = makeHarness({ environment: (fixture) => gatherEnvironment(baseContext, params('?fixture=F4'), fixture) })
    harness.setFixture(placeholder, null)
    const result = await harness.exportResult()
    await expect(verifyHarnessResult(result)).resolves.toEqual({ ok: true, errors: [] })
  })

  it('release mode on placeholder content fails to export instead of producing evidence', async () => {
    const harness = makeHarness({ environment: (fixture) => gatherEnvironment(baseContext, params('?mode=release'), fixture) })
    harness.setFixture(placeholder, null)
    await expect(harness.exportResult()).rejects.toThrow(/release/)
  })
})

describe('postResult (same-origin collector only)', () => {
  const result = { hash: 'x' } as never
  const location = (search: string) => ({ href: `http://192.168.1.20:4173/src/debug/interactive-viewer.html${search}`, search })

  it('does nothing without a token', async () => {
    let called = false
    const outcome = await postResult(result, location(''), (async () => { called = true; return new Response('{}') }) as typeof fetch)
    expect(outcome).toEqual({ status: 'no-collector' })
    expect(called).toBe(false)
  })

  it('posts to the fixed same-origin path with the token header, and ignores any ?collect= URL', async () => {
    const seen: Array<{ url: string; init: RequestInit }> = []
    const fetchImpl = (async (url: string, init: RequestInit) => {
      seen.push({ url: String(url), init })
      return new Response(JSON.stringify({ ok: true }), { status: 200 })
    }) as typeof fetch
    const outcome = await postResult(result, location('?token=s3cret&collect=https%3A%2F%2Fevil.example%2Fx'), fetchImpl)
    expect(outcome).toEqual({ status: 'posted' })
    expect(seen).toHaveLength(1)
    expect(new URL(seen[0].url).origin).toBe('http://192.168.1.20:4173')
    expect(new URL(seen[0].url).pathname).toBe(COLLECT_PATH)
    expect((seen[0].init.headers as Record<string, string>)['x-harness-token']).toBe('s3cret')
    expect(seen[0].init.method).toBe('POST')
  })

  it('surfaces the collector rejection reasons', async () => {
    const fetchImpl = (async () => new Response(JSON.stringify({ ok: false, errors: ['bad'] }), { status: 400 })) as typeof fetch
    await expect(postResult(result, location('?token=t'), fetchImpl)).resolves.toEqual({ status: 'rejected', errors: ['bad'] })
  })
})

describe('placeholder fixture', () => {
  it('builds a real XLSX zip and records adapter xlsx-placeholder', async () => {
    const { info, grid } = await fixtureLoader.loadPlaceholder()
    expect(info).toMatchObject({ adapter: 'xlsx-placeholder', rows: 120, cols: 10 })
    expect(info.bytes).toBeGreaterThan(1000)
    expect(grid?.length).toBe(120)
  })
})
