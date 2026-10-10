/**
 * Browser classification and environment gathering for the harness page.
 *
 * Labels are conservative: `safari` is claimed only for a real, non-automated
 * Safari. Every other WebKit (iOS Chrome/Firefox/Edge, Playwright WebKit, an
 * embedded WKWebView) is the generic `webkit`; the Playwright driver later
 * re-tags its own runs `webkit-playwright`.
 */
import { FIXTURE_IDS, PLACEHOLDER_ADAPTERS, checkProfileConformance, deriveEvidence } from './result'
import type { FixtureInfo } from './fixture'
import type { BrowserEngine, BrowserName, FixtureId, HarnessEnvironment, HarnessMode, HarnessProfile } from './types'

export interface BrowserClassification {
  name: BrowserName
  engine: BrowserEngine
  version: string
}

export function classifyBrowser(userAgent: string, webdriver: boolean): BrowserClassification {
  // iOS wrappers are WebKit whatever brand they carry (and EdgiOS also has Version/).
  const iosWrapper = /(?:CriOS|FxiOS|EdgiOS|OPiOS)\/([\d.]+)/.exec(userAgent)
  if (iosWrapper) return { name: 'webkit', engine: 'webkit', version: iosWrapper[1] }
  const blink = /(?:HeadlessChrome|Chrome|Chromium)\/([\d.]+)/.exec(userAgent) ?? /(?:Edg|OPR)\/([\d.]+)/.exec(userAgent)
  if (blink) return { name: 'chromium', engine: 'blink', version: blink[1] }
  const gecko = /Firefox\/([\d.]+)/.exec(userAgent)
  if (gecko) return { name: 'firefox', engine: 'gecko', version: gecko[1] }
  if (/AppleWebKit\//.test(userAgent)) {
    const version = /Version\/([\d.]+)/.exec(userAgent)?.[1]
    const isSafari = version !== undefined && /Safari\//.test(userAgent) && !webdriver
    return { name: isSafari ? 'safari' : 'webkit', engine: 'webkit', version: version ?? 'unknown' }
  }
  return { name: 'unknown', engine: 'unknown', version: 'unknown' }
}

export interface HarnessParams {
  commit: string
  fixture: FixtureId
  profile: HarnessProfile
  mode: HarnessMode
  cache: 'cold' | 'warm'
  token: string | null
  operator: HarnessEnvironment['operator']
}

function optionalText(params: URLSearchParams, key: string): string | null {
  const value = params.get(key)
  return value === null || value === '' ? null : value
}

/** Strict parsing: an invalid value is an error, never a silent default. */
export function parseHarnessParams(search: string): HarnessParams {
  const params = new URLSearchParams(search)
  const fixture = params.get('fixture') ?? 'F1'
  if (!FIXTURE_IDS.includes(fixture as FixtureId)) throw new Error(`fixture must be one of ${FIXTURE_IDS.join(', ')}, got ${fixture}`)
  const profile = params.get('profile') ?? 'desktop'
  if (profile !== 'desktop' && profile !== 'mobile') throw new Error(`profile must be desktop|mobile, got ${profile}`)
  const mode = params.get('mode') ?? 'smoke'
  if (mode !== 'smoke' && mode !== 'release') throw new Error(`mode must be smoke|release, got ${mode}`)
  const cache = params.get('cache') ?? 'cold'
  if (cache !== 'cold' && cache !== 'warm') throw new Error(`cache must be cold|warm, got ${cache}`)
  return {
    commit: params.get('commit') ?? 'unknown',
    fixture: fixture as FixtureId,
    profile,
    mode,
    cache,
    token: optionalText(params, 'token'),
    operator: {
      deviceModel: optionalText(params, 'deviceModel'),
      osVersion: optionalText(params, 'osVersion'),
      thermalState: optionalText(params, 'thermal'),
      batteryState: optionalText(params, 'battery'),
    },
  }
}

/** Raw browser facts, separated from `window` so tests can supply them. */
export interface BrowserContext {
  userAgent: string
  webdriver: boolean
  uaDataPlatform: string | null
  navigatorPlatform: string | null
  maxTouchPoints: number
  hardwareConcurrency: number | null
  deviceMemory: number | null
  dpr: number
  innerWidth: number
  innerHeight: number
  fontsStatus: string | null
  gpu: string | null
}

function detectPlatform(ctx: BrowserContext): { platform: string; platformNote: string | null } {
  if (ctx.uaDataPlatform) return { platform: ctx.uaDataPlatform, platformNote: null }
  const touchMac = ctx.maxTouchPoints > 1 // iPadOS reports a Mac platform in desktop mode
  const fromNavigator = (platform: string): string | null => {
    if (/^Mac/.test(platform)) return touchMac ? 'iPadOS' : 'macOS'
    if (/^iPad/.test(platform)) return 'iPadOS'
    if (/^(iPhone|iPod)/.test(platform)) return 'iOS'
    if (/^Win/.test(platform)) return 'Windows'
    if (/Android/.test(platform)) return 'Android'
    if (/^Linux/.test(platform)) return /Android/.test(ctx.userAgent) ? 'Android' : 'Linux'
    return null
  }
  const viaNavigator = ctx.navigatorPlatform ? fromNavigator(ctx.navigatorPlatform) : null
  if (viaNavigator) return { platform: viaNavigator, platformNote: 'userAgentData unavailable (insecure origin or engine); derived from navigator.platform' }
  const ua = ctx.userAgent
  const viaUa = /iPad/.test(ua) ? 'iPadOS' : /iPhone/.test(ua) ? 'iOS' : /Android/.test(ua) ? 'Android' : /Windows/.test(ua) ? 'Windows' : /Macintosh/.test(ua) ? (touchMac ? 'iPadOS' : 'macOS') : /Linux|X11/.test(ua) ? 'Linux' : null
  if (viaUa) return { platform: viaUa, platformNote: 'userAgentData and navigator.platform unavailable; derived from the user agent' }
  return { platform: 'unknown', platformNote: 'userAgentData, navigator.platform and the user agent do not identify the OS' }
}

export function gatherEnvironment(ctx: BrowserContext, params: HarnessParams, fixture: FixtureInfo | null): HarnessEnvironment {
  const browser = classifyBrowser(ctx.userAgent, ctx.webdriver)
  const source = {
    adapter: fixture?.adapter ?? 'none',
    bytes: fixture?.bytes ?? 0,
    rows: fixture?.rows ?? null,
    cols: fixture?.cols ?? null,
  }
  const browserFacts: HarnessEnvironment['browser'] = {
    name: browser.name,
    version: browser.version,
    engine: browser.engine,
    userAgent: ctx.userAgent.slice(0, 512),
    automation: ctx.webdriver ? 'webdriver' : 'none',
  }
  const viewport = { width: ctx.innerWidth, height: ctx.innerHeight }
  const placeholder = PLACEHOLDER_ADAPTERS.includes(source.adapter)
  const os = detectPlatform(ctx)
  return {
    buildCommit: params.commit,
    hardware: { logicalCores: ctx.hardwareConcurrency, deviceMemoryGiB: ctx.deviceMemory, gpu: ctx.gpu },
    os: { ...os, userAgent: ctx.userAgent.slice(0, 512) },
    browser: browserFacts,
    dpr: ctx.dpr,
    viewport,
    profileConformance: checkProfileConformance(params.profile, viewport, ctx.dpr),
    fonts:
      ctx.fontsStatus === null
        ? { status: 'unavailable', resolved: false }
        : { status: ctx.fontsStatus === 'loaded' ? 'ready' : 'loading', resolved: ctx.fontsStatus === 'loaded' },
    cache: {
      state: params.cache,
      note: `${params.cache} state requested by the URL; the page cannot verify browser cache or process state`,
    },
    source,
    backend: 'canvas2d-placeholder',
    requestedFixtureId: params.fixture,
    fixtureId: placeholder ? null : params.fixture,
    evidence: deriveEvidence({ source, browser: browserFacts }),
    operator: params.operator,
    profile: params.profile,
    mode: params.mode,
    note: 'S0.3 scaffold: synthetic placeholder fixture rendered by a minimal canvas grid; interactive viewer kernels are not implemented yet',
  }
}

function detectGpu(): string | null {
  try {
    const canvas = document.createElement('canvas')
    const gl = canvas.getContext('webgl2') ?? canvas.getContext('webgl')
    if (!gl) return null
    const info = gl.getExtension('WEBGL_debug_renderer_info')
    const renderer = info ? gl.getParameter(info.UNMASKED_RENDERER_WEBGL) : gl.getParameter(gl.RENDERER)
    return typeof renderer === 'string' ? renderer.slice(0, 256) : null
  } catch {
    return null
  }
}

export function readBrowserContext(): BrowserContext {
  const nav = window.navigator as Navigator & {
    deviceMemory?: number
    userAgentData?: { platform?: string }
  }
  return {
    userAgent: nav.userAgent,
    webdriver: nav.webdriver === true,
    uaDataPlatform: nav.userAgentData?.platform || null,
    navigatorPlatform: nav.platform || null,
    maxTouchPoints: nav.maxTouchPoints ?? 0,
    hardwareConcurrency: nav.hardwareConcurrency || null,
    deviceMemory: nav.deviceMemory ?? null,
    dpr: window.devicePixelRatio,
    innerWidth: window.innerWidth,
    innerHeight: window.innerHeight,
    fontsStatus: document.fonts ? document.fonts.status : null,
    gpu: detectGpu(),
  }
}
