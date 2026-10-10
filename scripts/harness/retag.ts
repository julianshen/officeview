/**
 * Re-tag a page-exported result with facts only the driver knows (it launched
 * the browser), then recompute the integrity hash. Chromium keeps its own
 * label; WebKit is `webkit-playwright` and can never become `safari`.
 */
import { deriveEvidence, hashResult } from '../../src/debug/harness/result'
import type { HarnessResult, OperatorMetadata } from '../../src/debug/harness/types'

export interface DriverFacts {
  engine: 'chromium' | 'webkit'
  browserVersion: string
  /** Node's process.platform of the machine running the driver. */
  hostPlatform: string
  /** Commit label (sha, optionally -dirty) of the checkout under test. */
  commit: string
  operator: OperatorMetadata
}

const PLATFORM_LABEL: Record<string, string> = { darwin: 'macOS', win32: 'Windows', linux: 'Linux' }

export async function retagResult(result: HarnessResult, facts: DriverFacts): Promise<HarnessResult> {
  const { hash: _previousHash, ...rest } = result
  const environment = {
    ...rest.environment,
    buildCommit: facts.commit,
    browser: {
      name: facts.engine === 'webkit' ? ('webkit-playwright' as const) : ('chromium' as const),
      version: facts.browserVersion,
      engine: facts.engine === 'webkit' ? ('webkit' as const) : ('blink' as const),
      userAgent: rest.environment.browser.userAgent,
      automation: 'playwright' as const,
    },
    os: {
      platform: PLATFORM_LABEL[facts.hostPlatform] ?? facts.hostPlatform,
      platformNote: 'reported by the driver host (process.platform); the page user agent is not used for automated runs',
      userAgent: rest.environment.os.userAgent,
    },
    cache: { state: 'cold' as const, note: 'fresh Playwright browser context per run' },
    operator: facts.operator,
  }
  const tagged = { ...rest, environment: { ...environment, evidence: deriveEvidence(environment) } }
  return { ...tagged, hash: await hashResult(tagged) }
}
