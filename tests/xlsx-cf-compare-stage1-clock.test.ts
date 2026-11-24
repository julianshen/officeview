// @vitest-environment node
/**
 * cf-compare on a Stage 1 clock case evaluates it as the viewer Excel measured: on the anchor day, in the required zone,
 * whatever the host's day and zone. The resident (interactive) adapter parsed Stage 1-4 cases with no clock pin, so the
 * load-time recalculation of a `ca` TODAY() cell (stage1-timeperiod-today-clock B15, stage1-timeperiod-today-clock-2 B7)
 * read the host's day while the rules read the anchor serial through `dateContext`: from the day after the anchor both
 * cases failed "expected matched: true, but no CF formatting applied". Production parses and paints at one instant, so
 * the two always agree there; only the validation clock was split.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { compareCfCase } from '../scripts/cf-compare'
import type { CfOracleCase } from '../scripts/cf-refs/cases'

const root = resolve(__dirname, '..')
const manifest = JSON.parse(readFileSync(resolve(root, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as { cases: CfOracleCase[] }
/** A host day far from every anchor: an adapter that reads the host clock instead of the anchor names the wrong day. */
const HOST_DAY = new Date('2027-03-01T12:00:00Z')
const CASES = ['stage1-timeperiod-today-clock', 'stage1-timeperiod-today-clock-2']

const previousTz = process.env.TZ
afterEach(() => {
  vi.useRealTimers()
  if (previousTz === undefined) delete process.env.TZ; else process.env.TZ = previousTz
})

describe('cf-compare: Stage 1 clock cases validate on their anchor day, whatever the host', () => {
  it('both cases are measured Stage 1 clock cases', () => {
    for (const id of CASES) {
      const c = manifest.cases.find(e => e.id === id)!
      expect(c).toMatchObject({ stage: 1, status: 'measured' })
      expect(c.clockRequirement?.requiredTimeZone).toBe('Asia/Taipei')
    }
  })

  for (const zone of ['UTC', 'America/Los_Angeles']) {
    for (const id of CASES) {
      for (const mode of ['interactive', 'paged'] as const) {
        it(`${id} [${mode}] passes under TZ=${zone} on another host day`, async () => {
          process.env.TZ = zone
          vi.useFakeTimers({ toFake: ['Date'] })
          vi.setSystemTime(HOST_DAY)
          const c = manifest.cases.find(e => e.id === id)!
          const res = await compareCfCase(c, mode, { rootDir: root })
          expect(res.outcome, res.reason).toBe('pass')
        }, 300_000)
      }
    }
  }
})
