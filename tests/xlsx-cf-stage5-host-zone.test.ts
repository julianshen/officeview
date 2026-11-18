// @vitest-environment node
/**
 * Stage 5 final review F2: the Stage 5 validation of a clock case is independent of the host's time zone and clock. A clock
 * case was measured by a viewer in its `requiredTimeZone` on its anchor day; the audit rows, the counterfactual and the
 * conformance/mutant verdict helper must all evaluate it there and then, whatever zone and day the host is in (CI runs on
 * UTC Linux, on any day).
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { caseClockOf, caseDocumentEvaluationOf, withCaseZone } from '../scripts/cf-refs/case-clock'
import { cfDocumentEvaluationOf } from '../src/xlsx/cf/formula'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const manifest = () => JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as import('../scripts/cf-refs/cases').CfManifest
const clockCases = () => manifest().cases.filter(c => c.stage === 5 && c.status === 'measured' && c.clockRequirement)
/** A host day far from every anchor: a tool that reads the host clock instead of the anchor names the wrong day. */
const HOST_DAY = new Date('2027-03-01T12:00:00Z')

const previousTz = process.env.TZ
afterEach(() => {
  vi.useRealTimers()
  if (previousTz === undefined) delete process.env.TZ; else process.env.TZ = previousTz
})

describe('the Stage 5 clock cases validate in their own zone and day, whatever the host (final review F2)', () => {
  it('the case clock is the anchor instant in the required zone', () => {
    const c = clockCases().find(x => x.id === 'stage5-clock-today')!
    const clock = caseClockOf(c)!
    expect(clock.timeZone).toBe('Asia/Taipei')
    expect(clock.now()).toBe(Date.parse(`${c.clockRequirement!.anchorLocalDate}T04:00:00+08:00`))
    expect(caseClockOf({})).toBeUndefined()
  })

  it('a document parsed under another zone is evaluated in the case zone', async () => {
    process.env.TZ = 'America/Los_Angeles'
    const c = clockCases()[0]!
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))))
    expect(doc.semantics?.timeZone).toBe('America/Los_Angeles')
    expect(caseDocumentEvaluationOf(c, doc)).toEqual({ ...cfDocumentEvaluationOf(doc), timeZone: 'Asia/Taipei' })
    expect(caseDocumentEvaluationOf({}, doc)).toBe(cfDocumentEvaluationOf(doc))
  })

  it('withCaseZone sets the ambient zone for the case and restores it', async () => {
    process.env.TZ = 'UTC'
    const c = clockCases()[0]!
    const inside = await withCaseZone(c, async () => Intl.DateTimeFormat().resolvedOptions().timeZone)
    expect(inside).toBe('Asia/Taipei')
    expect(Intl.DateTimeFormat().resolvedOptions().timeZone).toBe('UTC')
    expect(await withCaseZone({}, async () => Intl.DateTimeFormat().resolvedOptions().timeZone)).toBe('UTC')
  })

  for (const zone of ['UTC', 'America/Los_Angeles']) {
    it(`under TZ=${zone} on another host day, the audit rows and the verdict helper agree with Excel on every clock-case cell`, async () => {
      process.env.TZ = zone
      vi.useFakeTimers({ toFake: ['Date'] })
      vi.setSystemTime(HOST_DAY)
      const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
      const { stage5SlotVerdict } = await import('./helpers/cf-stage5-verdict')
      const cases = clockCases()
      expect(cases.map(c => c.id)).toContain('stage5-clock-today')
      const rows = await stage5CaseRecords({ ...manifest(), cases }, REPO_ROOT)
      expect(rows.filter(r => r.compared).length).toBeGreaterThan(0)
      expect(rows.filter(r => r.compared && !r.agrees).map(r => `${r.caseId}:${r.ref} excel=${JSON.stringify(r.excel)} production=${JSON.stringify(r.production)}`)).toEqual([])
      const violations: string[] = []
      for (const c of cases) {
        // The caller parses under the host zone, as the conformance and mutant tests do.
        const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))))
        const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
        for (const e of c.expected.cells) { const v = await stage5SlotVerdict(c, doc, sheet, e.ref, REPO_ROOT); if (v) violations.push(v) }
      }
      expect(violations).toEqual([])
    }, 300000)
  }
})
