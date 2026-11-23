// @vitest-environment node
/**
 * cf-compare reports the interactive text scope's whole-request refusal (its O(rules x cells) estimate over the scratch
 * bound, a deliberate admission policy pinned by the Stage 2-4 pressure audits) as recipe-pending, withheld fail-closed,
 * not as a failure. So that a case newly refused never hides as pending, the set of manifest cases whose compared range
 * the estimator refuses is pinned here: the two rule-heavy glyph-key sheets (76 rules over A1:K42), refused at the text
 * scope, and three Stage 5 sheets cf-compare already reports withheld at the service batch.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { estimateCfPrepareBytes } from '../src/xlsx/cf/service'
import { compareCfCase } from '../scripts/cf-compare'
import { parseRangeA1 } from '../scripts/cf-refs/verify'
import { caseParseOptionsOf } from '../scripts/cf-refs/case-clock'
import type { CfOracleCase } from '../scripts/cf-refs/cases'

const root = resolve(__dirname, '..')
const manifest = JSON.parse(readFileSync(resolve(root, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as { cases: CfOracleCase[] }
// Refused at the text scope (the service batches of 8 rows admit): the two rule-heavy glyph-key sheets.
const TEXT_SCOPE_REFUSED = ['stage4-icon-glyph-key', 'stage4-icon-glyph-key-centred']
// Refused earlier, at the service batch (cf-compare already reports these withheld): three Stage 5 sheets.
const SERVICE_REFUSED = ['stage5-cfvo-relative-a1', 'stage5-cfvo-scope-v2', 'stage5-formula-truthiness']
const REFUSED = [...TEXT_SCOPE_REFUSED, ...SERVICE_REFUSED].sort()

describe('cf-compare: interactive text-scope refusal is withheld, and pinned to the cases it covers', () => {
  it('the estimator refuses the compared range of exactly these five cases', async () => {
    const refused: string[] = []
    for (const c of manifest.cases) {
      if (c.status === 'unsupported') continue
      const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(root, c.saved)))), c.stage === 5 ? caseParseOptionsOf(c) : {})
      const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]
      const rules = sheet?.conditionalFormatting
      const rect = parseRangeA1(c.range)
      if (!rules || rules.rules.length === 0 || !rect) continue
      const bound = estimateCfPrepareBytes(rules, [{ firstRow: rect.minRow, lastRow: rect.maxRow, firstCol: rect.minCol, lastCol: rect.maxCol }])
      if (bound === undefined) refused.push(c.id)
    }
    expect(refused.sort()).toEqual(REFUSED)
  }, 300_000)

  for (const id of TEXT_SCOPE_REFUSED) {
    it(`${id} [interactive] is recipe-pending (withheld at the text scope), not fail`, async () => {
      const c = manifest.cases.find(e => e.id === id)!
      const res = await compareCfCase(c, 'interactive', { rootDir: root })
      expect(res.outcome, res.reason).toBe('recipe-pending')
      expect(res.reason).toMatch(/Interactive text scope refused at its admitted batch/)
    }, 300_000)
  }
})
