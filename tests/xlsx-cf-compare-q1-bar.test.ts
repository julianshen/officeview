// @vitest-environment node
/**
 * Codex pre-capture review P2 (the all-equal follow-up): cf-compare's interactive adapter counts a bar with no descriptor
 * as applied ONLY when the presentation path itself paints it (`barAt` of the resident text scope, which applies the
 * painter's own withholds: merged hosts, paint gates, `barWithheld`) and it carries no `valueFraction` (lo = hi, owner
 * ruling Q1). A ready service outcome alone never counts.
 */
import { readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { q1BarPainted } from '../scripts/cf-compare'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { prepareCfTextScope } from '../src/xlsx/cf/text'
import { VERIFIED_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { CfBarDescriptor } from '../src/xlsx/cf/model'

const REPO_ROOT = resolve(__dirname, '..')
const budget = { reserve() {}, release() {}, retain: (_t: unknown, _b: number, commit: () => void) => commit(), settle() {} }
const Q1_BAR: CfBarDescriptor = { ruleId: 0, axisFraction: 0, startFraction: 0, endFraction: 0.5, fillColor: 0xff0000ff, isGradient: false, isNegative: false, showValue: true, direction: 'leftToRight' }

async function degenerateScope() {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/stage3-bar-degenerate.xlsx')))))
  return prepareCfTextScope(doc.sheets[0]! as never, { rects: [{ r0: 0, r1: 12, c0: 0, c1: 8 }], signal: new AbortController().signal, locale: 'en-US', budget: budget as never })
}

describe('cf-compare interactive: a descriptor-less bar is applied only where the presentation path paints it', () => {
  it('stage3-bar-degenerate B2 (the Q1 bar, lo = hi) is painted by the text scope and counts', async () => {
    const scope = (await degenerateScope())!
    try {
      expect(scope.describe(1, 1)).toBeUndefined()
      expect(q1BarPainted(scope, 1, 1)).toBe(true)
      // A painted bar WITH a value position (D3, inverted bounds are withheld; use the non-degenerate G column) never needs it.
      expect(q1BarPainted(scope, 2, 6)).toBe(false)
    } finally { scope.dispose() }
  }, 300_000)

  it('a ready bar the painter withholds (a paint gate not held) does not count', async () => {
    await withExactCfEvidence([...VERIFIED_CAPABILITIES].filter(c => c !== 'bar:geometry-insets'), async () => {
      const scope = (await degenerateScope())!
      try {
        expect(scope.barAt?.(1, 1)).toBeUndefined()
        expect(q1BarPainted(scope, 1, 1)).toBe(false)
      } finally { scope.dispose() }
    })
  }, 300_000)

  it('a bar withheld at a merged host (barAt reports none, the service outcome still carries it) does not count', () => {
    const merged = { barAt: (_r: number, _c: number): CfBarDescriptor | undefined => undefined }
    expect(q1BarPainted(merged, 1, 1)).toBe(false)
    expect(q1BarPainted({ barAt: () => Q1_BAR }, 1, 1)).toBe(true)
    expect(q1BarPainted({ barAt: () => ({ ...Q1_BAR, valueFraction: 0.5 }) }, 1, 1)).toBe(false)
    expect(q1BarPainted(undefined, 1, 1)).toBe(false)
    expect(q1BarPainted({}, 1, 1)).toBe(false)
  })
})
