import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import {
  auditStage3Manifest,
  BAR_LAYOUT_SETTLING,
  deriveBarLayout,
  renderBarLayoutBlock,
  renderBarLengthsBlock,
  renderVerifiedCapabilitiesSource,
} from '../scripts/cf-refs/audit'
import type { CfManifest } from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import { barLayoutEvidence, injectCfBarLayout, withExactCfEvidence } from '../src/xlsx/cf/evidence'

// Ruling S3-B: exactly one layout variant is enabled, by bar:lengths-* and bar:axis-* evidence. The harness derives it
// from the hypotheses each verified capture matched, and production reads it only from the published export.

const REPO_ROOT = resolve(__dirname, '..')
const AUDIT_TIMEOUT_MS = 30_000
const manifest = (): CfManifest => JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
const SPAN_CELL = { origin: 'span', lengthBasis: 'cell' } as const
const AXIS_SIDE = { origin: 'axis', lengthBasis: 'side' } as const

afterEach(() => injectCfBarLayout(null))

describe('Stage 3 bar layout variant (Ruling S3-B)', () => {
  it('the current captures select span/cell for all-positive scopes and axis/side for mixed-sign scopes', () => {
    const report = auditStage3Manifest(manifest(), { rootDir: REPO_ROOT })
    expect(report.barLayout.layout).toEqual({
      allPositive: { ...SPAN_CELL, settledBy: ['bar:lengths-default-10-90', 'bar:lengths-explicit-0-100'] },
      mixedSign: { ...AXIS_SIDE, settledBy: ['bar:axis-automatic', 'bar:axis-middle'] },
    })
    expect(report.barLayout.findings).toEqual([])
  }, AUDIT_TIMEOUT_MS)

  it('Task 22b: capture run 7 verifies bar:lengths-default-10-90 and exports the matched rule; the block renders idempotently', () => {
    const report = auditStage3Manifest(manifest(), { rootDir: REPO_ROOT })
    expect(report.perCapability.find(p => p.capabilityId === 'bar:lengths-default-10-90')?.disposition).toBe('verified')
    // stage3-bar-lengths-default (run 7): half-min on the x14 form, the linear reading on base-only bars.
    // Task 22c (review I2): with the pairing form x fill type combinations the capture measured (no base-only solid).
    expect(report.barLengths).toEqual({ 'default-10-90': { rule: 'half-min@x14', measured: ['base-only/gradient', 'x14/gradient', 'x14/solid'], settledBy: ['bar:lengths-default-10-90'] } })
    expect(renderBarLengthsBlock({})).toBe('export const VERIFIED_STAGE3_BAR_LENGTHS: Readonly<Record<string, { rule: string; measured: readonly string[]; settledBy: readonly string[] }>> = {\n}\n')
    expect(renderBarLengthsBlock({ 'default-10-90': { rule: 'half-min@x14', measured: ['x14/solid'], settledBy: ['bar:lengths-default-10-90'] } }))
      .toBe("export const VERIFIED_STAGE3_BAR_LENGTHS: Readonly<Record<string, { rule: string; measured: readonly string[]; settledBy: readonly string[] }>> = {\n  'default-10-90': { rule: 'half-min@x14', measured: ['x14/solid'], settledBy: ['bar:lengths-default-10-90'] },\n}\n")
    // Every candidate rule keeps the span/cell layout: the rule changes lengths, not where a bar starts.
    for (const v of Object.values(BAR_LAYOUT_SETTLING['bar:lengths-default-10-90']!.variants)) expect(v).toEqual(SPAN_CELL)
  }, AUDIT_TIMEOUT_MS)

  it('derivation pools every case, and disagreement or an unmapped match selects no variant', () => {
    const verified = (...ids: string[]) => ids.map(capabilityId => ({ capabilityId, disposition: 'verified' as const }))
    // Agreement across two cases.
    const ok = deriveBarLayout(verified('bar:lengths-explicit-0-100', 'bar:axis-automatic'),
      [{ 'bar:lengths-explicit-0-100': 'effective' }, { 'bar:axis-automatic': 'automatic', 'bar:lengths-explicit-0-100': 'effective' }])
    expect(ok.layout).toEqual({ allPositive: { ...SPAN_CELL, settledBy: ['bar:lengths-explicit-0-100'] }, mixedSign: { ...AXIS_SIDE, settledBy: ['bar:axis-automatic'] } })
    // A second case matching a different hypothesis: no variant for that scope, with a finding.
    const split = deriveBarLayout(verified('bar:lengths-explicit-0-100', 'bar:axis-automatic'),
      [{ 'bar:lengths-explicit-0-100': 'effective', 'bar:axis-automatic': 'automatic' }, { 'bar:lengths-explicit-0-100': 'base' }])
    expect(split.layout.allPositive).toBeUndefined()
    expect(split.layout.mixedSign).toEqual({ ...AXIS_SIDE, settledBy: ['bar:axis-automatic'] })
    expect(split.findings.join('\n')).toMatch(/allPositive/)
    // Nothing verified: nothing selected.
    expect(deriveBarLayout([{ capabilityId: 'bar:axis-automatic', disposition: 'supported-structural' }], [{ 'bar:axis-automatic': 'automatic' }]).layout).toEqual({})
    // Every settling capability names its scope and only the hypotheses the predictor's layout embodies.
    for (const [cap, s] of Object.entries(BAR_LAYOUT_SETTLING)) {
      expect(cap).toMatch(/^bar:(lengths|axis)-/)
      expect(['allPositive', 'mixedSign']).toContain(s.scope)
    }
  })

  it('rendering without --publish-stage3 is byte-identical to the committed file; publishing carries the layout export', () => {
    const committed = readFileSync(resolve(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts'), 'utf8')
    expect(renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { committedSource: committed })).toBe(committed)
    const published = renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { publishStage3: true })
    const block = renderBarLayoutBlock(auditStage3Manifest(manifest(), { rootDir: REPO_ROOT }).barLayout.layout)
    expect(published).toContain(block)
    // Task 23 published it: the committed file is exactly the published rendering.
    expect(committed).toBe(published)
    expect(block).toContain("allPositive: { origin: 'span', lengthBasis: 'cell', settledBy: ['bar:lengths-default-10-90', 'bar:lengths-explicit-0-100'] }")
    // A published block is carried over verbatim by a later --write without --publish-stage3, and is idempotent.
    expect(renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { committedSource: published })).toBe(published)
    const edited = published.replace(block, renderBarLayoutBlock({ mixedSign: { ...AXIS_SIDE, settledBy: ['bar:axis-middle'] } }))
    expect(edited).not.toBe(published)
    expect(renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { committedSource: edited })).toBe(edited)
  }, AUDIT_TIMEOUT_MS)

  it('production resolves a variant only from the export and only while its settling capabilities are verified', () => {
    // The published layout (Task 23): both scopes under production evidence; a scope whose settling capability is not
    // verified drops out (allPositive is also settled by bar:lengths-default-10-90).
    expect(barLayoutEvidence()).toEqual({ allPositive: SPAN_CELL, mixedSign: AXIS_SIDE })
    withExactCfEvidence(['bar:lengths-explicit-0-100', 'bar:axis-automatic', 'bar:axis-middle'], () => {
      expect(barLayoutEvidence()).toEqual({ mixedSign: AXIS_SIDE })
    })
    withExactCfEvidence([], () => expect(barLayoutEvidence()).toBeUndefined())
    injectCfBarLayout({
      allPositive: { ...SPAN_CELL, settledBy: ['bar:lengths-explicit-0-100'] },
      mixedSign: { ...AXIS_SIDE, settledBy: ['bar:axis-automatic', 'bar:axis-middle'] },
    })
    withExactCfEvidence(['bar:lengths-explicit-0-100', 'bar:axis-automatic', 'bar:axis-middle'], () => {
      expect(barLayoutEvidence()).toEqual({ allPositive: SPAN_CELL, mixedSign: AXIS_SIDE })
    })
    // A settling capability missing from the verified set drops that scope's variant.
    withExactCfEvidence(['bar:lengths-explicit-0-100', 'bar:axis-automatic'], () => {
      expect(barLayoutEvidence()).toEqual({ allPositive: SPAN_CELL })
    })
    withExactCfEvidence([], () => expect(barLayoutEvidence()).toBeUndefined())
    // An explicit gate is honoured (the service passes its own evidence gate).
    expect(barLayoutEvidence(cap => cap === 'bar:lengths-explicit-0-100')).toEqual({ allPositive: SPAN_CELL })
    // Malformed exports never yield a variant.
    for (const bad of [
      { allPositive: { origin: 'edge', lengthBasis: 'cell', settledBy: ['bar:lengths-explicit-0-100'] } },
      { allPositive: { ...SPAN_CELL, settledBy: [] } },
      { allPositive: { ...SPAN_CELL } },
    ]) {
      injectCfBarLayout(bad as never)
      expect(barLayoutEvidence(() => true)).toBeUndefined()
    }
  })
})
