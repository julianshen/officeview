// @vitest-environment node
/**
 * The Stage 5 publication gate (`audit.ts --write --publish-stage5`) judges the records under the evidence about to be
 * published. Review P2 (percent-order full review, re-review): the set is taken structured from the renderer, never re-parsed
 * from the rendered text (a missed parse became an empty set, under which withheld visuals are not compared), and it may be
 * smaller than the committed set when the audit revokes an id.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { renderVerifiedCapabilities, renderVerifiedCapabilitiesSource } from '../scripts/cf-refs/audit'
import { VERIFIED_CAPABILITIES } from '../src/xlsx/cf/verifiedCapabilities'
import type { CfManifest } from '../scripts/cf-refs/cases'

const root = resolve(__dirname, '..')
const manifest = JSON.parse(readFileSync(resolve(root, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
const committedSource = readFileSync(resolve(root, 'src/xlsx/cf/verifiedCapabilities.ts'), 'utf8')
const options = { committedSource, publishStage3: true, publishStage5: true }

describe('audit: the pending Stage 5 evidence is the renderer\'s own set', () => {
  it('the structured set is exactly the VERIFIED_CAPABILITIES the source declares', () => {
    const { source, allVerified } = renderVerifiedCapabilities(manifest, root, options)
    const declared = /export const VERIFIED_CAPABILITIES: ReadonlySet<string> = new Set\(\[([\s\S]*?)\]\)/.exec(source)![1]!.match(/'[^']+'/g)!.map(x => x.slice(1, -1))
    expect([...allVerified]).toEqual(declared)
    expect(allVerified.length).toBeGreaterThan(0)
  }, 600_000)

  it('the source wrapper is unchanged, and at HEAD the rendered set equals the committed set', () => {
    const { source, allVerified } = renderVerifiedCapabilities(manifest, root, options)
    expect(renderVerifiedCapabilitiesSource(manifest, root, options)).toBe(source)
    expect(new Set(allVerified)).toEqual(VERIFIED_CAPABILITIES)
  }, 600_000)
})
