// @vitest-environment node
/**
 * S10.2 review finding 2: a cached package must never be reported as the
 * full-scale fixture it is named after. Cached bytes are verified against the
 * committed manifest (content SHA-256 over inflated entries, expanded bytes); a mismatch is
 * regenerated, and counts come from the manifest entry the bytes matched.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { FIXTURE_MANIFEST_PATH, ensureViewerFixtures, type FixtureManifest } from '../scripts/harness/viewerFixtures'
import { xlsxBytes } from './helpers/xlsx-package'

const manifest = JSON.parse(readFileSync(join(process.cwd(), FIXTURE_MANIFEST_PATH), 'utf8')) as FixtureManifest
let dir = ''
afterEach(() => { if (dir) rmSync(dir, { recursive: true, force: true }); delete process.env.OFFICEVIEW_VIEWER_FIXTURE_DIR })
const isolate = () => { dir = mkdtempSync(join(tmpdir(), 'ovf-')); process.env.OFFICEVIEW_VIEWER_FIXTURE_DIR = dir; return dir }
const oneCell = () => xlsxBytes({ sheetData: '<row r="1"><c r="A1"><v>1</v></c></row>' })

describe('viewer fixture identity', () => {
  it('a planted one-cell f2-full.xlsx is replaced by the verified 1M-row package, never reported as it', async () => {
    const d = isolate()
    writeFileSync(join(d, 'f2-full.xlsx'), oneCell())
    const out = await ensureViewerFixtures('full', manifest, ['F2'])
    const entry = manifest.scales.full.F2
    expect(out.F2.contentSha256).toBe(entry.contentSha256)
    expect(out.F2.nnz).toBe(8_238_808)
    expect([out.F2.rows, out.F2.cols]).toEqual([1_048_576, 8])
    // The file this run opens is the regenerated package, not the planted one.
    expect(readFileSync(join(d, 'f2-full.xlsx')).length).toBe(out.F2.bytes)
    expect(out.F2.bytes).toBeGreaterThan(40_000_000)
  }, 120_000)

  it('refuses when the generator no longer produces the manifest bytes', async () => {
    isolate()
    const stale: FixtureManifest = JSON.parse(JSON.stringify(manifest))
    stale.scales.reduced.F1.contentSha256 = '0'.repeat(64)
    await expect(ensureViewerFixtures('reduced', stale, ['F1'])).rejects.toThrow(/does not match manifest/)
  })

  it('refuses a manifest written for another cache version', async () => {
    isolate()
    await expect(ensureViewerFixtures('reduced', { ...manifest, cacheVersion: 999 }, ['F1'])).rejects.toThrow(/cache version/)
  })

  it('reuses a cached package only when every identity field matches', async () => {
    const d = isolate()
    const first = await ensureViewerFixtures('reduced', manifest, ['styled'])
    const again = await ensureViewerFixtures('reduced', manifest, ['styled'])
    expect(again.styled).toEqual(first.styled)
    expect(again.styled.contentSha256).toBe(manifest.scales.reduced.styled.contentSha256)
    expect(readFileSync(join(d, manifest.scales.reduced.styled.file)).length).toBe(again.styled.bytes)
  })
})
