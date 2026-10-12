/**
 * Rewrites tests/fixtures/xlsx-viewer/viewer-perf-fixtures.json from a fresh
 * generation of every S10.2 viewer fixture (full and reduced scale). Run after
 * a deliberate generator change, with FIXTURE_CACHE_VERSION bumped:
 *   bun scripts/write-viewer-fixture-manifest.ts
 */
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { FIXTURE_MANIFEST_PATH, buildFixtureManifest } from './harness/viewerFixtures'

const dir = mkdtempSync(join(tmpdir(), 'officeview-fixture-manifest-'))
try {
  const manifest = await buildFixtureManifest(dir)
  writeFileSync(resolve(FIXTURE_MANIFEST_PATH), JSON.stringify(manifest, null, 2) + '\n')
  console.log(`wrote ${FIXTURE_MANIFEST_PATH}`)
} finally {
  rmSync(dir, { recursive: true, force: true })
}
