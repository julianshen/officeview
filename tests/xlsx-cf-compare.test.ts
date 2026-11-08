// @vitest-environment node
/**
 * Task 11 fix round 1, finding 4: interactive/paged compare modes must be
 * actual viewport adapters (run the data path, bind outputs to workbook and
 * oracle hashes), not unconditional stubs. Pixel diff stays pending without
 * Excel viewport refs, but the adapter must prove the path runs and record
 * attributable hashes.
 */
import { describe, expect, it } from 'vitest'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { compareCfCase } from '../scripts/cf-compare'
import type { CfManifest } from '../scripts/cf-refs/cases'
import { readFileSync } from 'node:fs'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')

function loadManifest(): CfManifest {
  return JSON.parse(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
}

describe('viewport adapters run the data path and bind hashes', () => {
  it('interactive adapter prepares resident viewport states with hashes', async () => {
    const manifest = loadManifest()
    const c = manifest.cases.find(entry => entry.id === 'stage1-cellis-literals')!
    const res = await compareCfCase(c, 'interactive', { rootDir: REPO_ROOT })
    expect(res.outcome).toBe('recipe-pending')
    expect(res.outputHash, 'adapter records an output hash').toMatch(/^[0-9a-f]{64}$/)
    expect(res.workbookHash, 'adapter binds the workbook hash').toBe(c.expected.savedSha256)
    expect(res.oracleRef, 'adapter binds the oracle case').toContain('stage1-cellis-literals')
    expect(res.reads, 'adapter ran real reads').toBeGreaterThan(0)
  }, 60000)

  it('paged adapter reads worker viewport states with hashes', async () => {
    const manifest = loadManifest()
    const c = manifest.cases.find(entry => entry.id === 'stage1-cellis-literals')!
    const res = await compareCfCase(c, 'paged', { rootDir: REPO_ROOT })
    expect(res.outcome).toBe('recipe-pending')
    expect(res.outputHash, 'adapter records an output hash').toMatch(/^[0-9a-f]{64}$/)
    expect(res.workbookHash, 'adapter binds the workbook hash').toBe(c.expected.savedSha256)
    expect(res.oracleRef, 'adapter binds the oracle case').toContain('stage1-cellis-literals')
    expect(res.reads, 'adapter ran real reads').toBeGreaterThan(0)
  }, 60000)
})
