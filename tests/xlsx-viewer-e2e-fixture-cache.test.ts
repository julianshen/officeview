// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { SERVED_FIXTURES, servedFixturePath } from '../scripts/harness/fixture-server'
import { GENERATED_FIXTURES, fixturePath } from './browser/viewer-e2e'

const root = fileURLToPath(new URL('..', import.meta.url))
const cache = join(root, '.tmp', 'viewer-e2e-fixtures-v1')

describe('browser and manual viewer fixture cache', () => {
  it('keeps every generated browser fixture in the repository versioned cache', () => {
    for (const name of Object.keys(GENERATED_FIXTURES) as (keyof typeof GENERATED_FIXTURES)[]) {
      expect(fixturePath(name)).toBe(join(cache, name))
    }
  })

  it('shares generated paths between browser and manual harnesses', () => {
    for (const name of Object.keys(SERVED_FIXTURES) as (keyof typeof SERVED_FIXTURES)[]) {
      expect(servedFixturePath(root, name)).toBe(fixturePath(name))
      if (name !== 'f1-small.xlsx') expect(servedFixturePath(root, name)).toBe(join(cache, name))
    }
  })

  it('keeps the manual root argument for committed fixtures only', () => {
    const otherRoot = join(root, 'another-checkout')
    expect(servedFixturePath(otherRoot, 'f1-small.xlsx')).toBe(join(otherRoot, 'tests', 'fixtures', 'xlsx-viewer', 'f1-small.xlsx'))
    expect(servedFixturePath(otherRoot, 'f2-1k.xlsx')).toBe(join(cache, 'f2-1k.xlsx'))
  })
})
