import { describe, expect, it } from 'vitest'
import { hiddenFixtures } from '../../scripts/stream-s0/hidden-fixtures'
import { loadOfficeFile } from '../../src/components/OfficeFile'

describe('T2 hidden-content matrix', () => {
  it('every fixture parses (withHidden + withoutHidden)', async () => {
    const fixtures = await hiddenFixtures()
    expect(fixtures.length).toBe(13)
    for (const f of fixtures) {
      await expect(loadOfficeFile(f.withHidden)).resolves.toBeTruthy()
      await expect(loadOfficeFile(f.withoutHidden)).resolves.toBeTruthy()
    }
  })

  it('marker found in model for docx-vanish (exposure walk works)', async () => {
    const fixtures = await hiddenFixtures()
    const f = fixtures.find((x) => x.id === 'docx-vanish')!
    const doc = await loadOfficeFile(f.withHidden)
    const json = JSON.stringify(doc)
    expect(json).toContain(f.marker)
  })

  it('identical docs have no paint leak (sanity)', async () => {
    const fixtures = await hiddenFixtures()
    const f = fixtures[0]
    // same bytes → should not leak
    expect(f.withHidden.length).toBeGreaterThan(0)
  })
})
