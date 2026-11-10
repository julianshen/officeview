import { describe, expect, it } from 'vitest'
import {
  hasEvidence,
  injectCfEvidence,
  injectExactCfEvidence,
  VERIFIED_CAPABILITIES,
  withExactCfEvidence,
  withInjectedCfEvidence,
} from '../src/xlsx/cf/evidence'

const PROD = [...VERIFIED_CAPABILITIES].find(id => id.startsWith('scale:'))!

describe('exact evidence injection', () => {
  it('production evidence is non-empty for scales', () => {
    expect(PROD).toBeDefined()
    expect(hasEvidence(PROD)).toBe(true)
  })

  it('additive injection still sees production evidence', () => {
    withInjectedCfEvidence(['other:x'], () => {
      expect(hasEvidence(PROD)).toBe(true)
      expect(hasEvidence('other:x')).toBe(true)
    })
  })

  it('exact injection hides production evidence and restores afterwards', () => {
    withExactCfEvidence([], () => expect(hasEvidence(PROD)).toBe(false))
    withExactCfEvidence(['other:x'], () => {
      expect(hasEvidence('other:x')).toBe(true)
      expect(hasEvidence(PROD)).toBe(false)
    })
    expect(hasEvidence(PROD)).toBe(true)
  })

  it('exact injection is restored after a throw and after async completion', async () => {
    expect(() => withExactCfEvidence([], () => { throw new Error('x') })).toThrow('x')
    expect(hasEvidence(PROD)).toBe(true)
    await withExactCfEvidence([], async () => { expect(hasEvidence(PROD)).toBe(false) })
    expect(hasEvidence(PROD)).toBe(true)
  })

  it('injectExactCfEvidence(null) and injectCfEvidence clear exactness', () => {
    injectExactCfEvidence([])
    expect(hasEvidence(PROD)).toBe(false)
    injectCfEvidence(null)
    expect(hasEvidence(PROD)).toBe(true)
    injectExactCfEvidence([])
    injectExactCfEvidence(null)
    expect(hasEvidence(PROD)).toBe(true)
  })
})
