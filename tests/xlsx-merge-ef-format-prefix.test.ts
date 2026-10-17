import { expect, it } from 'vitest'
import { negativeDateOverflow } from '../src/xlsx/cellFormat'
import type { ResolvedSemantics } from '../src/xlsx/formula/types'

const semantics = { dateSystem: '1900', locale: 'en-US' } as ResolvedSemantics
it('does not build or reserve discarded negative-date sections after the fourth section', () => {
  const code = '0;'.repeat(2000) + '0_ '.repeat(2000)
  let live = 0, peak = 0, sections = 0, firstSectionBudget = -1
  const allocation = { allocate(bytes: number) { if (live + bytes > 64 * 1024) throw new RangeError('workspace refused'); live += bytes; peak = Math.max(peak, live) }, release(bytes: number) { live -= bytes } }
  const descriptor = Object.getOwnPropertyDescriptor(Array.prototype, 'push')!
  const original = Array.prototype.push
  Object.defineProperty(Array.prototype, 'push', { ...descriptor, value: function(this: unknown[], ...items: unknown[]) {
    for (const item of items) if (item !== null && typeof item === 'object' && Object.prototype.hasOwnProperty.call(item, 'body') && Object.prototype.hasOwnProperty.call(item, 'unsupported')) { if (firstSectionBudget < 0) firstSectionBudget = live; sections++ }
    return original.apply(this, items)
  } })
  let result: boolean | undefined
  try { result = negativeDateOverflow(-1, code, semantics, allocation) }
  finally { Object.defineProperty(Array.prototype, 'push', descriptor) }
  expect(result).toBe(false)
  expect(sections).toBe(4)
  expect(firstSectionBudget).toBeGreaterThanOrEqual(4 * (64 + 8)) // four section records and array slots before publication
  expect(peak).toBeLessThan(64 * 1024)
})

it('still refuses retained underscore expansions before negative-date section construction', () => {
  const allocation = { allocate(bytes: number) { if (bytes > 64 * 1024) throw new RangeError('workspace refused') }, release() {} }
  expect(() => negativeDateOverflow(-1, '0_ '.repeat(3000) + ';0', semantics, allocation)).toThrow('workspace refused')
})
