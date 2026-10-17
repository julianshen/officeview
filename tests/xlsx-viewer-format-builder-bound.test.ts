import { expect, it, vi } from 'vitest'
import { formatColor, formatExtended } from '../src/xlsx/cellFormat'
import { semanticsFor } from '../src/xlsx/view/semanticsFor'

for (const [name, code, body] of [
  ['padding', '[>0]' + '_x'.repeat(2048) + '0', '" "'.repeat(2048) + '0'],
  ['quoted/escaped', '[>0]' + '"abcd"\\x'.repeat(512) + '0', '"abcd"\\x'.repeat(512) + '0'],
  ['currency', '[>0][$' + '€'.repeat(2048) + '-' + '-'.repeat(2048) + ']0', '"' + '€'.repeat(2048) + '"0'],
] as const) it(`admits ${name} section bodies and both structural masks before replacement`, () => {
  let held = 0
  const reservations: number[] = [], replace = String.prototype.replace
  const spy = vi.spyOn(String.prototype, 'replace').mockImplementation(function(this: string, ...args: any[]) {
    if (String(this) === body) reservations.push(held)
    return (replace as any).apply(this, args)
  })
  try {
    expect(typeof formatExtended(3, code, semanticsFor('1900'), String, { allocate: n => { held += n }, release: n => { held -= n } })).toBe('string')
    expect(reservations.length).toBeGreaterThan(0)
    expect(reservations[0]).toBeGreaterThanOrEqual(body.length * 6 + 4 * 72)
  } finally { spy.mockRestore() }
})

it('currency strips only its first locale suffix without constructing a discarded split list', () => {
  const code = '[Red][$€-' + '-'.repeat(8192) + ']0', split = String.prototype.split
  let discardedSplit = false
  const spy = vi.spyOn(String.prototype, 'split').mockImplementation(function(this: string, ...args: any[]) {
    if (String(this).length > 8192 && args[0] === '-') discardedSplit = true
    return (split as any).apply(this, args)
  })
  try {
    expect(formatColor(3, code)).toBe('#FF0000')
    expect(discardedSplit).toBe(false)
  } finally { spy.mockRestore() }
})
