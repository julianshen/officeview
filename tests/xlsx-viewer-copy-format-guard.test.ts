import { expect, it } from 'vitest'
import { cellDisplayText } from '../src/xlsx/render'
import { formatCommon, type FormatAllocation } from '../src/xlsx/formula/format'
import { workbookOf } from './helpers/xlsx-find'
import type { XlsxCellStyle } from '../src/xlsx/types'

function allocation(capacity = Infinity) {
  let live = 0, peak = 0, borrowed = false
  const context: FormatAllocation = {
    allocate: bytes => {
      expect(Number.isSafeInteger(bytes) && bytes >= 0).toBe(true)
      if (live + bytes > capacity) throw new Error('reservation refused')
      live += bytes; peak = Math.max(peak, live)
    },
    release: bytes => { live -= bytes; expect(live).toBeGreaterThanOrEqual(0) },
    borrowed: () => { borrowed = true },
  }
  return { context, live: () => live, peak: () => peak, borrowed: () => borrowed }
}

it.each([
  [1234567.125, '#,##0.00'], [-123.5, '0.00;(0.00);"zero"'], [0, '0;-0;"zero"'],
  [0.125, '0.0%'], [1234567, '0,,'], [1.2345e-100, '0.' + '0'.repeat(2048)],
  [45292.56789, 'dddd, mmmm dd, yyyy h:mm:ss AM/PM'], [-45292.56789, 'yyyy-mm-dd'],
  ['abc', '0;0;0;"pre"@@"post"'], ['abc', '0;0;0;'], [true, '0;0;0;@@'],
  [123, '[Red]0'], [Infinity, '0.00'], ['abc', '0'],
] as const)('guarded canonical formatting preserves output/diagnostics for %s with %s', (value, code) => {
  const plain = workbookOf([{ sheetId: 'A', cells: [] }]).sheets[0]
  const guarded = workbookOf([{ sheetId: 'A', cells: [] }]).sheets[0]
  const style: XlsxCellStyle = { numFmtId: 0, formatCode: code }, guard = allocation()
  const expected = cellDisplayText(plain, { value, style })
  expect(cellDisplayText(guarded, { value, style }, guard.context)).toBe(expected)
  expect(guarded.diagnostics).toEqual(plain.diagnostics)
  expect(guard.peak()).toBeGreaterThanOrEqual(guard.live())
})

it.each([0, 2, 9, 10, 3, 4, 14, 18, 19, 20, 21, 22])('guarded built-in numFmtId %i stays canonical', numFmtId => {
  const sheet = workbookOf([{ sheetId: 'A', cells: [] }]).sheets[0], style = { numFmtId }
  const guard = allocation()
  expect(cellDisplayText(sheet, { value: 45292.56789, style }, guard.context)).toBe(cellDisplayText(sheet, { value: 45292.56789, style }))
  expect(guard.peak()).toBeGreaterThan(0)
})

it('unformatted and one-section authored source strings are borrowed', () => {
  const sheet = workbookOf([{ sheetId: 'A', cells: [] }]).sheets[0]
  const ordinary = allocation(), authored = allocation()
  expect(cellDisplayText(sheet, { value: 'raw' }, ordinary.context)).toBe('raw')
  expect(ordinary.peak()).toBe(0)
  expect(ordinary.borrowed()).toBe(true)
  expect(cellDisplayText(sheet, { value: 'raw', style: { numFmtId: 0, formatCode: '0' } }, authored.context)).toBe('raw')
  expect(authored.borrowed()).toBe(true)
})

it('refuses valid text expansion while the next canonical output is still below its final size', () => {
  const sheet = workbookOf([{ sheetId: 'A', cells: [] }]).sheets[0], guard = allocation(150_000)
  expect(() => cellDisplayText(sheet, { value: 'x'.repeat(256), style: { numFmtId: 0, formatCode: '0;0;0;' + '@'.repeat(1024) } }, guard.context)).toThrow('reservation refused')
  expect(guard.peak()).toBeLessThanOrEqual(150_000)
})

it('denies parser/token and numeric padding allocations without returning partial text', () => {
  const sheet = workbookOf([{ sheetId: 'A', cells: [] }]).sheets[0]
  for (const [value, code, cap] of [[1, '0'.repeat(128), 100], [0.5, '0.' + '0'.repeat(2048), 150_000]] as const) {
    const guard = allocation(cap)
    expect(() => formatCommon(value, code, sheet.semantics!, guard.context)).toThrow('reservation refused')
  }
})
