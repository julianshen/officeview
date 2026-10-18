/**
 * Stream F item 4 (iOS Simulator run): the viewer's own resource profile must
 * follow the admitted device profile. README / spec §12.4: admission selects
 * D for desktop and M for phones and tablets (iPadOS included, which reports a
 * desktop UA with touch); there is no public profile prop. Before the fix the
 * viewer presented with D budgets on every device (`props.profile ?? 'D'`),
 * so a phone's canvas/bitmap pool was capped at 128 MiB instead of 64 MiB.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc, { type SheetFrameStats } from '../src/components/xlsx/InteractiveSheetDoc'
import { BITMAP_BUDGETS } from '../src/xlsx/view/tileCache'
import type { XlsxDocument } from '../src/xlsx/types'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'

function workbook(): XlsxDocument {
  const rows = Array.from({ length: 30 }, (_, r) => ({ index: r, cells: Array.from({ length: 6 }, (_, c) => ({ col: c, value: r * 10 + c })) }))
  return { sheets: [{ name: 'S', rows, cols: [], merges: [], mergeRanges: [] }] } as unknown as XlsxDocument
}
const IPHONE = 'Mozilla/5.0 (iPhone; CPU iPhone OS 18_7 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/26.1 Mobile/15E148 Safari/604.1'
const IPADOS = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/26.1 Safari/605.1.15'

async function budgetUnder(userAgent?: string, maxTouchPoints = 0): Promise<number> {
  if (userAgent) {
    vi.spyOn(navigator, 'userAgent', 'get').mockReturnValue(userAgent)
    Object.defineProperty(navigator, 'maxTouchPoints', { configurable: true, get: () => maxTouchPoints })
  }
  const all: SheetFrameStats[] = []
  const ui = render(<InteractiveSheetDoc doc={workbook()} {...({ onFrame: (stats: SheetFrameStats) => { all.push(stats) } } as object)} />)
  await waitFor(() => expect(all.length).toBeGreaterThan(0), FIRST_PAINT_WAIT)
  const budget = all[all.length - 1].budgetBytes
  ui.unmount()
  return budget
}

describe('the viewer presents with the admitted device profile', () => {
  afterEach(() => {
    vi.restoreAllMocks()
    delete (navigator as { maxTouchPoints?: number }).maxTouchPoints
  })
  it('desktop keeps the D bitmap budget', async () => {
    expect(await budgetUnder()).toBe(BITMAP_BUDGETS.D)
  }, FIRST_PAINT_TEST.timeout)
  it('a phone gets the M bitmap budget', async () => {
    expect(await budgetUnder(IPHONE, 5)).toBe(BITMAP_BUDGETS.M)
  }, FIRST_PAINT_TEST.timeout)
  it('iPadOS (desktop UA with touch) gets the M bitmap budget', async () => {
    expect(await budgetUnder(IPADOS, 5)).toBe(BITMAP_BUDGETS.M)
  }, FIRST_PAINT_TEST.timeout)
})
