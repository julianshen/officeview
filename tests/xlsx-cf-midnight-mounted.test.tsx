/**
 * The across-midnight follow-up (brief cf-midnight), mounted: a resident viewer left open across local midnight. The tick
 * (`InteractiveSheetDoc` clockGen, `localDaySerial`) re-prepares the sheet; the new preparation withholds the clock rule's
 * cells and the cells whose rule reads a `ca` TODAY() cell, and keeps the clean cells. A sheet whose rules are clock-free
 * but read a tainted cell follows the tick too. Under the injected `rules-tick` reading the tick paints the new day, as
 * production did before the follow-up.
 */
import { afterAll, beforeAll, expect, it, vi } from 'vitest'
import { act, cleanup, render } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import * as paint from '../src/xlsx/cf/render'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { CF_ACROSS_MIDNIGHT_CAP, injectCfAcrossMidnightReadings } from '../src/xlsx/cf/midnight'
import { withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { GREEN, RED, cellIsRule, cfBlock, workbookBytes } from './helpers/cf-views'
import type { XlsxDocument } from '../src/xlsx/types'

const previousTz = process.env.TZ
beforeAll(() => { process.env.TZ = 'UTC' })
afterAll(() => { if (previousTz === undefined) delete process.env.TZ; else process.env.TZ = previousTz })

const DAY = 86_400_000
/** Ten seconds before midnight (UTC, the zone the test pins). */
const LATE = Date.UTC(2026, 10, 16, 23, 59, 50)
const S = Math.floor(LATE / DAY) + 25569
/** Stage 1 evidence every rule below needs: the measured capability set production ships (timePeriod and cellIs). */
const EVIDENCE = ['timePeriod:today', 'cellIs:equal', 'cellIs:greaterThan']

async function parsedAt(bytes: Uint8Array, at: number): Promise<XlsxDocument> {
  return parseXlsx(await OfficePackage.load(bytes), { evaluation: { now: () => at } })
}
/** A1 the load day under `timePeriod today` (green), B1 = 1 under `cellIs > 0` (red), C1 a `ca` TODAY() under `cellIs = S` (green). */
const clockSheet = (): Uint8Array => workbookBytes({ A1: { date: S }, B1: 1, C1: { f: 'TODAY()', v: S - 5, ca: true } },
  cfBlock('A1', '<cfRule type="timePeriod" timePeriod="today" priority="1" dxfId="1"><formula>FLOOR(A1,1)=TODAY()</formula></cfRule>')
  + cfBlock('B1', cellIsRule('greaterThan', '0', 2, 0)) + cfBlock('C1', cellIsRule('equal', String(S), 3, 1)))
/** No clock rule: B1 = 1 under `cellIs > 0` (red), C1 a `ca` TODAY() under `cellIs = S` (green). */
const taintedSheet = (): Uint8Array => workbookBytes({ B1: 1, C1: { f: 'TODAY()', v: S - 5, ca: true } },
  cfBlock('B1', cellIsRule('greaterThan', '0', 1, 0)) + cfBlock('C1', cellIsRule('equal', String(S), 2, 1)))

async function acrossMidnight(bytes: Uint8Array, refs: readonly [number, number][]): Promise<{ before: Array<string | undefined>; after: Array<string | undefined> }> {
  const doc = await parsedAt(bytes, LATE)
  const source = createObjectWorkbookSource(doc, `midnight-${Math.random()}`)
  const batches: Array<{ at: number; batch: import('../src/xlsx/cf/render').CfPaintBatch }> = []
  const actual = paint.prepareCfPaint
  const spy = vi.spyOn(paint, 'prepareCfPaint').mockImplementation(async (sheet, options) => {
    const at = Date.now()
    const out = await actual(sheet, options)
    if (out && options?.store) batches.push({ at, batch: out })
    return out
  })
  vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'Date'], now: LATE })
  const ui = render(<InteractiveSheetDoc doc={doc} source={source} />)
  const fills = (b: import('../src/xlsx/cf/render').CfPaintBatch) => refs.map(([row, col]) => b.layer.styleAt(row, col, undefined)?.fillColor)
  try {
    await vi.waitFor(() => expect(batches.length).toBeGreaterThan(0))
    const before = fills(batches[batches.length - 1]!.batch)
    await act(async () => { await vi.advanceTimersByTimeAsync(15_000) })
    await vi.waitFor(() => expect(batches.some(b => b.at >= LATE + 10_000)).toBe(true))
    const after = fills(batches.filter(b => b.at >= LATE + 10_000).at(-1)!.batch)
    return { before, after }
  } finally { ui.unmount(); cleanup(); spy.mockRestore(); vi.useRealTimers() }
}

it('a clock sheet: the tick withholds the clock rule and the tainted cell, the clean cell keeps painting', async () => {
  await withExactCfEvidence(EVIDENCE, async () => {
    const { before, after } = await acrossMidnight(clockSheet(), [[0, 0], [0, 1], [0, 2]])
    expect(before).toEqual([GREEN, RED, GREEN])
    expect(after).toEqual([undefined, RED, undefined])
  })
})

it('a sheet with clock-free rules over a ca TODAY() cell follows the tick and withholds that cell', async () => {
  await withExactCfEvidence(EVIDENCE, async () => {
    const { before, after } = await acrossMidnight(taintedSheet(), [[0, 1], [0, 2]])
    expect(before).toEqual([RED, GREEN])
    expect(after).toEqual([RED, undefined])
  })
})

it('under the injected rules-tick reading the tick paints the new day (production before the follow-up)', async () => {
  injectCfAcrossMidnightReadings({ [CF_ACROSS_MIDNIGHT_CAP]: 'rules-tick' })
  try {
    await withExactCfEvidence([...EVIDENCE, CF_ACROSS_MIDNIGHT_CAP], async () => {
      const { before, after } = await acrossMidnight(clockSheet(), [[0, 0], [0, 1], [0, 2]])
      expect(before).toEqual([GREEN, RED, GREEN])
      // A1 is yesterday now; C1 still names the load day, as the literal rule reads it.
      expect(after).toEqual([undefined, RED, GREEN])
    })
  } finally { injectCfAcrossMidnightReadings(null) }
})
