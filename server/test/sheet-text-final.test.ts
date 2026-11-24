/**
 * Stage 4b final review — R1 and R2 regressions, through the real `DocumentHost` parse/snapshot/text-answer path.
 *
 * R1: a table region's number format is part of the painted style. A cell whose own style is General but whose table
 * style's wholeTable DXF writes `;;;` paints nothing, so under `'visible-only'` a11y and fit must not see its value; under
 * `'full'` the view keeps it (fit measures it).
 *
 * R2: a row fit of a rotated or wrapped cell reads the canonical painter layout height, as the local viewer's fit does
 * (`InteractiveSheetDoc.tsx` passes `fit.layoutHeight`); the server's answer equals that local fit.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { autoFitSize, createFitMeasure } from '../../src/components/xlsx/sheetResize'
import { cellDisplayText } from '../../src/xlsx/render'
import { regionStylesFor } from '../../src/xlsx/tableStyleLoader'
import { RESIZE_LIMITS } from '../../src/xlsx/view/axisSizes'
import { buildSheetGeometry } from '../../src/xlsx/view/viewport'
import type { XlsxDocument, XlsxSheet } from '../../src/xlsx/types'
import { serverBackend } from '../src/backend'
import { DocumentHost } from '../src/render'
import type { SheetClock } from '../src/sheet/clock'
import type { SheetTextJob } from '../src/sheet/text-ops'
import { FIT_LAYOUT, layoutRow, TABLE_HIDDEN, tableHiddenFormat } from './sheet-fixtures'

const hosts: DocumentHost[] = []
afterAll(() => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
})

const CLOCK: SheetClock = { day: '2026-11-23', timeZone: 'UTC' }
const PANE = { x: 0, y: 0, width: 640, height: 400 }

async function opened(bytes: Uint8Array, policy: 'visible-only' | 'full') {
  const host = new DocumentHost(); hosts.push(host)
  const { docKey } = await host.open('final', bytes, policy, { sheetMode: 'tiles' })
  const saved = (host as unknown as { docs: Map<string, { savedParse: XlsxDocument }> }).docs.get(docKey)!.savedParse
  const text = (job: SheetTextJob) => host.sheetText(docKey, CLOCK, job)
  const fit = async (axis: 'row' | 'col', index: number) => {
    const r = await text({ op: 'fit', sheetIndex: 0, axis, index, pane: PANE, sizes: [] })
    if (r.op !== 'fit') throw new Error('fit reply')
    return r.size
  }
  return { host, docKey, saved, text, fit }
}

describe('R1: a table-region number format is the painted format', () => {
  test('visible-only: a11y and fit do not see a value the table style hides; full keeps it', async () => {
    const vis = await opened(await tableHiddenFormat(), 'visible-only')
    // The precondition (the reported reason): A2's own style is General, the table region gives it `;;;`, and the
    // region-resolved display text is ''.
    const snap = await vis.host.sheetSnapshot(vis.docKey, 0, CLOCK)
    const a2 = snap.sheet.rows.find((r) => r.index === 1)?.cells.find((c) => c.col === 0)
    expect(a2?.value).toBe(TABLE_HIDDEN.value)
    expect(a2?.style?.formatCode).toBeUndefined()
    expect(a2?.style?.numFmtId ?? 0).toBe(0)
    const region = regionStylesFor(snap.sheet)
    expect(region).toBeDefined()
    const painted = region!.styleAt(1, 0, a2!.style, 'screen')
    expect(painted?.formatCode).toBe(';;;')
    expect(cellDisplayText(snap.sheet, { value: TABLE_HIDDEN.value, style: painted })).toBe('')

    const a11y = await vis.text({ op: 'a11y', sheetIndex: 0, rows: [[0, 1]], cols: [0, 1], sizes: [] })
    expect(a11y).toEqual({ op: 'a11y', cells: [[0, 1, TABLE_HIDDEN.outside]] })
    expect(JSON.stringify(a11y)).not.toContain(String(TABLE_HIDDEN.value))
    // Fit: column A holds nothing drawn, so it fits as the empty column C does.
    const blankCol = await vis.fit('col', 2)
    const visibleA = await vis.fit('col', 0)
    expect(visibleA).toBe(blankCol)

    // Control: under 'full' the view keeps A1 and A2 (identity), so the fit measures their text.
    const full = await opened(await tableHiddenFormat(), 'full')
    expect(await full.fit('col', 0)).toBeGreaterThan(visibleA)
    expect(await full.fit('col', 2)).toBe(blankCol)
  }, 30_000)
})

/** The local viewer's row fit (`InteractiveSheetDoc.tsx` resizeLine): `createFitMeasure({ sheet, admit })` and its
 * `layoutHeight`, measuring on canvases a document creates; here the document hands out the server backend's canvases
 * (the reference the reviewer used). */
function localRowFit(sheet: XlsxSheet, row: number): number {
  const geometry = buildSheetGeometry(sheet)
  if (geometry.kind !== 'ready') throw new Error('fixture geometry')
  const g = globalThis as { document?: unknown }
  const had = 'document' in g, prior = g.document
  const backend = serverBackend()
  g.document = { createElement: () => backend.createCanvas(1, 1), documentElement: undefined }
  try {
    const fit = createFitMeasure({ sheet, admit: () => true })
    expect(fit.layoutHeight).toBeDefined()
    return autoFitSize('row', row, {
      geometry, sheet, pane: PANE, measure: fit.measure, allocation: fit.allocation, ...(fit.layoutHeight ? { layoutHeight: fit.layoutHeight } : {}),
      knownRows: (limit) => sheet.rows.slice(0, limit).map((r) => r.index),
    })
  } finally {
    if (had) g.document = prior
    else delete g.document
  }
}

describe('R2: a row fit reads the canonical layout height', () => {
  const cases: Array<[string, 'rotated' | 'wrapped', string]> = [
    ['rotated 90°, long', 'rotated', FIT_LAYOUT.rotatedLong],
    ['rotated 90°, short', 'rotated', FIT_LAYOUT.rotatedShort],
    ['wrapped', 'wrapped', FIT_LAYOUT.wrapped],
  ]
  for (const [name, kind, text] of cases) {
    test(`${name}: the server row fit equals the local viewer's`, async () => {
      const o = await opened(await layoutRow(kind, text), 'visible-only')
      const local = localRowFit(o.saved.sheets[0]!, 0)
      // The fixture needs more than one line height and stays under the clamp, so a collapsed or clamped fit shows.
      expect(local).toBeGreaterThan(40)
      expect(local).toBeLessThan(RESIZE_LIMITS.row.max)
      expect(await o.fit('row', 0)).toBe(local)
    }, 30_000)
  }
})
