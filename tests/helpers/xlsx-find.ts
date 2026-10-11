import type { XlsxCell, XlsxDocument } from '../../src/xlsx/types'
import type { RowOrder, WorkbookOwner } from '../../src/xlsx/view/contracts'
import { createIdentityOrder } from '../../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../../src/xlsx/view/visibility'
import { createObjectWorkbookSource } from '../../src/worker/viewerObjectSource'
import { initialWorkbookOwner } from '../../src/worker/viewerProtocol'
import type { WorkbookSource } from '../../src/worker/viewerSession'
import type { FindHit, FindInput, WorkbookFind } from '../../src/xlsx/view/find'

export type FixtureValue = string | number | boolean | { error: string } | { unavailable: true }
export interface FixtureSheet { sheetId: string; cells: { row: number; col: number; value: FixtureValue }[]; hiddenRows?: number[] }

/** Parser-shaped workbook: rows/cells only, hidden rows authored on XlsxRow. */
export function workbookOf(sheets: readonly FixtureSheet[]): XlsxDocument {
  return {
    sheets: sheets.map(sheet => {
      const rows = new Map<number, { index: number; hidden?: boolean; cells: XlsxCell[] }>()
      const rowOf = (index: number) => {
        let row = rows.get(index)
        if (!row) rows.set(index, row = { index, cells: [] })
        return row
      }
      for (const index of sheet.hiddenRows ?? []) rowOf(index).hidden = true
      for (const { row, col, value } of sheet.cells) {
        const base = { ref: `R${row}C${col}`, row, col, styleIndex: 0 }
        rowOf(row).cells.push(typeof value !== 'object' ? { ...base, value, hasCachedValue: true }
          : 'error' in value ? { ...base, value: value.error, hasCachedValue: true, valueIsError: true }
            // A formula cell without a cached value.
            : { ...base, value: null, formula: 'NOW()', hasCachedValue: false })
      }
      return { name: `Sheet ${sheet.sheetId}`, sheetId: sheet.sheetId, rows: [...rows.values()], cols: [], merges: [], mergeRanges: [] }
    }),
    tables: [],
  }
}

export interface InputOptions {
  docId?: string
  /** Extra filter-hidden DISPLAY slots per sheet, on top of authored hidden rows. */
  filtered?: Record<string, number[]>
  orders?: Record<string, RowOrder>
  owner?: WorkbookOwner
  source?: WorkbookSource
}
/** Find input built the way the viewer holds it: the resident object source per
 * sheet and the single display-space visibility mask (authored hidden ∪ filter). */
export function findInputOf(doc: XlsxDocument, options: InputOptions = {}): { input: FindInput; source: WorkbookSource } {
  const docId = options.docId ?? 'doc'
  const source = options.source ?? createObjectWorkbookSource(doc, docId)
  const owner = options.owner ?? initialWorkbookOwner(docId, source.sheets)
  return {
    source,
    input: {
      owner,
      sheets: source.sheets.map(({ sheetId }) => {
        const entry = owner.sheetVector.find(item => item.sheetId === sheetId)!
        const order = options.orders?.[sheetId] ?? createIdentityOrder()
        const authored = source.geometry(sheetId).baseHidden.intervals.map(({ r0, r1 }) => ({ r0: order.sourceToDisplay(r0), r1: order.sourceToDisplay(r1) }))
        const filtered = (options.filtered?.[sheetId] ?? []).map(slot => ({ r0: slot, r1: slot }))
        return {
          sheetId,
          store: source.store(sheetId),
          identity: { docId, sheetId, valueGen: entry.valueGen, fontGen: owner.fontGen, viewGen: entry.revision },
          order,
          visibility: visibilityFromHidden(createIntervalSet([...authored, ...filtered])),
        }
      }),
    },
  }
}

export async function drain(session: WorkbookFind): Promise<FindHit[]> {
  const out: FindHit[] = []
  for (;;) {
    const page = await session.nextPage()
    out.push(...page.hits)
    if (page.complete) return out
  }
}
export const where = (hits: readonly FindHit[]) => hits.map(hit => [hit.sheetId, hit.source.row, hit.source.col])
