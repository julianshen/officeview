import type { XlsxCell, XlsxSheet } from '../../xlsx/types'
import { cellDisplayText } from '../../xlsx/render'
import { FindError } from '../../xlsx/view/find'
import { withFormatWorkspace, type FormatBudget } from '../../xlsx/view/formatWorkspace'
import { boundedSelectionText } from '../../xlsx/view/selection'

/** Only the bounded label leaves this scope; its text is in the DOM cost model. */
export function formatCellPresentation(sheet: XlsxSheet, cell: XlsxCell, budget: FormatBudget, cacheInputs = false): { text: string; state?: 'error' } {
  try {
    return withFormatWorkspace(budget, allocation => {
      const text = cellDisplayText(sheet, cell, allocation)
      const units = Math.min(text.length, 256)
      const inputs = cacheInputs ? (typeof cell.value === 'string' ? cell.value.length * 2 : 0) + (cell.style?.formatCode?.length ?? 0) * 2 : 0
      const bytes = units * 2 + inputs + 256
      allocation.allocate(units * 4 + inputs + 256)
      const presentation = { text: boundedSelectionText(text) }
      allocation.retain?.([presentation], bytes, () => {})
      return presentation
    })
  } catch (error) {
    if (!(error instanceof FindError) || error.code !== 'resource-limit') throw error
    return { text: 'Value unavailable (formatting memory budget).', state: 'error' }
  }
}

/** A bounded viewer cache; each entry's owned text/metadata already has a lease. */
export class SemanticPresentationCache {
  private readonly entries = new Map<XlsxCell, {
    value: XlsxCell['value']; code: string | undefined; numFmtId: number | undefined
    dateSystem: unknown; locale: string | undefined; presentation: { text: string; state?: 'error' }
  }>()
  constructor(private readonly sheet: XlsxSheet, private readonly budget: FormatBudget) {}
  get(cell: XlsxCell): { text: string; state?: 'error' } {
    const { value, style } = cell, dateSystem = this.sheet.semantics?.dateSystem, locale = this.sheet.semantics?.locale
    // Cache snapshots may outlive mutable source strings. Cache only bounded
    // inputs and charge those snapshots too; larger formats/values still format.
    if ((typeof value === 'string' && value.length > 256) || (style?.formatCode?.length ?? 0) > 256) {
      this.entries.delete(cell)
      return formatCellPresentation(this.sheet, cell, this.budget)
    }
    const hit = this.entries.get(cell)
    if (hit && Object.is(hit.value, value) && hit.code === style?.formatCode && hit.numFmtId === style?.numFmtId && hit.dateSystem === dateSystem && hit.locale === locale) {
      this.entries.delete(cell)
      this.entries.set(cell, hit)
      return hit.presentation
    }
    this.entries.delete(cell)
    const presentation = formatCellPresentation(this.sheet, cell, this.budget, true)
    // The presentation's 256 B fixed reservation includes this cache entry.
    // Mutable resident cells/styles are compared by their actual format inputs.
    if (!presentation.state) {
      if (this.entries.size >= 512) this.entries.delete(this.entries.keys().next().value!)
      this.entries.set(cell, { value, code: style?.formatCode, numFmtId: style?.numFmtId, dateSystem, locale, presentation })
    }
    return presentation
  }
}
