/** Async CF preparation; painters only consume admitted source-space snapshots. */
import { cfRuntimeEvidence, type CfEvidenceGate } from './predicates'
import { createCfService } from './service'
import { CF_LIMITS, createCfMemory, createCfWorkBudget, type CfMemory } from './limits'
import { resolveCfCellStyle } from './dxf'
import type { CfLease, CfLiteral, CfPrepared, CfPrepareRequest, CfStamp, CfValueSource, SourceRect } from './model'
import type { XlsxSheet } from '../types'
import type { RegionStyleLayer } from '../tableStyles/regionLayer'
import { regionStylesFor } from '../tableStyleLoader'
import { cellSourceRow } from '../view/sheetCells'
import { hostTask } from '../view/scheduler'

export interface CfPaintBatch {
  readonly stamp: CfStamp
  readonly key: string
  readonly layer: RegionStyleLayer
  dispose(): void
}
/** No inferred values: cached formula errors stay errors; missing formula
 * values refuse rather than masquerading as blanks. */
function residentSource(sheet: XlsxSheet, stamp: CfStamp, memory: CfMemory): CfValueSource {
  return {
    stamp: () => stamp,
    async bounds(_id, budget, signal) {
      let firstRow = 1048576, lastRow = -1, firstCol = 16384, lastCol = -1
      for (const row of sheet.rows) for (const cell of row.cells) {
        if (signal?.aborted) throw new Error('CF cancelled')
        const work = budget?.spend(1)
        if (work === 'refused') throw new Error('CF work limit')
        if (work === 'yield') { await hostTask(); budget?.resume() }
        const r = cellSourceRow(row, cell)
        firstRow = Math.min(firstRow, r); lastRow = Math.max(lastRow, r); firstCol = Math.min(firstCol, cell.col); lastCol = Math.max(lastCol, cell.col)
      }
      return lastRow < 0 ? { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 } : { firstRow, lastRow, firstCol, lastCol }
    },
    async *pages(_id, rects, budget, signal) {
      // One-cell borrowed-value pages avoid copying strings or indexing the
      // entire resident sheet. The service owns its admitted result map.
      const hold = memory.reserve(256)
      if (hold.kind !== 'reserved') throw new Error('CF page admission')
      try {
        const rows = new Uint32Array(1), cols = new Uint16Array(1)
        for (const row of sheet.rows) for (const cell of row.cells) {
          if (signal.aborted) throw new Error('CF cancelled')
          const work = budget.spend(1)
          if (work === 'refused') throw new Error('CF work limit')
          if (work === 'yield') { await hostTask(); budget.resume() }
          const r = cellSourceRow(row, cell)
          if (!rects.some(b => r >= b.firstRow && r <= b.lastRow && cell.col >= b.firstCol && cell.col <= b.lastCol)) continue
          if (cell.formula && (cell.hasCachedValue === false || cell.value === null)) throw new Error('CF unavailable formula value')
          const v = cell.value
          const value: CfLiteral = cell.valueIsError ? { kind: 'error', value: String(v) } : v === null ? { kind: 'blank' }
            : typeof v === 'number' ? { kind: 'number', value: v } : typeof v === 'boolean' ? { kind: 'boolean', value: v } : { kind: 'string', value: v }
          rows[0] = r; cols[0] = cell.col
          yield { value: { rows, cols, values: [value] }, bytes: 256, release() {}, transfer() { return { kind: 'unresolved', reason: 'borrowed-page' } } }
        }
      } finally { hold.release() }
    },
  }
}
/** Shared by resident assets and async paged callers. Callers pass the exact
 * viewport source rectangles/stamp/source for paged preparation. */
export async function prepareCfPaint(sheet: XlsxSheet, options: {
  memory?: CfMemory; request?: CfPrepareRequest; source?: CfValueSource; evidence?: CfEvidenceGate
} = {}): Promise<CfPaintBatch | undefined> {
  const rules = sheet.conditionalFormatting
  if (!rules || rules.rules.length > CF_LIMITS.rules || (options.request?.sourceRects.length ?? 0) > CF_LIMITS.rangesPerRule) return undefined
  let rangeCount = 0
  for (const rule of rules.rules) {
    if (rule.ranges.length > CF_LIMITS.rangesPerRule || (rangeCount += rule.ranges.length) > CF_LIMITS.rangesPerSheet) return undefined
  }
  const memory = options.memory ?? createCfMemory()
  const hold = memory.reserve(1024 + Math.max(rangeCount, options.request?.sourceRects.length ?? 0) * 192)
  if (hold.kind !== 'reserved') return undefined
  const evidence = options.evidence ?? cfRuntimeEvidence
  const service = createCfService(rules, memory, { evidence })
  let lease: CfLease<CfPrepared> | undefined
  try {
    const stamp = options.request?.stamp ?? { owner: 'static', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
    const source = options.source ?? residentSource(sheet, stamp, memory)
    const rects: SourceRect[] = []
    if (options.request) rects.push(...options.request.sourceRects)
    else {
      const fullAxis = rules.rules.some(rule => rule.ranges.some(r => r.lastRow === 1048575 || r.lastCol === 16383))
      const bounds = fullAxis ? await source.bounds(stamp.sheetId, createCfWorkBudget(), new AbortController().signal) : undefined
      for (const rule of rules.rules) for (const r of rule.ranges) {
        // Whole-axis declarations are scoped to populated content.
        const firstRow = r.firstRow, firstCol = r.firstCol
        const lastRow = r.firstRow === 0 && r.lastRow === 1048575 ? bounds?.lastRow ?? 0 : r.lastRow
        const lastCol = r.firstCol === 0 && r.lastCol === 16383 ? bounds?.lastCol ?? 0 : r.lastCol
        if (!rects.some(q => q.firstRow === firstRow && q.lastRow === lastRow && q.firstCol === firstCol && q.lastCol === lastCol)) rects.push({ firstRow, lastRow, firstCol, lastCol })
        if (rects.length > CF_LIMITS.rangesPerRule) throw new Error('CF range limit')
      }
    }
    let cells = 0
    for (const r of rects) {
      cells += (r.lastRow - r.firstRow + 1) * (r.lastCol - r.firstCol + 1)
      if (!Number.isSafeInteger(cells) || cells < 0 || cells > CF_LIMITS.preparedCells) throw new Error('CF cell limit')
    }
    // Shared preparation and paint scratch have separate ownership. Covers
    // transient styles, synthetic cell records and gridline masks before paint.
    if (!hold.grow(hold.bytes + cells * 2048)) throw new Error('CF paint admission')
    lease = await service.prepareCf(options.request ?? { stamp, sourceRects: rects, signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, source)
    const base = regionStylesFor(sheet)
    const layer: RegionStyleLayer = {
      retainedBytes: hold.bytes + lease.bytes,
      rects: [...base?.rects ?? [], ...rects.map(r => ({ minRow: r.firstRow, maxRow: r.lastRow, minCol: r.firstCol, maxCol: r.lastCol }))],
      styleAt(row, col, own, mode) {
        const style = base?.styleAt(row, col, own, mode) ?? own
        const merge = sheet.mergeRanges?.find(m => row >= m.minRow && row <= m.maxRow && col >= m.minCol && col <= m.maxCol)
        if (merge && !evidence.allows('stage1-blank-merge-order', 'layout:merge-hidden-endpoint')) return style
        const state = lease?.value.at(row, col)
        return state?.kind === 'ready' ? resolveCfCellStyle(style, state.value) : style
      },
    }
    let disposed = false
    return { stamp, key: JSON.stringify([stamp.owner, stamp.sheetId, stamp.workbookValueGen, stamp.cfGen, stamp.viewGen, stamp.structureGen, stamp.clockGen]), layer, dispose() { if (disposed) return; disposed = true; lease?.release(); lease = undefined; service.dispose(); hold.release() } }
  } catch {
    lease?.release(); service.dispose(); hold.release()
    const diagnostics = (sheet.diagnostics ??= [])
    if (!diagnostics.some(d => d.feature === 'cf-paint-unavailable')) diagnostics.push({ kind: 'unsupported', feature: 'cf-paint-unavailable', message: 'Conditional formatting could not be prepared within the admitted batch; base formatting is shown.' })
    return undefined
  }
}

export { createCfMemory as createPaintMemory } from './limits'
