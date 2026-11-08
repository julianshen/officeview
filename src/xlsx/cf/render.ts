/** Async CF preparation; painters only consume admitted source-space snapshots. */
import { createCfStatistics } from './statistics'
import { sameCfStamp } from './cache'
import { cfRuntimeEvidence, type CfEvidenceGate } from './predicates'
import { createCfService, type CfPaintStatistics } from './service'
import { CF_LIMITS, createCfMemory, createCfWorkBudget, measureCfPayload, type CfMemory } from './limits'
import { resolveCfCellStyle } from './dxf'
import type { CfCellState, CfLiteral, CfPrepareRequest, CfStamp, CfValueSource, SourceRect } from './model'
import type { XlsxSheet } from '../types'
import type { RegionStyleLayer } from '../tableStyles/regionLayer'
import { regionStylesFor } from '../tableStyleLoader'
import { cellSourceRow } from '../view/sheetCells'
import { isValueUnavailable, type ValuePolicy } from '../view/values'
import { cfNeedsViewport, cfExplicitBounds } from '../cfLoader'
import { STATIC_SHEET_BUDGET } from '../render'
import { CfPageFault, cfRulesVersion, pagedStoreCfPages } from '../view/pagedSource'
import type { SourceStore } from '../view/contracts'
import { hostTask } from '../view/scheduler'

export interface CfPaintBatch {
  readonly stamp: CfStamp
  readonly key: string
  readonly layer: RegionStyleLayer
  dispose(): void
}
/** No inferred values: cached formula errors stay errors; missing formula
 * values refuse rather than masquerading as blanks. */
/** Task 9 reuses the borrowed-value resident source for per-batch text scopes. */
export function residentSource(sheet: XlsxSheet, stamp: CfStamp, memory: CfMemory, policy: ValuePolicy): CfValueSource {
  let bounds: SourceRect | undefined, ordered = true
  return {
    stamp: () => stamp,
    async bounds(_id, budget, signal) {
      if (bounds) return bounds
      let previousRow = -1
      let firstRow = 1048576, lastRow = -1, firstCol = 16384, lastCol = -1
      for (const row of sheet.rows) {
        if (!Number.isSafeInteger(row.index) || row.index < previousRow) ordered = false
        previousRow = row.index
        const rowWork = budget?.spend(1)
        if (rowWork === 'refused') throw new Error('CF work limit')
        if (rowWork === 'yield') { await hostTask(); budget?.resume() }
        for (const cell of row.cells) {
        if (signal?.aborted) throw new Error('CF cancelled')
        const work = budget?.spend(1)
        if (work === 'refused') throw new Error('CF work limit')
        if (work === 'yield') { await hostTask(); budget?.resume() }
        const r = cellSourceRow(row, cell)
        if (r !== row.index) ordered = false
        firstRow = Math.min(firstRow, r); lastRow = Math.max(lastRow, r); firstCol = Math.min(firstCol, cell.col); lastCol = Math.max(lastCol, cell.col)
      } }
      return bounds = lastRow < 0 ? { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 } : { firstRow, lastRow, firstCol, lastCol }
    },
    async *pages(_id, rects, budget, signal) {
      // One-cell borrowed-value pages avoid copying strings or indexing the
      // entire resident sheet. The service owns its admitted result map.
      const hold = memory.reserve(256)
      if (hold.kind !== 'reserved') throw new Error('CF page admission')
      try {
        const rows = new Uint32Array(1), cols = new Uint16Array(1)
        const yielded = new Set<number>()
        let begin = 0, end = sheet.rows.length
        if (ordered && bounds && rects.length) {
          const first = Math.min(...rects.map(r => r.firstRow)), last = Math.max(...rects.map(r => r.lastRow))
          const lower = (value: number) => { let a = 0, b = sheet.rows.length; while (a < b) { const m = (a + b) >>> 1; if (sheet.rows[m]!.index < value) a = m + 1; else b = m }; return a }
          begin = lower(first); end = lower(last + 1)
        }
        for (let ri = begin; ri < end; ri++) { const row = sheet.rows[ri]!; for (const cell of row.cells) {
          if (signal.aborted) throw new Error('CF cancelled')
          const work = budget.spend(1)
          if (work === 'refused') throw new Error('CF work limit')
          if (work === 'yield') { await hostTask(); budget.resume() }
          const r = cellSourceRow(row, cell)
          if (!rects.some(b => r >= b.firstRow && r <= b.lastRow && cell.col >= b.firstCol && cell.col <= b.lastCol)) continue
          if (isValueUnavailable(cell, policy)) throw new Error('CF unavailable formula value')
          let v = cell.value
          let isErr = cell.valueIsError
          const merge = sheet.mergeRanges?.find(m => r >= m.minRow && r <= m.maxRow && cell.col >= m.minCol && cell.col <= m.maxCol)
          if (merge && (r !== merge.minRow || cell.col !== merge.minCol)) {
            const anchorRow = sheet.rows.find(ar => ar.index === merge.minRow)
            const anchorCell = anchorRow?.cells.find(ac => ac.col === merge.minCol)
            if (anchorCell) {
              v = anchorCell.value
              isErr = anchorCell.valueIsError
            }
          }
          const value: CfLiteral = isErr ? { kind: 'error', value: String(v) } : v === null ? { kind: 'blank' }
            : typeof v === 'number' ? { kind: 'number', value: v } : typeof v === 'boolean' ? { kind: 'boolean', value: v } : { kind: 'string', value: v }
          rows[0] = r; cols[0] = cell.col
          yielded.add(r * 16384 + cell.col)
          yield { value: { rows, cols, values: [value] }, bytes: 256, release() {}, transfer() { return { kind: 'unresolved', reason: 'borrowed-page' } } }
        } }
        if (sheet.mergeRanges?.length) {
          for (const m of sheet.mergeRanges) {
            if (!rects.some(b => m.minRow <= b.lastRow && m.maxRow >= b.firstRow && m.minCol <= b.lastCol && m.maxCol >= b.firstCol)) continue
            const anchorRow = sheet.rows.find(ar => ar.index === m.minRow)
            const anchorCell = anchorRow?.cells.find(ac => ac.col === m.minCol)
            if (!anchorCell) continue
            const v = anchorCell.value
            const isErr = anchorCell.valueIsError
            const value: CfLiteral = isErr ? { kind: 'error', value: String(v) } : v === null ? { kind: 'blank' }
              : typeof v === 'number' ? { kind: 'number', value: v } : typeof v === 'boolean' ? { kind: 'boolean', value: v } : { kind: 'string', value: v }
            for (let mr = m.minRow; mr <= m.maxRow; mr++) {
              for (let mc = m.minCol; mc <= m.maxCol; mc++) {
                if (yielded.has(mr * 16384 + mc)) continue
                if (!rects.some(b => mr >= b.firstRow && mr <= b.lastRow && mc >= b.firstCol && mc <= b.lastCol)) continue
                rows[0] = mr; cols[0] = mc
                yielded.add(mr * 16384 + mc)
                yield { value: { rows, cols, values: [value] }, bytes: 256, release() {}, transfer() { return { kind: 'unresolved', reason: 'borrowed-page' } } }
              }
            }
          }
        }
      } finally { hold.release() }
    },
  }
}
/** Read the committed store, including pivot overlays. A borrowed snapshot stops
 * painting synchronously when its source commits, before any async replacement. */
export function storeCfSource(store: SourceStore, stamp: CfStamp, memory: CfMemory): CfValueSource {
  return {
    stamp: () => ({ ...stamp, owner: store.stamp.docId, workbookValueGen: store.stamp.valueGen, structureGen: store.structureGen ?? 0 }),
    bounds: async () => { if (store.completeness !== 'complete') return undefined; const used = store.extent.used; return used
      ? { firstRow: used.r0, lastRow: used.r1, firstCol: used.c0, lastCol: used.c1 }
      : { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 } },
    pages: pagedStoreCfPages(store, memory),
  }
}
/** One diagnostic surface for initial and late refusals, with the sheet cap. */
function unavailable(sheet: XlsxSheet): undefined {
  const ds = sheet.diagnostics ??= []
  if (ds.length < CF_LIMITS.diagnostics && !ds.some(d => d.feature === 'cf-paint-unavailable')) ds.push({ kind: 'unsupported', feature: 'cf-paint-unavailable', message: 'Conditional formatting could not be prepared within the admitted batch; base formatting is shown.' })
  return undefined
}
// Retained statistics sessions across preparations / viewport scrolls per (sheet, generation) (Ruling I5)
const sheetStatsSessions = new WeakMap<object, { session: import('./statistics').CfStatisticsSession; key: string }>()

/** Shared by resident assets and paged callers. paintRects scopes whole-axis
 * rules to the visible/printed source coordinates, including unstored blanks. */
export async function prepareCfPaint(sheet: XlsxSheet, options: {
  store?: SourceStore; memory?: CfMemory; request?: CfPrepareRequest; source?: CfValueSource; evidence?: CfEvidenceGate;
  paintRects?: readonly SourceRect[]; valuePolicy?: ValuePolicy; dateContext?: import('./predicates').CfDateContext;
  viewport?: { metrics: import('../view/contracts').SheetMetrics; boxes: readonly import('../view/contracts').LogicalCellBox[] }
} = {}): Promise<CfPaintBatch | undefined> {
  const rules = sheet.conditionalFormatting
  if (!rules) return undefined
  // An importing prefix cannot certify either absence or range statistics.
  if (options.store && options.store.completeness !== 'complete') return undefined
  let requested = options.request?.sourceRects ?? options.paintRects
  if (rules.rules.length > CF_LIMITS.rules || (requested?.length ?? 0) > CF_LIMITS.rangesPerSheet) return unavailable(sheet)
  let rangeCount = 0, propertyBytes = 0
  for (const rule of rules.rules) {
    if (rule.ranges.length > CF_LIMITS.rangesPerRule || (rangeCount += rule.ranges.length) > CF_LIMITS.rangesPerSheet) return unavailable(sheet)
    for (const key in rule.dxf) { const v = rule.dxf[key]; propertyBytes += 192 + key.length * 4 + (typeof v === 'string' ? v.length * 2 : 16) }
    for (const key of rule.writes) propertyBytes += 128 + key.length * 4
  }
  const memory = options.memory ?? createCfMemory()
  const hold = memory.reserve(2048 + Math.max(rangeCount, requested?.length ?? 0) * 512)
  if (hold.kind !== 'reserved') return unavailable(sheet)
  const evidence = options.evidence ?? cfRuntimeEvidence
  const stamp = options.request?.stamp ?? { owner: options.store?.stamp.docId ?? 'static', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: options.store?.stamp.valueGen ?? 0, cfGen: options.store ? cfRulesVersion(rules) : 1, viewGen: 0, structureGen: options.store?.structureGen ?? 0, clockGen: 0 }
  const signal = options.request?.signal ?? new AbortController().signal
  const source = options.source ?? (options.store ? storeCfSource(options.store, stamp, memory) : residentSource(sheet, stamp, memory, options.valuePolicy ?? 'cached'))
  const budget = createCfWorkBudget({ signal })
  const hiddenRows: SourceRect[] = []
  for (const r of sheet.rows) {
    if (r.hidden) hiddenRows.push({ firstRow: r.index, lastRow: r.index, firstCol: 0, lastCol: 16383 })
  }
  const statisticsOptions = hiddenRows.length > 0 ? { hiddenRows } : undefined
  // Ruling I5: Cache range statistics per (sheet, rule range, data version) across preparations.
  const currentStatsKey = `${stamp.owner}|${stamp.sheetId}|${stamp.workbookValueGen}|${stamp.structureGen}`
  let retainedStats = sheetStatsSessions.get(sheet)
  if (!retainedStats || retainedStats.key !== currentStatsKey) {
    if (retainedStats) retainedStats.session.dispose()
    const session = createCfStatistics(memory, { maxEntries: 64 })
    retainedStats = { session, key: currentStatsKey }
    sheetStatsSessions.set(sheet, retainedStats)
  }
  const statistics: CfPaintStatistics = { session: retainedStats.session, stamp, results: new Map() }
  const dateContext = options.dateContext ?? (sheet.semantics ? {
    dateSystem: sheet.semantics.dateSystem,
    timeZone: sheet.semantics.timeZone,
    epochNowMs: sheet.semantics.epochNowMs,
  } : undefined)
  let units = 0
  const work = async () => { if (++units > CF_LIMITS.workUnits) throw new Error('CF paint work'); const result = budget.spend(1); if (result === 'refused') throw new Error('CF paint work'); if (result === 'yield') { await hostTask(); budget.resume() } }
  try {
    const scoped = Boolean(options.viewport && cfNeedsViewport(sheet))
    if (options.viewport && scoped) {
      const { metrics, boxes } = options.viewport
      // Union column intervals immediately, so duplicate rules/panes never
      // multiply retained row records. Every new row/interval is admitted first.
      const rowsBySource = new Map<number, Array<[number, number]>>()
      for (const box of boxes) {
        const rows = metrics.rows.visibleBetween(box.y, box.y + box.height), cols = metrics.cols.visibleBetween(box.x, box.x + box.width)
        for (const rr of rows.runs) for (let r = rr.start; r <= rr.end; r++) for (const cc of cols.runs) {
          const sourceRow = metrics.order.displayToSource(r)
          for (const rule of rules.rules) for (const range of rule.ranges) {
            await work()
            if (sourceRow < range.firstRow || sourceRow > range.lastRow) continue
            let from = Math.max(cc.start, range.firstCol), to = Math.min(cc.end, range.lastCol)
            if (from > to) continue
            let intervals = rowsBySource.get(sourceRow)
            if (!intervals) {
              if (rowsBySource.size >= CF_LIMITS.rangesPerSheet || !hold.grow(hold.bytes + 512)) throw new Error('CF viewport framing')
              rowsBySource.set(sourceRow, intervals = [])
            }
            let first = 0
            while (first < intervals.length && intervals[first][1] < from - 1) { await work(); first++ }
            let end = first
            while (end < intervals.length && intervals[end][0] <= to + 1) {
              await work(); from = Math.min(from, intervals[end][0]); to = Math.max(to, intervals[end][1]); end++
            }
            if (end === first && !hold.grow(hold.bytes + 256)) throw new Error('CF viewport interval')
            intervals.splice(first, end - first, [from, to])
          }
        }
      }
      // Sort in source order (including filtered/reordered rows), then extend
      // identical intervals on adjacent rows rather than rescanning row slabs.
      type Slab = { -readonly [K in keyof SourceRect]: SourceRect[K] }
      const visible: Slab[] = [], sourceRows = [...rowsBySource.keys()].sort((a,b)=>a-b)
      let previous = new Map<number, Slab>(), previousRow = -2
      for (const row of sourceRows) {
        const current = new Map<number, Slab>()
        for (const [from,to] of rowsBySource.get(row)!) {
          await work()
          const prior = previousRow === row - 1 ? previous.get(from) : undefined
          if (prior && prior.lastCol === to) { prior.lastRow = row; current.set(from, prior) }
          else {
            if (visible.length >= CF_LIMITS.rangesPerSheet || !hold.grow(hold.bytes + 512)) throw new Error('CF viewport slabs')
            const rect = {firstRow:row,lastRow:row,firstCol:from,lastCol:to}
            visible.push(rect); current.set(from,rect)
          }
        }
        previous = current; previousRow = row
      }
      requested = visible
    }
    const rects: SourceRect[] = []
    if (requested) rects.push(...requested)
    else {
      // computeMetrics includes merges, declarations and finite CF extents (bounded by I3). Open
      // tails do not enlarge it; their coverage is clipped to this paint region.
      const finite = cfExplicitBounds(sheet)
      let bottom = finite && sheet.rows.length === 0 ? finite.r1 : 0
      let right = finite && sheet.rows.length === 0 ? finite.c1 : 0
      for (const row of sheet.rows) {
        await work(); bottom = Math.max(bottom, row.index)
        for (const cell of row.cells) {
          await work(); bottom = Math.max(bottom, cellSourceRow(row, cell)); right = Math.max(right, cell.col)
        }
      }
      for (const m of sheet.mergeRanges ?? []) { await work(); bottom = Math.max(bottom, m.maxRow); right = Math.max(right, m.maxCol) }
      for (const r of regionStylesFor(sheet)?.rects ?? []) { await work(); bottom = Math.max(bottom, r.maxRow); right = Math.max(right, r.maxCol) }
      for (const [row, cols] of sheet.resolvedSparklines ?? []) { bottom = Math.max(bottom, row); for (const col of cols.keys()) { await work(); right = Math.max(right, col) } }
      const declared = sheet.cols.reduce((n, c) => Math.max(n, c.max), right)
      if (declared <= right + 64) right = declared
      if (sheet.drawings?.length) { bottom = Math.max(bottom, sheet.drawingMarkers?.maxRow ?? 0); right = Math.max(right, sheet.drawingMarkers?.maxCol ?? 0) }

      bottom = Math.min(bottom, STATIC_SHEET_BUDGET.maxGridRows - 1)
      right = Math.min(right, STATIC_SHEET_BUDGET.maxGridCols - 1)

      const aggregateRuleWork = rules.rules.reduce((sum, rule) => {
        return sum + rule.ranges.reduce((rSum, range) => {
          const rLastRow = range.lastRow === 1048575 ? bottom : range.lastRow
          const rLastCol = range.lastCol === 16383 ? right : range.lastCol
          return rSum + Math.max(0, rLastRow - range.firstRow + 1) * Math.max(0, rLastCol - range.firstCol + 1)
        }, 0)
      }, 0)
      if (aggregateRuleWork > CF_LIMITS.workUnits) return unavailable(sheet)

      for (const rule of rules.rules) for (const r of rule.ranges) {
        const lastRow = Math.min(r.lastRow, bottom)
        const lastCol = r.lastCol === 16383 ? right : Math.min(r.lastCol, STATIC_SHEET_BUDGET.maxGridCols - 1)
        if (lastRow >= r.firstRow && lastCol >= r.firstCol) rects.push({ ...r, lastRow, lastCol })
      }
    }
    for (const r of rects) if (![r.firstRow, r.lastRow, r.firstCol, r.lastCol].every(Number.isSafeInteger) || r.firstRow < 0 || r.lastRow > 1048575 || r.firstCol < 0 || r.lastCol > 16383 || r.lastRow < r.firstRow || r.lastCol < r.firstCol) throw new Error('CF rectangle')
    const base = regionStylesFor(sheet)
    const colors = { theme: sheet.drawingTheme?.palette ?? sheet.tableStyleSource?.colors.theme, indexed: sheet.tableStyleSource?.colors.indexed }
    const states = new Map<number, CfCellState>()
    // Split the union into disjoint row slabs and merged column intervals. No
    // full-sheet cell set, and overlapping rules never multiply cell admission.
    const edges = scoped ? [] : [...new Set(rects.flatMap(r => [r.firstRow, r.lastRow + 1]))].sort((a, b) => a - b)
    const maxCells = Math.max(1, Math.min(256, Math.floor((1024 * 1024 - rules.rules.length * 640 - 2208) / (512 + propertyBytes))))
    const paintCells = options.viewport || options.request ? Math.min(CF_LIMITS.preparedCells, rects.reduce((n, r) => n + (r.lastRow - r.firstRow + 1) * (r.lastCol - r.firstCol + 1), 0)) : 256
    if (!hold.grow(hold.bytes + paintCells * 2048)) throw new Error('CF paint workspace')
    let heldBytes = hold.bytes
    let hadWorkLimit = false
    for (let e = 0; e < (scoped ? rects.length : edges.length - 1); e++) {
      const firstRow = scoped ? rects[e].firstRow : edges[e], endRow = scoped ? rects[e].lastRow + 1 : edges[e + 1]
      const intervals: Array<[number, number]> = []
      if (scoped) intervals.push([rects[e].firstCol, rects[e].lastCol])
      else for (const r of rects) { await work(); if (r.firstRow <= firstRow && r.lastRow >= firstRow) intervals.push([r.firstCol, r.lastCol]) }
      intervals.sort((a, b) => a[0] - b[0])
      const merged: Array<[number, number]> = []
      for (const [a, b] of intervals) { const last = merged.at(-1); if (last && a <= last[1] + 1) last[1] = Math.max(last[1], b); else merged.push([a, b]) }
      for (let row = firstRow; row < endRow; row++) for (const [from, to] of merged) for (let col = from; col <= to; col += maxCells) {
        await work()
        const rect = { firstRow: row, lastRow: row, firstCol: col, lastCol: Math.min(to, col + maxCells - 1) }
        // Ruling I4: Per-batch work budget so static sheets (e.g. 3000x20) do not exhaust
        // the 1M whole-sheet limit. If a batch exceeds the budget, refusal is per-range.
        const batchBudget = createCfWorkBudget({ signal })
        const service = createCfService(rules, memory, { evidence, budget: () => batchBudget, statistics, statisticsOptions, dateContext })
        try {
          const lease = await service.prepareCf({ stamp, sourceRects: [rect], signal, grantBytes: options.request?.grantBytes ?? CF_LIMITS.scratchBytes }, source)
          try {
            for (let c = rect.firstCol; c <= rect.lastCol; c++) {
              const result = lease.value.at(row, c)
              if (result.kind !== 'ready' || !Object.keys(result.value.properties).length) continue
              const measured = measureCfPayload(result.value, memory)
              if (measured.kind !== 'ready' || !hold.grow(heldBytes + measured.value + 160)) throw new Error('CF retained paint admission')
              heldBytes += measured.value + 160
              // Ownership overlaps the service lease until the batch is drained.
              states.set(row * 16384 + c, result.value)
            }
          } finally { lease.release() }
        } catch (error) {
          if (error instanceof Error && (error.message === 'CF work limit' || error.message === 'CF cancelled')) {
            hadWorkLimit = true
            // Ruling I4: Refusal is per-range, not the whole sheet
            continue
          }
          throw error
        } finally { service.dispose() }
      }
    }
    if (hadWorkLimit && states.size === 0) return unavailable(sheet)
    let disposed = false
    const layer: RegionStyleLayer = {
      retainedBytes: hold.bytes,
      rects: [...base?.rects ?? [], ...rects.map(r => ({ minRow: r.firstRow, maxRow: r.lastRow, minCol: r.firstCol, maxCol: r.lastCol }))],
      styleAt(row, col, own, mode) {
        const style = base?.styleAt(row, col, own, mode) ?? own
        const merge = sheet.mergeRanges?.find(m => row >= m.minRow && row <= m.maxRow && col >= m.minCol && col <= m.maxCol)
        if (merge && !evidence.allows('layout:merge-hidden-endpoint') && !evidence.allows('layout:merge-hidden-anchor')) return style
        const state = !disposed && sameCfStamp(source.stamp(stamp.sheetId), stamp) && states.get(row * 16384 + col)
        return state ? resolveCfCellStyle(style, state, colors) : style
      },
    }
    return { stamp, key: JSON.stringify([stamp.owner, stamp.sheetId, stamp.workbookValueGen, stamp.cfGen, stamp.viewGen, stamp.structureGen, stamp.clockGen]), layer,
      dispose() { if (disposed) return; disposed = true; states.clear(); hold.release() } }
  } catch (error) { hold.release(); return error instanceof CfPageFault && (error.fault === 'incomplete' || error.fault === 'stale') ? undefined : unavailable(sheet) } finally {
    for (const result of statistics.results.values()) if (result.kind === 'ready') result.value.release()
  }
}

export { createCfMemory as createPaintMemory } from './limits'

/** One facet per receiving document and parent. Parent admission precedes every
 * local allocation, including overlapping viewport and print leases. */
const memories = new WeakMap<object, Map<object | undefined, CfMemory>>()
export function paintMemoryFor(owner: object, parent?: import('../../charts/workspace').ChartParentLedger): CfMemory {
  let map = memories.get(owner)
  if (!map) memories.set(owner, map = new Map())
  let memory = map.get(parent)
  if (!memory) {
    memory = createCfMemory({ context: parent ? {
      limitBytes: CF_LIMITS.workbookMetadataBytes, canReserve: () => true, owns: () => false,
      reserve() {
        let bytes = 0
        return { owner: 'cf-paint', get bytes() { return bytes },
          trySet(next) { if (next > bytes && !parent.reserve(next - bytes, 'cf-paint')) return false; if (next < bytes) parent.release('cf-paint', bytes - next); bytes = next; return true },
          charge(next) { if (!this.trySet(bytes + next)) throw new Error('CF parent admission') },
          release() { parent.release('cf-paint', bytes); bytes = 0 },
          adopt() { throw new Error('CF ownership is held by its paint lease') },
        }
      },
    } : undefined })
    map.set(parent, memory)
  }
  return memory
}
