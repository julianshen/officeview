/** Pivot value aggregation and show-values-as post-pass (Stage 2, library only). */
import {
  PIVOT_BASE_ITEM_DEFAULT, PIVOT_BASE_ITEM_NEXT, PIVOT_BASE_ITEM_PREVIOUS, PIVOT_VALUES_FIELD,
  type PivotCacheDefinition, type PivotDataFunction, type PivotDiagnostic, type PivotItemType, type PivotShowDataAs, type PivotTableDefinition,
} from '../model'
import type { PivotLine } from '../savedLayout'
import { lineCarriesValues, NO_LINE, type EngineResult, type PivotAxes } from './axes'
import { VALUE_TAG, type RecordColumn, type RecordStore } from './records'
import type { PivotCellValue, PivotValueGrid } from './place'
import type { PivotViewState } from './view'

const MAX_ACCUMULATOR_BYTES = 128 * 1024 * 1024
const ACC_BYTES_PER_SLOT = 57
const OUTPUT_BYTES_PER_SLOT = 9
const TOTAL_SHOW_AS = new Set(['percentOfTotal', 'percentOfRow', 'percentOfCol', 'index', 'percentOfParent', 'percentOfParentRow', 'percentOfParentCol', 'runTotal', 'percentOfRunningTotal'])

const fail = (feature: string, message: string): EngineResult<never> => ({
  ok: false, diagnostic: { kind: 'unsupported', feature, message } satisfies PivotDiagnostic,
})
const unmeasured = (message: string) => fail('pivot-engine-unmeasured', message)
const invalid = (message: string) => ({ ok: false as const, diagnostic: { kind: 'parse' as const, feature: 'pivot-values', message } })

interface Accumulator {
  records: Uint8Array
  nonEmpty: Uint32Array
  numeric: Uint32Array
  sum: Float64Array
  product: Float64Array
  mean: Float64Array
  m2: Float64Array
  min: Float64Array
  max: Float64Array
}

function accumulator(size: number): Accumulator {
  const product = new Float64Array(size); product.fill(1)
  const min = new Float64Array(size); min.fill(Number.POSITIVE_INFINITY)
  const max = new Float64Array(size); max.fill(Number.NEGATIVE_INFINITY)
  return {
    records: new Uint8Array(size), nonEmpty: new Uint32Array(size), numeric: new Uint32Array(size),
    sum: new Float64Array(size), product, mean: new Float64Array(size), m2: new Float64Array(size), min, max,
  }
}

function add(acc: Accumulator, at: number, value: ReadValue, fn: PivotDataFunction): void {
  acc.records[at] = 1
  if (value.kind === 'missing') return
  if (fn === 'count') { acc.nonEmpty[at]++; return }
  if (value.kind !== 'number') return
  if (fn === 'sum' || fn === 'average') acc.sum[at] += value.value
  if (fn === 'sum') return
  if (fn === 'countNums') { acc.numeric[at]++; return }
  if (fn === 'max') { acc.numeric[at]++; if (value.value > acc.max[at]!) acc.max[at] = value.value; return }
  if (fn === 'min') { acc.numeric[at]++; if (value.value < acc.min[at]!) acc.min[at] = value.value; return }
  if (fn === 'product') { acc.numeric[at]++; acc.product[at] *= value.value; return }
  const n = ++acc.numeric[at]
  const delta = value.value - acc.mean[at]!
  acc.mean[at] += delta / n
  if (fn === 'stdDev' || fn === 'stdDevp' || fn === 'var' || fn === 'varp') acc.m2[at] += delta * (value.value - acc.mean[at]!)
}

interface ReadValue { kind: 'missing' | 'other' | 'error' | 'date-unmeasured' | 'number'; value: number }
function readValue(column: RecordColumn | undefined, sharedItems: readonly import('../model').PivotCacheItem[] | undefined,
  row: number, into: ReadValue): boolean {
  if (!column) return false
  if (column.kind === 'shared') {
    const item = sharedItems?.[column.index[row]!]
    if (!item) return false
    switch (item.type) {
      case 'missing': into.kind = 'missing'; break
      case 'number': into.kind = 'number'; into.value = item.value; break
      case 'date': into.kind = 'date-unmeasured'; break
      case 'error': into.kind = 'error'; break
      case 'string': case 'boolean': into.kind = 'other'; break
    }
    return true
  }
  const tag = column.tag[row]
  switch (tag) {
    case VALUE_TAG.missing: into.kind = 'missing'; break
    case VALUE_TAG.number: case VALUE_TAG.date: into.kind = 'number'; into.value = column.num[row]!; break
    case VALUE_TAG.boolean: case VALUE_TAG.string: into.kind = 'other'; break
    case VALUE_TAG.error: into.kind = 'error'; break
    default: return false
  }
  return true
}

function valueOf(acc: Accumulator, at: number, fn: PivotDataFunction): number | undefined {
  const count = acc.numeric[at]!
  if (!acc.records[at]) return undefined
  switch (fn) {
    case 'sum': return acc.sum[at]!
    case 'count': return acc.nonEmpty[at]!
    case 'countNums': return count
    case 'average': return count ? acc.sum[at]! / count : Number.NaN
    case 'max': return count ? acc.max[at]! : Number.NaN
    case 'min': return count ? acc.min[at]! : Number.NaN
    case 'product': return count ? acc.product[at]! : Number.NaN
    case 'stdDev': return count > 1 ? Math.sqrt(acc.m2[at]! / (count - 1)) : Number.NaN
    case 'stdDevp': return count ? Math.sqrt(acc.m2[at]! / count) : Number.NaN
    case 'var': return count > 1 ? acc.m2[at]! / (count - 1) : Number.NaN
    case 'varp': return count ? acc.m2[at]! / count : Number.NaN
  }
}

function subtotalFunction(type: PivotItemType, fallback: PivotDataFunction): PivotDataFunction {
  switch (type) {
    case 'sum': case 'max': case 'min': case 'product': case 'stdDev': case 'var': return type
    case 'avg': return 'average'
    case 'countA': return 'count'
    case 'count': return 'countNums'
    case 'stdDevP': return 'stdDevp'
    case 'varP': return 'varp'
    default: return fallback
  }
}

function isNamedSubtotal(type: PivotItemType): boolean {
  return type !== 'data' && type !== 'default' && type !== 'grand' && type !== 'blank'
}

interface Targets { rows: number[][][]; cols: number[][][] }
interface LineIndex { prefixes: Map<string, number[]>; grand: number[]; grandByData: Map<number, number[]> }
function makePrefixIndex(lines: readonly PivotLine[], view: PivotViewState, axis: 'rows' | 'cols'): LineIndex {
  const prefixes = new Map<string, number[]>(), grand: number[] = [], grandByData = new Map<number, number[]>()
  lines.forEach((line, index) => {
    if (line.type === 'grand') {
      grand.push(index)
      let data = grandByData.get(line.data)
      if (!data) grandByData.set(line.data, data = [])
      data.push(index)
      return
    }
    if (!lineCarriesValues(view, axis, line)) return
    const key = line.items.join(',')
    let items = prefixes.get(key)
    if (!items) prefixes.set(key, items = [])
    items.push(index)
  })
  return { prefixes, grand, grandByData }
}

function targetsFor(lines: readonly PivotLine[], index: LineIndex, view: PivotViewState,
  axis: 'rows' | 'cols', leafIndex: number, data: number): number[] | undefined {
  const leaf = lines[leafIndex]
  if (!leaf) return undefined
  const path = [...leaf.items]
  const fields = axis === 'rows' ? view.rows : view.cols
  const valuesLevel = fields.indexOf(PIVOT_VALUES_FIELD)
  if (valuesLevel >= 0 && valuesLevel < path.length) path[valuesLevel] = data
  const found = new Set<number>()
  for (let length = 0; length <= path.length; length++) {
    for (const lineIndex of index.prefixes.get(path.slice(0, length).join(',')) ?? []) {
      const line = lines[lineIndex]!
      if (valuesLevel >= 0 && line.data !== data) continue
      found.add(lineIndex)
    }
  }
  for (const lineIndex of (valuesLevel < 0 ? index.grand : index.grandByData.get(data) ?? [])) found.add(lineIndex)
  return [...found].sort((a, b) => a - b)
}

function readTargets(store: RecordStore, view: PivotViewState, axes: PivotAxes): EngineResult<Targets> {
  const rows: number[][][] = Array.from({ length: view.data.length }, () => Array.from({ length: axes.rows.length }, () => []))
  const cols: number[][][] = Array.from({ length: view.data.length }, () => Array.from({ length: axes.cols.length }, () => []))
  for (let record = 0; record < store.count; record++) {
    const row = axes.rowKey[record], col = axes.colKey[record]
    if (row === undefined || col === undefined) return invalid(`record ${record} has an invalid leaf line`)
    if (row === NO_LINE || col === NO_LINE) return unmeasured('filtered-out records require measured visible-total semantics')
    if (row >= axes.rows.length || col >= axes.cols.length) return invalid(`record ${record} has an invalid leaf line`)
  }
  const rowLineIndex = makePrefixIndex(axes.rows, view, 'rows'), colLineIndex = makePrefixIndex(axes.cols, view, 'cols')
  for (let d = 0; d < view.data.length; d++) {
    for (let row = 0; row < axes.rows.length; row++) rows[d]![row] = targetsFor(axes.rows, rowLineIndex, view, 'rows', row, d) ?? []
    for (let col = 0; col < axes.cols.length; col++) cols[d]![col] = targetsFor(axes.cols, colLineIndex, view, 'cols', col, d) ?? []
  }
  for (let record = 0; record < store.count; record++) for (let d = 0; d < view.data.length; d++) {
    if (!rows[d]![axes.rowKey[record]!]!.length || !cols[d]![axes.colKey[record]!]!.length)
      return invalid(`record ${record} has no value-carrying row or column line`)
  }
  return { ok: true, value: { rows, cols } }
}

function cellIndex(row: number, col: number, data: number, colCount: number, dataCount: number): number {
  return (row * colCount + col) * dataCount + data
}

/** One-pass aggregation of source records, followed by the show-as calculation. */
export function computeValues(store: RecordStore, cache: PivotCacheDefinition, _def: PivotTableDefinition,
  view: PivotViewState, axes: PivotAxes): EngineResult<PivotValueGrid> {
  if (!view.data.length) return { ok: true, value: { get: () => ({ kind: 'empty' }) } }
  if (cache.olap || cache.source.type !== 'worksheet') return fail('pivot-engine-pending', 'OLAP and non-worksheet sources are not computed')
  if (axes.rowKey.length !== store.count || axes.colKey.length !== store.count) return invalid('record leaf-line arrays do not match the record store')
  if (view.data.some(data => !Number.isInteger(data.field) || data.field < 0 || data.field >= store.columns.length)) return invalid('a data field has no record column')

  const rowCount = axes.rows.length, colCount = axes.cols.length, dataCount = view.data.length
  const gridSize = rowCount * colCount * dataCount
  const rowTotalSize = rowCount * dataCount, colTotalSize = colCount * dataCount
  const estimatedSlots = gridSize + rowTotalSize + colTotalSize + dataCount
  // Include the per-line target lists and prefix-index map; their JS overhead dominates when a pivot has many lines.
  const targetBytes = (rowCount + colCount) * (128 + dataCount * 32)
  const estimatedBytes = estimatedSlots * ACC_BYTES_PER_SLOT + gridSize * OUTPUT_BYTES_PER_SLOT * 2 + targetBytes
  if (!Number.isSafeInteger(estimatedBytes) || estimatedBytes > MAX_ACCUMULATOR_BYTES)
    return fail('pivot-values-limit', 'the value grid exceeds its bounded accumulator budget')

  const recordTargets = readTargets(store, view, axes)
  if (!recordTargets.ok) return recordTargets

  const gridAcc = accumulator(gridSize), rowTotals = accumulator(rowTotalSize), colTotals = accumulator(colTotalSize), grandTotals = accumulator(dataCount)
  const sourceColumns = view.data.map(data => store.columns[data.field])
  const sourceItems = view.data.map(data => cache.fields[data.field]?.sharedItems?.items)
  const functions = view.data.map(data => data.fn)
  const valueScratch: ReadValue = { kind: 'missing', value: 0 }
  for (let record = 0; record < store.count; record++) {
    for (let d = 0; d < dataCount; d++) {
      const field = view.data[d]!.field
      if (!readValue(sourceColumns[d], sourceItems[d], record, valueScratch)) return invalid(`record ${record} has an invalid value in field ${field}`)
      if (valueScratch.kind === 'error') return unmeasured('error-valued records need a measured aggregation rule')
      if (valueScratch.kind === 'date-unmeasured') return unmeasured('shared date values need the workbook date-system context')
      const fn = functions[d]!
      const rowTargets = recordTargets.value.rows[d]![axes.rowKey[record]!]!
      const colTargets = recordTargets.value.cols[d]![axes.colKey[record]!]!
      add(grandTotals, d, valueScratch, fn)
      for (const row of rowTargets) {
        const rowLine = axes.rows[row]!
        add(rowTotals, row * dataCount + d, valueScratch, subtotalFunction(rowLine.type, fn))
      }
      for (const col of colTargets) {
        const colLine = axes.cols[col]!
        add(colTotals, col * dataCount + d, valueScratch, subtotalFunction(colLine.type, fn))
      }
      for (const row of rowTargets) for (const col of colTargets) {
        const rowLine = axes.rows[row]!, colLine = axes.cols[col]!
        const rowNamed = isNamedSubtotal(rowLine.type), colNamed = isNamedSubtotal(colLine.type)
        const rowFn = subtotalFunction(rowLine.type, fn), colFn = subtotalFunction(colLine.type, fn)
        if (rowNamed && colNamed && rowFn !== colFn)
          return unmeasured('intersecting row and column named subtotals with different functions are not measured')
        if ((rowNamed || colNamed) && view.data[d]!.showAs !== 'normal')
          return unmeasured('show-as applied to a named subtotal is not measured')
        const cellFn = rowNamed ? rowFn : colNamed ? colFn : fn
        add(gridAcc, cellIndex(row, col, d, colCount, dataCount), valueScratch, cellFn)
      }
    }
  }

  const raw = new Float64Array(gridSize), present = new Uint8Array(gridSize)
  for (let row = 0; row < rowCount; row++) for (let col = 0; col < colCount; col++) for (let d = 0; d < dataCount; d++) {
    const at = cellIndex(row, col, d, colCount, dataCount)
    const rowType = axes.rows[row]!.type, colType = axes.cols[col]!.type
    const rowNamed = isNamedSubtotal(rowType), colNamed = isNamedSubtotal(colType)
    const rowFn = subtotalFunction(rowType, view.data[d]!.fn), colFn = subtotalFunction(colType, view.data[d]!.fn)
    if (rowNamed && colNamed && rowFn !== colFn)
      return unmeasured('intersecting row and column named subtotals with different functions are not measured')
    const fn = rowNamed ? rowFn : colNamed ? colFn : view.data[d]!.fn
    const result = valueOf(gridAcc, at, fn)
    if (result === undefined) continue
    if (!Number.isFinite(result)) return unmeasured('aggregate overflow or sample variance with fewer than two numbers needs measured error behavior')
    raw[at] = result; present[at] = 1
  }

  // The [measure] rule on record-based totals is known for sums, but not for other functions. Keep those pivots
  // saved-only when a total line or total-based show-as needs the rule until Task 15 supplies an Excel oracle.
  const hasTotalLines = axes.rows.some(line => line.type !== 'data' && line.type !== 'blank')
    || axes.cols.some(line => line.type !== 'data' && line.type !== 'blank')
  for (let d = 0; d < dataCount; d++) {
    if (view.data[d]!.fn === 'sum') continue
    if (view.data[d]!.fn === 'average' && view.data[d]!.showAs === 'normal') continue
    if (hasTotalLines || TOTAL_SHOW_AS.has(view.data[d]!.showAs))
      return unmeasured(`record-based subtotal/grand-total behavior for ${view.data[d]!.fn} is not measured yet`)
  }

  const output = raw.slice(), outputPresent = present.slice()
  const totalValue = (acc: Accumulator, at: number, fn: PivotDataFunction): number | undefined => valueOf(acc, at, fn)
  const rowFields = view.rows, colFields = view.cols
  const findBaseLine = (axis: 'rows' | 'cols', current: number, baseField: number, baseItem: number): number | undefined => {
    const fields = axis === 'rows' ? rowFields : colFields, lines = axis === 'rows' ? axes.rows : axes.cols
    const level = fields.indexOf(baseField)
    if (level < 0 || level >= lines[current]!.items.length) return undefined
    const source = lines[current]!, path = source.items
    if (baseItem === PIVOT_BASE_ITEM_PREVIOUS || baseItem === PIVOT_BASE_ITEM_NEXT) {
      const siblings: number[] = []
      lines.forEach((line, index) => {
        if (line.type !== source.type || line.items.length !== path.length || !lineCarriesValues(view, axis, line)) return
        if (path.some((value, k) => k !== level && line.items[k] !== value)) return
        if (line.items[level] !== path[level]) siblings.push(index)
      })
      siblings.sort((a, b) => a - b)
      const priorCount = siblings.filter(index => index < current).length
      return baseItem === PIVOT_BASE_ITEM_PREVIOUS ? siblings[priorCount - 1] : siblings[priorCount]
    }
    const target = baseItem === PIVOT_BASE_ITEM_DEFAULT ? 0 : baseItem
    const found = lines.findIndex(line => line.type === source.type && line.items.length === path.length
      && line.items.every((value, k) => k === level ? value === target : value === path[k]))
    return found >= 0 ? found : undefined
  }

  const parentLine = (axis: 'rows' | 'cols', current: number, baseField?: number): number | undefined => {
    const fields = axis === 'rows' ? rowFields : colFields, lines = axis === 'rows' ? axes.rows : axes.cols
    const line = lines[current]!
    if (line.type === 'grand') return current
    const realLevels = fields.map((field, index) => field === PIVOT_VALUES_FIELD ? -1 : index).filter(index => index >= 0)
    const visibleLevels = realLevels.filter(index => index < line.items.length)
    if (!visibleLevels.length) return undefined
    const selectedOrdinal = baseField === undefined ? visibleLevels.length - 1 : realLevels.indexOf(fields.indexOf(baseField))
    if (selectedOrdinal < 0 || selectedOrdinal >= visibleLevels.length) return undefined
    const selectedDepth = selectedOrdinal + 1
    // At the selected base-field level the item is its own base, so its displayed share is 100%.
    if (baseField !== undefined && visibleLevels.length === selectedDepth) return current
    const parentDepth = baseField === undefined ? visibleLevels.length - 1 : selectedDepth
    if (parentDepth === 0) {
      const grand = lines.findIndex(candidate => candidate.type === 'grand' && candidate.data === line.data
        && lineCarriesValues(view, axis, candidate))
      return grand >= 0 ? grand : undefined
    }
    // Use the axis slot before the next real field; this retains a Values-field data item in the parent key.
    const pathEnd = realLevels[parentDepth]!
    const path = line.items.slice(0, pathEnd)
    let dataLine: number | undefined
    for (let index = 0; index < lines.length; index++) {
      const candidate = lines[index]!
      if (candidate.items.length !== path.length || candidate.type === 'blank' || !lineCarriesValues(view, axis, candidate)
        || candidate.items.some((value, k) => value !== path[k])) continue
      if (candidate.data !== line.data) continue
      if (candidate.type === 'default') return index
      if (candidate.type === 'data' && dataLine === undefined) dataLine = index
    }
    return dataLine
  }

  for (let row = 0; row < rowCount; row++) for (let col = 0; col < colCount; col++) for (let d = 0; d < dataCount; d++) {
    const at = cellIndex(row, col, d, colCount, dataCount)
    if (!present[at]) continue
    const data = view.data[d]!, mode = data.showAs as PivotShowDataAs | string
    if (mode === 'normal') continue
    let denominator: number | undefined, numerator = raw[at]!, result: number
    const rowTotal = totalValue(rowTotals, row * dataCount + d, data.fn)
    const colTotal = totalValue(colTotals, col * dataCount + d, data.fn)
    const grandTotal = totalValue(grandTotals, d, data.fn)
    if (mode === 'percentOfTotal') denominator = grandTotal
    else if (mode === 'percentOfRow') denominator = rowTotal
    else if (mode === 'percentOfCol') denominator = colTotal
    else if (mode === 'rankAscending' || mode === 'rankDescending') {
      const rowAxis = rowFields.includes(data.baseField), colAxis = colFields.includes(data.baseField)
      if (rowAxis === colAxis) return unmeasured('rank base field must identify exactly one row or column axis')
      const axis = rowAxis ? 'rows' : 'cols', lines = rowAxis ? axes.rows : axes.cols
      const current = rowAxis ? row : col, source = lines[current]!, level = (rowAxis ? rowFields : colFields).indexOf(data.baseField)
      if (source.type === 'grand' || level < 0 || level >= source.items.length) { outputPresent[at] = 0; continue }
      let rank = 1
      for (let sibling = 0; sibling < lines.length; sibling++) {
        if (sibling === current) continue
        const line = lines[sibling]!
        if (line.type !== source.type || line.items.length !== source.items.length || !lineCarriesValues(view, axis, line)) continue
        if (source.items.some((item, k) => k !== level && line.items[k] !== item)) continue
        const siblingAt = rowAxis ? cellIndex(sibling, col, d, colCount, dataCount) : cellIndex(row, sibling, d, colCount, dataCount)
        if (!present[siblingAt]) continue
        const siblingValue = raw[siblingAt]!
        if (siblingValue === numerator) return unmeasured('rank tie handling is not measured yet')
        if (mode === 'rankAscending' ? siblingValue < numerator : siblingValue > numerator) rank++
      }
      output[at] = rank
      continue
    }
    else if (mode === 'index') {
      const divisor = (rowTotal ?? 0) * (colTotal ?? 0)
      if (divisor === 0) return unmeasured('the index show-as zero denominator needs measured error behavior')
      result = numerator * (grandTotal ?? 0) / divisor
      output[at] = result
      continue
    } else if (mode === 'difference' || mode === 'percentDiff' || mode === 'percent') {
      const rowAxis = rowFields.includes(data.baseField), colAxis = colFields.includes(data.baseField)
      if (rowAxis === colAxis) return unmeasured('show-as base field must identify exactly one row or column axis')
      if (data.baseItem === PIVOT_BASE_ITEM_DEFAULT) return unmeasured('default base-item behavior is not measured yet')
      const axis = rowAxis ? 'rows' : 'cols', current = rowAxis ? row : col
      const base = findBaseLine(axis, current, data.baseField, data.baseItem)
      if (base === undefined) { outputPresent[at] = 0; continue }
      if (base === current && (mode === 'difference' || mode === 'percentDiff'))
        return unmeasured('difference-from values for the selected base item are not measured')
      const baseAt = rowAxis ? cellIndex(base, col, d, colCount, dataCount) : cellIndex(row, base, d, colCount, dataCount)
      if (!present[baseAt]) { outputPresent[at] = 0; continue }
      denominator = raw[baseAt]!
      if (mode === 'difference') { output[at] = numerator - denominator; continue }
      if (denominator === 0) return unmeasured('show-as division by zero needs measured error behavior')
      output[at] = mode === 'percent' ? numerator / denominator : (numerator - denominator) / denominator
      continue
    } else if (mode === 'percentOfParent' || mode === 'percentOfParentRow' || mode === 'percentOfParentCol') {
      const axis = mode === 'percentOfParentCol' ? 'cols' : 'rows'
      let baseField: number | undefined
      if (mode === 'percentOfParent' && data.baseField !== -1) baseField = data.baseField
      const parent = parentLine(axis, axis === 'rows' ? row : col, baseField)
      if (parent === undefined) return unmeasured('parent-total behavior for this line has no measured parent')
      const parentAt = axis === 'rows' ? cellIndex(parent, col, d, colCount, dataCount) : cellIndex(row, parent, d, colCount, dataCount)
      if (!present[parentAt]) { outputPresent[at] = 0; continue }
      denominator = raw[parentAt]!
    } else if (mode === 'runTotal' || mode === 'percentOfRunningTotal') {
      const rowAxis = rowFields.includes(data.baseField), colAxis = colFields.includes(data.baseField)
      if (rowAxis === colAxis) return unmeasured('running-total base field must identify exactly one row or column axis')
      const axis = rowAxis ? 'rows' : 'cols', lines = rowAxis ? axes.rows : axes.cols
      const current = rowAxis ? row : col, currentLine = lines[current]!, level = (rowAxis ? rowFields : colFields).indexOf(data.baseField)
      if (level < 0 || level >= currentLine.items.length) return unmeasured('running-total base field is not present on the value line')
      const siblings: number[] = []
      lines.forEach((line, index) => {
        if (line.type !== currentLine.type || line.items.length !== currentLine.items.length || !lineCarriesValues(view, axis, line)) return
        if (currentLine.items.some((item, k) => k !== level && item !== line.items[k])) return
        siblings.push(index)
      })
      let cumulative = 0, total = 0
      for (const sibling of siblings) {
        const sourceAt = rowAxis ? cellIndex(sibling, col, d, colCount, dataCount) : cellIndex(row, sibling, d, colCount, dataCount)
        if (present[sourceAt]) {
          total += raw[sourceAt]!
          if (sibling <= current) cumulative += raw[sourceAt]!
        }
      }
      if (mode === 'runTotal') { output[at] = cumulative; continue }
      numerator = cumulative
      denominator = total
    } else return fail('pivot-engine-pending', `show-as mode ${mode} is not implemented`)

    if (denominator === undefined) { outputPresent[at] = 0; continue }
    if (denominator === 0) return unmeasured('show-as division by zero needs measured error behavior')
    output[at] = numerator / denominator
  }

  for (let at = 0; at < output.length; at++) if (outputPresent[at] && !Number.isFinite(output[at]!))
    return unmeasured('show-as result overflow is not measured yet')

  const grid: PivotValueGrid = {
    get(row, col, data): PivotCellValue {
      if (!Number.isInteger(row) || !Number.isInteger(col) || !Number.isInteger(data) || row < 0 || col < 0 || data < 0
        || row >= rowCount || col >= colCount || data >= dataCount) return { kind: 'empty' }
      const at = cellIndex(row, col, data, colCount, dataCount)
      return outputPresent[at] ? { kind: 'number', value: output[at]! } : { kind: 'empty' }
    },
  }
  return { ok: true, value: grid }
}
