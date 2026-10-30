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
// One computeValues call is capped at 128 MiB, one quarter of the 512 MiB
// desktop worker pool, leaving room for axes, source pages and other worker jobs.
const OUTPUT_BYTES_PER_SLOT = 9
const TOTAL_SHOW_AS = new Set(['percentOfTotal', 'percentOfRow', 'percentOfCol', 'index', 'percentOfParent', 'percentOfParentRow', 'percentOfParentCol', 'runTotal', 'percentOfRunningTotal'])

const fail = (feature: string, message: string): EngineResult<never> => ({
  ok: false, diagnostic: { kind: 'unsupported', feature, message } satisfies PivotDiagnostic,
})
const unmeasured = (message: string) => fail('pivot-engine-unmeasured', message)
const invalid = (message: string) => ({ ok: false as const, diagnostic: { kind: 'parse' as const, feature: 'pivot-values', message } })

interface Accumulator {
  nonEmpty?: Uint32Array
  numeric?: Uint32Array
  sum?: Float64Array
  product?: Float64Array
  mean?: Float64Array
  m2?: Float64Array
  min?: Float64Array
  max?: Float64Array
}

// Keep the record loop's aggregation dispatch numeric and monomorphic. The
// string-valued PivotDataFunction is useful at the API boundary, but repeated
// string comparisons for every record/line pair dominate large value scans.
const VALUE_FN = {
  average: 0, count: 1, countNums: 2, max: 3, min: 4, product: 5,
  stdDev: 6, stdDevp: 7, sum: 8, var: 9, varp: 10,
} as const
type ValueFnCode = typeof VALUE_FN[keyof typeof VALUE_FN]

function valueFnCode(fn: PivotDataFunction): ValueFnCode { return VALUE_FN[fn] }

function subtotalFnCode(type: PivotItemType, fallback: ValueFnCode): ValueFnCode {
  switch (type) {
    case 'sum': return VALUE_FN.sum
    case 'max': return VALUE_FN.max
    case 'min': return VALUE_FN.min
    case 'product': return VALUE_FN.product
    case 'stdDev': return VALUE_FN.stdDev
    case 'var': return VALUE_FN.var
    case 'avg': return VALUE_FN.average
    case 'countA': return VALUE_FN.count
    case 'count': return VALUE_FN.countNums
    case 'stdDevP': return VALUE_FN.stdDevp
    case 'varP': return VALUE_FN.varp
    default: return fallback
  }
}

function accumulatorBytes(fn: ValueFnCode): number {
  switch (fn) {
    case VALUE_FN.sum: return 8
    case VALUE_FN.count: case VALUE_FN.countNums: return 4
    case VALUE_FN.average: case VALUE_FN.max: case VALUE_FN.min: case VALUE_FN.product: return 12
    default: return 20 // numeric count, Welford mean and M2
  }
}

function accumulator(size: number, fn: ValueFnCode): Accumulator {
  const result: Accumulator = {}
  switch (fn) {
    case VALUE_FN.sum: result.sum = new Float64Array(size); break
    case VALUE_FN.count: result.nonEmpty = new Uint32Array(size); break
    case VALUE_FN.countNums: result.numeric = new Uint32Array(size); break
    case VALUE_FN.average:
      result.sum = new Float64Array(size); result.numeric = new Uint32Array(size); break
    case VALUE_FN.max:
      result.max = new Float64Array(size); result.max.fill(Number.NEGATIVE_INFINITY)
      result.numeric = new Uint32Array(size); break
    case VALUE_FN.min:
      result.min = new Float64Array(size); result.min.fill(Number.POSITIVE_INFINITY)
      result.numeric = new Uint32Array(size); break
    case VALUE_FN.product:
      result.product = new Float64Array(size); result.product.fill(1)
      result.numeric = new Uint32Array(size); break
    default:
      result.mean = new Float64Array(size); result.m2 = new Float64Array(size); result.numeric = new Uint32Array(size)
  }
  return result
}

type AccumulatorBank = Array<Accumulator | undefined>
function accumulatorBank(size: number, codes: readonly number[]): AccumulatorBank {
  const bank: AccumulatorBank = new Array(11)
  for (const code of codes) bank[code] = accumulator(size, code as ValueFnCode)
  return bank
}

function accumulatorBankBytes(size: number, codes: readonly number[]): number {
  let bytes = 0
  for (const code of codes) bytes += size * accumulatorBytes(code as ValueFnCode)
  return bytes
}

function add(acc: Accumulator, present: Uint8Array, at: number, value: ReadValue, fn: ValueFnCode): void {
  present[at] = 1
  if (value.kind === 'missing') return
  if (fn === VALUE_FN.count) { acc.nonEmpty![at]++; return }
  if (value.kind !== 'number') return
  if (fn === VALUE_FN.sum || fn === VALUE_FN.average) acc.sum![at] += value.value
  if (fn === VALUE_FN.sum) return
  if (fn === VALUE_FN.average) { acc.numeric![at]++; return }
  if (fn === VALUE_FN.countNums) { acc.numeric![at]++; return }
  if (fn === VALUE_FN.max) { acc.numeric![at]++; if (value.value > acc.max![at]!) acc.max![at] = value.value; return }
  if (fn === VALUE_FN.min) { acc.numeric![at]++; if (value.value < acc.min![at]!) acc.min![at] = value.value; return }
  if (fn === VALUE_FN.product) { acc.numeric![at]++; acc.product![at] *= value.value; return }
  const n = ++acc.numeric![at]
  const delta = value.value - acc.mean![at]!
  acc.mean![at] += delta / n
  if (fn === VALUE_FN.stdDev || fn === VALUE_FN.stdDevp || fn === VALUE_FN.var || fn === VALUE_FN.varp)
    acc.m2![at] += delta * (value.value - acc.mean![at]!)
}

interface ReadValue { kind: 'missing' | 'other' | 'error' | 'date-unmeasured' | 'non-finite' | 'number'; value: number }
function readValue(column: RecordColumn | undefined, sharedItems: readonly import('../model').PivotCacheItem[] | undefined,
  row: number, into: ReadValue): boolean {
  if (!column) return false
  if (column.kind === 'shared') {
    const item = sharedItems?.[column.index[row]!]
    if (!item) return false
    switch (item.type) {
      case 'missing': into.kind = 'missing'; break
      case 'number': into.value = item.value; into.kind = Number.isFinite(item.value) ? 'number' : 'non-finite'; break
      case 'date': into.kind = 'date-unmeasured'; break
      case 'error': into.kind = 'error'; break
      case 'string': case 'boolean': into.kind = 'other'; break
    }
    return true
  }
  const tag = column.tag[row]
  switch (tag) {
    case VALUE_TAG.missing: into.kind = 'missing'; break
    case VALUE_TAG.number: case VALUE_TAG.date:
      into.value = column.num[row]!
      into.kind = Number.isFinite(into.value) ? 'number' : 'non-finite'
      break
    case VALUE_TAG.boolean: case VALUE_TAG.string: into.kind = 'other'; break
    case VALUE_TAG.error: into.kind = 'error'; break
    default: return false
  }
  return true
}

function valueOf(acc: Accumulator, present: Uint8Array, at: number, fn: ValueFnCode): number | undefined {
  const count = acc.numeric?.[at] ?? 0
  if (!present[at]) return undefined
  switch (fn) {
    case VALUE_FN.sum: return acc.sum![at]!
    case VALUE_FN.count: return acc.nonEmpty![at]!
    case VALUE_FN.countNums: return count
    case VALUE_FN.average: return count ? acc.sum![at]! / count : Number.NaN
    case VALUE_FN.max: return count ? acc.max![at]! : Number.NaN
    case VALUE_FN.min: return count ? acc.min![at]! : Number.NaN
    case VALUE_FN.product: return count ? acc.product![at]! : Number.NaN
    case VALUE_FN.stdDev: return count > 1 ? Math.sqrt(acc.m2![at]! / (count - 1)) : Number.NaN
    case VALUE_FN.stdDevp: return count ? Math.sqrt(acc.m2![at]! / count) : Number.NaN
    case VALUE_FN.var: return count > 1 ? acc.m2![at]! / (count - 1) : Number.NaN
    case VALUE_FN.varp: return count ? acc.m2![at]! / count : Number.NaN
  }
}

/** Aggregate a data field by caller-provided record keys using the engine's existing accumulator semantics. */
export function aggregateByKey(store: RecordStore, cache: PivotCacheDefinition, view: PivotViewState, dataIndex: number,
  keys: Uint32Array, keyCount: number, passes: Uint8Array, functionOverride?: PivotDataFunction): EngineResult<Array<number | undefined>> {
  if (!view.data[dataIndex] || keys.length !== store.count || passes.length !== store.count)
    return invalid('filter aggregate inputs do not match the pivot view')
  const data = view.data[dataIndex]!, sourceField = data.field
  const column = store.columns[sourceField], shared = cache.fields[sourceField]?.sharedItems?.items
  const fn = valueFnCode(functionOverride ?? data.fn), acc = accumulator(keyCount, fn), present = new Uint8Array(keyCount)
  const scratch: ReadValue = { kind: 'missing', value: 0 }
  for (let record = 0; record < store.count; record++) {
    if (!passes[record]) continue
    const key = keys[record]!
    if (key >= keyCount) return invalid(`filter aggregate record ${record} has an invalid item key`)
    if (!readValue(column, shared, record, scratch)) return invalid(`record ${record} has an invalid value in field ${sourceField}`)
    if (scratch.kind === 'error') return unmeasured('error-valued records need a measured aggregation rule')
    if (scratch.kind === 'date-unmeasured') return unmeasured('shared date values need the workbook date-system context')
    if (scratch.kind === 'non-finite') return unmeasured('aggregate overflow or sample variance with fewer than two numbers needs measured error behavior')
    add(acc, present, key, scratch, fn)
  }
  const result = new Array<number | undefined>(keyCount)
  for (let key = 0; key < keyCount; key++) {
    const value = valueOf(acc, present, key, fn)
    if (value !== undefined && !Number.isFinite(value)) return unmeasured('aggregate overflow or sample variance with fewer than two numbers needs measured error behavior')
    result[key] = value
  }
  return { ok: true, value: result }
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

interface LineTargets { offsets: Int32Array; indices: Int32Array }
interface Targets { rows: LineTargets[]; cols: LineTargets[] }
interface PrefixNode { children: Map<number, PrefixNode>; lines: number[] }
interface LineIndex { root: PrefixNode; grand: number[]; grandByData: Map<number, number[]> }
function makePrefixIndex(lines: readonly PivotLine[], view: PivotViewState, axis: 'rows' | 'cols'): LineIndex {
  const root: PrefixNode = { children: new Map(), lines: [] }, grand: number[] = [], grandByData = new Map<number, number[]>()
  lines.forEach((line, index) => {
    if (line.type === 'grand') {
      grand.push(index)
      let data = grandByData.get(line.data)
      if (!data) grandByData.set(line.data, data = [])
      data.push(index)
      return
    }
    if (!lineCarriesValues(view, axis, line)) return
    let node = root
    for (const item of line.items) {
      let child = node.children.get(item)
      if (!child) node.children.set(item, child = { children: new Map(), lines: [] })
      node = child
    }
    node.lines.push(index)
  })
  return { root, grand, grandByData }
}

function readTargets(store: RecordStore, view: PivotViewState, axes: PivotAxes): EngineResult<Targets> {
  for (let record = 0; record < store.count; record++) {
    const row = axes.rowKey[record], col = axes.colKey[record]
    if (row === undefined || col === undefined) return invalid(`record ${record} has an invalid leaf line`)
    if (row === NO_LINE || col === NO_LINE) return unmeasured('filtered-out records require measured visible-total semantics')
    if (row >= axes.rows.length || col >= axes.cols.length) return invalid(`record ${record} has an invalid leaf line`)
  }
  const rowLineIndex = makePrefixIndex(axes.rows, view, 'rows'), colLineIndex = makePrefixIndex(axes.cols, view, 'cols')
  const lists = (lines: readonly PivotLine[], index: LineIndex, axis: 'rows' | 'cols', data: number): LineTargets => {
    const offsets = new Int32Array(lines.length + 1), values: number[] = []
    const fields = axis === 'rows' ? view.rows : view.cols
    const valuesLevel = fields.indexOf(PIVOT_VALUES_FIELD)
    const append = (bucket: readonly number[]) => {
      for (const target of bucket) {
        if (valuesLevel < 0 || lines[target]!.data === data) values.push(target)
      }
    }
    for (let line = 0; line < lines.length; line++) {
      const leaf = lines[line]!
      let node: PrefixNode | undefined = index.root
      append(node.lines)
      for (let level = 0; level < leaf.items.length; level++) {
        const item = valuesLevel === level ? data : leaf.items[level]!
        node = node.children.get(item)
        if (!node) break
        append(node.lines)
      }
      append(valuesLevel < 0 ? index.grand : index.grandByData.get(data) ?? [])
      offsets[line + 1] = values.length
    }
    return { offsets, indices: Int32Array.from(values) }
  }
  const rows: LineTargets[] = [], cols: LineTargets[] = []
  const rowShared = !view.rows.includes(PIVOT_VALUES_FIELD), colShared = !view.cols.includes(PIVOT_VALUES_FIELD)
  for (let d = 0; d < view.data.length; d++) {
    rows.push(rowShared && d > 0 ? rows[0]! : lists(axes.rows, rowLineIndex, 'rows', d))
    cols.push(colShared && d > 0 ? cols[0]! : lists(axes.cols, colLineIndex, 'cols', d))
  }
  if (rowShared && colShared) {
    const rowList = rows[0]!, colList = cols[0]!
    for (let record = 0; record < store.count; record++) {
      const row = axes.rowKey[record]!, col = axes.colKey[record]!
      if (rowList.offsets[row] === rowList.offsets[row + 1] || colList.offsets[col] === colList.offsets[col + 1])
        return invalid(`record ${record} has no value-carrying row or column line`)
    }
  } else for (let record = 0; record < store.count; record++) for (let d = 0; d < view.data.length; d++) {
    const row = axes.rowKey[record]!, col = axes.colKey[record]!
    if (rows[d]!.offsets[row] === rows[d]!.offsets[row + 1] || cols[d]!.offsets[col] === cols[d]!.offsets[col + 1])
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
  const gridCellCount = rowCount * colCount
  // Reject clearly oversized requests before scanning lines to discover the
  // active function banks. Every accumulator cell needs at least one byte of
  // presence plus a four-byte count; output and target structures remain.
  const targetDepth = Math.max(view.rows.length, view.cols.length)
  // Per target/data field, cover retained Int32 indices and the temporary
  // number[] used to build them; deeper axes retain more ancestor targets.
  const targetBytes = (rowCount + colCount) * (128 + dataCount * Math.max(32, (targetDepth + 3) * 12))
  const minimumBytes = (gridCellCount + rowCount + colCount + 1) * dataCount * 5
    + gridSize * OUTPUT_BYTES_PER_SLOT * 2 + targetBytes
  if (!Number.isSafeInteger(minimumBytes) || minimumBytes > MAX_ACCUMULATOR_BYTES)
    return fail('pivot-values-limit', 'the value grid exceeds its bounded accumulator budget')

  const functions = view.data.map(data => valueFnCode(data.fn))
  const codesFor = (lines: readonly PivotLine[], d: number) => {
    const seen = new Uint8Array(11); seen[functions[d]!] = 1
    for (const line of lines) seen[subtotalFnCode(line.type, functions[d]!)] = 1
    const codes: number[] = []
    for (let code = 0; code < seen.length; code++) if (seen[code]) codes.push(code)
    return codes
  }
  const rowCodes = functions.map((_, d) => codesFor(axes.rows, d))
  const colCodes = functions.map((_, d) => codesFor(axes.cols, d))
  const gridCodes = functions.map((fn, d) => [...new Set([fn, ...rowCodes[d]!, ...colCodes[d]!])])
  const grandCodes = functions.map(fn => [fn])
  const unionCodes = (sets: readonly (readonly number[])[]) => [...new Set(sets.flat())]
  const gridBankCodes = unionCodes(gridCodes), rowBankCodes = unionCodes(rowCodes)
  const colBankCodes = unionCodes(colCodes), grandBankCodes = unionCodes(grandCodes)
  // The byte estimate includes per-line target lists and prefix-trie planning,
  // whose working set grows with the number of axis lines.
  const estimatedBytes = gridSize + rowCount * dataCount + colCount * dataCount + dataCount
    + accumulatorBankBytes(gridSize, gridBankCodes)
    + accumulatorBankBytes(rowCount * dataCount, rowBankCodes) + accumulatorBankBytes(colCount * dataCount, colBankCodes)
    + accumulatorBankBytes(dataCount, grandBankCodes) + gridSize * OUTPUT_BYTES_PER_SLOT * 2 + targetBytes
  if (!Number.isSafeInteger(estimatedBytes) || estimatedBytes > MAX_ACCUMULATOR_BYTES)
    return fail('pivot-values-limit', 'the value grid exceeds its bounded accumulator budget')

  const rowFunctions = new Uint8Array(rowCount * dataCount), colFunctions = new Uint8Array(colCount * dataCount)
  const rowNamed = new Uint8Array(rowCount), colNamed = new Uint8Array(colCount)
  for (let row = 0; row < rowCount; row++) {
    const line = axes.rows[row]!
    rowNamed[row] = +isNamedSubtotal(line.type)
    for (let d = 0; d < dataCount; d++) rowFunctions[row * dataCount + d] = subtotalFnCode(line.type, functions[d]!)
  }
  for (let col = 0; col < colCount; col++) {
    const line = axes.cols[col]!
    colNamed[col] = +isNamedSubtotal(line.type)
    for (let d = 0; d < dataCount; d++) colFunctions[col * dataCount + d] = subtotalFnCode(line.type, functions[d]!)
  }

  const recordTargets = readTargets(store, view, axes)
  if (!recordTargets.ok) return recordTargets

  const sourceColumns = view.data.map(data => store.columns[data.field])
  const sourceItems = view.data.map(data => cache.fields[data.field]?.sharedItems?.items)
  const sumOnly = functions.every(fn => fn === VALUE_FN.sum)
    && gridBankCodes.length === 1 && gridBankCodes[0] === VALUE_FN.sum
    && rowBankCodes.length === 1 && rowBankCodes[0] === VALUE_FN.sum
    && colBankCodes.length === 1 && colBankCodes[0] === VALUE_FN.sum
    && grandBankCodes.length === 1 && grandBankCodes[0] === VALUE_FN.sum
  const normalOnly = view.data.every(data => data.showAs === 'normal')
  if (sumOnly && !normalOnly && (rowNamed.includes(1) || colNamed.includes(1)))
    return unmeasured('show-as applied to a named subtotal is not measured')
  if (sumOnly && normalOnly) {
    // NaN is a private empty-cell sentinel (record values are finite); this
    // folds sum presence into the value array and keeps the hot update to one
    // typed-array read/write.
    const sums = new Float64Array(gridSize); sums.fill(Number.NaN)
    const scratch: ReadValue = { kind: 'missing', value: 0 }
    const commonTargets = !view.rows.includes(PIVOT_VALUES_FIELD) && !view.cols.includes(PIVOT_VALUES_FIELD)
    if (commonTargets) {
      const dataValues = new Float64Array(dataCount), dataNumeric = new Uint8Array(dataCount)
      const inlineColumns = sourceColumns.map(column => column?.kind === 'inline' ? column : undefined)
      const rowList = recordTargets.value.rows[0]!, colList = recordTargets.value.cols[0]!
      for (let record = 0; record < store.count; record++) {
        const rowLine = axes.rowKey[record]!, colLine = axes.colKey[record]!
        for (let d = 0; d < dataCount; d++) {
          const field = view.data[d]!.field
          const inline = inlineColumns[d]
          const tag = inline?.tag[record]
          if (inline && (tag === VALUE_TAG.number || tag === VALUE_TAG.date)) {
            dataNumeric[d] = 1
            dataValues[d] = inline.num[record]!
          } else {
            if (!readValue(sourceColumns[d], sourceItems[d], record, scratch)) return invalid(`record ${record} has an invalid value in field ${field}`)
            if (scratch.kind === 'error') return unmeasured('error-valued records need a measured aggregation rule')
            if (scratch.kind === 'date-unmeasured') return unmeasured('shared date values need the workbook date-system context')
            dataNumeric[d] = +(scratch.kind === 'number')
            dataValues[d] = scratch.value
            if (scratch.kind === 'non-finite') return unmeasured('aggregate overflow or sample variance with fewer than two numbers needs measured error behavior')
          }
          if (dataNumeric[d] && !Number.isFinite(dataValues[d]!))
            return unmeasured('aggregate overflow or sample variance with fewer than two numbers needs measured error behavior')
        }
        const rowStart = rowList.offsets[rowLine]!, rowEnd = rowList.offsets[rowLine + 1]!
        const colStart = colList.offsets[colLine]!, colEnd = colList.offsets[colLine + 1]!
        for (let r = rowStart; r < rowEnd; r++) {
          const row = rowList.indices[r]!, rowOffset = row * colCount
          for (let c = colStart; c < colEnd; c++) {
            const col = colList.indices[c]!, base = (rowOffset + col) * dataCount
            for (let d = 0; d < dataCount; d++) {
              const at = base + d, current = sums[at]!
              if (dataNumeric[d]) sums[at] = (Number.isNaN(current) ? 0 : current) + dataValues[d]!
              else if (Number.isNaN(current)) sums[at] = 0
            }
          }
        }
      }
    } else {
      for (let record = 0; record < store.count; record++) {
        const rowLine = axes.rowKey[record]!, colLine = axes.colKey[record]!
        for (let d = 0; d < dataCount; d++) {
          const field = view.data[d]!.field
          if (!readValue(sourceColumns[d], sourceItems[d], record, scratch)) return invalid(`record ${record} has an invalid value in field ${field}`)
          if (scratch.kind === 'error') return unmeasured('error-valued records need a measured aggregation rule')
          if (scratch.kind === 'date-unmeasured') return unmeasured('shared date values need the workbook date-system context')
          if (scratch.kind === 'non-finite') return unmeasured('aggregate overflow or sample variance with fewer than two numbers needs measured error behavior')
          const numeric = scratch.kind === 'number', value = scratch.value
          const rowList = recordTargets.value.rows[d]!, rowStart = rowList.offsets[rowLine]!, rowEnd = rowList.offsets[rowLine + 1]!
          const colList = recordTargets.value.cols[d]!, colStart = colList.offsets[colLine]!, colEnd = colList.offsets[colLine + 1]!
          for (let r = rowStart; r < rowEnd; r++) {
            const row = rowList.indices[r]!
            for (let c = colStart; c < colEnd; c++) {
              const col = colList.indices[c]!, at = (row * colCount + col) * dataCount + d
              const current = sums[at]!
              if (numeric) sums[at] = (Number.isNaN(current) ? 0 : current) + value
              else if (Number.isNaN(current)) sums[at] = 0
            }
          }
        }
      }
    }
    for (let at = 0; at < gridSize; at++) if (!Number.isNaN(sums[at]!) && !Number.isFinite(sums[at]!))
      return unmeasured('aggregate overflow or sample variance with fewer than two numbers needs measured error behavior')
    const grid: PivotValueGrid = {
      get(row, col, data): PivotCellValue {
        if (!Number.isInteger(row) || !Number.isInteger(col) || !Number.isInteger(data) || row < 0 || col < 0 || data < 0
          || row >= rowCount || col >= colCount || data >= dataCount) return { kind: 'empty' }
        const at = (row * colCount + col) * dataCount + data
        const value = sums[at]!
        return Number.isNaN(value) ? { kind: 'empty' } : { kind: 'number', value }
      },
    }
    return { ok: true, value: grid }
  }

  const gridPresent = new Uint8Array(gridSize), rowPresent = new Uint8Array(rowCount * dataCount)
  const colPresent = new Uint8Array(colCount * dataCount), grandPresent = new Uint8Array(dataCount)
  const gridAcc = accumulatorBank(gridSize, gridBankCodes), rowTotals = accumulatorBank(rowCount * dataCount, rowBankCodes)
  const colTotals = accumulatorBank(colCount * dataCount, colBankCodes), grandTotals = accumulatorBank(dataCount, grandBankCodes)
  const valueScratch: ReadValue = { kind: 'missing', value: 0 }
  if (sumOnly) {
    const gridSum = gridAcc[VALUE_FN.sum]!.sum!, rowSum = rowTotals[VALUE_FN.sum]!.sum!
    const colSum = colTotals[VALUE_FN.sum]!.sum!, grandSum = grandTotals[VALUE_FN.sum]!.sum!
    for (let record = 0; record < store.count; record++) {
      const rowLine = axes.rowKey[record]!, colLine = axes.colKey[record]!
      for (let d = 0; d < dataCount; d++) {
        const field = view.data[d]!.field
        if (!readValue(sourceColumns[d], sourceItems[d], record, valueScratch)) return invalid(`record ${record} has an invalid value in field ${field}`)
        if (valueScratch.kind === 'error') return unmeasured('error-valued records need a measured aggregation rule')
        if (valueScratch.kind === 'date-unmeasured') return unmeasured('shared date values need the workbook date-system context')
        const numeric = valueScratch.kind === 'number', value = valueScratch.value
        grandPresent[d] = 1
        if (numeric) grandSum[d] += value
        const rowList = recordTargets.value.rows[d]!, rowStart = rowList.offsets[rowLine]!, rowEnd = rowList.offsets[rowLine + 1]!
        for (let r = rowStart; r < rowEnd; r++) {
          const row = rowList.indices[r]!, at = row * dataCount + d
          rowPresent[at] = 1
          if (numeric) rowSum[at] += value
        }
        const colList = recordTargets.value.cols[d]!, colStart = colList.offsets[colLine]!, colEnd = colList.offsets[colLine + 1]!
        for (let c = colStart; c < colEnd; c++) {
          const col = colList.indices[c]!, at = col * dataCount + d
          colPresent[at] = 1
          if (numeric) colSum[at] += value
        }
        for (let r = rowStart; r < rowEnd; r++) {
          const row = rowList.indices[r]!
          for (let c = colStart; c < colEnd; c++) {
            const col = colList.indices[c]!, at = (row * colCount + col) * dataCount + d
            gridPresent[at] = 1
            if (numeric) gridSum[at] += value
          }
        }
      }
    }
  } else {
    for (let record = 0; record < store.count; record++) {
      for (let d = 0; d < dataCount; d++) {
        const field = view.data[d]!.field
        if (!readValue(sourceColumns[d], sourceItems[d], record, valueScratch)) return invalid(`record ${record} has an invalid value in field ${field}`)
        if (valueScratch.kind === 'error') return unmeasured('error-valued records need a measured aggregation rule')
        if (valueScratch.kind === 'date-unmeasured') return unmeasured('shared date values need the workbook date-system context')
        const fn = functions[d]!
        const rowLine = axes.rowKey[record]!, colLine = axes.colKey[record]!
        const rowList = recordTargets.value.rows[d]!, colList = recordTargets.value.cols[d]!
        const rowStart = rowList.offsets[rowLine]!, rowEnd = rowList.offsets[rowLine + 1]!
        const colStart = colList.offsets[colLine]!, colEnd = colList.offsets[colLine + 1]!
        add(grandTotals[fn]!, grandPresent, d, valueScratch, fn)
        for (let r = rowStart; r < rowEnd; r++) {
          const row = rowList.indices[r]!
          const rowAt = row * dataCount + d, rowFn = rowFunctions[rowAt]! as ValueFnCode
          add(rowTotals[rowFn]!, rowPresent, rowAt, valueScratch, rowFn)
        }
        for (let c = colStart; c < colEnd; c++) {
          const col = colList.indices[c]!
          const colAt = col * dataCount + d, colFn = colFunctions[colAt]! as ValueFnCode
          add(colTotals[colFn]!, colPresent, colAt, valueScratch, colFn)
        }
        for (let r = rowStart; r < rowEnd; r++) for (let c = colStart; c < colEnd; c++) {
          const row = rowList.indices[r]!, col = colList.indices[c]!
          const isRowNamed = !!rowNamed[row], isColNamed = !!colNamed[col]
          const rowFn = rowFunctions[row * dataCount + d]! as ValueFnCode
          const colFn = colFunctions[col * dataCount + d]! as ValueFnCode
          if (isRowNamed && isColNamed && rowFn !== colFn)
            return unmeasured('intersecting row and column named subtotals with different functions are not measured')
          if ((isRowNamed || isColNamed) && view.data[d]!.showAs !== 'normal')
            return unmeasured('show-as applied to a named subtotal is not measured')
          const cellFn = isRowNamed ? rowFn : isColNamed ? colFn : fn
          const at = cellIndex(row, col, d, colCount, dataCount)
          add(gridAcc[cellFn]!, gridPresent, at, valueScratch, cellFn)
        }
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
    const fnCode = valueFnCode(fn)
    const result = valueOf(gridAcc[fnCode]!, gridPresent, at, fnCode)
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
  const totalValue = (acc: AccumulatorBank, present: Uint8Array, at: number, d: number): number | undefined => {
    const fn = functions[d]!
    return valueOf(acc[fn]!, present, at, fn)
  }
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
    const rowTotal = totalValue(rowTotals, rowPresent, row * dataCount + d, d)
    const colTotal = totalValue(colTotals, colPresent, col * dataCount + d, d)
    const grandTotal = totalValue(grandTotals, grandPresent, d, d)
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
