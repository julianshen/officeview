/** Pivot value aggregation and show-values-as post-pass (Stage 2, library only). */
import {
  PIVOT_BASE_ITEM_DEFAULT, PIVOT_BASE_ITEM_NEXT, PIVOT_BASE_ITEM_PREVIOUS, PIVOT_VALUES_FIELD,
  type PivotCacheDefinition, type PivotDataFunction, type PivotDiagnostic, type PivotItemType, type PivotShowDataAs, type PivotTableDefinition,
} from '../model'
import type { PivotLine } from '../savedLayout'
import { lineCarriesValues, NO_LINE, type EngineResult, type PivotAxes } from './axes'
import { VALUE_TAG, type RecordColumn, type RecordStore } from './records'
import type { PivotCellValue, PivotValueGrid } from './place'
import type { DataFieldState, PivotViewState } from './view'
import { calculatedValue, compileCalculatedField, type CalculatedField } from './calculated'

const MAX_ACCUMULATOR_BYTES = 128 * 1024 * 1024
// One computeValues call is capped at 128 MiB, one quarter of the 512 MiB
// desktop worker pool, leaving room for axes, source pages and other worker jobs.
const OUTPUT_BYTES_PER_SLOT = 9

const fail = (feature: string, message: string): EngineResult<never> => ({
  ok: false, diagnostic: { kind: 'unsupported', feature, message } satisfies PivotDiagnostic,
})
const unmeasured = (message: string) => fail('pivot-engine-unmeasured', message)
const invalid = (message: string) => ({ ok: false as const, diagnostic: { kind: 'parse' as const, feature: 'pivot-values', message } })

interface Accumulator {
  nonEmpty?: Uint32Array
  numeric?: Uint32Array
  sum?: Float64Array
  sumSquares?: Float64Array
  product?: Float64Array
  min?: Float64Array
  max?: Float64Array
  error?: Array<string | undefined>
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
    case VALUE_FN.sum: return 12
    case VALUE_FN.count: return 4
    case VALUE_FN.countNums: return 8
    case VALUE_FN.average: case VALUE_FN.max: case VALUE_FN.min: case VALUE_FN.product: return 12
    case VALUE_FN.stdDev: case VALUE_FN.stdDevp: case VALUE_FN.var: case VALUE_FN.varp: return 20
    default: return 4
  }
}

function accumulator(size: number, fn: ValueFnCode): Accumulator {
  const result: Accumulator = {}
  switch (fn) {
    case VALUE_FN.sum: result.sum = new Float64Array(size); result.numeric = new Uint32Array(size); break
    case VALUE_FN.count: result.nonEmpty = new Uint32Array(size); break
    case VALUE_FN.countNums: result.numeric = new Uint32Array(size); result.nonEmpty = new Uint32Array(size); break
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
      result.sum = new Float64Array(size); result.sumSquares = new Float64Array(size); result.numeric = new Uint32Array(size)
  }
  result.error = new Array<string | undefined>(size)
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
  // Excel Count includes error-valued cells; numeric functions ignore them and
  // Sum alone propagates the source error (Task 18 quirk-error-value-count R19C4).
  if (value.kind === 'error') {
    if (fn === VALUE_FN.count || fn === VALUE_FN.countNums) acc.nonEmpty![at]++
    if (fn === VALUE_FN.count || fn === VALUE_FN.sum) acc.error![at] = value.text
    return
  }
  if (value.kind === 'missing') return
  if (fn === VALUE_FN.count) { acc.nonEmpty![at]++; return }
  // Count Numbers is shown (0) for a group with only text values and blank for
  // one with no values at all (Task 18 values-fn-countnums-mixed: R6C4 East/2022
  // is 0, R21C4 West/Cherries/2022, whose Channel is blank, is empty).
  if (fn === VALUE_FN.countNums) acc.nonEmpty![at]++
  if (value.kind !== 'number') return
  acc.numeric![at]++
  switch (fn) {
    case VALUE_FN.sum: case VALUE_FN.average: acc.sum![at] += value.value; return
    case VALUE_FN.countNums: return
    case VALUE_FN.max: if (value.value > acc.max![at]!) acc.max![at] = value.value; return
    case VALUE_FN.min: if (value.value < acc.min![at]!) acc.min![at] = value.value; return
    case VALUE_FN.product: acc.product![at] *= value.value; return
    default:
      acc.sum![at] += value.value
      acc.sumSquares![at] += value.value * value.value
  }
}

/** Fold one already-aggregated cell (`source`) into a total cell: the S2-AB leaf rollup. */
function foldCell(target: Accumulator, targetAt: number, source: Accumulator, sourceAt: number, fn: ValueFnCode): void {
  if (source.error?.[sourceAt] && !target.error![targetAt]) target.error![targetAt] = source.error[sourceAt]
  if (source.nonEmpty && target.nonEmpty) target.nonEmpty[targetAt] += source.nonEmpty[sourceAt]!
  if (source.numeric && target.numeric) target.numeric[targetAt] += source.numeric[sourceAt]!
  switch (fn) {
    case VALUE_FN.count: case VALUE_FN.countNums: return
    case VALUE_FN.sum: case VALUE_FN.average: target.sum![targetAt] += source.sum![sourceAt]!; return
    case VALUE_FN.max: if (source.max![sourceAt]! > target.max![targetAt]!) target.max![targetAt] = source.max![sourceAt]!; return
    case VALUE_FN.min: if (source.min![sourceAt]! < target.min![targetAt]!) target.min![targetAt] = source.min![sourceAt]!; return
    case VALUE_FN.product: target.product![targetAt] *= source.product![sourceAt]!; return
    default:
      target.sum![targetAt] += source.sum![sourceAt]!
      target.sumSquares![targetAt] += source.sumSquares![sourceAt]!
  }
}

/**
 * Totals from the leaf cells (S2-AB), as measured on the Task 15 oracle:
 * - Sum (and Count, Count Numbers, Max, Min, and the variance moments): each leaf row first sums its leaf cells across
 *   the columns in display order, then every row total sums those per-leaf-row cells down the rows in display order
 *   ("Σ leaf rows of Σ leaf cols"). This is bit-exact on every Sum cell of baseline, values-fn-sum and
 *   values-addition-order (e.g. baseline R6C8 East × grand and R24C8 grand × grand, which a flat row-major sum of
 *   leaf cells misses in the last bit), and it reproduces percent-of-total (R6C4) and index (R14C4) to 15 digits.
 * - Average: a total row × leaf column cell is the record-order sum of its records taken leaf row by leaf row
 *   (`seqRows`, filled by the caller); every total column then sums those cells across the leaf columns. This is the
 *   only tested order that reproduces every cell of values-fn-average (R24C8 grand × grand) and
 *   values-average-hidden-item (R20C6 grand × 2024, where per-leaf sums give …614 and Excel …616).
 * - Product keeps its direct record-order accumulator (values-fn-product R24C4); the caller fills it.
 */
function rollupTotals(gridAcc: AccumulatorBank, present: Uint8Array, rowLeaf: Uint8Array, colLeaf: Uint8Array,
  targets: Targets, gridCodes: ReadonlyArray<readonly number[]>, rowNamed: Uint8Array, colNamed: Uint8Array,
  rowCount: number, colCount: number, dataCount: number, rowData: Int32Array, colData: Int32Array): void {
  for (let data = 0; data < dataCount; data++) {
    const rows = targets.rows[data]!, cols = targets.cols[data]!
    const rowOk = (row: number) => rowData[row]! < 0 || rowData[row] === data
    const colOk = (col: number) => colData[col]! < 0 || colData[col] === data
    for (const code of gridCodes[data]!) {
      const fn = code as ValueFnCode
      if (fn === VALUE_FN.product) continue
      const acc = gridAcc[fn]!
      // Step 1: each row (leaf rows; for Average every row) folds its leaf-column cells into its total columns.
      for (let row = 0; row < rowCount; row++) {
        if (!rowOk(row) || (!rowLeaf[row] && fn !== VALUE_FN.average)) continue
        for (let col = 0; col < colCount; col++) {
          if (!colLeaf[col] || !colOk(col)) continue
          const source = cellIndex(row, col, data, colCount, dataCount)
          if (!present[source]) continue
          for (let c = cols.offsets[col]!; c < cols.offsets[col + 1]!; c++) {
            const target = cols.indices[c]!
            if (target === col || (rowNamed[row] && colNamed[target])) continue
            const at = cellIndex(row, target, data, colCount, dataCount)
            foldCell(acc, at, acc, source, fn); present[at] = 1
          }
        }
      }
      if (fn === VALUE_FN.average) continue
      // Step 2: each total row folds its leaf rows' cells (leaf and total columns) in display order.
      for (let row = 0; row < rowCount; row++) {
        if (!rowLeaf[row] || !rowOk(row)) continue
        for (let r = rows.offsets[row]!; r < rows.offsets[row + 1]!; r++) {
          const target = rows.indices[r]!
          if (target === row) continue
          for (let col = 0; col < colCount; col++) {
            if (!colOk(col) || (rowNamed[target] && colNamed[col])) continue
            const source = cellIndex(row, col, data, colCount, dataCount)
            if (!present[source]) continue
            const at = cellIndex(target, col, data, colCount, dataCount)
            foldCell(acc, at, acc, source, fn); present[at] = 1
          }
        }
      }
    }
  }
}

interface ReadValue { kind: 'missing' | 'other' | 'error' | 'date-unmeasured' | 'non-finite' | 'number'; value: number; text?: string }
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
      case 'error': into.kind = 'error'; into.text = item.value; break
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
    case VALUE_TAG.error: into.kind = 'error'; into.text = column.text[column.textIndex[row]!] ?? '#VALUE!'; break
    default: return false
  }
  return true
}

/** Sample and population variance from (n, Σx, Σx²). Every leaf cell of values-fn-var/-stdev matches
 * (Σx² − (Σx)²/n)/(n − 1), and every leaf cell of values-fn-varp/-stdevp matches (n·Σx² − (Σx)²)/n²; Welford and
 * two-pass forms miss leaf cells such as values-fn-stdevp R12C2 (Excel 27.65999999999757, two-pass …99968). */
function variance(acc: Accumulator, at: number, sample: boolean): number {
  const n = acc.numeric![at]!, s = acc.sum![at]!, q = acc.sumSquares![at]!
  return sample ? (q - s * s / n) / (n - 1) : (n * q - s * s) / (n * n)
}

function valueOf(acc: Accumulator, present: Uint8Array, at: number, fn: ValueFnCode): number | undefined {
  const count = acc.numeric?.[at] ?? 0
  if (!present[at]) return undefined
  switch (fn) {
    // Sum is blank for a group without numbers (Task 18 values-named-subtotals-intersecting R16C14, a record whose
    // Sales is blank); its error propagates.
    case VALUE_FN.sum: return acc.error?.[at] ? Number.NaN : count ? acc.sum![at]! : undefined
    // Count is blank for a group without values (values-named-subtotals-intersecting R30C2).
    case VALUE_FN.count: return acc.nonEmpty![at] ? acc.nonEmpty![at]! : undefined
    // Count Numbers is 0 over text values and blank without values (values-fn-countnums-mixed R6C4, R21C4).
    case VALUE_FN.countNums: return acc.nonEmpty![at] ? count : undefined
    case VALUE_FN.average: return count ? acc.sum![at]! / count : undefined
    case VALUE_FN.max: return count ? acc.max![at]! : undefined
    case VALUE_FN.min: return count ? acc.min![at]! : undefined
    case VALUE_FN.product: return count ? acc.product![at]! : undefined
    case VALUE_FN.stdDev: return count > 1 ? Math.sqrt(variance(acc, at, true)) : undefined
    case VALUE_FN.stdDevp: return count ? Math.sqrt(variance(acc, at, false)) : undefined
    case VALUE_FN.var: return count > 1 ? variance(acc, at, true) : undefined
    case VALUE_FN.varp: return count ? variance(acc, at, false) : undefined
  }
}

/** A sample statistic over a single number is #DIV/0! (values-fn-stdev R7C2 and values-fn-var R7C2, one record). */
function valueError(acc: Accumulator, present: Uint8Array, at: number, fn: ValueFnCode): string | undefined {
  if (!present[at] || (fn !== VALUE_FN.stdDev && fn !== VALUE_FN.var)) return undefined
  return acc.numeric![at] === 1 ? '#DIV/0!' : undefined
}

/** Aggregate a data field by caller-provided record keys using the engine's existing accumulator semantics. */
export function aggregateByKey(store: RecordStore, cache: PivotCacheDefinition, view: PivotViewState, dataIndex: number,
  keys: Uint32Array, keyCount: number, passes: Uint8Array, functionOverride?: PivotDataFunction): EngineResult<Array<number | undefined>> {
  if (!view.data[dataIndex] || keys.length !== store.count || passes.length !== store.count)
    return invalid('filter aggregate inputs do not match the pivot view')
  const data = view.data[dataIndex]!, sourceField = data.field
  const formula = cache.fields[sourceField]?.formula
  if (formula !== undefined) {
    // B2 rule (calc-value-sort): each item's value is the formula over that item's field sums.
    const compiled = compileCalculatedField(formula, cache, sourceField)
    if (!compiled.ok) return compiled
    const sums = new Map<number, Array<number | undefined>>()
    for (const field of compiled.value.fields) {
      const sum = aggregateByKey(store, cache, { ...view, data: view.data.map((d, k) => k === dataIndex ? { ...d, field, fn: 'sum', showAs: 'normal' } : d) }, dataIndex, keys, keyCount, passes, 'sum')
      if (!sum.ok) return sum
      sums.set(field, sum.value)
    }
    const out = new Array<number | undefined>(keyCount)
    for (let key = 0; key < keyCount; key++) {
      if (compiled.value.fields.every(field => sums.get(field)![key] === undefined)) continue
      const value = calculatedValue(compiled.value, field => sums.get(field)![key] ?? 0)
      if (value === 'unsupported' || value.kind !== 'number') return unmeasured(`a filter or sort by calculated data field ${dataIndex} with a non-number item value is not measured`)
      out[key] = value.value
    }
    return { ok: true, value: out }
  }
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
interface LineIndex { root: PrefixNode; grand: number[]; grandByData: Map<number, number[]>; nodes: Array<PrefixNode | undefined> }
/** Leaf lines: value-carrying data lines with no value-carrying line beneath them (their prefix-trie node has no
 * children). Values-field items sit in `items`, so lines of different data fields never share a node below it. */
function leafLineMask(lines: readonly PivotLine[], index: LineIndex): Uint8Array {
  const leaf = new Uint8Array(lines.length)
  for (let line = 0; line < lines.length; line++) {
    const node = index.nodes[line]
    if (node && lines[line]!.type === 'data' && node.children.size === 0) leaf[line] = 1
  }
  return leaf
}

const OVERFLOW = 'aggregate overflow or sample variance with fewer than two numbers needs measured error behavior'

/** A data field's values as numbers in record order, NaN where the record holds no number (blank, text, boolean). The
 * fast Sum path reads these instead of decoding every record per cell; an all-number inline column is used as is. */
function numericValues(column: RecordColumn | undefined, items: readonly import('../model').PivotCacheItem[] | undefined,
  count: number, field: number): EngineResult<Float64Array> {
  if (!column) return invalid(`a data field has no record column`)
  if (column.kind === 'inline') {
    const { tag, num } = column
    let plain = true
    for (let record = 0; record < count; record++) {
      const t = tag[record]
      if (t !== VALUE_TAG.number && t !== VALUE_TAG.date) { plain = false; break }
      if (!Number.isFinite(num[record]!)) return unmeasured(OVERFLOW)
    }
    if (plain) return { ok: true, value: num }
    const out = new Float64Array(count)
    for (let record = 0; record < count; record++) {
      switch (tag[record]) {
        case VALUE_TAG.number: case VALUE_TAG.date:
          if (!Number.isFinite(num[record]!)) return unmeasured(OVERFLOW)
          out[record] = num[record]!; break
        case VALUE_TAG.missing: case VALUE_TAG.string: case VALUE_TAG.boolean: out[record] = Number.NaN; break
        case VALUE_TAG.error: return unmeasured('error-valued records need a measured aggregation rule')
        default: return invalid(`record ${record} has an invalid value in field ${field}`)
      }
    }
    return { ok: true, value: out }
  }
  const lookup = new Float64Array(items?.length ?? 0)
  for (let k = 0; k < lookup.length; k++) {
    const item = items![k]!
    switch (item.type) {
      case 'number': if (!Number.isFinite(item.value)) return unmeasured(OVERFLOW); lookup[k] = item.value; break
      case 'date': return unmeasured('shared date values need the workbook date-system context')
      case 'error': return unmeasured('error-valued records need a measured aggregation rule')
      default: lookup[k] = Number.NaN
    }
  }
  const out = new Float64Array(count), index = column.index
  for (let record = 0; record < count; record++) {
    const k = index[record]!
    if (k >= lookup.length) return invalid(`record ${record} has an invalid value in field ${field}`)
    out[record] = lookup[k]!
  }
  return { ok: true, value: out }
}

/** Fast Sum path, record pass: each record's numbers into its leaf cell, in record order (S2-AB rule 1). Kept small
 * and separate so the JIT optimises the 1M-record loop. Returns the first record with an invalid line, or -1. */
function sumLeafCells(sums: Float64Array, filled: Uint8Array, values: readonly Float64Array[], rowKey: Uint32Array,
  colKey: Uint32Array, rowCount: number, colCount: number): number {
  const dataCount = values.length, count = rowKey.length
  if (dataCount === 2) {
    const first = values[0]!, second = values[1]!
    for (let record = 0; record < count; record++) {
      const row = rowKey[record]!, col = colKey[record]!
      if (row === NO_LINE || col === NO_LINE) continue
      if (row >= rowCount || col >= colCount) return record
      const base = (row * colCount + col) * 2
      const a = first[record]!, b = second[record]!
      if (a === a) { sums[base] += a; filled[base] = 1 }
      if (b === b) { sums[base + 1] += b; filled[base + 1] = 1 }
    }
    return -1
  }
  for (let record = 0; record < count; record++) {
    const row = rowKey[record]!, col = colKey[record]!
    if (row === NO_LINE || col === NO_LINE) continue
    if (row >= rowCount || col >= colCount) return record
    const base = (row * colCount + col) * dataCount
    for (let d = 0; d < dataCount; d++) {
      const value = values[d]![record]!
      if (value === value) { sums[base + d] += value; filled[base + d] = 1 }
    }
  }
  return -1
}

/**
 * Fast-path Sum rollup (S2-AB, see rollupTotals): first every leaf row folds its leaf cells into its total columns in
 * column display order, then every total row folds its leaf rows' cells in row display order. Work is
 * (leaf cells × column ancestors) + (leaf rows × row ancestors × columns), not records; adds are branch-free (an
 * empty cell holds 0) with presence OR-ed alongside.
 */
function rollupSumFromLeaves(sums: Float64Array, filled: Uint8Array, rowLeaf: Uint8Array, colLeaf: Uint8Array,
  rowTargets: LineTargets[], colTargets: LineTargets[], rowNamed: Uint8Array, colNamed: Uint8Array,
  rowCount: number, colCount: number, dataCount: number, sharedTargets: boolean): boolean {
  const groups = sharedTargets ? 1 : dataCount, width = colCount * dataCount
  const anyColNamed = colNamed.includes(1)
  for (let group = 0; group < groups; group++) {
    const rows = rowTargets[group]!, cols = colTargets[group]!
    const d0 = sharedTargets ? 0 : group, d1 = sharedTargets ? dataCount : group + 1
    // Leaf columns and their total columns (ancestors other than themselves), flattened.
    const leafCols: number[] = [], ancestorStart: number[] = [0], ancestors: number[] = []
    for (let col = 0; col < colCount; col++) {
      if (!colLeaf[col]) continue
      leafCols.push(col)
      for (let c = cols.offsets[col]!; c < cols.offsets[col + 1]!; c++) if (cols.indices[c] !== col) ancestors.push(cols.indices[c]!)
      ancestorStart.push(ancestors.length)
    }
    if (ancestors.length) for (let row = 0; row < rowCount; row++) {
      if (!rowLeaf[row]) continue
      const rowBase = row * width
      for (let k = 0; k < leafCols.length; k++) {
        const source = rowBase + leafCols[k]! * dataCount
        for (let a = ancestorStart[k]!; a < ancestorStart[k + 1]!; a++) {
          const target = rowBase + ancestors[a]! * dataCount
          for (let d = d0; d < d1; d++) { sums[target + d] += sums[source + d]!; filled[target + d] |= filled[source + d]! }
        }
      }
    }
    for (let row = 0; row < rowCount; row++) {
      if (!rowLeaf[row]) continue
      const source = row * width
      for (let r = rows.offsets[row]!; r < rows.offsets[row + 1]!; r++) {
        const targetRow = rows.indices[r]!
        if (targetRow === row) continue
        const target = targetRow * width
        if (sharedTargets && !(anyColNamed && rowNamed[targetRow])) {
          for (let k = 0; k < width; k++) { sums[target + k] += sums[source + k]!; filled[target + k] |= filled[source + k]! }
          continue
        }
        for (let col = 0; col < colCount; col++) {
          if (rowNamed[targetRow] && colNamed[col]) continue
          const offset = col * dataCount
          for (let d = d0; d < d1; d++) {
            sums[target + offset + d] += sums[source + offset + d]!; filled[target + offset + d] |= filled[source + offset + d]!
          }
        }
      }
    }
  }
  for (let at = 0; at < sums.length; at++) { const value = sums[at]!; if (value - value !== 0) return false }
  return true
}

/** Show-as denominators for a Sum slot, by the same S2-AB rule as its displayed totals: a row's total sums its leaf
 * rows' across-column totals, a column's total sums the leaf rows' cells in that column, and the grand total sums the
 * leaf rows' across-column totals. */
function sumTotalsFromLeaves(grid: Accumulator, gridPresent: Uint8Array, rowTotals: Accumulator, rowPresent: Uint8Array,
  colTotals: Accumulator, colPresent: Uint8Array, grandTotals: Accumulator, grandPresent: Uint8Array,
  rowLeaf: Uint8Array, colLeaf: Uint8Array, targets: Targets, rowCount: number, colCount: number, dataCount: number,
  data: number, rowData: Int32Array, colData: Int32Array): void {
  const rowOk = (row: number) => rowData[row]! < 0 || rowData[row] === data
  const colOk = (col: number) => colData[col]! < 0 || colData[col] === data
  const rows = targets.rows[data]!
  const reset = (acc: Accumulator, present: Uint8Array, at: number) => {
    acc.sum![at] = 0; acc.numeric![at] = 0; acc.error![at] = undefined; present[at] = 0
  }
  for (let row = 0; row < rowCount; row++) reset(rowTotals, rowPresent, row * dataCount + data)
  for (let col = 0; col < colCount; col++) reset(colTotals, colPresent, col * dataCount + data)
  reset(grandTotals, grandPresent, data)
  for (let row = 0; row < rowCount; row++) {
    if (!rowLeaf[row] || !rowOk(row)) continue
    const at = row * dataCount + data
    for (let col = 0; col < colCount; col++) {
      if (!colOk(col)) continue
      const source = cellIndex(row, col, data, colCount, dataCount)
      if (!gridPresent[source]) continue
      if (colLeaf[col]) { foldCell(rowTotals, at, grid, source, VALUE_FN.sum); rowPresent[at] = 1 }
      foldCell(colTotals, col * dataCount + data, grid, source, VALUE_FN.sum); colPresent[col * dataCount + data] = 1
    }
    if (!rowPresent[at]) continue
    for (let r = rows.offsets[row]!; r < rows.offsets[row + 1]!; r++) {
      const target = rows.indices[r]!
      if (target === row) continue
      foldCell(rowTotals, target * dataCount + data, rowTotals, at, VALUE_FN.sum); rowPresent[target * dataCount + data] = 1
    }
    foldCell(grandTotals, data, rowTotals, at, VALUE_FN.sum); grandPresent[data] = 1
  }
}

function makePrefixIndex(lines: readonly PivotLine[], view: PivotViewState, axis: 'rows' | 'cols'): LineIndex {
  const root: PrefixNode = { children: new Map(), lines: [] }, grand: number[] = [], grandByData = new Map<number, number[]>()
  const nodes = new Array<PrefixNode | undefined>(lines.length)
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
    nodes[index] = node
  })
  return { root, grand, grandByData, nodes }
}

function readTargets(store: RecordStore, view: PivotViewState, axes: PivotAxes, rowLineIndex: LineIndex, colLineIndex: LineIndex,
  validateRecords = true): EngineResult<Targets> {
  if (validateRecords) for (let record = 0; record < store.count; record++) {
    const row = axes.rowKey[record], col = axes.colKey[record]
    if (row === undefined || col === undefined) return invalid(`record ${record} has an invalid leaf line`)
    // Excel totals are based on records visible in the current pivot view.
    // A page selection, hidden item or filter marks excluded source rows NO_LINE.
    if (row === NO_LINE || col === NO_LINE) continue
    if (row >= axes.rows.length || col >= axes.cols.length) return invalid(`record ${record} has an invalid leaf line`)
  }
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
    if (row === NO_LINE || col === NO_LINE) continue
    if (rowList.offsets[row] === rowList.offsets[row + 1] || colList.offsets[col] === colList.offsets[col + 1])
        return invalid(`record ${record} has no value-carrying row or column line`)
    }
  } else for (let record = 0; record < store.count; record++) for (let d = 0; d < view.data.length; d++) {
    const row = axes.rowKey[record]!, col = axes.colKey[record]!
    if (row === NO_LINE || col === NO_LINE) continue
    if (rows[d]!.offsets[row] === rows[d]!.offsets[row + 1] || cols[d]!.offsets[col] === cols[d]!.offsets[col + 1])
      return invalid(`record ${record} has no value-carrying row or column line`)
  }
  return { ok: true, value: { rows, cols } }
}

function cellIndex(row: number, col: number, data: number, colCount: number, dataCount: number): number {
  return (row * colCount + col) * dataCount + data
}

/** One-pass aggregation of source records, followed by the show-as calculation; calculated data fields are the
 * formula over each cell's field sums (calculated.ts). */
export function computeValues(store: RecordStore, cache: PivotCacheDefinition, def: PivotTableDefinition,
  view: PivotViewState, axes: PivotAxes): EngineResult<PivotValueGrid> {
  return view.data.some(data => cache.fields[data.field]?.formula !== undefined)
    ? computeCalculatedValues(store, cache, def, view, axes)
    : computeSourceValues(store, cache, def, view, axes)
}

const sumOf = (data: DataFieldState, field: number): DataFieldState => ({ ...data, field, fn: 'sum', showAs: 'normal' })

/**
 * Calculated data fields (spec §5.6). Each slot's cells are its formula over the sums of the fields it names, in the
 * same cell: one all-sum pass per named field (same data slots, so Values-axis lines keep their slot), plus one pass
 * for the other data fields. Fails closed (S2-M) where Excel's rule is unmeasured: show-as or a function other than
 * sum on the calculated field, a named subtotal line, a text-valued named field, a value-carrying cell without
 * records, and an error or text result (division by zero's `#DIV/0!` or `errorCaption`).
 */
function computeCalculatedValues(store: RecordStore, cache: PivotCacheDefinition, def: PivotTableDefinition,
  view: PivotViewState, axes: PivotAxes): EngineResult<PivotValueGrid> {
  const compiled: Array<CalculatedField | undefined> = []
  for (let d = 0; d < view.data.length; d++) {
    const data = view.data[d]!, formula = cache.fields[data.field]?.formula
    if (formula === undefined) continue
    // B2 rule: the function is ignored (calc-function-count).
    // Percent of total is measured (calc-show-as): each unrounded calculated cell over the calculated grand total.
    if (data.showAs !== 'normal' && data.showAs !== 'percentOfTotal') return unmeasured(`show-as on calculated data field ${d} is not measured`)
    const field = compileCalculatedField(formula, cache, data.field)
    if (!field.ok) return field
    compiled[d] = field.value
  }
  if (axes.cols.some(line => isNamedSubtotal(line.type)))
    return unmeasured('a named subtotal of a calculated data field is not measured')
  // B2 rule (calc-named-subtotals): on a named row subtotal the calculated cell is blank, except at the column grand
  // total, where it is the formula over the item's sums (whatever the line's function).
  const namedRow = axes.rows.map(line => isNamedSubtotal(line.type))
  const sumAxes: PivotAxes = { ...axes, rows: axes.rows.map(line => isNamedSubtotal(line.type) ? { ...line, type: 'default' } : line) }
  const referenced = [...new Set(compiled.flatMap(field => field?.fields ?? []))]
  for (const field of referenced) {
    const column = store.columns[field], items = cache.fields[field]?.sharedItems?.items
    if (!column) return invalid(`calculated field source ${field} has no record column`)
    const text = column.kind === 'shared'
      ? items?.some(item => item.type === 'string' || item.type === 'boolean')
      : column.tag.some(tag => tag === VALUE_TAG.string || tag === VALUE_TAG.boolean)
    void text // B2 rule: text is left out of the sum (calc-text-field).
  }
  const gridSize = axes.rows.length * axes.cols.length * view.data.length
  if (!Number.isSafeInteger(gridSize) || (referenced.length + 1) * gridSize * OUTPUT_BYTES_PER_SLOT * 2 > MAX_ACCUMULATOR_BYTES)
    return fail('pivot-values-limit', 'the calculated value grids exceed their bounded accumulator budget')

  const anyRecordField = cache.fields.findIndex((field, index) => field.databaseField && store.columns[index] !== undefined)
  const base = computeSourceValues(store, cache, def, { ...view, data: view.data.map((data, d) => compiled[d] ? sumOf(data, compiled[d]!.fields[0] ?? anyRecordField) : data) }, axes)
  if (!base.ok) return base
  const sums = new Map<number, PivotValueGrid>()
  for (const field of referenced) {
    const grid = computeSourceValues(store, cache, def, { ...view, data: view.data.map(data => sumOf(data, field)) }, sumAxes)
    if (!grid.ok) return grid
    sums.set(field, grid.value)
  }

  const rowCount = axes.rows.length, colCount = axes.cols.length, dataCount = view.data.length
  const values = new Float64Array(gridSize), present = new Uint8Array(gridSize), errors = new Map<number, string>()
  const onRows = view.rows.includes(PIVOT_VALUES_FIELD), onCols = view.cols.includes(PIVOT_VALUES_FIELD)
  const colCarries = axes.cols.map(line => lineCarriesValues(view, 'cols', line))
  for (let row = 0; row < rowCount; row++) {
    if (!lineCarriesValues(view, 'rows', axes.rows[row]!)) continue
    for (let col = 0; col < colCount; col++) {
      if (!colCarries[col]) continue
      const d = onRows ? axes.rows[row]!.data : onCols ? axes.cols[col]!.data : 0
      const field = compiled[d]
      if (!field) continue
      if (namedRow[row] && axes.cols[col]!.type !== 'grand') continue
      // B2 rule: a cell without records is the formula over zero sums (calc-cell-without-records).
      const value = calculatedValue(field, source => {
        const sum = sums.get(source)!.get(row, col, d)
        return sum.kind === 'number' ? sum.value : 0
      })
      if (value === 'unsupported') return fail('pivot-calculated-field', `calculated data field ${d} uses a construct the formula engine lacks`)
      const at = cellIndex(row, col, d, colCount, dataCount)
      // B2 rule: an error shows as the error value with showError off (calc-division-by-zero).
      if (value.kind === 'error' && !def.showError) { errors.set(at, value.value); present[at] = 1; continue }
      if (value.kind !== 'number') return unmeasured(`calculated data field ${d} gives ${value.kind} ${value.kind === 'empty' ? '' : value.value}; its display is not measured`)
      values[at] = value.value; present[at] = 1
    }
  }
  // calc-show-as: percent of total divides every calculated cell (unrounded) by the slot's grand × grand cell; all 76
  // cells match, R6C7 included. Without both grand lines the denominator has no measured source.
  for (let d = 0; d < dataCount; d++) {
    if (!compiled[d] || view.data[d]!.showAs !== 'percentOfTotal') continue
    // An axis without a field other than Values has a single line per slot, which is that axis's total.
    const totalLine = (lines: readonly PivotLine[], fields: readonly number[], on: boolean) =>
      fields.some(field => field !== PIVOT_VALUES_FIELD)
        ? lines.findIndex(line => line.type === 'grand' && (!on || line.data === d))
        : lines.findIndex(line => line.type === 'data' && (!on || line.data === d))
    const grandRow = totalLine(axes.rows, view.rows, onRows), grandCol = totalLine(axes.cols, view.cols, onCols)
    if (grandRow < 0 || grandCol < 0) return unmeasured(`percent of total on calculated data field ${d} without grand totals is not measured`)
    const grandAt = cellIndex(grandRow, grandCol, d, colCount, dataCount)
    if (errors.has(grandAt) || !present[grandAt]) return unmeasured(`percent of total on calculated data field ${d} with a non-number grand total is not measured`)
    const grand = values[grandAt]!
    for (let row = 0; row < rowCount; row++) for (let col = 0; col < colCount; col++) {
      const at = cellIndex(row, col, d, colCount, dataCount)
      if (!present[at] || errors.has(at)) continue
      if (grand === 0) errors.set(at, '#DIV/0!')
      else values[at] = values[at]! / grand
    }
  }
  return {
    ok: true,
    value: {
      get(row, col, data): PivotCellValue {
        if (!compiled[data]) return base.value.get(row, col, data)
        if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0 || row >= rowCount || col >= colCount) return { kind: 'empty' }
        const at = cellIndex(row, col, data, colCount, dataCount)
        const error = errors.get(at)
        if (error !== undefined) return { kind: 'error', value: error }
        return present[at] ? { kind: 'number', value: values[at]! } : { kind: 'empty' }
      },
    },
  }
}

function computeSourceValues(store: RecordStore, cache: PivotCacheDefinition, _def: PivotTableDefinition,
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

  const sourceColumns = view.data.map(data => store.columns[data.field])
  const sourceItems = view.data.map(data => cache.fields[data.field]?.sharedItems?.items)
  const hasErrorValues = sourceItems.some(items => items?.some(item => item.type === 'error'))
    || sourceColumns.some(column => column?.kind === 'inline' && column.tag.includes(VALUE_TAG.error))
  const fastSumCandidate = !view.rows.includes(PIVOT_VALUES_FIELD) && !view.cols.includes(PIVOT_VALUES_FIELD)
    && view.data.every(data => data.showAs === 'normal')
    && functions.every(fn => fn === VALUE_FN.sum)
    && gridBankCodes.length === 1 && gridBankCodes[0] === VALUE_FN.sum
    && rowBankCodes.length === 1 && rowBankCodes[0] === VALUE_FN.sum
    && colBankCodes.length === 1 && colBankCodes[0] === VALUE_FN.sum
    && grandBankCodes.length === 1 && grandBankCodes[0] === VALUE_FN.sum
    && !hasErrorValues
  const rowLineIndex = makePrefixIndex(axes.rows, view, 'rows'), colLineIndex = makePrefixIndex(axes.cols, view, 'cols')
  const recordTargets = readTargets(store, view, axes, rowLineIndex, colLineIndex, !fastSumCandidate)
  if (!recordTargets.ok) return recordTargets

  const rowLeaf = leafLineMask(axes.rows, rowLineIndex), colLeaf = leafLineMask(axes.cols, colLineIndex)


  // S2-AA: variance totals stay refused. Their leaf cells match the textbook forms (see `variance`), but no tested
  // accumulation of (n, Σx, Σx²) reproduces every total cell of values-fn-stdev/-stdevp/-var/-varp: per-leaf sums
  // rolled up (S2-AB), records in display order, per-row or per-column partial sums, Welford with a parallel merge,
  // two-pass over the records and fused multiply-add forms each leave 1–9 total cells off in the 15th digit
  // (e.g. values-fn-var R24C8 grand × grand: Excel 6768554.289461315, closest 6768554.289461324).
  // A total line is any value-carrying line that is not a leaf: grand, default and named subtotals, and parent data
  // lines that carry their subtotal (on top or collapsed); only leaf data lines and blank lines are not.
  const hasTotals = (lines: readonly PivotLine[], leaf: Uint8Array, axis: 'rows' | 'cols') =>
    lines.some((line, index) => !leaf[index] && lineCarriesValues(view, axis, line))
  const varianceFns = new Set<ValueFnCode>([VALUE_FN.stdDev, VALUE_FN.stdDevp, VALUE_FN.var, VALUE_FN.varp])
  if (gridBankCodes.some(fn => varianceFns.has(fn as ValueFnCode)) && (hasTotals(axes.rows, rowLeaf, 'rows') || hasTotals(axes.cols, colLeaf, 'cols')))
    return unmeasured('variance subtotal rounding differs from Excel: its total cells differ in the 15th digit under every tested accumulation of the leaf moments')

  // A collapsed parent line is a leaf here: its records accumulate straight into its cell, in record order. Excel does
  // the same for Sum (layout-collapsed-item and -above-values are bit-exact, East × 2023 included, where summing the
  // hidden Product groups would differ in the last bit). Count, Count Numbers, Max and Min do not depend on the
  // accumulation order. S2-AA: Average, Product and the variance functions on a collapsed line have no oracle case.
  const collapsedLeaf = (lines: readonly PivotLine[], leaf: Uint8Array, fields: readonly number[]) => lines.some((line, index) =>
    leaf[index] && line.items.some((item, level) => {
      const field = fields[level]
      return field !== undefined && field !== PIVOT_VALUES_FIELD && item >= 0 && view.fields[field]?.collapsed[item] === true
        && fields.slice(level + 1).some(after => after !== PIVOT_VALUES_FIELD)
    }))
  const orderSensitive = new Set<ValueFnCode>([VALUE_FN.average, VALUE_FN.product, ...varianceFns])
  if (gridBankCodes.some(fn => orderSensitive.has(fn as ValueFnCode))
    && (collapsedLeaf(axes.rows, rowLeaf, view.rows) || collapsedLeaf(axes.cols, colLeaf, view.cols)))
    return unmeasured('average, product and variance on a collapsed item are not measured: no Excel oracle case covers them')

  const sumOnly = !hasErrorValues && functions.every(fn => fn === VALUE_FN.sum)
    && gridBankCodes.length === 1 && gridBankCodes[0] === VALUE_FN.sum
    && rowBankCodes.length === 1 && rowBankCodes[0] === VALUE_FN.sum
    && colBankCodes.length === 1 && colBankCodes[0] === VALUE_FN.sum
    && grandBankCodes.length === 1 && grandBankCodes[0] === VALUE_FN.sum
  const normalOnly = view.data.every(data => data.showAs === 'normal')
  if (sumOnly && normalOnly) {
    // Sums start at 0 with a separate presence byte per cell, so the rollup adds without branching on emptiness (an
    // empty cell adds 0, which leaves any total unchanged).
    const sums = new Float64Array(gridSize), filled = new Uint8Array(gridSize)
    const scratch: ReadValue = { kind: 'missing', value: 0 }
    const commonTargets = !view.rows.includes(PIVOT_VALUES_FIELD) && !view.cols.includes(PIVOT_VALUES_FIELD)
    if (commonTargets) {
      const values: Float64Array[] = []
      for (let d = 0; d < dataCount; d++) {
        const column = numericValues(sourceColumns[d], sourceItems[d], store.count, view.data[d]!.field)
        if (!column.ok) return column
        values.push(column.value)
      }
      const bad = sumLeafCells(sums, filled, values, axes.rowKey, axes.colKey, rowCount, colCount)
      if (bad >= 0) return invalid(`record ${bad} has an invalid leaf line`)
    } else {
      for (let record = 0; record < store.count; record++) {
        const rowLine = axes.rowKey[record]!, colLine = axes.colKey[record]!
        if (rowLine === NO_LINE || colLine === NO_LINE) continue
        for (let d = 0; d < dataCount; d++) {
          const field = view.data[d]!.field
          if (!readValue(sourceColumns[d], sourceItems[d], record, scratch)) return invalid(`record ${record} has an invalid value in field ${field}`)
          if (scratch.kind === 'error') return unmeasured('error-valued records need a measured aggregation rule')
          if (scratch.kind === 'date-unmeasured') return unmeasured('shared date values need the workbook date-system context')
          if (scratch.kind === 'non-finite') return unmeasured(OVERFLOW)
          if (scratch.kind !== 'number') continue
          const value = scratch.value
          const rowList = recordTargets.value.rows[d]!, rowStart = rowList.offsets[rowLine]!, rowEnd = rowList.offsets[rowLine + 1]!
          const colList = recordTargets.value.cols[d]!, colStart = colList.offsets[colLine]!, colEnd = colList.offsets[colLine + 1]!
          for (let r = rowStart; r < rowEnd; r++) {
            const row = rowList.indices[r]!
            if (!rowLeaf[row]) continue
            for (let c = colStart; c < colEnd; c++) {
              const col = colList.indices[c]!, at = (row * colCount + col) * dataCount + d
              if (!colLeaf[col]) continue
              sums[at] += value; filled[at] = 1
              if (!Number.isFinite(sums[at]!)) return unmeasured(OVERFLOW)
            }
          }
        }
      }
    }
    if (!rollupSumFromLeaves(sums, filled, rowLeaf, colLeaf, recordTargets.value.rows, recordTargets.value.cols,
      rowNamed, colNamed, rowCount, colCount, dataCount, commonTargets))
      return unmeasured(OVERFLOW)
    const grid: PivotValueGrid = {
      get(row, col, data): PivotCellValue {
        if (!Number.isInteger(row) || !Number.isInteger(col) || !Number.isInteger(data) || row < 0 || col < 0 || data < 0
          || row >= rowCount || col >= colCount || data >= dataCount) return { kind: 'empty' }
        const at = (row * colCount + col) * dataCount + data
        return filled[at] ? { kind: 'number', value: sums[at]! } : { kind: 'empty' }
      },
    }
    return { ok: true, value: grid }
  }

  const gridPresent = new Uint8Array(gridSize), rowPresent = new Uint8Array(rowCount * dataCount)
  const colPresent = new Uint8Array(colCount * dataCount), grandPresent = new Uint8Array(dataCount)
  const gridAcc = accumulatorBank(gridSize, gridBankCodes), rowTotals = accumulatorBank(rowCount * dataCount, rowBankCodes)
  const colTotals = accumulatorBank(colCount * dataCount, colBankCodes), grandTotals = accumulatorBank(dataCount, grandBankCodes)
  const valueScratch: ReadValue = { kind: 'missing', value: 0 }
  const targets = recordTargets.value
  const rowsOnValues = view.rows.includes(PIVOT_VALUES_FIELD), colsOnValues = view.cols.includes(PIVOT_VALUES_FIELD)
  const rowData = Int32Array.from(axes.rows, line => rowsOnValues ? line.data : -1)
  const colData = Int32Array.from(axes.cols, line => colsOnValues ? line.data : -1)
  // Phase 1, records in record order: every function of the slot into the leaf cells (the S2-AB rollup's inputs),
  // Product's direct record accumulator into its total cells, and the per-line record totals that show-as reads for
  // functions other than Sum.
  for (let record = 0; record < store.count; record++) {
    if (axes.rowKey[record] === NO_LINE || axes.colKey[record] === NO_LINE) continue
    for (let d = 0; d < dataCount; d++) {
      const field = view.data[d]!.field
      if (!readValue(sourceColumns[d], sourceItems[d], record, valueScratch)) return invalid(`record ${record} has an invalid value in field ${field}`)
      // See the Task 15 error-value oracle: only Sum propagates the error.
      if (valueScratch.kind === 'date-unmeasured') return unmeasured('shared date values need the workbook date-system context')
      if (valueScratch.kind === 'non-finite') return unmeasured('aggregate overflow or sample variance with fewer than two numbers needs measured error behavior')
      const fn = functions[d]!, codes = gridCodes[d]!
      const rowLine = axes.rowKey[record]!, colLine = axes.colKey[record]!
      const rowList = targets.rows[d]!, colList = targets.cols[d]!
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
        // Excel leaves intersections blank when row and column named subtotal
        // functions conflict (Task 15 values-named-subtotals-intersecting).
        if (isRowNamed && isColNamed) continue
        const cellFn = (isRowNamed ? rowFunctions[row * dataCount + d]! : isColNamed ? colFunctions[col * dataCount + d]! : fn) as ValueFnCode
        const at = cellIndex(row, col, d, colCount, dataCount)
        if (rowLeaf[row] && colLeaf[col]) for (const code of codes) add(gridAcc[code]!, gridPresent, at, valueScratch, code as ValueFnCode)
        else if (cellFn === VALUE_FN.product) add(gridAcc[cellFn]!, gridPresent, at, valueScratch, cellFn)
      }
    }
  }
  // Phase 2, Average: a total row × leaf column cell sums its records leaf row by leaf row, each leaf row's records
  // in record order (see rollupTotals). A stable counting sort of the records by leaf row gives that order.
  const averageData = gridCodes.map(codes => codes.includes(VALUE_FN.average))
  if (averageData.some(Boolean) && rowLeaf.some((leaf, row) => !leaf && lineCarriesValues(view, 'rows', axes.rows[row]!))) {
    const starts = new Uint32Array(rowCount + 1)
    for (let record = 0; record < store.count; record++) {
      const row = axes.rowKey[record]!
      if (row !== NO_LINE && axes.colKey[record] !== NO_LINE) starts[row + 1]++
    }
    for (let row = 0; row < rowCount; row++) starts[row + 1] += starts[row]!
    const order = new Uint32Array(starts[rowCount]!), next = starts.slice(0, rowCount)
    for (let record = 0; record < store.count; record++) {
      const row = axes.rowKey[record]!
      if (row !== NO_LINE && axes.colKey[record] !== NO_LINE) order[next[row]!++] = record
    }
    const average = gridAcc[VALUE_FN.average]!
    for (const record of order) for (let d = 0; d < dataCount; d++) {
      if (!averageData[d]) continue
      readValue(sourceColumns[d], sourceItems[d], record, valueScratch)
      const rowLine = axes.rowKey[record]!, colLine = axes.colKey[record]!
      const rowList = targets.rows[d]!, colList = targets.cols[d]!
      for (let c = colList.offsets[colLine]!; c < colList.offsets[colLine + 1]!; c++) {
        const col = colList.indices[c]!
        if (!colLeaf[col]) continue
        for (let r = rowList.offsets[rowLine]!; r < rowList.offsets[rowLine + 1]!; r++) {
          const row = rowList.indices[r]!
          if (rowLeaf[row]) continue
          add(average, gridPresent, cellIndex(row, col, d, colCount, dataCount), valueScratch, VALUE_FN.average)
        }
      }
    }
  }
  rollupTotals(gridAcc, gridPresent, rowLeaf, colLeaf, targets, gridCodes, rowNamed, colNamed,
    rowCount, colCount, dataCount, rowData, colData)
  for (let d = 0; d < dataCount; d++) if (functions[d] === VALUE_FN.sum)
    sumTotalsFromLeaves(gridAcc[VALUE_FN.sum]!, gridPresent, rowTotals[VALUE_FN.sum]!, rowPresent, colTotals[VALUE_FN.sum]!,
      colPresent, grandTotals[VALUE_FN.sum]!, grandPresent, rowLeaf, colLeaf, targets, rowCount, colCount, dataCount,
      d, rowData, colData)

  const raw = new Float64Array(gridSize), present = new Uint8Array(gridSize), rawErrors: Array<string | undefined> = new Array(gridSize)
  for (let row = 0; row < rowCount; row++) for (let col = 0; col < colCount; col++) for (let d = 0; d < dataCount; d++) {
    const at = cellIndex(row, col, d, colCount, dataCount)
    const rowType = axes.rows[row]!.type, colType = axes.cols[col]!.type
    const rowNamed = isNamedSubtotal(rowType), colNamed = isNamedSubtotal(colType)
    const rowFn = subtotalFunction(rowType, view.data[d]!.fn), colFn = subtotalFunction(colType, view.data[d]!.fn)
    if (rowNamed && colNamed) continue
    const fn = rowNamed ? rowFn : colNamed ? colFn : view.data[d]!.fn
    const fnCode = valueFnCode(fn)
    const result = valueOf(gridAcc[fnCode]!, gridPresent, at, fnCode)
    const error = gridAcc[fnCode]?.error?.[at] ?? valueError(gridAcc[fnCode]!, gridPresent, at, fnCode)
    // Excel keeps a Count detail cell numeric for an error-valued source, but
    // propagates that source error when it rolls the Count cell into a subtotal.
    const countedErrorDetail = fnCode === VALUE_FN.count && rowLeaf[row] && colLeaf[col]
    if (error && !countedErrorDetail) { rawErrors[at] = error; present[at] = 1; continue }
    if (result === undefined) continue
    if (!Number.isFinite(result)) return unmeasured('aggregate overflow or sample variance with fewer than two numbers needs measured error behavior')
    raw[at] = result; present[at] = 1
  }

  const output = raw.slice(), outputPresent = present.slice(), outputErrors: Array<string | undefined> = rawErrors.slice()
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
    const data = view.data[d]!, mode = data.showAs as PivotShowDataAs | string
    if (mode === 'normal') continue
    const hasLeafRecord = !!present[at]
    if (rawErrors[at]) continue
    const rowLine = axes.rows[row]!, colLine = axes.cols[col]!
    const difference = mode === 'difference' || mode === 'percent' || mode === 'percentDiff'
    const running = mode === 'runTotal' || mode === 'percentOfRunningTotal'
    const emptyItemIsRelevant = (difference && (rowLine.type === 'data' || colLine.type === 'data'))
      || (running && (rowLine.type === 'data' || colLine.type === 'data'))
    if (!hasLeafRecord && !emptyItemIsRelevant
      && (!lineCarriesValues(view, 'rows', rowLine) || !lineCarriesValues(view, 'cols', colLine))) continue
    let denominator: number | undefined, numerator = hasLeafRecord ? raw[at]! : 0, result: number
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
      let greater = new Set<number>()
      for (let sibling = 0; sibling < lines.length; sibling++) {
        if (sibling === current) continue
        const line = lines[sibling]!
        if (line.type !== source.type || line.items.length !== source.items.length || !lineCarriesValues(view, axis, line)) continue
        if (source.items.some((item, k) => k !== level && line.items[k] !== item)) continue
        const siblingAt = rowAxis ? cellIndex(sibling, col, d, colCount, dataCount) : cellIndex(row, sibling, d, colCount, dataCount)
        if (!present[siblingAt]) continue
        const siblingValue = raw[siblingAt]!
        if (mode === 'rankAscending' ? siblingValue < numerator : siblingValue > numerator) greater.add(siblingValue)
      }
      const rank = 1 + greater.size
      output[at] = rank
      continue
    }
    else if (mode === 'index') {
      // Index is ((value × grand) / row total) / column total, left to right: values-showas-index R15C4 and
      // values-showas-index-zero-total R19C7 differ in the 15th digit for value × grand / (row × column). A zero row
      // or column total is #DIV/0! (index-zero-total R22C3, Central's Units).
      const row0 = rowTotal ?? 0, col0 = colTotal ?? 0
      if (row0 === 0 || col0 === 0) { outputErrors[at] = '#DIV/0!'; outputPresent[at] = 1; continue }
      // A cell without records is 0 (values-showas-index R22C2, Central 2022).
      result = numerator * (grandTotal ?? 0) / row0 / col0
      output[at] = result; outputPresent[at] = 1
      continue
    } else if (mode === 'difference' || mode === 'percentDiff' || mode === 'percent') {
      // Measured on values-showas-difference-previous/-next/-item, -percent-of-item and -percent-diff-zero-base:
      // - previous/next without a sibling is blank (difference-previous column 2022);
      // - an explicit base item whose line does not exist under the same parent is #N/A (difference-item and
      //   percent-of-item R23, North has no Elderberries);
      // - the base item's own cells are blank for difference and % difference (difference-item rows 10–13,
      //   percent-diff-zero-base column C3) and 1 for % of (percent-of-item rows 10–13);
      // - a blank base cell counts as 0 for difference (difference-previous R22C4 Central 2023 = 410.25) and leaves
      //   % difference blank (percent-diff-zero-base R22C5); a blank current cell counts as 0 for difference
      //   (difference-next R22C2 = −410.25) and is #NULL! for % of (percent-of-item R22C2);
      // - a zero base is #DIV/0! (percent-diff-zero-base R17C5, South/Cherries 2022 Units = 0).
      const rowAxis = rowFields.includes(data.baseField), colAxis = colFields.includes(data.baseField)
      if (rowAxis === colAxis) return unmeasured('show-as base field must identify exactly one row or column axis')
      if (data.baseItem === PIVOT_BASE_ITEM_DEFAULT) return unmeasured('default base-item behavior is not measured yet')
      const axis = rowAxis ? 'rows' : 'cols', current = rowAxis ? row : col
      const relative = data.baseItem === PIVOT_BASE_ITEM_PREVIOUS || data.baseItem === PIVOT_BASE_ITEM_NEXT
      const lines = rowAxis ? axes.rows : axes.cols, level = (rowAxis ? rowFields : colFields).indexOf(data.baseField)
      if (level < 0 || level >= lines[current]!.items.length) { outputPresent[at] = 0; continue }
      const base = findBaseLine(axis, current, data.baseField, data.baseItem)
      if (base === undefined) {
        if (relative) { outputPresent[at] = 0; continue }
        if (mode === 'percentDiff') return unmeasured('% difference from a base item without a line is not measured')
        outputErrors[at] = '#N/A'; outputPresent[at] = 1; continue
      }
      const baseAt = rowAxis ? cellIndex(base, col, d, colCount, dataCount) : cellIndex(row, base, d, colCount, dataCount)
      const basePresent = !!present[baseAt], baseValue = basePresent ? raw[baseAt]! : 0
      if (base === current) {
        if (mode === 'percent' && hasLeafRecord) { output[at] = 1; outputPresent[at] = 1 } else outputPresent[at] = 0
        continue
      }
      if (!hasLeafRecord && !basePresent) { outputPresent[at] = 0; continue }
      if (mode === 'difference') { output[at] = numerator - baseValue; outputPresent[at] = 1; continue }
      if (mode === 'percent' && !hasLeafRecord) { outputErrors[at] = '#NULL!'; outputPresent[at] = 1; continue }
      if (mode === 'percentDiff' && !basePresent) { outputPresent[at] = 0; continue }
      if (!basePresent) return unmeasured('% of a base item with a blank base cell is not measured')
      if (!hasLeafRecord) return unmeasured('% difference of a blank cell from its base item is not measured')
      if (baseValue === 0) { outputErrors[at] = '#DIV/0!'; outputPresent[at] = 1; continue }
      output[at] = mode === 'percent' ? numerator / baseValue : (numerator - baseValue) / baseValue
      outputPresent[at] = 1
      continue
    } else if (mode === 'percentOfParent' || mode === 'percentOfParentRow' || mode === 'percentOfParentCol') {
      const axis = mode === 'percentOfParentCol' ? 'cols' : 'rows'
      let baseField: number | undefined
      if (mode === 'percentOfParent' && data.baseField !== -1) baseField = data.baseField
      // % of parent total with a base field is blank on the grand line (values-showas-percent-of-parent-region
      // R60C2); without a base field the grand line is its own parent (100%).
      if (baseField !== undefined && (axis === 'rows' ? rowLine : colLine).type === 'grand') { outputPresent[at] = 0; continue }
      const parent = parentLine(axis, axis === 'rows' ? row : col, baseField)
      if (parent === undefined) { outputPresent[at] = 0; continue }
      const parentAt = axis === 'rows' ? cellIndex(parent, col, d, colCount, dataCount) : cellIndex(row, parent, d, colCount, dataCount)
      if (!present[parentAt]) { outputPresent[at] = 0; continue }
      denominator = raw[parentAt]!
    } else if (mode === 'runTotal' || mode === 'percentOfRunningTotal') {
      const rowAxis = rowFields.includes(data.baseField), colAxis = colFields.includes(data.baseField)
      if (rowAxis === colAxis) return unmeasured('running-total base field must identify exactly one row or column axis')
      const axis = rowAxis ? 'rows' : 'cols', lines = rowAxis ? axes.rows : axes.cols
      const current = rowAxis ? row : col, currentLine = lines[current]!, level = (rowAxis ? rowFields : colFields).indexOf(data.baseField)
      if (level < 0 || level >= currentLine.items.length) {
        outputPresent[at] = 0
        continue
      }
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
      if (mode === 'runTotal') { output[at] = cumulative; outputPresent[at] = 1; continue }
      numerator = cumulative
      denominator = total
      if (denominator !== 0) outputPresent[at] = 1
    } else return fail('pivot-engine-pending', `show-as mode ${mode} is not implemented`)

    if (denominator === undefined) { outputPresent[at] = 0; continue }
    if (denominator === 0) { outputErrors[at] = '#DIV/0!'; outputPresent[at] = 1; continue }
    output[at] = numerator / denominator
    outputPresent[at] = 1
  }

  for (let at = 0; at < output.length; at++) if (outputPresent[at] && !Number.isFinite(output[at]!))
    return unmeasured('show-as result overflow is not measured yet')

  const grid: PivotValueGrid = {
    get(row, col, data): PivotCellValue {
      if (!Number.isInteger(row) || !Number.isInteger(col) || !Number.isInteger(data) || row < 0 || col < 0 || data < 0
        || row >= rowCount || col >= colCount || data >= dataCount) return { kind: 'empty' }
      const at = cellIndex(row, col, data, colCount, dataCount)
      if (outputErrors[at]) return { kind: 'error', value: outputErrors[at]! }
      return outputPresent[at] ? { kind: 'number', value: output[at]! } : { kind: 'empty' }
    },
  }
  return { ok: true, value: grid }
}
