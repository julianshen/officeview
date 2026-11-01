import { evaluateFormulaInternal, makeCellKey, formulaError, isFormulaError, isEvaluationError, isMatrixValue, publicFormulaValue, round15, scalarProjection, topLeftOfMatrix } from './evaluator'
import { parseFormula } from './parser'
import { translateSharedFormula } from './shared'
import type { AstNode, CellAddress, ElementValue, EvaluationContext, EvaluationValue, ModeFrameStore, ModeFrameStores, RefArea, ResolvedSemantics, SheetId } from './types'
import { createReferenceServices, gridAreas } from './refs'
import { SpillEngine, SpillSuspension, spillCellKey, type SpillBlockReason, type SpillIssue } from './spills'
import { resolveUnicodeSettings, unicodeVersionOf } from './unicode'
import { decideCellCalculation, isWhatIfDataTable, type CalcMode, type CalcPolicy } from './calc-policy'
import { canonicalOrder, findCyclicCells, ITERATIVE_PASS_HARD_CAP, runBoundedPasses, transitiveDependents } from './iterative'
import type { XlsxCell, XlsxDocument, XlsxRow, XlsxSheet } from '../types'
import { savedSpillFollowers } from './saved-spills'
import { workbookValuePolicy } from './value-policy'

/** Per-document value policy (spec §4.4): cached documents never reach the evaluator. */
export { markSavedValuesOnly, workbookValuePolicy, type WorkbookValuePolicy } from './value-policy'

/** Monotonic reference generation: each recalculation invalidates old cursors. */
let workbookReferenceGeneration = 0

export interface WorkbookEvaluationOptions {
  /** If true, recalculates all formula cells even if a cached <v> is present. */
  fullCalcOnLoad?: boolean
  /** Force recalculation of every formula cell. */
  forceRecalc?: boolean
  /** Evaluation intent. 'file-load' (default) preserves the established cache
   * policy; 'explicit-recalc' recomputes supported formulas (sampling the
   * injected clock once per generation) while unsupported constructs keep
   * their valid caches. */
  intent?: 'file-load' | 'explicit-recalc'
  /** Workbook text-length semantics (1 = legacy code units, 2 = current code
   * points). Legacy Stage A explicit flag; folded into the Unicode precedence. */
  textLengthVersion?: 1 | 2
  /** Explicit Unicode compatibility version override (wins over verified file
   * metadata). 1 = legacy UTF-16 code units, 2 = current code points. */
  unicodeVersion?: 1 | 2
  /**
   * Explicit array provenance mode. 'from-file' (default) is metadata-driven
   * and preserves legacy cached behavior; 'legacy' treats only fixed f.ref
   * arrays; 'dynamic' establishes dynamic spill owners for array-capable roots
   * (explicit-mode provenance — never guessed from file metadata).
   */
  arrayMode?: 'from-file' | 'legacy' | 'dynamic'
  /** Application resource policy for spill publication (never a native claim). */
  maxArrayCells?: number
  maxArraySide?: number
  /** Cumulative cap on cells materialized by array evaluation (results and
   * array-mode references), charged before each materialization and never
   * reset within one evaluation. Unset = no cumulative limit. */
  maxMaterializedCells?: number
  /** Explicit locale override (wins over parsed file metadata). */
  locale?: string
  /** Explicit IANA timeZone override. */
  timeZone?: string
  /** Injected clock sample (ms). Sampled lazily once per generation if absent. */
  now?: () => number
  /** Explicit date-system override (wins over the parsed workbookPr date1904). */
  date1904?: boolean
  /** Lazy pivot adapter; absent on load so GETPIVOTDATA retains its shipped value. */
  getPivotData?: (args: readonly AstNode[], evaluate: (node: AstNode) => EvaluationValue, currentSheet: string, markStale?: (feature: string) => void) => EvaluationValue
  /** Optional view-only cell values, used by a resident pivot commit without mutating saved sheet cells. */
  readCellValue?: (sheetName: string, col: number, row: number) => EvaluationValue | undefined
  /** Sparse cells owned by the currently committed view overlay, including cells
   * that were absent in the saved worksheet. Used by range aggregates. */
  overlayEntries?: readonly { sheetName: string; col: number; row: number; value: EvaluationValue }[]
}

function capturedWorkbookZone(): string {
  try {
    const zone = Intl.DateTimeFormat().resolvedOptions().timeZone
    return typeof zone === 'string' && zone !== '' ? zone : 'UTC'
  } catch {
    return 'UTC'
  }
}

/**
 * Resolve one semantics object per recalculation generation.
 * Precedence: explicit option > verified file metadata > entry default.
 * The clock is sampled LAZILY (first NOW/TODAY read) so a generation that
 * never runs a volatile function does not sample it; it is then reused for the
 * whole generation.
 */
function resolveWorkbookSemantics(
  doc: XlsxDocument,
  options?: WorkbookEvaluationOptions,
  onVolatile?: () => void,
): ResolvedSemantics {
  const base = doc.semantics
  const dateSystem: '1900' | '1904' = options?.date1904 !== undefined
    ? (options.date1904 ? '1904' : '1900')
    : (base?.dateSystem ?? '1900')
  // Unicode precedence: explicit option > verified metadata > entry default
  // (parsed workbook absent verified metadata = 1; standalone = 2). The legacy
  // textLengthVersion flag is an explicit option too.
  const entryDefault: 'workbook-default' | 'standalone-default' = base ? 'workbook-default' : 'standalone-default'
  const unicode = resolveUnicodeSettings(options?.unicodeVersion ?? options?.textLengthVersion, base?.unicode, entryDefault)
  let sampled = false
  let value = 0
  return {
    dateSystem,
    unicode,
    locale: options?.locale ?? base?.locale ?? 'en-US',
    timeZone: options?.timeZone ?? base?.timeZone ?? capturedWorkbookZone(),
    get epochNowMs(): number {
      if (!sampled) {
        sampled = true
        value = options?.now ? options.now() : Date.now()
      }
      onVolatile?.()
      return value
    },
  }
}

/** Array-capable root syntax/functions that can establish a dynamic owner. */
const ARRAY_CAPABLE_FUNCTIONS = new Set([
  'SEQUENCE', 'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'TRANSPOSE', 'IF', 'IFS', 'SWITCH',
  'IFERROR', 'IFNA', 'INDEX', 'CHOOSE', 'ROW', 'COLUMN', 'MMULT', 'MINVERSE', 'MUNIT',
])

function isArrayCapableRoot(node: AstNode): boolean {
  switch (node.type) {
    case 'arrayConst':
    case 'spill':
    case 'range':
    case 'wholeCol':
    case 'wholeRow':
    case 'union':
    case 'intersect':
    case 'ref3d':
    case 'table':
    case 'name':
      return true
    case 'call':
      return ARRAY_CAPABLE_FUNCTIONS.has(node.name.toUpperCase())
    case 'binary':
      return isArrayCapableRoot(node.left) || isArrayCapableRoot(node.right)
    case 'unary':
      return isArrayCapableRoot(node.expr)
    default:
      return false
  }
}

/** Persistent per-document spill ownership (survives recalculation calls). */
type CellContent = Omit<XlsxCell, 'ref' | 'col' | 'row'>
interface GeneratedCellRecord {
  key: string
  sheet: XlsxSheet
  cell: XlsxCell
  ownerKey: string
  /** Exact content we wrote, for detecting later user edits. */
  written: CellContent
  /** Original content when the cell pre-existed (restore); undefined = we created it. */
  original?: CellContent
}
interface CreatedRowRecord { sheet: XlsxSheet; index: number; ownerKey: string }
interface DocSpillState {
  generated: GeneratedCellRecord[]
  createdRows: CreatedRowRecord[]
  valueDeps: Map<string, Set<string>>
  /** Actual reads including readers that consumed a memoized formula value. */
  formulaDeps: Map<string, Set<string>>
  /** Cached reverse of formulaDeps: target cell key -> formulas that read it. */
  reverseFormulaDeps: Map<string, Set<string>>
  retainedUnsupportedKeys?: Set<string>
  runtimeIndex?: {
    sheetsByName: Map<string, XlsxSheet>; cellMap: Map<string, XlsxCell>; cellSheets: Map<string, string>
    sheetMaxCols: Map<string, number>; sheetMaxRows: Map<string, number>; inputKeys: Set<string>; formulaKeys: Set<string>; mergeKeys: Set<string>
    identitySheets: { sheetId: SheetId; name: string; workbookIndex: number }[]
    /** The document's cell-model revision (`noteWorkbookCellsChanged`) this index was built or last validated at. */
    revision: number
  }
}
const docSpillStates = new WeakMap<XlsxDocument, DocSpillState>()
const liveFormulaSnapshots = new WeakMap<XlsxDocument, Map<string, { cell: XlsxCell; content: CellContent }>>()
/** The saved state of one worksheet position a dependent recompute's spill work wrote or cleared (Ruling S3-AL): the
 * row and cell objects there before it first touched the position (absent: there was none), the cell's content, and
 * the position's runtime-index entries. */
interface SpillImage {
  sheet: XlsxSheet; row: number; col: number; rowObj?: XlsxRow; cell?: XlsxCell; content?: CellContent
  indexed?: XlsxCell; indexedSheet?: string; input: boolean
}
/** Per document, from its first dependent recompute until the restore: every position spill work touched, with its
 * saved state, and the saved generation's spill records (the cells it owned and the rows it created). */
interface SpillImages { positions: Map<string, SpillImage>; generated: GeneratedCellRecord[]; createdRows: CreatedRowRecord[] }
const liveSpillImages = new WeakMap<XlsxDocument, SpillImages>()
/** Set while a dependent recompute evaluates: told about a position before spill work writes or clears it. */
let spillImageRecorder: ((sheet: XlsxSheet, col: number, row: number) => void) | undefined

function reverseFormulaEdges(edges: Map<string, Set<string>>): Map<string, Set<string>> {
  const reverse = new Map<string, Set<string>>()
  for (const [reader, targets] of edges) for (const target of targets) {
    let readers = reverse.get(target)
    if (!readers) reverse.set(target, readers = new Set())
    readers.add(reader)
  }
  return reverse
}
function removeReverseFormulaEdge(reverse: Map<string, Set<string>>, target: string, reader: string): void {
  const readers = reverse.get(target)
  if (!readers) return
  readers.delete(reader)
  if (!readers.size) reverse.delete(target)
}
function addReverseFormulaEdge(reverse: Map<string, Set<string>>, target: string, reader: string): void {
  let readers = reverse.get(target)
  if (!readers) reverse.set(target, readers = new Set())
  readers.add(reader)
}

/** Engine-published spill footprint of one owner: the bounding box of its anchor and live generated cells. */
export interface EngineSpillFootprint { anchor: { row: number; col: number }; r0: number; r1: number; c0: number; c1: number }
/** Structural spill bridge (viewer S3.3): per-sheet footprints of the owners
 * that published generated cells in the last evaluation, as rectangles only.
 * Undefined when this document was never evaluated (no ownership is known). */
export function engineSpillFootprints(doc: XlsxDocument): ReadonlyMap<XlsxSheet, readonly EngineSpillFootprint[]> | undefined {
  const state = docSpillStates.get(doc)
  if (!state) return undefined
  const owners = new Map<string, { sheet: XlsxSheet; footprint: EngineSpillFootprint }>()
  for (const g of state.generated) {
    let entry = owners.get(g.ownerKey)
    if (!entry) {
      // Owner keys are spillCellKey(anchor): `${sheetId}!${col}:${row}`; spills stay on their anchor's sheet.
      const m = /!(\d+):(\d+)$/.exec(g.ownerKey)
      if (!m) continue
      const col = Number(m[1]), row = Number(m[2])
      entry = { sheet: g.sheet, footprint: { anchor: { row, col }, r0: row, r1: row, c0: col, c1: col } }
      owners.set(g.ownerKey, entry)
    }
    const f = entry.footprint
    f.r0 = Math.min(f.r0, g.cell.row); f.r1 = Math.max(f.r1, g.cell.row)
    f.c0 = Math.min(f.c0, g.cell.col); f.c1 = Math.max(f.c1, g.cell.col)
  }
  const out = new Map<XlsxSheet, EngineSpillFootprint[]>()
  for (const { sheet, footprint } of owners.values()) {
    const list = out.get(sheet)
    if (list) list.push(footprint)
    else out.set(sheet, [footprint])
  }
  return out
}

/** Snapshot/restore the full cell content (everything except its coordinates). */
function snapshotCell(cell: XlsxCell): CellContent {
  const { ref: _ref, col: _col, row: _row, ...rest } = cell
  void _ref
  void _col
  void _row
  return rest
}

function restoreCell(cell: XlsxCell, snap: CellContent): void {
  for (const key of Object.keys(cell)) {
    if (key !== 'ref' && key !== 'col' && key !== 'row') delete (cell as unknown as Record<string, unknown>)[key]
  }
  Object.assign(cell, snap)
}

/** True when the cell still holds EXACTLY the content we wrote (no user edit). */
function cellContentMatches(cell: XlsxCell, written: CellContent): boolean {
  const current = snapshotCell(cell) as Record<string, unknown>
  const expected = written as Record<string, unknown>
  const keys = new Set([...Object.keys(current), ...Object.keys(expected)])
  for (const key of keys) {
    if (current[key] !== expected[key]) return false
  }
  return true
}

function columnName(col: number): string {
  let name = ''
  let n = col
  while (n >= 0) {
    name = String.fromCharCode(65 + (n % 26)) + name
    n = Math.floor(n / 26) - 1
  }
  return name
}

function ensureCell(sheet: XlsxSheet, col: number, row: number): XlsxCell {
  let rowObj = sheet.rows.find((r) => r.index === row)
  if (!rowObj) {
    rowObj = { index: row, cells: [] }
    const at = sheet.rows.findIndex((r) => r.index > row)
    if (at < 0) sheet.rows.push(rowObj)
    else sheet.rows.splice(at, 0, rowObj)
  }
  let cell = rowObj.cells.find((c) => c.col === col)
  if (!cell) {
    cell = { ref: `${columnName(col)}${row + 1}`, col, row, value: null, styleIndex: 0 }
    const at = rowObj.cells.findIndex((c) => c.col > col)
    if (at < 0) rowObj.cells.push(cell)
    else rowObj.cells.splice(at, 0, cell)
  }
  return cell
}

function removeCell(sheet: XlsxSheet, cell: XlsxCell): boolean {
  const rowObj = sheet.rows.find((r) => r.index === cell.row)
  if (!rowObj) return false
  const at = rowObj.cells.indexOf(cell)
  if (at < 0) return false
  rowObj.cells.splice(at, 1)
  // NOTE: the row is never removed here. Rows we did not create carry original
  // formatting/metadata (heightPt/customHeight) and must survive a shrink.
  return true
}

/** Remove an empty row ONLY when this recalculation created it. */
function removeCreatedRow(sheet: XlsxSheet, index: number): void {
  const rowObj = sheet.rows.find((r) => r.index === index)
  if (!rowObj || rowObj.cells.length > 0) return
  const at = sheet.rows.indexOf(rowObj)
  if (at >= 0) sheet.rows.splice(at, 1)
}

/** Transitive readers of a set of changed cells (value-dependency closure). */
function affectedReaders(changed: ReadonlySet<string>, deps: Map<string, Set<string>>): Set<string> {
  const reverse = new Map<string, Set<string>>()
  for (const [reader, targets] of deps) {
    for (const target of targets) {
      const set = reverse.get(target)
      if (set) set.add(reader)
      else reverse.set(target, new Set([reader]))
    }
  }
  const out = new Set<string>()
  const queue = [...changed]
  while (queue.length > 0) {
    const key = queue.pop()!
    for (const reader of reverse.get(key) ?? []) {
      if (!out.has(reader)) {
        out.add(reader)
        queue.push(reader)
      }
    }
  }
  return out
}

export interface ChangedSheetRect { sheetId: number; rect: { minRow: number; minCol: number; maxRow: number; maxCol: number } }

type FormulaReference = { sheet: string; minRow: number; minCol: number; maxRow: number; maxCol: number }
type IndexedFormula = { key: string; sheet: string; row: number; col: number; refs: FormulaReference[]; pivotAnchorRefs: FormulaReference[] }
type IndexedReference = { ref: FormulaReference; key: string }
type ReferenceIntervalNode = {
  center: number
  byStart: IndexedReference[]
  byEnd: IndexedReference[]
  left?: ReferenceIntervalNode
  right?: ReferenceIntervalNode
}
type FormulaReferenceIndex = { formulas: IndexedFormula[]; byKey: Map<string, IndexedFormula>; bySheet: Map<string, ReferenceIntervalNode>; volatileReaders: Set<string>; revision: number }
const formulaReferenceIndexes = new WeakMap<XlsxDocument, FormulaReferenceIndex>()

/** The cell model's mutation revision per document (T31 r4 M-1). The formula engine's own writes (spill followers,
 * restored snapshots) never change formula text and keep the runtime cell index in step themselves; anything else
 * that adds, removes or replaces cells, or edits formula text in place, after the workbook was evaluated announces it
 * with `noteWorkbookCellsChanged`. A pivot commit then validates its caches by this number instead of walking every
 * cell of the workbook. */
const cellModelRevisions = new WeakMap<XlsxDocument, number>()
const cellModelRevision = (doc: XlsxDocument): number => cellModelRevisions.get(doc) ?? 0

/** Announces a direct edit of the document's cells (added, removed or replaced cells, or formula text edited in
 * place) made after its formulas were evaluated: the next dependent recompute rebuilds its indexes from the cells.
 * A full `evaluateWorkbookFormulas` resynchronises too. */
export function noteWorkbookCellsChanged(doc: XlsxDocument): void {
  cellModelRevisions.set(doc, cellModelRevision(doc) + 1)
}

function buildReferenceTree(items: IndexedReference[]): ReferenceIntervalNode | undefined {
  if (!items.length) return undefined
  const points = items.map(({ ref }) => Math.floor((ref.minRow + ref.maxRow) / 2)).sort((a, b) => a - b)
  const center = points[Math.floor(points.length / 2)]!
  const left: IndexedReference[] = [], right: IndexedReference[] = [], crossing: IndexedReference[] = []
  for (const item of items) {
    if (item.ref.maxRow < center) left.push(item)
    else if (item.ref.minRow > center) right.push(item)
    else crossing.push(item)
  }
  // A pathological distribution of point intervals at the same coordinate
  // still makes progress: those intervals belong to this node.
  return {
    center,
    byStart: crossing.slice().sort((a, b) => a.ref.minRow - b.ref.minRow),
    byEnd: crossing.slice().sort((a, b) => b.ref.maxRow - a.ref.maxRow),
    ...(left.length ? { left: buildReferenceTree(left) } : {}),
    ...(right.length ? { right: buildReferenceTree(right) } : {}),
  }
}

function queryReferenceTree(node: ReferenceIntervalNode | undefined, rect: { minRow: number; maxRow: number; minCol: number; maxCol: number }, out: Set<string>): void {
  if (!node) return
  const add = (item: IndexedReference) => {
    if (item.ref.minCol <= rect.maxCol && item.ref.maxCol >= rect.minCol) out.add(item.key)
  }
  if (rect.maxRow < node.center) {
    for (const item of node.byStart) { if (item.ref.minRow > rect.maxRow) break; add(item) }
    queryReferenceTree(node.left, rect, out)
  } else if (rect.minRow > node.center) {
    for (const item of node.byEnd) { if (item.ref.maxRow < rect.minRow) break; add(item) }
    queryReferenceTree(node.right, rect, out)
  } else {
    for (const item of node.byStart) add(item)
    queryReferenceTree(node.left, rect, out)
    queryReferenceTree(node.right, rect, out)
  }
}

function queryReferences(index: FormulaReferenceIndex, sheet: string, rect: { minRow: number; maxRow: number; minCol: number; maxCol: number }): Set<string> {
  const out = new Set<string>()
  queryReferenceTree(index.bySheet.get(sheet.toLowerCase()), rect, out)
  return out
}

function referenceIndexFor(doc: XlsxDocument): FormulaReferenceIndex {
  const revision = cellModelRevision(doc)
  const cached = formulaReferenceIndexes.get(doc)
  if (cached?.revision === revision) return cached
  const formulas: IndexedFormula[] = []
  const byKey = new Map<string, IndexedFormula>()
  const volatileReaders = new Set<string>()
  for (const sheet of doc.sheets) resolveSheetSharedFormulas(sheet, doc)
  const nameStack = new Set<string>()
  const collect = (node: AstNode, currentSheet: string, out: FormulaReference[], pivotOnly?: FormulaReference[], volatile?: { found: boolean }): void => {
    const add = (sheetName: string, minCol: number, minRow: number, maxCol: number, maxRow: number, target = out) => {
      target.push({ sheet: sheetName.toLowerCase(), minRow: Math.min(minRow, maxRow), minCol: Math.min(minCol, maxCol), maxRow: Math.max(minRow, maxRow), maxCol: Math.max(minCol, maxCol) })
    }
    switch (node.type) {
      case 'cell': add(node.ref.sheet ?? currentSheet, node.ref.col, node.ref.row, node.ref.col, node.ref.row); return
      case 'spill': add(node.anchor.sheet ?? currentSheet, node.anchor.col, node.anchor.row, node.anchor.col, node.anchor.row); return
      case 'range': add(node.ref.sheet ?? node.ref.from.sheet ?? currentSheet, node.ref.from.col, node.ref.from.row, node.ref.to.col, node.ref.to.row); return
      case 'wholeCol': add(node.ref.sheet ?? currentSheet, node.ref.from.index, 0, node.ref.to.index, 1_048_575); return
      case 'wholeRow': add(node.ref.sheet ?? currentSheet, 0, node.ref.from.index, 16_383, node.ref.to.index); return
      case 'ref3d': {
        const from = doc.sheets.findIndex(sheet => sheet.name.toLowerCase() === node.sheets.fromSheet.toLowerCase())
        const to = doc.sheets.findIndex(sheet => sheet.name.toLowerCase() === node.sheets.toSheet.toLowerCase())
        const names = from >= 0 && to >= 0 ? doc.sheets.slice(Math.min(from, to), Math.max(from, to) + 1).map(sheet => sheet.name) : []
        for (const name of names) collect(node.target as AstNode, name, out, pivotOnly, volatile)
        return
      }
      case 'name': {
        const localIndex = node.ref.sheet ? doc.sheets.findIndex(s => s.name.toLowerCase() === node.ref.sheet!.toLowerCase()) : doc.sheets.findIndex(s => s.name.toLowerCase() === currentSheet.toLowerCase())
        const match = (doc.definedNames ?? []).find(name => name.name.toLowerCase() === node.ref.name.toLowerCase() && name.localSheetIndex === localIndex)
          ?? (doc.definedNames ?? []).find(name => name.name.toLowerCase() === node.ref.name.toLowerCase() && name.localSheetIndex === undefined)
        if (match) {
          const id = `${match.localSheetIndex ?? 'w'}:${match.name.toLowerCase()}`
          if (!nameStack.has(id)) {
            nameStack.add(id)
            try { collect(parseFormula(match.source.replace(/^=/, '')), node.ref.sheet ?? currentSheet, out, pivotOnly, volatile) } catch { /* unsupported name remains governed by evaluator */ }
            nameStack.delete(id)
          }
        }
        return
      }
      case 'call': {
        if (node.name.toUpperCase() === 'INDIRECT' || node.name.toUpperCase() === 'OFFSET') if (volatile) volatile.found = true
        if (node.name.toUpperCase() === 'GETPIVOTDATA') {
          const anchor = node.args[1]
          if (anchor?.type === 'cell') add(anchor.ref.sheet ?? currentSheet, anchor.ref.col, anchor.ref.row, anchor.ref.col, anchor.ref.row, pivotOnly ?? out)
        }
        for (const arg of node.args) collect(arg, currentSheet, out, pivotOnly, volatile)
        return
      }
      case 'binary': case 'intersect': collect(node.left, currentSheet, out, pivotOnly, volatile); collect(node.right, currentSheet, out, pivotOnly, volatile); return
      case 'unary': case 'implicitIntersect': collect(node.expr, currentSheet, out, pivotOnly, volatile); return
      case 'union': for (const ref of node.refs) collect(ref, currentSheet, out, pivotOnly, volatile); return
      default: return
    }
  }
  for (const sheet of doc.sheets) for (const row of sheet.rows) for (const cell of row.cells) {
    if (!cell.formula) continue
    const key = makeCellKey(sheet.name, cell.col, cell.row)
    const entry: IndexedFormula = { key, sheet: sheet.name.toLowerCase(), row: cell.row, col: cell.col, refs: [], pivotAnchorRefs: [] }
    const volatile = { found: false }
    try { collect(parseFormula(cell.formula), sheet.name, entry.refs, entry.pivotAnchorRefs, volatile) } catch { /* malformed formulas are handled by the evaluator */ }
    if (volatile.found) volatileReaders.add(key)
    formulas.push(entry); byKey.set(key, entry)
  }
  const refsBySheet = new Map<string, IndexedReference[]>()
  for (const formula of formulas) for (const ref of [...formula.refs, ...formula.pivotAnchorRefs]) {
    const list = refsBySheet.get(ref.sheet)
    const item = { ref, key: formula.key }
    if (list) list.push(item); else refsBySheet.set(ref.sheet, [item])
  }
  const bySheet = new Map<string, ReferenceIntervalNode>()
  for (const [sheet, refs] of refsBySheet) {
    const tree = buildReferenceTree(refs)
    if (tree) bySheet.set(sheet, tree)
  }
  const result = { formulas, byKey, bySheet, volatileReaders, revision }
  formulaReferenceIndexes.set(doc, result)
  return result
}

function refreshRuntimeCellIndex(doc: XlsxDocument, runtime: NonNullable<DocSpillState['runtimeIndex']>, revision: number): void {
  const sheetsByName = new Map<string, XlsxSheet>(), cellMap = new Map<string, XlsxCell>(), cellSheets = new Map<string, string>()
  const sheetMaxCols = new Map<string, number>(), sheetMaxRows = new Map<string, number>()
  const inputKeys = new Set<string>(), formulaKeys = new Set<string>(), mergeKeys = new Set<string>()
  for (const sheet of doc.sheets) {
    resolveSheetSharedFormulas(sheet, doc)
    sheetsByName.set(sheet.name.toLowerCase(), sheet)
    let maxCol = -1, maxRow = -1
    for (const row of sheet.rows) {
      maxRow = Math.max(maxRow, row.index)
      for (const cell of row.cells) {
        maxCol = Math.max(maxCol, cell.col); maxRow = Math.max(maxRow, cell.row)
        const key = makeCellKey(sheet.name, cell.col, cell.row)
        cellMap.set(key, cell); cellSheets.set(key, sheet.name)
        if (cell.formula !== undefined || cell.sharedFormula !== undefined) formulaKeys.add(key)
        else if (cell.value !== null && cell.value !== undefined) inputKeys.add(key)
      }
    }
    sheetMaxCols.set(sheet.name.toLowerCase(), Math.max(0, maxCol))
    sheetMaxRows.set(sheet.name.toLowerCase(), Math.max(0, maxRow))
    for (const merge of sheet.mergeRanges ?? []) for (let row = merge.minRow; row <= merge.maxRow; row++) for (let col = merge.minCol; col <= merge.maxCol; col++) mergeKeys.add(makeCellKey(sheet.name, col, row))
  }
  Object.assign(runtime, { sheetsByName, cellMap, cellSheets, sheetMaxCols, sheetMaxRows, inputKeys, formulaKeys, mergeKeys,
    revision })
}

/** Conservative admission bound for subset recalculation. A range aggregate
 * currently scans all stored workbook cells in refs.ts; include that cost too. */
export const MAX_DEPENDENT_CELL_VISITS = 2_000_000
/** Passes a recompute makes after its spills grew or shrank, for the readers of the cells they moved over (a spill
 * whose readers spill again counts once per level); past it those readers are marked stale. */
const MAX_SPILL_FOLLOW_PASSES = 8

/** The used extents and stored-cell count the estimate charges against. With the evaluated workbook's runtime index
 * (the normal case) they are exactly what the engine reads: whole-row/column references clip to its per-sheet maxima,
 * and range aggregates enumerate its cell map, which follows spill writes and removals. No walk of the cells. */
function workExtents(doc: XlsxDocument, runtime: DocSpillState['runtimeIndex']): { used: Map<string, { rows: number; cols: number }>; storedCells: number } {
  const used = new Map<string, { rows: number; cols: number }>()
  if (runtime) {
    for (const [name, maxRow] of runtime.sheetMaxRows) used.set(name, { rows: maxRow + 1, cols: (runtime.sheetMaxCols.get(name) ?? 0) + 1 })
    return { used, storedCells: runtime.cellMap.size }
  }
  let storedCells = 0
  for (const sheet of doc.sheets) {
    let rows = 0, cols = 0
    for (const row of sheet.rows) {
      rows = Math.max(rows, row.index + 1)
      for (const cell of row.cells) { rows = Math.max(rows, cell.row + 1); cols = Math.max(cols, cell.col + 1); storedCells++ }
    }
    used.set(sheet.name.toLowerCase(), { rows, cols })
  }
  return { used, storedCells }
}

function dependentWorkEstimate(doc: XlsxDocument, index: FormulaReferenceIndex, keys: ReadonlySet<string>, runtime: DocSpillState['runtimeIndex']): number {
  const { used, storedCells } = workExtents(doc, runtime)
  let visits = 0
  for (const key of keys) {
    const formula = index.byKey.get(key)
    if (!formula) continue
    for (const ref of formula.refs) {
      const extent = used.get(ref.sheet)
      const rows = Math.max(0, Math.min(ref.maxRow + 1, extent?.rows ?? 0) - Math.max(0, ref.minRow))
      const cols = Math.max(0, Math.min(ref.maxCol + 1, extent?.cols ?? 0) - Math.max(0, ref.minCol))
      const span = rows * cols
      // Cell references are O(1); aggregate ranges enumerate the stored map.
      const cost = span <= 1 ? span : Math.max(span, storedCells)
      visits += cost
      if (visits > MAX_DEPENDENT_CELL_VISITS) return visits
    }
  }
  return visits
}

/** Paged sheets cannot build the resident formula graph; record one sheet-level
 * stale notice when any formula cell is outside the pivot rectangles. */
export function markPagedDependentsStale(
  prior: ReadonlySet<number>, sheetId: number, sheet: XlsxSheet,
  pivotRects: readonly { minRow: number; minCol: number; maxRow: number; maxCol: number }[],
  workbookHasFormulaOutsidePivots = false,
): ReadonlySet<number> {
  if (prior.has(sheetId)) return prior
  const hasDependent = workbookHasFormulaOutsidePivots || sheet.rows.some(row => row.cells.some(cell => Boolean(cell.formula)
    && !pivotRects.some(rect => cell.row >= rect.minRow && cell.row <= rect.maxRow && cell.col >= rect.minCol && cell.col <= rect.maxCol)))
  return hasDependent ? new Set([...prior, sheetId]) : prior
}

function chartRanges(model: import('../chartModel').XlsxChartModel): Array<{ sheet?: string; r0: number; c0: number; r1: number; c1: number }> {
  const ranges: Array<{ sheet?: string; r0: number; c0: number; r1: number; c1: number }> = []
  if (model.categoryRef) ranges.push(model.categoryRef)
  if (model.categoryPending) ranges.push(model.categoryPending)
  for (const group of model.groups) for (const series of group.series) {
    for (const range of Object.values(series.refs ?? {})) if (range) ranges.push(range)
    for (const range of Object.values(series.pending ?? {})) if (range) ranges.push(range)
  }
  return ranges
}

function chartsUsing(doc: XlsxDocument, formulaKeys: ReadonlySet<string>): number[] {
  const positions = new Set<string>()
  for (const key of formulaKeys) positions.add(key)
  const changed = new Set<number>()
  let chartIndex = 0
  const visit = (sheet: XlsxSheet, nodes: readonly import('../types').XlsxDrawing[] | undefined): void => {
    for (const node of nodes ?? []) {
      const chart = node.content?.kind === 'xlsxChart' ? node.content.chart : undefined
      if (chart) {
        const index = chartIndex++
        if (chartRanges(chart).some(range => {
          const name = (range.sheet ?? sheet.name).toLowerCase()
          for (const key of positions) {
            const bang = key.lastIndexOf('!'), coords = /^(\d+):(\d+)$/.exec(key.slice(bang + 1))
            if (bang < 0 || key.slice(0, bang) !== name || !coords) continue
            const col = Number(coords[1]), row = Number(coords[2])
            if (row >= range.r0 && row <= range.r1 && col >= range.c0 && col <= range.c1) return true
          }
          return false
        })) changed.add(index)
      }
      if (node.children?.length) visit(sheet, node.children)
    }
  }
  for (const sheet of doc.sheets) visit(sheet, sheet.drawings)
  return [...changed]
}

/** Re-evaluate just the transitive formula readers of committed pivot cells.
 * Dependency edges are those observed during the last workbook evaluation;
 * GETPIVOTDATA's anchor reference is also matched syntactically so its shipped
 * cache can survive the earlier unknown-function pass. */
export function recomputeDependents(
  doc: XlsxDocument,
  changed: readonly ChangedSheetRect[],
  options?: WorkbookEvaluationOptions,
): { recomputed: number; charts: readonly number[]; sheets: readonly number[]; stale: readonly string[]; /** Every reader this call covered: recomputed, or (over budget) left as it was. */ readers: readonly string[]; /** Over the work budget (S3-AE): nothing ran, and the stale readers keep whatever value they showed (maybe an earlier commit's live value), not necessarily the saved one (T31 r4 M-4). */ overBudget?: true } {
  if (workbookValuePolicy(doc) === 'cached') throw new RangeError('This workbook shows saved values only; its formulas are never evaluated')
  const prior = docSpillStates.get(doc)
  const index = referenceIndexFor(doc)
  // A host that edited the cells after the full pass announced it (noteWorkbookCellsChanged): rebuild only the
  // runtime cell maps; the syntactic index was rebuilt above and cached values remain intact. Unchanged, neither
  // index walks the workbook (M-1).
  if (prior?.runtimeIndex && prior.runtimeIndex.revision !== index.revision) refreshRuntimeCellIndex(doc, prior.runtimeIndex, index.revision)
  const formulaKeys = new Set<string>(), queue: string[] = []
  const enqueue = (key: string) => { if (index.byKey.has(key) && !formulaKeys.has(key)) { formulaKeys.add(key); queue.push(key) } }
  for (const entry of changed) {
    const sheet = doc.sheets[entry.sheetId]
    if (sheet) for (const key of queryReferences(index, sheet.name, entry.rect)) enqueue(key)
  }
  // S3-Z: runtime-computed references cannot be indexed syntactically. These
  // volatile formulas are a small conservative reader set and follow commits.
  for (const key of index.volatileReaders) enqueue(key)

  // Build the observed-edge reverse map once, then traverse it alongside the
  // cached spatial index. This is O(edges + visited readers), not readers × formulas.
  const dynamicReverse = prior?.reverseFormulaDeps ?? reverseFormulaEdges(prior?.formulaDeps ?? new Map())
  // Spill records key their owner by sheet identity (`spillCellKey`); the reader queue keys cells by lower-cased sheet
  // name (`makeCellKey`). Followers are looked up by the owner's cell key, so a recomputed spill's cells, and so their
  // readers, are always found (PR #57 review: with numeric sheet ids the two never met).
  const ownerCellKey = spillOwnerCellKeys(doc)
  const followersByOwner = (generated: readonly GeneratedCellRecord[]): Map<string, string[]> => {
    const out = new Map<string, string[]>()
    for (const record of generated) {
      const owner = ownerCellKey(record.ownerKey)
      if (owner === undefined) continue
      const follower = makeCellKey(record.sheet.name, record.cell.col, record.cell.row)
      const list = out.get(owner)
      if (list) list.push(follower); else out.set(owner, [follower])
    }
    return out
  }
  let generatedByOwner = followersByOwner(prior?.generated ?? [])
  /** The formulas reading one cell: by the syntactic index and by the edges the last evaluation observed. */
  const readersOfCell = (target: string, reverse: ReadonlyMap<string, ReadonlySet<string>>): string[] => {
    const out: string[] = []
    const bang = target.lastIndexOf('!'), coords = /^(\d+):(\d+)$/.exec(target.slice(bang + 1))
    if (bang >= 0 && coords) {
      const col = Number(coords[1]), row = Number(coords[2])
      for (const reader of queryReferences(index, target.slice(0, bang), { minRow: row, maxRow: row, minCol: col, maxCol: col })) out.push(reader)
    }
    for (const reader of reverse.get(target) ?? []) out.push(reader)
    return out
  }
  /** Drains `pending` into `add`: each target's readers, and the readers of the cells it spills. */
  const drain = (pending: string[], add: (key: string) => void, followers: ReadonlyMap<string, readonly string[]>, reverse: ReadonlyMap<string, ReadonlySet<string>>) => {
    for (let cursor = 0; cursor < pending.length; cursor++) {
      const target = pending[cursor]!
      for (const reader of readersOfCell(target, reverse)) add(reader)
      for (const follower of followers.get(target) ?? []) for (const reader of readersOfCell(follower, reverse)) add(reader)
    }
  }
  drain(queue, enqueue, generatedByOwner, dynamicReverse)
  if (!formulaKeys.size) return { recomputed: 0, charts: [], sheets: [], stale: [], readers: [] }
  const staleKeys = new Set<string>()
  const estimatedVisits = dependentWorkEstimate(doc, index, formulaKeys, prior?.runtimeIndex)
  if (estimatedVisits > MAX_DEPENDENT_CELL_VISITS) {
    for (const key of formulaKeys) staleKeys.add(key)
    const retained = prior?.retainedUnsupportedKeys ?? new Set<string>()
    for (const key of formulaKeys) retained.add(key)
    if (prior) prior.retainedUnsupportedKeys = retained
    const sheets = [...new Set([...formulaKeys].map(key => doc.sheets.findIndex(sheet => sheet.name.toLowerCase() === key.slice(0, key.lastIndexOf('!')))).filter(i => i >= 0))]
    return { recomputed: 0, charts: [], sheets, stale: [...staleKeys], readers: [...formulaKeys], overBudget: true }
  }
  const recomputeKeys = formulaKeys
  let snapshots = liveFormulaSnapshots.get(doc)
  if (!snapshots) liveFormulaSnapshots.set(doc, snapshots = new Map())
  const snapshotFormulas = (keys: Iterable<string>) => {
    for (const key of keys) {
      if (snapshots!.has(key)) continue
      const cell = prior?.runtimeIndex?.cellMap.get(key)
      if (cell) snapshots!.set(key, { cell, content: snapshotCell(cell) })
    }
  }
  snapshotFormulas(recomputeKeys)
  // Ruling S3-AL: the saved state of every position the spill work writes or clears (a spill that grows or shrinks), so
  // print and the restore put back exactly the saved cells, and the restored cells are the saved generation's again.
  let images = liveSpillImages.get(doc)
  if (!images) liveSpillImages.set(doc, images = { positions: new Map(), generated: [...(prior?.generated ?? [])], createdRows: [...(prior?.createdRows ?? [])] })
  const positions = images.positions, runtime = prior?.runtimeIndex
  const evaluateSubset = (keys: ReadonlySet<string>) => {
    spillImageRecorder = (sheet, col, row) => {
      const key = makeCellKey(sheet.name, col, row)
      if (positions.has(key)) return
      const rowObj = sheet.rows.find(candidate => candidate.index === row)
      const cell = rowObj?.cells.find(candidate => candidate.col === col)
      const indexed = runtime?.cellMap.get(key), indexedSheet = runtime?.cellSheets.get(key)
      positions.set(key, { sheet, row, col, ...(rowObj ? { rowObj } : {}), ...(cell ? { cell, content: snapshotCell(cell) } : {}),
        ...(indexed ? { indexed } : {}), ...(indexedSheet !== undefined ? { indexedSheet } : {}), input: runtime?.inputKeys.has(key) ?? false })
    }
    try { evaluateWorkbookFormulasInternal(doc, { ...options, forceRecalc: true, intent: 'explicit-recalc' }, keys) }
    finally { spillImageRecorder = undefined }
  }
  evaluateSubset(recomputeKeys)
  // A spill that grew or shrank changed cells no formula reader was queued for (a cell it now covers, or one it left):
  // their readers are recomputed after it, until no spill moves further (bounded; past the bound, or over the work
  // budget, they keep their values and are marked stale).
  const spilled = new Set<string>()
  for (const owner of recomputeKeys) for (const follower of generatedByOwner.get(owner) ?? []) spilled.add(follower)
  for (let pass = 0; ; pass++) {
    const state = docSpillStates.get(doc)
    const after = followersByOwner(state?.generated ?? [])
    // Every cell a recomputed formula spilled before or spills now: its readers not yet recomputed are owed.
    const touched = new Set<string>()
    for (const owner of recomputeKeys) for (const follower of [...(generatedByOwner.get(owner) ?? []), ...(after.get(owner) ?? [])]) { touched.add(follower); spilled.add(follower) }
    const extra = new Set<string>(), pending: string[] = []
    const add = (key: string) => { if (index.byKey.has(key) && !recomputeKeys.has(key) && !extra.has(key)) { extra.add(key); pending.push(key) } }
    const reverse = state?.reverseFormulaDeps ?? new Map<string, Set<string>>()
    for (const cell of touched) for (const reader of readersOfCell(cell, reverse)) add(reader)
    drain(pending, add, after, reverse)
    generatedByOwner = after
    if (!extra.size) break
    if (pass >= MAX_SPILL_FOLLOW_PASSES || dependentWorkEstimate(doc, index, extra, prior?.runtimeIndex) > MAX_DEPENDENT_CELL_VISITS) {
      const retained = state?.retainedUnsupportedKeys ?? new Set<string>()
      for (const key of extra) { staleKeys.add(key); retained.add(key) }
      if (state) state.retainedUnsupportedKeys = retained
      break
    }
    snapshotFormulas(extra)
    for (const key of extra) recomputeKeys.add(key)
    evaluateSubset(extra)
  }
  const finalFollowers = followersByOwner(docSpillStates.get(doc)?.generated ?? [])
  const affectedSheets = [...new Set([...formulaKeys].map(key => {
    const bang = key.lastIndexOf('!')
    return doc.sheets.findIndex(sheet => sheet.name.toLowerCase() === key.slice(0, bang))
  }).filter(index => index >= 0))]
  const retained = docSpillStates.get(doc)?.retainedUnsupportedKeys ?? new Set<string>()
  // A reader can become stale because an unsupported formula was first
  // discovered during this very subset evaluation. Propagate that result to
  // its transitive readers (the readers of its spilled cells included) and restore their own pre-run cached values.
  const newlyStale = [...recomputeKeys].filter(key => retained.has(key))
  const afterQueue = [...newlyStale]
  for (const key of newlyStale) staleKeys.add(key)
  const finalReverse = docSpillStates.get(doc)?.reverseFormulaDeps ?? new Map<string, Set<string>>()
  for (let i = 0; i < afterQueue.length; i++) {
    const target = afterQueue[i]!
    const readers = new Set(readersOfCell(target, finalReverse))
    for (const follower of finalFollowers.get(target) ?? []) for (const reader of readersOfCell(follower, finalReverse)) readers.add(reader)
    for (const reader of readers) if (index.byKey.has(reader) && !staleKeys.has(reader)) {
      staleKeys.add(reader); retained.add(reader)
      const original = snapshots.get(reader), formula = prior?.runtimeIndex?.cellMap.get(reader)
      if (original && formula) restoreCell(formula, original.content)
      afterQueue.push(reader)
    }
  }
  for (const key of staleKeys) retained.add(key)
  // A chart over the cells a recomputed formula spills (or spilled) follows it too, not only one over its anchor.
  return { recomputed: recomputeKeys.size, readers: [...recomputeKeys], charts: chartsUsing(doc, new Set([...recomputeKeys, ...spilled])), sheets: affectedSheets, stale: [...new Set([...staleKeys, ...recomputeKeys].filter(key => retained.has(key)))] }
}

/** By sheet, row and column: the saved cell at every position a dependent recompute changed (a copy), or null where
 * the saved workbook had no cell; undefined when nothing changed. That covers the formula cells it recomputed and the
 * cells their spills wrote or cleared. Ruling S3-AL: print, which never runs a pivot session, shows these beside the
 * saved pivot. */
export type SavedDependentCells = ReadonlyMap<XlsxSheet, ReadonlyMap<number, ReadonlyMap<number, XlsxCell | null>>>
export function savedDependentCells(doc: XlsxDocument): SavedDependentCells | undefined {
  const snapshots = liveFormulaSnapshots.get(doc), images = liveSpillImages.get(doc)
  if (!snapshots?.size && !images?.positions.size) return undefined
  const out = new Map<XlsxSheet, Map<number, Map<number, XlsxCell | null>>>()
  const put = (sheet: XlsxSheet, row: number, col: number, cell: XlsxCell | null) => {
    let rows = out.get(sheet)
    if (!rows) out.set(sheet, rows = new Map())
    let cols = rows.get(row)
    if (!cols) rows.set(row, cols = new Map())
    if (!cols.has(col)) cols.set(col, cell)
  }
  // Copies of the saved content, never the live cell spread over it: the live cell may carry keys the saved one lacks.
  for (const image of images?.positions.values() ?? []) put(image.sheet, image.row, image.col, image.cell ? { ...image.content!, ref: image.cell.ref, col: image.col, row: image.row } : null)
  for (const [key, { cell, content }] of snapshots ?? []) {
    const name = key.slice(0, key.lastIndexOf('!')), sheet = doc.sheets.find(candidate => candidate.name.toLowerCase() === name)
    if (sheet) put(sheet, cell.row, cell.col, { ...content, ref: cell.ref, col: cell.col, row: cell.row })
  }
  return out
}

/** Restore the saved formula values after a resident viewer releases its live pivot contexts, and every position their
 * spills wrote or cleared: cells, rows, the runtime index and the saved generation's spill records, so a later
 * recompute owns the saved spills again. Returns the indexes of the sheets whose cells it restored. */
export function restoreDependentFormulaValues(doc: XlsxDocument): number[] {
  const snapshots = liveFormulaSnapshots.get(doc), images = liveSpillImages.get(doc)
  if (!snapshots && !images) return []
  const names = new Set<string>()
  if (images) {
    const state = docSpillStates.get(doc), runtime = state?.runtimeIndex
    const emptied: Array<{ sheet: XlsxSheet; row: number }> = []
    for (const [key, image] of images.positions) {
      const { sheet } = image
      names.add(sheet.name.toLowerCase())
      let rowObj = sheet.rows.find(candidate => candidate.index === image.row)
      const now = rowObj?.cells.find(candidate => candidate.col === image.col)
      if (!image.cell) {
        if (rowObj && now) rowObj.cells.splice(rowObj.cells.indexOf(now), 1)
        if (!image.rowObj) emptied.push({ sheet, row: image.row })
      } else {
        if (!rowObj) {
          rowObj = image.rowObj ?? { index: image.row, cells: [] }
          const at = sheet.rows.findIndex(candidate => candidate.index > image.row)
          if (at < 0) sheet.rows.push(rowObj); else sheet.rows.splice(at, 0, rowObj)
        }
        if (now !== image.cell) {
          if (now) rowObj.cells.splice(rowObj.cells.indexOf(now), 1)
          const at = rowObj.cells.findIndex(candidate => candidate.col > image.col)
          if (at < 0) rowObj.cells.push(image.cell); else rowObj.cells.splice(at, 0, image.cell)
        }
        restoreCell(image.cell, image.content!)
      }
      if (runtime) {
        if (image.indexed) runtime.cellMap.set(key, image.indexed); else runtime.cellMap.delete(key)
        if (image.indexedSheet !== undefined) runtime.cellSheets.set(key, image.indexedSheet); else runtime.cellSheets.delete(key)
        if (image.input) runtime.inputKeys.add(key); else runtime.inputKeys.delete(key)
      }
    }
    // Rows the saved workbook did not have go again once empty.
    for (const { sheet, row } of emptied) removeCreatedRow(sheet, row)
    if (state) { state.generated = images.generated; state.createdRows = images.createdRows }
    liveSpillImages.delete(doc)
  }
  for (const [key, { cell, content }] of snapshots ?? []) { restoreCell(cell, content); names.add(key.slice(0, key.lastIndexOf('!'))) }
  snapshots?.clear()
  liveFormulaSnapshots.delete(doc)
  // The restored cells carry the same formula text: the syntactic index stays valid (M-1).
  return doc.sheets.flatMap((sheet, index) => names.has(sheet.name.toLowerCase()) ? [index] : [])
}

/**
 * Whether an AST contains structured table identifiers. Shared copies keep
 * table expressions verbatim, so a master whose tables cannot bind in this
 * document cannot produce usable follower formulas either.
 */
function astContainsTable(node: AstNode): boolean {
  switch (node.type) {
    case 'table':
      return true
    case 'unary':
      return astContainsTable(node.expr)
    case 'implicitIntersect':
      return astContainsTable(node.expr)
    case 'binary':
      return astContainsTable(node.left) || astContainsTable(node.right)
    case 'call':
      return node.args.some(astContainsTable)
    case 'union':
      return node.refs.some(astContainsTable)
    case 'intersect':
      return astContainsTable(node.left) || astContainsTable(node.right)
    default:
      return false
  }
}

/**
 * Resolves shared formulas within a sheet by translating relative offsets
 * from the master cell (`cell.formula` defined with `si`) to follower cells.
 */
export function resolveSheetSharedFormulas(sheet: XlsxSheet, doc?: XlsxDocument): void {
  const masterBySi = new Map<number, XlsxCell>()

  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      if (cell.sharedFormula?.si !== undefined && cell.formula) {
        masterBySi.set(cell.sharedFormula.si, cell)
      }
    }
  }

  // A master that cannot be parsed must NEVER emit its diagnostic sentence as
  // follower formula source: followers keep their cached values (bounded by
  // the recalc-flag rules) and the ORIGINAL master text is preserved verbatim
  // in the diagnostic for provenance.
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      if (cell.sharedFormula?.si !== undefined && !cell.formula) {
        const master = masterBySi.get(cell.sharedFormula.si)
        if (master && master.formula) {
          const masterAst = parseFormula(master.formula)
          // A CANONICAL error literal (#N/A/#DIV/0!/#REF!/#NAME?…) is a VALID
          // master — the group resolves as a real error, no provenance
          // diagnostic. Only parser-sentence failures skip with provenance.
          const canonicalErrorMaster = masterAst.type === 'error' && isFormulaError(masterAst.error)
          // A parsed master whose table identifiers cannot bind in this
          // document (no table inventory) keeps followers cached with the
          // same master-source provenance instead of emitting unusable
          // follower formulas.
          const unboundTables =
            !canonicalErrorMaster &&
            masterAst.type !== 'error' &&
            (doc?.tables?.length ?? 0) === 0 &&
            astContainsTable(masterAst)
          if ((masterAst.type === 'error' && !canonicalErrorMaster) || unboundTables) {
            const holder = sheet as XlsxSheet & { diagnostics?: Array<{ kind: string; feature: string; message: string }> }
            holder.diagnostics ??= []
            const message = `Shared formula group si=${cell.sharedFormula.si} master could not be parsed; follower ${cell.ref || `r${cell.row}c${cell.col}`} keeps its cached value. Master source preserved: ${master.formula}`
            if (!holder.diagnostics.some(d => d.message === message)) {
              holder.diagnostics.push({ kind: 'unsupported-formula', feature: 'shared-formula', message })
            }
            continue
          }
          const dCol = cell.col - master.col
          const dRow = cell.row - master.row
          const { formula } = translateSharedFormula(master.formula, dCol, dRow)
          cell.formula = formula
        }
      }
    }
  }
}

/**
 * Evaluates all formulas across an XLSX workbook using a single shared
 * EvaluationContext and per-cell memoization, supporting cross-sheet references,
 * shared formula groups, circular reference cycle poisoning, and the C2
 * phase-two spill/legacy lifecycle (discovery, atomic publication, own-only
 * shrink cleanup and A1# resolution).
 */
export function evaluateWorkbookFormulas(
  doc: XlsxDocument,
  options?: WorkbookEvaluationOptions
): void {
  evaluateWorkbookFormulasInternal(doc, options)
}

/** A bounded subset run used only by recomputeDependents. `onlyKeys` also
 * scopes old spill cleanup and preserves unaffected dependency edges. */
/** Each sheet's identity as the spill engine keys owners by it (`spillCellKey`: the workbook sheet id, as the
 * evaluation's identity sheets name it), by lower-cased sheet name, and back. */
function sheetIdentities(doc: XlsxDocument): { idOf: Map<string, string>; nameOf: Map<string, string> } {
  const idOf = new Map<string, string>(), nameOf = new Map<string, string>()
  for (const sheet of doc.workbookSheets ?? doc.sheets.map(candidate => ({ sheetId: candidate.sheetId ?? candidate.name, name: candidate.name }))) {
    const name = sheet.name.toLowerCase(), id = String(sheet.sheetId)
    if (!idOf.has(name)) idOf.set(name, id)
    if (!nameOf.has(id)) nameOf.set(id, name)
  }
  return { idOf, nameOf }
}
/** `keys` (cell keys, `makeCellKey`) as spill owner keys (`spillCellKey`), with the cell keys themselves kept too. */
function spillOwnerKeys(doc: XlsxDocument, keys: ReadonlySet<string>): Set<string> {
  const { idOf } = sheetIdentities(doc)
  const out = new Set<string>(keys)
  for (const key of keys) {
    const bang = key.lastIndexOf('!'), id = idOf.get(key.slice(0, bang))
    if (bang >= 0 && id !== undefined) out.add(`${id}${key.slice(bang)}`)
  }
  return out
}
/** A spill owner key (`spillCellKey`) as the owner's cell key (`makeCellKey`), the key every reader queue uses. */
function spillOwnerCellKeys(doc: XlsxDocument): (ownerKey: string) => string | undefined {
  const { nameOf } = sheetIdentities(doc)
  return ownerKey => {
    const bang = ownerKey.lastIndexOf('!'), name = nameOf.get(ownerKey.slice(0, bang))
    return bang >= 0 && name !== undefined ? `${name}${ownerKey.slice(bang)}` : undefined
  }
}

function evaluateWorkbookFormulasInternal(
  doc: XlsxDocument,
  options?: WorkbookEvaluationOptions,
  onlyKeys?: ReadonlySet<string>,
): void {
  if (workbookValuePolicy(doc) === 'cached') {
    throw new RangeError('This workbook shows saved values only; its formulas are never evaluated')
  }
  // Cheap early-exit: skip evaluation work entirely when no cell has a formula.
  // A what-if `<f t="dataTable"/>` may carry no formula text at all; it is
  // still unsupported and must be diagnosed, so the actual what-if type keeps
  // the scan from exiting before the diagnosis branch. Ordinary formula-less,
  // type-less content (the common case) still exits cheaply here.
  let hasFormulas = Boolean(onlyKeys?.size)
  if (!onlyKeys) for (const sheet of doc.sheets) {
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        if (cell.formula !== undefined || cell.sharedFormula !== undefined || isWhatIfDataTable(cell.formulaType)) {
          hasFormulas = true
          break
        }
      }
      if (hasFormulas) break
    }
    if (hasFormulas) break
  }
  if (!hasFormulas) return

  // One resolved semantics object per recalculation generation, shared by
  // evaluation AND render (propagated to every sheet) so DATE/TEXT and the
  // cell formatter agree on the actual date system/locale without callers
  // injecting semantics manually.
  /** Phase 2: cohort cells that ACTUALLY reached a volatile function (NOW/TODAY
   * sample the clock through this semantics getter). A dead lazy branch that
   * never runs records nothing. */
  const volatileKeys = new Set<string>()
  const semantics = resolveWorkbookSemantics(doc, options, () => {
    if (currentFrame) volatileKeys.add(currentFrame.key)
  })
  doc.semantics = semantics
  for (const sheet of doc.sheets) sheet.semantics = semantics

  const arrayMode = options?.arrayMode ?? 'from-file'

  // 0. Own-only cleanup of the previous generation's generated cells. A cell
  //    is reverted ONLY when it is still the exact object we wrote AND its
  //    content is byte-for-byte what we wrote; a user replacement (new object)
  //    or ANY user content edit (value, formula, style, …) is preserved and
  //    becomes ordinary input. An ORIGINAL blank/styled cell is RESTORED to its
  //    snapshot (never removed); only cells/rows we created are removed.
  const prior = docSpillStates.get(doc)
  const retainedUnsupportedKeys = new Set(onlyKeys ? prior?.retainedUnsupportedKeys ?? [] : [])
  if (onlyKeys) for (const key of onlyKeys) retainedUnsupportedKeys.delete(key)
  // A no-op manual file-load recalculates nothing (manual gates ca /
  // fullCalcOnLoad / uncached too), so the previous generation's owned spill
  // followers, created rows and reverse-dependency registry are preserved: no
  // premature cleanup before the policy decision. An explicit
  // force/explicit-recalc still performs the own-only cleanup below. This is a
  // cleanup filter, not an early return: semantics propagation and the
  // what-if diagnostics in the decision loop still run.
  const manualFileLoadNoOp = options?.forceRecalc !== true &&
    (options?.intent ?? 'file-load') === 'file-load' &&
    doc.calc?.calcMode === 'manual'
  // The spill owners among `onlyKeys`, keyed as the spill records key them (`spillCellKey`: the sheet's identity, not
  // its lower-cased name), so a subset recalculation reverts, and does not retain, the followers of the owners it
  // recalculates: a dependent spill that shrinks leaves none of its earlier cells behind (Ruling S3-AL).
  const onlyOwners = onlyKeys ? spillOwnerKeys(doc, onlyKeys) : undefined
  const removedKeys = new Set<string>()
  if (prior && !manualFileLoadNoOp) {
    for (const g of prior.generated) {
      if (onlyOwners && !onlyOwners.has(g.ownerKey)) continue
      const rowObj = g.sheet.rows.find((r) => r.index === g.cell.row)
      if (!rowObj || !rowObj.cells.includes(g.cell)) continue // replaced object → user input
      if (!cellContentMatches(g.cell, g.written)) continue // user content edit → user input
      spillImageRecorder?.(g.sheet, g.cell.col, g.cell.row)
      if (g.original) restoreCell(g.cell, g.original)
      else {
        removeCell(g.sheet, g.cell)
        prior.runtimeIndex?.cellMap.delete(g.key)
        prior.runtimeIndex?.cellSheets.delete(g.key)
      }
      removedKeys.add(g.key)
    }
    for (const row of prior.createdRows) {
      if (onlyOwners && !onlyOwners.has(row.ownerKey)) continue
      removeCreatedRow(row.sheet, row.index)
    }
  }
  const valueDeps = prior?.valueDeps ?? new Map<string, Set<string>>()
  if (onlyKeys) for (const key of onlyKeys) valueDeps.delete(key)
  const formulaDeps = prior?.formulaDeps ?? new Map<string, Set<string>>()
  const reverseDeps = prior?.reverseFormulaDeps ?? reverseFormulaEdges(formulaDeps)
  if (onlyKeys) for (const key of onlyKeys) {
    for (const target of formulaDeps.get(key) ?? []) removeReverseFormulaEdge(reverseDeps, target, key)
    formulaDeps.delete(key)
  }
  /** ACTUAL (lazily discovered) dependency edges for this generation, kept
   * separate from `valueDeps` so established spill-cleanup semantics are
   * unchanged. Populated only by real reads — dead lazy branches add nothing. */
  const generationDeps = new Map<string, Set<string>>()
  const forceKeys = prior && removedKeys.size > 0 ? affectedReaders(removedKeys, prior.valueDeps) : new Set<string>()

  const cachedIndex = onlyKeys ? prior?.runtimeIndex : undefined
  const sheetsByName = cachedIndex?.sheetsByName ?? new Map<string, XlsxSheet>()
  const cellMap = cachedIndex?.cellMap ?? new Map<string, XlsxCell>()
  const cellSheets = cachedIndex?.cellSheets ?? new Map<string, string>()
  const sheetMaxCols = cachedIndex?.sheetMaxCols ?? new Map<string, number>()
  const sheetMaxRows = cachedIndex?.sheetMaxRows ?? new Map<string, number>()

  // 1. Resolve shared formulas on all sheets and index sheets, cells, and bounds
  if (!cachedIndex) for (const sheet of doc.sheets) {
    resolveSheetSharedFormulas(sheet, doc)
    sheetsByName.set(sheet.name.toLowerCase(), sheet)

    let maxCol = -1
    let maxRow = -1
    for (const row of sheet.rows) {
      if (row.index > maxRow) maxRow = row.index
      for (const cell of row.cells) {
        if (cell.col > maxCol) maxCol = cell.col
        if (cell.row > maxRow) maxRow = cell.row
        const key = makeCellKey(sheet.name, cell.col, cell.row)
        cellMap.set(key, cell)
        cellSheets.set(key, sheet.name)
      }
    }
    sheetMaxCols.set(sheet.name.toLowerCase(), Math.max(0, maxCol))
    sheetMaxRows.set(sheet.name.toLowerCase(), Math.max(0, maxRow))
    // Surface per-sheet formula diagnostics (e.g. unparsable shared masters)
    // on the document channel, deduped by feature+message.
    const sheetDiags = (sheet as unknown as { diagnostics?: Array<{ kind: string; feature: string; message: string }> }).diagnostics
    if (sheetDiags?.length) {
      const holder = doc as unknown as { diagnostics?: Array<{ kind: string; feature: string; message: string }> }
      holder.diagnostics ??= []
      for (const d of sheetDiags) {
        if (!holder.diagnostics.some(x => x.feature === d.feature && x.message === d.message)) holder.diagnostics.push(d)
      }
    }
  }

  const generation = ++workbookReferenceGeneration

  // 2. Original occupancy snapshot BEFORE any formula evaluation. Generated
  //    cells from a prior generation were already removed above, so every
  //    remaining cell is an original input/formula or a user-preserved cell.
  const inputKeys = cachedIndex?.inputKeys ?? new Set<string>()
  const formulaKeys = cachedIndex?.formulaKeys ?? new Set<string>()
  const mergeKeys = cachedIndex?.mergeKeys ?? new Set<string>()
  if (!cachedIndex) for (const sheet of doc.sheets) {
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        const key = makeCellKey(sheet.name, cell.col, cell.row)
        if (cell.formula !== undefined || cell.sharedFormula !== undefined) formulaKeys.add(key)
        // Only NONBLANK original content blocks a spill; a blank (styled or not)
        // cell is a valid spill target and is restored on cleanup.
        else if (cell.value !== null && cell.value !== undefined) inputKeys.add(key)
      }
    }
    for (const merge of sheet.mergeRanges ?? []) {
      for (let row = merge.minRow; row <= merge.maxRow; row++) {
        for (let col = merge.minCol; col <= merge.maxCol; col++) mergeKeys.add(makeCellKey(sheet.name, col, row))
      }
    }
  }

  // 3. Identities + spill engine (before the frame loop).
  const identitySheets = cachedIndex?.identitySheets ?? (doc.workbookSheets ?? doc.sheets.map((s, i) => ({
    sheetId: s.sheetId ?? s.name,
    name: s.name,
    workbookIndex: s.workbookIndex ?? i,
  }))).map((s) => ({ sheetId: s.sheetId, name: s.name, workbookIndex: s.workbookIndex }))
  const sheetIdOfName = (name: string): SheetId | undefined =>
    identitySheets.find((s) => s.name.toLowerCase() === name.toLowerCase())?.sheetId
  const sheetNameOfId = (id: SheetId): string | undefined =>
    identitySheets.find((s) => s.sheetId === id)?.name

  const memo = new Map<string, EvaluationValue>()
  const activeStack: string[] = []
  const activeSet = new Set<string>()
  const cycleMembers = new Set<string>()
  const generatedList: GeneratedCellRecord[] = []
  const retainedGenerated = onlyOwners && prior ? prior.generated.filter(g => !onlyOwners.has(g.ownerKey)) : []
  /** Rows this recalculation CREATED (safe to drop when empty); original rows
   * carrying heightPt/customHeight are never removed. */
  const createdRows: CreatedRowRecord[] = []
  /** Pre-recalculation cache of each owner anchor, for arbitration rollback. */
  const prePublishCache = new Map<string, { value: string | number | boolean | null; valueIsError: boolean | undefined }>()
  const eligibleSaved = new Map<XlsxCell, ReturnType<typeof savedSpillFollowers>[number]>()
  const savedByAnchor = new Map<XlsxCell, Array<ReturnType<typeof savedSpillFollowers>[number]>>()
  for(const record of savedSpillFollowers(doc)){
    const key=makeCellKey(record.sheet.name,record.cell.col,record.cell.row)
    if(cellMap.get(key)===record.cell&&cellContentMatches(record.cell,record.written)){
      eligibleSaved.set(record.cell,record)
      const list=savedByAnchor.get(record.anchor)
      if(list)list.push(record);else savedByAnchor.set(record.anchor,[record])
    }
  }
  const activeSavedOwners=new Set<XlsxCell>()
  const blankSavedContent=(cell:XlsxCell):CellContent=>({...snapshotCell(cell),value:null,hasCachedValue:false,valueIsError:false})

  function addressOfName(sheetName: string, col: number, row: number): CellAddress {
    return { sheetId: sheetIdOfName(sheetName) ?? sheetName, col, row }
  }

  function isOriginalOccupant(address: CellAddress): boolean {
    const name = sheetNameOfId(address.sheetId)
    if (name === undefined) return false
    const key = makeCellKey(name, address.col, address.row)
    const cell=cellMap.get(key),saved=cell&&eligibleSaved.get(cell)
    if(saved&&activeSavedOwners.has(saved.anchor))return false
    if (formulaKeys.has(key)) return true
    // A live pivot cell is the overlay's: a value it shows is an original occupant (read through the overlay, never as a
    // spill's follower); a cell it vacated is not, even over a saved input (PR #57).
    const shown = options?.readCellValue?.(name, address.col, address.row)
    if (shown !== undefined) return shown !== null
    return inputKeys.has(key)
  }

  /** True when the region contains at least one non-original cell (so a spill
   * owner could legitimately own part of it). Conservative for huge regions. */
  function hasOwnableCell(areas: readonly RefArea[]): boolean {
    for (const area of areas) {
      if (area.cols <= 0 || area.rows <= 0) continue
      if (sheetNameOfId(area.sheetId) === undefined) return true
      if (area.cols * area.rows > 10000) return true
      for (let row = area.firstRow; row < area.firstRow + area.rows; row++) {
        for (let col = area.firstCol; col < area.firstCol + area.cols; col++) {
          if (!isOriginalOccupant({ sheetId: area.sheetId, col, row })) return true
        }
      }
    }
    return false
  }

  function tableCovers(sheetId: SheetId, col: number, row: number): boolean {
    for (const table of doc.tables ?? []) {
      const e = table.extent
      if (e.sheetId === sheetId && col >= e.firstCol && col < e.firstCol + e.cols && row >= e.firstRow && row < e.firstRow + e.rows) return true
    }
    return false
  }

  function originalBlocker(address: CellAddress): SpillBlockReason | undefined {
    const name = sheetNameOfId(address.sheetId)
    if (name === undefined) return undefined
    const key = makeCellKey(name, address.col, address.row)
    // A structured-table region blocks regardless of its cell contents (array
    // formulas inside an Excel Table are documented #SPILL!).
    if (tableCovers(address.sheetId, address.col, address.row)) return 'table'
    if (mergeKeys.has(key)) return 'merged-cell'
    if (formulaKeys.has(key)) return 'formula-cell'
    // A committed pivot's live cells (the view overlay, PR #57): a value it shows blocks a spill as any input does
    // (#SPILL! at the anchor, as Excel); a cell it vacated (null) is empty, whatever the saved sheet still holds there.
    const shown = options?.readCellValue?.(name, address.col, address.row)
    if (shown === null) return undefined
    if (shown !== undefined) return 'input-cell'
    const savedCell=cellMap.get(key),saved=savedCell&&eligibleSaved.get(savedCell)
    if(saved&&activeSavedOwners.has(saved.anchor))return undefined
    if (inputKeys.has(key)) return 'input-cell'
    return undefined
  }

  function reportIssue(issue: SpillIssue): void {
    const holder = doc as unknown as { diagnostics?: Array<{ kind: string; feature: string; message: string }> }
    holder.diagnostics ??= []
    if (!holder.diagnostics.some((d) => d.feature === issue.feature && d.message === issue.message)) {
      holder.diagnostics.push({ kind: issue.kind, feature: issue.feature, message: issue.message })
    }
  }

  /** Per-pass edge: during a bounded iterative pass only the current pass's
   * actual edges are tracked, so an adaptive/conditional dependency change is
   * detectable without polluting valueDeps/generationDeps. */
  function recordPassDependency(targetKey: string): void {
    const frame = currentFrame
    if (!frame) return
    const pass = passDeps.get(frame.key)
    if (pass) pass.add(targetKey)
    else passDeps.set(frame.key, new Set([targetKey]))
  }

  function recordDependency(targetKey: string): void {
    const frame = currentFrame
    if (!frame) return
    if (iterating) {
      recordPassDependency(targetKey)
      return
    }
    const set = valueDeps.get(frame.key)
    if (set) set.add(targetKey)
    else valueDeps.set(frame.key, new Set([targetKey]))
    const discovery = generationDeps.get(frame.key)
    if (discovery) discovery.add(targetKey)
    else generationDeps.set(frame.key, new Set([targetKey]))
    const graph = formulaDeps.get(frame.key)
    if (graph) graph.add(targetKey)
    else formulaDeps.set(frame.key, new Set([targetKey]))
    addReverseFormulaEdge(reverseDeps, targetKey, frame.key)
  }

  /** Discovery-only edge: records the ACTUAL read graph for iteration without
   * touching the established valueDeps graph. */
  function recordDiscoveryDependency(targetKey: string): void {
    const frame = currentFrame
    if (!frame) return
    if (iterating) {
      recordPassDependency(targetKey)
      return
    }
    const discovery = generationDeps.get(frame.key)
    if (discovery) discovery.add(targetKey)
    else generationDeps.set(frame.key, new Set([targetKey]))
    const graph = formulaDeps.get(frame.key)
    if (graph) graph.add(targetKey)
    else formulaDeps.set(frame.key, new Set([targetKey]))
    addReverseFormulaEdge(reverseDeps, targetKey, frame.key)
  }

  function writeGenerated(address: CellAddress, value: ElementValue, valueIsError: boolean, ownerKey: string): void {
    const name = sheetNameOfId(address.sheetId)
    if (name === undefined) return
    const sheet = sheetsByName.get(name.toLowerCase())
    if (!sheet) return
    spillImageRecorder?.(sheet, address.col, address.row)
    const existingRow = sheet.rows.find((r) => r.index === address.row)
    const existingCell = existingRow?.cells.find((c) => c.col === address.col)
    // Snapshot the ORIGINAL content of a pre-existing (blank/styled) cell so a
    // later cleanup restores it instead of destroying its formatting.
    const original = existingCell ? (eligibleSaved.has(existingCell) ? blankSavedContent(existingCell) : snapshotCell(existingCell)) : undefined
    const cell = ensureCell(sheet, address.col, address.row)
    const normalized = typeof value === 'number' ? round15(value) : value
    const primitive = isEvaluationError(normalized) ? normalized.code : normalized
    cell.value = primitive
    cell.valueIsError = valueIsError
    cell.hasCachedValue = true
    cell.formula = undefined
    if (!existingRow) createdRows.push({ sheet, index: address.row, ownerKey })
    const key = makeCellKey(name, address.col, address.row)
    cellMap.set(key, cell)
    cellSheets.set(key, name)
    memo.set(key, normalized)
    generatedList.push({ key, sheet, cell, ownerKey, written: snapshotCell(cell), ...(original ? { original } : {}) })
  }

  /** Revert one generated cell: restore an original cell's snapshot, or remove
   * a cell we created. Never destroys original formatting/metadata. */
  function revertGenerated(g: GeneratedCellRecord): void {
    spillImageRecorder?.(g.sheet, g.cell.col, g.cell.row)
    if (g.original) {
      restoreCell(g.cell, g.original)
    } else {
      removeCell(g.sheet, g.cell)
      cellMap.delete(g.key)
      cellSheets.delete(g.key)
    }
    memo.delete(g.key)
  }

  /** Arbitration rollback: revert a committed owner's followers, retain its
   * valid cache and gate it (internal #NAME? + marks). */
  function failCommittedOwner(ownerKey: string, anchor: CellAddress): void {
    for (let i = generatedList.length - 1; i >= 0; i--) {
      const g = generatedList[i]
      if (g.ownerKey !== ownerKey) continue
      revertGenerated(g)
      generatedList.splice(i, 1)
    }
    const name = sheetNameOfId(anchor.sheetId)
    if (name === undefined) return
    const anchorKey = makeCellKey(name, anchor.col, anchor.row)
    const cell = cellMap.get(anchorKey)
    const cache = prePublishCache.get(ownerKey)
    if (cell && cache) {
      cell.value = cache.value
      cell.valueIsError = cache.valueIsError
      memo.set(anchorKey, isFormulaError(cache.value) && cache.valueIsError !== false ? formulaError(cache.value) : cache.value)
      return
    }
    if (cell) {
      cell.value = '#NAME?'
      cell.valueIsError = true
      cell.hasCachedValue = true
    }
    memo.set(anchorKey, formulaError('#NAME?'))
  }

  const engine = new SpillEngine({
    sheetIdOfName,
    sheetNameOfId,
    isOriginalOccupant,
    hasOwnableCell,
    originalBlocker,
    writeGenerated,
    failCommittedOwner,
    report: reportIssue,
  }, generation, { maxArrayCells: options?.maxArrayCells, maxArraySide: options?.maxArraySide })

  // 4. Parse formulas that need calculation, then register potential owners
  //    from the prebuilt index (never from formula visitation order).
  const formulas = new Map<string, AstNode>()
  const policy: CalcPolicy = {
    intent: options?.intent ?? (options?.forceRecalc === true ? 'explicit-recalc' : 'file-load'),
    forceRecalc: options?.forceRecalc === true,
    fullCalcOnLoad: options?.fullCalcOnLoad === true || doc.calc?.fullCalcOnLoad === true,
    calcMode: (doc.calc?.calcMode ?? 'auto') as CalcMode,
    iterate: doc.calc?.iterate === true,
    iterateCount: doc.calc?.iterateCount ?? 100,
    iterateDelta: doc.calc?.iterateDelta ?? 0.001,
  }
  const reportWhatIfDataTable = (cell: XlsxCell): void => {
    const holder = doc as unknown as { diagnostics?: Array<{ kind: string; feature: string; message: string }> }
    holder.diagnostics ??= []
    const message = `What-if data-table formula at ${cell.ref} retained; what-if computation is outside the finite inventory`
    if (!holder.diagnostics.some((d) => d.feature === 'what-if-data-table-unverified' && d.message === message)) {
      holder.diagnostics.push({ kind: 'native-gate', feature: 'what-if-data-table-unverified', message })
    }
  }
  const reportManualDeferred = (cell: XlsxCell): void => {
    const holder = doc as unknown as { diagnostics?: Array<{ kind: string; feature: string; message: string }> }
    holder.diagnostics ??= []
    const message = `Formula at ${cell.ref} is uncomputed under calcMode=manual; recalculation deferred until an explicit request`
    if (!holder.diagnostics.some((d) => d.feature === 'calc-manual-deferred' && d.message === message)) {
      holder.diagnostics.push({ kind: 'calc-manual-retained', feature: 'calc-manual-deferred', message })
    }
  }
  for (const [key, cell] of cellMap) {
    if (onlyKeys && !onlyKeys.has(key)) continue
    if (!cell.formula) {
      // A what-if data-table cell may carry no formula text at all; it is still
      // unsupported and must be diagnosed, never silently computed.
      if (isWhatIfDataTable(cell.formulaType)) reportWhatIfDataTable(cell)
      continue
    }
    const decision = decideCellCalculation(policy, {
      formulaType: cell.formulaType,
      ca: cell.ca,
      hasCachedValue: cell.hasCachedValue,
      value: cell.value,
      forcedDependent: forceKeys.has(key) || Boolean(onlyKeys?.has(key)),
    })
    if (decision.whatIfDataTable) {
      reportWhatIfDataTable(cell)
      continue
    }
    if (decision.reason === 'manual-retained') {
      // An uncached formula left uncomputed by manual mode is a real deferred
      // calculation (no fabricated value); a valid cached file-load is not.
      const uncached = cell.hasCachedValue === undefined ? cell.value == null : !cell.hasCachedValue
      if (uncached) reportManualDeferred(cell)
      continue
    }
    if (decision.recalc) formulas.set(key, parseFormula(cell.formula))
  }
  // Phase 2: snapshot every recalculated cell's shipped cache so a cohort with
  // an untrusted seed can be left byte-for-byte as it arrived (never a guessed 0).
  const iterationSnapshot = new Map<string, { value: XlsxCell['value']; valueIsError: boolean | undefined; hasCachedValue: boolean | undefined }>()
  if (policy.iterate) {
    for (const key of formulas.keys()) {
      const snapshot = cellMap.get(key)!
      iterationSnapshot.set(key, { value: snapshot.value, valueIsError: snapshot.valueIsError, hasCachedValue: snapshot.hasCachedValue })
    }
  }
  for (const [key, ast] of formulas) {
    const cell = cellMap.get(key)!
    const sheetName = cellSheets.get(key)!
    const address = addressOfName(sheetName, cell.col, cell.row)
    if (cell.dynamicArray && arrayMode !== 'legacy') {
      const parsed=cell.arrayRef?parseFormula(cell.arrayRef):undefined
      const area=parsed&&(parsed.type==='range'||parsed.type==='cell')?gridAreas(parsed,address.sheetId)[0]:undefined
      engine.registerOwner(address,{kind:'dynamic',source:'verified-file'},area)
      activeSavedOwners.add(cell)
    } else if (cell.arrayRef) {
      // Legacy fixed f.ref: declared geometry is known before evaluation.
      const parsed = parseFormula(cell.arrayRef)
      const area = (parsed.type === 'range' || parsed.type === 'cell')
        ? gridAreas(parsed, address.sheetId)[0]
        : undefined
      if (area) {
        engine.registerOwner(address, { kind: 'legacy-array', source: 'verified-file' }, area)
      } else {
        reportIssue({
          kind: 'spill-provenance',
          feature: 'spill-provenance',
          message: `Legacy array range "${cell.arrayRef}" at ${cell.ref} could not be parsed; declared output geometry unavailable`,
          anchor: address,
        })
      }
    } else if (arrayMode === 'dynamic' && isArrayCapableRoot(ast)) {
      engine.registerOwner(address, { kind: 'dynamic', source: 'explicit-mode' })
    }
  }

  // Evaluation frames suspend only when a dynamically requested dependency is
  // unresolved. No recursive cell calls or static traversal of lazy branches.
  class PendingDependency {
    constructor(readonly key: string) {}
  }

  /** C4 r2 PRIVATE driver control (workbook-local, never exported): an active
   * cycle back-edge whose seed is UNVERIFIED. It is THROWN, never returned as a
   * value, so no Core formula function — not even IFERROR — can observe it,
   * catch it, or execute a fallback. `evaluateCell` catches it and retains the
   * shipped cache; runIterativeCohort reports `calc-iterative-seed-unverified`. */
  class UnavailableIterationSeed {
    constructor(readonly key: string) {}
  }

  interface RangeProgress { matrix: EvaluationValue[][]; next: number }
  interface Frame {
    key: string
    modeStores: ModeFrameStores
    ranges: Map<string, RangeProgress>
    /** FRAME-OWNED unsupported-feature marks: a child cell's suspension/run can
     * never clear the parent's marks (Agy's SUSPENSION.json defect). */
    unsupportedFeatures: Set<string>
    /** Scheduled/reached spill owner: evaluated once in array mode, published. */
    owner: boolean
    /** This frame's value was fed by an unresolved cycle (a provisional read),
     * so it must not publish a placeholder over the shipped cache. */
    provisional: boolean
  }

  function newStore(): ModeFrameStore {
    return { nodeValues: new WeakMap(), flatArgs: new WeakMap(), functionWork: new WeakMap() }
  }

  let currentFrame: Frame | undefined
  /** Phase 2 iteration state: true while a bounded in-place pass is in flight. */
  let iterating = false
  /** Current in-pass value of every cohort cell (seed → updated in place). */
  const iterValues = new Map<string, EvaluationValue>()
  /** Actual dependency edges of the cell currently being iterated (per pass). */
  const passDeps = new Map<string, Set<string>>()
  /** Cells whose generation-local value is provisional (an unresolved cycle fed
   * them); they must not publish a placeholder over their shipped cache. */
  const provisionalCells = new Set<string>()

  const ctx: EvaluationContext & { unsupportedFeatures?: Set<string> } = {
    typedValues: true,
    excelTextLengthVersion: unicodeVersionOf(semantics.unicode),
    unsupportedFeatures: new Set(),
    semantics,
    ...(options?.getPivotData ? { getPivotData: options.getPivotData } : {}),
    maxArrayCells: options?.maxArrayCells,
    maxArraySide: options?.maxArraySide,
    ...(options?.maxMaterializedCells !== undefined ? { materializationBudget: { limit: options.maxMaterializedCells, used: 0 } } : {}),
    getCellValue: (sheetName, col, row) => {
      const targetSheetName = sheetName ?? ctx.currentSheet
      if (!targetSheetName) return formulaError('#REF!')
      const resolvedSheet = sheetsByName.get(targetSheetName.toLowerCase())
      if (!resolvedSheet) return formulaError('#REF!')
      const key = makeCellKey(resolvedSheet.name, col, row)
      // Iterative pass: a cohort cell reads the current in-pass value (an
      // already-updated earlier cell or a previous-pass later cell) instead of
      // the legacy cycle sentinel.
      if (iterating && iterValues.has(key)) {
        recordPassDependency(key)
        return iterValues.get(key)!
      }
      if (activeSet.has(key)) {
        const index = activeStack.indexOf(key)
        for (let i = index; i < activeStack.length; i++) cycleMembers.add(activeStack[i])
        // Record the ACTUAL back/self edge for lazy dependency discovery; the
        // established valueDeps graph (spill cleanup) stays untouched.
        recordDiscoveryDependency(key)
        // Phase 2 discovery shadow: a cycle back-edge reads the active cell's
        // trusted numeric cache (never a guessed 0) and marks this frame so it
        // does not publish a placeholder over the shipped cache.
        if (currentFrame) currentFrame.provisional = true
        if (!policy.iterate) return 0
        const seed = iterationSeed(key)
        if (seed !== undefined) return seed
        // C4 r2: the active cell has NO trusted seed. Inventing a provisional 0
        // would decide a lazy branch and sample a volatile clock; returning a
        // typed error would be caught by IFERROR, which would execute its
        // fallback. Throw the PRIVATE driver control instead so no Core formula
        // function ever observes a value: evaluateCell abandons this frame and
        // retains the shipped cache.
        throw new UnavailableIterationSeed(key)
      }
      if (memo.has(key)) {
        // A memoized read is still an ACTUAL dependency edge for iteration
        // discovery (dependents of a cycle are otherwise invisible).
        recordDiscoveryDependency(key)
        if (currentFrame && provisionalCells.has(key)) currentFrame.provisional = true
        return memo.get(key)!
      }
      // C2 phase2: settle spill ownership BEFORE any absent/generated memo.
      // An original input/formula cell can never be a follower.
      const address = addressOfName(resolvedSheet.name, col, row)
      if (!isOriginalOccupant(address)) {
        engine.prepareCell(address, ctx)
        const owned = engine.readOwnedCell(address)
        if (owned !== undefined) {
          memo.set(key, owned)
          recordDependency(key)
          return owned
        }
        // A gated owner's unknown footprint is NOT a confirmed absence: mark
        // this frame and propagate the internal #NAME? gate so the reader keeps
        // its valid cache instead of memoizing a guessed blank.
        const gate = engine.gateForCell(address)
        if (gate) {
          ctx.unsupportedFeatures?.add(gate)
          memo.set(key, formulaError('#NAME?'))
          recordDependency(key)
          return formulaError('#NAME?')
        }
      }
      const cell = cellMap.get(key)
      const viewValue = options?.readCellValue?.(resolvedSheet.name, col, row)
      if (viewValue !== undefined) { memo.set(key, viewValue); recordDependency(key); return viewValue }
      if (!cell) {
        memo.set(key, null)
        recordDependency(key)
        return null
      }
      if (formulas.has(key)) throw new PendingDependency(key)
      // Parsed cells explicitly distinguish errors and strings. Legacy hand-built
      // models retain their previous canonical-error-string convention for all cells.
      const error = cell.valueIsError ?? isFormulaError(cell.value)
      const value = error && isFormulaError(cell.value) ? formulaError(cell.value) : cell.value
      memo.set(key, value)
      recordDependency(key)
      return value
    },

    getRangeValues: (sheetName, from, to) => {
      const targetSheetName = sheetName ?? ctx.currentSheet
      if (!targetSheetName) return [[formulaError('#REF!')]]

      const resolvedSheet = sheetsByName.get(targetSheetName.toLowerCase())
      if (!resolvedSheet) return [[formulaError('#REF!')]]

      const minCol = Math.min(from.col, to.col)
      const maxCol = Math.max(from.col, to.col)
      const minRow = Math.min(from.row, to.row)
      const maxRow = Math.max(from.row, to.row)

      // C2 phase2: settle every relevant spill owner before populated
      // enumeration or structural absent counting (whole axes included).
      engine.prepareRegion([{
        sheetId: sheetIdOfName(resolvedSheet.name) ?? resolvedSheet.name,
        firstCol: minCol,
        firstRow: minRow,
        cols: maxCol - minCol + 1,
        rows: maxRow - minRow + 1,
      }], ctx)

      const regionGate = engine.gateForRegion([{
        sheetId: sheetIdOfName(resolvedSheet.name) ?? resolvedSheet.name,
        firstCol: minCol,
        firstRow: minRow,
        cols: maxCol - minCol + 1,
        rows: maxRow - minRow + 1,
      }])
      if (regionGate) {
        ctx.unsupportedFeatures?.add(regionGate)
        return [[formulaError('#NAME?')]]
      }

      const usedMaxCol = sheetMaxCols.get(resolvedSheet.name.toLowerCase()) ?? 0
      const usedMaxRow = sheetMaxRows.get(resolvedSheet.name.toLowerCase()) ?? 0

      // Only clamp large/unbounded ranges (> 10k cells) to used bounds
      let clampedMaxCol = maxCol
      let clampedMaxRow = maxRow
      if ((maxCol - minCol + 1) * (maxRow - minRow + 1) > 10000) {
        clampedMaxCol = Math.min(maxCol, Math.max(minCol, usedMaxCol))
        clampedMaxRow = Math.min(maxRow, Math.max(minRow, usedMaxRow))
      }

      const totalCells = (clampedMaxCol - minCol + 1) * (clampedMaxRow - minRow + 1)
      if (totalCells > 100000) {
        return [[formulaError('#NUM!')]]
      }

      // Retain only completed range cells when a dependency suspends this frame.
      // Resuming SUM over N uncached cells needs O(N) reads, not O(N squared).
      const rangeKey = `${resolvedSheet.name}!${minCol}:${minRow}:${clampedMaxCol}:${clampedMaxRow}`
      let progress = currentFrame?.ranges.get(rangeKey)
      if (!progress) {
        progress = { matrix: [], next: 0 }
        currentFrame?.ranges.set(rangeKey, progress)
      }
      const width = clampedMaxCol - minCol + 1
      while (progress.next < totalCells) {
        const rowIndex = Math.floor(progress.next / width)
        const colIndex = progress.next % width
        const value = ctx.getCellValue!(resolvedSheet.name, minCol + colIndex, minRow + rowIndex)
        const row = progress.matrix[rowIndex] ?? (progress.matrix[rowIndex] = [])
        row.push(value)
        progress.next++
      }
      return progress.matrix
    },
  }

  // B1 reference services: original-order identities, sparse reads through
  // the frame-aware cell callback (memo/cycle/suspension preserved), B0
  // defined-name/table inventories, and native-gate diagnostics.
  ctx.references = createReferenceServices({
    generation: () => generation,
    sheets: identitySheets,
    sheetIdOfName,
    sheetNameOfId,
    store: {
      read: (address) => {
        const name = sheetNameOfId(address.sheetId)
        if (name === undefined) return undefined
        const key = makeCellKey(name, address.col, address.row)
        const cell = cellMap.get(key)
        // Frame-aware read: pending formula dependencies suspend here and
        // resume the consumer on the same cursor entry (no skip/duplicate).
        const value = ctx.getCellValue!(name, address.col, address.row)
        if (value === undefined) return undefined
        const isFormulaCell = cell !== undefined && (cell.formula !== undefined || cell.sharedFormula !== undefined)
        const isGenerated = engine.ownerKeyForGenerated(address) !== undefined
        if (!isFormulaCell && !isGenerated && value === null) return undefined
        return { value: scalarProjection(value), origin: isFormulaCell ? 'formula' : isGenerated ? 'spill' : 'input' }
      },
      stored: function* () {
        const overlayByKey = new Map<string, { sheetName: string; col: number; row: number; value: EvaluationValue }>()
        for (const entry of options?.overlayEntries ?? []) overlayByKey.set(makeCellKey(entry.sheetName, entry.col, entry.row), entry)
        const seen = new Set<string>()
        for (const [key, cell] of cellMap) {
          if (seen.has(key)) continue
          const overlaid = overlayByKey.get(key)
          const name = cellSheets.get(key)
          if (name === undefined) continue
          const sheetId = sheetIdOfName(name)
          if (sheetId === undefined) continue
          const isFormulaCell = cell.formula !== undefined || cell.sharedFormula !== undefined
          if (!isFormulaCell && (cell.value === null || cell.value === undefined)) continue
          const viewValue = options?.readCellValue?.(name, cell.col, cell.row)
          if (viewValue === null) continue
          let value: ElementValue = overlaid ? scalarProjection(overlaid.value) : viewValue !== undefined ? scalarProjection(viewValue) : cell.value
          if (overlaid && value === null) continue
          if (!isFormulaCell) {
            const error = cell.valueIsError ?? isFormulaError(cell.value)
            value = error && isFormulaError(cell.value) ? formulaError(cell.value) : cell.value
          }
          yield {
            address: { sheetId, col: cell.col, row: cell.row },
            value,
            origin: isFormulaCell ? 'formula' : 'input',
          }
          seen.add(key)
        }
        for (const [key, entry] of overlayByKey) {
          if (seen.has(key) || scalarProjection(entry.value) === null) continue
          const sheetId = sheetIdOfName(entry.sheetName)
          if (sheetId === undefined) continue
          yield { address: { sheetId, col: entry.col, row: entry.row }, value: scalarProjection(entry.value), origin: 'input' }
        }
      },
    },
    definedNames: doc.definedNames ?? [],
    tables: doc.tables,
    onIssue: (feature, message) => {
      const holder = doc as unknown as { diagnostics?: Array<{ kind: string; feature: string; message: string }> }
      holder.diagnostics ??= []
      if (!holder.diagnostics.some((d) => d.feature === feature && d.message === message)) {
        holder.diagnostics.push({ kind: 'native-gate', feature, message })
      }
    },
  })
  ctx.reportFormulaIssue = (issue) => {
    const holder = doc as unknown as { diagnostics?: Array<{ kind: string; feature: string; message: string }> }
    holder.diagnostics ??= []
    const feature = issue.feature ?? issue.gate ?? 'reference'
    if (!holder.diagnostics.some((d) => d.feature === feature && d.message === issue.message)) {
      holder.diagnostics.push({ kind: issue.kind, feature, message: issue.message })
    }
  }
  // ABI hook: the spill engine requests an owner through the frame queue.
  ctx.pendingSpill = (request) => {
    throw new SpillSuspension(request)
  }
  // ABI spill services (workbook-owned registry): reads settle owners first.
  ctx.spills = {
    prepareCell: (address, c) => engine.prepareCell(address, c),
    prepareRegion: (areas, c) => engine.prepareRegion(areas, c),
    readOwnedCell: (address) => engine.readOwnedCell(address),
    resolveSpill: (anchor, c) => engine.resolveSpill(anchor, c),
  }

  function evaluateOwnerFrame(frame: Frame, cell: XlsxCell, formula: AstNode, address: CellAddress): EvaluationValue {
    const ownerKey = spillCellKey(address)
    // An owner already gated (arbitration/discovery) publishes nothing: internal
    // #NAME? + a frame mark so the workbook retains its valid cache.
    const preGate = engine.gateFeature(address)
    if (preGate) {
      frame.unsupportedFeatures.add(preGate)
      return formulaError('#NAME?')
    }
    if (cell.hasCachedValue) prePublishCache.set(ownerKey, { value: cell.value, valueIsError: cell.valueIsError })
    ctx.mode = 'array'
    ctx.nodeValues = frame.modeStores.array.nodeValues
    ctx.flatArgs = frame.modeStores.array.flatArgs
    ctx.functionWork = frame.modeStores.array.functionWork
    const result = evaluateFormulaInternal(formula, ctx)
    // A gate raised DURING this owner's own evaluation (e.g. an unresolved
    // self-owned shape) must not publish: internal #NAME? + frame mark only.
    const midGate = engine.gateFeature(address)
    if (midGate) {
      frame.unsupportedFeatures.add(midGate)
      return formulaError('#NAME?')
    }
    // C4 r3 PRE-PUBLICATION CONTROLLER GUARD. During the discovery shadow a
    // frame on (or fed by) an unresolved cycle is PROVISIONAL: its computed
    // matrix must never commit — publishing here would write generated
    // followers (B1=9) before the cohort is verified and a later retain gate
    // would restore only the anchor, leaving a partial matrix and poisoning
    // cached readers (D1=B1=99 → 9). Gate the owner through the existing
    // unavailable-owner path (internal #NAME? + explicit feature) so readers of
    // its unknown footprint keep their valid caches instead of a guessed 0,
    // and return the generation-local value. A verified cohort re-evaluates
    // under `iterating`, and an acyclic dynamic array (no cycle/provisional
    // mark) still publishes normally.
    if (policy.iterate && !iterating && (cycleMembers.has(frame.key) || frame.provisional)) {
      engine.markFailed(
        address,
        'calc-iterative-spill-unverified',
        `Iterative recalculation of the cyclic array/spill owner at ${cell.ref || ownerKey} is unverified; publication withheld and the shipped cache retained`,
      )
      return isMatrixValue(result) ? topLeftOfMatrix(result) : result
    }
    const record = engine.recordFor(address)
    // Unimplemented roots keep their entire saved cache; publication must not
    // retire a native follower while the anchor itself will retain its cache.
    if(frame.unsupportedFeatures.size>0){
      const feature=frame.unsupportedFeatures.values().next().value!
      engine.markFailed(address,feature,`Unsupported spill at ${cell.ref}; saved output retained`)
      return result
    }
    let out: EvaluationValue
    if (record?.provenance.kind === 'legacy-array' && cell.arrayRef) {
      out = isEvaluationError(result) || !isMatrixValue(result)
        ? engine.publish(address, undefined, result)
        : engine.publishLegacy(address, result)
    } else if (isEvaluationError(result)) {
      out = engine.publish(address, undefined, result)
    } else if (isMatrixValue(result)) {
      out = engine.publish(address, result, topLeftOfMatrix(result))
    } else {
      out = engine.publish(address, undefined, result)
    }
    const gate = engine.gateFeature(address)
    if (gate) frame.unsupportedFeatures.add(gate)
    if(!gate&&activeSavedOwners.has(cell)){
      // Retire only unchanged parser-owned caches that were not replaced by a
      // newly published follower. Edited cells remain ordinary blockers/input.
      for(const saved of savedByAnchor.get(cell)??[]){
        if(engine.ownerKeyForGenerated(addressOfName(saved.sheet.name,saved.cell.col,saved.cell.row))!==undefined)continue
        spillImageRecorder?.(saved.sheet,saved.cell.col,saved.cell.row)
        Object.assign(saved.cell,blankSavedContent(saved.cell))
        const key=makeCellKey(saved.sheet.name,saved.cell.col,saved.cell.row)
        inputKeys.delete(key);memo.delete(key)
      }
    }
    return out
  }

  function evaluateCell(rootKey: string): void {
    if (!formulas.has(rootKey)) return
    // A bounded iterative pass re-evaluates every cohort cell even though a
    // previous-generation value is memoized.
    if (!iterating && memo.has(rootKey)) return
    const frames: Frame[] = []
    const push = (key: string) => {
      const cell = cellMap.get(key)!
      const sheetName = cellSheets.get(key)!
      frames.push({
        key,
        modeStores: { scalar: newStore(), array: newStore() },
        ranges: new Map(),
        unsupportedFeatures: new Set(),
        owner: engine.hasOwner(addressOfName(sheetName, cell.col, cell.row)),
        provisional: false,
      })
      activeStack.push(key)
      activeSet.add(key)
    }
    push(rootKey)
    while (frames.length > 0) {
      const frame = frames[frames.length - 1]
      currentFrame = frame
      const cell = cellMap.get(frame.key)!
      ctx.currentSheet = cellSheets.get(frame.key)!
      ctx.currentCell = { col: cell.col, row: cell.row, absCol: false, absRow: false }
      ctx.mode = 'scalar'
      ctx.modeStores = frame.modeStores
      ctx.nodeValues = frame.modeStores.scalar.nodeValues
      ctx.flatArgs = frame.modeStores.scalar.flatArgs
      ctx.functionWork = frame.modeStores.scalar.functionWork
      ctx.unsupportedFeatures = frame.unsupportedFeatures
      // Resolved workbook identity plus containing-table context for
      // unqualified structured references.
      const addressSheetId = sheetIdOfName(ctx.currentSheet)
      let currentAddress: CellAddress | undefined
      if (addressSheetId !== undefined) {
        currentAddress = { sheetId: addressSheetId, col: cell.col, row: cell.row }
      }
      ctx.currentAddress = currentAddress
      ctx.tableContext = undefined
      if (currentAddress && doc.tables) {
        for (const table of doc.tables) {
          if (
            table.sheetId === currentAddress.sheetId &&
            cell.col >= table.extent.firstCol &&
            cell.col < table.extent.firstCol + table.extent.cols &&
            cell.row >= table.extent.firstRow &&
            cell.row < table.extent.firstRow + table.extent.rows
          ) {
            ctx.tableContext = { tableId: table.id, currentCell: currentAddress }
            break
          }
        }
      }
      let value: EvaluationValue
      try {
        value = frame.owner && currentAddress
          ? evaluateOwnerFrame(frame, cell, formulas.get(frame.key)!, currentAddress)
          : evaluateFormulaInternal(formulas.get(frame.key)!, ctx)
      } catch (error) {
        if (error instanceof SpillSuspension) {
          const owner = error.request.owner
          const ownerKey = makeCellKey(sheetNameOfId(owner.sheetId) ?? owner.sheetId, owner.col, owner.row)
          if (memo.has(ownerKey)) continue
          if (!formulas.has(ownerKey) || activeSet.has(ownerKey)) {
            // The requested owner cannot be scheduled (not recalculated, or an
            // active ancestor). Do not guess a blank: gate it and mark this
            // frame so its valid cache is retained.
            engine.markFailed(owner, 'spill-discovery', `Spill owner ${ownerKey} cannot be settled (${activeSet.has(ownerKey) ? 'active discovery cycle' : 'not scheduled for recalculation'}); ownership discovery gated`)
            frame.unsupportedFeatures.add('spill-discovery')
            continue
          }
          push(ownerKey)
          continue
        }
        if (error instanceof UnavailableIterationSeed) {
          // C4 r2: an untrusted cycle back-edge was read. No Core formula
          // function saw a value (IFERROR cannot run a fallback). Keep this
          // frame's shipped cache generation-local; runIterativeCohort reports
          // `calc-iterative-seed-unverified` and restores the shipped bytes.
          memo.set(frame.key, iterationSnapshotMemo(frame.key))
          provisionalCells.add(frame.key)
          frames.pop()
          activeStack.pop()
          activeSet.delete(frame.key)
          const parentFrame = frames[frames.length - 1]
          ctx.unsupportedFeatures = parentFrame ? parentFrame.unsupportedFeatures : frame.unsupportedFeatures
          continue
        }
        if (!(error instanceof PendingDependency)) throw error
        push(error.key)
        continue
      }
      // The legacy cycle sentinel applies only to the non-iterative path; the
      // iterative path resolves the cycle below instead of publishing 0.
      if (cycleMembers.has(frame.key) && !policy.iterate) value = 0
      // A spill gate mark propagates the internal #NAME? convention so the
      // retention path below keeps a valid cache (no guessed blank/zero).
      if (
        frame.unsupportedFeatures.has('spill-discovery') ||
        frame.unsupportedFeatures.has('spill-arbitration') ||
        frame.unsupportedFeatures.has('spill-provenance') ||
        frame.unsupportedFeatures.has('calc-iterative-spill-unverified')
      ) {
        value = formulaError('#NAME?')
      }
      // Unsupported constructs reached during THIS cell's evaluated path keep a
      // valid existing cache (with a diagnostic) instead of overwriting #NAME?;
      // unsupported is decided from the evaluated path only — lazy IF/IFERROR
      // branches that never ran record nothing. Marks are FRAME-OWNED: children
      // evaluated (or resumed) during suspension cannot clear the parent's.
      const unsupported = frame.unsupportedFeatures
      const prevValue = cell.value
      const retained = isEvaluationError(value) && value.code === '#NAME?' &&
        unsupported.size > 0 &&
        cell.hasCachedValue === true && prevValue !== null && prevValue !== undefined
      if (retained) {
        retainedUnsupportedKeys.add(frame.key)
        // Retention error identity uses the SAME legacy tri-state rule as
        // normal cached reads (SPEC FRAME-PROBE legacyRetainedError):
        // error = cell.valueIsError ?? isFormulaError(cell.value) — a TRUE or
        // UNDEFINED flag with a canonical code tags the memo as a real error
        // (dependent IFERROR catches); an EXPLICIT FALSE keeps raw text.
        const legacyErrorFlag = cell.valueIsError ?? isFormulaError(String(prevValue))
        const retainedForDependents = typeof prevValue === 'string' &&
          legacyErrorFlag && isFormulaError(prevValue)
          ? formulaError(prevValue)
          : (prevValue as EvaluationValue)
        memo.set(frame.key, retainedForDependents)
        frames.pop()
        activeStack.pop()
        activeSet.delete(frame.key)
        const parentFrame = frames[frames.length - 1]
        ctx.unsupportedFeatures = parentFrame ? parentFrame.unsupportedFeatures : frame.unsupportedFeatures
        const features = [...unsupported]
        unsupported.clear()
        const docHolder = doc as unknown as { diagnostics?: Array<{ kind: string; feature: string; message: string }> }
        docHolder.diagnostics ??= []
        const ref = cell.ref || `${cellSheets.get(frame.key)}!r${cell.row}c${cell.col}`
        for (const feature of features) {
          const message = `Formula at ${ref} uses unsupported ${feature}; prior cached value retained`
          if (!docHolder.diagnostics.some(d => d.feature === feature && d.message === message)) {
            docHolder.diagnostics.push({ kind: 'unsupported-formula', feature, message })
          }
        }
        continue
      }
      // Phase 2 discovery shadow: a frame on (or fed by) an unresolved cycle
      // must NOT publish a placeholder over the shipped cache. Its computed
      // value stays generation-local for graph discovery and is resolved by the
      // bounded iteration or by an explicit retain gate.
      if (policy.iterate && (cycleMembers.has(frame.key) || frame.provisional)) {
        if (value === null || value === undefined) value = 0
        const provisionalValue = typeof value === 'number' ? round15(value) : value
        provisionalCells.add(frame.key)
        memo.set(frame.key, provisionalValue)
        frames.pop()
        activeStack.pop()
        activeSet.delete(frame.key)
        const parentFrame = frames[frames.length - 1]
        ctx.unsupportedFeatures = parentFrame ? parentFrame.unsupportedFeatures : frame.unsupportedFeatures
        continue
      }
      // Scalar formula results return 0 for blank/empty results (native) and
      // normalize BEFORE memoization so dependents and the model agree; truly
      // absent INPUT cells keep null. Stored cell values carry Excel's
      // 15-significant-digit precision (single normalization at this boundary).
      if (value === null || value === undefined) value = 0
      const finalValue = typeof value === 'number' ? round15(value) : value
      cell.value = finalValue === 0 ? 0 : publicFormulaValue(finalValue)
      cell.valueIsError = isEvaluationError(finalValue)
      cell.hasCachedValue = true
      memo.set(frame.key, finalValue)
      frames.pop()
      activeStack.pop()
      activeSet.delete(frame.key)
    }
    cycleMembers.clear()
    currentFrame = undefined
    ctx.nodeValues = undefined
    ctx.flatArgs = undefined
    ctx.functionWork = undefined
    ctx.currentCell = undefined
    ctx.currentAddress = undefined
    ctx.tableContext = undefined
  }

  // -------------------------------------------------------------------------
  // Phase 2 — actual bounded iterative calculation.
  //
  // Runs only when the file enables iteration AND the ACTUAL (lazily
  // discovered) dependency graph contains a cycle. Acyclic work is already
  // published above; only the cyclic cohort plus its downstream dependents are
  // re-evaluated in canonical row/column order, in place, seeded from trusted
  // finite numeric caches. A missing/nonnumeric/Boolean/text/error/non-finite
  // seed is UNVERIFIED — the shipped cache is retained, never assumed 0.
  // -------------------------------------------------------------------------
  function iterationSeed(key: string): number | undefined {
    const snap = iterationSnapshot.get(key)
    const cell = cellMap.get(key)
    const value = snap ? snap.value : cell?.value
    const valueIsError = snap ? snap.valueIsError : cell?.valueIsError
    const hasCachedValue = snap ? snap.hasCachedValue : cell?.hasCachedValue
    if (valueIsError === true) return undefined
    if (typeof value !== 'number' || !Number.isFinite(value)) return undefined
    if (hasCachedValue === false) return undefined
    return value
  }

  function iterationSeedReason(key: string): string {
    const snap = iterationSnapshot.get(key)
    const cell = cellMap.get(key)
    const value = snap ? snap.value : cell?.value
    const valueIsError = snap ? snap.valueIsError : cell?.valueIsError
    const hasCachedValue = snap ? snap.hasCachedValue : cell?.hasCachedValue
    if (valueIsError === true || (typeof value === 'string' && isFormulaError(value))) return 'an error cache'
    if (value === null || value === undefined || hasCachedValue === false) return 'a missing cache'
    if (typeof value !== 'number') return `a ${typeof value} cache`
    if (!Number.isFinite(value)) return 'a non-finite numeric cache'
    return 'no trusted seed'
  }

  function reportIterationGate(feature: string, message: string): void {
    const holder = doc as unknown as { diagnostics?: Array<{ kind: string; feature: string; message: string }> }
    holder.diagnostics ??= []
    if (!holder.diagnostics.some((d) => d.feature === feature && d.message === message)) {
      holder.diagnostics.push({ kind: 'native-gate', feature, message })
    }
  }

  /** The memo value a retained/abandoned frame exposes to readers: the shipped
   * cache with the established legacy tri-state error identity. */
  function iterationSnapshotMemo(key: string): EvaluationValue {
    const snap = iterationSnapshot.get(key)
    if (!snap) return null
    return isFormulaError(snap.value) && snap.valueIsError !== false ? formulaError(snap.value) : snap.value
  }

  function restoreIterationCache(key: string): void {
    const snap = iterationSnapshot.get(key)
    const cell = cellMap.get(key)
    if (!snap || !cell) return
    cell.value = snap.value
    cell.valueIsError = snap.valueIsError
    cell.hasCachedValue = snap.hasCachedValue
    // Keep reads consistent for any cohort cell that still references it without
    // scheduling the retained cell for re-evaluation.
    memo.set(key, iterationSnapshotMemo(key))
  }

  function iterativeChange(before: EvaluationValue, after: EvaluationValue): number {
    if (typeof before === 'number' && typeof after === 'number' && Number.isFinite(before) && Number.isFinite(after)) {
      return Math.abs(after - before)
    }
    if (isEvaluationError(before) && isEvaluationError(after)) return before.code === after.code ? 0 : Number.POSITIVE_INFINITY
    return Object.is(before, after) ? 0 : Number.POSITIVE_INFINITY
  }

  function sameDependencySet(a: ReadonlySet<string>, b: ReadonlySet<string>): boolean {
    if (a.size !== b.size) return false
    for (const value of a) if (!b.has(value)) return false
    return true
  }

  function runIterativeCohort(): void {
    const formulaKeySet = new Set(formulas.keys())
    const depsOf = (node: string): Iterable<string> => {
      const set = generationDeps.get(node)
      if (!set) return []
      const filtered: string[] = []
      for (const dep of set) if (formulaKeySet.has(dep)) filtered.push(dep)
      return filtered
    }

    const { cyclic } = findCyclicCells([...formulaKeySet], depsOf)
    if (cyclic.size === 0) return

    // A trusted seed is a finite numeric cache. Everything else (missing,
    // Boolean, text, error, non-finite) is UNVERIFIED and poisons its closure.
    const untrusted = new Set<string>()
    for (const key of cyclic) if (iterationSeed(key) === undefined) untrusted.add(key)
    for (;;) {
      const reachable = transitiveDependents([...cyclic].filter((key) => !untrusted.has(key)), depsOf, formulaKeySet)
      const poisonedNow = transitiveDependents(untrusted, depsOf, formulaKeySet)
      let added = false
      for (const key of reachable) {
        if (poisonedNow.has(key) || iterationSeed(key) !== undefined) continue
        untrusted.add(key)
        added = true
      }
      if (!added) break
    }
    const poisoned = transitiveDependents(untrusted, depsOf, formulaKeySet)
    const iterableCyclic = [...cyclic].filter((key) => !untrusted.has(key))
    const cohortAll = transitiveDependents(iterableCyclic, depsOf, formulaKeySet)

    const positionOf = (key: string): readonly [number, number, number] => {
      const cell = cellMap.get(key)!
      const sheetName = cellSheets.get(key)!
      const sheetIndex = identitySheets.find((s) => s.name.toLowerCase() === sheetName.toLowerCase())?.workbookIndex ?? 0
      return [sheetIndex, cell.row, cell.col]
    }

    const cohortCells = [...cohortAll].filter((key) => !poisoned.has(key))
    const cohortCellSet = new Set(cohortCells)
    const cohortSheets = new Set(cohortCells.map((key) => positionOf(key)[0]))

    const gates: Array<{ feature: string; message: string }> = []
    if (cohortSheets.size > 1) {
      gates.push({
        feature: 'calc-iterative-order-unverified',
        message: `Iterative recalculation spans ${cohortSheets.size} sheets; cross-sheet pass order is unverified, so cyclic caches are retained`,
      })
    }
    const count = policy.iterateCount
    const delta = policy.iterateDelta
    // The OOXML/app contract is a positive integer pass count. A non-integer is
    // out of contract — never silently rounded up.
    if (!Number.isSafeInteger(count) || count <= 0 || count > ITERATIVE_PASS_HARD_CAP) {
      gates.push({
        feature: 'calc-iterative-limit-unverified',
        message: `Iterative recalculation with iterateCount=${count} is unverified (supported: a positive safe integer in 1..${ITERATIVE_PASS_HARD_CAP}); cyclic caches are retained`,
      })
    }
    if (!Number.isFinite(delta) || delta < 0) {
      gates.push({
        feature: 'calc-iterative-limit-unverified',
        message: `Iterative recalculation with iterateDelta=${delta} is unverified; cyclic caches are retained`,
      })
    }
    for (const key of cohortCells) {
      const cell = cellMap.get(key)!
      const sheetName = cellSheets.get(key)!
      if (cell.arrayRef || engine.hasOwner(addressOfName(sheetName, cell.col, cell.row))) {
        gates.push({
          feature: 'calc-iterative-spill-unverified',
          message: `Iterative recalculation of the array/spill cohort at ${cell.ref || key} is unverified; cyclic caches are retained`,
        })
        break
      }
    }
    for (const key of cohortCells) {
      if (!volatileKeys.has(key)) continue
      gates.push({
        feature: 'calc-iterative-volatile-unverified',
        message: `Iterative recalculation of the volatile cohort at ${cellMap.get(key)?.ref || key} is unverified; cyclic caches are retained`,
      })
      break
    }
    if (doc.calcChainPresent === true) {
      gates.push({
        feature: 'calc-iterative-chain-unverified',
        message: 'The workbook actually ships a calcChain part; its pass-order provenance is unverified, so cyclic caches are retained',
      })
    }

    if (gates.length > 0) {
      for (const gate of gates) reportIterationGate(gate.feature, gate.message)
      for (const key of cohortAll) restoreIterationCache(key)
      for (const key of poisoned) restoreIterationCache(key)
      return
    }

    for (const key of untrusted) {
      const cell = cellMap.get(key)
      reportIterationGate(
        'calc-iterative-seed-unverified',
        `Iterative recalculation unavailable at ${cell?.ref || key}: ${iterationSeedReason(key)} (no trusted finite numeric seed); valid cache retained`,
      )
    }
    for (const key of poisoned) {
      if (untrusted.has(key)) continue
      const cell = cellMap.get(key)
      reportIterationGate(
        'calc-iterative-seed-unverified',
        `Iterative recalculation unavailable at ${cell?.ref || key}: downstream of a cohort with no trusted seed; valid cache retained`,
      )
    }

    // Retain every cell that will not be iterated exactly as the file shipped it.
    for (const key of cohortAll) if (!cohortCellSet.has(key)) restoreIterationCache(key)
    for (const key of poisoned) if (!cohortCellSet.has(key)) restoreIterationCache(key)

    const order = canonicalOrder(cohortCells, positionOf)
    // Discovery edges are the baseline for adaptive/conditional dependency-change
    // detection. C4 r2: the guard compares the ACTUAL read sets, plain INPUT
    // edges included — a branch that starts reading an ordinary input cell
    // (never a formula node) is a real dependency change. The SCC graph itself
    // stays restricted to formula nodes (see `depsOf`).
    const lastDeps = new Map<string, Set<string>>()
    for (const key of order) {
      const deps = generationDeps.get(key)
      lastDeps.set(key, new Set(deps ? [...deps] : []))
    }

    iterating = true
    for (const key of order) {
      const seed = iterationSeed(key)!
      iterValues.set(key, seed)
      memo.set(key, seed)
    }
    let dependencyChanged = false
    runBoundedPasses(count, delta, () => {
      let maxChange = 0
      for (const key of order) {
        passDeps.clear()
        const before = iterValues.get(key)!
        evaluateCell(key)
        const after = memo.has(key) ? memo.get(key)! : iterValues.get(key)!
        iterValues.set(key, after)
        const current = new Set<string>(passDeps.get(key) ?? [])
        const previous = lastDeps.get(key)
        if (previous && !sameDependencySet(previous, current)) {
          dependencyChanged = true
          return maxChange
        }
        lastDeps.set(key, current)
        const change = iterativeChange(before, after)
        if (change > maxChange) maxChange = change
      }
      return maxChange
    }, () => dependencyChanged)
    iterating = false
    iterValues.clear()

    if (dependencyChanged) {
      reportIterationGate(
        'calc-iterative-dependency-unverified',
        'Iterative recalculation unavailable: the actual dependency edges of the cyclic cohort changed between passes (adaptive/conditional dependency); cyclic caches retained',
      )
      for (const key of cohortAll) restoreIterationCache(key)
      for (const key of poisoned) restoreIterationCache(key)
      return
    }
  }

  for (const key of formulas.keys()) evaluateCell(key)

  if (policy.iterate && formulas.size > 0) runIterativeCohort()

  docSpillStates.set(doc, manualFileLoadNoOp && prior
    ? prior
    : { generated: [...retainedGenerated, ...generatedList], createdRows: onlyOwners && prior ? [...prior.createdRows.filter(row => !onlyOwners.has(row.ownerKey)), ...createdRows] : createdRows,
      valueDeps, formulaDeps, reverseFormulaDeps: reverseDeps, retainedUnsupportedKeys, runtimeIndex: { sheetsByName, cellMap, cellSheets, sheetMaxCols, sheetMaxRows, inputKeys, formulaKeys, mergeKeys, identitySheets, revision: cellModelRevision(doc) } })
  // A full pass reads every cell afresh: it is the resynchronisation point for direct edits nobody announced, so the
  // syntactic reader index is rebuilt from the cells at the next dependent recompute.
  if (!onlyKeys) formulaReferenceIndexes.delete(doc)
}
