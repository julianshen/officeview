import { evaluateFormulaInternal, makeCellKey, formulaError, isFormulaError, isEvaluationError, isMatrixValue, publicFormulaValue, round15, scalarProjection, topLeftOfMatrix } from './evaluator'
import { parseFormula } from './parser'
import { translateSharedFormula } from './shared'
import type { AstNode, CellAddress, ElementValue, EvaluationContext, EvaluationValue, ModeFrameStore, ModeFrameStores, RefArea, ResolvedSemantics, SheetId } from './types'
import { createReferenceServices, gridAreas } from './refs'
import { SpillEngine, SpillSuspension, spillCellKey, type SpillBlockReason, type SpillIssue } from './spills'
import { resolveUnicodeSettings, unicodeVersionOf } from './unicode'
import { decideCellCalculation, isWhatIfDataTable, type CalcMode, type CalcPolicy } from './calc-policy'
import { canonicalOrder, findCyclicCells, ITERATIVE_PASS_HARD_CAP, runBoundedPasses, transitiveDependents } from './iterative'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../types'
import { savedSpillFollowers } from './saved-spills'

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
  /** Explicit locale override (wins over parsed file metadata). */
  locale?: string
  /** Explicit IANA timeZone override. */
  timeZone?: string
  /** Injected clock sample (ms). Sampled lazily once per generation if absent. */
  now?: () => number
  /** Explicit date-system override (wins over the parsed workbookPr date1904). */
  date1904?: boolean
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
interface CreatedRowRecord { sheet: XlsxSheet; index: number }
interface DocSpillState {
  generated: GeneratedCellRecord[]
  createdRows: CreatedRowRecord[]
  valueDeps: Map<string, Set<string>>
}
const docSpillStates = new WeakMap<XlsxDocument, DocSpillState>()

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
  // Cheap early-exit: skip evaluation work entirely when no cell has a formula.
  // A what-if `<f t="dataTable"/>` may carry no formula text at all; it is
  // still unsupported and must be diagnosed, so the actual what-if type keeps
  // the scan from exiting before the diagnosis branch. Ordinary formula-less,
  // type-less content (the common case) still exits cheaply here.
  let hasFormulas = false
  for (const sheet of doc.sheets) {
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
  const removedKeys = new Set<string>()
  if (prior && !manualFileLoadNoOp) {
    for (const g of prior.generated) {
      const rowObj = g.sheet.rows.find((r) => r.index === g.cell.row)
      if (!rowObj || !rowObj.cells.includes(g.cell)) continue // replaced object → user input
      if (!cellContentMatches(g.cell, g.written)) continue // user content edit → user input
      if (g.original) restoreCell(g.cell, g.original)
      else removeCell(g.sheet, g.cell)
      removedKeys.add(g.key)
    }
    for (const row of prior.createdRows) removeCreatedRow(row.sheet, row.index)
  }
  const valueDeps = prior?.valueDeps ?? new Map<string, Set<string>>()
  /** ACTUAL (lazily discovered) dependency edges for this generation, kept
   * separate from `valueDeps` so established spill-cleanup semantics are
   * unchanged. Populated only by real reads — dead lazy branches add nothing. */
  const generationDeps = new Map<string, Set<string>>()
  const forceKeys = prior && removedKeys.size > 0 ? affectedReaders(removedKeys, prior.valueDeps) : new Set<string>()

  const sheetsByName = new Map<string, XlsxSheet>()
  const cellMap = new Map<string, XlsxCell>()
  const cellSheets = new Map<string, string>()
  const sheetMaxCols = new Map<string, number>()
  const sheetMaxRows = new Map<string, number>()

  // 1. Resolve shared formulas on all sheets and index sheets, cells, and bounds
  for (const sheet of doc.sheets) {
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
  const inputKeys = new Set<string>()
  const formulaKeys = new Set<string>()
  const mergeKeys = new Set<string>()
  for (const sheet of doc.sheets) {
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
  const identitySheets = (doc.workbookSheets ?? doc.sheets.map((s, i) => ({
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
    return formulaKeys.has(key) || inputKeys.has(key)
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
  }

  function writeGenerated(address: CellAddress, value: ElementValue, valueIsError: boolean, ownerKey: string): void {
    const name = sheetNameOfId(address.sheetId)
    if (name === undefined) return
    const sheet = sheetsByName.get(name.toLowerCase())
    if (!sheet) return
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
    if (!existingRow) createdRows.push({ sheet, index: address.row })
    const key = makeCellKey(name, address.col, address.row)
    cellMap.set(key, cell)
    cellSheets.set(key, name)
    memo.set(key, normalized)
    generatedList.push({ key, sheet, cell, ownerKey, written: snapshotCell(cell), ...(original ? { original } : {}) })
  }

  /** Revert one generated cell: restore an original cell's snapshot, or remove
   * a cell we created. Never destroys original formatting/metadata. */
  function revertGenerated(g: GeneratedCellRecord): void {
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
      forcedDependent: forceKeys.has(key),
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
    maxArrayCells: options?.maxArrayCells,
    maxArraySide: options?.maxArraySide,
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
        for (const [key, cell] of cellMap) {
          const name = cellSheets.get(key)
          if (name === undefined) continue
          const sheetId = sheetIdOfName(name)
          if (sheetId === undefined) continue
          const isFormulaCell = cell.formula !== undefined || cell.sharedFormula !== undefined
          if (!isFormulaCell && (cell.value === null || cell.value === undefined)) continue
          let value: ElementValue = cell.value
          if (!isFormulaCell) {
            const error = cell.valueIsError ?? isFormulaError(cell.value)
            value = error && isFormulaError(cell.value) ? formulaError(cell.value) : cell.value
          }
          yield {
            address: { sheetId, col: cell.col, row: cell.row },
            value,
            origin: isFormulaCell ? 'formula' : 'input',
          }
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
    : { generated: generatedList, createdRows, valueDeps })
}
