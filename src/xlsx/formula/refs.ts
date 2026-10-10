/**
 * B1 reference geometry, sparse cursors and workbook reference services.
 *
 * Pure geometry (areas, overlaps, scalar projection) plus a dependency-
 * injected service factory: traversal reads through a caller-supplied cell
 * reader, so workbook memoization/cycle/suspension semantics are preserved
 * and structural controls can run against in-memory fakes. No evaluation of
 * later families lives here; aggregate/error semantics stay in functions.ts.
 */
import type {
  AstNode,
  CellAddress,
  DefinedNameMetadata,
  ElementValue,
  EvaluationContext,
  EvaluationError,
  EvaluationValue,
  GridReference,
  NameBinding,
  RefArea,
  RefCursor,
  RefEntry,
  ReferenceNode,
  ReferenceServices,
  ResolvedRef,
  SheetId,
  TableMetadata,
  TableSyntax,
} from './types'
import { parseFormula } from './parser'
import { translateAst } from './shared'
import { isBindingActive } from './name-state'
import { contains3d, relativeNodeGate } from './ref-discovery'
import { referenceGenerationChanged, registerReferenceGeneration, registerReferenceCreator } from './reference-generation'

/** Single stored cell observed by a traversal (absent cells are unlisted). */
export interface StoredCell {
  address: CellAddress
  value: ElementValue
  origin: 'input' | 'formula' | 'spill'
}

export interface SparseCellStore {
  /** Read one cell; undefined means truly absent. May throw suspensions. */
  read(address: CellAddress): { value: ElementValue; origin: 'input' | 'formula' | 'spill' } | undefined
  /** All stored cells; callers filter per area (never enumerating blanks). */
  stored(): Iterable<StoredCell>
}

export interface ReferenceServiceDeps {
  /** Monotonic recalculation generation; cursors bind to it. */
  generation: () => number
  /** Original-order sheet identities for runs and local-name scopes. */
  sheets: readonly { sheetId: SheetId; name: string; workbookIndex: number }[]
  /** Case-insensitive sheet lookup by name. */
  sheetIdOfName(name: string): SheetId | undefined
  sheetNameOfId(id: SheetId): string | undefined
  store: SparseCellStore
  definedNames: readonly DefinedNameMetadata[]
  /** Undefined table inventory means unavailable, never known-empty. */
  tables: readonly TableMetadata[] | undefined
  onIssue?: (feature: string, message: string) => void
}

export function gridAreas(grid: GridReference, sheetId: SheetId): RefArea[] {
  switch (grid.type) {
    case 'cell':
      return [{ sheetId, firstCol: grid.ref.col, firstRow: grid.ref.row, cols: 1, rows: 1 }]
    case 'range': {
      const firstCol = Math.min(grid.ref.from.col, grid.ref.to.col)
      const firstRow = Math.min(grid.ref.from.row, grid.ref.to.row)
      return [{
        sheetId,
        firstCol,
        firstRow,
        cols: Math.max(grid.ref.from.col, grid.ref.to.col) - firstCol + 1,
        rows: Math.max(grid.ref.from.row, grid.ref.to.row) - firstRow + 1,
      }]
    }
    case 'wholeCol': {
      const firstCol = Math.min(grid.ref.from.index, grid.ref.to.index)
      return [{
        sheetId,
        firstCol,
        firstRow: 0,
        cols: Math.max(grid.ref.from.index, grid.ref.to.index) - firstCol + 1,
        rows: 1048576,
      }]
    }
    case 'wholeRow': {
      const firstRow = Math.min(grid.ref.from.index, grid.ref.to.index)
      return [{
        sheetId,
        firstCol: 0,
        firstRow,
        cols: 16384,
        rows: Math.max(grid.ref.from.index, grid.ref.to.index) - firstRow + 1,
      }]
    }
  }
}

function overlap(a: RefArea, b: RefArea): RefArea | null {
  if (a.sheetId !== b.sheetId) return null
  if (a.cols <= 0 || a.rows <= 0 || b.cols <= 0 || b.rows <= 0) return null
  const firstCol = Math.max(a.firstCol, b.firstCol)
  const firstRow = Math.max(a.firstRow, b.firstRow)
  const lastCol = Math.min(a.firstCol + a.cols - 1, b.firstCol + b.cols - 1)
  const lastRow = Math.min(a.firstRow + a.rows - 1, b.firstRow + b.rows - 1)
  if (lastCol < firstCol || lastRow < firstRow) return null
  return { sheetId: a.sheetId, firstCol, firstRow, cols: lastCol - firstCol + 1, rows: lastRow - firstRow + 1 }
}

/** Rectangular pairwise overlaps in source order; empty means no overlap. */
export function intersectAreas(left: readonly RefArea[], right: readonly RefArea[]): RefArea[] {
  const out: RefArea[] = []
  for (const a of left) {
    for (const b of right) {
      const hit = overlap(a, b)
      if (hit) out.push(hit)
    }
  }
  return out
}

export function areaCells(area: RefArea): number {
  return area.cols * area.rows
}

function inArea(area: RefArea, col: number, row: number): boolean {
  return (
    col >= area.firstCol &&
    col < area.firstCol + area.cols &&
    row >= area.firstRow &&
    row < area.firstRow + area.rows
  )
}

export function isTaggedError(value: ElementValue): value is EvaluationError {
  return typeof value === 'object' && value !== null && (value as EvaluationError).kind === 'formula-error'
}

export function errorOf(code: EvaluationError['code']): EvaluationError {
  return { kind: 'formula-error', code }
}

interface CursorState {
  id: string
  mode: 'populated' | 'all'
  generation: number
  areaIdx: number
  /** Populated mode: ordinal within the area's stored list. */
  ordinal: number
  /** All mode: offsets within the current area. */
  offsetRow: number
  offsetCol: number
}

interface PopulatedSlot { col: number; row: number }

class ResolvedRefImpl implements ResolvedRef {
  readonly kind = 'resolved-reference' as const
  readonly areas: readonly RefArea[]
  readonly generation: number
  private store: SparseCellStore
  private currentGeneration: () => number
  private populatedByArea: Array<PopulatedSlot[] | null> = []
  private cursorSeq = 0
  private cursors = new Map<string, CursorState>()

  constructor(
    areas: readonly RefArea[],
    store: SparseCellStore,
    generation: number,
    currentGeneration: () => number,
  ) {
    this.areas = areas
    this.store = store
    this.generation = generation
    this.currentGeneration = currentGeneration
    this.populatedByArea = areas.map(() => null)
    registerReferenceCreator(this, region => new ResolvedRefImpl(region, store, generation, currentGeneration))
  }

  private checkGeneration(kind: 'reference' | 'cursor' = 'reference'): void {
    const current = this.currentGeneration()
    if (this.generation !== current) throw referenceGenerationChanged(this.generation, current, kind)
  }

  private checkLive(cursor: RefCursor): CursorState {
    const st = this.cursors.get(cursor.id)
    if (!st || st.generation !== this.generation) {
      throw new Error(`stale cursor generation (expected ${this.generation})`)
    }
    this.checkGeneration('cursor')
    if (st.mode !== cursor.mode) throw new Error('cursor mode mismatch')
    return st
  }

  private populated(areaIdx: number): PopulatedSlot[] {
    this.checkGeneration()
    let list = this.populatedByArea[areaIdx]
    if (!list) {
      const area = this.areas[areaIdx]
      list = []
      for (const cell of this.store.stored()) {
        if (cell.address.sheetId === area.sheetId && inArea(area, cell.address.col, cell.address.row)) {
          list.push({ col: cell.address.col, row: cell.address.row })
        }
      }
      this.checkGeneration()
      // Row-major source order within the area; duplicates across union
      // areas stay separate occurrences by construction.
      list.sort((a, b) => a.row - b.row || a.col - b.col)
      this.populatedByArea[areaIdx] = list
    }
    return list
  }

  readAt(areaIndex: number, offsetRow: number, offsetCol: number): ElementValue {
    // Group 11 fix: positional reads are generation-guarded like cursors.
    this.checkGeneration()
    // Positional API inputs must be finite integer indices before lookup/read.
    // Use the existing domain-failure convention, not a native-offset claim.
    if (!Number.isInteger(areaIndex) || !Number.isInteger(offsetRow) || !Number.isInteger(offsetCol)) return errorOf('#REF!')
    const area = this.areas[areaIndex]
    if (!area) return errorOf('#REF!')
    if (offsetRow < 0 || offsetCol < 0 || offsetRow >= area.rows || offsetCol >= area.cols) {
      return errorOf('#REF!')
    }
    const hit = this.store.read({
      sheetId: area.sheetId,
      col: area.firstCol + offsetCol,
      row: area.firstRow + offsetRow,
    })
    this.checkGeneration()
    if (!hit) return null
    return hit.value
  }

  openCursor(mode: 'populated' | 'all'): RefCursor {
    const id = `c${++this.cursorSeq}`
    const cursor: RefCursor = { id, mode, generation: this.generation }
    this.cursors.set(id, { id, mode, generation: this.generation, areaIdx: 0, ordinal: 0, offsetRow: 0, offsetCol: 0 })
    return cursor
  }

  peek(cursor: RefCursor): RefEntry | undefined {
    const st = this.checkLive(cursor)
    if (st.mode === 'populated') {
      let areaIdx = st.areaIdx
      let ordinal = st.ordinal
      while (areaIdx < this.areas.length) {
        const list = this.populated(areaIdx)
        if (ordinal < list.length) {
          const slot = list[ordinal]
          const area = this.areas[areaIdx]
          // Group 1 fix: persist the EXACT peeked position before the fresh
          // read. A suspension during the read leaves the cursor pointing AT
          // this entry (retry peeks the same entry: no skip/duplicate), and a
          // successful advance() then moves past the exact committed address.
          st.areaIdx = areaIdx
          st.ordinal = ordinal
          const hit = this.store.read({ sheetId: area.sheetId, col: slot.col, row: slot.row })
          this.checkGeneration('cursor')
          return {
            areaIndex: areaIdx,
            offsetRow: slot.row - area.firstRow,
            offsetCol: slot.col - area.firstCol,
            address: { sheetId: area.sheetId, col: slot.col, row: slot.row },
            value: hit ? hit.value : null,
            origin: hit ? hit.origin : 'absent',
          }
        }
        areaIdx++
        ordinal = 0
      }
      return undefined
    }
    let areaIdx = st.areaIdx
    let offsetRow = st.offsetRow
    let offsetCol = st.offsetCol
    while (areaIdx < this.areas.length) {
      const area = this.areas[areaIdx]
      if (offsetRow < area.rows && offsetCol < area.cols) {
        const col = area.firstCol + offsetCol
        const row = area.firstRow + offsetRow
        // Group 1 fix: persist the exact peeked position (all mode).
        st.areaIdx = areaIdx
        st.offsetRow = offsetRow
        st.offsetCol = offsetCol
        const hit = this.store.read({ sheetId: area.sheetId, col, row })
        this.checkGeneration('cursor')
        return {
          areaIndex: areaIdx,
          offsetRow,
          offsetCol,
          address: { sheetId: area.sheetId, col, row },
          value: hit ? hit.value : null,
          origin: hit ? hit.origin : 'absent',
        }
      }
      areaIdx++
      offsetRow = 0
      offsetCol = 0
    }
    return undefined
  }

  advance(cursor: RefCursor): void {
    const st = this.checkLive(cursor)
    if (st.mode === 'populated') {
      st.ordinal++
      while (st.areaIdx < this.areas.length && st.ordinal >= this.populated(st.areaIdx).length) {
        st.areaIdx++
        st.ordinal = 0
      }
      return
    }
    if (st.areaIdx >= this.areas.length) return
    const area = this.areas[st.areaIdx]
    if (area.rows > 0 && area.cols > 0) {
      st.offsetCol++
      if (st.offsetCol < area.cols) return
      st.offsetCol = 0
      st.offsetRow++
      if (st.offsetRow < area.rows) return
    }
    // Consume only the committed position in the current area. A later real
    // area begins at offset zero; skipping empties never consumes that cell.
    st.areaIdx++
    st.offsetRow = 0
    st.offsetCol = 0
    while (st.areaIdx < this.areas.length && (this.areas[st.areaIdx].rows === 0 || this.areas[st.areaIdx].cols === 0)) st.areaIdx++
  }

  countAbsent(): number {
    // Group 11 fix: absent counts are generation-guarded like cursors.
    this.checkGeneration()
    // Structural arithmetic only: never enumerates blank cells.
    let total = 0
    let populated = 0
    for (let i = 0; i < this.areas.length; i++) {
      total += areaCells(this.areas[i])
      populated += this.populated(i).length
    }
    return total - populated
  }
}

function tag(code: EvaluationError['code']): EvaluationError {
  return errorOf(code)
}

/** Construct a resolved reference over caller-supplied geometry + store.
 * Used by the workbook spill engine for A1# owner rectangles; the cursor/
 * generation machinery is shared with every other resolved reference. */
export function createResolvedRef(
  areas: readonly RefArea[],
  store: SparseCellStore,
  generation: number,
  currentGeneration: () => number,
): ResolvedRef {
  return new ResolvedRefImpl(areas, store, generation, currentGeneration)
}

/** Scalar projection of resolved areas (implicit intersection semantics). */
export function projectScalar(
  areas: readonly RefArea[],
  read: (sheetId: SheetId, col: number, row: number) => ElementValue,
  cell: { col: number; row: number } | undefined,
): ElementValue {
  let cells = 0
  for (const area of areas) cells += areaCells(area)
  if (areas.length === 1 && cells === 1) {
    const area = areas[0]
    return read(area.sheetId, area.firstCol, area.firstRow)
  }
  // A diagnosed API-context #VALUE!: no invented A1 context when the
  // selection needs cell context that was never supplied.
  if (!cell) return tag('#VALUE!')
  if (areas.length === 1) {
    const area = areas[0]
    if (area.rows === 1 && area.cols > 1) {
      // Horizontal vector: the current column selects the cell within it.
      if (cell.col >= area.firstCol && cell.col < area.firstCol + area.cols) {
        return read(area.sheetId, cell.col, area.firstRow)
      }
      return tag('#VALUE!')
    }
    if (area.cols === 1 && area.rows > 1) {
      // Vertical vector: the current row selects the cell within it.
      if (cell.row >= area.firstRow && cell.row < area.firstRow + area.rows) {
        return read(area.sheetId, area.firstCol, cell.row)
      }
      return tag('#VALUE!')
    }
  }
  return tag('#VALUE!')
}

function gridHasRelative(grid: GridReference): boolean {
  switch (grid.type) {
    case 'cell':
      return !grid.ref.absCol || !grid.ref.absRow
    case 'range':
      return (
        !grid.ref.from.absCol ||
        !grid.ref.from.absRow ||
        !grid.ref.to.absCol ||
        !grid.ref.to.absRow
      )
    case 'wholeCol':
      return !grid.ref.from.absolute || !grid.ref.to.absolute
    case 'wholeRow':
      return !grid.ref.from.absolute || !grid.ref.to.absolute
  }
}

function hasRelativeEndpoints(node: ReferenceNode): boolean {
  switch (node.type) {
    case 'cell':
    case 'range':
    case 'wholeCol':
    case 'wholeRow':
      return gridHasRelative(node)
    case 'union':
      return node.refs.some(hasRelativeEndpoints)
    case 'intersect':
      return hasRelativeEndpoints(node.left) || hasRelativeEndpoints(node.right)
    case 'ref3d':
      return gridHasRelative(node.target)
    case 'spill':
      return !node.anchor.absCol || !node.anchor.absRow
    case 'name':
    case 'table':
      return false
  }
}

function gridSheet(grid: GridReference): string | undefined {
  switch (grid.type) {
    case 'cell':
      return grid.ref.sheet
    case 'range':
      return grid.ref.sheet
    case 'wholeCol':
      return grid.ref.sheet
    case 'wholeRow':
      return grid.ref.sheet
  }
}


function findTable(
  tables: readonly TableMetadata[] | undefined,
  name: string,
): TableMetadata | undefined {
  if (!tables) return undefined
  const upper = name.toUpperCase()
  return tables.find((t) => t.name.toUpperCase() === upper || t.displayName.toUpperCase() === upper)
}

function resolveTableAreas(
  syntax: TableSyntax,
  tables: readonly TableMetadata[] | undefined,
  ctx: EvaluationContext,
  onIssue: ((feature: string, message: string) => void) | undefined,
  useSiteSheetId: SheetId | undefined,
): RefArea[] | EvaluationError {
  // Unavailable inventory gates with a diagnostic (caches preserved) and is
  // distinct from a computed unknown-table error in complete metadata.
  if (tables === undefined) {
    markFeature(ctx, 'table-unavailable')
    onIssue?.(
      'table-unavailable',
      'Table inventory unavailable; structured references not resolved against a complete registry',
    )
    return tag('#NAME?')
  }
  let table: TableMetadata | undefined
  if (syntax.table !== undefined) {
    table = findTable(tables, syntax.table)
    if (!table) return tag('#NAME?')
  } else {
    const contextId = ctx.tableContext?.tableId
    table = contextId !== undefined ? tables.find((t) => t.id === contextId) : undefined
    if (!table) return tag('#REF!')
  }
  const full = table.extent
  const dataRows = Math.max(0, full.rows - table.headerRowCount - table.totalsRowCount)
  const dataFirstRow = full.firstRow + table.headerRowCount
  const colSpan = (name: string): { firstCol: number; cols: number } | undefined => {
    const idx = table.columns.findIndex((c) => c.name.toUpperCase() === name.toUpperCase())
    if (idx < 0) return undefined
    return { firstCol: full.firstCol + idx, cols: 1 }
  }

  const sections: RefArea[] = []
  const items = syntax.items.length > 0 ? syntax.items : (['data'] as const)
  for (const item of items) {
    switch (item) {
      case 'all':
        sections.push({ ...full })
        break
      case 'data':
        // Empty table-data references resolve to explicit zero-sized areas.
        sections.push({
          sheetId: full.sheetId,
          firstCol: full.firstCol,
          firstRow: dataFirstRow,
          cols: full.cols,
          rows: dataRows,
        })
        break
      case 'headers':
        if (table.headerRowCount === 0) return tag('#REF!')
        sections.push({
          sheetId: full.sheetId,
          firstCol: full.firstCol,
          firstRow: full.firstRow,
          cols: full.cols,
          rows: table.headerRowCount,
        })
        break
      case 'totals':
        if (table.totalsRowCount === 0) return tag('#REF!')
        sections.push({
          sheetId: full.sheetId,
          firstCol: full.firstCol,
          firstRow: full.firstRow + full.rows - table.totalsRowCount,
          cols: full.cols,
          rows: table.totalsRowCount,
        })
        break
      case 'thisRow': {
        const current = ctx.currentAddress ?? ctx.tableContext?.currentCell ?? ctx.currentCell
        if (!current || useSiteSheetId !== table.sheetId || current.row < dataFirstRow || current.row >= dataFirstRow + dataRows) {
          markFeature(ctx, 'table-this-row-placement')
          onIssue?.('table-this-row-placement', 'This-row reference requires a use-site on the owning table sheet and inside its data rows; native invalid-placement result unmeasured')
          return tag('#NAME?')
        }
        sections.push({
          sheetId: full.sheetId,
          firstCol: full.firstCol,
          firstRow: current.row,
          cols: full.cols,
          rows: 1,
        })
        break
      }
    }
  }

  const out: RefArea[] = []
  for (const section of sections) {
    switch (syntax.columns.kind) {
      case 'all':
        out.push(section)
        break
      case 'single': {
        const span = colSpan(syntax.columns.name)
        if (!span) return tag('#REF!')
        out.push({
          sheetId: section.sheetId,
          firstCol: span.firstCol,
          firstRow: section.firstRow,
          cols: span.cols,
          rows: section.rows,
        })
        break
      }
      case 'range': {
        const from = colSpan(syntax.columns.from)
        const to = colSpan(syntax.columns.to)
        if (!from || !to) return tag('#REF!')
        const firstCol = Math.min(from.firstCol, to.firstCol)
        const lastCol = Math.max(from.firstCol + from.cols - 1, to.firstCol + to.cols - 1)
        out.push({
          sheetId: section.sheetId,
          firstCol,
          firstRow: section.firstRow,
          cols: lastCol - firstCol + 1,
          rows: section.rows,
        })
        break
      }
    }
  }
  return out
}

/** Scalar projection of an already-resolved reference. */
export function scalarOfResolved(
  ref: ResolvedRef,
  cell: { col: number; row: number } | undefined,
  ctx?: EvaluationContext,
): ElementValue {
  const areas = ref.areas
  let cells = 0
  for (const area of areas) cells += areaCells(area)
  if (areas.length === 1 && cells === 1) return ref.readAt(0, 0, 0)
  // Group 2: missing currentCell = diagnosed API-context boundary (never an
  // invented A1 context); the capability is marked specifically.
  if (!cell) {
    markFeature(ctx, 'projection-context-missing')
    ctx?.reportFormulaIssue?.({
      kind: 'unsupported-reference',
      feature: 'projection-context-missing',
      message: 'Implicit projection needs current cell context that was not supplied',
    })
    return errorOf('#VALUE!')
  }
  if (areas.length === 1) {
    const area = areas[0]
    if (area.rows === 1 && area.cols > 1) {
      if (cell.col >= area.firstCol && cell.col < area.firstCol + area.cols) {
        return ref.readAt(0, 0, cell.col - area.firstCol)
      }
      return errorOf('#VALUE!') // documented out-of-range vector selection
    }
    if (area.cols === 1 && area.rows > 1) {
      if (cell.row >= area.firstRow && cell.row < area.firstRow + area.rows) {
        return ref.readAt(0, cell.row - area.firstRow, 0)
      }
      return errorOf('#VALUE!') // documented out-of-range vector selection
    }
  }
  // Group 2: ambiguous projection shape (2D rectangle or multi-area union):
  // capability gate with marked diagnostics and a frame-visible error.
  markFeature(ctx, 'projection-ambiguous')
  ctx?.reportFormulaIssue?.({
    kind: 'unsupported-reference',
    feature: 'projection-ambiguous',
    message: 'Implicit projection shape is ambiguous (2D area or multi-area); native result unmeasured',
  })
  return errorOf('#NAME?')
}

export function isReferenceNode(node: AstNode): node is ReferenceNode {
  switch (node.type) {
    case 'cell':
    case 'range':
    case 'wholeCol':
    case 'wholeRow':
    case 'name':
    case 'table':
    case 'ref3d':
    case 'union':
    case 'intersect':
    case 'spill':
      return true
    default:
      return false
  }
}

export type RefEvalNode = (node: AstNode, ctx?: EvaluationContext) => EvaluationValue

function markFeature(ctx: EvaluationContext | undefined, feature: string): void {
  if (!ctx) return
  const record = ctx as EvaluationContext & { unsupportedFeatures?: Set<string> }
  if (!record.unsupportedFeatures) record.unsupportedFeatures = new Set()
  record.unsupportedFeatures.add(feature)
}

/**
 * Scalar value of a reference node. With reference services the full
 * resolve/project path runs; standalone only syntactic geometry resolves
 * (names/tables/3D need a registry, spill needs a lifecycle).
 */
export function scalarOfReference(
  node: ReferenceNode,
  ctx: EvaluationContext | undefined,
  evalNode: RefEvalNode,
): EvaluationValue {
  const services = ctx?.references
  if (services && ctx) {
    const resolved = services.resolve(node, ctx)
    if (resolved === undefined) {
      // A name denoting a value, not a reference: evaluate its source.
      if (node.type !== 'name') return errorOf('#VALUE!')
      // Request the owning name node so the driver acquires its binding
      // identity. Evaluating the source directly bypasses that cycle scope.
      return evalNode(node, ctx)
    }
    if (isTaggedError(resolved as ElementValue)) return resolved as EvaluationError
    services.prepare(resolved as ResolvedRef, ctx)
    return scalarOfResolved(resolved as ResolvedRef, ctx.currentCell, ctx)
  }
  // Standalone path: syntactic projection through the legacy cell callback.
  // Registry operands (names/tables) stay unavailable; 3D/spill resolve to
  // typed reference errors without a sheet inventory or lifecycle.
  const readLegacy = (sheetId: SheetId, col: number, row: number): ElementValue => {
    const callback = ctx?.getCellValue
    if (!callback) return null
    const value = callback(sheetId === '' ? undefined : sheetId, col, row)
    if (value === undefined || value === null) return null
    if (!ctx?.typedValues && typeof value === 'string' && CANONICAL_ERROR_STRINGS.has(value)) {
      return errorOf(value as EvaluationError['code'])
    }
    return value as ElementValue
  }
  switch (node.type) {
    case 'cell':
    case 'range':
    case 'wholeCol':
    case 'wholeRow':
    case 'union':
    case 'intersect': {
      const areas = syntacticAreas(node, ctx?.currentSheet)
      if (isTaggedError(areas as ElementValue)) return areas as EvaluationError
      return projectScalar(areas as RefArea[], readLegacy, ctx?.currentCell)
    }
    case 'name':
      markFeature(ctx, typeof node.ref.name === 'string' ? node.ref.name.toUpperCase() : 'NAME')
      return errorOf('#NAME?')
    case 'table':
      markFeature(ctx, 'table-unavailable')
      ctx?.reportFormulaIssue?.({
        kind: 'unsupported-reference',
        feature: 'table-unavailable',
        message: 'Table inventory unavailable without reference services',
      })
      return errorOf('#NAME?')
    case 'ref3d':
      return errorOf('#REF!')
    case 'spill':
      // Spill ownership is a later lifecycle: gate with a diagnostic that
      // preserves valid caches instead of inventing a genuine error.
      markFeature(ctx, 'spill-unavailable')
      ctx?.reportFormulaIssue?.({
        kind: 'unsupported-reference',
        feature: 'spill-unavailable',
        message: 'Spill ownership is not implemented in B1',
      })
      return errorOf('#NAME?')
  }
}

/** Scalar value of an implicit-intersection (@) operand. */
export function scalarOfImplicitIntersect(
  expr: AstNode,
  ctx: EvaluationContext | undefined,
  evalNode: RefEvalNode,
): EvaluationValue {
  if (isReferenceNode(expr)) {
    const services = ctx?.references
    if (ctx && services && contains3d(expr, (name, useCtx) => services.bindName(name, useCtx), ctx)) {
      markFeature(ctx, 'implicit-intersection-3d')
      ctx.reportFormulaIssue?.({
        kind: 'unsupported-reference',
        feature: 'implicit-intersection-3d',
        message: 'Implicit intersection of a 3D sheet run is a prohibited capability',
      })
      return errorOf('#NAME?')
    }
    return scalarOfReference(expr, ctx, evalNode)
  }
  return evalNode(expr, ctx)
}

/**
 * Syntactic reference areas without a registry: grids compose by geometry
 * (unions concatenate, intersections overlap). Registry operands stay
 * unavailable errors; callers decide the error/value policy.
 */
export function syntacticAreas(
  node: AstNode,
  currentSheet: string | undefined,
): RefArea[] | EvaluationError {
  const sheetOf = (explicit: string | undefined): string => explicit ?? currentSheet ?? ''
  switch (node.type) {
    case 'cell':
    case 'range':
    case 'wholeCol':
    case 'wholeRow':
      return gridAreas(node, sheetOf(node.ref.sheet))
    case 'union': {
      const out: RefArea[] = []
      for (const part of node.refs) {
        const areas = syntacticAreas(part, currentSheet)
        if (isTaggedError(areas as ElementValue)) return areas as EvaluationError
        out.push(...(areas as RefArea[]))
      }
      return out
    }
    case 'intersect': {
      const left = syntacticAreas(node.left, currentSheet)
      if (isTaggedError(left as ElementValue)) return left as EvaluationError
      const right = syntacticAreas(node.right, currentSheet)
      if (isTaggedError(right as ElementValue)) return right as EvaluationError
      return intersectAreas(left as RefArea[], right as RefArea[])
    }
    default:
      return errorOf('#NAME?')
  }
}

const CANONICAL_ERROR_STRINGS = new Set<string>([
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#NULL!',
])

export function createReferenceServices(deps: ReferenceServiceDeps): ReferenceServices {
  const workbookIndexOf = (name: string): number | undefined => {
    const lower = name.toLowerCase()
    return deps.sheets.find((s) => s.name.toLowerCase() === lower)?.workbookIndex
  }

  // Group 12: physical sheet bounds domain check (0-based; 16384 cols x
  // 1048576 rows). Direct invalid coordinates are a contract-approved typed
  // #REF! — never a resolved out-of-range area.
  const areasInDomain = (areas: readonly RefArea[]): boolean =>
    areas.every(
      (a) =>
        a.firstCol >= 0 &&
        a.firstRow >= 0 &&
        a.firstCol + a.cols <= 16384 &&
        a.firstRow + a.rows <= 1048576,
    )

  const resolveGrid = (grid: GridReference, ctx: EvaluationContext): RefArea[] | EvaluationError => {
    const sheetName = gridSheet(grid) ?? ctx.currentSheet
    if (sheetName === undefined) return tag('#REF!')
    const sheetId = deps.sheetIdOfName(sheetName)
    if (sheetId === undefined) return tag('#REF!')
    const areas = gridAreas(grid, sheetId)
    if (!areasInDomain(areas)) return tag('#REF!')
    return areas
  }

  const resolveSpillAddress = (node: Extract<ReferenceNode, { type: 'spill' }>, ctxArg: EvaluationContext): CellAddress => {
    const sheetName = node.anchor.sheet ?? ctxArg.currentSheet
    const sheetId = sheetName === undefined ? '' : (deps.sheetIdOfName(sheetName) ?? '')
    return { sheetId, col: node.anchor.col, row: node.anchor.row }
  }

  const gate3dIntersection = (ctxArg: EvaluationContext): EvaluationError => {    markFeature(ctxArg, '3d-intersection')
    deps.onIssue?.(
      '3d-intersection',
      'Intersection operands include a 3D sheet run; prohibited capability gate (no native outcome invented)',
    )
    return tag('#NAME?')
  }

  const bindNameDirect = (name: { name: string; sheet?: string }, ctxArg: EvaluationContext): NameBinding | EvaluationError => {
    // (deferred to services.bindName via lazy property access at call time)
    return servicesRef.bindName(name, ctxArg)
  }
  const servicesRef: { bindName(name: { name: string; sheet?: string }, ctx: EvaluationContext): NameBinding | EvaluationError } = {
    bindName: undefined as never,
  }

  // Stable binding cache: key = lower name + scope sheetId/source scope +
  // generation + source text + base provenance/base coordinates. Hits reuse
  // the SAME AST object identity (frame scheduler memo stability).
  let bindingAstCache = new Map<string, { ast: NameBinding['ast'] }>()
  const bindingCacheKey = (
    d: DefinedNameMetadata, scopeSheetId: SheetId | undefined,
  ): string => {
    const base = d.relativeBase
    return [
      d.name.toLowerCase(),
      d.localSheetIndex === undefined ? 'global' : String(d.localSheetIndex),
      scopeSheetId ?? '',
      String(deps.generation()),
      d.source,
      d.baseProvenance,
      base ? `${base.sheetId}:${base.col}:${base.row}` : 'nobase',
    ].join('#')
  }

  // Positive VALUE classification follows the caller's existing memo lifetime:
  // an unchanged alias source may select a child whose metadata later changes.
  // Fresh memos must reclassify; supplied memo ownership stays untouched.
  // Source identities and sheet basis distinguish generation/base/local scope.
  const knownValueBindingsByMemo = new WeakMap<WeakMap<AstNode, EvaluationValue>, WeakMap<AstNode, Set<SheetId | undefined>>>()

  const services: ReferenceServices = {
    bindName(name, ctx): NameBinding | EvaluationError {
      // Group 3 minimal repair: binding ASTs are cached by the stable key
      // (lower-cased name + scope + source + generation + relative-base
      // identity). A hit reuses the SAME cached AST object (stable identity
      // for the frame scheduler's memo); generation/source changes refresh.
      // Group 12: an explicit qualifier naming an UNKNOWN sheet must not
      // silently resolve an unrelated global name. The exact native name-error
      // code is unmeasured: preserve an explicit capability diagnostic and let
      // the evaluator's valid-cache policy apply; the structural tag stays an
      // honest reference failure without claiming a native code.
      if (name.sheet !== undefined && workbookIndexOf(name.sheet) === undefined) {
        // Internal convention (not a fabricated native code claim): an
        // unavailable-capability gate marks the evaluated frame and returns
        // the internal #NAME? shape, exactly like every other unavailable
        // capability gate (spill/table/reversed-3D). The specific marker +
        // issue carry the diagnostic; the exact native name-error code stays
        // an unmeasured evidence question.
        markFeature(ctx, 'qualified-name-sheet-unavailable')
        deps.onIssue?.(
          'qualified-name-sheet-unavailable',
          `Qualified name references unknown sheet "${name.sheet}"`,
        )
        return tag('#NAME?')
      }
      const upper = name.name.toUpperCase()
      const candidates = deps.definedNames.filter((d) => d.name.toUpperCase() === upper)
      if (candidates.length === 0) return tag('#NAME?')
      // Explicit sheet qualification selects that local scope; otherwise the
      // use-site local scope shadows workbook-global names.
      const scopeIndex = name.sheet !== undefined
        ? workbookIndexOf(name.sheet)
        : ctx.currentSheet !== undefined
          ? workbookIndexOf(ctx.currentSheet)
          : undefined
      let picked: DefinedNameMetadata | undefined
      if (scopeIndex !== undefined) {
        picked = candidates.find((d) => d.localSheetIndex === scopeIndex)
      }
      picked ??= candidates.find((d) => d.localSheetIndex === undefined)
      if (!picked) return tag('#NAME?')
      const scopeSheetId = picked.localSheetIndex !== undefined
        ? deps.sheets.find((s) => s.workbookIndex === picked.localSheetIndex)?.sheetId
        : undefined
      let ast: NameBinding['ast']
      const cacheKey = bindingCacheKey(picked, scopeSheetId)
      const cachedAst = bindingAstCache.get(cacheKey)
      if (cachedAst) {
        ast = cachedAst.ast
      } else {
        try {
          ast = parseFormula(picked.source)
        } catch {
          return tag('#NAME?')
        }
        bindingAstCache.set(cacheKey, { ast })
      }
      return {
        name: picked.name,
        ...(scopeSheetId !== undefined ? { scopeSheetId } : {}),
        source: picked.source,
        ast,
        ...(picked.relativeBase !== undefined ? { relativeBase: picked.relativeBase } : {}),
        baseProvenance: picked.baseProvenance,
      }
    },

    resolve(node, ctx): ResolvedRef | EvaluationError | undefined {
      type Result = ResolvedRef | EvaluationError | undefined
      type ResolveFrame = {
        node: ReferenceNode
        phase: 'enter' | 'name' | 'union' | 'intersectLeft' | 'intersectRight'
        source?: AstNode
        index?: number
        areas?: RefArea[]
        left?: ResolvedRef
      }
      const generation = deps.generation()
      const make = (areas: readonly RefArea[]): ResolvedRef =>
        new ResolvedRefImpl(areas, deps.store, generation, deps.generation)
      const basis = ctx.currentSheet === undefined ? undefined : deps.sheetIdOfName(ctx.currentSheet)
      let knownValueBindings = ctx.nodeValues ? knownValueBindingsByMemo.get(ctx.nodeValues) : undefined
      if (!knownValueBindings) {
        knownValueBindings = new WeakMap<AstNode, Set<SheetId | undefined>>()
        if (ctx.nodeValues) knownValueBindingsByMemo.set(ctx.nodeValues, knownValueBindings)
      }
      const stack: ResolveFrame[] = [{ node, phase: 'enter' }]
      const pathSources = new Set<AstNode>()
      let result: Result
      const finish = (value: Result): void => {
        const frame = stack.pop()!
        if (frame.source) {
          pathSources.delete(frame.source)
          if (value === undefined) {
            let bases = knownValueBindings.get(frame.source)
            if (!bases) { bases = new Set(); knownValueBindings.set(frame.source, bases) }
            bases.add(basis)
          }
        }
        result = value
      }
      const cycle = (bound: NameBinding): EvaluationError => {
        markFeature(ctx, 'name-cycle')
        const message = `Defined name "${bound.name}" reaches an active binding; bounded capability gate (native circular initialization/value unmeasured)`
        deps.onIssue?.('name-cycle', message)
        ctx.reportFormulaIssue?.({ kind: 'unsupported-reference', feature: 'name-cycle', message })
        return tag('#NAME?')
      }
      while (stack.length > 0) {
        const frame = stack[stack.length - 1]
        const current = frame.node
        const relativeGate = frame.phase === 'enter' ? relativeNodeGate(current) : undefined
        if (relativeGate) {
          markFeature(ctx, relativeGate.feature)
          deps.onIssue?.(relativeGate.feature, relativeGate.message)
          ctx.reportFormulaIssue?.({ kind: 'unsupported-reference', ...relativeGate })
          finish(tag('#NAME?'))
          continue
        }
        if (frame.phase === 'name') { finish(result); continue }
        if (frame.phase === 'union' && current.type === 'union') {
          if (result === undefined) { finish(tag('#VALUE!')); continue }
          if (isTaggedError(result as ElementValue)) { finish(result); continue }
          for (const area of (result as ResolvedRef).areas) frame.areas!.push(area)
          frame.index!++
          if (frame.index! === current.refs.length) { finish(make(frame.areas!)); continue }
          stack.push({ node: current.refs[frame.index!], phase: 'enter' })
          continue
        }
        if (frame.phase === 'intersectLeft' && current.type === 'intersect') {
          if (result === undefined) { finish(tag('#NULL!')); continue }
          if (isTaggedError(result as ElementValue)) { finish(result); continue }
          frame.left = result as ResolvedRef
          frame.phase = 'intersectRight'
          stack.push({ node: current.right, phase: 'enter' })
          continue
        }
        if (frame.phase === 'intersectRight') {
          if (result === undefined) { finish(tag('#NULL!')); continue }
          if (isTaggedError(result as ElementValue)) { finish(result); continue }
          const hit = intersectAreas(frame.left!.areas, (result as ResolvedRef).areas)
          finish(hit.length === 0 ? tag('#NULL!') : make(hit))
          continue
        }
        switch (current.type) {
          case 'name': {
            const bound = services.bindName(current.ref, ctx)
            if ('kind' in bound) { finish(bound); break }
            const source = bound.ast
            // The driver owns executing VALUE names; this resolver owns only
            // reference topology. Both paths use the same bound source identity.
            if (isBindingActive(ctx, source) || pathSources.has(source)) { finish(cycle(bound)); break }
            if (knownValueBindings.get(source)?.has(basis)) { finish(undefined); break }
            if (!isReferenceNode(source)) {
              frame.source = source
              finish(undefined)
              break
            }
            let target = source
            if (hasRelativeEndpoints(target)) {
              if (!bound.relativeBase || bound.baseProvenance === 'unknown') {
                markFeature(ctx, 'relative-name')
                deps.onIssue?.('relative-name', `Relative defined name "${bound.name}" has unknown base provenance; offsets not guessed`)
                finish(tag('#NAME?'))
                break
              }
              const useSite = ctx.currentAddress ?? ctx.currentCell
              if (!useSite) {
                markFeature(ctx, 'relative-name-context-missing')
                deps.onIssue?.('relative-name-context-missing', `Relative defined name "${bound.name}" needs a supplied use-site address`)
                finish(tag('#NAME?'))
                break
              }
              const moved = translateAst(target, useSite.col - bound.relativeBase.col, useSite.row - bound.relativeBase.row)
              if (moved.type === 'error') { finish(tag('#REF!')); break }
              if (!isReferenceNode(moved)) { finish(tag('#NAME?')); break }
              target = moved
            }
            pathSources.add(source)
            frame.source = source
            frame.phase = 'name'
            stack.push({ node: target, phase: 'enter' })
            break
          }
          case 'cell':
          case 'range':
          case 'wholeCol':
          case 'wholeRow': {
            const areas = resolveGrid(current, ctx)
            finish(isTaggedError(areas as ElementValue) ? areas as EvaluationError : make(areas as RefArea[]))
            break
          }
          case 'union':
            frame.phase = 'union'
            frame.areas = []
            frame.index = 0
            if (current.refs.length === 0) finish(make([]))
            else stack.push({ node: current.refs[0], phase: 'enter' })
            break
          case 'intersect':
            if (contains3d(current, bindNameDirect, ctx)) { finish(gate3dIntersection(ctx)); break }
            frame.phase = 'intersectLeft'
            stack.push({ node: current.left, phase: 'enter' })
            break
          case 'ref3d': {
            const order = [...deps.sheets].sort((a, b) => a.workbookIndex - b.workbookIndex)
            const fromIdx = order.findIndex((s) => s.name.toLowerCase() === current.sheets.fromSheet.toLowerCase())
            const toIdx = order.findIndex((s) => s.name.toLowerCase() === current.sheets.toSheet.toLowerCase())
            // Unknown sheets are a computed #REF! (no such reference target).
            // Reversed runs have unmeasured native semantics: keep an explicit
            // unknown/provenance gate instead of assuming a computed error.
            if (fromIdx < 0 || toIdx < 0) { finish(tag('#REF!')); break }
            if (fromIdx > toIdx) {
              markFeature(ctx, 'reversed-3d')
              deps.onIssue?.(
                'reversed-3d',
                `3D sheet run "${current.sheets.fromSheet}:${current.sheets.toSheet}" is reversed; native ordering semantics unmeasured`,
              )
              finish(tag('#NAME?'))
              break
            }
            const out: RefArea[] = []
            let invalid = false
            for (let i = fromIdx; i <= toIdx; i++) {
              const part = gridAreas(current.target, order[i].sheetId)
              if (!areasInDomain(part)) { invalid = true; break }
              out.push(...part)
            }
            finish(invalid ? tag('#REF!') : make(out))
            break
          }
          case 'table': {
            const areas = resolveTableAreas(current.ref, deps.tables, ctx, deps.onIssue, ctx.currentAddress?.sheetId ?? ctx.tableContext?.currentCell.sheetId ?? (ctx.currentSheet === undefined ? undefined : deps.sheetIdOfName(ctx.currentSheet)))
            finish(isTaggedError(areas as ElementValue) ? areas as EvaluationError : make(areas as RefArea[]))
            break
          }
          case 'spill': {
            // C2 phase2: the workbook adapter owns spill ownership. Without a
            // spill service the capability stays gated (valid caches retained).
            if (ctx.spills) {
              const spilled = ctx.spills.resolveSpill(resolveSpillAddress(current, ctx), ctx)
              // A gated # outcome is a capability/provenance gate, never a
              // fabricated genuine error: mark the frame so the workbook keeps
              // its valid cache with a specific diagnostic.
              if (isTaggedError(spilled as ElementValue)) markFeature(ctx, 'spill-provenance')
              finish(spilled)
              break
            }
            markFeature(ctx, 'spill-unavailable')
            deps.onIssue?.('spill-unavailable', 'Spill ownership is not implemented without a workbook spill service')
            finish(tag('#NAME?'))
            break
          }
        }
      }
      return result
    },

    prepare(ref: ResolvedRef, ctx: EvaluationContext): void {
      // C2 phase2: settle relevant spill owners (all geometric candidates)
      // before any populated enumeration or structural absent count. The
      // request suspends through the workbook frame queue.
      ctx.spills?.prepareRegion(ref.areas, ctx)
    },
  }
  servicesRef.bindName = services.bindName
  registerReferenceGeneration(services, deps.generation)
  return services
}
