/**
 * C2 phase2 workbook spill lifecycle engine (private module).
 *
 * Owns the ownership registry, potential-owner index, transactional
 * publication/rollback, own-only cleanup bookkeeping and A1# resolution for
 * the workbook adapter (reference-contract-final/SPILLS.md). It is deliberately
 * host-driven: the workbook supplies occupancy classification, generated-cell
 * writes and diagnostics, so no evaluation of formulas happens here and no
 * public ABI escapes beyond the approved SPILLS.md shapes.
 *
 * Scheduling reuses the existing Stage A frame queue: an unresolved owner is
 * requested through `ctx.pendingSpill` (throwing `SpillSuspension`), the
 * workbook pushes the owner frame and resumes the same requesting frame — never
 * a recursive evaluateWorkbook/evaluateArrayFormula call.
 */
import { formulaError, isEvaluationError, topLeftOfMatrix } from './evaluator'
import { createResolvedRef, type SparseCellStore, type StoredCell } from './refs'
import type {
  ArrayProvenance,
  CellAddress,
  ElementValue,
  EvaluationContext,
  EvaluationError,
  MatrixValue,
  PendingSpillRequest,
  RefArea,
  ResolvedRef,
  SheetId,
  SpillRecord,
} from './types'

/** Verified documented #SPILL! causes (root primary-spill-root-r1/FACTS.json). */
export type SpillBlockReason =
  | 'input-cell'
  | 'formula-cell'
  | 'merged-cell'
  | 'table'
  | 'sheet-edge'
  | 'foreign-spill'

export interface SpillIssue {
  kind:
    | 'spill-collision'
    | 'spill-resource'
    | 'spill-provenance'
    | 'spill-discovery'
    | 'spill-arbitration'
    | 'native-gate'
  feature: string
  message: string
  anchor?: CellAddress
  blockedAt?: CellAddress
}

export interface SpillHost {
  sheetIdOfName(name: string): SheetId | undefined
  sheetNameOfId(id: SheetId): string | undefined
  /** True when the address holds an original (non-generated) input or formula. */
  isOriginalOccupant(address: CellAddress): boolean
  /** True when any cell in the region is NOT an original occupant (so a spill
   * owner could legitimately own part of it). Conservative for huge regions. */
  hasOwnableCell(areas: readonly RefArea[]): boolean
  /** Reason an ORIGINAL occupant blocks a footprint cell (undefined = clear). */
  originalBlocker(address: CellAddress): SpillBlockReason | undefined
  /** Install/overwrite a generated follower's primitive model value. */
  writeGenerated(address: CellAddress, value: ElementValue, valueIsError: boolean, ownerKey: string): void
  /** Mark a committed owner as failed (arbitration): clear its followers and
   * set its anchor to #SPILL! with a diagnostic. */
  failCommittedOwner(ownerKey: string, anchor: CellAddress): void
  report(issue: SpillIssue): void
}

/** Private scheduler suspension: never a native error, never caught by coercion. */
export class SpillSuspension extends Error {
  constructor(readonly request: PendingSpillRequest) {
    super(`spill ${request.phase} for ${request.owner.sheetId}!${request.owner.col}:${request.owner.row}`)
    this.name = 'SpillSuspension'
  }
}

export const SHEET_MAX_COLS = 16384
export const SHEET_MAX_ROWS = 1048576

export function spillCellKey(address: CellAddress): string {
  return `${address.sheetId}!${address.col}:${address.row}`
}

function inArea(area: RefArea, col: number, row: number): boolean {
  return col >= area.firstCol && col < area.firstCol + area.cols && row >= area.firstRow && row < area.firstRow + area.rows
}

function areasOverlap(a: RefArea, b: RefArea): boolean {
  if (a.sheetId !== b.sheetId || a.cols <= 0 || a.rows <= 0 || b.cols <= 0 || b.rows <= 0) return false
  return a.firstCol < b.firstCol + b.cols && b.firstCol < a.firstCol + a.cols &&
    a.firstRow < b.firstRow + b.rows && b.firstRow < a.firstRow + a.rows
}

function describe(address: CellAddress): string {
  return `${address.sheetId}!r${address.row + 1}c${address.col + 1}`
}

export interface SpillEngineOptions {
  maxArrayCells?: number
  maxArraySide?: number
}

/**
 * Per-recalculation spill engine. One instance is created for each workbook
 * recalculation; persistent ownership lives in the workbook adapter (doc
 * identity) and is replayed as generated-cell cleanup before construction.
 */
export class SpillEngine {
  private readonly records = new Map<string, SpillRecord>()
  private readonly ownerOfGenerated = new Map<string, string>()
  /** Explicit UNAVAILABLE gates (arbitration/discovery): internal #NAME? + marks. */
  private readonly gates = new Map<string, string>()
  private readonly maxCells: number
  private readonly maxSide: number

  constructor(
    private readonly host: SpillHost,
    private readonly generation: number,
    opts?: SpillEngineOptions,
  ) {
    this.maxCells = opts?.maxArrayCells ?? 1048576
    this.maxSide = opts?.maxArraySide ?? 1048576
  }

  // -------------------------------------------------------------------------
  // Registration (before any formula evaluation)
  // -------------------------------------------------------------------------

  registerOwner(anchor: CellAddress, provenance: ArrayProvenance, declaredArea?: RefArea): void {
    const key = spillCellKey(anchor)
    if (this.records.has(key)) return
    const record: SpillRecord = {
      anchor,
      provenance,
      generation: this.generation,
      sourceRevision: 0,
      state: 'undiscovered',
      ownedCells: new Set(),
      dependents: new Set(),
      ...(declaredArea ? { cachedArea: declaredArea } : {}),
    }
    this.records.set(key, record)
  }

  hasOwner(address: CellAddress): boolean {
    return this.records.has(spillCellKey(address))
  }

  recordFor(address: CellAddress): SpillRecord | undefined {
    return this.records.get(spillCellKey(address))
  }

  ownerKeyForGenerated(address: CellAddress): string | undefined {
    return this.ownerOfGenerated.get(spillCellKey(address))
  }

  /** Feature name when an owner is gated (internal #NAME? + frame marks). */
  gateFeature(address: CellAddress): string | undefined {
    return this.gates.get(spillCellKey(address))
  }

  private isSelf(anchor: CellAddress, ctx: EvaluationContext): boolean {
    const current = ctx.currentAddress
    return current !== undefined &&
      current.sheetId === anchor.sheetId && current.col === anchor.col && current.row === anchor.row
  }

  // -------------------------------------------------------------------------
  // Candidate index (geometric only; never enumerates blank cells)
  // -------------------------------------------------------------------------

  candidatesForCell(address: CellAddress): readonly string[] {
    const out: string[] = []
    for (const [key, rec] of this.records) {
      if (rec.anchor.sheetId !== address.sheetId) continue
      if (rec.state === 'committed') {
        if (rec.liveArea && inArea(rec.liveArea, address.col, address.row)) out.push(key)
        continue
      }
      if (rec.state === 'nonspilling' || rec.state === 'failed') continue
      if (rec.cachedArea && rec.provenance.kind === 'legacy-array') {
        if (inArea(rec.cachedArea, address.col, address.row)) out.push(key)
        continue
      }
      // Potential dynamic output extends only down/right from the anchor.
      if (rec.anchor.col <= address.col && rec.anchor.row <= address.row) out.push(key)
    }
    return out
  }

  candidatesForRegion(areas: readonly RefArea[]): readonly string[] {
    const out = new Set<string>()
    for (const area of areas) {
      if (area.cols <= 0 || area.rows <= 0) continue
      const lastCol = area.firstCol + area.cols - 1
      const lastRow = area.firstRow + area.rows - 1
      for (const [key, rec] of this.records) {
        if (rec.anchor.sheetId !== area.sheetId) continue
        if (rec.state !== 'undiscovered' && rec.state !== 'discovering') continue
        if (rec.cachedArea && rec.provenance.kind === 'legacy-array') {
          if (areasOverlap(rec.cachedArea, area)) out.add(key)
          continue
        }
        if (rec.anchor.col <= lastCol && rec.anchor.row <= lastRow) out.add(key)
      }
    }
    return [...out]
  }

  // -------------------------------------------------------------------------
  // Read-path preparation (before any absent/generated memo)
  // -------------------------------------------------------------------------

  prepareCell(address: CellAddress, ctx: EvaluationContext): void {
    const key = spillCellKey(address)
    if (this.ownerOfGenerated.has(key)) return
    const self = this.records.get(key)
    if (self && (self.state === 'committed' || self.state === 'nonspilling' || self.state === 'failed')) return
    for (const candidate of this.candidatesForCell(address)) {
      const rec = this.records.get(candidate)!
      if (rec.state === 'failed' || rec.state === 'nonspilling') continue
      if (this.isSelf(rec.anchor, ctx)) {
        // A known (committed/cached) disjoint self shape was already excluded
        // by candidatesForCell; an UNKNOWN self shape reading an absent cell is
        // an unresolved discovery loop: gate it, never guess a blank/zero.
        if (rec.liveArea ?? (rec.provenance.kind === 'legacy-array' ? rec.cachedArea : undefined)) continue
        this.markFailed(rec.anchor, 'spill-discovery', `Spill owner ${describe(rec.anchor)} reads ${describe(address)} inside its own undiscovered output region; ownership discovery gated (native circular initialization unmeasured), valid cache retained`)
        continue
      }
      if (rec.state === 'undiscovered' || rec.state === 'discovering') this.request(rec, ctx, address)
    }
  }

  prepareRegion(areas: readonly RefArea[], ctx: EvaluationContext): void {
    // C4 r4 SHARED REGION-AVAILABILITY BOUNDARY. A gated (unavailable) owner
    // whose output footprint is UNKNOWN is not a confirmed absence. Surface the
    // region gate BEFORE any populated cursor drain or structural absent count,
    // so a sparse region reader (SUM/COUNTIF/COUNT over the potential
    // footprint) keeps its valid cache instead of publishing a guessed blank/0.
    // `gateForRegion` only consults KNOWN footprints plus the conservative
    // down/right potential region — no precise failed extent is invented.
    const regionGate = this.gateForRegion(areas)
    if (regionGate) {
      const record = ctx as EvaluationContext & { unsupportedFeatures?: Set<string> }
      if (!record.unsupportedFeatures) record.unsupportedFeatures = new Set()
      record.unsupportedFeatures.add(regionGate)
    }
    let ownable: boolean | undefined
    for (const candidate of this.candidatesForRegion(areas)) {
      const rec = this.records.get(candidate)!
      if (rec.state === 'failed' || rec.state === 'nonspilling') continue
      if (this.isSelf(rec.anchor, ctx)) {
        if (rec.liveArea ?? (rec.provenance.kind === 'legacy-array' ? rec.cachedArea : undefined)) continue
        // A region made entirely of original inputs/formulas needs no owner
        // discovery (Stage A normal); otherwise the self shape is unresolved.
        ownable ??= this.host.hasOwnableCell(areas)
        if (!ownable) continue
        this.markFailed(rec.anchor, 'spill-discovery', `Spill owner ${describe(rec.anchor)} reads a region intersecting its own undiscovered output; ownership discovery gated (native circular initialization unmeasured), valid cache retained`)
        continue
      }
      if (rec.state === 'undiscovered' || rec.state === 'discovering') {
        const area = areas[0]
        this.request(rec, ctx, { sheetId: area.sheetId, col: area.firstCol, row: area.firstRow })
      }
    }
  }

  private request(rec: SpillRecord, ctx: EvaluationContext, requestedBy: CellAddress): never {
    rec.state = 'discovering'
    const request: PendingSpillRequest = {
      owner: rec.anchor,
      requestedBy,
      phase: rec.cachedArea ? 'publish' : 'discover-shape',
      generation: this.generation,
    }
    if (ctx.pendingSpill) ctx.pendingSpill(request)
    throw new SpillSuspension(request)
  }

  readOwnedCell(address: CellAddress): ElementValue | undefined {
    const ownerKey = this.ownerOfGenerated.get(spillCellKey(address))
    if (!ownerKey) return undefined
    const rec = this.records.get(ownerKey)
    if (!rec || !rec.matrix || !rec.liveArea) return undefined
    const row = address.row - rec.liveArea.firstRow
    const col = address.col - rec.liveArea.firstCol
    if (row < 0 || col < 0 || row >= rec.matrix.rows || col >= rec.matrix.cols) return undefined
    return rec.matrix.values[row][col]
  }

  /**
   * A gated owner (arbitration/discovery) whose footprint is UNKNOWN is not a
   * confirmed absence: a reader of a cell in its potential output region must
   * propagate the internal #NAME? gate + frame mark (retaining its own valid
   * cache) instead of memoizing a guessed blank.
   */
  gateForCell(address: CellAddress): string | undefined {
    for (const [key, feature] of this.gates) {
      const rec = this.records.get(key)
      if (!rec || rec.anchor.sheetId !== address.sheetId) continue
      const area = rec.liveArea ?? (rec.provenance.kind === 'legacy-array' ? rec.cachedArea : undefined)
      if (area) {
        if (inArea(area, address.col, address.row)) return feature
      } else if (rec.anchor.col <= address.col && rec.anchor.row <= address.row) {
        return feature
      }
    }
    return undefined
  }

  gateForRegion(areas: readonly RefArea[]): string | undefined {
    for (const [key, feature] of this.gates) {
      const rec = this.records.get(key)
      if (!rec) continue
      const area = rec.liveArea ?? (rec.provenance.kind === 'legacy-array' ? rec.cachedArea : undefined)
      for (const a of areas) {
        if (rec.anchor.sheetId !== a.sheetId || a.cols <= 0 || a.rows <= 0) continue
        if (area) {
          if (areasOverlap(area, a)) return feature
        } else if (rec.anchor.col <= a.firstCol + a.cols - 1 && rec.anchor.row <= a.firstRow + a.rows - 1) {
          return feature
        }
      }
    }
    return undefined
  }

  // -------------------------------------------------------------------------
  // Publication (transactional)
  // -------------------------------------------------------------------------

  /**
   * Publish an owner's evaluated result. `matrix` is the raw array-mode
   * EvaluationValue when the root produced an actual matrix; `undefined` means
   * the root resolved to a scalar outcome (nonspilling). Returns the primitive
   * value the workbook must write for the anchor cell.
   */
  publish(anchor: CellAddress, matrix: MatrixValue | undefined, scalarValue: ElementValue): ElementValue {
    const key = spillCellKey(anchor)
    const rec = this.records.get(key)
    if (!rec) return scalarValue
    if (!matrix) {
      rec.state = 'nonspilling'
      rec.sourceRevision++
      return scalarValue
    }
    const rows = matrix.rows
    const cols = matrix.cols
    if (
      !Number.isSafeInteger(rows) || !Number.isSafeInteger(cols) || rows < 1 || cols < 1 ||
      rows > this.maxSide || cols > this.maxSide || rows * cols > this.maxCells
    ) {
      rec.state = 'failed'
      this.host.report({
        kind: 'spill-resource',
        feature: 'spill-resource',
        message: `Spill anchor ${describe(anchor)} requests ${rows}x${cols}; resource policy (maxArrayCells ${this.maxCells}, maxArraySide ${this.maxSide}) maps to #SPILL!`,
        anchor,
      })
      return formulaError('#SPILL!')
    }
    const area: RefArea = { sheetId: anchor.sheetId, firstCol: anchor.col, firstRow: anchor.row, cols, rows }
    for (let row = 0; row < rows; row++) {
      for (let col = 0; col < cols; col++) {
        if (row === 0 && col === 0) continue // anchor itself is permitted
        const address: CellAddress = { sheetId: anchor.sheetId, col: anchor.col + col, row: anchor.row + row }
        if (address.col >= SHEET_MAX_COLS || address.row >= SHEET_MAX_ROWS) {
          return this.collide(rec, address, 'sheet-edge')
        }
        const occupantOwner = this.ownerOfGenerated.get(spillCellKey(address))
        if (occupantOwner && occupantOwner !== key) return this.collide(rec, address, 'foreign-spill')
        const blocker = this.host.originalBlocker(address)
        // An original formula cell is ALWAYS a known #SPILL! collision — it is
        // never converted into arbitration merely because it also happens to be
        // a registered array owner (root correction).
        if (blocker) return this.collide(rec, address, blocker)
      }
    }
    // True UNKNOWN output overlap is identified BEFORE any publication: a
    // competing dynamic owner whose anchor is not in this footprint, but whose
    // output rectangle (known, or the conservative down/right potential region)
    // intersects it, is a native-gated arbitration. Both anchors are gated with
    // no winner by publication/committed order and no transient followers.
    for (const [otherKey, other] of this.records) {
      if (otherKey === key || other.provenance.kind === 'legacy-array') continue
      if (other.state === 'failed' || other.state === 'nonspilling') continue
      // O's anchor inside this footprint is the known collision handled above.
      if (inArea(area, other.anchor.col, other.anchor.row)) continue
      const otherArea = other.liveArea
      if (otherArea) {
        if (!areasOverlap(area, otherArea)) continue
      } else {
        // Unknown shape: O's output can only extend down/right from its anchor.
        if (other.anchor.col > area.firstCol + area.cols - 1 || other.anchor.row > area.firstRow + area.rows - 1) continue
        // If THIS anchor lies inside O's potential region, O will collide with
        // this formula cell — a known blockage for O, not arbitration.
        if (other.anchor.col <= anchor.col && other.anchor.row <= anchor.row) continue
      }
      return this.arbitrate(rec, other)
    }
    for (let row = 0; row < rows; row++) {
      for (let col = 0; col < cols; col++) {
        if (row === 0 && col === 0) continue
        const address: CellAddress = { sheetId: anchor.sheetId, col: anchor.col + col, row: anchor.row + row }
        const value = matrix.values[row][col]
        this.host.writeGenerated(address, value, isEvaluationError(value), key)
        this.ownerOfGenerated.set(spillCellKey(address), key)
        rec.ownedCells.add(spillCellKey(address))
      }
    }
    rec.state = 'committed'
    rec.liveArea = area
    rec.matrix = matrix
    rec.sourceRevision++
    return topLeftOfMatrix(matrix)
  }

  /**
   * Order-independent competing-owner gate: neither anchor wins by traversal
   * or committed order. Both are failed with an explicit UNAVAILABLE gate
   * (internal #NAME? + frame marks) and any committed followers of the other
   * owner are rolled back with its valid cache retained.
   */
  private arbitrate(rec: SpillRecord, other: SpillRecord): EvaluationError {
    this.host.report({
      kind: 'spill-arbitration',
      feature: 'spill-arbitration',
      message: `Competing spill anchors ${describe(rec.anchor)} and ${describe(other.anchor)} have overlapping output without a verified arbitration rule; both gated (internal #NAME?, valid caches retained), no loop-order winner`,
      anchor: rec.anchor,
    })
    const otherKey = spillCellKey(other.anchor)
    const otherWasCommitted = other.state === 'committed'
    other.state = 'failed'
    other.liveArea = undefined
    other.matrix = undefined
    this.gates.set(otherKey, 'spill-arbitration')
    if (otherWasCommitted) this.host.failCommittedOwner(otherKey, other.anchor)
    rec.state = 'failed'
    rec.liveArea = undefined
    rec.matrix = undefined
    this.gates.set(spillCellKey(rec.anchor), 'spill-arbitration')
    return formulaError('#NAME?')
  }

  private collide(rec: SpillRecord, blockedAt: CellAddress, reason: SpillBlockReason): EvaluationError {    rec.state = 'failed'
    rec.liveArea = undefined
    rec.matrix = undefined
    this.host.report({
      kind: 'spill-collision',
      feature: 'spill-collision',
      message: `Spill anchor ${describe(rec.anchor)} blocked at ${describe(blockedAt)} by ${reason}; #SPILL! published with no partial followers`,
      anchor: rec.anchor,
      blockedAt,
    })
    return formulaError('#SPILL!')
  }

  /**
   * Legacy fixed f.ref projection: map the evaluated matrix by declared output
   * offsets, no dynamic collision transaction. A declared/actual mismatch is a
   * native-gated case (general padding/scalar repetition unmeasured), never an
   * invented #VALUE! or fabricated padding.
   */
  publishLegacy(anchor: CellAddress, matrix: MatrixValue): ElementValue {
    const key = spillCellKey(anchor)
    const rec = this.records.get(key)
    const declared = rec?.cachedArea
    const topLeft = topLeftOfMatrix(matrix)
    if (!rec || !declared) return topLeft
    if (declared.cols !== matrix.cols || declared.rows !== matrix.rows) {
      rec.state = 'failed'
      this.host.report({
        kind: 'native-gate',
        feature: 'legacy-array-output-mismatch',
        message: `Legacy fixed f.ref ${declared.cols}x${declared.rows} does not match the evaluated ${matrix.cols}x${matrix.rows} matrix; general padding/scalar repetition is unmeasured and not invented`,
        anchor,
      })
      return topLeft
    }
    for (let row = 0; row < matrix.rows; row++) {
      for (let col = 0; col < matrix.cols; col++) {
        if (row === 0 && col === 0) continue
        const address: CellAddress = { sheetId: declared.sheetId, col: declared.firstCol + col, row: declared.firstRow + row }
        const value = matrix.values[row][col]
        this.host.writeGenerated(address, value, isEvaluationError(value), key)
        this.ownerOfGenerated.set(spillCellKey(address), key)
        rec.ownedCells.add(spillCellKey(address))
      }
    }
    rec.state = 'committed'
    rec.liveArea = declared
    rec.matrix = matrix
    rec.sourceRevision++
    return topLeft
  }

  /** Mark an owner as failed without publication (circular/unsupported discovery). */
  markFailed(anchor: CellAddress, feature: string, message: string): void {
    const rec = this.records.get(spillCellKey(anchor))
    if (!rec) return
    rec.state = 'failed'
    rec.liveArea = undefined
    rec.matrix = undefined
    this.gates.set(spillCellKey(anchor), feature)
    this.host.report({ kind: 'spill-discovery', feature, message, anchor })
  }

  // -------------------------------------------------------------------------
  // A1# resolution
  // -------------------------------------------------------------------------

  resolveSpill(anchor: CellAddress, ctx: EvaluationContext): ResolvedRef | EvaluationError {
    const rec = this.records.get(spillCellKey(anchor))
    if (!rec) {
      this.host.report({
        kind: 'spill-provenance',
        feature: 'spill-provenance',
        message: `# at ${describe(anchor)} has no verified dynamic/legacy provenance; the failed/nonspilling-anchor error is unmeasured and not invented`,
        anchor,
      })
      return formulaError('#NAME?')
    }
    if (rec.state === 'undiscovered' || rec.state === 'discovering') {
      this.request(rec, ctx, ctx.currentAddress ?? anchor)
    }
    if (rec.state !== 'committed') {
      this.host.report({
        kind: 'spill-provenance',
        feature: 'spill-provenance',
        message: `# at ${describe(anchor)} refers to an anchor whose state is ${rec.state}; exact failed/nonspilling-anchor error is unmeasured`,
        anchor,
      })
      return formulaError('#NAME?')
    }
    const liveArea = rec.liveArea
    const matrix = rec.matrix
    if (!liveArea || !matrix) {
      this.host.report({
        kind: 'spill-provenance',
        feature: 'spill-provenance',
        message: `# at ${describe(anchor)} refers to a committed anchor without a live footprint; exact failed/nonspilling-anchor error is unmeasured`,
        anchor,
      })
      return formulaError('#NAME?')
    }
    const store: SparseCellStore = {
      read(address) {
        const row = address.row - liveArea.firstRow
        const col = address.col - liveArea.firstCol
        if (row < 0 || col < 0 || row >= matrix.rows || col >= matrix.cols) return undefined
        return { value: matrix.values[row][col], origin: 'spill' }
      },
      *stored(): Iterable<StoredCell> {
        for (let row = 0; row < matrix.rows; row++) {
          for (let col = 0; col < matrix.cols; col++) {
            yield {
              address: { sheetId: liveArea.sheetId, col: liveArea.firstCol + col, row: liveArea.firstRow + row },
              value: matrix.values[row][col],
              origin: 'spill',
            }
          }
        }
      },
    }
    return createResolvedRef([liveArea], store, this.generation, () => this.generation)
  }
}
