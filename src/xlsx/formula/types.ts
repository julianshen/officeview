/**
 * Pure domain types for Excel formula tokenization, AST, and evaluation.
 * Zero DOM, zero Canvas, zero XML dependencies.
 */

export type TokenType =
  | 'number'
  | 'string'
  | 'boolean'
  | 'cell'
  | 'range'
  | 'ident'
  | 'op'
  | 'comma'
  | 'lparen'
  | 'rparen'
  | 'colon'
  | 'error'
  | 'eof'
  | 'wholeCol'
  | 'wholeRow'
  | 'table'
  | 'spill'
  | 'at'
  | 'lbrace'
  | 'rbrace'
  | 'semicolon'

export interface CellRef {
  sheet?: string
  col: number // 0-based
  row: number // 0-based
  absCol: boolean
  absRow: boolean
}

export interface RangeRef {
  sheet?: string
  from: CellRef
  to: CellRef
}

export interface Token {
  type: TokenType
  value: string
  start: number
  numValue?: number
  sheet?: string
  cellRef?: CellRef
  rangeRef?: RangeRef
  wholeColRef?: WholeColSyntax
  wholeRowRef?: WholeRowSyntax
  /** Structured-reference table qualifier; absent for unqualified `[...]` refs. */
  tableName?: string
  /** Whitespace trivia preceded this token (reference intersection). */
  spaceBefore?: boolean
  /** Original identifier spelling (function names match uppercased value). */
  raw?: string
}

/** Approved modern error codes (reference-contract-final/ABI.md). */
export type SupportedModernError = '#SPILL!' | '#CALC!'
export type FormulaError =
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#NULL!'
  | SupportedModernError

export type FormulaValue = number | string | boolean | null | FormulaError

/** Tagged errors are used only during evaluation; public/model values stay primitive. */
export interface EvaluationError { readonly kind: 'formula-error'; readonly code: FormulaError }
/** Single reference/entry value (ABI ElementValue). */
export type ElementValue = FormulaValue | EvaluationError
/** Nonempty, rectangular, finite internal matrix; contains no matrices. */
export interface MatrixValue {
  readonly kind: 'formula-matrix'
  readonly rows: number
  readonly cols: number
  readonly values: readonly (readonly ElementValue[])[]
}
export type EvaluationValue = ElementValue | MatrixValue
export type EvaluationMode = 'scalar' | 'array'
export type ArrayLiteralElement = number | string | boolean | EvaluationError

export interface EvaluationContext {

  /** Typed callbacks return tagged errors and ordinary strings as text. Legacy callbacks
   * may return canonical error strings, which are interpreted as errors for compatibility.
   * Recursive typed callbacks should use evaluateFormulaInternal to preserve tags. */
  typedValues?: boolean
  /** Workbook text-length semantics: 1 = legacy Excel (UTF-16 code units),
   * 2 = current standalone Excel (code points). Default is 2. Legacy v1
   * workbooks must pass 1 via workbook/model metadata (integration item). */
  excelTextLengthVersion?: 1 | 2
  /** Actually-evaluated unsupported features for this cell's driver run
   * (unknown functions, parser-syntax constructs). Recorded at evaluation
   * time, so lazy IF/IFERROR branches that never run record nothing. */
  unsupportedFeatures?: Set<string>
  /** Per-cell completed AST nodes, retained while an iterative dependency frame resumes. */
  nodeValues?: WeakMap<AstNode, EvaluationValue>
  /** Completed function arguments and next argument, owned by one suspended cell frame. */
  flatArgs?: WeakMap<AstNode[], {
    next: number
    items: Array<{ value: EvaluationValue; fromRef: boolean }>
    /** Live sparse cursor for a reference argument being drained. */
    refCursor?: { ref: ResolvedRef; cursor: RefCursor; drainStartLen: number } | undefined
  }>

  getCellValue?(sheet: string | undefined, col: number, row: number): EvaluationValue
  /**
   * Evaluates a range of cells into a row-major 2D matrix of values.
   * Contract:
   * - Coordinates are 0-based.
   * - Returns empty/blank cells as null.
   * - Bounded to populated sheet dimensions to avoid unbounded memory allocation.
   * - Must cooperate with cycle detection by delegating or tracking visited cells.
   */
  getRangeValues?(sheet: string | undefined, from: CellRef, to: CellRef): EvaluationValue[][]
  currentSheet?: string
  currentCell?: CellRef
  /** Resolved workbook identity; does not replace currentSheet/currentCell. */
  currentAddress?: CellAddress
  /** Resolved date/locale/zone/clock semantics (C3). Explicit override wins;
   * adapters resolve once per recalculation generation and reuse the object. */
  semantics?: ResolvedSemantics
  /** Optional B1 reference services (sparse geometry, cursors, names/tables). */
  references?: ReferenceServices
  /** Containing-table context for unqualified structured references. */
  tableContext?: { tableId: string; currentCell: CellAddress }
  /** Active-mode, per-cell frame-owned scratch for resumable aggregation. */
  functionWork?: WeakMap<AstNode[], unknown>
  /** Explicit scalar/array mode; undefined is the scalar default. */
  mode?: EvaluationMode
  /** One scalar/array store pair owned by each suspended/completed cell frame. */
  modeStores?: ModeFrameStores
  /** Application resource policy for array materialization (never a native
   * shape claim): default 1048576 cells / 1048576 side. */
  maxArrayCells?: number
  maxArraySide?: number
  /** Mode-boundary evaluator used by array functions; never recursive. */
  evaluateInMode?: (node: AstNode, mode: EvaluationMode) => EvaluationValue
  /** Settle a known/possible spill owner before reading its cell (C2 phase2). */
  pendingSpill?: (request: PendingSpillRequest) => never
  /** Workbook-owned spill discovery/read services (C2 phase2). */
  spills?: SpillServices
  /** Legacy fixed f.ref projection context (C2 phase2). */
  legacyArray?: { anchor: CellAddress; output: RefArea }
  /** Native/provenance gate + unsupported-reference diagnostics channel. */
  reportFormulaIssue?: (issue: ReferenceIssue) => void
  /** In-flight evaluation path for cycle detection. Note: short-circuit-hidden cycles (e.g. IF-skipped branches) are dynamically avoided and not statically traversed. */
  visited?: Set<string>
  evalDepth?: number
  hasCycle?: boolean
}

export interface ArgumentProgress {
  next: number
  items: Array<{ value: EvaluationValue; fromRef: boolean; origin?: 'direct' | 'omitted' | 'reference' | 'array' }>
}
export interface ModeFrameStore {
  nodeValues: WeakMap<AstNode, EvaluationValue>
  flatArgs: WeakMap<AstNode[], ArgumentProgress>
  functionWork: WeakMap<AstNode[], unknown>
}
export interface ModeFrameStores { scalar: ModeFrameStore; array: ModeFrameStore }
export interface PendingSpillRequest {
  owner: CellAddress
  requestedBy: CellAddress
  phase: 'discover-shape' | 'publish'
  generation: number
}
export interface SpillServices {
  prepareCell(address: CellAddress, ctx: EvaluationContext): void
  prepareRegion(areas: readonly RefArea[], ctx: EvaluationContext): void
  readOwnedCell(address: CellAddress): ElementValue | undefined
  resolveSpill(anchor: CellAddress, ctx: EvaluationContext): ResolvedRef | EvaluationError
}

export type BinaryOp =
  | '+'
  | '-'
  | '*'
  | '/'
  | '^'
  | '&'
  | '='
  | '<>'
  | '<'
  | '<='
  | '>'
  | '>='

export type AstNode =
  | { type: 'number'; value: number }
  | { type: 'string'; value: string }
  | { type: 'boolean'; value: boolean }
  | { type: 'cell'; ref: CellRef }
  | { type: 'range'; ref: RangeRef }
  | { type: 'wholeCol'; ref: WholeColSyntax }
  | { type: 'wholeRow'; ref: WholeRowSyntax }
  | { type: 'name'; ref: NameSyntax }
  | { type: 'table'; ref: TableSyntax }
  | { type: 'ref3d'; sheets: SheetRunSyntax; target: GridReference }
  | { type: 'union'; refs: readonly ReferenceNode[] }
  | { type: 'intersect'; left: ReferenceNode; right: ReferenceNode }
  | { type: 'spill'; anchor: CellRef }
  | { type: 'implicitIntersect'; expr: AstNode }
  | { type: 'unary'; op: '+' | '-' | '%'; expr: AstNode }
  | { type: 'binary'; op: BinaryOp; left: AstNode; right: AstNode }
  | { type: 'call'; name: string; args: AstNode[] }
  | { type: 'arrayConst'; rows: readonly (readonly ArrayLiteralElement[])[] }
  | { type: 'empty' }
  | { type: 'error'; error: string }

/**
 * B0 workbook metadata (reference-contract-final/ABI.md, additive only).
 * Extraction-only placeholders; reference/array/conditional families remain
 * unimplemented and must not be advertised through these types.
 */
export type SheetId = string
export interface CellAddress { sheetId: SheetId; col: number; row: number }
export interface RefArea {
  sheetId: SheetId
  firstCol: number
  firstRow: number
  cols: number
  rows: number
}
export interface WorkbookSheetIdentity {
  sheetId: SheetId
  workbookIndex: number
  name: string
  kind: 'worksheet' | 'other'
}
export interface DefinedNameMetadata {
  name: string
  localSheetIndex?: number
  source: string
  relativeBase?: CellAddress
  baseProvenance: 'explicit' | 'verified-file' | 'unknown'
}
export interface TableMetadata {
  id: string
  name: string
  displayName: string
  sheetId: SheetId
  partPath: string
  extent: RefArea
  headerRowCount: 0 | 1
  totalsRowCount: 0 | 1
  columns: readonly { id: string; name: string; index: number }[]
}
export interface CalcSettings {
  calcMode: 'auto' | 'manual' | 'autoNoTable'
  fullCalcOnLoad: boolean
  iterate: boolean
  iterateCount: number
  iterateDelta: number
}
export interface UnicodeSettings {
  version: 1 | 2
  source: 'explicit' | 'verified-workbook' | 'workbook-default' | 'standalone-default'
  provenance?: { part: string; path: string; raw: string }
}
export interface ResolvedSemantics {
  dateSystem: '1900' | '1904'
  unicode: UnicodeSettings
  locale: string
  timeZone: string
  epochNowMs: number
}

/** Shared serial/civil date model (reference-contract-final/ABI.md). */
export type DateSystem = '1900' | '1904'
export interface SerialCivil {
  year: number
  month: number
  day: number
  special?: '1900-january-day-zero' | '1900-fictitious-february-29'
}
export interface SerialParts {
  system: '1900' | '1904'
  daySerial: number
  fraction: number
  civil: SerialCivil
}
/** Finite common-format outcome (format.ts; shared by TEXT and cell render). */
export type FormatResult =
  | { kind: 'text'; text: string }
  | { kind: 'unsupported'; reason: string }
  | EvaluationError

/**
 * B1 reference syntax (reference-contract-final/ABI.md, additive only).
 * Grammar-only shapes: the parser never consults a name/table registry.
 */
export interface AxisEndpoint { index: number; absolute: boolean }
export interface WholeColSyntax { sheet?: string; from: AxisEndpoint; to: AxisEndpoint }
export interface WholeRowSyntax { sheet?: string; from: AxisEndpoint; to: AxisEndpoint }
/** Qualification only; scope resolution happens at evaluation, never in grammar. */
export interface NameSyntax { name: string; sheet?: string }
export interface SheetRunSyntax { fromSheet: string; toSheet: string }
export type TableColumns =
  | { kind: 'all' }
  | { kind: 'single'; name: string }
  | { kind: 'range'; from: string; to: string }
export interface TableSyntax {
  /** Unqualified reference resolved from current table context. */
  table?: string
  items: readonly ('all' | 'data' | 'headers' | 'totals' | 'thisRow')[]
  columns: TableColumns
  /** Source text + translation provenance; shared copies keep identifiers. */
  raw: string
}
export type GridReference =
  | { type: 'cell'; ref: CellRef }
  | { type: 'range'; ref: RangeRef }
  | { type: 'wholeCol'; ref: WholeColSyntax }
  | { type: 'wholeRow'; ref: WholeRowSyntax }
export type ReferenceNode =
  | GridReference
  | { type: 'name'; ref: NameSyntax }
  | { type: 'table'; ref: TableSyntax }
  | { type: 'ref3d'; sheets: SheetRunSyntax; target: GridReference }
  | { type: 'union'; refs: readonly ReferenceNode[] }
  | { type: 'intersect'; left: ReferenceNode; right: ReferenceNode }
  | { type: 'spill'; anchor: CellRef }

/**
 * B1 sparse resolved references and transactional cursors (ABI.md).
 * Areas keep full geometry (A:A is 1x1048576, 1:1 is 16384x1, zero sizes
 * permitted); union duplicate occurrences are preserved as separate areas.
 */
export interface RefEntry {
  areaIndex: number
  offsetRow: number
  offsetCol: number
  address: CellAddress
  value: ElementValue
  origin: 'input' | 'formula' | 'spill' | 'absent'
}
export interface RefCursor {
  readonly id: string
  readonly mode: 'populated' | 'all'
  readonly generation: number
}
export interface ResolvedRef {
  readonly kind: 'resolved-reference'
  readonly areas: readonly RefArea[]
  readonly generation: number
  readAt(areaIndex: number, offsetRow: number, offsetCol: number): ElementValue
  openCursor(mode: 'populated' | 'all'): RefCursor
  /** Next entry WITHOUT advancing; only advance after committing the result. */
  peek(cursor: RefCursor): RefEntry | undefined
  advance(cursor: RefCursor): void
  /** Truly absent/blank cells, not formula-returned empty strings. */
  countAbsent(): number
}
export interface NameBinding {
  name: string
  scopeSheetId?: SheetId
  source: string
  ast: AstNode
  relativeBase?: CellAddress
  baseProvenance: 'explicit' | 'verified-file' | 'unknown'
}
export interface ReferenceServices {
  resolve(node: ReferenceNode, ctx: EvaluationContext): ResolvedRef | EvaluationError | undefined
  bindName(name: NameSyntax, ctx: EvaluationContext): NameBinding | EvaluationError
  /** Settle relevant spill owners first (B1: workbook-owned discovery hook). */
  prepare(ref: ResolvedRef, ctx: EvaluationContext): void
}
/** Diagnostic issue reported through EvaluationContext.reportFormulaIssue. */
export interface ReferenceIssue {
  kind:
    | 'native-gate'
    | 'unsupported-reference'
    | 'spill-collision'
    | 'spill-resource'
    | 'spill-provenance'
    | 'spill-discovery'
    | 'spill-arbitration'
  gate?: string
  feature?: string
  message: string
}

/**
 * SPILLS.md owner metadata and registry contract (C2 phase2). The workbook
 * adapter owns the registry, original-input occupancy, reverse dependency
 * graph and frame queues; these types are additive and never change existing
 * Stage A/B1 shapes.
 */
export type ArrayProvenanceKind = 'dynamic' | 'legacy-array' | 'possible-dynamic' | 'ambiguous'
export interface ArrayProvenance {
  kind: ArrayProvenanceKind
  source: 'verified-file' | 'explicit-mode' | 'syntax-capability'
  evidence?: { part: string; path: string; raw: string }
}
export type SpillState =
  | 'undiscovered'
  | 'discovering'
  | 'computing'
  | 'reserved'
  | 'committed'
  | 'cached'
  | 'nonspilling'
  | 'failed'
export interface SpillRecord {
  anchor: CellAddress
  provenance: ArrayProvenance
  generation: number
  sourceRevision: number
  state: SpillState
  cachedArea?: RefArea
  liveArea?: RefArea
  matrix?: MatrixValue
  /** Generated cells only, excluding the anchor. */
  ownedCells: Set<string>
  dependents: Set<string>
}
export type ShapeDiscovery =
  | { kind: 'scalar' }
  | { kind: 'array'; rows: number; cols: number }
  | { kind: 'value-dependent'; dependencies: readonly CellAddress[] }
export interface SpillRegistry {
  byAnchor: Map<string, SpillRecord>
  ownerOfGenerated: Map<string, string>
  candidatesForCell(address: CellAddress): readonly string[]
  candidatesForRegion(areas: readonly RefArea[]): readonly string[]
  discoverShape(anchor: CellAddress, ctx: EvaluationContext): ShapeDiscovery
  invalidate(anchor: CellAddress, nextSourceRevision: number): void
}


