/** Sort/filter operation range detection and structural admission (spec §6.1).
 * Pure module over caller-owned source state. Resolution precedence is strict:
 * a table containing the active cell wins with its metadata header/totals, then
 * an explicit contiguous selection, then the current region grown from the
 * active cell by bounded ring probes, then (for a blank seed only) the used
 * fallback. The used extent is a hint: the fallback resolves to the sole
 * current region R of the first populated cell in it, and only when nothing
 * else in the used range is populated; R itself is returned (it need not equal
 * the used extent, whose blank margins would otherwise move styled blank rows),
 * with source 'used'. Two blocks around a blank cell refuse as ambiguous.
 * Disjoint Shift selections are never coerced into a bounding rectangle: only
 * the active cell seeds detection. A derived range (explicit, region or used)
 * that intersects any table it does not exactly equal refuses with
 * 'ambiguous-selection' (it would mix table and non-table rows, or slice a
 * table seeded from outside it); one exactly equal to a table takes the table
 * path with its metadata. Guards run through the shared S3 guard; nothing here publishes,
 * mutates, or invents values. Unknown coverage waits, refusals carry typed
 * diagnostics, source failures map to typed source-error refusals, and
 * cancellation maps to cancelled. Outcomes are deep-frozen.
 *
 * No resume token is carried across an `incomplete` outcome: a retry re-runs
 * detection from scratch once the caller has completed the missing ranges. */
import type {
  AuthoritativeCell, CompactSourceRowSet, FilterRange, GridSelection, OperationRefusal, Page, ReadOptions,
  SourceFailure, SourceStamp, SourceStore, TableExtent,
} from './contracts'
import {
  MAX_COLS, MAX_ROWS, isValidCoordinate, isValidSourceRectInput,
  type SourceCell, type SourceRect,
} from './coordinates'
import type { PopulatedCellSource, StructuralIndex } from './structure'
import { guardSort, isChildOnlyRange } from './operationGuard'

export interface RangeDetectionRequest {
  store: SourceStore
  structure: StructuralIndex
  selection: GridSelection
  options: ReadOptions
  /** Sort-dialog header override for the non-table heuristic. Spec §6.4 offers
   * a header override only, so a totals override is refused as invalid-rule.
   * Tables keep their metadata and report the override as ignored. */
  override?: { headerRows: 0 | 1; totalsRows?: 0 } | null
  pageLimit?: number
  /** Used extent hint for the blank-seed fallback; defaults to the store extent,
   * else the whole sheet. Region growth never stops at it: data beyond the hint
   * still joins. */
  usedRange?: SourceRect
  /** Detection work cap: populated scan pages requested plus cells received
   * (default DEFAULT_DETECTION_WORK_BUDGET). Exceeding it refuses with
   * resource-limit; the shared guard's streaming scan is not counted. */
  workBudget?: number
}
export type RangeSource =
  | { kind: 'table'; id: string }
  | { kind: 'region' }
  | { kind: 'used' }
  | { kind: 'explicit' }
/** metadata: declared table counts; heuristic: §6.1 first-row rule; override: dialog. */
export type HeaderProvenance = 'metadata' | 'heuristic' | 'override'
export interface DetectedRange {
  kind: 'complete'
  range: FilterRange
  /** Source owner the range was resolved against; structure, store and read stamps agree. */
  stamp: SourceStamp
  source: RangeSource
  header: HeaderProvenance
  /** True when a dialog override was supplied for a table and metadata won (§6.1). */
  overrideIgnored: boolean
}
export type RangeDetection =
  | DetectedRange
  | { kind: 'incomplete'; missing: readonly SourceRect[] }
  | { kind: 'refused'; refusal: OperationRefusal }
  | { kind: 'cancelled' }

/** Generous enough for a full-capacity single-column region (~1M cells plus a
 * few thousand pages) while bounding pathological sources. */
export const DEFAULT_DETECTION_WORK_BUDGET = 2 ** 24
/** Largest gallop strip, in lines; also capped by the caller's page limit. */
const STRIP_MAX_LINES = 1024
const YIELD_EVERY_PAGES = 16
const YIELD_EVERY_MS = 4
const FULL_DOMAIN: SourceRect = { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }

/** Typed source failure raised by scan pages; mapped to a source-error refusal. */
class ScanFailure extends Error {
  readonly failure: SourceFailure
  constructor(failure: SourceFailure) {
    super(failure.message)
    this.failure = failure
  }
}
class WorkBudgetExceeded extends Error {}
/** A page that breaks the store contract (cells outside the request, stalled
 * cursors, missing pages) is a non-retryable malformed page, never a throw. */
function contractViolation(message: string): ScanFailure {
  return new ScanFailure({ code: 'page-error', retryable: false, message })
}

function frozenCopy<T>(value: T): T {
  if (Array.isArray(value)) return Object.freeze(value.map(frozenCopy)) as unknown as T
  if (typeof value === 'object' && value !== null) {
    const out: Record<string, unknown> = {}
    for (const [key, entry] of Object.entries(value)) out[key] = frozenCopy(entry)
    return Object.freeze(out) as T
  }
  return value
}
function refused(refusal: OperationRefusal): RangeDetection {
  return { kind: 'refused', refusal }
}
/** Guard and child-only outcomes share the refusal shape, but an 'incomplete'
 * reason means unknown coverage (wait and retry), not a terminal refusal. */
function translateRefusal(refusal: OperationRefusal): RangeDetection {
  if (refusal.reason === 'incomplete') return { kind: 'incomplete', missing: refusal.missing }
  return refused(refusal)
}
function sourceError(failure: SourceFailure): RangeDetection {
  return refused({ reason: 'source-error', code: failure.code, retryable: failure.retryable, message: failure.message })
}
function sameStamp(a: SourceStamp, b: SourceStamp): boolean {
  return a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen
}
/** Refusal for a non-complete structural index: a source failure outranks a
 * bare wait (it carries the code and retryability), else the missing ranges. */
function structuralWait(structure: StructuralIndex): RangeDetection {
  if (structure.failure !== undefined) return sourceError(structure.failure)
  return { kind: 'incomplete', missing: structure.missing }
}

/** Host yield that gives timers (and abort callbacks) a real macrotask turn. */
function yieldToHost(): Promise<void> {
  const host = globalThis as typeof globalThis & {
    setImmediate?: (callback: () => void) => unknown
    scheduler?: { yield?: () => Promise<void> }
  }
  if (host.setImmediate) return new Promise(resolve => { host.setImmediate!(resolve) })
  if (host.scheduler?.yield) return host.scheduler.yield()
  if (typeof MessageChannel !== 'undefined') {
    return new Promise(resolve => {
      const channel = new MessageChannel()
      channel.port1.onmessage = () => {
        channel.port1.close()
        channel.port2.close()
        resolve()
      }
      channel.port2.postMessage(null)
    })
  }
  return new Promise(resolve => setTimeout(resolve, 0))
}
/** Cooperative pacing: yield every few pages or milliseconds, then re-check the
 * signal, so an abort scheduled mid-run lands even when source pages resolve
 * as microtasks. */
class Pacer {
  private pages = 0
  private last = performance.now()
  constructor(private readonly signal: AbortSignal) {}
  async tick(): Promise<void> {
    this.signal.throwIfAborted()
    this.pages++
    if (this.pages < YIELD_EVERY_PAGES && performance.now() - this.last < YIELD_EVERY_MS) return
    await yieldToHost()
    this.pages = 0
    this.last = performance.now()
    this.signal.throwIfAborted()
  }
}

/** One budgeted, paced, typed view of store.scanPopulated. Error pages and
 * foreign page stamps raise ScanFailure; nothing is retained between pages. */
class Scanner {
  private work = 0
  constructor(
    private readonly store: SourceStore,
    private readonly options: ReadOptions,
    private readonly pacer: Pacer,
    private readonly budget: number,
  ) {}
  async page(rect: SourceRect, limit: number, cursor?: string): Promise<Page<AuthoritativeCell>> {
    if (this.work + 1 > this.budget) throw new WorkBudgetExceeded(`Detection exceeded its work budget of ${this.budget}`)
    this.work++
    await this.pacer.tick()
    const request = cursor === undefined ? { range: rect, limit } : { range: rect, limit, cursor }
    let first: Page<AuthoritativeCell> | undefined
    for await (const page of this.store.scanPopulated(request, this.options)) {
      first = page
      break
    }
    // An abort that landed while this page was in flight wins over its content.
    this.options.signal.throwIfAborted()
    if (first === undefined) throw contractViolation('Populated scan ended without a page')
    if (first.error !== undefined) throw new ScanFailure(first.error)
    if (!sameStamp(first.stamp, this.options.stamp)) {
      throw new ScanFailure({ code: 'stale-source', retryable: true, message: 'Populated scan page belongs to another source stamp' })
    }
    this.work += first.items.length
    for (const cell of first.items) {
      if (!isValidCoordinate(cell.row, cell.col)) throw contractViolation('Populated cell outside the worksheet domain')
      if (cell.row < rect.r0 || cell.row > rect.r1 || cell.col < rect.c0 || cell.col > rect.c1) {
        throw contractViolation('Populated cell outside the requested range')
      }
    }
    return first
  }
  /** Existence probe: follows cursors until a cell or exhaustion. A found cell
   * is decisive even on a loading page; without one, a loading page is unknown. */
  async exists(rect: SourceRect): Promise<'occupied' | 'blank' | 'unknown'> {
    let cursor: string | undefined
    let complete = true
    const seen = new Set<string>()
    for (;;) {
      const page = await this.page(rect, 1, cursor)
      if (page.items.length > 0) return 'occupied'
      if (page.completeness !== 'complete') complete = false
      if (page.nextCursor === undefined) return complete ? 'blank' : 'unknown'
      if (seen.has(page.nextCursor)) throw contractViolation('Populated cursor made no progress')
      seen.add(page.nextCursor)
      cursor = page.nextCursor
    }
  }
  /** First populated cell in rect, or why there is none. */
  async first(rect: SourceRect): Promise<SourceCell | 'blank' | 'unknown'> {
    let cursor: string | undefined
    let complete = true
    const seen = new Set<string>()
    for (;;) {
      const page = await this.page(rect, 1, cursor)
      if (page.items.length > 0) return { row: page.items[0].row, col: page.items[0].col }
      if (page.completeness !== 'complete') complete = false
      if (page.nextCursor === undefined) return complete ? 'blank' : 'unknown'
      if (seen.has(page.nextCursor)) throw contractViolation('Populated cursor made no progress')
      seen.add(page.nextCursor)
      cursor = page.nextCursor
    }
  }
  /** Guard feed: same typed failures and pacing, no detection budget. */
  asPopulatedSource(defaultLimit: number): PopulatedCellSource {
    return {
      iterate: async (range, cursor, limit) => {
        const page = await this.page(range, limit ?? defaultLimit, cursor)
        const items = page.items.map(cell => ({ row: cell.row, col: cell.col }))
        return page.completeness === 'complete'
          ? { items, completeness: 'complete', missing: [], nextCursor: page.nextCursor }
          : { items, completeness: page.completeness, missing: [{ ...range }], nextCursor: page.nextCursor }
      },
    }
  }
}

type Side = 'bottom' | 'top' | 'right' | 'left'
const SIDES: readonly Side[] = ['bottom', 'top', 'right', 'left']
type RegionProbe =
  | { kind: 'region'; rect: SourceRect }
  | { kind: 'blank' }
  | { kind: 'incomplete'; missing: readonly SourceRect[] }
/** Excel current region by bounded ring growth. Row sides probe the adjacent
 * row over [c0-1, c1+1] (corners included); column sides probe the adjacent
 * column over [r0, r1]. Growth to the fixpoint only ever asks for the ring,
 * never the band: a side gallops through dense stretches with one page per
 * strip (only a contiguous run of occupied lines from the edge is trusted, so
 * correctness never depends on page order), falls back to a limit-1 line probe,
 * and remembers the span it verified blank so a later growth on the other axis
 * probes only the newly exposed sub-spans. Working state is O(strip), not
 * O(band); work is O(region lines) pages. Unknown coverage is never a boundary. */
async function probeCurrentRegion(seed: SourceCell, scanner: Scanner, pageLimit: number): Promise<RegionProbe> {
  const seedRect = { r0: seed.row, r1: seed.row, c0: seed.col, c1: seed.col }
  const at = await scanner.exists(seedRect)
  if (at === 'blank') return { kind: 'blank' }
  if (at === 'unknown') return { kind: 'incomplete', missing: [seedRect] }
  const box = { ...seedRect }
  const cap = Math.min(STRIP_MAX_LINES, pageLimit)
  const gallop: Record<Side, number> = { bottom: 1, top: 1, right: 1, left: 1 }
  const verified: Record<Side, { line: number; lo: number; hi: number } | null> = { bottom: null, top: null, right: null, left: null }
  const rowSide = (side: Side) => side === 'bottom' || side === 'top'
  const outward = (side: Side): 1 | -1 => (side === 'bottom' || side === 'right' ? 1 : -1)
  const edgeLine = (side: Side) => side === 'bottom' ? box.r1 + 1 : side === 'top' ? box.r0 - 1 : side === 'right' ? box.c1 + 1 : box.c0 - 1
  const lineCount = (side: Side) => (rowSide(side) ? MAX_ROWS : MAX_COLS)
  /** Rect over lines [a, b] (either order) of a side, restricted to span [lo, hi]. */
  const strip = (side: Side, a: number, b: number, lo: number, hi: number): SourceRect => rowSide(side)
    ? { r0: Math.min(a, b), r1: Math.max(a, b), c0: lo, c1: hi }
    : { r0: lo, r1: hi, c0: Math.min(a, b), c1: Math.max(a, b) }
  const grow = (side: Side, lines: number) => {
    if (side === 'bottom') box.r1 += lines
    else if (side === 'top') box.r0 -= lines
    else if (side === 'right') box.c1 += lines
    else box.c0 -= lines
  }
  /** Probe one side once: 'grew', 'closed' (verified blank), or a missing rect. */
  const probeSide = async (side: Side): Promise<'grew' | 'closed' | SourceRect> => {
    const line = edgeLine(side)
    if (line < 0 || line >= lineCount(side)) return 'closed'
    const lo = rowSide(side) ? Math.max(0, box.c0 - 1) : box.r0
    const hi = rowSide(side) ? Math.min(MAX_COLS - 1, box.c1 + 1) : box.r1
    const known = verified[side]
    if (known !== null && known.line === line) {
      for (const [a, b] of [[lo, known.lo - 1], [known.hi + 1, hi]]) {
        if (a > b) continue
        const rect = strip(side, line, line, a, b)
        const answer = await scanner.exists(rect)
        if (answer === 'unknown') return rect
        if (answer === 'occupied') {
          grow(side, 1)
          verified[side] = null
          return 'grew'
        }
      }
      verified[side] = { line, lo: Math.min(lo, known.lo), hi: Math.max(hi, known.hi) }
      return 'closed'
    }
    const dir = outward(side)
    const available = dir === 1 ? lineCount(side) - line : line + 1
    const lines = Math.min(gallop[side], available)
    if (lines > 1) {
      // One page over a strip of lines; seen[] marks occupied lines by offset.
      const rect = strip(side, line, line + dir * (lines - 1), lo, hi)
      const page = await scanner.page(rect, lines)
      const seen = new Uint8Array(lines)
      for (const cell of page.items) seen[dir * ((rowSide(side) ? cell.row : cell.col) - line)] = 1
      let prefix = 0
      while (prefix < lines && seen[prefix] === 1) prefix++
      const exhausted = page.nextCursor === undefined && page.completeness === 'complete'
      if (prefix > 0) {
        grow(side, prefix)
        gallop[side] = prefix === lines ? Math.min(lines * 2, cap) : 1
        verified[side] = exhausted && prefix < lines ? { line: line + dir * prefix, lo, hi } : null
        return 'grew'
      }
      gallop[side] = 1
      if (exhausted) {
        verified[side] = { line, lo, hi }
        return 'closed'
      }
    }
    const rect = strip(side, line, line, lo, hi)
    const answer = await scanner.exists(rect)
    if (answer === 'unknown') return rect
    if (answer === 'occupied') {
      grow(side, 1)
      gallop[side] = Math.min(gallop[side] * 2, cap)
      verified[side] = null
      return 'grew'
    }
    verified[side] = { line, lo, hi }
    return 'closed'
  }
  for (;;) {
    let grew = false
    for (const side of SIDES) {
      const outcome = await probeSide(side)
      if (typeof outcome === 'object') return { kind: 'incomplete', missing: [outcome] }
      if (outcome === 'grew') grew = true
    }
    // A full pass without growth verified every side of the unchanged box.
    if (!grew) return { kind: 'region', rect: { ...box } }
  }
}

type HeaderProbe =
  | { kind: 'counts'; headerRows: 0 | 1 }
  | { kind: 'incomplete'; missing: readonly SourceRect[] }
  | { kind: 'refused'; refusal: OperationRefusal }
/** Non-table header heuristic (spec §6.1): the first row holds text and no
 * non-text while the next body row holds at least one non-text value. Blank
 * cells are ignored; errors count as non-text. Formula cells without a value
 * are permanently unavailable (§4.4): they refuse with a count and the first
 * address instead of a wait that can never resolve. A pending page waits
 * first, so the reported count and address are exact once coverage is full. */
async function probeHeaderRows(store: SourceStore, options: ReadOptions, pacer: Pacer, rect: SourceRect): Promise<HeaderProbe> {
  const band: SourceRect = { r0: rect.r0, r1: Math.min(rect.r0 + 1, MAX_ROWS - 1), c0: rect.c0, c1: rect.c1 }
  let row0Text = false, row0NonText = false, row1NonText = false
  let pending = false
  let unavailable = 0
  let firstUnavailable: SourceCell | null = null
  let cursor: string | undefined
  const seen = new Set<string>()
  for (;;) {
    await pacer.tick()
    const page = await store.readRows(band, cursor, options)
    options.signal.throwIfAborted()
    if (page.error !== undefined) {
      return { kind: 'refused', refusal: { reason: 'source-error', code: page.error.code, retryable: page.error.retryable, message: page.error.message } }
    }
    if (!sameStamp(page.stamp, options.stamp)) {
      return { kind: 'refused', refusal: { reason: 'source-error', code: 'stale-source', retryable: true, message: 'Row page belongs to another source stamp' } }
    }
    if (page.completeness !== 'complete') pending = true
    for (const item of page.items) {
      if (item.row < band.r0 || item.row > band.r1) throw contractViolation('Read row outside the requested band')
      for (const cell of item.cells) {
        if (!isValidCoordinate(cell.row, cell.col)) throw contractViolation('Read cell outside the worksheet domain')
        if (cell.row < band.r0 || cell.row > band.r1 || cell.col < band.c0 || cell.col > band.c1) {
          throw contractViolation('Read cell outside the requested band')
        }
        if (cell.unavailable) {
          unavailable++
          if (firstUnavailable === null || cell.row < firstUnavailable.row
            || (cell.row === firstUnavailable.row && cell.col < firstUnavailable.col)) {
            firstUnavailable = { row: cell.row, col: cell.col }
          }
          continue
        }
        const isText = !cell.valueIsError && typeof cell.value === 'string'
        const isNonText = cell.valueIsError || typeof cell.value === 'number' || typeof cell.value === 'boolean'
        if (item.row === band.r0) {
          row0Text = row0Text || isText
          row0NonText = row0NonText || isNonText
        } else if (item.row === band.r1) {
          row1NonText = row1NonText || isNonText
        }
      }
    }
    if (page.nextCursor === undefined) break
    if (seen.has(page.nextCursor)) throw contractViolation('Read cursor made no progress')
    seen.add(page.nextCursor)
    cursor = page.nextCursor
  }
  if (pending) return { kind: 'incomplete', missing: [{ ...band }] }
  if (firstUnavailable !== null) {
    return { kind: 'refused', refusal: { reason: 'formula-value-unavailable', count: unavailable, firstAddress: firstUnavailable } }
  }
  return { kind: 'counts', headerRows: row0Text && !row0NonText && row1NonText ? 1 : 0 }
}

interface DetectionContext {
  request: RangeDetectionRequest
  pacer: Pacer
  /** Source owner captured before the first await; the result is stamped with it. */
  owner: SourceStamp
}
/** Shared tail: resolve protected header/totals first, then refuse child-only
 * slices over the body band only (protected rows never move, §6.3), then the
 * frozen range with its provenance. */
async function finishRect(
  context: DetectionContext,
  rect: SourceRect,
  source: RangeSource,
  table?: { headerRows: 0 | 1; totalsRows: 0 | 1 },
): Promise<RangeDetection> {
  const { request } = context
  const override = request.override ?? null
  let headerRows: 0 | 1, totalsRows: 0 | 1, header: HeaderProvenance
  if (table !== undefined) {
    headerRows = table.headerRows
    totalsRows = table.totalsRows
    header = 'metadata'
  } else if (override !== null) {
    headerRows = override.headerRows
    totalsRows = 0
    header = 'override'
  } else if (rect.r0 === rect.r1) {
    headerRows = 0
    totalsRows = 0
    header = 'heuristic'
  } else {
    const probe = await probeHeaderRows(request.store, request.options, context.pacer, rect)
    if (probe.kind === 'incomplete') return { kind: 'incomplete', missing: probe.missing }
    if (probe.kind === 'refused') return refused(probe.refusal)
    headerRows = probe.headerRows
    totalsRows = 0
    header = 'heuristic'
  }
  const body = { r0: rect.r0 + headerRows, r1: rect.r1 - totalsRows }
  if (body.r0 <= body.r1) {
    const childOnly = isChildOnlyRange({ r0: body.r0, r1: body.r1, c0: rect.c0, c1: rect.c1 }, request.structure)
    if (childOnly !== null) return translateRefusal(childOnly)
  }
  return {
    kind: 'complete',
    range: { r0: rect.r0, r1: rect.r1, c0: rect.c0, c1: rect.c1, headerRows, totalsRows },
    stamp: context.owner,
    source,
    header,
    overrideIgnored: table !== undefined && override !== null,
  }
}
/** Membership of one selection axis: a single inclusive run, several runs
 * (disjoint), malformed input, or a worker-resident ref this pure module
 * cannot resolve. */
type AxisRun = { lo: number; hi: number } | 'disjoint' | 'invalid' | 'unsupported'
/** Ascending, non-overlapping inclusive spans; adjacent spans merge, so split
 * encodings of one contiguous run are the same run. */
function spanRun(spans: readonly { lo: number; hi: number }[], limit: number): Exclude<AxisRun, 'unsupported'> {
  if (spans.length === 0) return 'invalid'
  let disjoint = false
  let lo = -1, hi = -2
  for (const span of spans) {
    if (!Number.isSafeInteger(span.lo) || !Number.isSafeInteger(span.hi)
      || span.lo < 0 || span.lo > span.hi || span.hi >= limit || span.lo <= hi) return 'invalid'
    if (lo < 0) lo = span.lo
    else if (span.lo !== hi + 1) disjoint = true
    hi = span.hi
  }
  return disjoint ? 'disjoint' : { lo, hi }
}
/** Normalizes every CompactSourceRowSet encoding (§4.1) to the same run, so
 * equivalent intervals, lists and bitsets resolve identically (§6.1). */
function rowRun(rows: CompactSourceRowSet): AxisRun {
  if (rows.kind === 'ref') return 'unsupported'
  if (rows.kind === 'intervals') return spanRun(rows.intervals.map(span => ({ lo: span.r0, hi: span.r1 })), MAX_ROWS)
  if (!isValidCoordinate(rows.r0, 0) || !isValidCoordinate(rows.r1, 0) || rows.r0 > rows.r1) return 'invalid'
  if (rows.kind === 'list') {
    if (!(rows.rows instanceof Uint32Array) || rows.rows.length === 0) return 'invalid'
    let previous = -1
    for (const row of rows.rows) {
      if (row < rows.r0 || row > rows.r1 || row <= previous) return 'invalid'
      previous = row
    }
    const first = rows.rows[0], last = rows.rows[rows.rows.length - 1]
    return last - first === rows.rows.length - 1 ? { lo: first, hi: last } : 'disjoint'
  }
  // Bitset: bit i (LSB-first per word) is row r0+i; padding must be zero and
  // count must equal the popcount (shared encoding, §4.1). Mask-word bit tests
  // only; cell IDs are never bit-packed here.
  const span = rows.r1 - rows.r0 + 1
  if (rows.space !== 'source' || !(rows.words instanceof Uint32Array) || rows.words.length !== Math.ceil(span / 32)) return 'invalid'
  let count = 0, first = -1, last = -1
  for (let index = 0; index < rows.words.length; index++) {
    const word = rows.words[index]
    if (word === 0) continue
    for (let bit = 0; bit < 32; bit++) {
      if (((word >>> bit) & 1) === 0) continue
      const offset = index * 32 + bit
      if (offset >= span) return 'invalid'
      if (first < 0) first = offset
      last = offset
      count++
    }
  }
  if (count === 0 || count !== rows.count) return 'invalid'
  return count === last - first + 1 ? { lo: rows.r0 + first, hi: rows.r0 + last } : 'disjoint'
}
/** A single contiguous multi-cell selection containing the active cell, after
 * normalizing its encoding. Full rows/columns keep region detection from the
 * active cell (spec §6.1). Malformed membership is refused, a ref is
 * unsupported here, and disjoint membership seeds from the active cell. */
function explicitRect(selection: GridSelection, active: SourceCell): SourceRect | null | 'invalid' | 'unsupported' {
  if (selection.ranges.length !== 1) return null
  const part = selection.ranges[0]
  const rows = rowRun(part.rows)
  const cols = spanRun(part.cols.map(span => ({ lo: span.c0, hi: span.c1 })), MAX_COLS)
  if (rows === 'invalid' || cols === 'invalid') return 'invalid'
  if (rows === 'unsupported') return 'unsupported'
  if (rows === 'disjoint' || cols === 'disjoint') return null
  const rect = { r0: rows.lo, r1: rows.hi, c0: cols.lo, c1: cols.hi }
  if (rect.r0 === rect.r1 && rect.c0 === rect.c1) return null
  if ((rect.r0 === 0 && rect.r1 === MAX_ROWS - 1) || (rect.c0 === 0 && rect.c1 === MAX_COLS - 1)) return null
  if (active.row < rect.r0 || active.row > rect.r1 || active.col < rect.c0 || active.col > rect.c1) return null
  return rect
}
function finishTable(context: DetectionContext, table: TableExtent): Promise<RangeDetection> | RangeDetection {
  if (!isValidSourceRectInput(table.range)
    || (table.headerRows !== 0 && table.headerRows !== 1)
    || (table.totalsRows !== 0 && table.totalsRows !== 1)) {
    return refused({ reason: 'invalid-range', message: 'Declared table range is outside the worksheet domain' })
  }
  return finishRect(context, table.range, { kind: 'table', id: table.id }, table)
}
function sameRect(a: SourceRect, b: SourceRect): boolean {
  return a.r0 === b.r0 && a.r1 === b.r1 && a.c0 === b.c0 && a.c1 === b.c1
}
/** Table-overlap admission for ranges not resolved from a containing table:
 * any intersecting table other than an exact match refuses; an exact match is
 * the table itself and keeps its metadata. */
function finishDerived(context: DetectionContext, rect: SourceRect, source: RangeSource): Promise<RangeDetection> | RangeDetection {
  const structure = context.request.structure
  const other = structure.firstWhere(rect, hit => hit.kind === 'table' && !sameRect(hit.record.range, rect))
  if (other !== undefined && other.kind === 'table') {
    return refused({ reason: 'ambiguous-selection', message: `Derived range overlaps table ${other.record.id} without matching it` })
  }
  const exact = structure.firstWhere(rect, hit => hit.kind === 'table')
  if (exact !== undefined && exact.kind === 'table') return finishTable(context, exact.record)
  return finishRect(context, rect, source)
}
/** used minus region: up to four rectangles, clipped where the region overhangs. */
function outsideRegion(used: SourceRect, region: SourceRect): SourceRect[] {
  const out: SourceRect[] = []
  const r0 = Math.max(used.r0, region.r0), r1 = Math.min(used.r1, region.r1)
  if (r0 > r1 || used.c1 < region.c0 || used.c0 > region.c1) return [{ ...used }]
  if (used.r0 < r0) out.push({ r0: used.r0, r1: r0 - 1, c0: used.c0, c1: used.c1 })
  if (r1 < used.r1) out.push({ r0: r1 + 1, r1: used.r1, c0: used.c0, c1: used.c1 })
  if (used.c0 < region.c0) out.push({ r0, r1, c0: used.c0, c1: region.c0 - 1 })
  if (region.c1 < used.c1) out.push({ r0, r1, c0: region.c1 + 1, c1: used.c1 })
  return out
}
async function detectRange(request: RangeDetectionRequest, pacer: Pacer, owner: SourceStamp): Promise<RangeDetection> {
  const { store, structure, selection, options } = request
  if (request.pageLimit !== undefined && (!Number.isSafeInteger(request.pageLimit) || request.pageLimit <= 0)) {
    return refused({ reason: 'invalid-range', message: 'Invalid chunk page limit' })
  }
  if (request.workBudget !== undefined && (!Number.isSafeInteger(request.workBudget) || request.workBudget <= 0)) {
    return refused({ reason: 'invalid-range', message: 'Invalid detection work budget' })
  }
  const override = request.override
  if (override !== undefined && override !== null
    && ((override.headerRows !== 0 && override.headerRows !== 1)
      || (override.totalsRows !== undefined && override.totalsRows !== 0))) {
    return refused({ reason: 'invalid-rule', message: 'Only a 0/1 header override is supported' })
  }
  const active = selection.active
  if (active === null) {
    return refused({ reason: 'ambiguous-selection', message: 'No active cell to seed range detection' })
  }
  if (!isValidCoordinate(active.row, active.col)) {
    return refused({ reason: 'invalid-range', message: 'Active cell is outside the worksheet domain' })
  }
  options.signal.throwIfAborted()
  // Nothing may authorize a range against another document or value generation.
  if (!sameStamp(structure.stamp, owner) || !sameStamp(options.stamp, owner)) {
    return sourceError({ code: 'stale-source', retryable: true, message: 'Structure, store and read stamps disagree' })
  }
  if (structure.completeness !== 'complete') return structuralWait(structure)
  const pageLimit = request.pageLimit ?? 4096
  const context: DetectionContext = { request, pacer, owner }
  const scanner = new Scanner(store, options, pacer, request.workBudget ?? DEFAULT_DETECTION_WORK_BUDGET)
  const tableHit = structure.firstWhere(
    { r0: active.row, r1: active.row, c0: active.col, c1: active.col },
    hit => hit.kind === 'table',
  )
  if (tableHit !== undefined && tableHit.kind === 'table') return finishTable(context, tableHit.record)
  const explicit = explicitRect(selection, active)
  if (explicit === 'invalid') return refused({ reason: 'invalid-range', message: 'Malformed explicit selection range' })
  if (explicit === 'unsupported') return refused({ reason: 'unsupported', message: 'Row-set refs are resolved by the worker, not range detection' })
  if (explicit !== null) return finishDerived(context, explicit, { kind: 'explicit' })
  const region = await probeCurrentRegion(active, scanner, pageLimit)
  if (region.kind === 'incomplete') return { kind: 'incomplete', missing: region.missing }
  if (region.kind === 'region') return finishDerived(context, region.rect, { kind: 'region' })
  // Blank seed: the used range applies only when it holds exactly one current
  // region; two blocks around a blank cell must never become one range (§6.1).
  // The hint is optional: without one, the same limit-1 probes run over the
  // whole sheet (bounded by the work budget), so a proven-empty source refuses
  // instead of waiting on a hint that may never arrive.
  const used = request.usedRange ?? store.extent.used ?? FULL_DOMAIN
  if (!isValidSourceRectInput(used)) {
    return refused({ reason: 'invalid-range', message: 'Used range is outside the worksheet domain' })
  }
  const first = await scanner.first(used)
  if (first === 'unknown') return { kind: 'incomplete', missing: [{ ...used }] }
  if (first === 'blank') {
    return refused({ reason: 'ambiguous-selection', message: 'Blank active cell and no data in the used range' })
  }
  const sole = await probeCurrentRegion(first, scanner, pageLimit)
  if (sole.kind === 'incomplete') return { kind: 'incomplete', missing: sole.missing }
  if (sole.kind === 'blank') throw contractViolation('Populated cell vanished during detection')
  for (const rect of outsideRegion(used, sole.rect)) {
    const answer = await scanner.exists(rect)
    if (answer === 'unknown') return { kind: 'incomplete', missing: [rect] }
    if (answer === 'occupied') {
      return refused({ reason: 'ambiguous-selection', message: 'Blank active cell and the used range holds more than one region' })
    }
  }
  return finishDerived(context, sole.rect, { kind: 'used' })
}
/** Typed boundary: cancellation, scan failures (including contract-violating
 * pages, as page-error) and budget exhaustion become outcomes; anything else
 * is a programming error and throws.
 * The source owner is captured before any await and revalidated after the
 * last one (spec §10.1): an outcome computed while the store moved to another
 * generation is about old data, so it refuses as stale instead of being
 * attributed to the new owner. An abort that landed during the final await
 * still cancels. */
async function boundary(
  request: RangeDetectionRequest, run: (pacer: Pacer, owner: SourceStamp) => Promise<RangeDetection>,
): Promise<RangeDetection> {
  const signal = request.options.signal
  const owner: SourceStamp = Object.freeze({ ...request.store.stamp })
  let outcome: RangeDetection
  try {
    outcome = await run(new Pacer(signal), owner)
  } catch (error) {
    if (signal?.aborted) outcome = { kind: 'cancelled' }
    else if (error instanceof ScanFailure) outcome = sourceError(error.failure)
    else if (error instanceof WorkBudgetExceeded) outcome = refused({ reason: 'resource-limit', message: error.message })
    else throw error
  }
  if (signal?.aborted) outcome = { kind: 'cancelled' }
  else if (!sameStamp(request.store.stamp, owner)) {
    outcome = sourceError({ code: 'stale-source', retryable: true, message: 'Source generation changed during detection' })
  }
  return frozenCopy(outcome)
}
/** Resolve the operation range: table metadata, explicit selection, current
 * region, or the single-region used fallback, with header protection and
 * child-only refusal. Never throws for domain inputs or source failures. */
export async function detectOperationRange(request: RangeDetectionRequest): Promise<RangeDetection> {
  return boundary(request, (pacer, owner) => detectRange(request, pacer, owner))
}
/** Detect the range, then run the shared structural sort guard over it. A
 * refusal or wait propagates unchanged; a complete outcome means the range is
 * structurally admitted. Body value availability is checked by the S7.2 sort. */
export async function admitSortRange(request: RangeDetectionRequest): Promise<RangeDetection> {
  return boundary(request, async (pacer, owner) => {
    const detected = await detectRange(request, pacer, owner)
    if (detected.kind !== 'complete') return detected
    const pageLimit = request.pageLimit ?? 4096
    // The guard streams through its own unbudgeted scanner so S3's bounded
    // outside-column search keeps typed failures and cooperative pacing.
    const guardScanner = new Scanner(request.store, request.options, pacer, Number.POSITIVE_INFINITY)
    const guard = await guardSort(detected.range, request.structure, guardScanner.asPopulatedSource(pageLimit), { pageLimit: request.pageLimit })
    if (!guard.ok) return translateRefusal(guard.refusal)
    return detected
  })
}
