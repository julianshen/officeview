import { CF_PAINT_BATCH_CELLS } from '../cfLoader'
import type { AllocationContext, AllocationReservation } from '../view/allocationAccount'
import { measurePayloadBytes } from '../view/payloadBytes'
import { measureCfCellBytes } from '../view/contracts'
import type { CfDiagnostic, CfLease, CfOutcome } from './model'

/** Upper bound of the wire bytes one scale visual adds to a cell state, DERIVED from the wire measurer so the grant
 * estimate cannot drift from `measureCfCellBytes` (it includes the optional `fraction`). */
export const CF_SCALE_CELL_BYTES: number = measureCfCellBytes({ properties: {}, scale: { ruleId: 0, argb: 0, fraction: 0 } }) - measureCfCellBytes({ properties: {} })

/** Upper bound of the wire bytes one bar descriptor adds to a cell state (Ruling S3-S), DERIVED from the wire measurer
 * with every optional field present, so the grant estimate cannot drift from `measureCfCellBytes` (Lesson 6). */
export const CF_BAR_CELL_BYTES: number = measureCfCellBytes({ properties: {}, bar: {
  ruleId: 0, axisFraction: 0, startFraction: 0, endFraction: 0, fillColor: 0, valueFraction: 0, borderColor: 0, axisColor: 0,
  isGradient: false, isNegative: false, showValue: false, rtl: false,
} }) - measureCfCellBytes({ properties: {} })

/** The wire bytes one icon descriptor adds to a cell state (Ruling S4-Q), DERIVED from the wire measurer so the grant
 * estimate cannot drift from `measureCfCellBytes` (Lesson 6). The descriptor has no optional field. */
export const CF_ICON_CELL_BYTES: number = measureCfCellBytes({ properties: {}, icon: {
  ruleId: 0, set: 0, iconId: 0, bucket: 0, buckets: 0, showValue: false,
} }) - measureCfCellBytes({ properties: {} })

/** Upper bound of one normalized data-bar definition's own fields (the rule origin is priced by the rule bound),
 * DERIVED from the payload measurer with every optional field present and the longest enum spellings. */
export const CF_BAR_DEFINITION_BYTES: number = measurePayloadBytes({
  kind: 'dataBar', bounds: [{ type: 'percentile', value: 0 }, { type: 'percentile', value: 0 }],
  fillArgb: 0, negativeFillArgb: 0, borderArgb: 0, negativeBorderArgb: 0, axisArgb: 0, minLength: 0, maxLength: 0,
  gradient: false, border: false, negativeBarColorSameAsPositive: false, negativeBarBorderColorSameAsPositive: false,
  axisPosition: 'automatic', direction: 'rightToLeft', showValue: false, sheetRightToLeft: false, pairing: 'base-only', autoColors: ['negativeFill', 'axis'],
})!

/** Upper bound of one normalized icon-set definition's own fields (the rule origin is priced by the rule bound),
 * DERIVED from the payload measurer for the largest shape: five thresholds with the longest bound spelling, five custom
 * refs and the longest set names (Lesson 6). */
export const CF_ICON_DEFINITION_BYTES: number = measurePayloadBytes({
  kind: 'iconSet', set: '3TrafficLights1',
  thresholds: Array.from({ length: 5 }, () => ({ bound: { type: 'percentile', value: 0 }, gte: false })),
  reverse: false, showValue: false, percentAttr: false,
  custom: Array.from({ length: 5 }, () => ({ set: '3TrafficLights1', iconId: 0 })), sheetRightToLeft: false,
})!

/** Upper bound of one compiled formula AST node's own fields (Ruling S5-B, round-1 #3), DERIVED from the payload measurer
 * over the largest node shapes (a range qualified by a 31-character sheet name is the widest): the compile-time
 * reservation is `tokens x CF_FORMULA_NODE_BYTES` and a token never yields more than one node's own fields. Children are
 * priced as their own nodes, so each shape is measured with its child slots empty. */
const CF_NODE_SAMPLE_SHEET = 'S'.repeat(31)
const CF_NODE_SAMPLE_CELL = { sheet: CF_NODE_SAMPLE_SHEET, col: 16383, row: 1048575, absCol: true, absRow: true }
export const CF_FORMULA_NODE_BYTES: number = Math.max(...[
  { type: 'cell', ref: CF_NODE_SAMPLE_CELL },
  { type: 'range', ref: { sheet: CF_NODE_SAMPLE_SHEET, from: CF_NODE_SAMPLE_CELL, to: CF_NODE_SAMPLE_CELL } },
  { type: 'wholeCol', ref: { sheet: CF_NODE_SAMPLE_SHEET, from: { index: 16383, absolute: true }, to: { index: 16383, absolute: true } } },
  { type: 'call', name: 'AVERAGEIFS', args: [null] },
  { type: 'binary', op: '<=', left: null, right: null },
  { type: 'unary', op: '-', expr: null },
  { type: 'number', value: -1.7976931348623157e308 },
  { type: 'string', value: '' },
].map(node => measurePayloadBytes(node)!))
/** Bytes of one lexer token's own fields at their widest (every optional field present, strings empty: string payloads are
 * priced per code unit by the compile bound). The lexer builds one token object per token before the AST exists. */
const CF_TOKEN_SAMPLE_CELL = { sheet: '', col: 16383, row: 1048575, absCol: true, absRow: true }
export const CF_FORMULA_TOKEN_BYTES: number = measurePayloadBytes({
  type: 'cell', value: '', start: 8191, numValue: -1.7976931348623157e308, sheet: '', cellRef: CF_TOKEN_SAMPLE_CELL,
  rangeRef: { sheet: '', from: CF_TOKEN_SAMPLE_CELL, to: CF_TOKEN_SAMPLE_CELL }, spaceBefore: true, raw: '',
})!
/** One reference atom at its widest (an empty sheet name; its characters are priced per code unit). */
export const CF_FORMULA_ATOM_BYTES: number = measurePayloadBytes({ sheet: '', firstRow: 1048575, lastRow: 1048575, firstCol: 16383, lastCol: 16383, absRow: true, absCol: true, axis: 'col' })!
/** The compiled formula wrapper at its widest: every family, every truth kind, a full dependency list and an anchor. */
export const CF_FORMULA_FIXED_BYTES: number = measurePayloadBytes({
  ast: null, nodes: 4096, atoms: [], invariant: true, usesClock: true,
  families: ['aggregate', 'conditional-aggregate', 'date', 'information', 'logical', 'math', 'row-column', 'text', 'today'],
  dependencySheets: Array.from({ length: 16 }, () => ''), anchor: { row: 1048575, col: 16383 }, pinned: true, rootArea: true,
  truth: ['blank', 'boolean', 'error', 'number', 'range', 'text'], nonAsciiLiterals: [],
})!
/** One element of a character array the evaluator builds for a text function (`textChars`): a one-code-unit string slot. */
export const CF_CHAR_SLOT_BYTES: number = measurePayloadBytes(['a'])! - measurePayloadBytes([])!

/* Stage 5 Task 34: dependency snapshot and evaluator scratch prices (Ruling S5-N, Lessons 6 and 15). Each is DERIVED from
 * the payload measurer over the widest value the structure holds; string payloads are priced separately at 2 B per code
 * unit (the measurer's string rule), never hand-counted. */
const WIDEST_ERROR = Object.freeze({ kind: 'formula-error', code: '#DIV/0!' })
/** One snapshot entry: a row slot in its column's sorted row list, a value slot holding the widest non-string value (a
 * tagged error object; a string's characters are priced per unit on top), its string-unit and kind/ASCII flag slots, and its slot in the
 * snapshot's row-major key index (fix round 2, finding 1). */
export const CF_DEPENDENCY_CELL_BYTES: number = measurePayloadBytes([[1048575], [WIDEST_ERROR], [32767], [2], [17179869183]])! - measurePayloadBytes([[], [], [], [], []])!
/** One slot of the row-major key index alone (a rebuild after an out-of-order source holds old plus new). */
export const CF_DEPENDENCY_KEY_BYTES: number = measurePayloadBytes([17179869183])! - measurePayloadBytes([])!
/** One memo or map entry (key and value slots): the evaluator's per-node memo and the CF-owned invariant memo. */
export const CF_SLOT_BYTES: number = measurePayloadBytes(new Map([[0, 0]]))! - measurePayloadBytes(new Map())!
/** One populated slot `{ col, row }` the engine's sparse cursor builds per matched cell (refs.ts `populated`). */
export const CF_POPULATED_SLOT_BYTES: number = measurePayloadBytes([{ col: 16383, row: 1048575 }])! - measurePayloadBytes([])!
/** One flattened argument item `{ value, fromRef, omitted }` (functions.ts `FlatArgItem`) at its widest value. */
export const CF_FLAT_ITEM_BYTES: number = measurePayloadBytes([{ value: WIDEST_ERROR, fromRef: true, omitted: true }])! - measurePayloadBytes([])!
/** One key of a per-call view's sorted key list. */
export const CF_VIEW_KEY_BYTES: number = measurePayloadBytes([0])! - measurePayloadBytes([])!
/** One frequency-map entry of a COUNTIF hoist (an empty key string; its characters are priced per unit on top). */
export const CF_FREQUENCY_ENTRY_BYTES: number = measurePayloadBytes(new Map([['', 0]]))! - measurePayloadBytes(new Map())!
/** [call] Per-call engine bookkeeping the payload measurer cannot see (closures, WeakMaps, the reference-services object,
 * the evaluation context): a policy allowance, not a measurement, reviewed with the caps table (S5-N). */
export const CF_FORMULA_CALL_BYTES = 8 * 1024
/** [call] Per traversal occurrence: the engine's resolved-reference object, its cursor map and argument progress. */
export const CF_TRAVERSAL_BYTES = 1024
/** [call] The conditional matcher's typed rows at a conservative 8 B per DP cell (Ruling S5-N caps table). */
export const CF_MATCHER_CELL_BYTES = 8

/** Thrown by a CF-owned store or callback when a per-call spend is refused (Ruling S5-N): it unwinds the synchronous
 * engine call (no engine frame catches it) and the evaluation is discarded as `cf-resource-limit`. */
export const CF_WORK_REFUSED: unique symbol = Symbol('cf-work-refused')

/** One preparation's per-cell evaluation ceiling (Global Constraints "Bounded work"): every snapshot attempt and every
 * per-cell call draws from it, so fallbacks cannot multiply total work. Unlike a CfWorkBudget its refusal is not final:
 * a call too large for what is left is refused alone, and a smaller one after it may still run (S5-N: one rule's
 * refusal never refuses another). */
export interface CfUnitCeiling { readonly limit: number; readonly used: number; readonly remaining: number; take(units: number): boolean }
export function createCfUnitCeiling(limit: number = CF_LIMITS.formulaCeilingUnits): CfUnitCeiling {
  if (!validBytes(limit)) throw new RangeError('Invalid CF unit ceiling')
  let used = 0
  return {
    limit, get used() { return used }, get remaining() { return limit - used },
    take(units) { if (!validBytes(units) || units > limit - used) return false; used += units; return true },
  }
}

/** Ruling S3-I1-UNITS: one preparation call's own work units (calibrated in docs/validation/xlsx-cf.md section 11).
 * Stage 5 Task 34 recalibration: 250,000 -> 300,000, so it stays >= 5x the largest production call with formula rules
 * present (55,011 units: 1,000 cells, four formula rules, one COUNTIF($A:$A,...); 5.45x). */
const ACTIVE_UNITS = 300_000
export const CF_LIMITS = Object.freeze({ rules: 4096, rangesPerRule: 256, rangesPerSheet: 16384,
  sheetMetadataBytes: 4 * 1024 ** 2, workbookMetadataBytes: 16 * 1024 ** 2,
  textUnits: 8192, formulaTokens: 4096, formulaChildren: 3, cfvo: 5, customIcons: 5,
  // Stage 5 (Ruling S5-B caps table): reference atoms per compiled formula and distinct sheets one rule may reference.
  formulaAtoms: 64, referencedSheets: 16,
  // Stage 5 Task 34 (Ruling S5-N caps table): the dependency snapshot's cells per preparation; one MaterializationBudget
  // per preparation (array materialization only); the per-cell evaluation ceiling every snapshot attempt and every call
  // of one preparation draws from (2 x activeUnits, Global Constraints "Bounded work"); and the per-call evaluator
  // scratch cap (S5-B round-2 #1, 64 KiB).
  dependencyCells: 262_144, materializationCells: 1_000_000, formulaCeilingUnits: 2 * ACTIVE_UNITS, formulaScratchBytes: 64 * 1024,
  statisticsBytes: 8 * 1024 ** 2, scratchBytes: 8 * 1024 ** 2, workUnits: 1_000_000,
  // Ruling S3-I1-UNITS: one preparation call's own work, in deterministic units (never wall time). Calibrated in
  // docs/validation/xlsx-cf.md section 11 (scripts/cf-refs/active-units.ts): >= 5x the largest production call.
  activeUnits: ACTIVE_UNITS, checkpointUnits: 128, sliceMs: 4, preparedCells: CF_PAINT_BATCH_CELLS,
  matchingRules: 4096, descriptorUnits: 128, diagnostics: 64 })
/** The worst-case workspace one per-call view builds (fix round 2, finding 1): its key set and sorted key list at the
 * per-call cap, one area index per reference occurrence (at most the cap), and per distinct area (at most one per admitted
 * atom) its slot and measured metrics. DERIVED from the measurer; the evaluator reserves it with the per-call
 * bookkeeping before `viewFor`, which refuses without it. */
export const CF_VIEW_WORKSPACE_BYTES: number = (() => {
  const cap = CF_LIMITS.checkpointUnits, atoms = CF_LIMITS.formulaAtoms
  const keys = Array.from({ length: cap }, (_, i) => 17179869183 - i)
  const shared = { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }
  return measurePayloadBytes({
    keySet: new Set(keys), ordered: keys, indexOf: keys,
    distinct: Array.from({ length: atoms }, () => shared).map((_, i) => i), // slots only: the areas are the caller's
    perArea: Array.from({ length: atoms }, () => ({ cells: cap, stringUnits: 1e9, longestString: 32767, strings: cap, nonAscii: cap })),
  })!
})()
export interface CfRefusal { readonly kind: 'unresolved'; readonly reason: string }
const refuse = (reason = 'cf-budget-exceeded'): CfRefusal => ({ kind: 'unresolved', reason })
const validBytes = (n: number): boolean => Number.isSafeInteger(n) && n >= 0
export interface CfReservation {
  readonly kind: 'reserved'; readonly bytes: number
  /** Absolute new construction bound. Growth must precede the corresponding allocation. */
  grow(bytes: number): boolean
  adopt<T>(value: T, retainedBytes?: number): CfLease<T>
  release(): void
}
export interface CfMemory {
  readonly currentBytes: number; readonly constructionBytes: number; readonly retainedBytes: number
  readonly peakBytes: number; readonly closed: boolean
  reserve(bytes: number): CfReservation | CfRefusal
  /** Denies new work. Live construction/consumer leases remain charged until their explicit release. */
  close(): void
}
/** One workbook account shared by its sheets. Same adapter works with main and worker AllocationContexts.
 * Without an external account the bounded local parse account owns the leases. */
export function createCfMemory(options: { limitBytes?: number; context?: AllocationContext } = {}): CfMemory {
  const limit = options.limitBytes ?? CF_LIMITS.workbookMetadataBytes
  if (!validBytes(limit)) throw new RangeError('Invalid CF memory limit')
  let construction = 0, retained = 0, peak = 0, closed = false
  const context = options.context
  const memory: CfMemory = {
    get currentBytes() { return construction + retained }, get constructionBytes() { return construction },
    get retainedBytes() { return retained }, get peakBytes() { return peak }, get closed() { return closed || context?.closed === true },
    close() { closed = true },
    reserve(initial) {
      if (!validBytes(initial) || memory.closed || initial > limit - memory.currentBytes) return refuse()
      let account: AllocationReservation | undefined
      if (context) {
        if (!context.canReserve(initial)) return refuse()
        account = context.reserve()
        if (!account.trySet(initial)) { account.release(); return refuse() }
      }
      let bytes = initial, phase: 'construction' | 'retained' | 'released' = 'construction'
      let lease: CfLease<unknown> | undefined
      construction += bytes; peak = Math.max(peak, memory.currentBytes)
      const end = () => {
        if (phase === 'released') return
        if (phase === 'construction') construction -= bytes
        else retained -= bytes
        phase = 'released'; bytes = 0; account?.release()
      }
      const reservation: CfReservation = {
        kind: 'reserved', get bytes() { return bytes },
        grow(next) {
          if (!validBytes(next) || phase !== 'construction' || memory.closed) return false
          if (next <= bytes) return true // no premature refund of potentially live workspace
          if (next - bytes > limit - memory.currentBytes || (account && !account.trySet(next))) return false
          construction += next - bytes; bytes = next; peak = Math.max(peak, memory.currentBytes); return true
        },
        adopt<T>(value: T, retainedBytes = bytes): CfLease<T> {
          if (lease) {
            if (lease.value !== value || lease.bytes !== retainedBytes) throw new Error('CF reservation already adopted')
            return lease as CfLease<T>
          }
          if (phase !== 'construction' || memory.closed || !validBytes(retainedBytes) || retainedBytes > bytes) throw new Error('Invalid CF adoption')
          // Caller has dropped construction-only workspace before requesting this shrink.
          if (account && !account.trySet(retainedBytes)) throw new Error('CF adoption account closed')
          construction -= bytes; retained += retainedBytes; bytes = retainedBytes; phase = 'retained'
          const owned: CfLease<T> = {
            value, get bytes() { return bytes }, release: end,
            transfer(destination) {
              if (phase !== 'retained') return refuse('cf-owner-released')
              if (destination === memory) return { kind: 'ready', value: owned }
              const incoming = destination.reserve(bytes)
              if (incoming.kind === 'unresolved') return incoming
              try {
                const moved = incoming.adopt(value)
                end() // receiver admission succeeded before source ownership ends
                return { kind: 'ready', value: moved }
              } catch (error) { incoming.release(); throw error }
            },
          }
          lease = owned; return owned
        },
        release() { if (phase === 'construction') end() },
      }
      return reservation
    },
  }
  return memory
}
/** The supplied bound includes candidate + simultaneous construction workspace. Builder is synchronous;
 * async providers must keep their reservation until they drain, even after close/cancellation. */
export function buildCfCandidate<T>(memory: CfMemory, bound: number, builder: () => T): CfOutcome<CfLease<T>> {
  const reservation = memory.reserve(bound)
  if (reservation.kind === 'unresolved') return reservation
  try { return { kind: 'ready', value: reservation.adopt(builder()) } }
  catch (error) { reservation.release(); throw error }
}
/** Exact repository payload model: full distinct backing buffers, strings at 2 B/unit,
 * object/map/array slots. The measuring walk's scratch is itself admitted before it grows.
 * This measures an existing payload; use an input-derived bound BEFORE constructing candidates. */
export function measureCfPayload(value: unknown, memory: CfMemory): CfOutcome<number> {
  const hold = memory.reserve(0)
  if (hold.kind === 'unresolved') return hold
  let workspace = 0
  try {
    const bytes = measurePayloadBytes(value, extra => {
      const next = workspace + extra
      if (next > CF_LIMITS.scratchBytes || !hold.grow(next)) return false
      workspace = next; return true
    })
    return bytes === undefined ? refuse() : { kind: 'ready', value: bytes }
  } finally { hold.release() }
}

/** Walked byte size with the traversal workspace admitted on the given
 * account (grown incrementally, capped at scratchBytes): returns undefined
 * instead of measuring when the account cannot admit the walk. */
export function measureCfWalkedBytes(value: unknown, memory: CfMemory, capBytes = CF_LIMITS.scratchBytes): number | undefined {
  const hold = memory.reserve(0)
  if (hold.kind !== 'reserved') return undefined
  let workspace = 0
  try {
    const bytes = measurePayloadBytes(value, extra => {
      const next = workspace + extra
      if (next > capBytes || !hold.grow(next)) return false
      workspace = next
      return true
    })
    return bytes ?? undefined
  } finally {
    try { hold.release() } catch { /* release-only path */ }
  }
}

/** Walked byte size with only a hard traversal cap and no account (for
 * callers whose admission lives elsewhere, e.g. the session outgoing
 * measure covered by its pool charge): undefined on overflow or cap breach. */
export function measureBoundedPayload(value: unknown, capBytes: number): number | undefined {
  let workspace = 0
  const bytes = measurePayloadBytes(value, extra => {
    workspace += extra
    return workspace <= capBytes
  })
  return bytes ?? undefined
}

/** Admitted retained copy of an already-built value: measures with the
 * traversal workspace admitted, reserves the exact walked size, and adopts.
 * Each stage fails closed independently — unmeasurable walks, refused
 * reservations and failed adoptions all return undefined with nothing held
 * — so callers only publish shared ownership when the owning lease exists. */
export function admitSharedCopy<T>(memory: CfMemory, value: T): CfLease<T> | undefined {
  const walked = measureCfWalkedBytes(value, memory)
  if (walked === undefined) return undefined
  const hold = memory.reserve(walked)
  if (hold.kind !== 'reserved') return undefined
  try {
    return hold.adopt(value, walked)
  } catch {
    try { hold.release() } catch { /* release-only path */ }
    return undefined
  }
}

export function createCfDiagnostics(memory: CfMemory = createCfMemory({ limitBytes: CF_LIMITS.sheetMetadataBytes })) {
  const entries: CfDiagnostic[] = []
  let hold: CfReservation | undefined, refused = false, disposed = false
  let bytes = 0
  return {
    get entries(): readonly CfDiagnostic[] { return entries },
    get refused() { return refused },
    dispose() { disposed = true; entries.length = 0; hold?.release(); hold = undefined },
    add(value: CfDiagnostic): void {
      if (disposed || entries.length >= CF_LIMITS.diagnostics) return
      // Compare bounded prefixes without allocating duplicate strings or consuming room.
      if (entries.some(d => d.feature.length === Math.min(value.feature.length, 128) && value.feature.startsWith(d.feature) && d.message.length === Math.min(value.message.length, CF_LIMITS.textUnits) && value.message.startsWith(d.message))) return
      // Reserve the bounded record and its copied strings before constructing either.
      // 192 covers record/array slots, three string headers and the fixed kind text.
      const next = bytes + 192 + Math.min(value.feature.length, 128) * 2 + Math.min(value.message.length, CF_LIMITS.textUnits) * 2
      if (!hold) {
        const candidate = memory.reserve(next)
        if (candidate.kind === 'unresolved') { refused = true; return }
        hold = candidate
      } else if (!hold.grow(next)) { refused = true; return }
      bytes = next
      // Bounded strings and no unbounded auxiliary dedup Set after the cap.
      const diagnostic = { ...value, feature: value.feature.slice(0, 128) as CfDiagnostic['feature'], message: value.message.slice(0, CF_LIMITS.textUnits) }
      if (!entries.some(d => d.feature === diagnostic.feature && d.message === diagnostic.message)) entries.push(diagnostic)
    },
  }
}
const ready = { kind: 'ready', value: undefined } as const
export function createCfCaps(workbook: CfMemory = createCfMemory()) {
  let rules = 0, ranges = 0, metadata = 0, disposed = false
  let status: typeof ready | (CfRefusal & { scope: 'sheet'; writes: 'ALL' }) = ready
  let metadataHold: CfReservation | undefined
  const fail = () => status = { ...refuse('cf-sheet-scope-unresolved'), scope: 'sheet', writes: 'ALL' }
  return {
    get status() { return status },
    addRule(rangeCount: number) {
      if (status.kind === 'unresolved' || disposed) return fail()
      if (!validBytes(rangeCount) || rangeCount > CF_LIMITS.rangesPerRule || rules >= CF_LIMITS.rules || ranges + rangeCount > CF_LIMITS.rangesPerSheet) return fail()
      rules++; ranges += rangeCount; return ready
    },
    check(name: string, count: number) {
      const cap = CF_LIMITS[name as keyof typeof CF_LIMITS]
      return status.kind === 'ready' && !disposed && validBytes(count) && cap !== undefined && count <= cap ? ready : fail()
    },
    reserveMetadata(bytes: number): CfOutcome<void> {
      if (disposed || status.kind === 'unresolved' || !validBytes(bytes) || bytes > CF_LIMITS.sheetMetadataBytes - metadata) return fail()
      if (bytes === 0) return ready
      if (!metadataHold) {
        const hold = workbook.reserve(bytes)
        if (hold.kind === 'unresolved') return fail()
        metadataHold = hold
      } else if (!metadataHold.grow(metadata + bytes)) return fail()
      metadata += bytes; return ready
    },
    dispose() { if (!disposed) { disposed = true; metadataHold?.release(); metadataHold = undefined; metadata = 0 } },
  }
}
export interface CfWorkBudget {
  spend(units: number): 'continue' | 'yield' | 'refused'; resume(): void; readonly units: number
  /** Units a spend may still take without refusing (0 once refused or aborted). A caller that must not latch a refusal
   * (one call's admission among many) checks this first and spends only what fits. */
  readonly remaining: number
}
/** A per-call budget (the default) caps one preparation call's own work at `CF_LIMITS.activeUnits`; a `scan` budget
 * caps one whole-scan pass (statistics, bounds) at `maxUnits`, default `CF_LIMITS.workUnits`. Ruling S3-I1-UNITS:
 * refusal depends on the units spent alone, never on the clock; `now` only paces the cooperative yield (every
 * `checkpointUnits` units or `sliceMs` of active time), which changes scheduling, never the result. */
export type CfWorkBudgetOptions = { now?: () => number; signal?: AbortSignal } & ({ scan?: false } | { scan: true; maxUnits?: number })
/** Caller must stop on yield and call resume after its awaited host yield. A spend made while paused still counts
 * (and still refuses past the cap), so the unit count does not depend on when the clock yields. */
export function createCfWorkBudget(options: CfWorkBudgetOptions = {}): CfWorkBudget {
  const now = options.now ?? (() => performance.now())
  const maxUnits = options.scan ? options.maxUnits ?? CF_LIMITS.workUnits : CF_LIMITS.activeUnits
  let start = now(), units = 0, sliceUnits = 0, paused = false, refused = false
  return {
    get units() { return units },
    get remaining() { return refused || options.signal?.aborted ? 0 : maxUnits - units },
    spend(n) {
      if (refused || options.signal?.aborted || !validBytes(n) || n > maxUnits - units) { refused = true; return 'refused' }
      units += n
      if (paused) return 'yield'
      sliceUnits += n
      if (sliceUnits >= CF_LIMITS.checkpointUnits || now() - start >= CF_LIMITS.sliceMs) { paused = true; return 'yield' }
      return 'continue'
    },
    resume() { start = now(); sliceUnits = 0; paused = false },
  }
}
