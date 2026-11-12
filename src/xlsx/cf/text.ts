import type { XlsxSheet } from '../types'
import type { CfBarDescriptor, CfLease, CfPrepared, CfStamp, SourceRect as ModelRect } from './model'
import { CF_LIMITS, createCfMemory } from './limits'
import { createCfService, estimateCfPrepareBytes, type CfPaintStatistics } from './service'
import { createCfStatistics } from './statistics'
import { resolveCfCellStyle } from './dxf'
import { residentSource, storeCfSource } from './render'
import { cfRuntimeEvidence } from './predicates'
import { cfRulesVersion } from '../view/pagedSource'
import type { EffectiveFormatScope, FormatBudget } from '../view/formatWorkspace'
import type { AllocationContext, AllocationReservation } from '../view/allocationAccount'
import type { SheetColorContext } from '../sheetColors'
import type { ValuePolicy } from '../view/values'
import type { CfCellStateWire, SourceStore, CfReadResult } from '../view/contracts'
import { measureCfCellBytes, versionCfRules } from '../view/contracts'
import type { SourceRect } from '../view/coordinates'
import { describeCfUnavailable, describeCfWin, type CfDescriptor } from './describe'
import { barHidesText, barPaintable, paintDataBars } from './dataBarPaint'
import { PackageError } from '../../core/package-error'

const refused = (message: string): PackageError => new PackageError('resource-limit', message)

/** Bounded per-batch effective-format preparation for derived text (Task 9). */
export interface CfTextScopeRequest {
  /** Exact-cover source rects (see coverCellsToRects); chunked internally. */
  readonly rects: readonly SourceRect[]
  readonly signal: AbortSignal
  readonly locale?: string
  readonly budget: FormatBudget
  readonly colors?: SheetColorContext
  readonly valuePolicy?: ValuePolicy
  readonly workbookValueGen?: number
  readonly store?: SourceStore
  readonly dateContext?: import('./predicates').CfDateContext
}

/** Fixed admitted descriptor workspace per scope with a charged budget. */
const DESCRIPTOR_WORKSPACE_BYTES = 2048
const DESCRIPTOR_CACHE_KEYS = 32
/** Mirrors the service per-call limits so every chunk is admittable. */
const CHUNK_RECTS = 256
const CHUNK_CELLS = 1000

/** Bridges a caller's FormatBudget into a CfMemory allocation context so every
 * resident byte — service workspace, shared statistics, transient text and
 * retained results — is admitted in the aggregate ledger before allocation and
 * released on disposal. Refusals fail closed through the existing quota paths. */
function budgetAllocationContext(budget: FormatBudget): AllocationContext {
  return {
    get limitBytes() { return CF_LIMITS.scratchBytes },
    reserve: () => {
      let held = 0, released = false
      const reservation: AllocationReservation = {
        owner: 'cf-text-scope',
        get bytes() { return held },
        trySet: (next: number) => {
          if (released || !Number.isSafeInteger(next) || next < 0) return false
          if (next > held) {
            try {
              budget.reserve(next - held)
            } catch {
              return false
            }
          } else if (next < held) {
            try {
              budget.release(held - next)
            } catch {
              return false
            }
          }
          held = next
          return true
        },
        charge: (amount: number) => {
          if (!Number.isSafeInteger(amount) || amount < 0) throw new Error('cf-text-scope: invalid charge')
          try {
            budget.reserve(amount)
          } catch {
            throw new Error('cf-text-scope: aggregate budget refused')
          }
          held += amount
        },
        release: (amount?: number) => {
          if (released) return
          const value = amount ?? held
          held = Math.max(0, held - value)
          try {
            budget.release(value)
          } catch { /* release-only path */ }
        },
        adopt: () => {
          throw new Error('cf-text-scope: adopt unsupported')
        },
      }
      return reservation
    },
    canReserve: () => true,
    owns: () => false,
  }
}

/** Copied winning state: the scope owns this data, charged in the budget. */
interface CopiedState {
  readonly properties: Record<string, string | number | boolean | null>
  readonly writes: readonly string[] | undefined
  readonly reason: string | undefined
  readonly scale?: Readonly<{ fraction?: number; ruleId?: number; argb?: number }>
  /** The resolved bar descriptor, identical on the resident and paged paths (Ruling S3-S). */
  readonly bar?: CfBarDescriptor
}

type CfWireBar = NonNullable<CfCellStateWire['bar']>
/** Wire form of a descriptor (`direction` as `rtl`), the shape every copy is priced in. */
const wireBar = (d: CfBarDescriptor): CfWireBar => ({
  ruleId: d.ruleId, axisFraction: d.axisFraction, startFraction: d.startFraction, endFraction: d.endFraction, fillColor: d.fillColor,
  ...(d.borderColor !== undefined ? { borderColor: d.borderColor } : {}),
  ...(d.axisColor !== undefined ? { axisColor: d.axisColor } : {}),
  isGradient: d.isGradient, isNegative: d.isNegative, showValue: d.showValue, rtl: d.direction === 'rightToLeft',
})
/** The descriptor a wire bar encodes: the copy owns plain fields only, never the received object. */
const descriptorOf = (w: CfWireBar): CfBarDescriptor => ({
  ruleId: w.ruleId, axisFraction: w.axisFraction, startFraction: w.startFraction, endFraction: w.endFraction, fillColor: w.fillColor,
  ...(w.borderColor !== undefined ? { borderColor: w.borderColor } : {}),
  isGradient: w.isGradient, isNegative: w.isNegative,
  ...(w.axisColor !== undefined ? { axisColor: w.axisColor } : {}),
  showValue: w.showValue, direction: w.rtl ? 'rightToLeft' : 'leftToRight',
})

/** Structural key for one cell's descriptor inputs: shared aspects share one charge. A painted bar keys by its
 * announced position, so the descriptor cache never hands one cell's bar text to another cell. */
const winKey = (
  properties: Readonly<Record<string, string | number | boolean | null>>,
  unresolvedWrites: readonly string[] | undefined,
  scale?: Readonly<{ fraction?: number; ruleId?: number; argb?: number }> | number,
  bar?: CfBarDescriptor,
): string => {
  const paths = Object.keys(properties).sort().join(',')
  const writes = unresolvedWrites ? [...unresolvedWrites].sort().join(',') : ''
  const fraction = typeof scale === 'number' ? scale : scale?.fraction
  const scaleKey = fraction !== undefined ? `|scale:${Math.round(fraction * 100)}` : ''
  const barKey = bar ? `|bar:${Math.round(barPosition(bar) * 100)}` : ''
  return `${paths}|${writes}${scaleKey}${barKey}`
}

/** The value's position in the bar's range: the bar end away from its anchor (axis, leading or trailing edge). */
const barPosition = (bar: CfBarDescriptor): number => bar.startFraction === bar.axisFraction ? bar.endFraction : bar.startFraction

const winDescriptor = (
  properties: Readonly<Record<string, string | number | boolean | null>>,
  unresolvedWrites: readonly string[] | undefined,
  unresolvedReason: string | undefined,
  locale: string,
  cache: Map<string, CfDescriptor>,
  scale?: Readonly<{ fraction?: number; ruleId?: number; argb?: number }> | number,
  bar?: CfBarDescriptor,
): CfDescriptor | undefined => {
  const key = winKey(properties, unresolvedWrites, scale, bar)
  const cached = cache.get(key)
  if (cached) return cached
  const descriptor = unresolvedReason !== undefined
    ? describeCfUnavailable(locale)
    : describeCfWin(properties, unresolvedWrites, undefined, locale, scale, bar ? { fraction: barPosition(bar) } : undefined)
  if (descriptor && cache.size < DESCRIPTOR_CACHE_KEYS) cache.set(key, descriptor)
  return descriptor ?? undefined
}

/** Coordinates rects must be well-formed before chunking measures them. */
function rectsAreValid(rects: readonly SourceRect[]): boolean {
  for (const rect of rects) {
    if (![rect.r0, rect.r1, rect.c0, rect.c1].every(n => Number.isSafeInteger(n) && n >= 0)) return false
    if (rect.r0 > rect.r1 || rect.c0 > rect.c1 || rect.r1 > 1048575 || rect.c1 > 16383) return false
  }
  return true
}

/** Greedy chunks honoring the service per-call limits; a single over-cap rect
 * stands alone so the service refusal (not an over-wide request) governs it. */
function chunkRects(rects: readonly SourceRect[]): SourceRect[][] {
  const chunks: SourceRect[][] = []
  let current: SourceRect[] = [], area = 0
  const flush = (): void => { if (current.length) { chunks.push(current); current = []; area = 0 } }
  for (const rect of rects) {
    const cols = rect.c1 - rect.c0 + 1
    const maxRows = Math.max(1, Math.floor(CHUNK_CELLS / Math.max(1, cols)))
    for (let r0 = rect.r0; r0 <= rect.r1; r0 += maxRows) {
      const r1 = Math.min(rect.r1, r0 + maxRows - 1)
      const slice: SourceRect = { r0, r1, c0: rect.c0, c1: rect.c1 }
      const sliceArea = (r1 - r0 + 1) * cols
      if (current.length >= CHUNK_RECTS || area + sliceArea > CHUNK_CELLS) flush()
      current.push(slice)
      area += sliceArea
    }
  }
  flush()
  return chunks
}

const toModel = (rect: SourceRect): ModelRect => ({ firstRow: rect.r0, lastRow: rect.r1, firstCol: rect.c0, lastCol: rect.c1 })

function assembleScope(
  gen: number,
  states: Map<number, CopiedState>,
  colors: SheetColorContext,
  locale: string,
  bytes: number,
  release: () => void,
  current: () => boolean = () => true,
  regions?: EffectiveFormatScope['regions'],
): EffectiveFormatScope {
  const cache = new Map<string, CfDescriptor>()
  let disposed = false
  // The bar cells of the copied states (descriptors are shared, not copied again): the painter walks only these.
  const bars = new Map<number, CfBarDescriptor>()
  for (const [key, state] of states) if (state.bar) bars.set(key, state.bar)
  const barAt = (row: number, col: number): CfBarDescriptor | undefined => {
    if (disposed || !current()) return undefined
    const bar = bars.get(row * 16384 + col)
    return barPaintable(bar) ? bar : undefined
  }
  return {
    gen,
    bytes,
    regions,
    effectiveStyle: (row, col, base) => {
      if (disposed || !current()) return base
      const state = states.get(row * 16384 + col)
      if (!state) return base
      // Surviving winning properties apply even when other aspects are
      // blocked: only the descriptor reports the unavailable aspect.
      const scale = state.scale
      const paintsScale = scale?.argb !== undefined
      if (Object.keys(state.properties).length === 0 && !paintsScale) return base
      return resolveCfCellStyle(base, {
        properties: state.properties,
        ...(paintsScale ? { visual: { colorScale: { ruleId: scale!.ruleId ?? 0, argb: scale!.argb!, ...(scale!.fraction !== undefined ? { fraction: scale!.fraction } : {}) } } } : {}),
      }, colors)
    },
    barAt,
    paintBars: (ctx, cellRect, visible, scale) => paintDataBars(ctx, bars, barAt, cellRect, visible, scale),
    hidesTextAt: (row, col, host, scale) => barHidesText(barAt(row, col), host, scale),
    describe: (row, col) => {
      if (disposed || !current()) return undefined
      const state = states.get(row * 16384 + col)
      if (!state) return undefined
      // A bar announces itself only where it paints; one the painter withholds is unavailable formatting, like a bar
      // the evaluator left unresolved (that one reaches here through its writes or reason).
      const bar = barAt(row, col)
      if (state.bar && !bar) return describeCfUnavailable(locale)
      if (Object.keys(state.properties).length === 0 && (state.writes?.length ?? 0) === 0 && state.reason === undefined && !state.scale && !bar) return undefined
      return winDescriptor(state.properties, state.writes, state.reason, locale, cache, state.scale, bar)
    },
    dispose: () => {
      if (disposed) return
      disposed = true
      states.clear()
      bars.clear()
      cache.clear()
      release()
    },
  }
}

/**
 * Resident preparation over exact-cover rects. Returns undefined (fail closed,
 * base formatting) when the sheet has no CF, any cap is exceeded, the budget
 * refuses the admitted envelope, or preparation fails. Never throws on a live
 * signal: engines convert an aborted signal into operation cancellation.
 */
export async function prepareCfTextScope(
  sheet: XlsxSheet,
  request: CfTextScopeRequest,
): Promise<EffectiveFormatScope | undefined> {
  const rules = sheet.conditionalFormatting
  // Absence (nothing to prepare) resolves undefined; every other failure
  // throws, so engines can tell a refused preparation from an absent one.
  if (!rules || rules.rules.length === 0 || !rectsAreValid(request.rects) || request.rects.length === 0 || request.signal.aborted) return undefined
  // A single requested rect over the per-call prepared-cell cap is refused BEFORE it is
  // sliced (it always was: it used to stand alone and fail the estimator). Slicing only
  // bounds transient scratch for admissible rects; it must never turn an over-cap request
  // such as a whole column into an enumeration of many small slices.
  for (const rect of request.rects) {
    if ((rect.r1 - rect.r0 + 1) * (rect.c1 - rect.c0 + 1) > CF_LIMITS.preparedCells) throw refused('CF text preparation exceeds its admitted batch')
  }
  const chunks = chunkRects(request.rects)
  const modelChunks = chunks.map(chunk => chunk.map(toModel))
  // Fail fast on caps before any allocation; the bounded account below admits
  // every constructed byte in the caller's budget instead of a second estimate.
  for (const chunk of modelChunks) {
    if (estimateCfPrepareBytes(rules, chunk) === undefined) throw refused('CF text preparation exceeds its admitted batch')
  }
  // Reserve through a helper so budget refusals surface as refusals while
  // unrelated failures still fail closed to no scope.
  const admit = (amount: number): void => {
    try {
      request.budget.reserve(amount)
    } catch {
      throw refused('CF text preparation exceeds the admitted format budget')
    }
  }
  let bytes = DESCRIPTOR_WORKSPACE_BYTES
  const states = new Map<number, CopiedState>()
  admit(bytes)
  const fail = (): undefined => {
    states.clear()
    try { request.budget.release(bytes) } catch { /* release-only path */ }
    return undefined
  }
  const account = createCfMemory({ limitBytes: CF_LIMITS.scratchBytes, context: budgetAllocationContext(request.budget) })
  // One statistics session spans every chunk: each statistical rule's scope (scale, top10, aboveAverage,
  // duplicates) is scanned once per generation instead of once per 1000-cell chunk.
  const statistics: CfPaintStatistics = { session: createCfStatistics(account), stamp: undefined as unknown as CfStamp, results: new Map() }
  let statisticsReleased = false
  const closeAccount = (): void => {
    if (!statisticsReleased) {
      statisticsReleased = true
      for (const outcome of statistics.results.values()) if (outcome.kind === 'ready') { try { outcome.value.release() } catch { /* release-only path */ } }
      statistics.results.clear()
      try { statistics.session.dispose() } catch { /* release-only path */ }
    }
    account.close()
  }
  const hiddenRows: ModelRect[] = []
  for (const r of sheet.rows) if (r.hidden) hiddenRows.push({ firstRow: r.index, lastRow: r.index, firstCol: 0, lastCol: 16383 })
  const statisticsOptions = hiddenRows.length > 0 ? { hiddenRows } : undefined
  const sheetMerges: ModelRect[] = (sheet.mergeRanges ?? []).map(m => ({ firstRow: m.minRow, lastRow: m.maxRow, firstCol: m.minCol, lastCol: m.maxCol }))
  const dateContext = request.dateContext ?? (sheet.semantics ? {
    dateSystem: sheet.semantics.dateSystem,
    timeZone: sheet.semantics.timeZone,
    epochNowMs: Date.now(),
  } : undefined)
    try {
      const gen = cfRulesVersion(rules)
      const stamp: CfStamp = {
        owner: request.store?.stamp.docId ?? 'cf-text',
        sheetId: sheet.sheetId ?? sheet.name,
        workbookValueGen: request.store?.stamp.valueGen ?? request.workbookValueGen ?? 0,
        cfGen: gen,
        viewGen: 0,
        structureGen: request.store?.structureGen ?? 0,
        clockGen: 0,
      }
      statistics.stamp = stamp
      const source = request.store ? storeCfSource(request.store, stamp, account) : residentSource(sheet, stamp, account, request.valuePolicy ?? 'engine')
      const colors = request.colors ?? { theme: sheet.drawingTheme?.palette ?? sheet.tableStyleSource?.colors.theme, indexed: sheet.tableStyleSource?.colors.indexed }
      const locale = request.locale ?? 'en-US'
      for (const chunk of modelChunks) {
        if (request.signal.aborted) { closeAccount(); return fail() }
        const service = createCfService(rules, account, { evidence: cfRuntimeEvidence, dateContext, statistics, statisticsOptions, merges: sheetMerges })
        let lease: CfLease<CfPrepared>
        try {
          lease = await service.prepareCf(
            { stamp, sourceRects: chunk, signal: request.signal, grantBytes: CF_LIMITS.scratchBytes },
            source,
            account,
          )
        } catch (error) {
          service.dispose()
          throw error
        }
        try {
          const prepared: CfPrepared = lease.value
          for (const rect of chunk) {
            for (let row = rect.firstRow; row <= rect.lastRow; row++) {
              for (let col = rect.firstCol; col <= rect.lastCol; col++) {
                const outcome = prepared.at(row, col)
                if (outcome.kind === 'no-match') continue
                if (outcome.kind === 'unresolved') {
                  if (outcome.reason === 'cf-outside-preparation') continue
                  // Admit the retained copy before constructing it, like wins.
                  const occupied = measureCfCellBytes({ properties: {}, unresolvedWrites: undefined, unresolvedReason: outcome.reason })
                  if (!Number.isSafeInteger(bytes + occupied)) { service.dispose(); closeAccount(); throw refused('CF text preparation exceeds its admitted batch') }
                  admit(occupied)
                  bytes += occupied
                  states.set(row * 16384 + col, { properties: {}, writes: undefined, reason: outcome.reason })
                  continue
                }
                // Measure the source-owned winning state first, admit the retained
                // copy before constructing it, then store: no retained byte
                // ever exists uncharged (the spread below is transient scratch
                // for measurement; the stored copy is built after admission).
                const writes = outcome.value.unresolvedWrites ? [...outcome.value.unresolvedWrites] : undefined
                const dataBar = outcome.value.visual?.dataBar
                const occupied = measureCfCellBytes({ properties: outcome.value.properties, unresolvedWrites: writes, scale: outcome.value.visual?.colorScale, ...(dataBar ? { bar: wireBar(dataBar) } : {}) })
                if (!Number.isSafeInteger(bytes + occupied)) { service.dispose(); closeAccount(); throw refused('CF text preparation exceeds its admitted batch') }
                admit(occupied)
                bytes += occupied
                const properties: Record<string, string | number | boolean | null> = {}
                for (const key of Object.keys(outcome.value.properties)) properties[key] = outcome.value.properties[key]!
                states.set(row * 16384 + col, { properties, writes, reason: undefined, scale: outcome.value.visual?.colorScale, ...(dataBar ? { bar: descriptorOf(wireBar(dataBar)) } : {}) })
              }
            }
          }
        } finally {
          try { lease.release() } finally { service.dispose() }
        }
      }
    closeAccount()
    let disposed = false
    const scope = assembleScope(gen, states, colors, locale, bytes, () => {
      if (!disposed) {
        disposed = true
        request.budget.release(bytes)
      }
    }, () => source.stamp(stamp.sheetId).workbookValueGen === stamp.workbookValueGen && source.stamp(stamp.sheetId).structureGen === stamp.structureGen)
    if (request.signal.aborted) {
      scope.dispose()
      return undefined
    }
    return scope
  } catch (error) {
    // Abort converges to no scope (engines detect cancellation on their own
    // signal); admitted-budget refusal propagates so engines can classify it;
    // anything else fails closed to no scope. Held bytes are always released.
    states.clear()
    closeAccount()
    try { request.budget.release(bytes) } catch { /* release-only path */ }
    if (request.signal.aborted) return undefined
    if (error instanceof PackageError && error.code === 'resource-limit') throw error
    return undefined
  }
}

/** Wire grant is sized by the caller; the preflight scales with cell count. */
export const CF_TEXT_SCOPE_GRANT_BYTES = 2 * 1024 * 1024

/**
 * Wait until `isSettled` holds, the signal aborts, or `timeoutMs` elapses.
 * A worker read that never settles (no failure, no notification) resolves
 * `'timeout'` instead of hanging forever, so the caller can drop the stuck
 * entry and re-read fresh (bounded recovery). Every path removes the abort
 * listener, clears the timer, and releases the subscription — including a
 * synchronous notification that fires before subscription installation.
 */
export async function awaitSettled(
  isSettled: () => boolean,
  subscribe: (notify: () => void) => () => void,
  signal: AbortSignal,
  timeoutMs: number,
): Promise<'settled' | 'timeout'> {
  if (isSettled() || signal.aborted) return 'settled'
  return await new Promise<'settled' | 'timeout'>(resolve => {
    let done = false
    let timer: ReturnType<typeof setTimeout> | undefined
    let unsub: (() => void) | undefined
    let abortAdded = false
    const onAbort = (): void => {
      finish('settled')
    }
    const finish = (value: 'settled' | 'timeout'): void => {
      if (done) return
      done = true
      if (timer !== undefined) {
        clearTimeout(timer)
        timer = undefined
      }
      const release = unsub
      unsub = undefined
      release?.()
      if (abortAdded) {
        signal.removeEventListener('abort', onAbort)
        abortAdded = false
      }
      resolve(value)
    }
    unsub = subscribe(() => {
      if (isSettled() || signal.aborted) finish('settled')
    })
    // A synchronous notification may have finished during subscription; in
    // that case finish ran before the abort listener/timer existed, so only
    // release the (now-assigned) subscription and install nothing further.
    if (done) {
      const release = unsub
      unsub = undefined
      release?.()
      return
    }
    signal.addEventListener('abort', onAbort)
    abortAdded = true
    timer = setTimeout(() => finish('timeout'), timeoutMs)
    if (isSettled() || signal.aborted) finish('settled')
  })
}

/**
 * Scope over paged/worker read results. The caller performs one read per chunk
 * and owns the results; states are copied into budget-charged scope ownership.
 * Any non-ready chunk that is not a pending import fails the whole scope
 * closed; an empty first chunk means absent CF and resolves undefined.
 * `onEmpty` fires exactly in that absent case so callers can cache presence.
 */
export async function prepareCfWireScope(
  read: (rects: readonly SourceRect[]) => Promise<{ result: CfReadResult | undefined; release(): void }>,
  rects: readonly SourceRect[],
  options: {
    locale?: string
    colors?: SheetColorContext
    budget: FormatBudget
    onEmpty?: () => void
    current?: () => boolean
    base?: import('../tableStyles/regionLayer').RegionStyleLayer
  },
): Promise<EffectiveFormatScope | undefined> {
  if (!rects.length) return undefined
  const chunks = chunkRects(rects)
  let bytes = DESCRIPTOR_WORKSPACE_BYTES + rects.length * 128 + (options.base?.rects.length ?? 0) * 8
  const states = new Map<number, CopiedState>()
  // The initial workspace admission throws like any other refusal so callers
  // can tell a refused preparation apart from an absent one.
  try {
    options.budget.reserve(bytes)
  } catch {
    throw refused('CF wire preparation exceeds the admitted format budget')
  }
  const fail = (): undefined => {
    states.clear()
    try { options.budget.release(bytes) } catch { /* release-only path */ }
    return undefined
  }
  // Refusals throw so callers can tell them apart from absence; anything else
  // fails closed to no scope. Held bytes are always released first.
  const admitWire = (amount: number): void => {
    try {
      options.budget.reserve(amount)
    } catch {
      throw refused('CF wire preparation exceeds the admitted format budget')
    }
  }
  try {
    const regions: Array<{minRow:number;maxRow:number;minCol:number;maxCol:number}> = []
    let first = true
    let gen = 0
    for (const chunk of chunks) {
      const { result, release } = await read(chunk)
      try {
        // Absence is deterministic and needs no consistency tracking. Every
        // other non-ready outcome (including pending imports) is a refusal:
        // recording it keeps scan and materialization in agreement.
        if (options.current && !options.current()) return fail()
        if (!result) throw refused('CF wire preparation produced no result')
        if (first && !result.states.length && result.kind === 'ready' && result.definitionsVersion === versionCfRules('[]')) {
          try { options.onEmpty?.() } finally { /* observer-only */ }
          return fail()
        }
        if (result.kind !== 'ready') throw refused(`CF wire preparation is ${result.kind}`)
        if (first) {
          first = false
          gen = result.stamp.cfGen
        } else if (result.stamp.cfGen !== gen) {
          // Rules changed mid-preparation: never publish a mixed generation.
          throw refused('CF wire preparation crossed a definition generation')
        }
        for (const state of result.states) {
          const occupied = measureCfCellBytes(state) + 128
          const next = bytes + occupied
          if (!Number.isSafeInteger(next)) throw refused('CF wire preparation exceeds its admitted batch')
          admitWire(occupied)
          bytes = next
          const properties: Record<string, string | number | boolean | null> = {}
          for (const key of Object.keys(state.properties)) properties[key] = state.properties[key]!
          if (Object.keys(properties).length || state.scale?.argb !== undefined || state.bar) regions.push({minRow:state.row,maxRow:state.row,minCol:state.col,maxCol:state.col})
          states.set(state.row * 16384 + state.col, {
            properties,
            writes: state.unresolvedWrites ? [...state.unresolvedWrites] : undefined,
            reason: state.unresolvedReason,
            scale: state.scale,
            ...(state.bar ? { bar: descriptorOf(state.bar) } : {}),
          })
        }
      } finally {
        release()
      }
    }
    const colors = options.colors ?? {}
    const locale = options.locale ?? 'en-US'
    let disposed = false
    const scope = assembleScope(gen, states, colors, locale, bytes, () => {
      if (!disposed) {
        disposed = true
        options.budget.release(bytes)
      }
    }, options.current, regions)
    const paintLayer: import('../tableStyles/regionLayer').RegionStyleLayer = {
      rects: [...options.base?.rects ?? [], ...regions], retainedBytes: 0,
      styleAt: (row,col,own,mode) => scope.effectiveStyle(row,col,options.base?.styleAt(row,col,own,mode) ?? own),
    }
    if (options.current && !options.current()) { scope.dispose(); return undefined }
    return { ...scope, paintLayer }
  } catch (error) {
    // Abort converges to no scope (callers detect cancellation on their own
    // signal); admitted-budget refusal propagates for classification.
    states.clear()
    try { options.budget.release(bytes) } catch { /* release-only path */ }
    if (!(error instanceof PackageError) || error.code !== 'resource-limit') return undefined
    throw error
  }
}
