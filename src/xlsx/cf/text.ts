import type { XlsxSheet } from '../types'
import type { CfPrepared, CfStamp, SourceRect as ModelRect } from './model'
import { CF_LIMITS, createCfMemory } from './limits'
import { createCfService, estimateCfPrepareBytes } from './service'
import { resolveCfCellStyle } from './dxf'
import { residentSource } from './render'
import { cfRuntimeEvidence } from './predicates'
import { cfRulesVersion } from '../view/pagedSource'
import type { EffectiveFormatScope, FormatBudget } from '../view/formatWorkspace'
import type { SheetColorContext } from '../sheetColors'
import type { ValuePolicy } from '../view/values'
import type { CfReadResult } from '../view/contracts'
import { measureCfCellBytes, versionCfRules } from '../view/contracts'
import type { SourceRect } from '../view/coordinates'
import { describeCfUnavailable, describeCfWin, type CfDescriptor } from './describe'

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
}

/** Fixed admitted descriptor workspace per scope with a charged budget. */
const DESCRIPTOR_WORKSPACE_BYTES = 2048
const DESCRIPTOR_CACHE_KEYS = 32
/** Mirrors the service per-call limits so every chunk is admittable. */
const CHUNK_RECTS = 256
const CHUNK_CELLS = 5000

/** Copied winning state: the scope owns this data, charged in the budget. */
interface CopiedState {
  readonly properties: Record<string, string | number | boolean | null>
  readonly writes: readonly string[] | undefined
  readonly reason: string | undefined
}

/** Structural key for one cell's descriptor inputs: shared aspects share one charge. */
const winKey = (properties: Readonly<Record<string, string | number | boolean | null>>, unresolvedWrites: readonly string[] | undefined): string => {
  const paths = Object.keys(properties).sort().join(',')
  const writes = unresolvedWrites ? [...unresolvedWrites].sort().join(',') : ''
  return `${paths}|${writes}`
}

const winDescriptor = (
  properties: Readonly<Record<string, string | number | boolean | null>>,
  unresolvedWrites: readonly string[] | undefined,
  unresolvedReason: string | undefined,
  locale: string,
  cache: Map<string, CfDescriptor>,
): CfDescriptor | undefined => {
  const key = winKey(properties, unresolvedWrites)
  const cached = cache.get(key)
  if (cached) return cached
  const descriptor = unresolvedReason !== undefined
    ? describeCfUnavailable(locale)
    : describeCfWin(properties, unresolvedWrites, undefined, locale)
  if (descriptor && cache.size < DESCRIPTOR_CACHE_KEYS) cache.set(key, descriptor)
  return descriptor ?? undefined
}

/** Greedy chunks honoring the service per-call limits; a single over-cap rect
 * stands alone so the service refusal (not an over-wide request) governs it. */
function chunkRects(rects: readonly SourceRect[]): SourceRect[][] {
  const chunks: SourceRect[][] = []
  let current: SourceRect[] = [], area = 0
  const flush = (): void => { if (current.length) { chunks.push(current); current = []; area = 0 } }
  for (const rect of rects) {
    const rectArea = (rect.r1 - rect.r0 + 1) * (rect.c1 - rect.c0 + 1)
    if (current.length >= CHUNK_RECTS || area + rectArea > CHUNK_CELLS) flush()
    current.push(rect)
    area += rectArea
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
): EffectiveFormatScope {
  const cache = new Map<string, CfDescriptor>()
  let disposed = false
  return {
    gen,
    bytes,
    effectiveStyle: (row, col, base) => {
      if (disposed) return base
      const state = states.get(row * 16384 + col)
      if (!state) return base
      // Surviving winning properties apply even when other aspects are
      // blocked: only the descriptor reports the unavailable aspect.
      if (Object.keys(state.properties).length === 0) return base
      return resolveCfCellStyle(base, { properties: state.properties }, colors)
    },
    describe: (row, col) => {
      if (disposed) return undefined
      const state = states.get(row * 16384 + col)
      if (!state) return undefined
      if (Object.keys(state.properties).length === 0 && (state.writes?.length ?? 0) === 0 && state.reason === undefined) return undefined
      return winDescriptor(state.properties, state.writes, state.reason, locale, cache)
    },
    dispose: () => {
      if (disposed) return
      disposed = true
      states.clear()
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
  if (!rules || rules.rules.length === 0 || !request.rects.length || request.signal.aborted) return undefined
  const chunks = chunkRects(request.rects)
  const modelChunks = chunks.map(chunk => chunk.map(toModel))
  let envelope = DESCRIPTOR_WORKSPACE_BYTES
  for (const chunk of modelChunks) {
    const estimate = estimateCfPrepareBytes(rules, chunk)
    if (estimate === undefined) return undefined
    envelope += estimate
    if (!Number.isSafeInteger(envelope)) return undefined
  }
  try {
    request.budget.reserve(envelope)
  } catch {
    return undefined
  }
  const fail = (held: number): undefined => {
    if (held > 0) {
      try { request.budget.release(held) } catch { /* release-only path */ }
    }
    return undefined
  }
  const account = createCfMemory({ limitBytes: CF_LIMITS.scratchBytes })
  const service = createCfService(rules, account, { evidence: cfRuntimeEvidence })
  const states = new Map<number, CopiedState>()
  try {
    const gen = cfRulesVersion(rules)
    const stamp: CfStamp = {
      owner: 'cf-text',
      sheetId: sheet.sheetId ?? sheet.name,
      workbookValueGen: request.workbookValueGen ?? 0,
      cfGen: gen,
      viewGen: 0,
      structureGen: 0,
      clockGen: 0,
    }
    const source = residentSource(sheet, stamp, account, request.valuePolicy ?? 'engine')
    const colors = request.colors ?? {}
    const locale = request.locale ?? 'en-US'
    for (const chunk of modelChunks) {
      if (request.signal.aborted) { service.dispose(); account.close(); return fail(envelope) }
      const lease = await service.prepareCf(
        { stamp, sourceRects: chunk, signal: request.signal, grantBytes: CF_LIMITS.scratchBytes },
        source,
        account,
      )
      try {
        const prepared: CfPrepared = lease.value
        for (const rect of chunk) {
          for (let row = rect.firstRow; row <= rect.lastRow; row++) {
            for (let col = rect.firstCol; col <= rect.lastCol; col++) {
              const outcome = prepared.at(row, col)
              if (outcome.kind === 'no-match') continue
              if (outcome.kind === 'unresolved') {
                if (outcome.reason === 'cf-outside-preparation') continue
                states.set(row * 16384 + col, { properties: {}, writes: undefined, reason: outcome.reason })
                continue
              }
              const properties: Record<string, string | number | boolean | null> = {}
              for (const key of Object.keys(outcome.value.properties)) properties[key] = outcome.value.properties[key]!
              const writes = outcome.value.unresolvedWrites ? [...outcome.value.unresolvedWrites] : undefined
              states.set(row * 16384 + col, { properties, writes, reason: undefined })
            }
          }
        }
      } finally {
        lease.release()
      }
    }
    service.dispose()
    account.close()
    let disposed = false
    const scope = assembleScope(gen, states, colors, locale, envelope, () => {
      if (!disposed) {
        disposed = true
        request.budget.release(envelope)
      }
    })
    if (request.signal.aborted) {
      scope.dispose()
      return undefined
    }
    return scope
  } catch {
    try { service.dispose() } finally { account.close() }
    return fail(envelope)
  }
}

/** Wire grant is sized by the caller; the preflight scales with cell count. */
export const CF_TEXT_SCOPE_GRANT_BYTES = 2 * 1024 * 1024

/**
 * Scope over paged/worker read results. The caller performs one read per chunk
 * and owns the results; states are copied into budget-charged scope ownership.
 * Any non-ready chunk fails the whole scope closed. `onEmpty` fires when the
 * first chunk is ready, stateless and definition-free (absent CF).
 */
export async function prepareCfWireScope(
  read: (rects: readonly SourceRect[]) => Promise<{ result: CfReadResult | undefined; release(): void }>,
  rects: readonly SourceRect[],
  options: {
    locale?: string
    colors?: SheetColorContext
    budget: FormatBudget
    onEmpty?: () => void
  },
): Promise<EffectiveFormatScope | undefined> {
  if (!rects.length) return undefined
  const chunks = chunkRects(rects)
  let bytes = DESCRIPTOR_WORKSPACE_BYTES
  const states = new Map<number, CopiedState>()
  try {
    options.budget.reserve(bytes)
  } catch {
    return undefined
  }
  const fail = (): undefined => {
    states.clear()
    try { options.budget.release(bytes) } catch { /* release-only path */ }
    return undefined
  }
  try {
    let first = true
    let gen = 0
    for (const chunk of chunks) {
      const { result, release } = await read(chunk)
      try {
        if (!result || result.kind !== 'ready') return fail()
        if (first) {
          first = false
          gen = result.stamp.cfGen
          if (!result.states.length && result.definitionsVersion === versionCfRules('[]')) {
            try { options.onEmpty?.() } finally { /* observer-only */ }
            return fail()
          }
        } else if (result.stamp.cfGen !== gen) {
          // Rules changed mid-preparation: never publish a mixed generation.
          return fail()
        }
        for (const state of result.states) {
          const occupied = measureCfCellBytes(state)
          const next = bytes + occupied
          if (!Number.isSafeInteger(next)) return fail()
          try {
            options.budget.reserve(occupied)
          } catch {
            return fail()
          }
          bytes = next
          const properties: Record<string, string | number | boolean | null> = {}
          for (const key of Object.keys(state.properties)) properties[key] = state.properties[key]!
          states.set(state.row * 16384 + state.col, {
            properties,
            writes: state.unresolvedWrites ? [...state.unresolvedWrites] : undefined,
            reason: state.unresolvedReason,
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
    })
    return scope
  } catch {
    return fail()
  }
}
