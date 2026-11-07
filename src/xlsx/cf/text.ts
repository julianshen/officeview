import type { XlsxSheet } from '../types'
import type { CfStamp, SourceRect } from './model'
import { CF_LIMITS, createCfMemory } from './limits'
import { createCfService } from './service'
import { resolveCfCellStyle } from './dxf'
import { residentSource } from './render'
import { cfRuntimeEvidence } from './predicates'
import { cfRulesVersion } from '../view/pagedSource'
import type { EffectiveFormatScope, FormatBudget } from '../view/formatWorkspace'
import type { SheetColorContext } from '../sheetColors'
import type { ValuePolicy } from '../view/values'
import type { CfCellState, CfWriteMask } from './model'
import type { CfReadResult } from '../view/contracts'
import { describeCfUnavailable, describeCfWin, type CfDescriptor } from './describe'

/** Bounded per-batch effective-format preparation for derived text (Task 9). */
export interface CfTextScopeRequest {
  readonly rects: readonly SourceRect[]
  readonly signal: AbortSignal
  readonly locale?: string
  readonly budget?: FormatBudget
  readonly colors?: SheetColorContext
  readonly valuePolicy?: ValuePolicy
  readonly workbookValueGen?: number
}

/** Fixed admitted descriptor workspace per scope with a charged budget. */
const DESCRIPTOR_WORKSPACE_BYTES = 2048
const DESCRIPTOR_CACHE_KEYS = 32

/** Wire grant for one text-scope batch: covers a full scan/copy page of typical
 * wins (the preflight scales with cell count, not wins). Dense or taint-heavy
 * pages refuse to base text (fail closed) instead of growing. */
export const CF_TEXT_SCOPE_GRANT_BYTES = 2 * 1024 * 1024

/** Structural key for one cell's descriptor inputs: shared aspects share one charge. */
const winKey = (
  properties: Readonly<Record<string, string | number | boolean | null>>,
  unresolvedWrites: CfWriteMask | readonly string[] | undefined,
): string => {
  const paths = Object.keys(properties).sort().join(',')
  const writes = unresolvedWrites ? [...unresolvedWrites].sort().join(',') : ''
  return `${paths}|${writes}`
}

const winDescriptor = (
  properties: Readonly<Record<string, string | number | boolean | null>>,
  unresolvedWrites: CfWriteMask | readonly string[] | undefined,
  unresolvedReason: string | undefined,
  locale: string,
  cache: Map<string, CfDescriptor> | undefined,
): CfDescriptor | undefined => {
  const key = winKey(properties, unresolvedWrites)
  const cached = cache?.get(key)
  if (cached) return cached
  const descriptor = unresolvedReason !== undefined
    ? describeCfUnavailable(locale)
    : describeCfWin(properties, unresolvedWrites, undefined, locale)
  if (descriptor && cache && cache.size < DESCRIPTOR_CACHE_KEYS) cache.set(key, descriptor)
  return descriptor ?? undefined
}

/**
 * Resident preparation over the given rects. Returns undefined (fail closed,
 * base formatting) when the sheet has no CF, the scope exceeds caps, the
 * descriptor workspace is refused, or preparation fails. Never throws on a
 * live signal: engines convert an aborted signal into operation cancellation.
 */
export async function prepareCfTextScope(
  sheet: XlsxSheet,
  request: CfTextScopeRequest,
): Promise<EffectiveFormatScope | undefined> {
  const rules = sheet.conditionalFormatting
  if (!rules || rules.rules.length === 0 || !request.rects.length) return undefined
  if (request.signal.aborted) return undefined
  if (request.budget) {
    try {
      request.budget.reserve(DESCRIPTOR_WORKSPACE_BYTES)
    } catch {
      return undefined
    }
  }
  const account = createCfMemory({ limitBytes: CF_LIMITS.scratchBytes })
  const service = createCfService(rules, account, { evidence: cfRuntimeEvidence })
  const fail = (): undefined => {
    try { service.dispose() } finally { account.close() }
    request.budget?.release(DESCRIPTOR_WORKSPACE_BYTES)
    return undefined
  }
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
    const lease = await service.prepareCf(
      { stamp, sourceRects: request.rects, signal: request.signal, grantBytes: CF_LIMITS.scratchBytes },
      source,
      account,
    )
    if (request.signal.aborted) {
      try { lease.release() } finally { fail() }
      return undefined
    }
    const prepared = lease.value
    const colors = request.colors ?? {}
    const locale = request.locale ?? 'en-US'
    const cache = request.budget ? new Map<string, CfDescriptor>() : undefined
    let disposed = false
    return {
      gen,
      effectiveStyle: (row, col, base) => {
        if (disposed) return base
        const outcome = prepared.at(row, col)
        if (outcome.kind !== 'ready') return base
        return resolveCfCellStyle(base, outcome.value, colors)
      },
      ...(cache ? {
        describe: (row, col) => {
          if (disposed) return undefined
          const outcome = prepared.at(row, col)
          if (outcome.kind === 'no-match') return undefined
          if (outcome.kind === 'unresolved') {
            // Outside the prepared rects there is no information, not a failure.
            if (outcome.reason === 'cf-outside-preparation') return undefined
            return winDescriptor({}, undefined, outcome.reason, locale, cache)
          }
          const state: CfCellState = outcome.value
          const blocked = state.unresolvedWrites && state.unresolvedWrites.size > 0 ? [...state.unresolvedWrites] : undefined
          if (Object.keys(state.properties).length === 0 && !blocked) return undefined
          return winDescriptor(state.properties, blocked, undefined, locale, cache)
        },
      } : {}),
      dispose: () => {
        if (disposed) return
        disposed = true
        try { lease.release() } finally { try { service.dispose() } finally { account.close() } }
        request.budget?.release(DESCRIPTOR_WORKSPACE_BYTES)
      },
    }
  } catch {
    return fail()
  }
}

/**
 * Scope over a ready paged/worker read result. The caller owns the result and
 * any acquired cache entry: pass its release as `onDispose`. No descriptors
 * without a budget (find/copy need text only).
 */
export function cfTextScopeFromWire(
  result: CfReadResult,
  options: { locale?: string; colors?: SheetColorContext; budget?: FormatBudget; onDispose?: () => void },
): EffectiveFormatScope | undefined {
  if (result.kind !== 'ready') return undefined
  if (options.budget) {
    try {
      options.budget.reserve(DESCRIPTOR_WORKSPACE_BYTES)
    } catch {
      return undefined
    }
  }
  const fail = (): undefined => {
    options.budget?.release(DESCRIPTOR_WORKSPACE_BYTES)
    return undefined
  }
  try {
    const byCell = new Map<number, (typeof result.states)[number]>()
    for (const state of result.states) byCell.set(state.row * 16384 + state.col, state)
    const colors = options.colors ?? {}
    const locale = options.locale ?? 'en-US'
    const cache = options.budget ? new Map<string, CfDescriptor>() : undefined
    let disposed = false
    const unresolvedOf = (row: number, col: number): { writes: readonly string[]; reason: string | undefined } | undefined => {
      const state = byCell.get(row * 16384 + col)
      if (!state) return undefined
      if ((state.unresolvedWrites?.length ?? 0) === 0 && state.unresolvedReason === undefined) return undefined
      return { writes: state.unresolvedWrites ?? [], reason: state.unresolvedReason }
    }
    return {
      gen: result.stamp.cfGen,
      effectiveStyle: (row, col, base) => {
        if (disposed) return base
        const state = byCell.get(row * 16384 + col)
        if (!state) return base
        if ((state.unresolvedWrites?.length ?? 0) > 0 || state.unresolvedReason !== undefined) return base
        return resolveCfCellStyle(base, { properties: state.properties }, colors)
      },
      ...(cache ? {
        describe: (row, col) => {
          if (disposed) return undefined
          const state = byCell.get(row * 16384 + col)
          if (!state) return undefined
          if (Object.keys(state.properties).length === 0 && (state.unresolvedWrites?.length ?? 0) === 0 && state.unresolvedReason === undefined) return undefined
          const blocked = unresolvedOf(row, col)
          return winDescriptor(state.properties, blocked?.writes, blocked?.reason, locale, cache)
        },
      } : {}),
      dispose: () => {
        if (disposed) return
        disposed = true
        byCell.clear()
        cache?.clear()
        options.budget?.release(DESCRIPTOR_WORKSPACE_BYTES)
        options.onDispose?.()
      },
    }
  } catch {
    return fail()
  }
}
