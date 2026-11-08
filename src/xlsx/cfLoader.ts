/** Tiny bridge only: a CF-free sheet creates no account, cache or lazy import. */
import type { XmlNode } from '../core/xml'
import { attrs, getChildren } from '../core/xml'
import type { XlsxSheet } from './types'
import type { CfRule } from './cf/model'
type CfFoundation = typeof import('./cf/limits')
let pending: Promise<CfFoundation> | undefined
export function preloadCf(): Promise<CfFoundation> {
  return pending ??= import('./cf/limits').catch(error => {
    pending = undefined
    throw error
  })
}
export async function ensureCf<T = CfFoundation>(
  sheets: readonly Pick<XlsxSheet, 'conditionalFormatting'>[],
  load: () => Promise<T> = preloadCf as () => Promise<T>,
): Promise<T | undefined> {
  return sheets.some(sheet => sheet.conditionalFormatting !== undefined) ? load() : undefined
}

type CfParser = typeof import('./cf/parse')
let parserPending: Promise<CfParser> | undefined
/** Retryable lazy load of the CF rule parser (resident sheet parse path). */
export function preloadCfParser(): Promise<CfParser> {
  return parserPending ??= import('./cf/parse').catch(error => {
    parserPending = undefined
    throw error
  })
}
/** Loads the rule parser only when the worksheet actually carries CF content. */
export async function ensureCfParser<T = CfParser>(
  hasCfContent: boolean,
  load: () => Promise<T> = preloadCfParser as () => Promise<T>,
): Promise<T | undefined> {
  return hasCfContent ? load() : undefined
}

/**
 * Worksheet `ext` URI carrying `x14:conditionalFormattings`. Kept as a literal
 * (mirroring `CF_X14_EXTENSION_URI` in `./cf/parse`) so this bridge stays
 * eager-trivial: no runtime import of the lazy CF chunk.
 */
export const CF_PRESENCE_EXTENSION_URI = '{78C0D931-6437-407d-A8EE-F0AA7539E650}'

/**
 * Lexical presence check for conditional formatting on a worksheet node.
 * Local-name based and allocation-free: deciding *whether* to load the lazy
 * chunk must not itself import, price, or cache anything. Namespace strictness
 * is enforced later by the lazy parser, never here.
 */
export function hasConditionalFormatting(data: XmlNode | undefined): boolean {
  if (!data) return false
  if (getChildren(data, 'conditionalFormatting').length > 0) return true
  for (const extList of getChildren(data, 'extLst')) {
    for (const ext of getChildren(extList, 'ext')) {
      if (attrs(ext).uri !== CF_PRESENCE_EXTENSION_URI) continue
      if (getChildren(ext, 'conditionalFormattings').length > 0) return true
    }
  }
  return false
}

/** Explicit finite CF applies-to bounds, without loading the evaluator. Full
 * rows/columns never enlarge the raster into a worksheet-sized allocation. */
export function cfExplicitBounds(sheet: Pick<XlsxSheet, 'conditionalFormatting'>): { r0: number; r1: number; c0: number; c1: number } | undefined {
  const cf = sheet.conditionalFormatting
  if (!cf || cf.rules.length > 4096) return undefined
  let r0 = 1048576, r1 = -1, c0 = 16384, c1 = -1, count = 0
  for (const rule of cf.rules) for (const r of rule.ranges) {
    if (++count > 16384) return undefined
    if (r.lastRow === 1048575 || r.lastCol === 16383) continue
    if (!Number.isSafeInteger(r.firstRow) || !Number.isSafeInteger(r.lastRow) || !Number.isSafeInteger(r.firstCol) || !Number.isSafeInteger(r.lastCol)
      || r.firstRow < 0 || r.lastRow >= 1048576 || r.firstCol < 0 || r.lastCol >= 16384 || r.lastRow < r.firstRow || r.lastCol < r.firstCol) continue
    r0 = Math.min(r0, r.firstRow); r1 = Math.max(r1, r.lastRow); c0 = Math.min(c0, r.firstCol); c1 = Math.max(c1, r.lastCol)
  }
  return r1 < 0 ? undefined : { r0, r1, c0, c1 }
}
let renderPending: Promise<typeof import('./cf/render')> | undefined
export function ensureConditionalFormatting(sheets: readonly Pick<XlsxSheet, 'conditionalFormatting'>[]): Promise<typeof import('./cf/render') | undefined> {
  if (!sheets.some(sheet => sheet.conditionalFormatting !== undefined)) return Promise.resolve(undefined)
  return renderPending ??= import('./cf/render').catch(error => { renderPending = undefined; throw error })
}

/** Shared paint-batch boundary; metadata-only so the viewer need not load CF. */
export const CF_PAINT_BATCH_CELLS = 5000
export function cfNeedsViewport(sheet: Pick<XlsxSheet, 'conditionalFormatting'>): boolean {
  let span = 0
  for (const rule of sheet.conditionalFormatting?.rules ?? []) for (const r of rule.ranges) {
    if (r.lastRow === 1048575 || r.lastCol === 16383) return true
    span += (r.lastRow - r.firstRow + 1) * (r.lastCol - r.firstCol + 1)
    if (span > CF_PAINT_BATCH_CELLS) return true
  }
  return false
}

/**
 * Typed check for whether a rule definition can potentially match blank cells (M4).
 * Centralized next to cfExplicitBounds without dynamic casting.
 */
export function canRuleMatchBlanks(rule: Pick<CfRule, 'definition'>): boolean {
  const def = rule.definition
  if (def.kind === 'predicate') return def.type === 'containsBlanks'
  if (def.kind === 'unresolved') return def.type === 'expression'
  return false
}
