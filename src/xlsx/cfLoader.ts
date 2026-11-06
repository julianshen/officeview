/** Tiny bridge only: a CF-free sheet creates no account, cache or lazy import. */
import type { XmlNode } from '../core/xml'
import { attrs, getChildren } from '../core/xml'
import type { XlsxSheet } from './types'
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
