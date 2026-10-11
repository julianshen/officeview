import type { XlsxDocument } from '../types'

/**
 * Per-document value policy flag (spec §4.4), re-exported by ./workbook. Kept
 * free of evaluator imports so source adapters can read it without pulling the
 * formula engine into their chunk. Only a cached document carries the flag, so
 * an engine document is byte-identical to one parsed before it existed; a plain
 * own property (not a WeakSet/Symbol) survives structured cloning to a worker.
 */
export type WorkbookValuePolicy = NonNullable<XlsxDocument['valuePolicy']>
export function workbookValuePolicy(doc: XlsxDocument): WorkbookValuePolicy {
  return doc.valuePolicy === 'cached' ? 'cached' : 'engine'
}
export function markSavedValuesOnly(doc: XlsxDocument): void {
  doc.valuePolicy = 'cached'
}
