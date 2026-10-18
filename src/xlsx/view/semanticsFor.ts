import type { XlsxSheet } from '../types'

/** The parser's resolved workbook semantics for a date system (parse.ts
 * defaults: en-US, host time zone, workbook-default Unicode). Worker-backed
 * (S5.4) shell sheets and the worker filter evaluator format with it. */
export function semanticsFor(dateSystem: '1900' | '1904'): NonNullable<XlsxSheet['semantics']> {
  let timeZone = 'UTC'
  try { timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC' } catch { /* keep UTC */ }
  return { dateSystem, unicode: { version: 1, source: 'workbook-default' }, locale: 'en-US', timeZone, epochNowMs: Date.now() }
}
