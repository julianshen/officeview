/**
 * Stage 5 final review F2: a clock case's captured instant and time zone, pinned through parse, production and the
 * counterfactual, so a validation run is independent of the host's zone and day (CI runs on UTC Linux, on any day).
 *
 * Production's own semantics stay "the viewer's local time", as Excel's: the resident document takes its zone from the
 * ambient zone at parse, and the paged workbook from the ambient zone when its policy is first built. A clock case was
 * measured by a viewer in `requiredTimeZone` on its anchor day, so validation evaluates as that viewer:
 * - parse: `caseParseOptionsOf` pins the load-time worksheet evaluation's zone and instant;
 * - resident production and the counterfactual: `caseDocumentEvaluationOf` / `caseEvaluationOf` take the document's policy
 *   with the case zone, at the anchor instant (`caseNowOf`), whatever zone the document was parsed under;
 * - the worker (paged) route, which has no policy seam: `withCaseZone` sets the ambient zone before the session opens.
 * No production seam changes.
 */
import type { ParseXlsxOptions } from '../../src/xlsx/parse'
import type { XlsxDocument } from '../../src/xlsx/types'
import { cfDocumentEvaluationOf, cfEvaluationContextFor, type CfDocumentEvaluation, type CfEvaluationContext } from '../../src/xlsx/cf/formula'

export interface CaseClockLike { clockRequirement?: Readonly<{ anchorLocalDate: string; requiredTimeZone: string }> }
export interface CaseClock {
  /** The anchor instant: 04:00 local time on the anchor day (inside the 00:05-07:55 UTC/local distinguishing window). */
  readonly epochMs: number
  readonly now: () => number
  readonly timeZone: string
}

/** Fixed offsets of the zones a manifest may require (Asia/Taipei has no daylight saving time). */
const ZONE_OFFSETS: Readonly<Record<string, string>> = { 'Asia/Taipei': '+08:00' }

/** The case's pinned clock, or undefined for a case with no clock requirement. */
export function caseClockOf(c: CaseClockLike): CaseClock | undefined {
  const req = c.clockRequirement
  if (!req) return undefined
  const offset = ZONE_OFFSETS[req.requiredTimeZone]
  if (!offset) throw new Error(`case clock: no fixed offset recorded for ${req.requiredTimeZone}`)
  const epochMs = Date.parse(`${req.anchorLocalDate}T04:00:00${offset}`)
  if (!Number.isFinite(epochMs)) throw new Error(`case clock: invalid anchor ${req.anchorLocalDate}`)
  return Object.freeze({ epochMs, now: () => epochMs, timeZone: req.requiredTimeZone })
}

/** The instant a case's evaluation reads: the anchor for a clock case, else `fallback` (default `Date.now`). */
export function caseNowOf(c: CaseClockLike, fallback: () => number = Date.now): () => number {
  return caseClockOf(c)?.now ?? fallback
}

/** Parse options that pin the worksheet formulas' clock and zone to the case's (load-time evaluation honours both). */
export function caseParseOptionsOf(c: CaseClockLike, base: ParseXlsxOptions = {}): ParseXlsxOptions {
  const clock = caseClockOf(c)
  return clock ? { ...base, evaluation: { ...base.evaluation, timeZone: clock.timeZone, now: clock.now } } : base
}

/** The document's CF policy with a clock case's zone pinned (whatever zone the document was parsed under); unchanged for
 * a case with no clock requirement. */
export function caseDocumentEvaluationOf(c: CaseClockLike, doc: Pick<XlsxDocument, 'calc' | 'valuePolicy' | 'semantics'>): CfDocumentEvaluation {
  const policy = cfDocumentEvaluationOf(doc)
  const clock = caseClockOf(c)
  return clock && policy.timeZone !== clock.timeZone ? Object.freeze({ ...policy, timeZone: clock.timeZone }) : policy
}

/** One preparation's evaluation context for a case: the pinned policy at the case's instant (`fallback` for a case with
 * no clock requirement). */
export function caseEvaluationOf(c: CaseClockLike, doc: Pick<XlsxDocument, 'calc' | 'valuePolicy' | 'semantics'>, fallback: () => number = Date.now): CfEvaluationContext {
  return cfEvaluationContextFor(caseDocumentEvaluationOf(c, doc), caseNowOf(c, fallback))
}

const ambientZone = (): string => {
  try { return Intl.DateTimeFormat().resolvedOptions().timeZone } catch { return '' }
}

/**
 * Runs `fn` with the ambient zone set to the case's required zone (a no-op for a case with no clock requirement), and
 * restores the previous `TZ` afterwards. Throws if the runtime does not honour the change, rather than run in the host's
 * zone. Enter it before the parse or the paged session opens: the resident and paged policies are built once per document.
 */
export async function withCaseZone<T>(c: CaseClockLike, fn: () => Promise<T>): Promise<T> {
  const clock = caseClockOf(c)
  if (!clock) return fn()
  const previous = process.env.TZ
  process.env.TZ = clock.timeZone
  try {
    if (ambientZone() !== clock.timeZone) throw new Error(`case clock: the runtime did not adopt TZ=${clock.timeZone} (ambient ${ambientZone()})`)
    return await fn()
  } finally {
    if (previous === undefined) delete process.env.TZ
    else process.env.TZ = previous
  }
}
