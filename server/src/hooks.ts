/** The developer's server-side decision hooks (§6.1) and the effective
 * per-session policy with the protected defaults (§2 item 5). */

export interface AuthRequest { headers: Headers; ip: string; url: URL; token?: string; resume?: { sid: string } }
export interface AuthDecision {
  viewerId?: string
  role?: 'presenter' | 'audience'
  /** On a resume whose proof is NOT the live session token: the hook vouches
   * that it issued this proof for this viewer (§6.2 rule 7). */
  resumeGranted?: boolean
  policy: {
    contentPolicy?: 'visible-only' | 'full'
    projection?: 'protected' | 'relaxed'
    visibleMark?: 'fixed' | 'viewer' | 'off'
    forensicMark?: boolean
    /** Print service (Stage 3b). On by default; `false` turns it off. A cap
     * that is not a positive safe integer turns print off (fail closed). */
    print?: boolean | { maxPages: number; maxSessionPages?: number }
    /** Copy service (Stage 3a). Off by default; `'full'` is uncapped; a cap
     * that is not a positive safe integer turns copy off (fail closed). */
    copy?: false | 'full' | { maxParagraphs: number; maxSessionParagraphs?: number }
    /** Screen-reader text. `false` always turns it off; `true` serves per-unit
     * text when copy is off. Otherwise it follows `copy` (see `a11yMode`). */
    textA11y?: boolean
  }
}
export type Authorize = (req: AuthRequest) => Promise<AuthDecision | null>
/** Which document this viewer opens. The server has no file route; this hook
 * is the only way bytes enter a session. */
export type ResolveDocument = (req: AuthRequest, decision: AuthDecision) => Promise<{ bytes: Uint8Array; name?: string } | null>

export interface EffectivePolicy {
  contentPolicy: 'visible-only' | 'full'
  projection: 'final-raster' | 'commands'
  visibleMark: 'fixed' | 'viewer' | 'off'
  /** Layer-1 on final-raster units (Stage 2a). Command units carry no
   * forensic mark until Layer 2; notices say so (never claim an unapplied mark). */
  forensicMark: boolean
  /** What the developer asked for, recorded in the audit row. */
  forensicRequested: boolean
  /** Resolved print caps; `true`/unset ⇒ both null. Mutable on resume. */
  print: PrintPolicy
  role: 'presenter' | 'audience'
  /** Any departure from the protected defaults (audited, §6.1). Copy and
   * a11y do not count (R4): the audit row records the resolved services at
   * open (`AuditRow.services`), and every resume that changes them appends a
   * timestamped entry (`AuditRow.servicesChanged`, the audit `services` table). */
  relaxed: boolean
  /** Resolved copy caps; `'full'` ⇒ both null. Mutable on resume. */
  copy: CopyPolicy
  /** Resolved a11y mode. Mutable on resume. */
  textA11y: A11yMode
}
export type CopyPolicy = false | { maxParagraphs: number | null; maxSessionParagraphs: number | null }
export type PrintPolicy = false | { maxPages: number | null; maxSessionPages: number | null }
export type A11yMode = 'off' | 'unit' | 'focused'

const isCap = (v: unknown): v is number => Number.isSafeInteger(v) && (v as number) > 0

function copyPolicy(c: unknown): CopyPolicy {
  if (c === 'full') return { maxParagraphs: null, maxSessionParagraphs: null }
  if (!c || typeof c !== 'object' || Array.isArray(c)) return false
  const { maxParagraphs, maxSessionParagraphs } = c as Record<string, unknown>
  if (!isCap(maxParagraphs)) return false
  if (maxSessionParagraphs !== undefined && !isCap(maxSessionParagraphs)) return false
  return { maxParagraphs, maxSessionParagraphs: maxSessionParagraphs ?? null }
}

/** `true`/unset ⇒ uncapped; a cap object needs a positive safe-integer
 * `maxPages` (and `maxSessionPages` when given); anything else is off. */
export function printPolicy(p: unknown): PrintPolicy {
  if (p === undefined || p === true) return { maxPages: null, maxSessionPages: null }
  if (!p || typeof p !== 'object' || Array.isArray(p)) return false
  const { maxPages, maxSessionPages } = p as Record<string, unknown>
  if (!isCap(maxPages)) return false
  if (maxSessionPages !== undefined && !isCap(maxSessionPages)) return false
  return { maxPages, maxSessionPages: maxSessionPages ?? null }
}

/** 'unit' when copy is uncapped or a11y is asked for with copy off; 'focused'
 * when copy is capped; an explicit `textA11y: false` is always 'off'. */
function a11yMode(copy: CopyPolicy, textA11y: unknown): A11yMode {
  if (textA11y === false) return 'off'
  if (copy) return copy.maxParagraphs === null ? 'unit' : 'focused'
  return textA11y === true ? 'unit' : 'off'
}

/** The policy fields a resume must match exactly (R3): every field but the
 * mutable services (text, and print caps per Stage 3b P4). */
export function resumeInvariant(p: EffectivePolicy): Omit<EffectivePolicy, 'copy' | 'textA11y' | 'print'> {
  const { copy: _c, textA11y: _a, print: _p, ...rest } = p
  return rest
}

/** Manifest notices for what the forensic mark actually covers (§13.1). */
export function forensicNotices(p: Pick<EffectivePolicy, 'forensicMark' | 'projection'>): string[] {
  if (!p.forensicMark) return ['forensic-off']
  return p.projection === 'commands' ? ['forensic-commands-unmarked'] : []
}

export function effectivePolicy(d: AuthDecision): EffectivePolicy {
  const p = d.policy ?? {}
  const contentPolicy = p.contentPolicy === 'full' ? 'full' : 'visible-only'
  const projection = p.projection === 'relaxed' ? 'commands' : 'final-raster'
  const visibleMark = p.visibleMark === 'viewer' || p.visibleMark === 'off' ? p.visibleMark : 'fixed'
  const forensicRequested = p.forensicMark !== false
  const copy = copyPolicy(p.copy)
  return {
    contentPolicy, projection, visibleMark, forensicMark: forensicRequested, forensicRequested,
    print: printPolicy(p.print), role: d.role === 'presenter' ? 'presenter' : 'audience',
    relaxed: contentPolicy !== 'visible-only' || projection !== 'final-raster' || visibleMark === 'off' || !forensicRequested,
    copy, textA11y: a11yMode(copy, p.textA11y),
  }
}
