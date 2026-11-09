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
    print?: boolean
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
  print: boolean
  role: 'presenter' | 'audience'
  /** Any departure from the protected defaults (audited, §6.1). */
  relaxed: boolean
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
  return {
    contentPolicy, projection, visibleMark, forensicMark: forensicRequested, forensicRequested,
    print: p.print !== false, role: d.role === 'presenter' ? 'presenter' : 'audience',
    relaxed: contentPolicy !== 'visible-only' || projection !== 'final-raster' || visibleMark === 'off' || !forensicRequested,
  }
}
