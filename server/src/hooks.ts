/** The developer's server-side decision hooks (§6.1) and the effective
 * per-session policy with the protected defaults (§2 item 5). */

export interface AuthRequest { headers: Headers; ip: string; url: URL; token?: string; resume?: { sid: string } }
export interface AuthDecision {
  viewerId?: string
  role?: 'presenter' | 'audience'
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
  /** Always false in Stage 1: layer-1 is Stage 2 (manifest notice
   * `forensic-mark-stage-2`); never claim a mark that is not applied. */
  forensicMark: false
  /** What the developer asked for, recorded in the audit row. */
  forensicRequested: boolean
  print: boolean
  role: 'presenter' | 'audience'
  /** Any departure from the protected defaults (audited, §6.1). */
  relaxed: boolean
}

export const FORENSIC_NOTICE = 'forensic-mark-stage-2'

export function effectivePolicy(d: AuthDecision): EffectivePolicy {
  const p = d.policy ?? {}
  const contentPolicy = p.contentPolicy === 'full' ? 'full' : 'visible-only'
  const projection = p.projection === 'relaxed' ? 'commands' : 'final-raster'
  const visibleMark = p.visibleMark === 'viewer' || p.visibleMark === 'off' ? p.visibleMark : 'fixed'
  const forensicRequested = p.forensicMark !== false
  return {
    contentPolicy, projection, visibleMark, forensicMark: false, forensicRequested,
    print: p.print !== false, role: d.role === 'presenter' ? 'presenter' : 'audience',
    relaxed: contentPolicy !== 'visible-only' || projection !== 'final-raster' || visibleMark === 'off' || !forensicRequested,
  }
}
