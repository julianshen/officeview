/** The §6.2 rule-3 MANIFEST JSON (inside a kind-1 OVSD frame) and its
 * fail-closed client validation: numbers range-checked, unit sizes 1..8192,
 * enums closed, token ≤ 512 chars, unknown fields ignored. */
export interface Manifest {
  v: 1
  docRef: string
  kind: 'docx' | 'xlsx' | 'pptx'
  name: string
  units: { count: number; list: Array<{ w: number; h: number }> }
  layoutFingerprint: string
  session: { sid: string; token: string; expiresSec: number; resume: 'session' }
  policy: { contentPolicy: 'visible-only' | 'full'; protection: 'final-raster' | 'culled-commands' | 'commands'; semantic: 'visible-only' | 'off'; contentPolicyId: string; viewerMarkId: string }
  can: { search: boolean; copy: boolean; print: boolean; textA11y: boolean; present: boolean; notes: boolean; links: boolean }
  recorder: { commands: 'ovsd-1'; textMode: 'text'; scalesBuckets: number[]; requiredFaces: Array<{ id: string; hash: string; alias?: string }>; projection: 'final-raster' | 'culled-commands' | 'commands' }
  limits: { maxUnitPx: number; maxAssetBytes: number; wsFrameBytes: number; tilePx: number }
  notices: string[]
}
export type ManifestResult = { ok: true; manifest: Manifest } | { ok: false; field: string }

export const MANIFEST_CAP = 256 * 1024
const MAX_UNITS = 100_000, MAX_FACES = 64, MAX_NOTICES = 64

class Bad extends Error { constructor(readonly field: string) { super(field) } }
const obj = (v: unknown, f: string): Record<string, unknown> => { if (!v || typeof v !== 'object' || Array.isArray(v)) throw new Bad(f); return v as Record<string, unknown> }
const str = (v: unknown, f: string, max = 256, re?: RegExp): string => { if (typeof v !== 'string' || v.length > max || (re && !re.test(v))) throw new Bad(f); return v }
const int = (v: unknown, f: string, min: number, max: number): number => { if (!Number.isInteger(v) || (v as number) < min || (v as number) > max) throw new Bad(f); return v as number }
const bool = (v: unknown, f: string): boolean => { if (typeof v !== 'boolean') throw new Bad(f); return v }
const oneOf = <T extends string>(v: unknown, f: string, set: readonly T[]): T => { if (!set.includes(v as T)) throw new Bad(f); return v as T }
const arr = (v: unknown, f: string, max: number): unknown[] => { if (!Array.isArray(v) || v.length > max) throw new Bad(f); return v }

const HEX = /^[0-9a-f]+$/
const PROJ = ['final-raster', 'culled-commands', 'commands'] as const

export function validateManifest(json: unknown): ManifestResult {
  try {
    const m = obj(json, '$')
    if (m.v !== 1) throw new Bad('v')
    const units = obj(m.units, 'units')
    const list = arr(units.list, 'units.list', MAX_UNITS).map((u, i) => { const o = obj(u, `units.list[${i}]`); return { w: int(o.w, `units.list[${i}].w`, 1, 8192), h: int(o.h, `units.list[${i}].h`, 1, 8192) } })
    if (int(units.count, 'units.count', 0, MAX_UNITS) !== list.length) throw new Bad('units.count')
    const s = obj(m.session, 'session'), p = obj(m.policy, 'policy'), c = obj(m.can, 'can'), r = obj(m.recorder, 'recorder'), l = obj(m.limits, 'limits')
    const manifest: Manifest = {
      v: 1,
      docRef: str(m.docRef, 'docRef', 24, /^[0-9a-f]{24}$/),
      kind: oneOf(m.kind, 'kind', ['docx', 'xlsx', 'pptx'] as const),
      name: str(m.name, 'name', 512),
      units: { count: list.length, list },
      layoutFingerprint: str(m.layoutFingerprint, 'layoutFingerprint', 128, HEX),
      session: { sid: str(s.sid, 'session.sid', 12, /^[0-9a-f]{12}$/), token: str(s.token, 'session.token', 512), expiresSec: int(s.expiresSec, 'session.expiresSec', 1, 86_400), resume: oneOf(s.resume, 'session.resume', ['session'] as const) },
      policy: {
        contentPolicy: oneOf(p.contentPolicy, 'policy.contentPolicy', ['visible-only', 'full'] as const),
        protection: oneOf(p.protection, 'policy.protection', PROJ),
        semantic: oneOf(p.semantic, 'policy.semantic', ['visible-only', 'off'] as const),
        contentPolicyId: str(p.contentPolicyId, 'policy.contentPolicyId', 64),
        viewerMarkId: str(p.viewerMarkId, 'policy.viewerMarkId', 64),
      },
      can: Object.fromEntries((['search', 'copy', 'print', 'textA11y', 'present', 'notes', 'links'] as const).map((k) => [k, bool(c[k], `can.${k}`)])) as Manifest['can'],
      recorder: {
        commands: oneOf(r.commands, 'recorder.commands', ['ovsd-1'] as const),
        textMode: oneOf(r.textMode, 'recorder.textMode', ['text'] as const),
        scalesBuckets: arr(r.scalesBuckets, 'recorder.scalesBuckets', 8).map((b, i) => int(b, `recorder.scalesBuckets[${i}]`, 1, 8)),
        requiredFaces: arr(r.requiredFaces, 'recorder.requiredFaces', MAX_FACES).map((f, i) => {
          const o = obj(f, `recorder.requiredFaces[${i}]`)
          return { id: str(o.id, `recorder.requiredFaces[${i}].id`, 64, /^[a-z0-9-]+$/), hash: str(o.hash, `recorder.requiredFaces[${i}].hash`, 24, /^[0-9a-f]{24}$/), ...(o.alias !== undefined ? { alias: str(o.alias, `recorder.requiredFaces[${i}].alias`, 256) } : {}) }
        }),
        projection: oneOf(r.projection, 'recorder.projection', PROJ),
      },
      limits: { maxUnitPx: int(l.maxUnitPx, 'limits.maxUnitPx', 1, 8192), maxAssetBytes: int(l.maxAssetBytes, 'limits.maxAssetBytes', 1, 1 << 30), wsFrameBytes: int(l.wsFrameBytes, 'limits.wsFrameBytes', 1024, 1 << 20), tilePx: int(l.tilePx, 'limits.tilePx', 16, 8192) },
      notices: arr(m.notices, 'notices', MAX_NOTICES).map((n, i) => str(n, `notices[${i}]`, 128)),
    }
    return { ok: true, manifest }
  } catch (e) {
    if (e instanceof Bad) return { ok: false, field: e.field }
    throw e
  }
}

const enc = new TextEncoder()
export function encodeManifest(m: Manifest): Uint8Array {
  const b = enc.encode(JSON.stringify(m))
  if (b.length > MANIFEST_CAP) throw new Error('manifest exceeds 256 KiB')
  return b
}
