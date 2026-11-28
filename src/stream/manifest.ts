/** The §6.2 rule-3 MANIFEST JSON (inside a kind-1 OVSD frame) and its
 * fail-closed client validation: numbers range-checked, unit sizes 1..8192,
 * enums closed, token ≤ 512 chars, unknown fields ignored. */
import { MAX_SHEETS, SHEET_NAME_CHARS, TILE_BATCH_MAX, TILE_DPRS, TILE_PALETTES, TILE_PX, TILE_ZOOM_STEPS } from './sheet-wire'

export interface Manifest {
  v: 1
  docRef: string
  kind: 'docx' | 'xlsx' | 'pptx'
  name: string
  units: { count: number; list: ManifestUnit[] }
  layoutFingerprint: string
  session: { sid: string; token: string; expiresSec: number; resume: 'session' }
  policy: { contentPolicy: 'visible-only' | 'full'; protection: 'final-raster' | 'culled-commands' | 'commands'; semantic: 'visible-only' | 'off'; contentPolicyId: string; viewerMarkId: string }
  /** `media` (piece 5): audio media plays. Absent = off (a piece-4 server); a piece-5 server always sends it. */
  can: { search: boolean; copy: boolean; print: boolean; textA11y: boolean; present: boolean; notes: boolean; links: boolean; media?: boolean }
  /** Service caps. `copy` null = copy off (Stage 3a); `print` null = print off (Stage 3b); a null cap = uncapped.
   * `media` (piece 5) is keyed by kind: this piece defines `audio` only, and any other member is ignored, so a later
   * piece can add `video` beside it. Null or absent = media off; required when `can.media`. */
  caps: { copy: null | { maxParagraphs: number | null; maxSessionParagraphs: number | null }; a11y: 'off' | 'unit' | 'focused'; print: null | { maxPages: number | null; maxSessionPages: number | null }; media?: null | MediaCaps }
  recorder: { commands: 'ovsd-1'; textMode: 'text'; scalesBuckets: number[]; requiredFaces: Array<{ id: string; hash: string; alias?: string }>; projection: 'final-raster' | 'culled-commands' | 'commands' }
  limits: { maxUnitPx: number; maxAssetBytes: number; wsFrameBytes: number; tilePx: number }
  notices: string[]
  /** Stage 4a: `'tiles'` for an XLSX session the server serves as tiles (the client sent `sheet-tiles@1`). Absent =
   * units mode, and then `sheets` and `tiles` are absent too (ignored if sent). */
  sheetMode?: 'tiles'
  /** Tiles mode: the visible sheets in tab order; `sheetId` is the decimal index `"0".."n-1"` (R12). */
  sheets?: ManifestSheet[]
  /** Tiles mode: the fixed tile grid (S4 §6). */
  tiles?: ManifestTiles
  /** Stage 4b (P1): `1` iff the client sent `sheet-view@1` and the server serves the session in tiles mode with the 4b ops
   * (find, copy, a11y text, fit, resize). Valid only with `sheetMode: 'tiles'`; absent for a 4a server or session. */
  sheetView?: 1
  /** Stage 4c (P1): `1` iff the server serves sort, filter, value lists, range detection and the view fetch to this
   * session. Valid only with `sheetView: 1`; absent for a 4b server. */
  sheetSort?: 1
  /** Stage 4c (P1): `true` iff typed number and date filter conditions are served (contentPolicy `'full'`). Valid only
   * with `sheetSort`. */
  typedFilters?: true
}
export interface ManifestSheet { sheetId: string; name: string }
export interface ManifestTiles { px: 512; zoomSteps: number[]; dprs: [1, 2]; palettes: ['light', 'dark']; maxBatch: 64 }
/** `hidden`: a hidden slide under contentPolicy 'full' (it is a unit, but a show skips it). Only `true` is valid.
 * `anim`: the slide has timing, a transition, or both, so a show asks for its SHOW_PLAN (piece 4 D5). Only `true` is valid; without it the slide is a still. */
/** `media` (piece 5, D19): the slide ships an audio record or a transition sound, so a show asks for its SHOW_PLAN
 * even without `anim`. Only `true` is valid. */
export interface ManifestUnit { w: number; h: number; hidden?: true; anim?: true; media?: true }
/** Piece 5 (spec §6): the audio limits, verbatim (30 min, 32 MiB). */
export const MEDIA_AUDIO_CAPS = Object.freeze({ maxSeconds: 1800, maxBytes: 33554432 } as const)
export interface MediaCaps { audio: { maxSeconds: 1800; maxBytes: 33554432 } }
export type ManifestResult = { ok: true; manifest: Manifest } | { ok: false; field: string }

export const MANIFEST_CAP = 256 * 1024
const MAX_UNITS = 100_000, MAX_FACES = 64, MAX_NOTICES = 64

class Bad extends Error { constructor(readonly field: string) { super(field) } }
const obj = (v: unknown, f: string): Record<string, unknown> => { if (!v || typeof v !== 'object' || Array.isArray(v)) throw new Bad(f); return v as Record<string, unknown> }
const str = (v: unknown, f: string, max = 256, re?: RegExp): string => { if (typeof v !== 'string' || v.length > max || (re && !re.test(v))) throw new Bad(f); return v }
const int = (v: unknown, f: string, min: number, max: number): number => { if (!Number.isInteger(v) || (v as number) < min || (v as number) > max) throw new Bad(f); return v as number }
const bool = (v: unknown, f: string): boolean => { if (typeof v !== 'boolean') throw new Bad(f); return v }
const literalTrue = (v: unknown, f: string): true => { if (v !== true) throw new Bad(f); return v }
const oneOf = <T extends string>(v: unknown, f: string, set: readonly T[]): T => { if (!set.includes(v as T)) throw new Bad(f); return v as T }
const cap = (v: unknown, f: string): number | null => v === null ? null : int(v, f, 1, Number.MAX_SAFE_INTEGER)
const printCaps = (o: Record<string, unknown>) => ({ maxPages: cap(o.maxPages, 'caps.print.maxPages'), maxSessionPages: cap(o.maxSessionPages, 'caps.print.maxSessionPages') })
const arr = (v: unknown, f: string, max: number): unknown[] => { if (!Array.isArray(v) || v.length > max) throw new Bad(f); return v }

/** `caps.media`: absent, null, or `{audio}` with exactly the spec values; other members are ignored (spec §4.1). */
function mediaCaps(v: unknown): MediaCaps | null | undefined {
  if (v === undefined || v === null) return v
  const a = obj(obj(v, 'caps.media').audio, 'caps.media.audio')
  if (a.maxSeconds !== MEDIA_AUDIO_CAPS.maxSeconds) throw new Bad('caps.media.audio.maxSeconds')
  if (a.maxBytes !== MEDIA_AUDIO_CAPS.maxBytes) throw new Bad('caps.media.audio.maxBytes')
  return { audio: { maxSeconds: MEDIA_AUDIO_CAPS.maxSeconds, maxBytes: MEDIA_AUDIO_CAPS.maxBytes } }
}

/** `tiles`: exactly the S4 §6 grid; anything else rejects the manifest. */
function tilesBlock(v: unknown): ManifestTiles {
  const t = obj(v, 'tiles')
  if (t.px !== TILE_PX) throw new Bad('tiles.px')
  const same = (a: unknown, b: readonly unknown[]) => Array.isArray(a) && a.length === b.length && a.every((x, i) => x === b[i])
  if (!same(t.zoomSteps, TILE_ZOOM_STEPS)) throw new Bad('tiles.zoomSteps')
  if (!same(t.dprs, TILE_DPRS)) throw new Bad('tiles.dprs')
  if (!same(t.palettes, TILE_PALETTES)) throw new Bad('tiles.palettes')
  if (t.maxBatch !== TILE_BATCH_MAX) throw new Bad('tiles.maxBatch')
  return { px: TILE_PX, zoomSteps: [...TILE_ZOOM_STEPS], dprs: [...TILE_DPRS], palettes: [...TILE_PALETTES], maxBatch: TILE_BATCH_MAX }
}
/** `sheets`: ≤ 256 entries, ids `"0".."n-1"` in sequence, names ≤ 31 characters. */
const sheetList = (v: unknown): ManifestSheet[] => arr(v, 'sheets', MAX_SHEETS).map((x, i) => {
  const o = obj(x, `sheets[${i}]`)
  const sheetId = str(o.sheetId, `sheets[${i}].sheetId`, 3, /^\d{1,3}$/)
  if (sheetId !== String(i)) throw new Bad(`sheets[${i}].sheetId`)
  return { sheetId, name: str(o.name, `sheets[${i}].name`, SHEET_NAME_CHARS) }
})

const HEX = /^[0-9a-f]+$/
const PROJ = ['final-raster', 'culled-commands', 'commands'] as const

export function validateManifest(json: unknown): ManifestResult {
  try {
    const m = obj(json, '$')
    if (m.v !== 1) throw new Bad('v')
    const units = obj(m.units, 'units')
    const list = arr(units.list, 'units.list', MAX_UNITS).map((u, i): ManifestUnit => {
      const o = obj(u, `units.list[${i}]`), size = { w: int(o.w, `units.list[${i}].w`, 1, 8192), h: int(o.h, `units.list[${i}].h`, 1, 8192) }
      if (o.hidden !== undefined && o.hidden !== true) throw new Bad(`units.list[${i}].hidden`)
      if (o.anim !== undefined && o.anim !== true) throw new Bad(`units.list[${i}].anim`)
      if (o.media !== undefined && o.media !== true) throw new Bad(`units.list[${i}].media`)
      return { ...size, ...(o.hidden ? { hidden: true as const } : {}), ...(o.anim ? { anim: true as const } : {}), ...(o.media ? { media: true as const } : {}) }
    })
    if (int(units.count, 'units.count', 0, MAX_UNITS) !== list.length) throw new Bad('units.count')
    const s = obj(m.session, 'session'), p = obj(m.policy, 'policy'), c = obj(m.can, 'can'), r = obj(m.recorder, 'recorder'), l = obj(m.limits, 'limits'), k = obj(m.caps, 'caps')
    if (m.sheetView !== undefined && m.sheetMode === undefined) throw new Bad('sheetView')
    if (m.sheetSort !== undefined && m.sheetView === undefined) throw new Bad('sheetSort')
    if (m.typedFilters !== undefined && m.sheetSort === undefined) throw new Bad('typedFilters')
    const kc = k.copy === null ? null : obj(k.copy, 'caps.copy')
    const canMedia = c.media === undefined ? undefined : bool(c.media, 'can.media')
    const media = mediaCaps(k.media)
    if (canMedia && !media) throw new Bad('caps.media')
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
      can: { ...Object.fromEntries((['search', 'copy', 'print', 'textA11y', 'present', 'notes', 'links'] as const).map((k) => [k, bool(c[k], `can.${k}`)])) as Manifest['can'], ...(canMedia !== undefined ? { media: canMedia } : {}) },
      caps: {
        copy: kc && { maxParagraphs: cap(kc.maxParagraphs, 'caps.copy.maxParagraphs'), maxSessionParagraphs: cap(kc.maxSessionParagraphs, 'caps.copy.maxSessionParagraphs') },
        a11y: oneOf(k.a11y, 'caps.a11y', ['off', 'unit', 'focused'] as const),
        print: k.print === null ? null : printCaps(obj(k.print, 'caps.print')),
        ...(media !== undefined ? { media } : {}),
      },
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
      ...(m.sheetMode === undefined ? {} : { sheetMode: oneOf(m.sheetMode, 'sheetMode', ['tiles'] as const), sheets: sheetList(m.sheets), tiles: tilesBlock(m.tiles), ...(m.sheetView === undefined ? {} : { sheetView: int(m.sheetView, 'sheetView', 1, 1) as 1 }), ...(m.sheetSort === undefined ? {} : { sheetSort: int(m.sheetSort, 'sheetSort', 1, 1) as 1 }), ...(m.typedFilters === undefined ? {} : { typedFilters: literalTrue(m.typedFilters, 'typedFilters') }) }),
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
