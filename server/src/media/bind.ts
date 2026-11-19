/**
 * Piece 5 (plan D2, D3, D21, X13): media placeholders, identities and the per-session binding.
 *
 * The show-plan caches (L1/L2) are shared across sessions, so a cached plan names its media by document-level
 * placeholders: `@a<k>` for a spid-less audio use (an effect sound, the transition sound) and `@a<k>.<spid>` for a
 * shape's own audio, `k` being the part's index in the document's sorted media part paths (D21; a part-less sound,
 * a `builtIn:` one or an external link, gets a synthetic `k` at or past the part count, refused). At send time
 * `bindMedia` replaces each one with the session's `mediaId = base64url(HMAC-SHA256(sessionMediaKey, identity))[0, 22)`,
 * in the records, every `target.snd` (condition targets included) and the transition's `sound.start.mediaId`.
 *
 * Identity: `audio|k` (spid-less: a reused effect sound is one encode and one record per slide, X13) or
 * `audio|k|unit|spid` (a shape audio: the player keys shape media by spid). `unit` is the plan's own slide.
 */
import { createHmac } from 'node:crypto'
import { MEDIA_PLACEHOLDER, type ShowPlanRecord } from '../../../src/stream/show-plan'
import type { TimeNode } from '../../../src/slideshow/types'

export type MediaRef = { kind: 'audio'; k: number; unit?: number; spid?: string /* a shape audio's own record */ }

/** The placeholder a cached plan holds for `ref` (the unit is the plan's own). Throws past the D2 grammar. */
export function placeholderOf(ref: Pick<MediaRef, 'k' | 'spid'>): string {
  const s = ref.spid === undefined ? `@a${ref.k}` : `@a${ref.k}.${ref.spid}`
  if (!Number.isInteger(ref.k) || ref.k < 0 || !MEDIA_PLACEHOLDER.test(s)) throw new RangeError(`media: no placeholder for k=${ref.k} spid=${ref.spid}`)
  return s
}

/** The ref a placeholder of slide `unit` names; undefined for anything else. */
export function parsePlaceholder(s: string, unit: number): MediaRef | undefined {
  if (!MEDIA_PLACEHOLDER.test(s)) return undefined
  const [k, spid] = s.slice(2).split('.')
  return spid === undefined ? { kind: 'audio', k: Number(k) } : { kind: 'audio', k: Number(k), unit, spid }
}

/** D2's identity string, over `k` (D21), so `ws.ts` binds without knowing part paths. */
export function mediaIdentity(ref: MediaRef): string {
  if (ref.spid === undefined) return `audio|${ref.k}`
  if (ref.unit === undefined) throw new RangeError('media: a shape audio identity needs its unit')
  return `audio|${ref.k}|${ref.unit}|${ref.spid}`
}

/** `base64url(HMAC-SHA256(key, identity))[0, 22)`: 132 bits, the client's `MEDIA_ID`. */
export function sessionMediaId(key: Uint8Array, identity: string): string {
  return createHmac('sha256', key).update(identity).digest('base64url').slice(0, 22)
}

/**
 * `plan` with every placeholder replaced by `ids(ref)` (pure: `plan` is not changed). Throws on a media id or
 * `snd` target that is not a placeholder: an unbound plan never reaches a session.
 */
export function bindMedia(plan: ShowPlanRecord, ids: (ref: MediaRef) => string): ShowPlanRecord {
  const unit = plan.slide.index
  const memo = new Map<string, string>()
  const bind = (s: string): string => {
    let id = memo.get(s)
    if (id === undefined) {
      const ref = parsePlaceholder(s, unit)
      if (!ref) throw new Error('media: not a placeholder')
      memo.set(s, id = ids(ref))
    }
    return id
  }
  const target = <T>(t: T): T => (t && typeof t === 'object' && 'snd' in t ? { ...t, snd: bind((t as { snd: string }).snd) } : t)
  const node = (n: TimeNode, depth: number): TimeNode => {
    if (depth > 256) throw new Error('media: timing tree too deep to bind')
    const ctn = { ...n.ctn, start: n.ctn.start.map((c) => (c.target ? { ...c, target: target(c.target) } : c)), end: n.ctn.end.map((c) => (c.target ? { ...c, target: target(c.target) } : c)), ...(n.ctn.endSync?.target ? { endSync: { ...n.ctn.endSync, target: target(n.ctn.endSync.target) } } : {}) }
    if ('children' in n) return { ...n, ctn, children: n.children.map((c) => node(c, depth + 1)) }
    return { ...n, ctn, target: target(n.target) } as TimeNode
  }
  const timing = plan.show.timing && !('refused' in plan.show.timing) ? { ...plan.show.timing, root: node(plan.show.timing.root, 0) } : plan.show.timing
  const tr = plan.show.transition
  const start = tr?.sound?.start
  const transition = tr && start ? { ...tr, sound: { ...tr.sound, start: { ...start, mediaId: bind(start.mediaId) } } } : tr
  return {
    ...plan,
    show: { ...plan.show, timing, ...(transition ? { transition } : {}) },
    ...(plan.media ? { media: plan.media.map((r) => ({ ...r, mediaId: bind(r.mediaId) })) } : {}),
  }
}
