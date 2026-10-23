/** Target identity (slideshow spec 6.2/6.3): one string per animated target. */
import type { TargetKey, TargetRef } from './types'

/**
 * Shape `sp:<spid>`, text range `sp:<spid>/p:<st>-<end>` or `/c:<st>-<end>`,
 * sub-target `sp:<spid>#<sub>`, then `bg`, `sld`, `snd:<rId>`. Ids are
 * URI-encoded, so a hostile spid cannot forge another target's key.
 */
export function targetKey(ref: TargetRef): TargetKey {
  if ('spid' in ref) {
    let key = `sp:${encodeURIComponent(ref.spid)}`
    if (ref.sub) key += `#${ref.sub}`
    if (ref.text) key += `/${ref.text.kind === 'p' ? 'p' : 'c'}:${ref.text.st}-${ref.text.end}`
    return key
  }
  if ('bg' in ref) return 'bg'
  if ('sld' in ref) return 'sld'
  return `snd:${encodeURIComponent(ref.snd)}`
}
