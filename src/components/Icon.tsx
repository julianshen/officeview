/**
 * The single icon contract for viewer chrome (spec §6.3): a Lucide icon,
 * decorative only (the owning control carries the accessible label), tagged
 * with `data-icon` so tests select `svg[data-icon="…"]`. Lazy-only: never
 * import this, or any lucide-react icon, from eager code (scripts/check-eager.ts).
 */
import type { LucideIcon } from 'lucide-react'
import type { ReactElement } from 'react'

/** Icon names in use. Tasks 3 and 5 extend the union. */
export type IconName = 'chevron-up' | 'chevron-down' | 'plus' | 'minus' | 'search' | 'copy' | 'check' | 'ban' | 'command'

export function Icon({ icon: Glyph, name, size = 16, className }: { icon: LucideIcon; name: IconName; size?: 14 | 16 | 20; className?: string }): ReactElement {
  return <Glyph data-icon={name} size={size} strokeWidth={2} aria-hidden="true" focusable="false" className={className} />
}
