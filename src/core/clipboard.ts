/**
 * Copying text to the clipboard.
 *
 * The async Clipboard API is the primary path, but it needs a secure context
 * and can be blocked by permissions, so we fall back to the legacy
 * `execCommand('copy')` on a hidden textarea — which still works in older
 * WebViews and on plain-http origins.
 */

export type CopyOutcome = 'clipboard' | 'fallback' | 'failed'

/**
 * Copy text, preferring the async Clipboard API and falling back to a hidden
 * textarea + execCommand (which still works in older WebViews and on plain
 * http origins). Reports which path was used so callers can tell the user when
 * copying did not work — silently doing nothing is the worst outcome.
 */
export async function copyText(text: string): Promise<CopyOutcome> {
  if (typeof navigator !== 'undefined' && navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(text)
      return 'clipboard'
    } catch {
      // fall through to the legacy path
    }
  }
  if (typeof document === 'undefined') return 'failed'
  const area = document.createElement('textarea')
  area.value = text
  // keep it out of view and out of the tab order, but still selectable
  area.setAttribute('readonly', '')
  area.style.position = 'fixed'
  area.style.top = '0'
  area.style.left = '-9999px'
  area.style.opacity = '0'
  document.body.appendChild(area)
  const previous = document.activeElement as HTMLElement | null
  area.select()
  area.setSelectionRange(0, text.length)
  let ok = false
  try {
    ok = typeof document.execCommand === 'function' && document.execCommand('copy')
  } catch {
    ok = false
  }
  document.body.removeChild(area)
  previous?.focus?.()
  return ok ? 'fallback' : 'failed'
}