/** Dev-only input acknowledgment mutation rules. */
export type AckKind = 'status' | 'dialog-open' | 'dialog-close' | 'apply' | 'find'
export const ACK_ATTRIBUTES = ['aria-busy', 'aria-sort', 'aria-disabled', 'hidden', 'aria-hidden', 'style']
function removesDialog(mutation: MutationRecord): boolean {
  return mutation.type === 'childList' && [...mutation.removedNodes].some(node => node instanceof Element && (node.matches('[role=dialog]') || node.querySelector('[role=dialog]') !== null))
}
/** The restyled popup becomes inert at exit-start; its 120 ms visual fade
 * may retain the DOM afterwards. Observe the closed state, not animation cleanup.
 * Require both signals on the dialog itself so unrelated hiding cannot qualify. */
function exitsDialog(mutation: MutationRecord): boolean {
  return mutation.type === 'attributes' && (mutation.attributeName === 'aria-hidden' || mutation.attributeName === 'style')
    && mutation.target instanceof HTMLElement && mutation.target.matches('[role=dialog]')
    && mutation.target.getAttribute('aria-hidden') === 'true' && mutation.target.style.pointerEvents === 'none'
}
export function qualifiesAck(kind: AckKind, mutation: MutationRecord): boolean {
  const el = mutation.target instanceof Element ? mutation.target : mutation.target.parentElement
  if (!el) return false
  if (mutation.type === 'attributes' && (mutation.attributeName === 'aria-busy' || mutation.attributeName === 'aria-sort' || mutation.attributeName === 'aria-disabled')) return true
  const inStatus = el.closest('[role=status],[role=progressbar],[role=alert]') !== null
  switch (kind) {
    case 'status': return inStatus || el.closest('[role=progressbar]') !== null
    case 'dialog-open': return mutation.type === 'childList' && [...mutation.addedNodes].some(node => node instanceof Element && (node.matches('[role=dialog]') || node.querySelector('[role=dialog]') !== null))
    case 'dialog-close': return removesDialog(mutation) || exitsDialog(mutation)
    // Applying from a dialog: the dialog closing, or a status/aria change, acknowledges it.
    case 'apply': return removesDialog(mutation) || inStatus
    case 'find': return inStatus || (el.closest('[data-officeview-ui]') !== null && mutation.type !== 'attributes')
  }
}
