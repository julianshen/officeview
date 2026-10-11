/** Includes inherited editability and UI descendants such as an icon inside a
 * header button. Grid roots own their own gestures instead of legacy pan/zoom. */
export function isViewerUiTarget(target: EventTarget | null): boolean {
  const element = target instanceof Element ? target : target instanceof Node ? target.parentElement : null
  return Boolean(element?.closest('input, textarea, select, button, [contenteditable]:not([contenteditable="false"]), [data-officeview-ui], [data-officeview-grid], [role="dialog"]'))
}
