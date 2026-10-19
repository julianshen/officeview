/**
 * The missing-stylesheet check (spec §5, R10). Runs after first paint from the
 * lazy chunks (PagedChrome, InteractiveSheetDoc), so it stays out of the eager
 * entry.
 */

let warned = false

/** Warns once per page when officeview.css is not loaded (its --ov-sentinel is
 * missing on the theme root). A runtime check on purpose, not gated on
 * import.meta.env: lib mode would bake PROD in and silence it for every host
 * (spec §5, R10). Call it from an effect. */
export function warnIfStylesMissing(themeRoot: Element): void {
  if (warned) return
  if (getComputedStyle(themeRoot).getPropertyValue('--ov-sentinel').trim() !== '1') {
    warned = true
    console.warn('[officeview] style.css not loaded: import "officeview/style.css"')
  }
}

export function resetStyleWarningForTests(): void {
  warned = false
}
