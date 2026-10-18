/**
 * Theme root attributes and the missing-stylesheet check (spec §6.2, §6.8).
 * Eager and small: it must not import src/styles/tokens.ts.
 */

export type ThemeMode = 'light' | 'dark' | 'system'

/** Attributes for a viewer's outer wrapper: the theme root every token and
 * scoped rule hangs off. */
export function themeRootProps(mode: ThemeMode | undefined, instance: string): { 'data-ov-theme-root': ''; 'data-ov-theme': ThemeMode; 'data-ov-instance': string } {
  return { 'data-ov-theme-root': '', 'data-ov-theme': mode ?? 'system', 'data-ov-instance': instance }
}

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
