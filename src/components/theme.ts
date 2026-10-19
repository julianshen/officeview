/**
 * Theme root attributes (spec §6.2, §6.8). Eager and small: it must not
 * import src/styles/tokens.ts. The missing-stylesheet check lives in
 * styleWarning.ts, which only the lazy chunks import.
 */

export type ThemeMode = 'light' | 'dark' | 'system'

/** Attributes for a viewer's outer wrapper: the theme root every token and
 * scoped rule hangs off. */
export function themeRootProps(mode: ThemeMode | undefined, instance: string): { 'data-ov-theme-root': ''; 'data-ov-theme': ThemeMode; 'data-ov-instance': string } {
  return { 'data-ov-theme-root': '', 'data-ov-theme': mode ?? 'system', 'data-ov-instance': instance }
}
