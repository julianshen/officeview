/**
 * Theme token names and defaults (spec §6.1, §6.8). Each colour token is the
 * CSS custom property `--ov-color-<name>` on the theme root; officeview.css
 * defines the same values (tests/ui-theme-foundation.test.tsx keeps the two in
 * step). Hosts re-brand by overriding the properties from any selector that
 * targets [data-ov-theme-root].
 *
 * Not eager: chrome reads tokens through CSS, and only the lazy canvas
 * painters need these values in script (plan Global Constraints, eager graph).
 */

/** Chrome colours: DOM toolbar, menus, headers and backdrop. `gridline` themes
 * only the DOM header dividers, never the canvas gridlines. */
const CHROME_TOKENS = [
  'surface', 'chrome', 'field', 'border', 'text', 'muted', 'icon', 'hover',
  'accent', 'on-accent', 'header', 'header-text', 'header-active', 'gridline', 'backdrop',
] as const

/** Overlay colours drawn on paper (white in both themes, decision B). */
export const CANVAS_TOKENS = [
  'selection', 'selection-fill', 'find', 'find-active', 'text-selection', 'pending', 'pending-text', 'freeze',
] as const

export type CanvasToken = typeof CANVAS_TOKENS[number]

export const COLOR_TOKENS = [...CHROME_TOKENS, ...CANVAS_TOKENS] as const

export type ColorToken = typeof COLOR_TOKENS[number]

const CANVAS_LIGHT: Record<CanvasToken, string> = {
  'selection': '#107c41',
  'selection-fill': 'rgba(16,124,65,.12)',
  'find': 'rgba(255,214,0,.45)',
  'find-active': '#ca5010',
  'text-selection': 'rgba(64,128,255,.35)',
  'pending': '#eceff3',
  'pending-text': '#605e5c',
  'freeze': '#5f6368',
}

export const THEME_DEFAULTS: { light: Record<ColorToken, string>; dark: Record<ColorToken, string> } = {
  light: {
    'surface': '#ffffff',
    'chrome': '#faf9f8',
    'field': '#ffffff',
    'border': '#e1dfdd',
    'text': '#201f1e',
    'muted': '#605e5c',
    'icon': '#323130',
    'hover': '#f3f2f1',
    'accent': '#0f6cbd',
    'on-accent': '#ffffff',
    'header': '#f3f2f1',
    'header-text': '#323130',
    'header-active': '#deecf9',
    'gridline': '#edebe9',
    'backdrop': '#f0efee',
    ...CANVAS_LIGHT,
  },
  dark: {
    'surface': '#292827',
    'chrome': '#201f1e',
    'field': '#1b1a19',
    'border': '#3b3a39',
    'text': '#f3f2f1',
    'muted': '#a19f9d',
    'icon': '#d2d0ce',
    'hover': '#3b3a39',
    'accent': '#479ef5',
    // White on #479ef5 is about 2.8:1; near-black clears 4.5:1.
    'on-accent': '#1b1a19',
    'header': '#201f1e',
    'header-text': '#d2d0ce',
    'header-active': '#1c3149',
    'gridline': '#3b3a39',
    'backdrop': '#161514',
    ...CANVAS_LIGHT,
    'pending': '#e1dfdd',
  },
}
