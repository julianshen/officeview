/**
 * Slicer and timeline styles (Rulings A, Q, R, K). A style is two lookups, the
 * same for built-in and custom styles:
 * - the frame (`wholeTable`, `headerRow`) from the same-named table style —
 *   `builtinTableStyle(name)` or a custom `tableStyle` in styles.xml;
 * - the item states from its `SlicerStyleDef` — `builtinSlicerStyle(name)` or
 *   the x14 `slicerStyles` of styles.xml (`parseStylesExtras().slicerStyles`).
 * An unknown name (TimeSlicerStyleOther1–2 among them: Excel has no such
 * timeline styles) paints the default style and reports a diagnostic. A custom
 * style layers over the default built-in, so an element it leaves out keeps the
 * default's value. Hover states are never painted (view-only, Ruling R).
 */
import type { ThemeContext } from '../../../drawing/style'
import type { ColorSpec, DxfStyle, SlicerStyleDef, StylesExtras, TableStyleDef } from '../../tableStyles/model'
import { builtinSlicerStyle, builtinTableStyle } from '../../tableStyles/builtins'
import { resolveColorSpec } from '../../tableStyles/resolve'
import { applyTint } from '../../sheetColors'
import type { ControlDiagnostic, SlicerAppearance, SlicerItemPaint, SlicerItemState, TimelineAppearance } from './model'

export const DEFAULT_SLICER_STYLE = 'SlicerStyleLight1'
export const DEFAULT_TIMELINE_STYLE = 'TimeSlicerStyleLight1'
const ITEM_STATES: readonly SlicerItemState[] = ['selectedItemWithData', 'unselectedItemWithData', 'selectedItemWithNoData', 'unselectedItemWithNoData']

/** Timeline parts no SlicerStyleDef element names, read from the references (docs/validation/xlsx-pivots.md
 * "Timeline geometry"): thumb #AFB3B7 with a #BCBCBC edge, year labels #989898 (darkest glyph pixel), level picker
 * ≈ #A1A1A1 [measure, refs/timeline-light1-months.png]. The divider after a selected period, the scrollbar track and
 * the selection label follow the frame accent: tint 0.4 (#95B3D7 light1, #DA9694 dark2), tint 0.8 (#DAE4F1, #F1DBDA)
 * and tint −0.25 [measure]. The label's darkest glyph pixels are that shade at one coverage: #376092 at 90 % over white
 * is #4B709D (measure-timeline-selection), at 84 % #577AA4 (light1-months); #E36C09 at 90 % is #E47820 (dark6). */
const TIMELINE_FIXED = { thumb: '#AFB3B7', thumbEdge: '#BCBCBC', yearLabel: '#989898', timeLevel: '#A1A1A1' }

const hex = (argb: string | undefined): string | undefined => argb && /^[0-9A-F]{8}$/i.test(argb) ? `#${argb.slice(2).toUpperCase()}` : undefined
const colour = (spec: ColorSpec | undefined, theme: ThemeContext): string | undefined => hex(resolveColorSpec(spec, theme))
const fillOf = (dxf: DxfStyle | undefined, theme: ThemeContext) => dxf?.fill && dxf.fill.pattern !== 'none' ? colour(dxf.fill.bg ?? dxf.fill.fg, theme) : undefined
const borderOf = (dxf: DxfStyle | undefined, theme: ThemeContext, sides: readonly ('left' | 'top' | 'right' | 'bottom')[] = ['left', 'top', 'right', 'bottom']) => {
  for (const side of sides) { const b = dxf?.border?.[side]; if (b && b.style !== 'none' && b.color) return colour(b.color, theme) }
  return undefined
}

interface StyleParts { frame?: TableStyleDef; items?: SlicerStyleDef }
/** The style's two lookups, custom styles first; undefined when neither names it as a slicer/timeline style. */
function lookup(name: string, extras: StylesExtras): StyleParts | undefined {
  const items = extras.slicerStyles.get(name) ?? builtinSlicerStyle(name)
  if (!items) return undefined
  return { frame: extras.tableStyles.get(name) ?? builtinTableStyle(name), items }
}
function pick(name: string | undefined, fallback: string, extras: StylesExtras, feature: string, report?: (diagnostic: ControlDiagnostic) => void): { style: string; parts: StyleParts; base: StyleParts } {
  const base = lookup(fallback, extras) ?? {}
  const wanted = name || fallback
  const parts = lookup(wanted, extras)
  if (parts) return { style: wanted, parts, base }
  report?.({ kind: 'unsupported', feature, message: `Style ${wanted} is not a known slicer or timeline style; ${fallback} is painted instead` })
  return { style: fallback, parts: base, base }
}

function itemPaint(state: SlicerItemState, parts: StyleParts, theme: ThemeContext, inherited: SlicerItemPaint): SlicerItemPaint {
  const dxf = parts.items?.elements[state]?.dxf
  return {
    fill: fillOf(dxf, theme) ?? inherited.fill,
    border: borderOf(dxf, theme) ?? inherited.border,
    text: colour(dxf?.font?.color, theme) ?? inherited.text,
    bold: dxf?.font?.bold ?? inherited.bold,
  }
}

function frameOf(parts: StyleParts, theme: ThemeContext, inherited: { border: string; background: string; header: SlicerAppearance['header'] }) {
  const whole = parts.frame?.elements.wholeTable?.dxf, header = parts.frame?.elements.headerRow?.dxf
  return {
    border: borderOf(whole, theme) ?? inherited.border,
    background: fillOf(whole, theme) ?? inherited.background,
    header: {
      fill: fillOf(header, theme) ?? inherited.header.fill,
      text: colour(header?.font?.color, theme) ?? inherited.header.text,
      bold: header?.font?.bold ?? inherited.header.bold,
      separator: borderOf(header, theme, ['bottom']) ?? inherited.header.separator,
    },
  }
}

const ROOT_ITEM: SlicerItemPaint = { fill: '#FFFFFF', border: '#CCCCCC', text: '#000000', bold: false }
const ROOT_FRAME = { border: '#4F81BD', background: '#FFFFFF', header: { fill: '#FFFFFF', text: '#000000', bold: true, separator: '#4F81BD' } }

function slicerFrom(parts: StyleParts, theme: ThemeContext, inherited?: SlicerAppearance): SlicerAppearance {
  const frame = frameOf(parts, theme, inherited ?? ROOT_FRAME)
  const items = Object.fromEntries(ITEM_STATES.map(state => [state, itemPaint(state, parts, theme, inherited?.items[state] ?? ROOT_ITEM)])) as Record<SlicerItemState, SlicerItemPaint>
  return { ...frame, items, fontFamily: theme.fonts.minor.latin || 'Calibri' }
}

/** The appearance a slicer paints with, and the style name actually used. */
export function slicerAppearance(name: string | undefined, extras: StylesExtras, theme: ThemeContext, report?: (diagnostic: ControlDiagnostic) => void): { style: string; appearance: SlicerAppearance } {
  const { style, parts, base } = pick(name, DEFAULT_SLICER_STYLE, extras, 'slicer-style-unknown', report)
  const defaults = slicerFrom(base, theme)
  return { style, appearance: parts === base ? defaults : slicerFrom(parts, theme, defaults) }
}

/** The appearance a timeline paints with, and the style name actually used. */
export function timelineAppearance(name: string | undefined, extras: StylesExtras, theme: ThemeContext, report?: (diagnostic: ControlDiagnostic) => void): { style: string; appearance: TimelineAppearance } {
  const { style, parts, base } = pick(name, DEFAULT_TIMELINE_STYLE, extras, 'timeline-style-unknown', report)
  const defaults = slicerFrom(base, theme)
  const s = parts === base ? defaults : slicerFrom(parts, theme, defaults)
  const accent = s.border.slice(1)
  const tinted = (tint: number) => `#${applyTint(accent, tint)}`
  return {
    appearance: {
      border: s.border,
      background: s.background,
      header: { text: s.header.text, bold: s.header.bold, separator: s.header.separator },
      selectionLabel: tinted(-0.25),
      timeLevel: TIMELINE_FIXED.timeLevel,
      periodLabel: s.items.selectedItemWithData.text,
      yearLabel: TIMELINE_FIXED.yearLabel,
      selected: s.items.selectedItemWithData.fill,
      unselected: { fill: s.items.unselectedItemWithData.fill, border: s.items.unselectedItemWithData.border },
      divider: tinted(0.4),
      track: tinted(0.8),
      thumb: TIMELINE_FIXED.thumb,
      thumbEdge: TIMELINE_FIXED.thumbEdge,
      fontFamily: s.fontFamily,
    },
    style,
  }
}
