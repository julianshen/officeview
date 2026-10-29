/**
 * Workbook formatting metadata of a paged import (plan S5.4; spec §§4.2, 8.2,
 * 11.1): the resolved cell-style table (one entry per cellXfs index, bounded by
 * the styles part's small-part cap) and the date system. Paged pages keep only
 * a style index per cell; this table lets paint, copy, find and filter format
 * paged cells exactly as the object parser does for resident cells.
 */
import type { OfficePackage } from '../../core/zip'
import { attrs, getChildren, parseXml, type XmlNode } from '../../core/xml'
import { parseStyles, resolveXfStyle, workbookDigitWidth } from '../parse'
import { parseThemeContext } from '../../drawing/style'
import type { WorkbookFormatting } from '../view/contracts'
import type { SheetColorContext } from '../sheetColors'
import type { SlicerStyleDef, TableStyleDef } from '../tableStyles/model'
import { estimatePayloadBytes } from '../view/payloadBytes'
import { MAX_PAGED_TABLE_STYLE_DEFS_BYTES } from '../view/contracts'
import type { StylesExtras } from '../tableStyles/model'

export const DEFAULT_FORMATTING: WorkbookFormatting = Object.freeze({ dateSystem: '1900', styles: Object.freeze([]) })

/** A custom table, pivot or slicer style definition (not the `<tableStyles>`/`<slicerStyles>` lists Excel always writes). */
const CUSTOM_STYLE_DEFINITION = /<(?:[A-Za-z_][\w.-]*:)?(?:tableStyle|slicerStyle)[\s/>]/
/** Whether styles.xml defines a custom table, pivot or slicer style, so its extras parse with the workbook styles. */
export function stylesDefineCustomStyles(stylesText: string): boolean { return CUSTOM_STYLE_DEFINITION.test(stylesText) }

/** Same resolution as parse.ts (cellXfs → font/fill/border/numFmt). The styles extras (dxfs, custom table and
 * slicer styles) load lazily (spec §6.2): with `extras: 'auto'` they parse only when styles.xml defines a custom
 * style or a preferred name is given; the paged importer asks again (`'always'`) for the first sheet with a
 * styled table, a pivot, a slicer or a timeline. Without them `tableStyleExtras` is undefined. */
export async function workbookFormattingWithColorContext(workbookRoot: XmlNode, stylesText: string | undefined, themeText?: string, preferredTableStyleNames: ReadonlySet<string> = new Set(), preferredSlicerStyleNames: ReadonlySet<string> = new Set(), extras: 'auto' | 'always' = 'auto'): Promise<{ formatting: WorkbookFormatting; colorContext: SheetColorContext; tableStyleDefs: ReadonlyMap<string, TableStyleDef>; tableStyleExtras?: StylesExtras }> {
  const pr = getChildren(workbookRoot, 'workbookPr')[0]
  const date1904 = pr ? attrs(pr).date1904 : undefined
  const dateSystem = date1904 === '1' || date1904 === 'true' ? '1904' : '1900'
  // Theme colour references resolve against the workbook theme, as parse.ts does.
  let palette: Record<string, string> | undefined
  if (themeText !== undefined) {
    try { palette = parseThemeContext(themeText).palette } catch { /* optional theme */ }
  }
  if (stylesText === undefined) return { formatting: { dateSystem, styles: [] }, colorContext: palette ? { theme: palette } : {}, tableStyleDefs: new Map() }
  const root = parseXml(stylesText)
  const styles = await parseStyles({ xml: async (path: string) => path === 'xl/styles.xml' ? root : undefined } as unknown as OfficePackage, palette)
  const digitWidth = workbookDigitWidth(styles)
  const formatting: WorkbookFormatting = { dateSystem, styles: styles.xfs.map(xf => resolveXfStyle(styles, xf)), ...(digitWidth !== 7 ? { maxDigitWidthPx: digitWidth } : {}) }
  if (extras === 'auto' && !preferredTableStyleNames.size && !preferredSlicerStyleNames.size && !stylesDefineCustomStyles(stylesText)) {
    return { formatting, colorContext: styles.colorContext, tableStyleDefs: new Map() }
  }
  const { parseStylesExtras } = await import('../tableStyles/parseStylesExtras')
  const parsedExtras = parseStylesExtras(root, styles.colorContext, preferredTableStyleNames, preferredSlicerStyleNames)
  const parsedTableStyleDefs = parsedExtras.tableStyles
  const tableStyleDefs = new Map<string, TableStyleDef>()
  const slicerStyleDefs = new Map<string, SlicerStyleDef>()
  let retainedTableStyleBytes = 0
  const entries = [
    ...[...parsedTableStyleDefs].map(([name, def]) => ({ kind: 'table' as const, name, def, preferred: preferredTableStyleNames.has(name) })),
    ...[...parsedExtras.slicerStyles].map(([name, def]) => ({ kind: 'slicer' as const, name, def, preferred: preferredSlicerStyleNames.has(name) })),
  ].sort((a, b) => Number(b.preferred) - Number(a.preferred))
  for (const { kind, name, def } of entries) {
    const bytes = estimatePayloadBytes(def) + 32 + name.length * 2
    if (retainedTableStyleBytes + bytes > MAX_PAGED_TABLE_STYLE_DEFS_BYTES) continue
    retainedTableStyleBytes += bytes
    if (kind === 'table') tableStyleDefs.set(name, def as TableStyleDef)
    else slicerStyleDefs.set(name, def as SlicerStyleDef)
  }
  const tableStyleExtras: StylesExtras = { ...parsedExtras, tableStyles: tableStyleDefs, slicerStyles: slicerStyleDefs }
  return {
    formatting,
    colorContext: styles.colorContext,
    tableStyleDefs,
    tableStyleExtras,
  }
}

export async function workbookFormatting(workbookRoot: XmlNode, stylesText: string | undefined, themeText?: string): Promise<WorkbookFormatting> {
  return (await workbookFormattingWithColorContext(workbookRoot, stylesText, themeText)).formatting
}
