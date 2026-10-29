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
import { parseStylesExtras } from '../tableStyles/parseStylesExtras'
import type { TableStyleDef } from '../tableStyles/model'

export const DEFAULT_FORMATTING: WorkbookFormatting = Object.freeze({ dateSystem: '1900', styles: Object.freeze([]) })

/** Same resolution as parse.ts (cellXfs → font/fill/border/numFmt). */
export async function workbookFormattingWithColorContext(workbookRoot: XmlNode, stylesText: string | undefined, themeText?: string): Promise<{ formatting: WorkbookFormatting; colorContext: SheetColorContext; tableStyleDefs: ReadonlyMap<string, TableStyleDef> }> {
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
  const tableStyleDefs = parseStylesExtras(root, styles.colorContext).tableStyles
  const digitWidth = workbookDigitWidth(styles)
  return {
    formatting: { dateSystem, styles: styles.xfs.map(xf => resolveXfStyle(styles, xf)), ...(digitWidth !== 7 ? { maxDigitWidthPx: digitWidth } : {}) },
    colorContext: styles.colorContext,
    tableStyleDefs,
  }
}

export async function workbookFormatting(workbookRoot: XmlNode, stylesText: string | undefined, themeText?: string): Promise<WorkbookFormatting> {
  return (await workbookFormattingWithColorContext(workbookRoot, stylesText, themeText)).formatting
}
