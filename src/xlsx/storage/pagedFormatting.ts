/**
 * Workbook formatting metadata of a paged import (plan S5.4; spec §§4.2, 8.2,
 * 11.1): the resolved cell-style table (one entry per cellXfs index, bounded by
 * the styles part's small-part cap) and the date system. Paged pages keep only
 * a style index per cell; this table lets paint, copy, find and filter format
 * paged cells exactly as the object parser does for resident cells.
 */
import type { OfficePackage } from '../../core/zip'
import { attrs, getChildren, parseXml, type XmlNode } from '../../core/xml'
import { parseStyles, resolveXfStyle } from '../parse'
import type { WorkbookFormatting } from '../view/contracts'

export const DEFAULT_FORMATTING: WorkbookFormatting = Object.freeze({ dateSystem: '1900', styles: Object.freeze([]) })

/** Same resolution as parse.ts (cellXfs → font/fill/border/numFmt). */
export async function workbookFormatting(workbookRoot: XmlNode, stylesText: string | undefined): Promise<WorkbookFormatting> {
  const pr = getChildren(workbookRoot, 'workbookPr')[0]
  const date1904 = pr ? attrs(pr).date1904 : undefined
  const dateSystem = date1904 === '1' || date1904 === 'true' ? '1904' : '1900'
  if (stylesText === undefined) return { dateSystem, styles: [] }
  const root = parseXml(stylesText)
  const styles = await parseStyles({ xml: async (path: string) => path === 'xl/styles.xml' ? root : undefined } as unknown as OfficePackage)
  return {
    dateSystem,
    styles: styles.xfs.map(xf => resolveXfStyle(styles, xf)),
  }
}
