import { attrs, getChildren, elementNamespaceUri, type XmlNode } from '../../core/xml'
import type { XlsxSheet } from '../types'

export const X14_SLICER_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
export const X15_TIMELINE_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'
export const SLICER_EXTENSION_URI = '{A8765BA9-456A-4dab-B4F3-ACF838C121DE}'
export const TIMELINE_EXTENSION_URI = '{7E03D99C-DC04-49d9-9315-930204A7B6E9}'
export const PIVOT_PART_MAX_BYTES = 64 * 1024

export interface PivotRegionRef { partPath: string; ref: string; name?: string }
export interface ControlPartRef { partPath: string; anchorId?: string }
export interface TableStyleInfo {
  name?: string
  showFirstColumn: boolean
  showLastColumn: boolean
  showRowStripes: boolean
  showColumnStripes: boolean
}

export function sheetHasRegionStyles(sheet: XlsxSheet): boolean {
  return Boolean(sheet.tables?.some(table => Boolean(table.styleInfo?.name)) || sheet.pivotRegions?.length || sheet.slicerRefs?.length || sheet.timelineRefs?.length)
}

function validRef(ref: string): boolean {
  const cell = '\\$?[A-Z]{1,3}\\$?[1-9][0-9]*'
  const range = new RegExp(`^${cell}(?::${cell})?$`)
  if (!range.test(ref)) return false
  const cols = ref.replace(/\$?[0-9]+/g, '').split(':')
  const rows = ref.match(/\$?([0-9]+)/g) ?? []
  const colNum = (s: string) => [...s.replace(/\$/g, '')].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0)
  return cols.every(s => colNum(s) <= 16384) && rows.every(s => Number(s.replace(/\$/g, '')) <= 1048576)
}

function attr(raw: string, name: string): string | undefined {
  const escaped = name
  const m = new RegExp(`(?:^|\\s)${escaped}\\s*=\\s*(["'])(.*?)\\1`, 's').exec(raw)
  return m?.[2]
}

/** Bounded scan: never constructs an XML tree for a pivot definition. */
export function parsePivotRegion(text: string | undefined, partPath: string): { region?: PivotRegionRef; unreadable: boolean } {
  if (text === undefined || text.length > PIVOT_PART_MAX_BYTES) return { unreadable: true }
  const root = /<(?:[A-Za-z_][\w.-]*:)?pivotTableDefinition\b([^>]*)>/s.exec(text)
  const loc = /<(?:[A-Za-z_][\w.-]*:)?location\b([^>]*)\/?\s*>/s.exec(text)
  if (!root || !loc) return { unreadable: true }
  const ref = attr(loc[1], 'ref')
  if (!ref || !validRef(ref)) return { unreadable: true }
  const name = attr(root[1], 'name')
  return { region: { partPath, ref, ...(name ? { name } : {}) }, unreadable: false }
}

export function controlIds(root: XmlNode | undefined): { slicerIds: string[]; timelineIds: string[] } {
  const slicerIds: string[] = [], timelineIds: string[] = []
  const data = root && (getChildren(root, 'worksheet')[0] ?? root)
  if (!data) return { slicerIds, timelineIds }
  const main = elementNamespaceUri(data)
  for (const extList of getChildren(data, 'extLst')) {
    if (elementNamespaceUri(extList) !== main) continue
    for (const ext of getChildren(extList, 'ext')) {
      if (elementNamespaceUri(ext) !== main) continue
      const uri = attrs(ext).uri
      for (const node of getChildren(ext, 'slicerList')) {
        if (uri !== SLICER_EXTENSION_URI || elementNamespaceUri(node) !== X14_SLICER_NS) continue
        for (const slicer of getChildren(node, 'slicer')) {
          const id = attrs(slicer).id
          if (typeof id === 'string' && id !== '') slicerIds.push(id)
        }
      }
      for (const node of getChildren(ext, 'timelineRefs')) {
        if (uri !== TIMELINE_EXTENSION_URI || elementNamespaceUri(node) !== X15_TIMELINE_NS) continue
        for (const ref of getChildren(node, 'timelineRef')) {
          const id = attrs(ref).id
          if (typeof id === 'string' && id !== '') timelineIds.push(id)
        }
      }
    }
  }
  return { slicerIds, timelineIds }
}
