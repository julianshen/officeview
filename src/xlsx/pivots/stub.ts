import { attrs, getChildren, elementNamespaceUri, type XmlNode } from '../../core/xml'
import type { XlsxSheet } from '../types'
import { MAX_COLS, type SourceRect } from '../view/coordinates'

export const X14_SLICER_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
export const X15_TIMELINE_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'
export const SLICER_EXTENSION_URI = '{A8765BA9-456A-4dab-B4F3-ACF838C121DE}'
export const TIMELINE_EXTENSION_URI = '{7E03D99C-DC04-49d9-9315-930204A7B6E9}'
/** The prefix of a pivot table part that the location scan reads (resident parse and paged import). The paged
 * worker's cap on whole pivot, cache, slicer and timeline parts is the profile's `paged.pivotPartBytes` (Ruling AP). */
export const PIVOT_PART_MAX_BYTES = 64 * 1024

/** `rowPageCount`/`colPageCount` are the `<location>` page-field grid above `ref` (absent when 0). */
export interface PivotRegionRef { partPath: string; ref: string; name?: string; styleName?: string; rowPageCount?: number; colPageCount?: number }
/** Page-field grid counts above a few thousand cannot be real (a field is at most one column of the cache). */
const MAX_PAGE_GRID = 16_384

/** A saved pivot's guarded rectangles (S3-G): its location and the page-field cells above it. The page
 * grid is `rowPageCount` rows ending one gap row above the location, with `colPageCount` label/value
 * pairs three columns apart (Excel's placement, as savedLayout and the engine place them). */
export function savedPivotRects(table: SourceRect, rowPageCount = 0, colPageCount = 0): SourceRect[] {
  const rects = [table]
  const r0 = Math.max(0, table.r0 - 1 - rowPageCount), r1 = table.r0 - 2
  if (rowPageCount <= 0 || r1 < r0) return rects
  for (let j = 0; j < Math.max(1, colPageCount); j++) {
    const c0 = table.c0 + 3 * j
    if (c0 + 1 >= MAX_COLS) break
    rects.push({ r0, r1, c0, c1: c0 + 1 })
  }
  return rects
}
export interface ControlPartRef { partPath: string; anchorId?: string }
export interface TableStyleInfo {
  name?: string
  showFirstColumn: boolean
  showLastColumn: boolean
  showRowStripes: boolean
  showColumnStripes: boolean
}

export function sheetHasRegionStyles(sheet: XlsxSheet): boolean {
  return Boolean(sheet.regionStyles || sheet.tables?.some(table => Boolean(table.styleInfo?.name)) || sheet.pivotRegions?.length || sheet.pivotTables?.length || sheet.slicerRefs?.length || sheet.timelineRefs?.length)
}

/** An A1 cell's 1-based row and column, within the sheet's bounds. */
export function cellCoords(cell: string): { row: number; col: number } | undefined {
  const m = /^\$?([A-Z]{1,3})\$?([1-9][0-9]*)$/.exec(cell)
  if (!m) return undefined
  const col = [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0)
  const row = Number(m[2])
  return col <= 16384 && row <= 1048576 ? { row, col } : undefined
}

function validRef(ref: string): boolean {
  const parts = ref.split(':')
  if (parts.length < 1 || parts.length > 2) return false
  const start = cellCoords(parts[0]), end = cellCoords(parts[1] ?? parts[0])
  return Boolean(start && end && start.row <= end.row && start.col <= end.col)
}

function attr(raw: string, name: string): string | undefined {
  const m = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(["'])(.*?)\\1`, 's').exec(raw)
  return m?.[2]
}

function decodeXmlEntities(value: string): string {
  return value.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[\da-f]+);/gi, (entity) => {
    const named: Record<string, string> = { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" }
    const known = named[entity.toLowerCase()]
    if (known !== undefined) return known
    const hex = /^&#x([\da-f]+);$/i.exec(entity)
    const decimal = /^&#(\d+);$/.exec(entity)
    const codePoint = hex ? Number.parseInt(hex[1], 16) : decimal ? Number.parseInt(decimal[1], 10) : NaN
    return Number.isInteger(codePoint) && codePoint >= 0 && codePoint <= 0x10ffff && !(codePoint >= 0xd800 && codePoint <= 0xdfff)
      ? String.fromCodePoint(codePoint) : entity
  })
}

/** The prefix of a comments (or threaded comments) part a pivot sheet's overlap check reads (Ruling S3-AC). */
export const NOTE_PART_MAX_BYTES = 4 * 1024 * 1024
/** Noted cells kept per sheet; past this the sheet's notes are unknown (every cell then counts as occupied). */
export const MAX_NOTE_CELLS = 65_536
/** Bounded scan of a comments or threaded comments part: the cells its comments are on (row * MAX_COLS + col, 0-based);
 * undefined when the text is absent, cut short (no closing root tag) or names a cell it cannot place. */
export function noteCells(text: string | undefined): number[] | undefined {
  if (text === undefined || !/<\/(?:[A-Za-z_][\w.-]*:)?(?:comments|ThreadedComments)\s*>\s*$/.test(text)) return undefined
  const out: number[] = []
  const comment = /<(?:[A-Za-z_][\w.-]*:)?(?:comment|threadedComment)\b([^>]*)>/g
  for (let m = comment.exec(text); m; m = comment.exec(text)) {
    const at = cellCoords(attr(m[1]!, 'ref') ?? '')
    if (!at || out.length >= MAX_NOTE_CELLS) return undefined
    out.push((at.row - 1) * MAX_COLS + at.col - 1)
  }
  return out
}
/** A worksheet relationship to a comments or threaded comments part. */
export const isNoteRelationship = (type: string) => type.endsWith('/comments') || type.endsWith('/threadedComment')
/** Adds one part's noted cells to a sheet's (undefined: the part was unreadable, and the sheet's notes are unknown). */
export function recordNoteCells(sheet: { noteCells?: number[]; notesUnknown?: boolean }, cells: number[] | undefined): void {
  if (cells && (sheet.noteCells?.length ?? 0) + cells.length <= MAX_NOTE_CELLS) (sheet.noteCells ??= []).push(...cells)
  else { sheet.notesUnknown = true; sheet.noteCells = undefined }
}

/** Bounded scan: never constructs an XML tree for a pivot definition. */
export function parsePivotRegion(text: string | undefined, partPath: string): { region?: PivotRegionRef; unreadable: boolean } {
  if (text === undefined || text.length > PIVOT_PART_MAX_BYTES) return { unreadable: true }
  const root = /<(?:[A-Za-z_][\w.-]*:)?pivotTableDefinition\b([^>]*)>/s.exec(text)
  const loc = /<(?:[A-Za-z_][\w.-]*:)?location\b([^>]*)\/?\s*>/s.exec(text)
  if (!root || !loc) return { unreadable: true }
  const ref = attr(loc[1], 'ref')
  if (!ref || !validRef(ref)) return { unreadable: true }
  const rawName = attr(root[1], 'name')
  const name = rawName === undefined ? undefined : decodeXmlEntities(rawName)
  const style = /<(?:[A-Za-z_][\w.-]*:)?pivotTableStyleInfo\b([^>]*)>/s.exec(text)
  const rawStyleName = style ? attr(style[1], 'name') : undefined
  const styleName = rawStyleName === undefined ? undefined : decodeXmlEntities(rawStyleName)
  const count = (key: string) => { const raw = attr(loc[1], key), n = raw === undefined || !/^\d+$/.test(raw) ? 0 : Number(raw); return n > 0 && n <= MAX_PAGE_GRID ? n : undefined }
  const rowPageCount = count('rowPageCount'), colPageCount = count('colPageCount')
  return { region: { partPath, ref, ...(name ? { name } : {}), ...(styleName ? { styleName } : {}), ...(rowPageCount ? { rowPageCount } : {}), ...(colPageCount ? { colPageCount } : {}) }, unreadable: false }
}

/** Distinct slicer or timeline relationship ids read per worksheet (each names a part of up to 256 controls). */
export const MAX_CONTROL_IDS = 1_024

/** The worksheet's slicer and timeline relationship ids, de-duplicated and capped (a repeated id would read its part again). */
export function controlIds(root: XmlNode | undefined): { slicerIds: string[]; timelineIds: string[] } {
  const slicers = new Set<string>(), timelines = new Set<string>()
  const add = (set: Set<string>, id: unknown) => { if (typeof id === 'string' && id !== '' && set.size < MAX_CONTROL_IDS) set.add(id) }
  const result = () => ({ slicerIds: [...slicers], timelineIds: [...timelines] })
  const data = root && (getChildren(root, 'worksheet')[0] ?? root)
  if (!data) return result()
  const main = elementNamespaceUri(data)
  for (const extList of getChildren(data, 'extLst')) {
    if (elementNamespaceUri(extList) !== main) continue
    for (const ext of getChildren(extList, 'ext')) {
      if (elementNamespaceUri(ext) !== main) continue
      const uri = attrs(ext).uri
      for (const node of getChildren(ext, 'slicerList')) {
        if (uri !== SLICER_EXTENSION_URI || elementNamespaceUri(node) !== X14_SLICER_NS) continue
        for (const slicer of getChildren(node, 'slicer')) add(slicers, attrs(slicer).id)
      }
      for (const node of getChildren(ext, 'timelineRefs')) {
        if (uri !== TIMELINE_EXTENSION_URI || elementNamespaceUri(node) !== X15_TIMELINE_NS) continue
        for (const ref of getChildren(node, 'timelineRef')) add(timelines, attrs(ref).id)
      }
    }
  }
  return result()
}
