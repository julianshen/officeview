/** Bounded worksheet drawing anchor records (S3.3), shared by the object parser
 * (drawing.ts) and the paged importer so both adapters read anchors identically.
 * Only placement is read; content and images are never touched here. */
import type { XlsxAnchorMarker, XlsxDrawingAnchorRecord } from './types'

/** Child text of an xdr:from/xdr:to marker and the attributes of xdr:ext/xdr:pos,
 * first occurrence of each direct child only. */
export interface AnchorParts {
  from?: Record<string, string>
  to?: Record<string, string>
  ext?: Record<string, string>
  pos?: Record<string, string>
}
export const ANCHOR_ELEMENTS: ReadonlySet<string> = new Set(['oneCellAnchor', 'twoCellAnchor', 'absoluteAnchor'])

/** Local name of a drawing part's document element; only `wsDr` holds anchors. */
export const DRAWING_ROOT = 'wsDr'
/** Document element local name of an XML text (prolog, comments and doctype skipped). */
export function rootLocalName(xml: string): string | undefined {
  const m = /^(?:\uFEFF)?(?:\s|<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<!DOCTYPE[^>]*>)*<(?:[A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)/.exec(xml)
  return m?.[1]
}

/** Stable anchor identity: the drawing part plus the anchor's index among its top-level anchors. */
export const anchorId = (partPath: string, index: number): string => `${partPath}#${index}`

const MAX_ROWS = 1_048_576, MAX_COLS = 16_384
const finite = (raw: string | undefined): number | undefined =>
  raw !== undefined && raw.trim() !== '' && Number.isFinite(Number(raw)) ? Number(raw) : undefined
const extent = (value: number | undefined): value is number => value !== undefined && value >= 0

/** Markers address the full sheet grid (the static renderer's narrower caps are
 * a paint limit, not proof the anchor is absent). */
function marker(parts: Record<string, string> | undefined): XlsxAnchorMarker | undefined {
  const col = finite(parts?.col), row = finite(parts?.row), colOffEmu = finite(parts?.colOff), rowOffEmu = finite(parts?.rowOff)
  if (col === undefined || row === undefined || colOffEmu === undefined || rowOffEmu === undefined) return undefined
  if (!Number.isSafeInteger(col) || !Number.isSafeInteger(row) || col < 0 || row < 0 || col >= MAX_COLS || row >= MAX_ROWS) return undefined
  return { col, row, colOffEmu, rowOffEmu }
}

/** A record for a readable anchor, undefined when its placement cannot be read. */
export function anchorRecord(name: string, id: string, parts: AnchorParts): XlsxDrawingAnchorRecord | undefined {
  if (name === 'absoluteAnchor') {
    const x = finite(parts.pos?.x), y = finite(parts.pos?.y), cx = finite(parts.ext?.cx), cy = finite(parts.ext?.cy)
    return x !== undefined && y !== undefined && extent(cx) && extent(cy) ? { id, kind: 'absolute', x, y, cx, cy } : undefined
  }
  const from = marker(parts.from)
  if (!from) return undefined
  if (name === 'oneCellAnchor') {
    const cx = finite(parts.ext?.cx), cy = finite(parts.ext?.cy)
    return extent(cx) && extent(cy) ? { id, kind: 'oneCell', from, cx, cy } : undefined
  }
  if (name !== 'twoCellAnchor') return undefined
  const to = marker(parts.to)
  return to ? { id, kind: 'twoCell', from, to } : undefined
}
