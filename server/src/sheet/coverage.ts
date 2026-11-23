/**
 * Stage 4a (S4 §6.2 steps 4–5, R8): base coverage — which cells of the author view (identity order, authored sizes,
 * no session filter) lie under a drawing. A cell is covered iff its grid rect, or its R8 spill rect, meets any
 * drawing's transformed, effect-padded painted extent. Fill and transparency are ignored (owner, §11.1): the whole box
 * covers. Anything the analysis cannot bound fails closed to `unknown`, never to "absent" (F16).
 *
 * Task 15 fix round 1 (ruling "measure the spill"): a spill rect reaches only as far as the cell's text measurably
 * overflows its cell, laid out by the tile renderer's own cell layout (`createCellLayoutHost`, the same fonts and
 * measurement on the installed canvas backend), on the sides `spillSides` allows, plus `SPILL_SLACK`, capped by the
 * empty-run room. Text that fits never spills. Where the paint cannot be measured here (no canvas backend, rotated or
 * stacked text, a layout that throws, or a sheet whose conditional formats or region styles can restyle or inset the
 * text at paint time), the rect keeps the whole empty run (fail closed).
 */
import type { XlsxCell, XlsxCellStyle, XlsxDrawing, XlsxSheet } from '../../../src/xlsx/types'
import type { DrawingContentShape } from '../../../src/drawing/content'
import type { DrawingLine, DrawingStyle } from '../../../src/drawing/style'
import { resolveGeometry } from '../../../src/drawing/geometry'
import { cellSourceRow, drawingExtent } from '../../../src/xlsx/view/sheetCells'
import { cellTextOverflow, type SheetGeometryResult } from '../../../src/xlsx/view/viewport'
import { alignmentOverflows } from '../../../src/xlsx/view/textOverflow'
import { isValidCoordinate, MAX_COLS } from '../../../src/xlsx/view/coordinates'
import { installedCanvasBackend } from '../../../src/core/canvas-backend'
import { createCellLayoutHost } from '../../../src/xlsx/view/renderViewport'
import { sheetHasRegionStyles } from '../../../src/xlsx/pivots/stub'

export type BaseCoverage =
  | { kind: 'known'; covered(row: number, col: number): boolean; count: number }
  | { kind: 'unknown'; reason: 'unresolved-anchors' | 'unbounded-effect' | 'transform' }

interface Box { left: number; top: number; right: number; bottom: number }
type Node = XlsxDrawing | DrawingContentShape

/** The issue `resolveShadowDefinition` records for `ShadowDefinition.scaledOrSkewed` (src/drawing/style.ts): the resolved
 * `DrawingShadow` keeps no trace of it, so this issue on the node's style (or on a group's, whose shadow the children
 * inherit) is how the parse-time flag reaches a drawing node. */
const SCALED_SHADOW = 'outerShdw-scale'
/** Canvas antialiasing past a stroke's geometric edge, in CSS px. */
const AA_PX = 1

class Unknown extends Error { constructor(readonly reason: 'unbounded-effect' | 'transform') { super(reason) } }

/** Every condition under which `drawingExtent` silently skips a node or a point (sheetCells.ts), as a failure. */
function checkTransform(node: Node): void {
  const x = node.xEmu / 9525, y = node.yEmu / 9525, w = node.widthEmu / 9525, h = node.heightEmu / 9525
  const deg = node.rotationDeg ?? 0
  if ('transformValid' in node && node.transformValid === false) throw new Unknown('transform')
  if (![x, y, w, h, deg].every(Number.isFinite) || w < 0 || h < 0) throw new Unknown('transform')
  if (node.group) {
    const g = node.group, sx = g.ext.width / g.chExt.width, sy = g.ext.height / g.chExt.height
    if (![sx, sy, g.chOff.x, g.chOff.y].every(Number.isFinite) || sx < 0 || sy < 0) throw new Unknown('transform')
    for (const child of (node.children ?? []) as Node[]) checkTransform(child)
    return
  }
  if (node.drawingGeometry && !node.drawingGeometry.preset) {
    for (const path of resolveGeometry(node.drawingGeometry, w, h).paths) {
      for (const command of path.commands) for (let i = 1; i < command.length; i++) {
        const v = command[i]
        if (typeof v === 'number' && !Number.isFinite(v)) throw new Unknown('transform')
      }
    }
  }
  if ('content' in node && node.content?.kind === 'diagram') for (const shape of node.content.shapes as Node[]) checkTransform(shape)
}

/** How far a stroke paints past the path (the painter's own rules, src/drawing/paint.ts: round joins by default, the
 * miter limit for miter joins, arrowheads), plus antialiasing. */
function strokePad(line: DrawingLine | undefined): number {
  if (!line || line.fill?.kind === 'none') return 0
  const w = Number.isFinite(line.width) ? Math.max(0, line.width ?? 1) : 1
  let pad = w / 2
  if (line.join === 'miter') pad = Math.max(pad, w * (Number.isFinite(line.miterLimit) && (line.miterLimit ?? 0) > 0 ? line.miterLimit! : 8) / 2)
  for (const end of [line.headEnd, line.tailEnd]) {
    if (!end?.type || end.type === 'none') continue
    const base = Math.max(10 / 3, w), scale = { sm: 1.5, med: 2.25, lg: 3.375 }
    pad = Math.max(pad, Math.hypot(base * scale[end.length ?? 'med'], base * scale[end.width ?? 'med'] / 2) + w / 2)
  }
  return pad + AA_PX
}

/** One node's own paint past its box: its stroke, then its shadow (a copy of the stroked silhouette, offset in any
 * direction once rotated with the shape, blurred: canvas σ = blur/2, so 3σ = 1.5 × blur). Glow, reflection, 3D and the
 * other `EffectStyleDefinition.names` are never painted, so they add nothing. */
function nodePad(node: Node): number {
  const style = node.drawingStyle as DrawingStyle | undefined
  if (style?.issues?.some((i) => i.feature === SCALED_SHADOW)) throw new Unknown('unbounded-effect')
  let stroke = strokePad(style?.line)
  const legacy = (node as { line?: { widthEmu?: number } }).line
  if (legacy) stroke = Math.max(stroke, Math.max(1, (legacy.widthEmu ?? 12700) / 9525) / 2 + AA_PX)
  const shadow = style?.shadow
  if (!shadow) return stroke
  const reach = Math.hypot(shadow.offsetX, shadow.offsetY) + shadow.blurPx * 1.5
  if (!Number.isFinite(reach)) throw new Unknown('unbounded-effect')
  return stroke + reach
}

/** The largest own paint of any node of a tree, in page px: a group's scale scales its children's strokes and shadows. */
function treePad(node: Node, scale: number): number {
  let pad = nodePad(node) * scale
  if (node.group) {
    const g = node.group, s = scale * Math.max(1, g.ext.width / g.chExt.width, g.ext.height / g.chExt.height)
    for (const child of (node.children ?? []) as Node[]) pad = Math.max(pad, treePad(child, s))
  }
  if ('content' in node && node.content?.kind === 'diagram') for (const shape of node.content.shapes as Node[]) pad = Math.max(pad, treePad(shape, scale))
  return pad
}

const hasLeaf = (node: Node): boolean => !node.group || ((node.children ?? []) as Node[]).some(hasLeaf)

/** The padded painted extent of every top-level drawing, or the reason none can be given. */
function paintedExtents(drawings: readonly XlsxDrawing[]): Box[] {
  for (const node of drawings) checkTransform(node)
  const out: Box[] = []
  for (const node of drawings) {
    const e = drawingExtent([node])
    if (![e.left, e.top, e.right, e.bottom].every(Number.isFinite)) {
      if (hasLeaf(node)) throw new Unknown('transform')
      continue // an empty group paints nothing
    }
    const pad = treePad(node, 1)
    out.push({ left: e.left - pad, top: e.top - pad, right: e.right + pad, bottom: e.bottom + pad })
  }
  return out
}

/** Open-interval overlap: touching edges do not meet. */
const meets = (a: Box, b: Box) => a.left < b.right && a.right > b.left && a.top < b.bottom && a.bottom > b.top

const isNumber = (cell: XlsxCell) => typeof cell.value === 'number' && !cell.valueIsError
const holdsContent = (cell: XlsxCell) => (cell.value !== null && cell.value !== undefined) || cell.formula !== undefined

/** R8's conservative spill sides: general/left text spills right, right-aligned left, centred both ways; wrap, fill,
 * justify, distributed, shrink-to-fit and numbers never spill. A boolean or error under general alignment (centred by
 * Excel) and rotated text take both sides (fail closed). */
function spillSides(cell: XlsxCell, style: XlsxCellStyle | undefined): { left: boolean; right: boolean } {
  const none = { left: false, right: false }
  if (cell.value === null || cell.value === undefined || cell.value === '' || isNumber(cell)) return none
  if (style?.wrapText || style?.shrinkToFit || !alignmentOverflows(style?.horizontal)) return none
  const h = style?.horizontal
  if (style?.textRotation !== undefined || h === 'center') return { left: true, right: true }
  if (h === 'right') return { left: true, right: false }
  if (h === 'left' || typeof cell.value === 'string') return { left: false, right: true }
  return { left: true, right: true }
}

/**
 * Slack added to a measured overflow, in CSS px: `AA_PX` for antialiasing past the advance box, plus a quarter em of
 * the cell's font for ink past the advance (italic overhang, side bearings) and for hinting at another zoom or dpr,
 * where the scaled paint's advances need not equal the logical measurement. A quarter em (≈3.7 px at 11 pt) is well
 * above either (Calibri's widest overhang is under 0.1 em); it only widens a box that already overflows.
 */
export function spillSlack(fontSizePt: number | undefined): number {
  return AA_PX + 0.25 * (fontSizePt ?? 10) * 96 / 72
}
/** How far a cell's painted text reaches past its own box on each side (CSS px), or undefined when it cannot be known. */
export type SpillMeasure = (cell: XlsxCell, width: number, height: number, room: { left: number; right: number }) => { left: number; right: number } | undefined

/** The tile renderer's layout of one cell, on the installed canvas backend; undefined when nothing can be measured. */
function rendererSpillMeasure(sheet: XlsxSheet): SpillMeasure | undefined {
  if (sheetHasRegionStyles(sheet)) return undefined
  const cf = sheet.conditionalFormatting
  if (cf && (cf.rules.length > 0 || cf.unresolvedSheet)) return undefined
  const backend = installedCanvasBackend()
  if (!backend) return undefined
  let host: ReturnType<typeof createCellLayoutHost>
  try {
    const ctx = backend.createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
    // Format diagnostics land on a copy, never on the projected sheet.
    host = createCellLayoutHost(ctx, { ...sheet })
  } catch { return undefined }
  return (cell, width, height, room) => {
    let body: ReturnType<typeof host.layoutBody>
    try { body = host.layoutBody(cell, width, height, room) } catch { return undefined }
    const t = body.text
    if (!t) return { left: 0, right: 0 }
    if (t.kind === 'plain') return { left: Math.max(0, -t.tx), right: Math.max(0, t.tx + t.width - width) }
    if (t.rotation || t.stacked) return undefined
    // A laid-out line's clip is the band it paints in: its own box, widened by the room the measured text needs.
    return { left: Math.max(0, -t.clip.x), right: Math.max(0, t.clip.x + t.clip.width - width) }
  }
}

export function baseCoverage(sheet: XlsxSheet, geometry: SheetGeometryResult): BaseCoverage {
  const drawings = sheet.drawings ?? []
  if ((sheet.drawingAnchors?.unresolved ?? 0) > 0 || (drawings.length > 0 && !sheet.drawingAnchors)) return { kind: 'unknown', reason: 'unresolved-anchors' }
  if (geometry.kind !== 'ready') return { kind: 'unknown', reason: 'transform' }
  let extents: Box[]
  try { extents = paintedExtents(drawings) } catch (e) { if (e instanceof Unknown) return { kind: 'unknown', reason: e.reason }; throw e }
  const metrics = geometry.metrics
  const gridBox = (row: number, col: number): Box | undefined => {
    if (!isValidCoordinate(row, col)) return undefined
    const b = metrics.sourceCellToLogical({ row, col })
    return b ? { left: b.x, top: b.y, right: b.x + b.width, bottom: b.y + b.height } : undefined
  }
  const anyMeets = (b: Box) => extents.some((e) => meets(b, e))
  const covered = new Set<number>()
  let count = 0
  let measure: SpillMeasure | undefined | null = null
  if (extents.length > 0) {
    const union = extents.reduce((u, e) => ({ left: Math.min(u.left, e.left), top: Math.min(u.top, e.top), right: Math.max(u.right, e.right), bottom: Math.max(u.bottom, e.bottom) }))
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        const r = cellSourceRow(row, cell), c = cell.col
        let box = gridBox(r, c)
        if (!box || box.top >= union.bottom || box.bottom <= union.top) continue
        const merge = geometry.merges.first({ r0: r, r1: r, c0: c, c1: c })
        if (merge) {
          const a = gridBox(merge.minRow, merge.minCol), z = gridBox(merge.maxRow, merge.maxCol)
          if (a && z) box = { left: Math.min(a.left, box.left), top: Math.min(a.top, box.top), right: Math.max(z.right, box.right), bottom: Math.max(z.bottom, box.bottom) }
        } else {
          const sides = spillSides(cell, cell.style)
          if (sides.left || sides.right) {
            const room = cellTextOverflow(geometry, { row: r, col: c })
            if (room && (room.left > 0 || room.right > 0)) {
              if (measure === null) measure = rendererSpillMeasure(sheet)
              const rotated = cell.style?.textRotation !== undefined
              const used = rotated || !measure ? undefined : measure(cell, box.right - box.left, box.bottom - box.top, room)
              const reach = (side: 'left' | 'right') => {
                if (!sides[side]) return 0
                if (!used) return room[side]
                return used[side] > 0 ? Math.min(room[side], used[side] + spillSlack(cell.style?.fontSizePt)) : 0
              }
              box = { ...box, left: box.left - reach('left'), right: box.right + reach('right') }
            }
          }
        }
        if (!meets(box, union) || !anyMeets(box)) continue
        const key = r * MAX_COLS + c
        if (covered.has(key)) continue
        covered.add(key)
        if (holdsContent(cell)) count++
      }
    }
  }
  return {
    kind: 'known',
    count,
    covered(row: number, col: number): boolean {
      if (!isValidCoordinate(row, col)) return false
      if (covered.has(row * MAX_COLS + col)) return true
      const box = extents.length ? gridBox(row, col) : undefined
      return box !== undefined && anyMeets(box)
    },
  }
}
