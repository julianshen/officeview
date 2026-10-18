/**
 * Leaf helpers shared by the static renderer and the viewport planner (S2.3/S5.1).
 * Kept free of planner/axis imports so the eager static path (`render.ts`) does
 * not pull the interactive engine into the main bundle (plan S5.2).
 */
import type { XlsxCell, XlsxDrawing, XlsxRow } from '../types'
import type { DrawingContentShape } from '../../drawing/content'
import { resolveGeometry } from '../../drawing/geometry'

const EXPLICIT_REF = /^[A-Z]+\d+$/
/** Source row of a cell. An explicit cell reference wins over its row record:
 * `<row>` without `r` parses at index 0, but `<c r="A16386">` still names row
 * 16,385. Cells without a reference inherit the row record's index. */
export function cellSourceRow(row: XlsxRow, cell: XlsxCell): number {
  return EXPLICIT_REF.test(cell.ref) ? cell.row : row.index
}

type Matrix = [number, number, number, number, number, number]
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0]
function multiply(a: Matrix, b: Matrix): Matrix {
  return [a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1], a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3], a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5]]
}
/** Painted extent of drawing nodes in worksheet CSS px, including rotation, flips,
 * group transforms, custom geometry and diagram children. Empty input yields
 * infinite bounds (left/top +Infinity, right/bottom -Infinity). */
export function drawingExtent(nodes: readonly XlsxDrawing[]): { left: number; top: number; right: number; bottom: number } {
  const bounds = { left: Infinity, top: Infinity, right: -Infinity, bottom: -Infinity }
  type BoundsNode = XlsxDrawing | DrawingContentShape
  const walk = (items: readonly BoundsNode[], parent: Matrix): void => {
    for (const node of items) {
      const x = node.xEmu / 9525, y = node.yEmu / 9525, w = node.widthEmu / 9525, h = node.heightEmu / 9525
      const deg = node.rotationDeg ?? 0
      if ('transformValid' in node && node.transformValid === false || ![x, y, w, h, deg].every(Number.isFinite) || w < 0 || h < 0) continue
      const rad = deg * Math.PI / 180, cos = Math.cos(rad), sin = Math.sin(rad)
      const fx = node.flipH ? -1 : 1, fy = node.flipV ? -1 : 1
      const cx = x + w / 2, cy = y + h / 2
      const local: Matrix = [cos * fx, sin * fx, -sin * fy, cos * fy, cx - cos * fx * w / 2 + sin * fy * h / 2, cy - sin * fx * w / 2 - cos * fy * h / 2]
      const matrix = multiply(parent, local)
      const include = (px: number, py: number) => {
        const tx = matrix[0] * px + matrix[2] * py + matrix[4]
        const ty = matrix[1] * px + matrix[3] * py + matrix[5]
        if (Number.isFinite(tx) && Number.isFinite(ty)) {
          bounds.left = Math.min(bounds.left, tx); bounds.top = Math.min(bounds.top, ty)
          bounds.right = Math.max(bounds.right, tx); bounds.bottom = Math.max(bounds.bottom, ty)
        }
      }
      if (node.group) {
        const g = node.group, sx = g.ext.width / g.chExt.width, sy = g.ext.height / g.chExt.height
        if (![sx, sy, g.chOff.x, g.chOff.y].every(Number.isFinite) || sx < 0 || sy < 0) continue
        const childMatrix = multiply(matrix, [sx, 0, 0, sy, -sx * g.chOff.x / 9525, -sy * g.chOff.y / 9525])
        walk((node.children ?? []) as BoundsNode[], childMatrix)
      } else {
        include(0, 0); include(w, 0); include(w, h); include(0, h)
        if (node.drawingGeometry && !node.drawingGeometry.preset) {
          for (const path of resolveGeometry(node.drawingGeometry, w, h).paths) {
            for (const command of path.commands) {
              if (command[0] === 'moveTo' || command[0] === 'lnTo') include(command[1], command[2])
              else if (command[0] === 'quadBezTo') { include(command[1], command[2]); include(command[3], command[4]) }
              else if (command[0] === 'cubicBezTo') { include(command[1], command[2]); include(command[3], command[4]); include(command[5], command[6]) }
              else if (command[0] === 'arcTo') {
                const [, cx, cy, rx, ry] = command
                include(cx - rx, cy - ry); include(cx + rx, cy - ry)
                include(cx - rx, cy + ry); include(cx + rx, cy + ry)
              }
            }
          }
        }
        if ('content' in node && node.content?.kind === 'diagram') walk(node.content.shapes as BoundsNode[], matrix)
      }
    }
  }
  walk(nodes, IDENTITY)
  return bounds
}
