/**
 * The viewer's frame host source (Rulings S3-A11Y-PAGED, S4-L): what InteractiveSheetDoc hands its forced-paged grid
 * semantics so a scope that never paints decides each bar and icon on the host the frame paints it on. Per cell: the
 * logical box from the projection metrics (as the viewport placements are laid out), the stored cell and its screen
 * base style; and the frame's own measurement of the cell's un-inset text width through the viewer's layout host, on
 * the cell carrying the CF-effective style the frame lays it out with. A cell not loaded, hidden or outside the metrics
 * has no host (withheld). Nothing is retained per cell. Lives in the viewer's lazy chunk; never imports the cf chunk.
 */
import type { XlsxCell, XlsxCellStyle } from '../types'
import type { SourceCell } from './coordinates'
import { cellTextWidth, type CellLayoutHost } from './textLayout'

export interface FrameHostInputs {
  readonly metrics: { sourceCellToLogical(at: SourceCell): { x: number; y: number; width: number; height: number } | null | undefined }
  cellAt(at: SourceCell): XlsxCell | undefined
  styleOf(cell: XlsxCell): XlsxCellStyle | undefined
  layoutHost(): CellLayoutHost | undefined
}

export function createFrameHostSource(read: () => FrameHostInputs | undefined): import('../cf/render').CfFrameHostSource {
  return {
    cell(row, col) {
      const inputs = read()
      if (!inputs) return undefined
      const at = { row, col }
      const cell = inputs.cellAt(at)
      if (!cell) return undefined
      let box: ReturnType<FrameHostInputs['metrics']['sourceCellToLogical']>
      try { box = inputs.metrics.sourceCellToLogical(at) } catch { return undefined }
      return box ? { box, base: inputs.styleOf(cell), cell } : undefined
    },
    textWidth(cell, box) {
      const host = read()?.layoutHost()
      if (!host) return undefined
      try { return cellTextWidth(host.layoutBody(cell, box.width, box.height)) } catch { return undefined }
    },
  }
}
