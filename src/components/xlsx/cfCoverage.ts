import type { LogicalCellBox, SheetMetrics } from '../../xlsx/view/contracts'
import type { FormatBudget } from '../../xlsx/view/formatWorkspace'
import type { SourceRect } from '../../xlsx/view/coordinates'
import { CF_PAINT_BATCH_CELLS } from '../../xlsx/cfLoader'

/** Full tile/ring coverage, never the CSS viewport's partial edge cells.
 * Metadata is admitted before retention; overview refuses before row enumeration. */
export function collectCfCoverage(metrics: SheetMetrics, boxes: readonly LogicalCellBox[], budget: FormatBudget) {
  let bytes = 2048, disposed = false
  budget.reserve(bytes)
  try {
    const rects: SourceRect[] = []
    let cells = 0
    for (const box of boxes) {
      const rows = metrics.rows.visibleBetween(box.y, box.y + box.height)
      const cols = metrics.cols.visibleBetween(box.x, box.x + box.width)
      let nr=0,nc=0
      for (const r of rows.runs) nr+=r.end-r.start+1
      for (const c of cols.runs) nc+=c.end-c.start+1
      cells+=nr*nc
      if (cells>CF_PAINT_BATCH_CELLS) throw new Error('Conditional formatting tile coverage exceeds the admitted batch; base formatting is shown.')
      for (const run of rows.runs) for(let slot=run.start;slot<=run.end;slot++) {
        const row=metrics.order.displayToSource(slot)
        for(const col of cols.runs) {
          budget.reserve(128);bytes+=128
          rects.push({r0:row,r1:row,c0:col.start,c1:col.end})
        }
      }
    }
    budget.reserve(boxes.length*64);bytes+=boxes.length*64
    const owned=boxes.map(b=>({...b}))
    return {rects, contains:(box:LogicalCellBox)=>!disposed && owned.some(b=>box.x>=b.x-1e-7&&box.y>=b.y-1e-7&&box.x+box.width<=b.x+b.width+1e-7&&box.y+box.height<=b.y+b.height+1e-7),
      dispose(){if(!disposed){disposed=true;rects.length=0;owned.length=0;budget.release(bytes)}}}
  } catch(error) {budget.release(bytes);throw error}
}
