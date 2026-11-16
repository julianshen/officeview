/**
 * The visible mark (§7.6): its spec, its styling and the one way the server paints it. Moved out of
 * `render.ts` (which re-exports it) so the show-plan builder can burn it into layers without a cycle.
 */
import { paintWatermark, setWatermarkMeasurer, type WatermarkOptions } from '../../../src/core/watermark'

export const FIXED_MARK_TEXT = 'officeview · protected'
export type MarkSpec = { kind: 'fixed' } | { kind: 'viewer'; text: string } | { kind: 'off' }

/** Visible-mark styling [call]: mid-grey reads on light and dark pages alike. */
export function markOptions(mark: MarkSpec): WatermarkOptions | undefined {
  if (mark.kind === 'off') return undefined
  return { text: mark.kind === 'fixed' ? FIXED_MARK_TEXT : mark.text, color: '#808080', opacity: 0.18 }
}

/** Paints the visible mark from a clean watermark module state.
 *
 * `src/core/watermark.ts` keeps a process-global text-width cache keyed on the
 * ROUNDED point size, but (with no measurer registered, as here) stores a
 * width computed from the UNROUNDED pixel size. Two pages whose sizes round to
 * the same pt (a 960-px deck and a 955-px sheet are both 50 pt) then share one
 * entry, and whichever painted first sets the tile step of the other: the
 * canonical a long-lived worker serves would depend on what it rendered
 * before, and a fresh trace process would refuse the reference (F1, Stage 2b).
 * `setWatermarkMeasurer(null)` is the module's only reset (measurer stays
 * null, cache cleared), so every mark is laid out from its own page alone. */
export function paintMark(ctx: CanvasRenderingContext2D, page: { widthPx: number; heightPx: number }, mark: WatermarkOptions): void {
  setWatermarkMeasurer(null)
  paintWatermark(ctx, page, mark)
}
