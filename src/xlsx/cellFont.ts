/**
 * The one cell font string: the static, interactive and worker painters and
 * auto-fit measurement all build it here, so what is measured is what is
 * painted (a bare "Calibri" resolves to a different fallback face on hosts
 * without it). See render.ts for why the stack is explicit.
 */
export const CELL_FONT_STACK = '"Calibri", Carlito, Arial, Helvetica, "Liberation Sans", sans-serif'

/** The font-family list for a cell face: an authored name goes ahead of the
 * stack (so a missing face still falls back deterministically); Calibri and an
 * absent name are the stack itself. */
export function cellFontFamily(name: string | undefined): string {
  const face = cellFace(name)
  return !face || face === 'Calibri' ? CELL_FONT_STACK : `"${face}", ${CELL_FONT_STACK}`
}

/** Whether a cell paints in Calibri, its own or the default (an absent name): the
 * only face whose Excel layout metrics the renderer carries (render.ts calibriWidth). */
export function isCalibriFace(name: string | undefined): boolean {
  const face = cellFace(name)
  return !face || face === 'Calibri'
}

const cellFace = (name: string | undefined) => name?.replace(/["\\;{}]/g, '').trim()

/** The canvas font string of a cell style at `sizePt`. */
export function cellFontString(style: { bold?: boolean; italic?: boolean; fontName?: string } | undefined, sizePt: number): string {
  return `${style?.italic ? 'italic ' : ''}${style?.bold ? 'bold ' : ''}${sizePt}pt ${cellFontFamily(style?.fontName)}`
}

/** Mixed narrow and wide glyphs, so two different faces measure differently. */
const FACE_PROBE = 'mmmmmmmmmmlli1WQ@#0.'
/** Per realm (main thread, each worker): a face's availability does not change per cell. */
const faceAvailability = new Map<string, boolean>()

/**
 * Whether an authored cell face draws as itself on this canvas backend, rather than
 * through the fallback stack: measured once per face (cached), by checking that the
 * face measures the same ahead of two different generics (a missing face takes each
 * generic's width). A backend whose generics measure alike gives no signal and reports
 * unavailable, so the caller keeps its fallback behaviour. A face that only becomes
 * available later in the session (a loading web font) keeps its first answer.
 */
export function cellFaceAvailable(ctx: CanvasRenderingContext2D, name: string | undefined): boolean {
  const face = cellFace(name)
  if (!face) return false
  let known = faceAvailability.get(face)
  if (known === undefined) faceAvailability.set(face, known = probeFace(ctx, face))
  return known
}

function probeFace(ctx: CanvasRenderingContext2D, face: string): boolean {
  if (typeof ctx.measureText !== 'function') return false
  const previous = ctx.font
  try {
    const width = (stack: string) => { ctx.font = `40px ${stack}`; return ctx.measureText(FACE_PROBE).width }
    const mono = width('monospace'), serif = width('serif')
    if (!(mono > 0) || !(serif > 0) || mono === serif) return false
    return width(`"${face}", monospace`) === width(`"${face}", serif`)
  } catch { return false } finally { ctx.font = previous }
}
