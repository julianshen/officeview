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
