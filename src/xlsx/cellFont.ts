/**
 * The one cell font string: the static, interactive and worker painters and
 * auto-fit measurement all build it here, so what is measured is what is
 * painted (a bare "Calibri" resolves to a different fallback face on hosts
 * without it). See render.ts for why the stack is explicit.
 */
export const CELL_FONT_STACK = '"Calibri", Carlito, Arial, Helvetica, "Liberation Sans", sans-serif'

/** The canvas font string of a cell style at `sizePt`. */
export function cellFontString(style: { bold?: boolean; italic?: boolean } | undefined, sizePt: number): string {
  return `${style?.italic ? 'italic ' : ''}${style?.bold ? 'bold ' : ''}${sizePt}pt ${CELL_FONT_STACK}`
}
