/**
 * Paged chrome geometry (spec §6.10). The only eager module under paged/:
 * <OfficeDoc>'s placeholders and the lazy toolbar and status line both size
 * themselves from these, so the swap cannot shift layout.
 */
export const PAGED_TOOLBAR_HEIGHT = { regular: 40, compact: 48 } as const
export const PAGED_STATUS_HEIGHT = 24
export const COMPACT_PANE_PX = 600

export function isCompact(containerWidth: number, coarse: boolean): boolean {
  return coarse || containerWidth < COMPACT_PANE_PX
}

/** The scroll container's padding around the page stack (PAGE_PADDING_PX). */
export const PAGE_PADDING_PX = 16
