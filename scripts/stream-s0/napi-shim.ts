/**
 * Stage-0 Task 1 — napi shim.
 *
 * The server-side backend is @napi-rs/canvas (§2 item 1), but every consumer in
 * this stage (probes, tests, pixel-diff style helpers) is written against the
 * node `canvas` API names. This shim re-exports the @napi-rs/canvas equivalents
 * under those node-canvas names so test code can swap the backend purely by
 * a vitest resolve.alias (`vitest.stream-s0-napi.config.ts` aliases `canvas`
 * to this file) — the `CanvasBackend` seam (§7.3) is Stage 1.
 */
import {
  Canvas as NapiCanvas,
  GlobalFonts,
  Image as NapiImage,
  createCanvas as napiCreateCanvas,
  loadImage as napiLoadImage,
} from '@napi-rs/canvas'

/** The @napi-rs/canvas canvas class, exposed under the node-canvas name. */
export { NapiCanvas as Canvas }

/** node-canvas `createCanvas(w, h)` → the @napi-rs/canvas factory. */
export function createCanvas(w: number, h: number): NapiCanvas {
  return napiCreateCanvas(w, h)
}

/** node-canvas `loadImage(src)` → the @napi-rs/canvas loader. */
export function loadImage(src: Parameters<typeof napiLoadImage>[0]): ReturnType<typeof napiLoadImage> {
  return napiLoadImage(src)
}

/** The @napi-rs/canvas Image class, exposed under the node-canvas name. */
export { NapiImage as Image }

/**
 * node-canvas-style `registerFont(path, { family, weight?, style? })`.
 *
 * → `GlobalFonts.registerFromPath`. The descriptor's `weight`/`style` are
 * accepted for node-canvas compatibility; @napi-rs/canvas resolves face
 * weight/style from the font file's own metadata under the family alias, so
 * the path is registered once per family (re-registration of the same path
 * updates the family in place).
 */
export function registerFont(
  path: string,
  options?: { family: string; weight?: string; style?: string },
): void {
  if (!options || typeof options.family !== 'string' || options.family.length === 0) {
    throw new TypeError('registerFont requires a font family name')
  }
  GlobalFonts.registerFromPath(path, options.family)
}
