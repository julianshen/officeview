/**
 * Stage-0 Task 1 — napi shim.
 *
 * The server-side backend is @napi-rs/canvas (§2 item 1), but every consumer in
 * this stage (probes, tests, pixel-diff style helpers) is written against the
 * node `canvas` API names. This shim re-exports the @napi-rs/canvas equivalents
 * under those node-canvas names so test code can swap the backend purely by
 * a vitest resolve.alias (`vitest.stream-s0-napi.config.ts` aliases `canvas`
 * to this file) — the `CanvasBackend` seam (§7.3) is Stage 1.
 *
 * It also self-identifies: `STREAM_S0_BACKEND = '@napi-rs/canvas'` is exported
 * so `results.resolveBackend()` can distinguish the module the `canvas`
 * specifier resolved to (the real `canvas` package never exports that name);
 * `scripts/stream-s0/napi-preload.ts` serves the same map under bun scripts.
 */
import {
  Canvas as NapiCanvas,
  GlobalFonts,
  Image as NapiImage,
  ImageData as NapiImageData,
  createCanvas as napiCreateCanvas,
  loadImage as napiLoadImage,
} from '@napi-rs/canvas'

/** Marker export: the module the `canvas` specifier resolved to is this shim. */
export const STREAM_S0_BACKEND = '@napi-rs/canvas'

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

/** node-canvas `ImageData`, exposed via the @napi-rs/canvas class. */
export { NapiImageData as ImageData }

/** node-canvas `createImageData(w, h)`. */
export function createImageData(w: number, h: number): NapiImageData {
  return new NapiImageData(w, h)
}

/**
 * family → the `FontKey` `GlobalFonts.registerFromPath` returned, so callers
 * can distinguish a family's registered faces for lease bookkeeping. The
 * registry in @napi-rs/canvas 1.1.0 has no unregister-by-family — `deregisterFont`
 * drops the map entry; the faces themselves stay process-global (documented).
 */
const fontKeys = new Map<string, unknown>()

/**
 * node-canvas-style `registerFont(path, { family, weight?, style? })`.
 *
 * → `GlobalFonts.registerFromPath`. The descriptor's `weight`/`style` are
 * accepted for node-canvas compatibility; @napi-rs/canvas resolves face
 * weight/style from the font file's own metadata under the family alias, so
 * the path is registered once per family. Throws on a failed registration
 * (`registerFromPath` returns null) instead of silently continuing.
 */
export function registerFont(
  path: string,
  options: { family: string; weight?: string; style?: string },
): void {
  if (!options || typeof options.family !== 'string' || options.family.length === 0) {
    throw new TypeError('registerFont requires a font family name')
  }
  const key = GlobalFonts.registerFromPath(path, options.family)
  if (key === null) {
    throw new Error(`registerFont: GlobalFonts.registerFromPath failed for ${options.family} (path: ${path})`)
  }
  fontKeys.set(options.family, key)
}

/** The @napi-rs/canvas global registry, for probe code that counts faces. */
export { GlobalFonts }

/**
 * Drop this family from the shim's face map. @napi-rs/canvas 1.1.0 has no
 * unregister-by-family; the registered faces remain process-global for the
 * backend's lifetime regardless (registered faces cannot be individually
 * freed), which Task 2's worker recycling handles.
 */
export function deregisterFont(family: string): unknown | undefined {
  const key = fontKeys.get(family)
  fontKeys.delete(family)
  return key
}
