/** The server's CanvasBackend (§9.2): @napi-rs/canvas with the pinned pack
 * registered and every handed-out 2-D context font-pinned (Task 4 note). */
import { GlobalFonts } from '@napi-rs/canvas'
import { napiBackend } from '../../scripts/stream-s1/napi-backend'
import type { CanvasBackend, CanvasLike } from '../../src/core/canvas-backend'
import { loadPack, pinFonts, registerPack } from './fonts'
import type { Pack } from '../../scripts/stream-s1/fontpack-lock'

export function pinCanvas<T extends CanvasLike>(canvas: T, pack: Pack): T {
  const native = (canvas as unknown as { getContext(kind: string, o?: unknown): object }).getContext.bind(canvas)
  const cache = new Map<string, object>()
  // The canvas itself stays native (drawImage/createPattern need the real
  // object); only its 2-D context view pins fonts.
  ;(canvas as unknown as { getContext(kind: string, o?: unknown): unknown }).getContext = (kind: string, o?: unknown) => {
    const ctx = native(kind, o)
    if (kind !== '2d' || !ctx) return ctx
    let view = cache.get(kind)
    if (!view) { view = pinFonts(ctx, pack); cache.set(kind, view) }
    return view
  }
  return canvas
}

export function serverBackend(pack: Pack = loadPack()): CanvasBackend {
  registerPack(GlobalFonts, pack)
  return { ...napiBackend, name: '@napi-rs/canvas+pack', createCanvas: (w, h) => pinCanvas(napiBackend.createCanvas(w, h), pack) }
}
