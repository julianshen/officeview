/** The server's CanvasBackend (§9.2): @napi-rs/canvas with the pinned pack
 * registered and every handed-out 2-D context font-pinned (Task 4 note). */
import { GlobalFonts } from '@napi-rs/canvas'
import { napiBackend } from '../../scripts/stream-s1/napi-backend'
import type { CanvasBackend, CanvasLike } from '../../src/core/canvas-backend'
import { loadPack, pinFonts, registerPack } from './fonts'
import type { Pack } from '../../scripts/stream-s1/fontpack-lock'

/** A canvas whose 2-D context view pins fonts: alias families always; with `packOnly`, every family (Task 14). */
export function pinCanvas<T extends CanvasLike>(canvas: T, pack: Pack, pin: { packOnly?: boolean } = {}): T {
  const native = (canvas as unknown as { getContext(kind: string, o?: unknown): object }).getContext.bind(canvas)
  const cache = new Map<string, object>()
  // The canvas itself stays native (drawImage/createPattern need the real
  // object); only its 2-D context view pins fonts.
  ;(canvas as unknown as { getContext(kind: string, o?: unknown): unknown }).getContext = (kind: string, o?: unknown) => {
    const ctx = native(kind, o)
    if (kind !== '2d' || !ctx) return ctx
    let view = cache.get(kind)
    if (!view) { view = pinFonts(ctx, pack, pin); cache.set(kind, view) }
    return view
  }
  return canvas
}

export function serverBackend(pack: Pack = loadPack()): CanvasBackend {
  registerPack(GlobalFonts, pack)
  return { ...napiBackend, name: '@napi-rs/canvas+pack', createCanvas: (w, h) => pinCanvas(napiBackend.createCanvas(w, h), pack) }
}

/**
 * The show path's backend (piece 4 Task 14 ruling): `backend`'s canvases with fonts resolved through the pack
 * only — an alias to its pack face, a generic to its pack family, an unknown family to the pack's fallback, never
 * a host face — so layers, keyframes and Morph assets never depend on the host's installed fonts. Units, print
 * and notes keep alias-only pinning (a follow-up: `docs/validation/streaming-s3-p4.md`).
 */
export function showBackend(backend: CanvasBackend, pack: Pack = loadPack()): CanvasBackend {
  return { ...backend, name: `${backend.name}+pack-only`, createCanvas: (w, h) => pinCanvas(backend.createCanvas(w, h), pack, { packOnly: true }) }
}
