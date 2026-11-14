/**
 * The show path's canvases (piece 4 Task 14 controller ruling): the installed backend (the server's, in a pool
 * worker) with every font string resolved through the pinned pack only, so no layer, keyframe or Morph asset
 * depends on the host's installed fonts (`showBackend`).
 */
import { canvasBackend, type CanvasBackend } from '../../../src/core/canvas-backend'
import { showBackend } from '../backend'

const wrapped = new WeakMap<CanvasBackend, CanvasBackend>()
export function showCanvasBackend(): CanvasBackend {
  const b = canvasBackend()
  let s = wrapped.get(b)
  if (!s) wrapped.set(b, s = showBackend(b))
  return s
}
