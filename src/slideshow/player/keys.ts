/**
 * Audience input mapping (slideshow spec 4.3 key table). Pure: events in,
 * `ShowInput` out. Precedence beyond the mapping (kiosk, the number buffer,
 * black/white resume) is applied by the session's `dispatch`.
 *
 * Wheel and swipe (my calls beyond the brief's 120 units and 50 px):
 * - a wheel step needs 120 delta units in one direction (line and page
 *   deltas are converted at 40 px per line and 800 px per page);
 * - after a step, wheel input is dropped for WHEEL_GAP_MS on the show
 *   clock, so a free-spinning wheel or a trackpad fling moves one step;
 * - partial deltas are forgotten after WHEEL_IDLE_MS without wheel input,
 *   or when the direction changes;
 * - a swipe is a one-finger touch that ends at least 50 CSS px from its
 *   start horizontally, and further horizontally than vertically.
 */
import type { ShowInput } from './session'

export const WHEEL_STEP = 120
export const WHEEL_GAP_MS = 250
export const WHEEL_IDLE_MS = 500
export const SWIPE_PX = 50

const NEXT = new Set([' ', 'Spacebar', 'ArrowRight', 'Right', 'ArrowDown', 'Down', 'PageDown', 'n', 'N'])
const PREVIOUS = new Set(['ArrowLeft', 'Left', 'ArrowUp', 'Up', 'PageUp', 'p', 'P', 'Backspace'])
const BUFFER_MAX = 6

/**
 * One key of the spec 4.3 table. `buffer` is the session's number-entry
 * buffer: Enter after digits is "go to slide n" (1-based on screen, 0-based
 * here); Enter alone is `enter` (next). Esc and `-` always map to `exit`,
 * whatever else is pressed with them. Keys with Ctrl, Alt or Meta are left
 * to the browser.
 */
export function keyToInput(e: KeyboardEvent, buffer: string): ShowInput | undefined {
  const k = e.key
  if (k === 'Escape' || k === 'Esc' || k === '-' || k === 'Subtract') return { kind: 'exit' }
  if (e.ctrlKey || e.altKey || e.metaKey) return undefined
  if (k.length === 1 && k >= '0' && k <= '9') return buffer.length < BUFFER_MAX ? { kind: 'digit', d: Number(k) } : undefined
  if (k === 'Enter') return buffer ? { kind: 'goto', index: Number(buffer) - 1 } : { kind: 'enter' }
  if (NEXT.has(k)) return { kind: 'next' }
  if (PREVIOUS.has(k)) return { kind: 'previous' }
  if (k === 'Home') return { kind: 'first' }
  if (k === 'End') return { kind: 'last' }
  if (k === 'b' || k === 'B' || k === '.') return { kind: 'black' }
  if (k === 'w' || k === 'W' || k === ',') return { kind: 'white' }
  if (k === 's' || k === 'S' || k === '+') return { kind: 'pause' }
  return undefined
}

/** Keys that never resume a black or white screen on their own (they arrive with the key that matters). Tab is excluded by the session too: it is never taken from the page. */
export const MODIFIER_KEYS: ReadonlySet<string> = new Set(['Shift', 'Control', 'Alt', 'Meta', 'CapsLock', 'AltGraph', 'Fn', 'OS'])

export interface GestureState {
  /** The show clock, injected. */
  now(): number
  wheel: number
  wheelT: number
  stepT: number
  touch?: { x: number; y: number }
}

export function createGestureState(now: () => number): GestureState {
  return { now, wheel: 0, wheelT: -Infinity, stepT: -Infinity }
}

type Point = { clientX: number; clientY: number }
const first = (list: ArrayLike<Point> | undefined): Point | undefined => (list && list.length > 0 ? list[0] : undefined)

/** Wheel down/up and swipe left/right; undefined for everything else (touch start, a short swipe, a partial or debounced notch). */
export function gestureToInput(e: WheelEvent | TouchEvent, state: GestureState): ShowInput | undefined {
  if (e.type === 'wheel') {
    const w = e as WheelEvent
    const t = state.now()
    const scale = w.deltaMode === 1 ? 40 : w.deltaMode === 2 ? 800 : 1
    const delta = (Number.isFinite(w.deltaY) ? w.deltaY : 0) * scale
    if (t - state.wheelT > WHEEL_IDLE_MS || Math.sign(delta) !== Math.sign(state.wheel)) state.wheel = 0
    state.wheelT = t
    if (t - state.stepT < WHEEL_GAP_MS) return undefined
    state.wheel += delta
    if (Math.abs(state.wheel) < WHEEL_STEP) return undefined
    const down = state.wheel > 0
    state.wheel = 0
    state.stepT = t
    return { kind: down ? 'next' : 'previous' }
  }
  const touch = e as TouchEvent
  if (e.type === 'touchstart') {
    const p = touch.touches?.length === 1 ? first(touch.touches) : undefined
    state.touch = p ? { x: p.clientX, y: p.clientY } : undefined
    return undefined
  }
  if (e.type === 'touchcancel') { state.touch = undefined; return undefined }
  if (e.type !== 'touchend' || !state.touch) return undefined
  const start = state.touch
  state.touch = undefined
  const p = first(touch.changedTouches)
  if (!p) return undefined
  const dx = p.clientX - start.x, dy = p.clientY - start.y
  if (!(Math.abs(dx) >= SWIPE_PX) || Math.abs(dx) <= Math.abs(dy)) return undefined
  return { kind: dx < 0 ? 'next' : 'previous' }
}
