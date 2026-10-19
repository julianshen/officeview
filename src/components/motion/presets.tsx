/**
 * Chrome motion (spec §6.4, owner decision A). Lazy-only: check-eager
 * forbids this folder and the motion packages in the eager graph.
 *
 * Only transform and opacity animate, in 120–220 ms on a Fluent decelerate
 * curve. Under prefers-reduced-motion every variant resolves to a zero
 * duration. The grid and page canvases and the content transform never
 * animate: nothing here wraps a canvas.
 */
import { forwardRef, useCallback, useRef, type Ref, type ComponentPropsWithoutRef, type ReactElement, type ReactNode } from 'react'
import { AnimatePresence, LazyMotion, domAnimation, m, useIsPresent, visualElementStore, type Transition, type Variants } from 'motion/react'
import { useControlMedia } from '../xlsx/controlMedia'

export const EASE: [number, number, number, number] = [0.1, 0.9, 0.2, 1]
const ms = (duration: number): Transition => ({ duration: duration / 1000, ease: EASE })

/** Menus and anchored popups: a short drop from their anchor. No scale: a
 * popup's targets keep their full 44px size from the first frame. */
export const popover: Variants = {
  initial: { opacity: 0, y: -4 },
  animate: { opacity: 1, y: 0, transition: ms(160) },
  exit: { opacity: 0, y: -4, transition: ms(120) },
}
/** Dialogs (the sort dialog). */
export const dialog: Variants = {
  initial: { opacity: 0, scale: 0.96 },
  animate: { opacity: 1, scale: 1, transition: ms(200) },
  exit: { opacity: 0, scale: 0.96, transition: ms(140) },
}
/** The mobile filter bottom sheet slides up. */
export const sheet: Variants = {
  initial: { y: '100%' },
  animate: { y: 0, transition: ms(220) },
  exit: { y: '100%', transition: ms(180) },
}
/** Opacity only: safe around a position: fixed descendant and on inline text. */
export const fade: Variants = {
  initial: { opacity: 0 },
  animate: { opacity: 1, transition: ms(120) },
  exit: { opacity: 0, transition: ms(120) },
}
const TICK_MS = 140
/** A changed readout (sheet, zoom, page position) or a confirmation. */
export const tick: Variants = {
  initial: { opacity: 0, y: 4 },
  animate: { opacity: 1, y: 0, transition: ms(TICK_MS) },
  exit: { opacity: 0, y: -4, transition: ms(120) },
}

const INSTANT: Transition = { duration: 0 }
const reduced = new WeakMap<Variants, Variants>()
/** The same targets, reached at once. */
export function reduce(variants: Variants): Variants {
  let out = reduced.get(variants)
  if (!out) {
    out = Object.fromEntries(Object.entries(variants).map(([label, target]) => [label, { ...(target as object), transition: INSTANT }]))
    reduced.set(variants, out)
  }
  return out
}

/** `variants`, or their zero-duration form under prefers-reduced-motion. */
export function useVariants(variants: Variants): Variants {
  return useControlMedia('(prefers-reduced-motion: reduce)') ? reduce(variants) : variants
}

/** The lazy chunks' motion root: domAnimation only (no layout or drag
 * features). `strict`: `motion.*` throws, use `m.*`. */
export function MotionRoot({ children }: { children: ReactNode }): ReactElement {
  return <LazyMotion features={domAnimation} strict>{children}</LazyMotion>
}

/** Exit animations for chrome only; the viewer's first mount never animates. */
export function Presence({ children }: { children: ReactNode }): ReactElement {
  return <AnimatePresence initial={false}>{children}</AnimatePresence>
}

const NAMES = new Map<Variants, string>([[popover, 'popover'], [dialog, 'dialog'], [sheet, 'sheet'], [fade, 'fade'], [tick, 'tick']])

/**
 * The ref every `m.*` element here carries. Motion keeps an element → visual
 * element entry (and with it the last props: handlers, children) until the
 * element itself is collected; React clears its own props from a removed
 * node, so this drops motion's entry too. A removed node that something
 * still holds then no longer keeps the unmounted viewer alive.
 */
export function useMotionRef<T extends Element>(own?: Ref<T> | null): (node: T | null) => void {
  const last = useRef<T | null>(null)
  return useCallback((node: T | null) => {
    if (node) last.current = node
    else if (last.current) { visualElementStore?.delete(last.current); last.current = null }
    if (typeof own === 'function') own(node)
    else if (own) (own as { current: T | null }).current = node
  }, [own])
}

/** Props for an `m.*` element that enters with `v` (and exits with it under
 * <Presence>); `enter: false` renders the first mount in its final state. */
export function useMotionProps<T extends Element = HTMLElement>(v: Variants, enter = true, own?: Ref<T> | null) {
  return { ref: useMotionRef(own), 'data-ov-motion': NAMES.get(v), variants: useVariants(v), initial: enter ? 'initial' : false, animate: 'animate', exit: 'exit' } as const
}

/**
 * A text run keyed by `k`: a new key remounts it and plays `v` in. Nothing
 * exits, so the old text never lingers beside the new one. The first mount
 * renders in its final state.
 */
export function Tick({ k, v = tick, className = 'ov:inline-block', children }: { k: string | number; v?: Variants; className?: string; children: ReactNode }): ReactElement {
  // Decided once per key: the first key renders final; a later key enters,
  // unless it follows the previous change within one tick (a ctrl-wheel or
  // pinch gesture steps the zoom every frame; the readout must not sit dim).
  const state = useRef({ k, enter: false, at: 0 })
  if (state.current.k !== k) {
    const now = performance.now()
    state.current = { k, enter: now - state.current.at > TICK_MS, at: now }
  }
  return <m.span key={k} ref={useMotionRef<HTMLSpanElement>()} data-ov-motion={NAMES.get(v)} data-ov-tick={k} className={className}
    variants={useVariants(v)} initial={state.current.enter ? 'initial' : false} animate="animate">{children}</m.span>
}

const pulses = new WeakMap<Variants, Variants>()
/** The tick as keyframes under two labels: switching labels replays it. */
function pulseOf(v: Variants): Variants {
  let out = pulses.get(v)
  if (!out) {
    const { initial, animate } = v as { initial: { opacity: number; y: number }; animate: { opacity: number; y: number; transition: Transition } }
    const run = { opacity: [initial.opacity, animate.opacity], y: [initial.y, animate.y], transition: animate.transition }
    out = { a: run, b: run }
    pulses.set(v, out)
  }
  return out
}

/**
 * Plays the tick on its child when `k` changes, without remounting it (a
 * focused native control keeps focus). The first mount does not animate.
 */
export function Pulse({ k, children }: { k: string; children: ReactNode }): ReactElement {
  // Each new key flips the label; render-time and idempotent for one key.
  const state = useRef({ k, flip: false })
  if (state.current.k !== k) state.current = { k, flip: !state.current.flip }
  return <m.span ref={useMotionRef<HTMLSpanElement>()} data-ov-motion="tick" data-ov-tick={k} className="ov:inline-block"
    variants={pulseOf(useVariants(tick))} initial={false} animate={state.current.flip ? 'b' : 'a'}>{children}</m.span>
}

type PopProps = ComponentPropsWithoutRef<typeof m.div> & { v: Variants }
/**
 * A popup under <Presence>: enters and exits with `v`. While exiting it is
 * hidden from assistive technology and pointer input, so a closed menu or
 * dialog is gone for everyone but the eye at once.
 */
export const Pop = forwardRef<HTMLDivElement, PopProps>(function Pop({ v, style, ...props }, ref) {
  const present = useIsPresent()
  return <m.div ref={useMotionRef(ref)} data-ov-motion={NAMES.get(v)} {...props} variants={useVariants(v)} initial="initial" animate="animate" exit="exit"
    aria-hidden={present ? props['aria-hidden'] : true} style={present ? style : { ...style, pointerEvents: 'none' }} />
})
