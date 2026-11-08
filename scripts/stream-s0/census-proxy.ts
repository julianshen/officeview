/**
 * Stage-0 Task 3 — recorder census proxy.
 *
 * `wrapCensus(ctx)` returns a Proxy over a real 2-D context that forwards
 * everything and counts what happened:
 *  - method calls, by name + arg shape (`drawImage/source=canvas#9`,
 *    `fillText/3`, `measureText/1`)
 *  - property writes, by name + value class (`fillStyle=<string>`)
 *  - text-index protocol entries (plan Task 3: `RECORD_TEXT` access is
 *    forwarded — the census hands painters a recorder — but it is never a
 *    canvas op: entries land in `recordText`, not in `methods`)
 *  - gradient/pattern wrapper calls (`addColorStop`, the §7.3 recorded
 *    wrapper `setTransform`) as `{wrapper}.{method}` keys
 *
 * The recorded subset is copied from the normative contract §7.3/§7.5 as
 * `RECORDED_SUBSET`; every method/property a corpus paint touches through the
 * proxy that the subset does not cover is a RED spec gap (`notInSubset`) —
 * a gap is reported, never fixed by editing the subset (Global Constraints).
 */
import { RECORD_TEXT } from '../../src/core/text-recording'

import { RECORDED_SUBSET } from '../../src/stream/record/subset'
export { RECORDED_SUBSET } from '../../src/stream/record/subset'

/** gradient/pattern wrapper calls (§7.3 wrapper setTransform + addColorStop). */
const WRAPPER_SUBSET = {
  methods: new Set(['addColorStop', 'setTransform']),
} as const

/** Round-9 census report per plan Task 3. */
export interface CensusReport {
  methods: Record<string, number>
  props: Record<string, { count: number; values: string[] }>
  notInSubset: string[]
  faceResolution: Array<{ font: string; widthNode: number; widthNapi: number; relDiff: number }>
  recordText: Array<{ text: string; scope: string; start: number; end: number }>
  /** every distinct `ctx.font` string the census saw (face-resolution input). */
  fontStrings: string[]
  /** round-9 extras for Task 2: classed painter evidence. */
  classEvidence: {
    shadowBlurSeen: boolean
    setTransformSeen: boolean
    createPatternSeen: boolean
    drawImageCanvasSourceSeen: boolean
  }
}

export interface CensusProxy {
  ctx: CanvasRenderingContext2D
  /** Only Task 2 consumes this (units with `shadowBlur > 0`, `setTransform`,
   * `createPattern`, or a canvas-source `drawImage`). */
  classEvidence(): CensusReport['classEvidence']
  report(): CensusReport
}


function methodShape(name: string, target: Rec, args: IArguments): string {
  const values = Array.from(args as ArrayLike<unknown>)
  if (name === 'drawImage') {
    const first = values[0] as { constructor?: { name?: string }; width?: unknown } | undefined
    const source = isCanvasSource(first) ? 'source=canvas' : isImageSource(first) ? 'source=image' : 'source=unknown'
    const arity = values.length === 9 ? '9' : values.length === 5 ? '5' : values.length === 3 ? '3' : String(values.length)
    return `${name}/${source}#${arity}`
  }
  if (name === 'fillText' || name === 'strokeText') return `${name}/${values.length}`
  void target
  return `${name}/${values.length}`
}

function isCanvasSource(value: unknown): boolean {
  const ctor = (value as { constructor?: { name?: string } } | undefined)?.constructor
  // napi names its main canvas `Canvas` (SVGCanvas for SVG); node-canvas's
  // own class carries the same `Canvas` name — matching on constructor name
  // alone fails, so detect a real paintable surface instead (napi instances
  // use getContext, node-canvas instances use getContext too — probe both).
  if (ctor?.name === 'Canvas' || ctor?.name === 'SVGCanvas' || ctor?.name === 'CanvasSurface') return true
  if (typeof value !== 'object' || value === null) return false
  return 'getContext' in (value as object) && 'width' in (value as object) && 'height' in (value as object)
}

function isImageSource(value: unknown): boolean {
  const ctor = (value as { constructor?: { name?: string } } | undefined)?.constructor
  return ctor?.name === 'Image' || (typeof value === 'object' && value !== null && 'naturalWidth' in (value as object))
}

const valueTag = (v: unknown): string => {
  if (v === null) return 'null'
  if (Array.isArray(v)) return 'array'
  const ctor = (v as { constructor?: { name?: string } }).constructor
  return typeof v === 'object' || typeof v === 'function' ? ctor?.name ?? typeof v : typeof v
}

type Rec = Record<string, any>
export function wrapCensus(ctx: CanvasRenderingContext2D): CensusProxy {
  const fontStringsForRes = new Set<string>()
  const methods: Record<string, number> = {}
  const props: Record<string, { count: number; values: string[] }> = {}
  const notInSubset: string[] = []
  const shadowBlurSeen = { value: false }
  const setTransformSeen = { value: false }
  const createPatternSeen = { value: false }
  const drawImageCanvasSourceSeen = { value: false }

  const count = (key: string, inSubset: boolean): void => {
    methods[key] = (methods[key] ?? 0) + 1
    if (!inSubset && !notInSubset.includes(key)) notInSubset.push(key)
  }

  function isFn2(v: unknown): v is (...args: unknown[]) => unknown {
    return typeof v === 'function'
  }

  const proxy = new Proxy(ctx as unknown as Rec, {
    get(target, prop) {
      if (prop === RECORD_TEXT) return
      const name = String(prop)
      const value = (target as Rec)[String(prop)]
      if (isFn2(value)) {
        return (...args: unknown[]) => {
          let key = methodShape(name, target, args as unknown as IArguments)
          if (name === 'createPattern') createPatternSeen.value = true
          if (name === 'drawImage' && key.includes('source=canvas')) drawImageCanvasSourceSeen.value = true
          if (name === 'setTransform' || name === 'transform') setTransformSeen.value = true
          count(key, (RECORDED_SUBSET.methods as ReadonlySet<string>).has(name))
          const out = Reflect.apply(value, target, args)
          // Gradients/patterns are §7.3 wrapper-recorded: their calls are
          // census evidence, classified against WRAPPER_SUBSET.
          if ((name === 'createPattern' || name === 'createLinearGradient' || name === 'createRadialGradient') && typeof out === 'object' && out !== null) {
            return wrapWrapper(name, out)
          }
          return out
        }
        function wrapWrapper(wname: string, wrapper: object): object {
          // Native-class methods (addColorStop, setTransform) throw when
          // detached from their `this`: call bound to the wrapper (round-9
          // "Illegal invocation" fix for the §7.3 gradient wrapper).
          return new Proxy(wrapper, {
            get(w: object, p: string) {
              const fn = (w as Rec)[p]
              if (isFn2(fn)) {
                return (...args: unknown[]) => {
                  const key = `${wname}.${p}/${args.length}`
                  methods[key] = (methods[key] ?? 0) + 1
                  const inWrapper = (WRAPPER_SUBSET.methods as ReadonlySet<string>).has(p)
                  const inSubset = (RECORDED_SUBSET.methods as ReadonlySet<string>).has(wname)
                  if (!(inWrapper || inSubset) && !notInSubset.includes(key)) notInSubset.push(key)
                  return fn.apply(w, args)
                }
              }
              return (w as Rec)[p]
            },
          })
        }
      }
      // Direct read: Reflect.get with a proxy receiver makes native canvas
      // getters throw "Illegal invocation" (the round-6 text-index recorder
      // hit the same trap and resolved it with direct target reads).
      return (target as Rec)[String(prop)]
    },
    has(target, prop) {
      return Reflect.has(target as object, prop)
    },
    // node-canvas/napi setters reject Reflect.set called with a proxy
    // receiver ("Invalid argument"): assign straight to the target, as the
    // round-6 text-index recorder does.
    set(target, prop, value) {
      const name = String(prop)
      if (prop === RECORD_TEXT) return true
      if (name === 'shadowBlur' && typeof value === 'number' && value > 0) shadowBlurSeen.value = true
      const inSubset = (RECORDED_SUBSET.setters as ReadonlySet<string>).has(name)
      count(name, inSubset)
      const slot = props[name] ?? (props[name] = { count: 0, values: [] })
      slot.count += 1
      const tag = valueTag(value)
      const label = tag === 'string' || tag === 'number' ? `${tag}:${String(value)}` : tag
      if (slot.values.length < 20 && !slot.values.includes(label)) slot.values.push(label)
      // Round-9: the face-resolution census needs every distinct font string —
      // the `.font` sample list must not be capped (round-10 F2 note).
      if (name === 'font' && typeof value === 'string' && value.length > 0) fontStringsForRes.add(value)
      ;(target as Rec)[String(prop)] = value
      return true
    },
  }) as unknown as CanvasRenderingContext2D

  return {
    ctx: proxy,
    classEvidence(): CensusReport['classEvidence'] {
      return {
        shadowBlurSeen: shadowBlurSeen.value,
        setTransformSeen: setTransformSeen.value,
        createPatternSeen: createPatternSeen.value,
        drawImageCanvasSourceSeen: drawImageCanvasSourceSeen.value,
      }
    },
    report(): CensusReport {
      return {
        methods,
        props,
        notInSubset,
        faceResolution: [],
        recordText: [],
        fontStrings: [...fontStringsForRes],
        classEvidence: {
          shadowBlurSeen: shadowBlurSeen.value,
          setTransformSeen: setTransformSeen.value,
          createPatternSeen: createPatternSeen.value,
          drawImageCanvasSourceSeen: drawImageCanvasSourceSeen.value,
        },
      }
    },
  }
}

/** round-9 scratch: face resolution is implemented in t3-census.ts (napi vs
 * node-canvas widths per distinct `ctx.font`), so the census leaves the field
 * populated there. */
