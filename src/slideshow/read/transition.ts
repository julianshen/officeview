/**
 * Slide transition reader (slideshow spec 6.1, 4.4).
 *
 * - Extension variants come from `mc:AlternateContent` (Choices gated by the
 *   namespace-aware requirements table, then the Fallback) and from
 *   `p:transition/p:extLst/p:ext`. An extension carrying a known transition
 *   element replaces the base effect; any other extension URI is diagnosed.
 * - The shared XML layer strips element and attribute prefixes, so an effect's
 *   namespace is found by its local name in the catalogue (the transition
 *   element names of p, p14, p15 and p159 are disjoint), and `p14:dur` arrives
 *   as `dur` (PresentationML 2006 defines no `dur` on `p:transition`).
 * - `spd` without a duration maps through `SPD_MS`; absent `spd` is the schema
 *   default `fast`.
 * - A `p:transition` with no effect element (it only sets the advance) reads as
 *   element `none` in namespace `p`.
 * - Transitions are read from the slide only. A layout or master transition is
 *   diagnosed by `diagnoseInheritedTransition`, never inherited.
 */
import { attrs, orderedChildren, type XmlNode } from '../../core/xml'
import { CATALOGUE, rowForTransition } from '../catalogue'
import type { DiagnosticSink, TransitionSpec } from '../types'
import { compatibilityBranches } from './requirements'

/** Working values (LibreOffice's importer); to verify against D4 captures. */
export const SPD_MS = Object.freeze({ fast: 500, med: 750, slow: 1000 })

/** Transition element local name to its namespace, from the catalogue. */
const ELEMENT_NS: ReadonlyMap<string, TransitionSpec['ns']> = new Map(
  CATALOGUE.filter(r => r.kind === 'transition' && r.element && r.ns).map(r => [r.element!, r.ns as TransitionSpec['ns']]))

const children = (node: XmlNode | undefined) => orderedChildren(node).filter(([name]) => name !== '#text')
const bool = (v: string | undefined): boolean | undefined => v === '1' || v === 'true' ? true : v === '0' || v === 'false' ? false : undefined
const uint = (v: string | undefined): number | undefined => v !== undefined && /^\s*\d+\s*$/.test(v) && Number.isSafeInteger(Number(v)) ? Number(v) : undefined

/** The slide-level `p:transition` candidates in preference order, with rejected Requires prefixes. */
function candidates(root: XmlNode): { list: XmlNode[]; rejected: string[] } {
  const list: XmlNode[] = []
  const rejected: string[] = []
  for (const [name, node] of children(root)) {
    if (name === 'transition') list.push(node)
    else if (name === 'AlternateContent') {
      const { branches, rejected: r } = compatibilityBranches(node)
      const found = branches.flatMap(b => children(b).filter(([n]) => n === 'transition').map(([, t]) => t))
      const any = children(node).some(([, b]) => children(b).some(([n]) => n === 'transition'))
      if (any) rejected.push(...r)
      list.push(...found)
    }
  }
  return { list, rejected }
}

interface Parsed { spec: TransitionSpec; unknown?: string }

function parseOne(node: XmlNode, slideIndex: number, sink: DiagnosticSink): Parsed {
  const a = attrs(node)
  const diagnose = (code: string, effect: string, message: string) => sink.add({ code, slideIndex, effect, message })
  let effect: XmlNode | undefined
  let effectName: string | undefined
  const unknowns: string[] = []
  let sound: TransitionSpec['sound']
  for (const [name, child] of children(node)) {
    if (name === 'sndAc') {
      for (const [sname, s] of children(child)) {
        if (sname === 'stSnd') {
          const snd = children(s).find(([n]) => n === 'snd')?.[1]
          const rId = attrs(snd).embed
          if (rId) sound = { start: { rId, loop: bool(attrs(s).loop) ?? false } }
          else diagnose('transition-sound', 'stSnd', 'Transition start sound without a relationship')
        } else if (sname === 'endSnd') sound = { end: true }
        else diagnose('transition-unknown', sname, `Unknown transition sound element ${sname}`)
      }
    } else if (name === 'extLst') {
      for (const [ename, ext] of children(child)) {
        if (ename !== 'ext') { diagnose('transition-extension', ename, `Unknown transition extension element ${ename}`); continue }
        const known = children(ext).find(([n]) => ELEMENT_NS.has(n))
        if (known) { effect = known[1]; effectName = known[0] }
        else diagnose('transition-extension', attrs(ext).uri ?? '', `Unknown transition extension ${attrs(ext).uri ?? '(no uri)'}`)
      }
    } else if (ELEMENT_NS.has(name)) {
      if (!effectName) { effect = child; effectName = name }
    } else unknowns.push(name)
  }
  // With a known effect, stray elements are only diagnosed; without one, the
  // first unknown element makes this branch unusable (the caller tries the next).
  const unknown = effectName ? undefined : unknowns[0]
  for (const name of effectName ? unknowns : unknowns.slice(1)) diagnose('transition-unknown', name, `Unknown transition element ${name}`)
  let durationMs = uint(a.dur)
  if (a.dur !== undefined && durationMs === undefined) diagnose('transition-invalid-value', 'dur', `Invalid transition duration ${a.dur}`)
  if (durationMs === undefined) {
    const spd = a.spd ?? 'fast'
    if (spd === 'fast' || spd === 'med' || spd === 'slow') durationMs = SPD_MS[spd]
    else { diagnose('transition-invalid-value', 'spd', `Invalid transition speed ${spd}`); durationMs = SPD_MS.fast }
  }
  const advClick = bool(a.advClick)
  if (a.advClick !== undefined && advClick === undefined) diagnose('transition-invalid-value', 'advClick', `Invalid advClick ${a.advClick}`)
  const advTmMs = uint(a.advTm)
  if (a.advTm !== undefined && advTmMs === undefined) diagnose('transition-invalid-value', 'advTm', `Invalid advTm ${a.advTm}`)
  const element = effectName ?? 'none'
  const ns = effectName ? ELEMENT_NS.get(effectName)! : 'p'
  const preset = effectName === 'prstTrans' ? attrs(effect).prst ?? 'fade' : undefined
  if (effectName && !rowForTransition(ns, element, preset)) diagnose('transition-unknown-preset', `${element}:${preset ?? ''}`, `Transition ${element} ${preset ?? ''} is not catalogued`)
  const spec: TransitionSpec = {
    element, ns,
    ...(preset !== undefined ? { preset } : {}),
    attrs: Object.freeze({ ...attrs(effect) }),
    durationMs,
    advClick: advClick ?? true,
    ...(advTmMs !== undefined ? { advTmMs } : {}),
    ...(sound ? { sound } : {}),
  }
  return { spec, ...(unknown !== undefined ? { unknown } : {}) }
}

/** The slide's own transition, or undefined when it has none. */
export function readTransition(slideXml: XmlNode, slideIndex: number, sink: DiagnosticSink): TransitionSpec | undefined {
  const { list, rejected } = candidates(slideXml)
  for (const prefix of rejected) sink.add({ code: 'transition-requires', slideIndex, effect: prefix, message: `mc:Choice requires unsupported namespace prefix ${prefix}; next branch used` })
  for (let i = 0; i < list.length; i++) {
    const parsed = parseOne(list[i], slideIndex, sink)
    if (parsed.unknown === undefined) return parsed.spec
    sink.add({ code: 'transition-unknown', slideIndex, effect: parsed.unknown, message: `Unknown transition element ${parsed.unknown}${i + 1 < list.length ? '; next branch used' : '; no effect'}` })
    if (i + 1 === list.length) return parsed.spec
  }
  return undefined
}

/** Diagnoses a transition on a layout or master: it is not inherited by the slide. */
export function diagnoseInheritedTransition(partXml: XmlNode, part: 'layout' | 'master', slideIndex: number, sink: DiagnosticSink): boolean {
  if (candidates(partXml).list.length === 0) return false
  sink.add({ code: 'transition-inherited', slideIndex, effect: part, message: `A transition on the slide ${part} is not inherited` })
  return true
}
