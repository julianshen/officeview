// S1-11 stand-ins (spec D3). For preset classes the rebuild tables cannot
// approximate, the state keeps the class's nearest effect; each use surfaces
// one diagnosis (the sink dedupes).
import type { DiagnosticSink } from '../types'

export type StandInKind = 'visible' | 'none' | 'hidden' | 'path-end'

/** 'entr' → visible, 'emph' → none, 'exit' → hidden, 'path' → path-end. */
export function standInFor(cls: string): StandInKind {
  switch (cls) {
    case 'entr': return 'visible'
    case 'emph': return 'none'
    case 'exit': return 'hidden'
    case 'path': return 'path-end'
    default: return 'none'
  }
}

/** Transition stages (spec 6.4) without a catalogue row reuse the fade. */
export function transitionStandIn(sink?: DiagnosticSink, effect?: string): 'fade' {
  sink?.add({ code: 'transition-stand-in', effect: effect ?? 'transition', message: 'Unknown transition; using a fade stand-in' })
  return 'fade'
}
