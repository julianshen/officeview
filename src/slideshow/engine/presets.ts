// S1-11: preset rebuild (spec 6.2, D1). For a CTn carrying only preset
// metadata, this module synthesizes keep-alike behaviour nodes approximating
// the named preset. Known ids rebuild whatever the catalogue names (the rows
// only document the mapping; the synthesis is this task's own preset tables);
// unknown ids fall to the class stand-ins.
import { rowForPreset } from '../catalogue'
import { standInFor } from './stand-ins'
import type { AnimEffectBehaviour, BehaviourNode, CTn, DiagnosticSink, SetBehaviour, TargetRef } from '../types'

type SubTypeMap = Readonly<Record<string, string>>

/** Class-level stand-in nodes. `none` emits no state (documented for the report). */
function classStandIn(cls: string, ctn: CTn, target: TargetRef): BehaviourNode[] {
  const kind = standInFor(cls)
  if (kind === 'visible')
    return [{
      kind: 'set', ctn, target, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'visible',
    } satisfies SetBehaviour]
  if (kind === 'hidden')
    return [{
      kind: 'set', ctn, target, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'hidden',
    } satisfies SetBehaviour]
  return []
}

/** maskKind the filter arg mapping for the wipe presets; subtype → direction. */
const WIPE_SIDES: SubTypeMap = { '8': 'right', '2': 'left', '1': 'up', '4': 'down' }

export function rebuildFromPreset(ctn: CTn, target: TargetRef, sink?: DiagnosticSink): BehaviourNode[] | undefined {
  const cls = ctn.presetClass
  if (!cls) return undefined
  const diagnose = (code: string, effect: string, message: string): void =>
    sink?.add({ code, effect, message })
  const id = ctn.presetId
  const subtype = ctn.presetSubtype === undefined ? undefined : String(ctn.presetSubtype)

  const known = id === undefined ? undefined : rowForPreset(cls, id, ctn.presetSubtype)
  const wipe = known !== undefined && id === 22
  const fadeKnown = known !== undefined && id === 10

  const effectAttr = (a: string[] = ['style.visibility']): Pick<AnimEffectBehaviour, 'kind' | 'attrNames' | 'additive' | 'accumulate'> => ({
    kind: 'animEffect', attrNames: a, additive: 'base', accumulate: 'none',
  })

  const approx = (message: string): BehaviourNode[] | undefined => {
    diagnose('preset-approximate', `preset ${cls}:${id}`, message)
    return undefined
  }
  try {
    switch (cls) {
      case 'entr':
      case 'exit': {
        if (wipe) {
          const side = WIPE_SIDES[subtype ?? '8'] ?? 'right'
          const node: AnimEffectBehaviour = {
            ...effectAttr(), ctn, target,
            transition: cls === 'exit' ? 'out' : 'in',
            filter: `wipe(${side})`,
          }
          approx('entr 22 wipe rebuilt as the mask wipe keep-alike')
          return [node]
        }
        if (fadeKnown) {
          const node: AnimEffectBehaviour = { ...effectAttr(), ctn, target, transition: cls === 'exit' ? 'out' : 'in', filter: 'fade()' }
          approx('entr 10 fade rebuilt as the fade filter keep-alike')
          return [node]
        }
        if (known !== undefined && id === 1) {
          const node: SetBehaviour = { kind: 'set', ctn, target, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'visible' }
          approx('appear preset rebuilt as a plain visibility set')
          return [node]
        }
        approx('preset id has no keep-alike table; using the class stand-in')
        return classStandIn(cls, ctn, target)
      }
      case 'emph':
      case 'path':
      default:
        diagnose('preset-approximate', `preset ${cls}:${id}`, `no preset table for ${cls}${id === undefined ? '' : `:${id}`}; using the class stand-in`)
        return classStandIn(cls, ctn, target)
    }
  } catch {
    return classStandIn(cls, ctn, target)
  }
}
