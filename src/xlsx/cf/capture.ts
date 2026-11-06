/**
 * Streaming CF subtree collector (spec §6.2; Task 4 wiring seam).
 *
 * Drives off `XmlEvents` (the same contract `sheetRecords` uses): namespace
 * scopes are tracked per element, and only the two CF subtrees are captured —
 * direct-child `<conditionalFormatting>` of the worksheet and
 * `x14:conditionalFormattings` under the CF `ext` — everything else is
 * ignored without allocation. Captured nodes are registered namespace-aware so
 * the shared `finalizeCfSubtrees` core sees exactly what the resident parser
 * sees; URI policy (strict x14/xm, worksheet-namespace base) lives in that
 * core, not here.
 *
 * Bounds (structural, pre-screening): at most 32768 captured nodes, depth 64,
 * `textUnits` per text node. Exceeding them, or ending with unclosed frames,
 * yields the same sheet-wide `cf-resource-limit` sentinel as the resident
 * path rather than a partial model. Exact per-rule caps are enforced by the
 * shared finalizer, so both paths refuse the same oversized inputs.
 */
import { getChildren, registerNamespaceAware, type XmlNode } from '../../core/xml'
import { localName, type XmlEvents } from '../../core/xml-stream'
import { CF_LIMITS } from './limits'
import type { CfMemory, CfReservation } from './limits'
import { CF_X14_EXTENSION_URI, finalizeCfSubtrees, type CfParseMeta, type CfStylesSource } from './parse'
import { X14_MAIN_NS } from '../tableStyles/parseStylesExtras'
import type { CfLease, CfSheetRules } from './model'

const MAX_CAPTURE_NODES = 32768
const MAX_CAPTURE_DEPTH = 64

interface Frame {
  local: string
  uri: string | undefined
  scope: ReadonlyMap<string, string>
  attrs: Record<string, string>
  children: Map<string, XmlNode[]>
  childUris: Map<string, Array<string | undefined>>
  texts: string[]
  textLen: number
}

export interface CfStreamHandle {
  readonly events: XmlEvents
  finish(): CfLease<CfSheetRules>
}

export function createCfStreamEvents(
  styles: CfStylesSource,
  memory: CfMemory,
  meta: CfParseMeta,
): CfStreamHandle {
  const scopes: Array<ReadonlyMap<string, string>> = []
  const stack: Array<{ local: string; uri: string | undefined; extUri?: string }> = []
  // One entry per captured open element: a built frame, or null once over cap.
  const entries: Array<{ frame: Frame | null; isRoot: boolean }> = []
  let sheetNs: string | undefined
  let sheetSeen = false
  let nodeCount = 0
  let overCap = false
  let finished = false
  // Working reservation: frames, copied attributes and accumulated text are
  // admitted as they grow, before they are built.
  let hold: CfReservation | undefined
  const admit = (delta: number): boolean => {
    if (overCap) return false
    try {
      if (!hold) {
        const first = memory.reserve(256)
        if (first.kind === 'unresolved') {
          overCap = true
          return false
        }
        hold = first
      }
      if (!hold.grow(hold.bytes + delta)) {
        overCap = true
        return false
      }
      return true
    } catch {
      overCap = true
      return false
    }
  }
  // Wrong-namespace containers skipped during capture; published as diagnostics.
  const skipped = new Set<string>()
  const SKIP_WRAPPER = 'non-x14 conditionalFormattings content is skipped'
  const SKIP_PATH = 'conditionalFormatting extension outside the worksheet namespace is skipped'
  // Current capture root kind; undefined outside a root.
  let rootKind: 'base' | 'x14' | undefined
  const base: XmlNode[] = []
  const x14blocks: XmlNode[] = []

  const uriOf = (qname: string, scope: ReadonlyMap<string, string>): string | undefined => {
    const colon = qname.indexOf(':')
    return scope.get(colon < 0 ? '' : qname.slice(0, colon))
  }

  const finishFrame = (frame: Frame): XmlNode => {
    // Wrap text-only children as registered nodes: the shared finalizer
    // filters by elementNamespaceUri, which is undefined for bare strings.
    const wrapped = new Map<string, XmlNode[]>()
    for (const [name, children] of frame.children) {
      const uris = frame.childUris.get(name) ?? []
      wrapped.set(
        name,
        children.map((child, index) =>
          typeof child === 'string'
            ? registerNamespaceAware({ '@attrs': {}, '#text': child }, uris[index], frame.scope, new Map())
            : child,
        ),
      )
    }
    const node: XmlNode = { '@attrs': frame.attrs }
    if (frame.texts.length > 0) node['#text'] = frame.texts.join('')
    for (const [name, children] of wrapped) {
      node[name] = children.length === 1 ? children[0] : [...children]
    }
    registerNamespaceAware(node, frame.uri, frame.scope, wrapped)
    return node
  }

  const events: XmlEvents = {
    open(qname, attrNames, attrValues, count, selfClosing) {
      const parentScope = scopes.length > 0 ? scopes[scopes.length - 1]! : new Map<string, string>()
      let scope: ReadonlyMap<string, string> = parentScope
      for (let i = 0; i < count; i++) {
        const name = attrNames[i]!
        if (name !== 'xmlns' && !name.startsWith('xmlns:')) continue
        if (scope === parentScope) scope = new Map(parentScope)
        ;(scope as Map<string, string>).set(name === 'xmlns' ? '' : name.slice(6), attrValues[i]!)
      }
      const local = localName(qname)
      const uri = uriOf(qname, scope)
      const depth = stack.length
      if (!sheetSeen) {
        sheetSeen = true
        if (local === 'worksheet') sheetNs = uri
      }

      const isRoot = rootKind === undefined
      let kind: 'base' | 'x14' | undefined
      if (isRoot) {
        if (depth === 1 && stack[0]?.local === 'worksheet' && local === 'conditionalFormatting') {
          kind = 'base'
        } else if (
          depth === 3 &&
          stack[0]?.local === 'worksheet' &&
          stack[1]?.local === 'extLst' &&
          stack[2]?.local === 'ext' &&
          stack[2]?.extUri === CF_X14_EXTENSION_URI &&
          local === 'conditionalFormattings'
        ) {
          if (stack[1]?.uri === sheetNs && stack[2]?.uri === sheetNs && uri === X14_MAIN_NS) {
            kind = 'x14'
          } else {
            // Claimed CF extension with a non-conforming path or wrapper:
            // skip the subtree but say so (fail-closed, never silent).
            skipped.add(stack[1]?.uri === sheetNs && stack[2]?.uri === sheetNs ? SKIP_WRAPPER : SKIP_PATH)
          }
        }
        if (kind !== undefined) rootKind = kind
      } else {
        kind = rootKind
      }

      if (kind !== undefined) {
        if (overCap || nodeCount >= MAX_CAPTURE_NODES || entries.length >= MAX_CAPTURE_DEPTH) {
          overCap = true
          entries.push({ frame: null, isRoot })
        } else {
          nodeCount++
          let attrBytes = 0
          for (let i = 0; i < count; i++) {
            const name = attrNames[i]!
            if (name === 'xmlns' || name.startsWith('xmlns:')) continue
            attrBytes += localName(name).length + String(attrValues[i]).length
          }
          if (!admit(192 + attrBytes * 2)) {
            entries.push({ frame: null, isRoot })
          } else {
            const frameAttrs: Record<string, string> = {}
            for (let i = 0; i < count; i++) {
              const name = attrNames[i]!
              if (name === 'xmlns' || name.startsWith('xmlns:')) continue
              const stripped = localName(name)
              if (!(stripped in frameAttrs)) frameAttrs[stripped] = attrValues[i]!
            }
            entries.push({ frame: { local, uri, scope, attrs: frameAttrs, children: new Map(), childUris: new Map(), texts: [], textLen: 0 }, isRoot })
          }
        }
      }

    const extUriOf = (): string | undefined => {
      if (local !== 'ext') return undefined
      for (let i = 0; i < count; i++) {
        const name = attrNames[i]!
        if (name === 'xmlns' || name.startsWith('xmlns:')) continue
        if (localName(name) === 'uri') return attrValues[i]
      }
      return undefined
    }
      stack.push({ local, uri, ...(local === 'ext' ? { extUri: extUriOf() } : {}) })
      scopes.push(scope)
      if (selfClosing) events.close(qname)
    },

    close(_qname) {
      stack.pop()
      scopes.pop()
      if (rootKind === undefined || entries.length === 0) return
      const entry = entries.pop()!
      if (entry.frame !== null && !overCap) {
        const node = finishFrame(entry.frame)
        // Nearest enclosing real frame, if any.
        let parent: Frame | undefined
        for (let i = entries.length - 1; i >= 0; i--) {
          if (entries[i]!.frame !== null) {
            parent = entries[i]!.frame!
            break
          }
        }
        if (parent) {
          const list = parent.children.get(entry.frame.local) ?? []
          list.push(node)
          parent.children.set(entry.frame.local, list)
          const uris = parent.childUris.get(entry.frame.local) ?? []
          uris.push(entry.frame.uri)
          parent.childUris.set(entry.frame.local, uris)
        } else if (entry.isRoot) {
          if (rootKind === 'base') {
            base.push(node)
          } else {
            for (const inner of getChildren(node, 'conditionalFormatting')) x14blocks.push(inner)
          }
        }
      }
      if (entry.isRoot) rootKind = undefined
    },

    text(text, _more) {
      void _more
      for (let i = entries.length - 1; i >= 0; i--) {
        const frame = entries[i]!.frame
        if (frame !== null) {
          if (frame.textLen + text.length > CF_LIMITS.textUnits) {
            overCap = true
          } else if (!admit(text.length * 2)) {
            // Admission failed: withhold everything at finish.
          } else {
            frame.texts.push(text)
            frame.textLen += text.length
          }
          return
        }
      }
    },
  }

  function releaseHold(): void {
    // Idempotent: the hold exists only in the construction phase.
    try {
      hold?.release()
    } catch {
      /* best effort on teardown paths */
    } finally {
      hold = undefined
    }
  }

  function finish(): CfLease<CfSheetRules> {
    if (finished) throw new Error('CF capture already finished')
    finished = true
    try {
      if (entries.length > 0 || overCap || rootKind !== undefined) {
        // Truncated stream or structural overflow: withhold everything.
        return finalizeCfSubtrees({ base: [], x14: [], sheetNs, truncated: true, skipped: [...skipped] }, styles, memory, meta)
      }
      return finalizeCfSubtrees({ base, x14: x14blocks, sheetNs, skipped: [...skipped] }, styles, memory, meta)
    } finally {
      // The finalizer admits its own normalized copies; drop the transient hold.
      releaseHold()
    }
  }

  return { events, finish }
}
