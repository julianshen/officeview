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

/** Shared empty scope: the document root inherits no bindings, and scopes are
 * never mutated in place (extensions copy first), so one instance is safe. */
const EMPTY_SCOPE: ReadonlyMap<string, string> = new Map()

/** Local-part length without allocating the sliced string. */
function localLen(qname: string): number {
  const colon = qname.indexOf(':')
  return colon < 0 ? qname.length : qname.length - colon - 1
}

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
      // Terminal state is inert: post-finish events must neither build nor charge.
      if (finished) return
      // Account for every attribute with plain arithmetic (no allocation), so
      // admission precedes any construction below. xmlns prefixes and values
      // are retained in scope maps; ordinary attributes are copied into
      // captured frames.
      let frameAttrBytes = 0
      let bindingsBytes = 0
      let newBindings = 0
      for (let i = 0; i < count; i++) {
        const name = attrNames[i]!
        const valueLen = String(attrValues[i]).length
        if (name === 'xmlns' || name.startsWith('xmlns:')) {
          newBindings++
          // Retained prefix slice, retained value string and the map slot.
          bindingsBytes += ((name === 'xmlns' ? 0 : name.length - 6) + valueLen) * 2 + 64
        } else {
          frameAttrBytes += localLen(name) + valueLen
        }
      }
      const parentScope = scopes.length > 0 ? scopes[scopes.length - 1]! : EMPTY_SCOPE

    const extUriOf = (): string | undefined => {
      if (localName(qname) !== 'ext') return undefined
      for (let i = 0; i < count; i++) {
        const name = attrNames[i]!
        if (name === 'xmlns' || name.startsWith('xmlns:')) continue
        if (localName(name) === 'uri') return attrValues[i]
      }
      return undefined
    }
      const local = localName(qname)
      // Admit namespace storage BEFORE copying or extending any scope.
      // Inherited bindings were admitted at their own declaration; this covers
      // the copy of the inherited map plus every new binding retained below.
      let scope: ReadonlyMap<string, string> = parentScope
      if (newBindings > 0) {
        if (!admit(64 + parentScope.size * 64 + bindingsBytes)) {
          // Refuse without constructing or retaining new scope state: keep
          // tracking balanced by reusing the parent scope.
          overCap = true
          stack.push({ local, uri: uriOf(qname, parentScope), ...(local === 'ext' ? { extUri: extUriOf() } : {}) })
          scopes.push(parentScope)
          if (selfClosing) events.close(qname)
          return
        }
        const fresh = new Map(parentScope)
        for (let i = 0; i < count; i++) {
          const name = attrNames[i]!
          if (name !== 'xmlns' && !name.startsWith('xmlns:')) continue
          fresh.set(name === 'xmlns' ? '' : name.slice(6), attrValues[i]!)
        }
        scope = fresh
      }
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
          // Admit the frame record, copied attribute strings and child-map
          // slots before building any of them. Namespace storage was admitted
          // at the scope step above, so xmlns declarations are excluded here.
          if (!admit(192 + frameAttrBytes * 2 + 64)) {
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

      stack.push({ local, uri, ...(local === 'ext' ? { extUri: extUriOf() } : {}) })
      scopes.push(scope)
      if (selfClosing) events.close(qname)
    },

    close(_qname) {
      // Terminal state is inert: post-finish events must not build or charge.
      if (finished) return
      stack.pop()
      scopes.pop()
      if (rootKind === undefined || entries.length === 0) return
      const entry = entries.pop()!
      if (entry.frame !== null && !overCap) {
        // Admit the assembled node (object graph, child lists and wrapped
        // text nodes) before building it from the admitted frame.
        let childItems = 0
        for (const list of entry.frame.children.values()) childItems += list.length
        if (!admit(256 + childItems * 64 + entry.frame.texts.length * 16)) {
          // Withhold everything at finish; the partial node is dropped.
        } else {
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
      }
      if (entry.isRoot) rootKind = undefined
    },

    text(text, _more) {
      // Terminal state is inert: late text must not build or charge.
      if (finished) return
      void _more
      for (let i = entries.length - 1; i >= 0; i--) {
        const frame = entries[i]!.frame
        if (frame !== null) {
          if (frame.textLen + text.length > CF_LIMITS.textUnits) {
            overCap = true
          } else if (!admit(text.length * 2 + 64)) {
            // Admission failed (including the per-chunk slot): withhold everything at finish.
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
        // Truncated stream or structural overflow: withhold everything. The
        // transient capture hold bought the frames we are now discarding, so
        // drop it *before* finalizing: the account may be exhausted, and the
        // small fail-closed skip still needs room to be admitted.
        releaseHold()
        base.length = 0
        x14blocks.length = 0
        entries.length = 0
        scopes.length = 0
        stack.length = 0
        return finalizeCfSubtrees({ base: [], x14: [], sheetNs, truncated: true, skipped: [...skipped] }, styles, memory, meta)
      }
      return finalizeCfSubtrees({ base, x14: x14blocks, sheetNs, skipped: [...skipped] }, styles, memory, meta)
    } finally {
      // The finalizer admits its own normalized copies; drop the transient
      // hold and every captured reference so nothing stays alive past finish.
      releaseHold()
      base.length = 0
      x14blocks.length = 0
      entries.length = 0
      scopes.length = 0
      stack.length = 0
      skipped.clear()
      sheetNs = undefined
      nodeCount = 0
    }
  }

  return { events, finish }
}
