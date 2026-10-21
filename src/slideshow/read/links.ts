import { attrs, getChild, getChildren, elementChildren, type XmlNode } from '../../core/xml'
import type { DiagnosticSink, LinkSpec } from '../types'
import type { Relationship } from './deck'

const jumps = new Set(['nextslide', 'previousslide', 'firstslide', 'lastslide', 'endshow'])

/**
 * `hlinkClick` on shapes, groups and text runs. Unsupported actions (custom shows,
 * lastslideviewed, macros), unsafe schemes and links that LinkSpec cannot address
 * (table cells have no run address) are explicit refusals, never silent drops.
 * Actions: MS-OE376 2.1.1315. `run.r` counts `r`, `fld` and `br` children in order,
 * the same sequence as `PptxParagraph.runs` (src/drawing/text-parse.ts).
 */
export function readLinks(xml: XmlNode, rels: readonly Relationship[], parts: readonly string[], index: number, sink: DiagnosticSink): LinkSpec[] {
  const links: LinkSpec[] = []

  function action(n: XmlNode): LinkSpec['action'] {
    const a = attrs(n)
    const r = rels.find(r => r.id === a.id)
    const verb = a.action ?? ''
    if (verb.startsWith('ppaction://hlinkshowjump?jump=')) {
      const v = verb.slice('ppaction://hlinkshowjump?jump='.length)
      if (jumps.has(v)) return { kind: 'jump', verb: v as 'nextslide' }
    }
    if (verb === 'ppaction://hlinksldjump' || !verb) {
      if (r && !r.external && r.type.endsWith('/slide')) {
        const i = parts.indexOf(r.target)
        if (i >= 0) return { kind: 'slide', index: i }
      }
    }
    if (!verb && r?.external) {
      try {
        const u = new URL(r.target)
        if (['http:', 'https:', 'mailto:'].includes(u.protocol)) return { kind: 'url', href: r.target }
      } catch {
        // Falls through to the refusal below.
      }
    }
    return { kind: 'refused', reason: verb || 'unsafe-or-missing-target' }
  }

  function push(spid: string, a: LinkSpec['action'], run?: { p: number; r: number }) {
    links.push({ spid, ...(run ? { run } : {}), action: a })
    if (a.kind === 'refused') sink.add({ code: 'link-refused', slideIndex: index, effect: `${spid}:${run?.p ?? ''}:${run?.r ?? ''}`, message: a.reason })
  }

  function add(n: XmlNode, spid: string, run?: { p: number; r: number }) {
    push(spid, action(n), run)
  }

  /** The cells of `a:tbl` carry text bodies, but a LinkSpec has no cell address: refuse them visibly. */
  function tableLinks(frame: XmlNode, spid: string) {
    const table = getChild(getChild(getChild(frame, 'graphic'), 'graphicData'), 'tbl')
    for (const tr of getChildren(table, 'tr'))
      for (const tc of getChildren(tr, 'tc'))
        for (const p of getChildren(getChild(tc, 'txBody'), 'p'))
          for (const [k, r] of elementChildren(p))
            if (k === 'r' || k === 'fld')
              for (const _ of getChildren(getChild(r, 'rPr'), 'hlinkClick')) push(spid, { kind: 'refused', reason: 'table-cell' })
  }

  function walk(n: XmlNode) {
    for (const [kind, shape] of elementChildren(n)) {
      if (kind === 'grpSp') {
        const own = getChild(getChild(getChild(shape, 'nvGrpSpPr'), 'cNvPr'), 'hlinkClick')
        const groupId = attrs(getChild(getChild(shape, 'nvGrpSpPr'), 'cNvPr')).id
        if (own && groupId) add(own, groupId)
        walk(shape)
        continue
      }
      if (kind === 'AlternateContent') {
        // One branch is rendered: the first Choice, else the Fallback. Walking both would repeat every shape.
        const branch = getChild(shape, 'Choice') ?? getChild(shape, 'Fallback')
        if (branch) walk(branch)
        continue
      }
      const nv = getChild(shape, 'nvSpPr') ?? getChild(shape, 'nvPicPr') ?? getChild(shape, 'nvGraphicFramePr') ?? getChild(shape, 'nvCxnSpPr')
      const cn = getChild(nv, 'cNvPr')
      const spid = attrs(cn).id
      if (!spid) continue
      for (const link of getChildren(cn, 'hlinkClick')) add(link, spid)
      if (kind === 'graphicFrame') tableLinks(shape, spid)
      getChildren(getChild(shape, 'txBody'), 'p').forEach((p, pi) => {
        let ri = 0
        for (const [k, r] of elementChildren(p)) {
          if (k === 'r' || k === 'fld' || k === 'br') {
            for (const link of getChildren(getChild(r, 'rPr'), 'hlinkClick')) add(link, spid, { p: pi, r: ri })
            ri++
          }
        }
      })
    }
  }

  walk(getChild(getChild(xml, 'cSld'), 'spTree') ?? {})
  return links
}
