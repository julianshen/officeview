/**
 * Corpus coverage (slideshow plan S1-22, owner rule): every effect except the
 * "never" rows renders as its closest equivalent. This classifies each
 * transition, effect and effect sound of a deck as
 *
 * - `native`: drawn as the file says (a stage-1 transition row; an effect
 *   whose file behaviours all render);
 * - `closest`: the closest equivalent plays, and a diagnostic says so (a 2-D
 *   look-alike or a fade stand-in for a transition; an effect with a dropped
 *   attribute, an unknown filter drawn as a fade, a preset rebuilt from the
 *   table, a whole-graphic build; media the run could not play, such as a
 *   codec jsdom cannot decode (the poster stays), or an effect sound on no
 *   shape, which the player does not follow);
 * - `refused`: nothing of it plays (a "never" row; every behaviour dropped;
 *   a target that is not on the slide).
 *
 * The classification reads the timing tree the engine reads and the
 * diagnostics a full play-through emitted (`playThrough`), so a stand-in that
 * was never diagnosed is reported, not assumed.
 */
import JSZip from 'jszip'
import { parseXmlOrdered, attrs, orderedChildren, type XmlNode } from '../../src/core/xml'
import { rowForPreset, rowForTransition } from '../../src/slideshow/catalogue'
import { createDiagnosticSink } from '../../src/slideshow/diagnostics'
import { readSlideTiming } from '../../src/slideshow/read/timing'
import { shapeResolver } from '../../src/slideshow/read/resolve'
import { resolveTransition2D } from '../../src/slideshow/render/transitions2d'
import type { PptxDocument } from '../../src/pptx/types'
import type { BehaviourNode, ContainerNode, SlideShowDiagnostic, TimeNode, TimingTree } from '../../src/slideshow/types'
import type { ShowRig } from './slideshow-show-rig'

export type Coverage = 'native' | 'closest' | 'refused'
export interface CoverageItem {
  /** 1-based. */
  slide: number
  kind: 'transition' | 'effect' | 'sound'
  what: string
  row?: string
  coverage: Coverage
  note: string
}

const kids = (n: XmlNode | undefined) => orderedChildren(n).filter(([k]) => k !== '#text')
const child = (n: XmlNode | undefined, name: string) => kids(n).find(([k]) => k === name)?.[1]
const isContainer = (n: TimeNode): n is ContainerNode => n.kind === 'par' || n.kind === 'seq' || n.kind === 'excl'
const MEDIA = new Set(['audio', 'video', 'cmd'])

function behavioursOf(node: TimeNode, out: BehaviourNode[] = []): BehaviourNode[] {
  if (isContainer(node)) for (const c of node.children) behavioursOf(c, out)
  else out.push(node)
  return out
}
function effectsOf(node: TimeNode, out: ContainerNode[] = []): ContainerNode[] {
  if (!isContainer(node)) return out
  if (node.ctn.presetClass !== undefined) { out.push(node); return out }
  for (const c of node.children) effectsOf(c, out)
  return out
}

/** The slide XML of each deck slide, in show order. */
export async function slideXmls(bytes: Uint8Array): Promise<XmlNode[]> {
  const zip = await JSZip.loadAsync(bytes)
  const text = async (p: string) => { const f = zip.file(p); if (!f) throw new Error(`missing ${p}`); return f.async('string') }
  const pres = parseXmlOrdered(await text('ppt/presentation.xml'))
  const rels = parseXmlOrdered(await text('ppt/_rels/presentation.xml.rels'))
  const targets = new Map(kids(rels).map(([, r]) => [attrs(r).Id, attrs(r).Target]))
  const out: XmlNode[] = []
  for (const [, s] of kids(child(pres, 'sldIdLst'))) {
    const target = targets.get(attrs(s)['r:id'] ?? attrs(s).id)
    if (!target) continue
    out.push(parseXmlOrdered(await text('ppt/' + target.replace(/^\/?ppt\//, '').replace(/^\.\//, ''))))
  }
  return out
}

/** Plays every slide: each slide's transition, every trigger once, every click step, until the end of the show. */
export async function playThrough(r: ShowRig, maxSteps = 200): Promise<void> {
  const triggered = new Set<string>()
  for (let i = 0; i < maxSteps && r.s.state.slide !== 'end' && !r.s.state.exited; i++) {
    await r.tick(5000)
    const slide = r.s.state.slide
    for (const b of [...r.host.querySelectorAll<HTMLElement>('[data-officeview-show-targets] button')]) {
      const t = b.dataset.target ?? ''
      if (!t.startsWith('trigger:') || triggered.has(`${slide}:${t}`)) continue
      triggered.add(`${slide}:${t}`)
      r.s.activate({ spid: t.slice('trigger:'.length), kind: 'trigger' })
      await r.tick(5000)
    }
    r.s.next()
    await r.tick(0)
  }
}

/** Classifies a deck's transitions, effects and effect sounds from its timing and the diagnostics of a play-through. */
export async function deckCoverage(bytes: Uint8Array, doc: PptxDocument, diagnostics: readonly SlideShowDiagnostic[]): Promise<CoverageItem[]> {
  const xmls = await slideXmls(bytes)
  const items: CoverageItem[] = []
  xmls.forEach((xml, i) => {
    const shapes = doc.slides[i]
    if (!shapes) return
    const read = readSlideTiming(xml, i, shapeResolver(shapes), createDiagnosticSink())
    const on = (code: string, effect?: string) => diagnostics.some(d => d.code === code && d.slideIndex === i && (effect === undefined || d.effect === effect))
    const slide = i + 1
    const tr = read.transition
    if (tr && tr.element !== 'none') {
      const name = `${tr.ns}:${tr.element}${tr.preset ? `:${tr.preset}` : ''}`
      const row = rowForTransition(tr.ns, tr.element, tr.preset)
      const drawn = resolveTransition2D(tr)
      if (row?.status === 'never') items.push({ slide, kind: 'transition', what: name, row: row.id, coverage: 'refused', note: 'never row' })
      else if (!row) items.push({ slide, kind: 'transition', what: name, coverage: 'closest', note: `uncatalogued: ${drawn.spec.element} stand-in${on('transition-unknown-preset') ? ', diagnosed' : ', NOT diagnosed'}` })
      else if (drawn.approximated) items.push({ slide, kind: 'transition', what: name, row: row.id, coverage: 'closest', note: `${row.status}: 2-D look-alike ${drawn.spec.element}${on('transition-approximate') ? ', diagnosed' : ', NOT diagnosed'}` })
      else items.push({ slide, kind: 'transition', what: name, row: row.id, coverage: 'native', note: `${row.status} (${tr.durationMs} ms)` })
    }
    if (tr?.sound?.start) {
      // S2-7's transition-sound controller plays it, or diagnoses why not (effect 'transition-sound').
      const told = diagnostics.filter(d => d.slideIndex === i && d.effect === 'transition-sound' && d.code.startsWith('media-')).map(d => d.code)
      items.push({ slide, kind: 'sound', what: 'transition sound', row: 'sound:transition', coverage: told.length ? 'closest' : 'native',
        note: told.length ? `not played in this run: silent (${[...new Set(told)].join(', ')}), diagnosed` : 'plays (S2-7)' })
    } else if (tr?.sound?.end) {
      items.push({ slide, kind: 'sound', what: 'transition endSnd', row: 'sound:transition', coverage: 'native', note: 'stops the playing sound (S2-7)' })
    }
    const tree = read.timing && !('refused' in read.timing) ? read.timing as TimingTree : undefined
    if (read.timing && 'refused' in read.timing) {
      items.push({ slide, kind: 'effect', what: 'whole timing tree', coverage: 'closest', note: `timing refused (${String((read.timing as { refused: unknown }).refused)}): the slide's static state shows, diagnosed` })
    }
    if (!tree) return
    for (const fx of effectsOf(tree.root)) {
      const { presetClass: cls, presetId: id, presetSubtype: sub } = fx.ctn
      const row = cls !== undefined && id !== undefined ? rowForPreset(cls, id, sub) : undefined
      const what = `${cls} ${id ?? '?'}${sub ? `/${sub}` : ''}${row ? ` ${row.uiName}` : ''}${fx.ctn.nodeType === 'interactiveSeq' ? '' : ''}`
      const all = behavioursOf(fx)
      const media = all.filter(b => MEDIA.has(b.kind))
      const visual = all.filter(b => !MEDIA.has(b.kind))
      const diagOf = (b: BehaviourNode) => diagnostics.filter(d => d.slideIndex === i && d.effect === `behaviour ${b.ctn.id}`)
      if (media.length) {
        // Shape media, effect sounds and commands are played by the media player (S2-7, S2-10), which diagnoses
        // on the target shape's spid, or on `behaviour <id>` for an effect sound (p:sndTgt: no shape)
        // (media-codec, media-refused, media-missing, media-error, media-autoplay, media-command).
        const parts = media.map(b => {
          const label = b.kind === 'cmd' ? `cmd ${(b as { cmd?: string }).cmd ?? ''}`.trim() : b.kind
          if (!('spid' in b.target)) {
            const codes = [...new Set(diagOf(b).filter(d => d.code.startsWith('media-')).map(d => d.code))]
            return codes.length
              ? { closest: true, row: 'sound:effect', note: `${label} effect sound: not played in this run (${codes.join(', ')}; the effect runs silent), diagnosed` }
              : { closest: false, row: 'sound:effect', note: `${label} effect sound: plays with its effect (S2-10)` }
          }
          const spid = b.target.spid
          const codes = [...new Set(diagnostics.filter(d => d.slideIndex === i && d.code.startsWith('media-') && (d.effect === spid || d.effect === String(b.ctn.id))).map(d => d.code))]
          return codes.length
            ? { closest: true, row: `behaviour:p:${b.kind}`, note: `${label}: not played in this run (${codes.join(', ')}; the poster stays), diagnosed` }
            : { closest: false, row: `behaviour:p:${b.kind}`, note: `${label}: plays (S2-7)` }
        })
        items.push({ slide, kind: cls === 'mediacall' ? 'effect' : 'sound', what: cls === 'mediacall' ? what : `${what}: ${media.map(b => b.kind).join('+')}`,
          row: parts[0].row, coverage: parts.some(p => p.closest) ? 'closest' : 'native', note: parts.map(p => p.note).join('; ') })
      }
      if (cls === 'mediacall' && !visual.length) continue
      if (!visual.length) {
        const rebuilt = diagnostics.some(d => d.code === 'preset-approximate' && d.effect === `preset ${cls}:${id}`)
        items.push({ slide, kind: 'effect', what, row: row?.id, coverage: 'closest', note: `no file behaviours: rebuilt from the preset table${rebuilt ? ', diagnosed' : ', NOT diagnosed'}` })
        continue
      }
      const notes: string[] = []
      let dropped = 0
      for (const b of visual) {
        const ds = diagOf(b)
        const unsupported = ds.filter(d => d.code === 'behaviour-unsupported')
        if (ds.some(d => d.code === 'anim-effect-unknown-filter')) notes.push(`filter ${(b as { filter?: string }).filter ?? '?'} drawn as a fade`)
        if (ds.some(d => d.code === 'anim-clr-unreadable')) notes.push('colour value unreadable')
        if (ds.some(d => d.code === 'behaviour-invalid-formula')) { notes.push('formula refused'); dropped++ }
        else if (unsupported.length) {
          notes.push(`${unsupported.map(d => d.message.replace(/^Attr (\S+).*$/, '$1')).join(', ')} dropped`)
          if ('attrNames' in b && Array.isArray(b.attrNames) && unsupported.length >= b.attrNames.length) dropped++
        }
        if ('spid' in b.target) {
          if (on('unresolved-target', b.target.spid) || on('master-target', b.target.spid)) { notes.push(`target ${b.target.spid} not on the slide`); dropped++ }
          if (b.target.sub === 'graphic' || b.target.sub === 'chart') notes.push(`${b.target.sub} build animates the whole ${b.target.sub}`)
        }
      }
      const rebuilt = diagnostics.some(d => d.code === 'preset-approximate' && d.slideIndex === i && d.effect === `preset ${cls}:${id}`)
      if (rebuilt) notes.push('preset keep-alike')
      const unique = [...new Set(notes)]
      const coverage: Coverage = dropped >= visual.length ? 'refused' : unique.length ? 'closest' : 'native'
      items.push({ slide, kind: 'effect', what, row: row?.id, coverage, note: unique.length ? unique.join('; ') : 'file behaviours' })
    }
  })
  return items
}

export function coverageCounts(items: readonly CoverageItem[]): Record<Coverage, number> {
  const out: Record<Coverage, number> = { native: 0, closest: 0, refused: 0 }
  for (const it of items) out[it.coverage]++
  return out
}

/** The Markdown table for the validation doc and the task report. */
export function coverageMarkdown(decks: ReadonlyArray<{ deck: string; items: readonly CoverageItem[] }>): string {
  const lines = ['| Deck | Native | Closest | Refused | Items (slide: what → coverage, note) |', '|---|---|---|---|---|']
  for (const { deck, items } of decks) {
    const c = coverageCounts(items)
    const list = items.map(it => `${it.slide}: ${it.what} → **${it.coverage}** (${it.note})`).join('<br>').replace(/\|/g, '\\|') || 'no transitions or effects'
    lines.push(`| ${deck} | ${c.native} | ${c.closest} | ${c.refused} | ${list} |`)
  }
  return lines.join('\n')
}
