// Slideshow D4 capture plan (S1-21). Reads a corpus deck with the slideshow
// reader and scheduler and lists, per slide and step, what the owner captures
// in PowerPoint and when: the step's span in ms after its trigger, the five
// frame offsets (0/25/50/75/100 %) and the exact reference file names.
//
// The durations are the engine's reading of the file (for example `spd="med"`
// is 750 ms, a D4 question). The capture procedure keeps the raw video so a
// frame can be re-extracted if PowerPoint's real duration differs.
//
// Usage: bun scripts/slideshow-capture-plan.ts [corpus-dir] [--write]
//        bun scripts/slideshow-capture-plan.ts --deck <built-deck.pptx>
// Prints the generated deck tables; --write splices them into
// docs/validation/slideshow-captures.md between its capture-plan markers.
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { resolve as resolvePath } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { attrs, orderedChildren, parseXmlOrdered, type XmlNode } from '../src/core/xml'
import { rowForPreset, rowForTransition } from '../src/slideshow/catalogue'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { createScheduler } from '../src/slideshow/engine/scheduler'
import { readSlideTiming } from '../src/slideshow/read/timing'
import { shapeResolver } from '../src/slideshow/read/resolve'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import type { BehaviourNode, ContainerNode, ShowEvent, TargetResolver, TimeNode, TimingTree, TransitionSpec } from '../src/slideshow/types'
import { CAPTURE_PCTS, refFileName } from './slideshow-compare'

/** Every deck in docs/validation/slideshow-corpus.md, in its table order. */
export const CORPUS_DECKS = Object.freeze([
  'tdf104788', 'tdf115394', 'tdf142915', 'tdf44223', 'bnc591147', 'connector-shape-animations',
  'tdf104792-smart-art-animation', 'narration', 'airplane-transition', 'tdf156808', 'tdf149588_transparentSolidFill',
  'tdf105739', 'tdf146223', 'tdf123684', 'ooxtheme', 'tdf93868',
] as const)

export interface StepEffect { cls: string; id: number; subtype?: number; nodeType?: string; label: string; row?: string }
export interface CaptureStep {
  slide: number
  step: number
  /** What starts the step: the slide's arrival (step 0) or a click (step ≥ 1). */
  trigger: 'arrive' | 'click'
  /** Step 0 only: the incoming transition. */
  transition?: { element: string; ns: string; durationMs: number; spd?: string; dur?: string; row?: string }
  effects: StepEffect[]
  /** Ms from the trigger until the step settles; 'indefinite' or 'unresolved' when the engine cannot bound it. */
  spanMs: number | 'indefinite' | 'unresolved'
  /** Frame offsets in ms after the trigger, keyed by percent. */
  frames: Array<{ pct: number; ms: number; file: string }>
  /** Ms from the trigger until the first visual behaviour starts (an effect's start delay); omitted when 0. */
  leadMs?: number
}
export interface CaptureSlide {
  slide: number
  hidden: boolean
  advClick: boolean
  advTmMs?: number
  interactive: string[]
  steps: CaptureStep[]
}
export interface CapturePlan { deck: string; slideSize: { cx: number; cy: number }; slides: CaptureSlide[] }

const kids = (n: XmlNode | undefined) => orderedChildren(n).filter(([k]) => k !== '#text')
const child = (n: XmlNode | undefined, name: string) => kids(n).find(([k]) => k === name)?.[1]
const isContainer = (n: TimeNode): n is ContainerNode => n.kind === 'par' || n.kind === 'seq' || n.kind === 'excl'

function find(node: TimeNode, pred: (c: ContainerNode) => boolean, out: ContainerNode[] = []): ContainerNode[] {
  if (!isContainer(node)) return out
  if (pred(node)) out.push(node)
  for (const c of node.children) find(c, pred, out)
  return out
}

function effectsOf(group: ContainerNode): StepEffect[] {
  return find(group, c => c.ctn.presetClass !== undefined && c.ctn.presetId !== undefined).map(c => {
    const { presetClass: cls, presetId: id, presetSubtype: subtype, nodeType } = c.ctn as { presetClass: string; presetId: number; presetSubtype?: number; nodeType?: string }
    const row = rowForPreset(cls, id, subtype)
    return { cls, id, ...(subtype !== undefined ? { subtype } : {}), ...(nodeType ? { nodeType } : {}), label: row?.uiName ?? `${cls} ${id}`, ...(row ? { row: row.id } : {}) }
  })
}

/** A step that starts on its own (a group whose begin is not `indefinite`) plays at the slide's arrival. */
const waitsForClick = (group: ContainerNode) => group.ctn.start.some(c => c.delayMs === 'indefinite')

/** Spans shorter than three 60 fps video frames get only their 0 and 100 % frames. */
export const MIN_SPAN_FOR_QUARTERS_MS = 50

function frames(slide: number, step: number, span: CaptureStep['spanMs']): CaptureStep['frames'] {
  if (typeof span !== 'number') return [{ pct: 0, ms: 0, file: refFileName(slide, step, 0) }]
  const pcts = span === 0 ? [100] : span < MIN_SPAN_FOR_QUARTERS_MS ? [0, 100] : CAPTURE_PCTS
  return pcts.map(pct => ({ pct, ms: Math.round(span * pct / 100), file: refFileName(slide, step, pct) }))
}

const SPACING = 1_000_000
const LEAD_STEP_MS = 5, LEAD_MAX_MS = 10_000

const MEDIA = new Set(['audio', 'video', 'cmd'])
function behaviours(node: TimeNode, out: BehaviourNode[] = []): BehaviourNode[] {
  if (isContainer(node)) for (const c of node.children) behaviours(c, out)
  else out.push(node)
  return out
}

export function planSlide(slideXml: XmlNode, slideIndex: number, resolver: TargetResolver): CaptureSlide {
  const sink = createDiagnosticSink()
  const read = readSlideTiming(slideXml, slideIndex, resolver, sink)
  const tr: TransitionSpec | undefined = read.transition
  const slide = slideIndex + 1
  const hidden = attrs(slideXml).show === '0' || attrs(slideXml).show === 'false'
  const tree = read.timing && !('refused' in read.timing) ? read.timing as TimingTree : undefined
  const main = tree ? find(tree.root, c => c.ctn.nodeType === 'mainSeq')[0] : undefined
  const interactive = tree ? find(tree.root, c => c.ctn.nodeType === 'interactiveSeq').map(c => {
    const on = c.ctn.start.find(s => s.evt === 'onClick')
    return on?.target && 'spid' in on.target ? `shape id ${on.target.spid}` : 'a trigger shape'
  }) : []
  const groups = main ? main.children.filter(isContainer) : []
  const scheduler = tree ? createScheduler(tree, { sink, slideIndex }) : undefined
  const transMs = tr?.durationMs ?? 0
  const auto = groups.filter(g => !waitsForClick(g))
  const clicks = groups.filter(waitsForClick)
  // A step settles when its last visual behaviour ends; media and commands
  // (a sound held until onStopAudio, a media call) do not change the frame.
  const endOf = (g: ContainerNode, events: ShowEvent[]): number | 'indefinite' | undefined => {
    let end = 0
    for (const b of behaviours(g)) {
      if (MEDIA.has(b.kind)) continue
      const e = scheduler?.endTimeOf(b.ctn.id, events)
      if (e === undefined || e === 'indefinite') return e
      end = Math.max(end, e)
    }
    return end
  }
  // When the first visual behaviour of the groups starts, in ms after the
  // trigger at `t`: a step whose effects open with a delay shows nothing
  // until then. Sampled every LEAD_STEP_MS, up to the span (or LEAD_MAX_MS).
  const leadOf = (gs: ContainerNode[], events: ShowEvent[], t: number, span: CaptureStep['spanMs']): number => {
    if (!scheduler) return 0
    const visual = new Set(gs.flatMap(g => behaviours(g)).filter(b => !MEDIA.has(b.kind)))
    if (!visual.size) return 0
    const limit = typeof span === 'number' ? span : LEAD_MAX_MS
    for (let d = 0; d <= limit; d += LEAD_STEP_MS) {
      const snap = scheduler.evaluate(events, t + d)
      if (snap.active.some(e => visual.has(e.node)) || snap.frozen.some(e => visual.has(e.node))) return d
    }
    return 0
  }
  const lead = (ms: number) => ms > 0 ? { leadMs: ms } : {}

  // Step 0: the transition, then every group that starts without a click.
  let autoEnd: number | 'indefinite' | 'unresolved' = 0
  for (const g of auto) {
    const e = endOf(g, [])
    if (e === undefined) { autoEnd = 'unresolved'; break }
    if (e === 'indefinite') { autoEnd = 'indefinite'; break }
    if (typeof autoEnd === 'number') autoEnd = Math.max(autoEnd, e)
  }
  const span0 = typeof autoEnd === 'number' ? transMs + autoEnd : autoEnd
  const trRow = tr ? rowForTransition(tr.ns, tr.element, tr.preset) : undefined
  const steps: CaptureStep[] = [{
    slide, step: 0, trigger: 'arrive',
    ...(tr ? { transition: { element: tr.preset ? `${tr.element}:${tr.preset}` : tr.element, ns: tr.ns, durationMs: tr.durationMs, ...(trRow ? { row: trRow.id } : {}) } } : {}),
    effects: auto.flatMap(effectsOf), spanMs: span0, frames: frames(slide, 0, span0),
    ...lead(transMs > 0 ? 0 : leadOf(auto, [], 0, span0)),
  }]
  const events: ShowEvent[] = []
  clicks.forEach((g, i) => {
    const t = SPACING * (i + 1)
    events.push({ kind: 'next', t })
    const e = endOf(g, events.slice())
    const span: CaptureStep['spanMs'] = e === undefined ? 'unresolved' : e === 'indefinite' ? 'indefinite' : Math.max(0, e - t)
    steps.push({ slide, step: i + 1, trigger: 'click', effects: effectsOf(g), spanMs: span, frames: frames(slide, i + 1, span), ...lead(leadOf([g], events, t, span)) })
  })
  return { slide, hidden, advClick: tr?.advClick ?? true, ...(tr?.advTmMs !== undefined ? { advTmMs: tr.advTmMs } : {}), interactive, steps }
}

/** The `spd` and `p14:dur` attributes as written on the slide's first `p:transition` (the reader keeps only the resolved duration). */
function written(slideXml: XmlNode): { spd?: string; dur?: string } {
  let found: { spd?: string; dur?: string } | undefined
  const walk = (n: XmlNode) => {
    for (const [k, c] of kids(n)) {
      if (found) return
      if (k === 'transition') { const a = attrs(c); found = { ...(a.spd ? { spd: a.spd } : {}), ...(a.dur ? { dur: a.dur } : {}) }; return }
      walk(c)
    }
  }
  walk(slideXml)
  return found ?? {}
}

export async function planDeck(deck: string, bytes: Uint8Array): Promise<CapturePlan> {
  const zip = await JSZip.loadAsync(bytes)
  const text = async (p: string) => { const f = zip.file(p); if (!f) throw new Error(`${deck}: missing ${p}`); return f.async('string') }
  const pres = parseXmlOrdered(await text('ppt/presentation.xml'))
  const rels = parseXmlOrdered(await text('ppt/_rels/presentation.xml.rels'))
  const targets = new Map(kids(rels).map(([, r]) => [attrs(r).Id, attrs(r).Target]))
  const sz = attrs(child(pres, 'sldSz'))
  const doc = await parsePptx(await OfficePackage.load(bytes))
  const ids = kids(child(pres, 'sldIdLst')).map(([, s]) => attrs(s)['r:id'] ?? attrs(s).id)
  const slides: CaptureSlide[] = []
  for (const [i, rid] of ids.entries()) {
    const target = targets.get(rid)
    if (!target) throw new Error(`${deck}: slide relationship ${rid} missing`)
    const path = 'ppt/' + target.replace(/^\/?ppt\//, '').replace(/^\.\//, '')
    const xml = parseXmlOrdered(await text(path))
    const shapes = doc.slides[i]
    if (!shapes) throw new Error(`${deck}: slide ${i + 1} did not parse`)
    const s = planSlide(xml, i, shapeResolver(shapes))
    const t = s.steps[0].transition
    if (t) Object.assign(t, written(xml))
    slides.push(s)
  }
  return { deck, slideSize: { cx: Number(sz.cx), cy: Number(sz.cy) }, slides }
}

const msText = (s: CaptureStep['spanMs']) => typeof s === 'number' ? `${s}` : s

function effectText(s: CaptureStep): string {
  const parts: string[] = []
  if (s.transition) {
    const t = s.transition
    const how = [t.spd ? `spd=${t.spd}` : '', t.dur ? `p14:dur=${t.dur}` : ''].filter(Boolean).join(' ')
    parts.push(`transition ${t.ns}:${t.element}${how ? ` (${how})` : ''}, ${t.durationMs} ms`)
  } else if (s.step === 0) parts.push('no transition (cut)')
  for (const e of s.effects) parts.push(`${e.cls} ${e.id}${e.subtype !== undefined ? `/${e.subtype}` : ''} ${e.label}${e.nodeType && e.nodeType !== 'clickEffect' ? ` [${e.nodeType}]` : ''}`)
  return parts.join('; ').replace(/\|/g, '\\|')
}

/** The Markdown block for one deck, as embedded in the checklist between generated markers. */
export function renderPlan(plan: CapturePlan): string {
  const lines: string[] = []
  const ratio = plan.slideSize.cx / plan.slideSize.cy
  const shape = Math.abs(ratio - 16 / 9) < 0.01 ? '16:9' : Math.abs(ratio - 4 / 3) < 0.01 ? '4:3' : `${plan.slideSize.cx}×${plan.slideSize.cy} EMU`
  const total = plan.slides.reduce((n, s) => n + s.steps.reduce((m, st) => m + st.frames.length, 0), 0)
  lines.push(`Slide size ${shape}. ${plan.slides.length} slide(s), ${total} frame(s). Save into \`tests/fixtures/slideshow-refs/${plan.deck}/\`.`)
  lines.push('')
  lines.push('| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |')
  lines.push('|---|---|---|---|---|---|---|---|')
  for (const s of plan.slides) {
    for (const st of s.steps) {
      const notes: string[] = []
      if (st.step === 0) {
        if (s.hidden) notes.push('hidden slide: reach it with its number + Enter')
        if (!s.advClick) notes.push('advClick=0: advance with → or Space, not the mouse')
        if (s.advTmMs !== undefined) notes.push(`advTm=${s.advTmMs}: auto-advances`)
        if (s.interactive.length) notes.push(`trigger on ${s.interactive.join(', ')} (not a step)`)
      }
      if (typeof st.spanMs !== 'number') notes.push('span not bounded: save the 0 % frame and keep 5 s of video')
      else if (st.spanMs > 0 && st.spanMs < MIN_SPAN_FOR_QUARTERS_MS) notes.push(`under ${MIN_SPAN_FOR_QUARTERS_MS} ms: 0 and 100 % only`)
      const what = effectText(st) + (notes.length ? ` — ${notes.join('; ')}` : '')
      const rows = [...new Set([st.transition?.row, ...st.effects.map(e => e.row)].filter(Boolean))].map(r => `\`${r}\``).join('<br>') || '—'
      const shots = st.frames.map(f => `${f.ms} → \`${f.file}\``).join('<br>')
      lines.push(`| [ ] | ${st.slide} | ${st.step} | ${st.trigger} | ${what} | ${rows} | ${msText(st.spanMs)} | ${shots} |`)
    }
  }
  return lines.join('\n')
}

export const PLAN_BEGIN = (deck: string) => `<!-- capture-plan:${deck}:begin -->`
export const PLAN_END = (deck: string) => `<!-- capture-plan:${deck}:end -->`

/** Replace each deck's generated block in the checklist text; throws when a marker pair is missing. */
export function spliceChecklist(doc: string, blocks: ReadonlyMap<string, string>): string {
  let out = doc
  for (const [deck, body] of blocks) {
    const b = out.indexOf(PLAN_BEGIN(deck)), e = out.indexOf(PLAN_END(deck))
    if (b < 0 || e < b) throw new Error(`checklist has no ${PLAN_BEGIN(deck)} … ${PLAN_END(deck)} pair`)
    out = out.slice(0, b + PLAN_BEGIN(deck).length) + '\n' + body + '\n' + out.slice(e)
  }
  return out
}

export const CHECKLIST_PATH = 'docs/validation/slideshow-captures.md'

async function main(args: string[]) {
  const one = args.indexOf('--deck')
  if (one >= 0) {
    // A deck the owner built for a behavioural question: print its plan (with
    // the transition element PowerPoint wrote for each slide) under its file name.
    const file = args[one + 1]
    if (!file) throw new Error('--deck needs a .pptx path')
    const deck = file.replace(/^.*[\\/]/, '').replace(/\.pptx$/i, '')
    console.log(renderPlan(await planDeck(deck, new Uint8Array(readFileSync(file)))))
    return
  }
  const write = args.includes('--write')
  const dir = args.find(a => !a.startsWith('--')) ?? 'corpus'
  const blocks = new Map<string, string>()
  for (const deck of CORPUS_DECKS) {
    const file = resolvePath(dir, `lo-${deck}.pptx`)
    if (!existsSync(file)) throw new Error(`missing ${file}: run bun scripts/fetch-corpus.ts`)
    blocks.set(deck, renderPlan(await planDeck(deck, new Uint8Array(readFileSync(file)))))
  }
  if (write) {
    writeFileSync(CHECKLIST_PATH, spliceChecklist(readFileSync(CHECKLIST_PATH, 'utf8'), blocks))
    console.log(`updated ${CHECKLIST_PATH} (${blocks.size} decks)`)
  } else for (const [deck, body] of blocks) console.log(`${PLAN_BEGIN(deck)}\n${body}\n${PLAN_END(deck)}\n`)
}

if (process.argv[1] && resolvePath(process.argv[1]) === fileURLToPath(import.meta.url)) await main(process.argv.slice(2))
