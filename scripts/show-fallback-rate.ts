/**
 * Piece 4 Task 9 (spec §6 rule 7, §9 item 5): the show-plan fallback rate
 * per corpus deck — the "fallback rate is recorded" gate.
 *
 *   bun scripts/fetch-corpus.ts && bun scripts/show-fallback-rate.ts
 *
 * Every fetched corpus `.pptx` is opened in a server `DocumentHost` ('full'
 * policy) and every `anim` slide (timing or a transition) is built with
 * `showPlan` at scale 1, full motion, the fixed mark: the production path,
 * which runs `classifySlide` and then the build's own rules (a trigger that
 * moves geometry, the rest rule, mark validation, the limits and budget).
 *
 * Per deck it writes the brief's fields:
 *   { deck, slides, steps, composited, rate, reasons{} }
 * - `slides`: the deck's shown (non-hidden) slides;
 * - `steps`: click steps over its `anim` slides; `composited`: those that
 *   play as keyframes or as the still; `rate` = composited / steps (null
 *   when there is no step);
 * - `reasons`: composited steps by reason (a refused slide's steps count
 *   under `still:<code>`);
 * and, additively: `anim` (slides asked for a plan), `stills` (slides the
 * server refused: they play the piece-3 still), `stillReasons` (refusal
 * code and its diagnostics), `error` (a deck that could not be opened).
 * Output: `docs/validation/streaming-s3-p4-rates.json`.
 */
import { readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import './stream-s0/napi-preload'

const { DocumentHost } = await import('../server/src/render')
const { loadShowDeck } = await import('../src/testdata/show-decks')
const { createSlideTimeline } = await import('../src/slideshow/engine/steps')

const ROOT = resolve(import.meta.dir, '..')
const CORPUS = resolve(ROOT, 'corpus')
const OUT = resolve(ROOT, 'docs/validation/streaming-s3-p4-rates.json')

interface DeckRate {
  deck: string
  slides: number
  steps: number
  composited: number
  rate: number | null
  reasons: Record<string, number>
  anim: number
  stills: number
  stillReasons: Record<string, number>
  error?: string
}

const bump = (o: Record<string, number>, k: string, n = 1) => { o[k] = (o[k] ?? 0) + n }

/** Click steps of slide `i` (for a refused slide, whose refusal carries none). */
async function stepCounter(bytes: Uint8Array): Promise<(i: number) => Promise<number>> {
  try {
    const d = await loadShowDeck(bytes)
    return async (i) => { try { return createSlideTimeline((await d.meta(i)).timing, { showAnimation: true, reducedMotion: false }).steps } catch { return 0 } }
  } catch { return async () => 0 }
}

const decks = readdirSync(CORPUS).filter((f) => f.endsWith('.pptx')).sort()
if (decks.length === 0) throw new Error('no corpus decks: run `bun scripts/fetch-corpus.ts` first')
const out: DeckRate[] = []
for (const name of decks) {
  const r: DeckRate = { deck: name, slides: 0, steps: 0, composited: 0, rate: null, reasons: {}, anim: 0, stills: 0, stillReasons: {} }
  const host = new DocumentHost()
  const t0 = performance.now()
  try {
    const bytes = new Uint8Array(readFileSync(resolve(CORPUS, name)))
    const opened = await host.open(name, bytes, 'full')
    r.slides = opened.units.length - opened.hidden.length
    r.anim = opened.anim.length
    const stepsOf = await stepCounter(bytes)
    for (const unitIndex of opened.anim) {
      let built
      try {
        built = await host.showPlan(opened.docKey, { unitIndex, scale: 1, motion: 'full', mark: { kind: 'fixed' }, strictMarks: false })
      } catch (e) {
        built = { refused: `error:${e instanceof Error ? e.message.slice(0, 60) : String(e)}`, diagnostics: [] }
      }
      if ('refused' in built) {
        r.stills++
        bump(r.stillReasons, built.refused)
        for (const d of new Set(built.diagnostics.map((x) => x.code))) bump(r.stillReasons, `${built.refused}/${d}`)
        const n = await stepsOf(unitIndex)
        r.steps += n
        r.composited += n
        if (n) bump(r.reasons, `still:${built.refused}`, n)
        continue
      }
      for (const s of built.plan.steps) {
        r.steps++
        if (s.eligibility !== 'composited') continue
        r.composited++
        bump(r.reasons, s.reason ?? 'unspecified')
      }
    }
    host.drop(opened.docKey)
  } catch (e) {
    r.error = e instanceof Error ? e.message.slice(0, 120) : String(e)
  }
  r.rate = r.steps ? Math.round((r.composited / r.steps) * 1e4) / 1e4 : null
  out.push(r)
  console.log(`${name}: slides ${r.slides}, anim ${r.anim}, stills ${r.stills}, steps ${r.steps}, composited ${r.composited}, rate ${r.rate} (${Math.round(performance.now() - t0)} ms)${r.error ? ` error ${r.error}` : ''}`)
}

const total = out.reduce((t, r) => ({ slides: t.slides + r.slides, anim: t.anim + r.anim, stills: t.stills + r.stills, steps: t.steps + r.steps, composited: t.composited + r.composited }), { slides: 0, anim: 0, stills: 0, steps: 0, composited: 0 })
const reasons: Record<string, number> = {}, stillReasons: Record<string, number> = {}
for (const r of out) { for (const [k, n] of Object.entries(r.reasons)) bump(reasons, k, n); for (const [k, n] of Object.entries(r.stillReasons)) bump(stillReasons, k, n) }
const doc = {
  generated: 'bun scripts/show-fallback-rate.ts',
  build: { scale: 1, motion: 'full', mark: 'fixed', strictMarks: false, contentPolicy: 'full' },
  total: { ...total, rate: total.steps ? Math.round((total.composited / total.steps) * 1e4) / 1e4 : null, reasons, stillReasons },
  decks: out,
}
writeFileSync(OUT, `${JSON.stringify(doc, null, 2)}\n`)
console.log(`wrote ${OUT}: ${JSON.stringify(doc.total)}`)
process.exit(0)
