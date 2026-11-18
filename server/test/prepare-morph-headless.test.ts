/**
 * Piece 4 Task 2: the DOM-free Morph preparation (`planMorph`) runs on the
 * server (bun, no DOM) under `@napi-rs/canvas` (`serverBackend`).
 */
import { afterAll, beforeAll, describe, expect, it } from 'bun:test'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { purityViolations } from '../../tests/helpers/core-gates'
import { sourceClosure } from './source-graph'
import { installCanvasBackend, type CanvasBackend } from '../../src/core/canvas-backend'
import { prepareSlide, type PreparedSlide } from '../../src/slideshow/prepare'
import { planMorph, type MorphEnv } from '../../src/slideshow/prepare/morph'
import { createDiagnosticSink } from '../../src/slideshow/diagnostics'
import type { CanvasFactory, DiagnosticSink, Lease, RenderState, SlideShowWorkspace } from '../../src/slideshow/types'
import { loadShowDeck, manyLayersDeck, morphDeck } from '../../src/testdata/show-decks'
import { serverBackend } from '../src/backend'

let backend: CanvasBackend
let previous: CanvasBackend | undefined
beforeAll(() => { backend = serverBackend(); previous = installCanvasBackend(backend) })
afterAll(() => { installCanvasBackend(previous) })

function roomyWorkspace(): SlideShowWorkspace {
  return {
    limit: Number.POSITIVE_INFINITY, used: () => 0, reserveTerminal: () => undefined,
    lease: (bytes, tag) => {
      const lease: Lease = { tag, bytes, grow: () => true, shrink() {}, transfer: () => lease, release() {} }
      return lease
    },
  }
}

const factory = (): CanvasFactory => (w, h) => backend.createCanvas(Math.max(1, Math.ceil(w)), Math.max(1, Math.ceil(h))) as never
const EMPTY: RenderState = new Map()

async function slides(bytes: Uint8Array): Promise<{ from: PreparedSlide; to: PreparedSlide }> {
  const d = await loadShowDeck(bytes)
  const prep = async (i: number) => prepareSlide({ index: i, slide: d.doc.slides[i], meta: await d.meta(i), watermark: false })
  return { from: await prep(0), to: await prep(1) }
}

function env(sink: DiagnosticSink): MorphEnv {
  return { sink, workspace: roomyWorkspace(), canvas: factory(), resolveFont: f => f, maxLayers: 256, scale: 1, watermark: false }
}

const codes = (sink: DiagnosticSink) => sink.list().map(d => d.code)

describe('prepare morph headless', () => {
  it("planMorph pairs morphDeck's objects under bun", async () => {
    expect(typeof document).toBe('undefined')
    const { from, to } = await slides(await morphDeck())
    const sink = createDiagnosticSink()
    const plan = planMorph({ ps: from, state: EMPTY }, to, EMPTY, to.meta.transition!, env(sink))
    if (typeof plan === 'string') throw new Error(plan)
    expect(plan.pairs).toEqual([{ fromKey: 'sp:2', toKey: 'sp:2', rule: 'name-kind' }])
    expect(plan.items.length).toBeGreaterThan(0)
    for (const item of plan.items) {
      for (const s of [item.from, item.to]) {
        if (!s) continue
        expect(Number.isInteger(s.ref)).toBe(true)
        expect(plan.refs[s.ref]).toBeDefined()
      }
    }
    expect(plan.kinds).toContain('morph:pair:sp:2')
  })

  it('over-bound unit split coarsens', async () => {
    // 128 words a side (the most `textItems` splits one shape into, `MAX_TEXT_BOXES`; past it the words coarsen to
    // paragraphs first, `text-build-coarsened`) split into ~130 layers each: more than the 256-layer Morph bound.
    const { from, to } = await slides(await morphDeck({ option: 'byWord', words: 128 }))
    const sink = createDiagnosticSink()
    const plan = planMorph({ ps: from, state: EMPTY }, to, EMPTY, to.meta.transition!, env(sink))
    expect(typeof plan).not.toBe('string')
    expect(codes(sink)).toContain('morph-units-coarsened')
    if (typeof plan === 'string') return
    // Coarsened: the box pairs whole, no word layer of its own.
    expect(plan.kinds).toContain('morph:pair:sp:2')
    expect(plan.kinds.some(k => k.startsWith('morph:pair:sp:2/'))).toBe(false)
    expect(plan.refs.length).toBeLessThanOrEqual(256)
  })

  it('over 256 objects refuses', async () => {
    const { from, to } = await slides(await manyLayersDeck(300))
    const sink = createDiagnosticSink()
    const plan = planMorph({ ps: from, state: EMPTY }, to, EMPTY, to.meta.transition!, env(sink))
    expect(typeof plan).toBe('string')
    expect(codes(sink)).toContain('morph-refused')
  })

  it('prepare/morph.ts reaches no stage, compositor, React or DOM', () => {
    const root = resolve(import.meta.dir, '../..')
    const { files, packages } = sourceClosure(resolve(root, 'src/slideshow/prepare/morph.ts'), root)
    expect(files.filter(f => f.endsWith('.tsx') || f === 'src/slideshow/player/stage.ts' || f.includes('compositor'))).toEqual([])
    expect(packages.filter(p => /^react(-dom)?(\/|$)/.test(p))).toEqual([])
    // The shared core (`canvas-backend`, `text-metrics`) guards its own browser fallback; the slideshow modules have none.
    const offenders = files.filter(f => f.startsWith('src/slideshow/')).flatMap(f => purityViolations(readFileSync(resolve(root, f), 'utf8')).map(v => `${f}: ${v}`))
    expect(offenders).toEqual([])
  })
})
