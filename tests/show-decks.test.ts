/**
 * Piece 4 D18: the shared show fixture decks read as intended. Every later
 * task builds on them, so each is checked here: its timing reads (never
 * refused), its transition and links are read, and every adversary deck's
 * hidden ink (#13F00D) never reaches the composed local show's frames
 * (`moverReveal` aside, whose ink shows only strictly between the
 * classifier's samples and the 15 fps frame times).
 */
import { afterEach, describe, expect, it } from 'vitest'
import {
  ADVERSARY_KINDS, HIDDEN_INK, MOVER_REVEAL, adversaryDeck, cubeDeck, flyInDeck, linkDeck, loadShowDeck, longStepDeck, manyLayersDeck, morphDeck, triggerDeck,
} from '../src/testdata/show-decks'
import { createSlideTimeline } from '../src/slideshow/engine/steps'
import type { TimeNode } from '../src/slideshow/types'
import { installShowSeams, loadDeck, startShow, type ShowRig, type ShowSeams } from './helpers/slideshow-show-rig'

const behaviours = (n: TimeNode, out: string[] = []): string[] => {
  if ('children' in n) n.children.forEach(c => behaviours(c, out))
  else out.push(`${n.kind}${n.attrNames.length ? `[${n.attrNames.join(',')}]` : ''}`)
  return out
}

async function timingOf(bytes: Uint8Array, i = 0) {
  const d = await loadShowDeck(bytes)
  const meta = await d.meta(i)
  const t = meta.timing
  if (!t || 'refused' in t) throw new Error(`slide ${i}: no timing (${t && 'refused' in t ? t.refused : 'none'})`)
  return { d, meta, root: t.root }
}

describe('show fixture decks', () => {
  it('flyInDeck: one click flying box 3 in', async () => {
    const { root } = await timingOf(await flyInDeck())
    expect(behaviours(root)).toEqual(['set[style.visibility]', 'anim[ppt_x]', 'anim[ppt_y]'])
  })

  it('morphDeck: shape 2 "Box" on both slides, moved and resized; the second enters by Morph', async () => {
    for (const option of ['byObject', 'byWord', 'byChar'] as const) {
      const d = await loadShowDeck(await morphDeck({ option, words: 3 }))
      const [a, b] = d.doc.slides.map(s => s.shapes)
      expect(a.map(s => [s.source?.id, s.source?.name])).toEqual([['2', 'Box']])
      expect(b.map(s => [s.source?.id, s.source?.name])).toEqual([['2', 'Box']])
      expect([a[0].xEmu, a[0].widthEmu]).not.toEqual([b[0].xEmu, b[0].widthEmu])
      const t = (await d.meta(1)).transition
      expect(t).toMatchObject({ ns: 'p159', element: 'morph', attrs: { option } })
      expect(a[0].textBody?.paragraphs[0].runs.map(r => r.text).join('')).toBe('w0 w1 w2')
    }
  })

  it('cubeDeck: the second slide enters by a 3-D prism (Cube)', async () => {
    const d = await loadShowDeck(await cubeDeck())
    expect((await d.meta(1)).transition).toMatchObject({ ns: 'p14', element: 'prism' })
  })

  it('triggerDeck: an interactive sequence on button 3 fades box 2 in', async () => {
    const { root } = await timingOf(await triggerDeck())
    const seq = (root as Extract<TimeNode, { children: TimeNode[] }>).children[0]
    expect(seq.ctn.nodeType).toBe('interactiveSeq')
    expect(seq.ctn.start[0]).toMatchObject({ evt: 'onClick', target: { spid: '3' } })
    expect(behaviours(root)).toEqual(['set[style.visibility]', 'animEffect'])
  })

  it('linkDeck: a slide jump, a URL and a run link', async () => {
    const d = await loadShowDeck(await linkDeck())
    expect((await d.meta(0)).links).toHaveLength(3)
  })

  it('longStepDeck: a 9 s spin', async () => {
    const { root } = await timingOf(await longStepDeck())
    expect(behaviours(root)).toEqual(['animRot[r]'])
  })

  it('manyLayersDeck: n boxes faded in over 5 steps, then a Morph of n objects', async () => {
    const { d, root } = await timingOf(await manyLayersDeck(12))
    expect(d.doc.slides[0].shapes).toHaveLength(12)
    expect(d.doc.slides[1].shapes).toHaveLength(12)
    const main = (root as Extract<TimeNode, { children: TimeNode[] }>).children[0] as Extract<TimeNode, { children: TimeNode[] }>
    expect(main.children).toHaveLength(5)
    expect(behaviours(root).filter(b => b === 'animEffect')).toHaveLength(12)
    expect((await d.meta(1)).transition).toMatchObject({ ns: 'p159', element: 'morph' })
  })

  it.each(ADVERSARY_KINDS)('adversaryDeck(%s): timing reads', async (kind) => {
    const { root } = await timingOf(await adversaryDeck(kind))
    expect(behaviours(root).length).toBeGreaterThan(0)
  })
})

describe('timeline keyTimes (piece 4 §6 rule 1)', () => {
  it('moverReveal: behaviour begin/end and tav times, never a motion-path vertex', async () => {
    const { meta } = await timingOf(await adversaryDeck('moverReveal'))
    const tl = createSlideTimeline(meta.timing, { showAnimation: true, reducedMotion: false })
    const c = tl.next(tl.checkpoint(0), 0)
    const { stepMs, dartAt, dartMs } = MOVER_REVEAL
    expect(tl.keyTimes(c, 0, stepMs)).toEqual([0, dartAt, dartAt + dartMs, stepMs])
    expect(tl.keyTimes(c, 1, 999)).toEqual([dartAt, dartAt + dartMs])
    expect(tl.keyTimes(tl.initial(), 0, stepMs)).toEqual([])
    expect(createSlideTimeline(undefined, { showAnimation: true, reducedMotion: false }).keyTimes(c, 0, stepMs)).toEqual([])
  })
})

describe('adversary decks hide their ink in the local show', () => {
  let seams: ShowSeams | undefined
  let rig: ShowRig | undefined
  afterEach(() => { rig?.exit(); rig = undefined; seams?.restore(); seams = undefined; document.body.replaceChildren() })

  const ink = [parseInt(HIDDEN_INK.slice(0, 2), 16), parseInt(HIDDEN_INK.slice(2, 4), 16), parseInt(HIDDEN_INK.slice(4, 6), 16)]
  /** Pixels of the frame within ΔRGB ≤ 6 of the hidden ink, with α > 0. */
  const inkPixels = (r: ShowRig) => {
    const c = r.frameCanvas()!
    const data = c.getContext('2d')!.getImageData(0, 0, c.width, c.height).data
    let n = 0
    for (let i = 0; i < data.length; i += 4) {
      if (data[i + 3] > 0 && Math.abs(data[i] - ink[0]) <= 6 && Math.abs(data[i + 1] - ink[1]) <= 6 && Math.abs(data[i + 2] - ink[2]) <= 6) n++
    }
    return n
  }
  /** Ink pixel counts before the click and every 50 ms for 1 s after it. */
  async function trace(kind: (typeof ADVERSARY_KINDS)[number]) {
    seams = installShowSeams({ width: 960, height: 720 })
    rig = await startShow(await loadDeck(await adversaryDeck(kind)), { seams, size: { width: 960, height: 720 } })
    await rig.tick(0)
    const counts = [inkPixels(rig)]
    await rig.key('ArrowRight')
    for (let t = 0; t <= 1000; t += 50) { counts.push(inkPixels(rig)); await rig.tick(50) }
    counts.push(inkPixels(rig))
    return counts
  }

  it.each(ADVERSARY_KINDS.filter(k => k !== 'moverReveal'))('%s: no frame shows the ink', async (kind) => {
    expect(await trace(kind)).toEqual(Array(23).fill(0))
  }, 30_000)

  it('moverReveal: the ink is invisible at every classifier sample and 15 fps frame, visible only inside the dart window', async () => {
    const { stepMs, dartAt, dartMs, reveal } = MOVER_REVEAL
    // 8 even samples, the keyframe and behaviour begin/end times, and every 15 fps frame of the step.
    const covered = [...Array.from({ length: 8 }, (_, i) => (i * stepMs) / 7), 0, dartAt, dartAt + dartMs, stepMs, ...Array.from({ length: 16 }, (_, i) => (i * 1000) / 15)]
    const inside = [reveal[0] + 0.5, dartAt + dartMs / 2, reveal[1] - 0.5]
    expect(covered.every(t => t <= reveal[0] || t >= reveal[1])).toBe(true)
    seams = installShowSeams({ width: 960, height: 720 })
    rig = await startShow(await loadDeck(await adversaryDeck('moverReveal')), { seams, size: { width: 960, height: 720 } })
    await rig.tick(0)
    expect(inkPixels(rig)).toBe(0)
    await rig.key('ArrowRight')
    const at = [...covered.map(t => ({ t, shown: false })), ...inside.map(t => ({ t, shown: true }))].sort((a, b) => a.t - b.t)
    let now = 0
    const seen: Array<{ t: number; ink: number }> = []
    for (const { t } of at) { await rig.tick(t - now); now = t; seen.push({ t, ink: inkPixels(rig) }) }
    for (const [i, { shown }] of at.entries()) {
      if (shown) expect(seen[i].ink, `t=${seen[i].t}`).toBeGreaterThan(0)
      else expect(seen[i].ink, `t=${seen[i].t}`).toBe(0)
    }
  }, 30_000)
})
