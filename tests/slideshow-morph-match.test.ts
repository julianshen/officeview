import { describe, expect, test } from 'vitest'
import type { PptxShape, PptxSlide } from '../src/pptx/types'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { matchMorph } from '../src/slideshow/morph/match'

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------
const PX = 9525
/** EMU coordinates of the 960x720 slide. The diagonal is 1200px, so the
 * geometry threshold is 300px; the inclusive-edge test puts centres exactly
 * 300 px apart. */
const W = 960 * PX
const H = 720 * PX

function text(text: string) {
  return {
    anchor: 't' as const, wrap: true, insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0,
    paragraphs: [{ align: 'left' as const, level: 0, runs: [{ text, fontSizePt: 12 }] }],
  }
}

interface ObjSpec {
  x?: number; y?: number; w?: number; h?: number
  name?: string
  text?: string
  image?: boolean
  geometry?: 'rect' | 'ellipse' | 'roundRect'
  children?: boolean
  sourceless?: boolean
}

function obj(id: number | undefined, over: ObjSpec = {}): PptxShape {
  const shape: PptxShape = {
    xEmu: (over.x ?? 0) * PX, yEmu: (over.y ?? 0) * PX,
    widthEmu: (over.w ?? 100) * PX, heightEmu: (over.h ?? 50) * PX,
    geometry: over.geometry ?? 'rect',
  }
  if (over.sourceless) shape.source = { partPath: 'ppt/slides/slide1.xml', treePath: 'spTree/x', element: 'sp', representation: 'native' }
  else shape.source = { partPath: 'ppt/slides/slide1.xml', treePath: `spTree/sp[${id}]`, element: 'sp', id: String(id), representation: 'native' }
  if (over.name !== undefined && shape.source) shape.source.name = over.name
  if (over.text !== undefined) shape.textBody = text(over.text)
  if (over.image) shape.image = { data: new Uint8Array(4), width: 10, height: 10 } as PptxShape['image']
  if (over.children) {
    shape.group = { off: { x: 0, y: 0 }, ext: { width: 100 * PX, height: 50 * PX }, chOff: { x: 0, y: 0 }, chExt: { width: 100 * PX, height: 50 * PX } }
    shape.children = [obj(undefined, { geometry: 'ellipse' })]
  }
  return shape
}

function slideOf(shapes: PptxShape[], index = 0): PptxSlide {
  return { index, widthEmu: W, heightEmu: H, shapes, background: '#f0f0f0' }
}

// ---------------------------------------------------------------------------
// Rule 1: !! names
// ---------------------------------------------------------------------------
describe('morph matching: !! names', () => {
  test('!! names pair first even across kinds', () => {
    const from = slideOf([obj(1, { name: '!!X', geometry: 'rect' })])
    const to = slideOf([obj(9, { name: '!!X', image: true })])
    expect(matchMorph(from, to, 'byObject')).toEqual({
      pairs: [{ from: 'sp:1', to: 'sp:9', rule: '!!' }],
      fadeOut: [],
      fadeIn: [],
    })
  })

  test('!! objects never pair with differently named or untagged objects', () => {
    // sp:1 is !!-tagged; sp:2 shares its text and sp:3 its geometry and kind,
    // and both would be legal matches for an untagged object. Morph must not
    // pair a !! object with a non-!! object (published rule), so it fades out.
    const from = slideOf([
      obj(1, { name: '!!X', text: 'hello', geometry: 'rect' }),
      obj(4, { name: 'Y', text: 'hello' }),
    ])
    const to = slideOf([
      obj(2, { name: 'R', text: 'hello' }),
      obj(3, { name: 'S', geometry: 'rect' }),
      obj(5, { name: 'Z', text: 'hello' }),
    ])
    const match = matchMorph(from, to, 'byObject')
    // Equal text: reading order pairs sp:4 with the first candidate, sp:2.
    expect(match.pairs).toEqual([
      { from: 'sp:4', to: 'sp:2', rule: 'text' },
    ])
    expect(match.fadeOut).toEqual(['sp:1'])
    expect(match.fadeIn).toEqual(['sp:3', 'sp:5'])
  })

  test('!! names pair before name-kind matches in the same slide', () => {
    const from = slideOf([
      obj(1, { name: '!!A' }),
      obj(2, { name: 'A' }),
    ])
    const to = slideOf([
      obj(3, { name: 'A' }),
      obj(4, { name: '!!A' }),
    ])
    expect(matchMorph(from, to, 'byObject').pairs).toEqual([
      { from: 'sp:1', to: 'sp:4', rule: '!!' },
      { from: 'sp:2', to: 'sp:3', rule: 'name-kind' },
    ])
  })
})

// ---------------------------------------------------------------------------
// Rule order: name-kind, then text, then geometry
// ---------------------------------------------------------------------------
describe('morph matching: rule order', () => {
  test('name and kind beat text; text beats geometry', () => {
    // Name+kind beats text: sp:1 (name B) pairs with sp:2 (name B) rather than
    // sp:3, which shares its text.
    const a = matchMorph(
      slideOf([obj(1, { name: 'B', text: 'hi' })]),
      slideOf([obj(2, { name: 'B', text: 'yo' }), obj(3, { name: 'C', text: 'hi' })]),
      'byObject',
    )
    expect(a.pairs).toEqual([{ from: 'sp:1', to: 'sp:2', rule: 'name-kind' }])
    expect(a.fadeIn).toEqual(['sp:3'])

    // Text beats geometry: same-geometry objects sit at swapped positions, so
    // a geometry pass would pair sp:1->sp:3 and sp:2->sp:4; the text rule runs
    // first and pairs by content instead.
    const b = matchMorph(
      slideOf([
        obj(1, { text: 'same', geometry: 'rect', x: 0, w: 400, h: 200 }),
        obj(2, { text: 'other', geometry: 'rect', x: 6000, w: 400, h: 200 }),
      ]),
      slideOf([
        obj(3, { text: 'other', geometry: 'rect', x: 0, w: 400, h: 200 }),
        obj(4, { text: 'same', geometry: 'rect', x: 6000, w: 400, h: 200 }),
      ]),
      'byObject',
    )
    expect(b.pairs).toEqual([
      { from: 'sp:1', to: 'sp:4', rule: 'text' },
      { from: 'sp:2', to: 'sp:3', rule: 'text' },
    ])
  })

  test('geometry pairs same preset shapes within a quarter of the slide diagonal', () => {
    // No names, no text: two rects in place pair by geometry; the third rect
    // sits beyond a quarter of the diagonal and stays unmatched.
    const from = slideOf([
      obj(1, { geometry: 'rect', x: 0, w: 200, h: 100 }),
      obj(2, { geometry: 'rect', x: 7000, w: 200, h: 100 }),
      obj(3, { geometry: 'ellipse', x: 0, w: 200, h: 100 }),
    ])
    const to = slideOf([
      obj(4, { geometry: 'rect', x: 100, w: 200, h: 100 }),
      obj(5, { geometry: 'rect', x: 7100, w: 200, h: 100 }),
      obj(6, { geometry: 'rect', x: W / PX, w: 200, h: 100 }),
    ])
    const match = matchMorph(from, to, 'byObject')
    expect(match.pairs).toEqual([
      { from: 'sp:1', to: 'sp:4', rule: 'geometry' },
      { from: 'sp:2', to: 'sp:5', rule: 'geometry' },
    ])
    expect(match.fadeOut).toEqual(['sp:3'])
    expect(match.fadeIn).toEqual(['sp:6'])
  })

  test('the geometry threshold is inclusive at a quarter of the diagonal', () => {
    // Slide diagonal is 1200 px, so the threshold is 300 px: the frame centres
    // sit exactly 300 px apart.
    const from = slideOf([obj(1, { geometry: 'rect', x: 0, y: 0, w: 200, h: 100 })])
    const to = slideOf([obj(2, { geometry: 'rect', x: 300, y: 0, w: 200, h: 100 })])
    expect(matchMorph(from, to, 'byObject').pairs).toEqual([
      { from: 'sp:1', to: 'sp:2', rule: 'geometry' },
    ])
  })
})

// ---------------------------------------------------------------------------
// Text units (byWord / byChar)
// ---------------------------------------------------------------------------
describe('morph matching: text units', () => {
  test('byWord and byChar pair text units', () => {
    const from = slideOf([obj(1, { text: 'ab cd' })])
    const to = slideOf([obj(2, { text: 'ab cd' })])

    const words = matchMorph(from, to, 'byWord')
    // Unnamed text-bearing objects pair their units by content first; the
    // container pair follows (its partner is only known after all unit
    // buckets matched) and keeps the frame animatable.
    expect(words.pairs).toEqual([
      { from: 'sp:1/c:0-1', to: 'sp:2/c:0-1', rule: 'text' },
      { from: 'sp:1/c:3-4', to: 'sp:2/c:3-4', rule: 'text' },
      { from: 'sp:1', to: 'sp:2', rule: 'text' },
    ])
    expect(words.fadeOut).toEqual([])
    expect(words.fadeIn).toEqual([])

    const chars = matchMorph(from, to, 'byChar')
    expect(chars.pairs).toEqual([
      { from: 'sp:1/c:0-0', to: 'sp:2/c:0-0', rule: 'text' },
      { from: 'sp:1/c:1-1', to: 'sp:2/c:1-1', rule: 'text' },
      { from: 'sp:1/c:3-3', to: 'sp:2/c:3-3', rule: 'text' },
      { from: 'sp:1/c:4-4', to: 'sp:2/c:4-4', rule: 'text' },
      { from: 'sp:1', to: 'sp:2', rule: 'text' },
    ])

    // byObject pairs the same content once, at object level, with no units.
    const whole = matchMorph(from, to, 'byObject')
    expect(whole.pairs).toEqual([{ from: 'sp:1', to: 'sp:2', rule: 'text' }])
  })

  test('byWord pairs words when names match but text differs', () => {
    const match = matchMorph(
      slideOf([obj(1, { name: 'T', text: 'ab cd' })]),
      slideOf([obj(2, { name: 'T', text: 'ab xy' })]),
      'byWord',
    )
    // The shared name pairs the objects; the common word pairs as a unit; the
    // changed words fade at unit level (their containers already have targets).
    expect(match.pairs).toEqual([
      { from: 'sp:1', to: 'sp:2', rule: 'name-kind' },
      { from: 'sp:1/c:0-1', to: 'sp:2/c:0-1', rule: 'text' },
    ])
    expect(match.fadeOut).toEqual(['sp:1/c:3-4'])
    expect(match.fadeIn).toEqual(['sp:2/c:3-4'])
  })

  test('byWord pairs common words between different texts', () => {
    const swapped = matchMorph(
      slideOf([obj(1, { text: 'ab cd' })]),
      slideOf([obj(2, { text: 'cd ab' })]),
      'byWord',
    )
    expect(swapped.pairs).toEqual([
      { from: 'sp:1/c:0-1', to: 'sp:2/c:3-4', rule: 'text' },
      { from: 'sp:1/c:3-4', to: 'sp:2/c:0-1', rule: 'text' },
      { from: 'sp:1', to: 'sp:2', rule: 'text' },
    ])
    expect(swapped.fadeOut).toEqual([])
    expect(swapped.fadeIn).toEqual([])

    const grown = matchMorph(
      slideOf([obj(1, { text: 'ab cd' })]),
      slideOf([obj(2, { text: 'ab cd ef' })]),
      'byWord',
    )
    expect(grown.pairs).toEqual([
      { from: 'sp:1/c:0-1', to: 'sp:2/c:0-1', rule: 'text' },
      { from: 'sp:1/c:3-4', to: 'sp:2/c:3-4', rule: 'text' },
      { from: 'sp:1', to: 'sp:2', rule: 'text' },
    ])
    // The container is paired, so leftover units fade; the frame does not.
    expect(grown.fadeIn).toEqual(['sp:2/c:6-7'])
  })
})

// ---------------------------------------------------------------------------
// Fades and one-to-one pairing
// ---------------------------------------------------------------------------
describe('morph matching: fades and one-to-one', () => {
  test('unmatched objects fade out and in', () => {
    const from = slideOf([obj(1, { name: 'A', text: 'one', geometry: 'rect', x: 0 })])
    const to = slideOf([obj(2, { name: 'C', text: 'two', geometry: 'ellipse', x: 8000 })])
    expect(matchMorph(from, to, 'byObject')).toEqual({
      pairs: [],
      fadeOut: ['sp:1'],
      fadeIn: ['sp:2'],
    })
  })

  test('one-to-one: no object pairs twice', () => {
    // Two same-text sources compete for one partner: reading order wins once,
    // the loser fades; duplicate !! names behave the same (published 1:1 rule).
    const a = matchMorph(
      slideOf([obj(1, { text: 'same' }), obj(2, { text: 'same' })]),
      slideOf([obj(3, { text: 'same' })]),
      'byObject',
    )
    expect(a.pairs).toEqual([{ from: 'sp:1', to: 'sp:3', rule: 'text' }])
    expect(a.fadeOut).toEqual(['sp:2'])

    const b = matchMorph(
      slideOf([obj(4, { name: '!!K' }), obj(5, { name: '!!K' })]),
      slideOf([obj(6, { name: '!!K' })]),
      'byObject',
    )
    expect(b.pairs).toEqual([{ from: 'sp:4', to: 'sp:6', rule: '!!' }])
    expect(b.fadeOut).toEqual(['sp:5'])

    const seen = new Set<string>()
    for (const pair of [...a.pairs, ...b.pairs]) {
      expect(seen.has(pair.from)).toBe(false)
      expect(seen.has(pair.to)).toBe(false)
      seen.add(pair.from)
      seen.add(pair.to)
    }
  })

  test('duplicate source ids yield one key per id', () => {
    // Two shapes share cNvPr id 1 on each slide. Matching keys off the id, so
    // first wins as elsewhere in the engine: the first pairing stands, and
    // the duplicate claims nothing (no pair, no second fade entry).
    const a = matchMorph(
      slideOf([obj(1, { text: 'a' }), obj(1, { text: 'b' })]),
      slideOf([obj(2, { text: 'a' }), obj(2, { text: 'b' })]),
      'byObject',
    )
    expect(a).toEqual({
      pairs: [{ from: 'sp:1', to: 'sp:2', rule: 'text' }],
      fadeOut: [],
      fadeIn: [],
    })

    const b = matchMorph(
      slideOf([obj(1, { name: '!!A' }), obj(1, { name: '!!B' })]),
      slideOf([obj(2, { name: '!!A' }), obj(2, { name: '!!B' })]),
      'byObject',
    )
    expect(b.pairs).toEqual([{ from: 'sp:1', to: 'sp:2', rule: '!!' }])
  })

  test('shapes without a source id claim nothing', () => {
    // Id-less shapes have no target key, like everywhere else in the engine:
    // they never pair and never enter a fade list, but their keyed partners
    // still behave normally.
    const match = matchMorph(
      slideOf([obj(undefined, { sourceless: true, text: 'hi' })]),
      slideOf([obj(2, { text: 'hi' })]),
      'byObject',
    )
    expect(match).toEqual({ pairs: [], fadeOut: [], fadeIn: ['sp:2'] })
  })
})

// ---------------------------------------------------------------------------
// Kinds
// ---------------------------------------------------------------------------
describe('morph matching: kinds', () => {
  test('adding text does not change kind', () => {
    const match = matchMorph(
      slideOf([obj(1, { name: 'R' })]),
      slideOf([obj(2, { name: 'R', text: 'hi' })]),
      'byObject',
    )
    expect(match.pairs).toEqual([{ from: 'sp:1', to: 'sp:2', rule: 'name-kind' }])
    expect(match.fadeOut).toEqual([])
    expect(match.fadeIn).toEqual([])
  })

  test('geometry pairs require the same kind', () => {
    // An untagged text shape and a picture share a geometry bucket but not a
    // kind; the geometry rule must not cross-fade one into the other.
    const a = matchMorph(
      slideOf([obj(1, { text: 'q', geometry: 'rect', x: 0, y: 0 })]),
      slideOf([obj(2, { image: true, geometry: 'rect', x: 0, y: 0 })]),
      'byObject',
    )
    expect(a).toEqual({ pairs: [], fadeOut: ['sp:1'], fadeIn: ['sp:2'] })

    const b = matchMorph(
      slideOf([obj(3, { children: true, x: 0, y: 0 })]),
      slideOf([obj(4, { image: true, x: 0, y: 0 })]),
      'byObject',
    )
    expect(b).toEqual({ pairs: [], fadeOut: ['sp:3'], fadeIn: ['sp:4'] })
  })

  test('groups pair as wholes and children claim nothing', () => {
    const match = matchMorph(
      slideOf([obj(1, { children: true, name: 'G' })]),
      slideOf([obj(2, { children: true, name: 'G' })]),
      'byObject',
    )
    expect(match).toEqual({
      pairs: [{ from: 'sp:1', to: 'sp:2', rule: 'name-kind' }],
      fadeOut: [],
      fadeIn: [],
    })
  })
})

// ---------------------------------------------------------------------------
// Bounds: refusal and diagnosis
// ---------------------------------------------------------------------------
describe('morph matching: bounds', () => {
  test('the object bound refuses and diagnoses', () => {
    const many = (id0: number): PptxShape[] =>
      Array.from({ length: 257 }, (_, i) => obj(id0 + i, { name: 'N' + i, text: 't' + i }))
    const sink = createDiagnosticSink()
    const match = matchMorph(slideOf(many(0), 3), slideOf(many(1000), 4), 'byObject', sink)
    expect(match).toEqual({ pairs: [], fadeOut: [], fadeIn: [] })
    expect(sink.list()).toEqual([
      { code: 'morph-refused', slideIndex: 4, effect: 'sld', message: expect.stringContaining('257') },
    ])
  })

  test('the object bound accepts exactly the limit', () => {
    const many = (start: number): PptxShape[] =>
      Array.from({ length: 256 }, (_, i) => obj(start + i, { name: 'N' + i }))
    const sink = createDiagnosticSink()
    const match = matchMorph(slideOf(many(0)), slideOf(many(1000)), 'byObject', sink)
    expect(sink.list()).toEqual([])
    expect(match.pairs.length).toBe(256)
    expect(match.pairs.every(p => p.rule === 'name-kind')).toBe(true)
  })

  test('the unit bound coarsens monotonically and is diagnosed', () => {
    // 82 pairs need 25 word units each: the cap (2048) is reached inside the
    // 82nd. From there on, every pair coarsens to its container pair — even
    // later one-word pairs that would still fit individually.
    const from: PptxShape[] = []
    const to: PptxShape[] = []
    for (let i = 0; i < 100; i++) {
      const body = i < 82 ? `T${i} ${Array.from({ length: 24 }, (_, k) => `w${k}`).join(' ')}` : `T${i}`
      from.push(obj(i, { name: `P${i}`, text: body }))
      to.push(obj(1000 + i, { name: `P${i}`, text: body }))
    }
    const sink = createDiagnosticSink()
    const match = matchMorph(slideOf(from), slideOf(to), 'byWord', sink)
    const units = match.pairs.filter(p => p.from.includes('/c:'))
    const objects = match.pairs.filter(p => !p.from.includes('/c:'))
    expect(objects.length).toBe(100)
    expect(units.length).toBe(81 * 25)
    // Monotonic: nothing after the 2048 cap gets unit treatment, not even the
    // one-word pairs that would still fit.
    expect(match.pairs.some(p => p.from.startsWith('sp:82/c:'))).toBe(false)
    expect(match.pairs.some(p => p.from.startsWith('sp:99/c:'))).toBe(false)
    expect(sink.list()).toEqual([
      {
        code: 'morph-units-coarsened', slideIndex: 0, effect: 'sld',
        message: expect.stringContaining('2048'),
      },
    ])
  })
})

// ---------------------------------------------------------------------------
// Degenerate slides
// ---------------------------------------------------------------------------
describe('morph matching: degenerate slides', () => {
  test('empty slides pair nothing and fade nothing', () => {
    expect(matchMorph(slideOf([]), slideOf([]), 'byWord')).toEqual({
      pairs: [], fadeOut: [], fadeIn: [],
    })
    expect(matchMorph(slideOf([]), slideOf([]), 'byObject')).toEqual({
      pairs: [], fadeOut: [], fadeIn: [],
    })
  })
})
