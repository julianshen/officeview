import { describe, expect, test } from 'vitest'
import type { PptxShape, PptxSlide } from '../src/pptx/types'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { matchMorph } from '../src/slideshow/morph/match'

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------
const PX = 9525
const W = 960 * PX
const H = 720 * PX

/** EMU diagonal of the 960x720 slide, and the 25% geometry-match threshold. */
const DIAG = Math.hypot(W, H)
const THRESHOLD = DIAG / 4

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
}

function obj(id: number, over: ObjSpec = {}): PptxShape {
  const shape: PptxShape = {
    xEmu: (over.x ?? 0) * PX, yEmu: (over.y ?? 0) * PX,
    widthEmu: (over.w ?? 100) * PX, heightEmu: (over.h ?? 50) * PX,
    geometry: over.geometry ?? 'rect',
    source: { partPath: 'ppt/slides/slide1.xml', treePath: `spTree/sp[${id}]`, element: 'sp', id: String(id), representation: 'native' },
  }
  if (over.name !== undefined && shape.source) shape.source.name = over.name
  if (over.text !== undefined) shape.textBody = text(over.text)
  if (over.image) shape.image = { data: new Uint8Array(4), width: 10, height: 10 } as PptxShape['image']
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
    // sp:1 is !!-tagged; sp:2 shares its text and sp:3 its geometry, and both
    // would be legal matches for a non-tagged object. Morph must not pair a
    // !! object with a non-!! object (published rule), so it fades out.
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
    // The far rect is farther than the geometry threshold from every source
    // shape, including diagonally displaced ones.
    expect(THRESHOLD < (W / 2)).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// Text units (byWord / byChar)
// ---------------------------------------------------------------------------
describe('morph matching: text units', () => {
  test('byWord and byChar pair text units', () => {
    const from = slideOf([obj(1, { text: 'ab cd' })])
    const to = slideOf([obj(2, { text: 'ab cd' })])

    const words = matchMorph(slideOf([obj(1, { text: 'ab cd' })]), slideOf([obj(2, { text: 'ab cd' })]), 'byWord')
    expect(words.pairs).toEqual([
      { from: 'sp:1/c:0-1', to: 'sp:2/c:0-1', rule: 'text' },
      { from: 'sp:1/c:3-4', to: 'sp:2/c:3-4', rule: 'text' },
    ])
    expect(words.fadeOut).toEqual([])
    expect(words.fadeIn).toEqual([])

    const chars = matchMorph(from, to, 'byChar')
    expect(chars.pairs).toEqual([
      { from: 'sp:1/c:0-0', to: 'sp:2/c:0-0', rule: 'text' },
      { from: 'sp:1/c:1-1', to: 'sp:2/c:1-1', rule: 'text' },
      { from: 'sp:1/c:3-3', to: 'sp:2/c:3-3', rule: 'text' },
      { from: 'sp:1/c:4-4', to: 'sp:2/c:4-4', rule: 'text' },
    ])

    // byObject pairs the same content once, at object level.
    const whole = matchMorph(from, to, 'byObject')
    expect(whole.pairs).toEqual([{ from: 'sp:1', to: 'sp:2', rule: 'text' }])
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
})
