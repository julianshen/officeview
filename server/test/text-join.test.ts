/** S3a Task 2: R1 — the visible-text join between index spans and occlusion runs. */
import { describe, expect, test } from 'bun:test'
import type { TextPlacement, TextSpan } from '../../src/core/search'
import type { RunDetail } from '../src/occlusion'
import { joinSpans } from '../src/text/join'

const IDENTITY = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }
/** A span drawn at local (x, baseline) under `transform`, band top/bottom = baseline -15 / +5. */
function span(text: string, x: number, y: number, width: number, transform = IDENTITY): TextSpan {
  const placement: TextPlacement = { x, y, width, top: y - 15, bottom: y + 5, transform }
  const px = transform.a * x + transform.c * y + transform.e, py = transform.b * x + transform.d * y + transform.f
  return { text, x: px, y: py, width: width * Math.hypot(transform.a, transform.b), fontSize: 20, placement }
}
const run = (text: string, x: number, y: number, width: number, height: number, servable = true): RunDetail =>
  servable ? { text, ink: { x, y, width, height }, servable } : { text, ink: { x, y, width, height }, servable, reason: 'covered' }

describe('joinSpans (R1)', () => {
  test('one span, one servable run at the same box ⇒ served', () => {
    const s = span('Hello', 20, 50, 50)
    expect(joinSpans([s], [run('Hello', 20, 36, 50, 18)], 1)).toEqual({ served: [s], withheld: 0 })
  })
  test('span whose run is covered ⇒ withheld', () => {
    expect(joinSpans([span('Hello', 20, 50, 50)], [run('Hello', 20, 36, 50, 18, false)], 1)).toEqual({ served: [], withheld: 1 })
  })
  test('span with no run ⇒ withheld', () => { // glyph-path text
    expect(joinSpans([span('Hello', 20, 50, 50)], [run('Other', 300, 300, 50, 18)], 1)).toEqual({ served: [], withheld: 1 })
  })
  test('span over two runs, one withheld ⇒ withheld', () => {
    const r = joinSpans([span('Hello world', 20, 50, 110)], [run('Hello', 20, 36, 50, 18), run(' world', 75, 36, 55, 18, false)], 1)
    expect(r).toEqual({ served: [], withheld: 1 })
  })
  test('per-glyph runs covering every character ⇒ served', () => {
    const s = span('abc', 20, 50, 30)
    const r = joinSpans([s], [run('a', 20, 38, 10, 14), run('b', 30, 38, 10, 14), run('c', 40, 38, 10, 14)], 1)
    expect(r.served).toEqual([s])
  })
  test('run text missing a character of the span ⇒ withheld', () => {
    expect(joinSpans([span('abc', 20, 50, 30)], [run('a', 20, 38, 10, 14), run('b', 30, 38, 10, 14)], 1).withheld).toBe(1)
  })
  test('coverage counts code points as a multiset (repeated letters need repeated runs)', () => {
    expect(joinSpans([span('aa', 20, 50, 20)], [run('a', 20, 38, 10, 14)], 1).withheld).toBe(1)
    expect(joinSpans([span('aa', 20, 50, 20)], [run('a', 20, 38, 10, 14), run('a', 30, 38, 10, 14)], 1).served).toHaveLength(1)
    // astral code points are one unit each, no normalisation
    expect(joinSpans([span('😀', 20, 50, 20)], [run('😀', 20, 38, 20, 14)], 1).served).toHaveLength(1)
    expect(joinSpans([span('é', 20, 50, 20)], [run('é', 20, 38, 20, 14)], 1).withheld).toBe(1) // NFC span vs NFD run
  })
  test('scale 2 ink maps onto page-px spans', () => {
    const s = span('Hello', 20, 50, 50)
    expect(joinSpans([s], [run('Hello', 40, 72, 100, 36)], 2).served).toEqual([s])
    // the same device box read at scale 1 lies elsewhere on the page
    expect(joinSpans([s], [run('Hello', 140, 172, 100, 36)], 2).withheld).toBe(1)
  })
  test('a translated and scaled placement is projected through its transform', () => {
    const s = span('Hi', 10, 20, 20, { a: 2, b: 0, c: 0, d: 2, e: 100, f: 100 }) // page band x 120..160, y 110..150
    expect(joinSpans([s], [run('Hi', 122, 112, 36, 36)], 1).served).toEqual([s])
    // a run at the untransformed local band does not match
    expect(joinSpans([s], [run('Hi', 10, 5, 20, 20)], 1).withheld).toBe(1)
  })
  test('a rotated placement uses the bounds of its projected band', () => {
    const t = { a: 0, b: 1, c: -1, d: 0, e: 200, f: 0 } // 90° — local (x, y) → page (200 - y, x)
    const s = span('Up', 10, 50, 40, t) // page band x 145..165, y 10..50
    expect(joinSpans([s], [run('Up', 146, 12, 18, 36)], 1).served).toEqual([s])
    expect(joinSpans([s], [run('Up', 10, 35, 40, 20)], 1).withheld).toBe(1)
  })
  test('the visual polygon, when present, is the span box', () => {
    const s = span('Arc', 20, 50, 30)
    s.placement!.visual = { polygon: [{ x: 300, y: 300 }, { x: 340, y: 300 }, { x: 340, y: 320 }, { x: 300, y: 320 }] }
    expect(joinSpans([s], [run('Arc', 302, 302, 36, 16)], 1).served).toEqual([s])
    expect(joinSpans([s], [run('Arc', 20, 36, 30, 18)], 1).withheld).toBe(1)
  })
  test('a span without placement uses {x, y - fontSize, width, fontSize * 1.25}', () => {
    const s: TextSpan = { text: 'Cell', x: 20, y: 50, width: 40, fontSize: 20 } // box 20..60, 30..55
    expect(joinSpans([s], [run('Cell', 22, 36, 36, 16)], 1).served).toEqual([s])
    expect(joinSpans([s], [run('Cell', 22, 70, 36, 16)], 1).withheld).toBe(1)
  })
  test('a run is a candidate when the span centre lies inside its ink (large run over a small span)', () => {
    const s = span('a', 50, 50, 10) // box 50..60, 35..55 — centre (55, 45)
    // centre of this run (100, 45) is outside the span box but the run ink contains the span centre
    expect(joinSpans([s], [run('a big run', 0, 30, 200, 30)], 1).served).toEqual([s])
    expect(joinSpans([s], [run('a big run', 0, 30, 200, 30, false)], 1).withheld).toBe(1)
  })
  test('the 2 px inflation admits a run centre just outside the box, not further', () => {
    const s: TextSpan = { text: 'x', x: 20, y: 50, width: 10, fontSize: 20 } // box 20..30, 30..55
    expect(joinSpans([s], [run('x', 25, 55, 6, 2)], 1).served).toHaveLength(1) // centre (28, 56)
    expect(joinSpans([s], [run('x', 25, 57, 6, 2)], 1).withheld).toBe(1) // centre (28, 58)
  })
  test('served spans are the original objects, in input order', () => {
    const a = span('One', 20, 50, 30), b = span('Two', 100, 50, 30), c = span('Six', 20, 100, 30)
    const r = joinSpans([a, b, c], [run('One', 20, 38, 30, 14), run('Two', 100, 38, 30, 14, false), run('Six', 20, 88, 30, 14)], 1)
    expect(r.served[0]).toBe(a)
    expect(r.served[1]).toBe(c)
    expect(r.withheld).toBe(1)
  })
})

describe('joinSpans — whitespace-only spans', () => {
  const left = span('Hello', 20, 50, 50), gap = span(' ', 70, 50, 5), right = span('world', 75, 50, 50)
  const lr = run('Hello', 20, 38, 50, 14), rr = run('world', 75, 38, 50, 14)
  test('served when both neighbours on the line are served', () => {
    expect(joinSpans([left, gap, right], [lr, rr], 1)).toEqual({ served: [left, gap, right], withheld: 0 })
  })
  test('withheld when either existing neighbour is withheld', () => {
    expect(joinSpans([left, gap, right], [lr, { ...rr, servable: false, reason: 'covered' }], 1)).toEqual({ served: [left], withheld: 2 })
    expect(joinSpans([left, gap, right], [{ ...lr, servable: false, reason: 'covered' }, rr], 1)).toEqual({ served: [right], withheld: 2 })
  })
  test('served with a single neighbour when only that one exists and is served', () => {
    const tail = span('  ', 125, 50, 10)
    expect(joinSpans([left, right, tail], [lr, rr], 1).served).toEqual([left, right, tail])
  })
  test('withheld with no non-whitespace neighbour on the same line', () => {
    const lone = span(' ', 20, 150, 5), empty = span('', 20, 200, 0)
    expect(joinSpans([left, lone, empty], [lr], 1)).toEqual({ served: [left], withheld: 2 })
  })
  test('neighbours on a different baseline (> 1 px) do not count', () => {
    const below = span(' ', 70, 52, 5)
    expect(joinSpans([left, below], [lr], 1).served).toEqual([left])
    const near = span(' ', 70, 50.8, 5)
    expect(joinSpans([left, near], [lr], 1).served).toEqual([left, near])
  })
  test('the nearest neighbour decides, skipping other whitespace and other lines', () => {
    const other = span('Elsewhere', 20, 150, 90)
    const g1 = span(' ', 70, 50, 3), g2 = span(' ', 73, 50, 2)
    const r = joinSpans([left, g1, other, g2, right], [lr, rr, run('Elsewhere', 20, 138, 90, 14, false)], 1)
    expect(r.served).toEqual([left, g1, g2, right])
  })
})
