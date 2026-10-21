import { describe, expect, it } from 'vitest'
import { parseMotionPath, pointAt } from '../src/slideshow/motion-path'
import type { MotionPath } from '../src/slideshow/types'

function parse(src: string): MotionPath {
  const r = parseMotionPath(src)
  if (!r.ok) throw new Error(`expected ok, got ${r.reason}: ${src.slice(0, 60)}`)
  return r.path
}

function refusal(src: string): string | null {
  const r = parseMotionPath(src)
  return r.ok ? null : r.reason
}

function expectPts(actual: readonly number[], expected: readonly number[]): void {
  expect(actual).toHaveLength(expected.length)
  expected.forEach((v, i) => expect(actual[i]).toBeCloseTo(v, 9))
}

function finite(pt: { x: number; y: number; angle: number }): boolean {
  return Number.isFinite(pt.x) && Number.isFinite(pt.y) && Number.isFinite(pt.angle)
}

/** Exactly `n` characters: a valid path padded with trailing spaces. */
function padded(n: number): string {
  const base = 'M 0 0 L 1 0'
  return base + ' '.repeat(n - base.length)
}

describe('slideshow motion path: parsing', () => {
  it('parses M L C Z E, absolute and relative', () => {
    const abs = parse('M 0 0 L 1 0 C 1 1 0 1 0 0 Z E')
    expect(abs.segments.map(s => s.cmd)).toEqual(['M', 'L', 'C', 'Z'])
    expectPts(abs.segments[0].pts, [0, 0])
    expectPts(abs.segments[1].pts, [1, 0])
    expectPts(abs.segments[2].pts, [1, 1, 0, 1, 0, 0])
    expectPts(abs.segments[3].pts, [])

    // Relative: every point is resolved against the current point when the command starts.
    const rel = parse('m 0.1 0.1 l 0.2 0 c 0 0.1 0.1 0.1 0.1 0 z e')
    expect(rel.segments.map(s => s.cmd)).toEqual(['M', 'L', 'C', 'Z'])
    expectPts(rel.segments[0].pts, [0.1, 0.1])
    expectPts(rel.segments[1].pts, [0.3, 0.1])
    expectPts(rel.segments[2].pts, [0.3, 0.2, 0.4, 0.2, 0.4, 0.1])
  })

  it('relative c measures all three points from the point where the command starts', () => {
    const c = parse('M 0 0 c 0 1 1 1 1 0')
    const abs = parse('M 0 0 C 0 1 1 1 1 0')
    expectPts(c.segments[1].pts, [0, 1, 1, 1, 1, 0])
    expectPts(c.segments[1].pts, abs.segments[1].pts)
    // The current point afterwards is the cubic's end point, so a following relative line is exact.
    const chained = parse('M 0.5 0.5 c 0.1 0 0.2 0.1 0.2 0.2 l 0.1 0.1')
    expectPts(chained.segments[1].pts, [0.6, 0.5, 0.7, 0.6, 0.7, 0.7])
    expectPts(chained.segments[2].pts, [0.8, 0.8])
  })

  it('E ends the path: nothing after it is parsed', () => {
    const p = parse('M 0 0 L 1 0 E L 5 5 C 1 2 3 4 5 6')
    expect(p.segments.map(s => s.cmd)).toEqual(['M', 'L'])
    expect(p.length).toBeCloseTo(1, 9)
    // The tail after E is ignored, whatever it holds.
    expect(refusal('M 0 0 L 1 0 E ???')).toBeNull()
    expect(parse('M 0 0 E').segments.map(s => s.cmd)).toEqual(['M'])
  })

  it('accepts exactly 65536 characters and 10000 commands; refuses 65537 characters and 10001 commands', () => {
    expect(padded(65536)).toHaveLength(65536)
    expect(refusal(padded(65536))).toBeNull()
    expect(padded(65537)).toHaveLength(65537)
    expect(refusal(padded(65537))).toBe('too-long')

    const lines = (n: number) => 'M 0 0' + ' L 1 0'.repeat(n - 1)
    expect(lines(10001).length).toBeLessThan(65536)
    expect(parse(lines(10000)).segments).toHaveLength(10000)
    expect(refusal(lines(10001))).toBe('too-many-commands')
    // Z counts as a command; E does not (it is a terminator, not a segment).
    expect(refusal(lines(9999) + ' Z')).toBeNull()
    expect(refusal(lines(10000) + ' Z')).toBe('too-many-commands')
    expect(refusal(lines(10000) + ' E')).toBeNull()
  })

  it('refuses NaN, Infinity and stray tokens', () => {
    expect(refusal('M 0 0 L NaN 0')).toBe('syntax')
    expect(refusal('M 0 0 L Infinity 0')).toBe('syntax')
    expect(refusal('M 0 0 L -Infinity 0')).toBe('syntax')
    expect(refusal('M 0 0 foo')).toBe('syntax')
    expect(refusal('M 0 0 L 1')).toBe('syntax')
    expect(refusal('M 0 0 L 1 0 2')).toBe('syntax')
    expect(refusal('M 0 0 L 1e999 0')).toBe('syntax')
    // Only decimal numbers: no hex, binary or octal forms that Number() would accept.
    expect(refusal('M 0x10 0 L 1 0')).toBe('syntax')
    expect(refusal('M 0b1 0 L 1 0')).toBe('syntax')
    expect(refusal('M 0 0 L 0o7 0')).toBe('syntax')
  })

  it('refuses coordinates beyond +-1e6 slide fractions, including accumulated relative ones', () => {
    expect(refusal('M 0 0 L 1000000 -1000000')).toBeNull()
    expect(refusal('M 0 0 L 1000001 0')).toBe('syntax')
    expect(refusal('M 0 0 L 0 -1000001')).toBe('syntax')
    expect(refusal('M 0 0 L 1e308 0 L -1e308 0')).toBe('syntax')
    expect(refusal('M 0 0 l 1000000 0 l 1 0')).toBe('syntax')
    expect(refusal('M 0 0 C 0 0 2000000 0 0 0')).toBe('syntax')
  })

  it('refuses an empty path and a path that does not start with M; accepts compact tokens', () => {
    expect(refusal('')).toBe('syntax')
    expect(refusal('   ')).toBe('syntax')
    expect(refusal('E')).toBe('syntax')
    expect(refusal('L 1 1')).toBe('syntax')
    expect(refusal('Z')).toBe('syntax')
    expect(refusal('M 0 0 M 1 1')).toBeNull()
    expectPts(parse('M0,0L1,0.5E').segments[1].pts, [1, 0.5])
    // Exponent numbers as PowerPoint writes them.
    expectPts(parse('M -4.16667E-6 0 L 1 0').segments[0].pts, [-4.16667e-6, 0])
    expectPts(parse('M .5 +.25 L 1 0').segments[0].pts, [0.5, 0.25])
  })
})

describe('slideshow motion path: length and pointAt', () => {
  it('pointAt 0, 0.5 and 1 on a known line, with angle in radians', () => {
    const p = parse('M 0 0 L 1 0')
    expect(p.length).toBeCloseTo(1, 9)
    for (const [f, x] of [[0, 0], [0.5, 0.5], [1, 1]] as const) {
      const pt = pointAt(p, f)
      expect(pt.x).toBeCloseTo(x, 9)
      expect(pt.y).toBeCloseTo(0, 9)
      expect(pt.angle).toBeCloseTo(0, 9)
    }
    const down = pointAt(parse('M 0 0 L 0 2'), 0.25)
    expect(down.x).toBeCloseTo(0, 9)
    expect(down.y).toBeCloseTo(0.5, 9)
    expect(down.angle).toBeCloseTo(Math.PI / 2, 9)
    expect(pointAt(parse('M 1 0 L 0 0'), 0.5).angle).toBeCloseTo(Math.PI, 9)
  })

  it('walks a polyline at constant speed with the angle of the edge it is on', () => {
    const p = parse('M 0 0 L 1 0 L 1 1')
    expect(p.length).toBeCloseTo(2, 9)
    const a = pointAt(p, 0.25)
    expect(a.x).toBeCloseTo(0.5, 9)
    expect(a.y).toBeCloseTo(0, 9)
    expect(a.angle).toBeCloseTo(0, 9)
    const b = pointAt(p, 0.75)
    expect(b.x).toBeCloseTo(1, 9)
    expect(b.y).toBeCloseTo(0.5, 9)
    expect(b.angle).toBeCloseTo(Math.PI / 2, 9)
    const end = pointAt(p, 1)
    expect(end.x).toBeCloseTo(1, 9)
    expect(end.y).toBeCloseTo(1, 9)
    expect(end.angle).toBeCloseTo(Math.PI / 2, 9)
  })

  it('samples a known cubic by arc length at 0, 0.25, 0.5, 0.75 and 1', () => {
    // Symmetric arch: B(0.5) = (0.5, 0.75) exactly, and the curve mirrors about x = 0.5.
    const p = parse('M 0 0 C 0 1 1 1 1 0')
    expect(p.length).toBeGreaterThan(1.9)
    expect(p.length).toBeLessThan(2.1)

    const p0 = pointAt(p, 0)
    expect(p0.x).toBeCloseTo(0, 9)
    expect(p0.y).toBeCloseTo(0, 9)
    // Leaves towards the first control point (straight up), as seen by the first 1/32 chord.
    expect(p0.angle).toBeGreaterThan(Math.PI / 2 - 0.05)
    expect(p0.angle).toBeLessThanOrEqual(Math.PI / 2)

    const mid = pointAt(p, 0.5)
    expect(mid.x).toBeCloseTo(0.5, 9)
    expect(mid.y).toBeCloseTo(0.75, 9)
    expect(Math.abs(mid.angle)).toBeLessThan(0.1) // apex: heading +x, to within one chord

    const q1 = pointAt(p, 0.25)
    const q3 = pointAt(p, 0.75)
    expect(q1.x + q3.x).toBeCloseTo(1, 9)
    expect(q1.y).toBeCloseTo(q3.y, 9)
    expect(q1.x).toBeGreaterThan(0)
    expect(q1.x).toBeLessThan(0.5)
    expect(q1.y).toBeGreaterThan(0.3)
    expect(q1.y).toBeLessThan(0.75)
    expect(q1.angle).toBeGreaterThan(0)
    expect(q1.angle).toBeLessThan(Math.PI / 2)
    expect(q3.angle).toBeLessThan(0)

    const p1 = pointAt(p, 1)
    expect(p1.x).toBeCloseTo(1, 9)
    expect(p1.y).toBeCloseTo(0, 9)
    expect(p1.angle).toBeLessThan(-Math.PI / 2 + 0.05)
    expect(p1.angle).toBeGreaterThanOrEqual(-Math.PI / 2)
  })

  it('cubic length is the sum of 32 chords', () => {
    // Collinear control points: the polyline is the straight line, length exactly 1.
    expect(parse('M 0 0 C 0.25 0 0.75 0 1 0').length).toBeCloseTo(1, 9)
    // Recompute the 32-chord polyline independently.
    const bez = (t: number, a: number, b: number, c: number, d: number) =>
      (1 - t) ** 3 * a + 3 * (1 - t) ** 2 * t * b + 3 * (1 - t) * t * t * c + t ** 3 * d
    let want = 0
    let px = 0
    let py = 0
    for (let s = 1; s <= 32; s++) {
      const x = bez(s / 32, 0, 0, 1, 1)
      const y = bez(s / 32, 0, 1, 1, 0)
      want += Math.hypot(x - px, y - py)
      px = x
      py = y
    }
    expect(parse('M 0 0 C 0 1 1 1 1 0').length).toBeCloseTo(want, 12)
  })

  it('Z closes the path with a straight edge that counts toward its length', () => {
    const p = parse('M 0 0 L 1 0 L 1 1 Z E')
    const total = 2 + Math.SQRT2
    expect(p.length).toBeCloseTo(total, 9)
    const end = pointAt(p, 1)
    expect(end.x).toBeCloseTo(0, 9)
    expect(end.y).toBeCloseTo(0, 9)
    expect(end.angle).toBeCloseTo(-3 * Math.PI / 4, 9) // heading back down-left
    const corner = pointAt(p, 2 / total)
    expect(corner.x).toBeCloseTo(1, 9)
    expect(corner.y).toBeCloseTo(1, 9)
    // Half way along the closing edge.
    const half = pointAt(p, (2 + Math.SQRT2 / 2) / total)
    expect(half.x).toBeCloseTo(0.5, 9)
    expect(half.y).toBeCloseTo(0.5, 9)
  })

  it('Z on an already-closed path adds nothing, and drawing continues from the subpath start', () => {
    expect(parse('M 0 0 L 1 0 L 0 0 Z').length).toBeCloseTo(2, 9)
    expect(parse('M 0 0 L 1 0 Z L 0 1').length).toBeCloseTo(3, 9)
    const rel = parse('M 0.2 0.2 L 0.6 0.2 z l 0.1 0')
    expectPts(rel.segments[3].pts, [0.3, 0.2])
  })

  it('treats a non-finite or out-of-range fraction as a position on the path, never NaN', () => {
    const p = parse('M 0 0 L 1 0')
    expect(pointAt(p, Number.NaN)).toEqual({ x: 0, y: 0, angle: 0 })
    expect(pointAt(p, Number.POSITIVE_INFINITY).x).toBeCloseTo(1, 9)
    expect(pointAt(p, Number.NEGATIVE_INFINITY).x).toBeCloseTo(0, 9)
    expect(pointAt(p, -3).x).toBeCloseTo(0, 9)
    expect(pointAt(p, 3).x).toBeCloseTo(1, 9)
    for (const f of [Number.NaN, Infinity, -Infinity, -1, 2, 0.3]) expect(finite(pointAt(p, f))).toBe(true)
  })

  it('stays finite for a hand-built path whose coordinates overflow or are NaN', () => {
    const huge: MotionPath = {
      segments: [
        { cmd: 'M', pts: [0, 0] },
        { cmd: 'L', pts: [1e308, 0] },
        { cmd: 'L', pts: [-1e308, 0] },
      ],
      length: Infinity,
    }
    for (const f of [0, 0.5, 1]) expect(finite(pointAt(huge, f))).toBe(true)
    const nanPath: MotionPath = { segments: [{ cmd: 'M', pts: [0, 0] }, { cmd: 'L', pts: [Number.NaN, 1] }], length: 0 }
    for (const f of [0, 0.5, 1]) expect(finite(pointAt(nanPath, f))).toBe(true)
  })

  it('an M jump is not a traversable edge, and a zero-length edge gives no angle', () => {
    const start = pointAt(parse('M 0 0 M 1 1 L 2 2'), 0)
    expect(start.x).toBeCloseTo(1, 9)
    expect(start.y).toBeCloseTo(1, 9)
    expect(start.angle).toBeCloseTo(Math.PI / 4, 9)
    expect(pointAt(parse('M 0 0 L 0 0 L 0 1'), 0).angle).toBeCloseTo(Math.PI / 2, 9)
  })

  it('a path with no length stays at its start point', () => {
    expect(pointAt(parse('M 0.25 0.5'), 0.7)).toEqual({ x: 0.25, y: 0.5, angle: 0 })
    expect(pointAt({ segments: [], length: 0 }, 0.5)).toEqual({ x: 0, y: 0, angle: 0 })
  })

  it('samples once per path: pointAt does not re-read or re-sample the segments', () => {
    const src = parse('M 0 0 C 0 1 1 1 1 0 L 2 0')
    // A hand-built path is sampled on its first pointAt and the table is reused after that.
    let reads = 0
    const handBuilt = {
      get segments() { reads++; return src.segments },
      length: src.length,
    } as MotionPath
    const first = pointAt(handBuilt, 0.4)
    for (let i = 0; i < 50; i++) expect(pointAt(handBuilt, 0.4)).toEqual(first)
    expect(reads).toBe(1)

    // A parsed path is sampled by the parser (it needs the length anyway), so pointAt reads nothing.
    const parsed = parse('M 0 0 C 0 1 1 1 1 0 L 2 0')
    let parsedReads = 0
    const segments = parsed.segments
    Object.defineProperty(parsed, 'segments', { get() { parsedReads++; return segments } })
    for (let i = 0; i < 50; i++) pointAt(parsed, i / 50)
    expect(parsedReads).toBe(0)
  })
})
