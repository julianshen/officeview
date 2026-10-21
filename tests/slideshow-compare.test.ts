import { describe, expect, it } from 'vitest'
import { compareFrames, readRefs, TOLERANCE_EXACT, TOLERANCE_APPROX } from '../scripts/slideshow-compare'

function makeImage(w: number, h: number, fill: (x: number, y: number) => [number, number, number]): ImageData {
  const data = new Uint8ClampedArray(w * h * 4)
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const [r, g, b] = fill(x, y)
    const i = (y * w + x) * 4
    data[i] = r; data[i+1] = g; data[i+2] = b; data[i+3] = 255
  }
  return { data, width: w, height: h, colorSpace: 'srgb' } as ImageData
}

describe('slideshow-compare', () => {
  it('identical frames pass; a one-row stripe fails rowMax but passes meanAbs', () => {
    const a = makeImage(32, 32, () => [10, 20, 30])
    const b = makeImage(32, 32, () => [10, 20, 30])
    const r1 = compareFrames(a, b, TOLERANCE_EXACT)
    expect(r1.pass).toBe(true)
    expect(r1.meanAbs).toBe(0)
    // one row differs by a lot
    const c = makeImage(32, 32, (_x, y) => y === 1 ? [25, 35, 45] : [10, 20, 30])
    const r2 = compareFrames(a, c, TOLERANCE_EXACT)
    expect(r2.meanAbs).toBeLessThanOrEqual(TOLERANCE_EXACT.meanAbs)
    expect(r2.rowMax).toBeGreaterThan(TOLERANCE_EXACT.rowMax)
    expect(r2.pass).toBe(false)
  })
  it('approximate rows use the looser tolerance', () => {
    const a = makeImage(32, 32, () => [10, 20, 30])
    const c = makeImage(32, 32, (_x, y) => y === 1 ? [25, 35, 45] : [10, 20, 30])
    const exact = compareFrames(a, c, TOLERANCE_EXACT)
    const approx = compareFrames(a, c, TOLERANCE_APPROX)
    expect(exact.pass).toBe(false)
    expect(approx.pass).toBe(true)
  })
  it('a missing ref leaves the row unverified, never failed', async () => {
    const set = await readRefs('/tmp/nonexistent-slideshow-refs')
    expect(set.files).toEqual([])
  })
  it('the checklist lists every corpus deck and every behavioural question', async () => {
    const fs = await import('node:fs/promises')
    const txt = await fs.readFile('docs/validation/slideshow-captures.md', 'utf8')
    expect(txt).toContain('tdf104788')
    expect(txt).toContain('spd')
    expect(txt).toContain('advTm')
    expect(txt).toContain('kiosk')
  })
})
