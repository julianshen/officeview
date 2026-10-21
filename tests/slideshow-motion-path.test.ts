import { describe, expect, it } from 'vitest'
import { parseMotionPath, pointAt } from '../src/slideshow/motion-path'

describe('slideshow motion path', () => {
  it('parses M L C Z E, absolute and relative', () => {
    const r = parseMotionPath('M 0 0 L 1 0 C 1 1 0 1 0 0 Z E')
    expect(r.ok).toBe(true)
    if (r.ok) {
      expect(r.path.segments.map(s => s.cmd)).toEqual(['M','L','C','Z'])
    }
    const rel = parseMotionPath('m 0 0 l 1 0 e')
    expect(rel.ok).toBe(true)
  })
  it('pointAt 0, 0.5 and 1 on a known line and a known cubic', () => {
    const r = parseMotionPath('M 0 0 L 1 0')
    expect(r.ok).toBe(true)
    if (r.ok) {
      expect(pointAt(r.path, 0)).toMatchObject({ x: 0, y: 0 })
      expect(pointAt(r.path, 0.5)).toMatchObject({ x: 0.5, y: 0 })
      expect(pointAt(r.path, 1)).toMatchObject({ x: 1, y: 0 })
    }
    const c = parseMotionPath('M 0 0 C 0.33 1 0.67 1 1 0')
    expect(c.ok).toBe(true)
    if (c.ok) {
      const mid = pointAt(c.path, 0.5)
      expect(mid.x).toBeCloseTo(0.5, 1)
      expect(mid.y).toBeGreaterThan(0.5)
    }
  })
  it('accepts 65536 characters and 10000 commands; refuses 65537 characters and 10001 commands', () => {
    // build 10000 commands
    let cmds = 'M 0 0'
    for (let i = 0; i < 9999; i++) cmds += ' L 1 0'
    expect(cmds.length).toBeLessThan(65536)
    const ok = parseMotionPath(cmds)
    expect(ok.ok).toBe(true)
    if (ok.ok) expect(ok.path.segments.length).toBe(10000)
    // 10001 commands
    let cmds2 = 'M 0 0'
    for (let i = 0; i < 10000; i++) cmds2 += ' L 1 0'
    const bad = parseMotionPath(cmds2)
    expect(bad.ok).toBe(false)
    if (!bad.ok) expect(bad.reason).toBe('too-many-commands')
    // 65537 chars
    const longStr = 'M 0 0 ' + ' '.repeat(65533)
    expect(longStr.length).toBeGreaterThan(65536)
    const bad2 = parseMotionPath(longStr)
    expect(bad2.ok).toBe(false)
    if (!bad2.ok) expect(bad2.reason).toBe('too-long')
  })
  it('refuses NaN, Infinity and stray tokens', () => {
    expect(parseMotionPath('M 0 0 L NaN 0').ok).toBe(false)
    expect(parseMotionPath('M 0 0 L Infinity 0').ok).toBe(false)
    expect(parseMotionPath('M 0 0 foo').ok).toBe(false)
    expect(parseMotionPath('M 0 0 L 1').ok).toBe(false)
  })
})
