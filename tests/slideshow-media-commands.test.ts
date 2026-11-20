/** Piece 5 Task 4 (D15, X10, X11): the shared media command whitelist. */
import { describe, expect, it } from 'vitest'
import { canonicalMediaCommand, parseMediaCommand, type MediaCommand } from '../src/slideshow/media/commands'

describe('parseMediaCommand', () => {
  it('parses the whitelist; playFrom canonical; everything else undefined', () => {
    expect(parseMediaCommand('call', 'playFrom(1.50)')).toEqual({ verb: 'playFrom', seconds: 1.5 })
    expect(canonicalMediaCommand(parseMediaCommand('call', 'playFrom(1.50)')!)).toBe('playFrom(1.5)')
    expect(parseMediaCommand('call', ' playFrom( 2 ) ')).toEqual({ verb: 'playFrom', seconds: 2 })
    for (const v of ['play', 'resume', 'pause', 'stop', 'togglePause'] as const) expect(parseMediaCommand('call', v)).toEqual({ verb: v })
    expect(parseMediaCommand('call', 'playFrom(-1)')).toBeUndefined()
    expect(parseMediaCommand('call', 'playFrom(1e3)')).toBeUndefined()
    expect(parseMediaCommand('call', 'verb:1')).toBeUndefined()
    expect(parseMediaCommand('verb', '1')).toBeUndefined()
    expect(parseMediaCommand('call', 'Play')).toBeUndefined()
    expect(parseMediaCommand('evt', 'onstopaudio')).toEqual({ verb: 'onstopaudio' })
    expect(parseMediaCommand('evt', 'OnStopAudio')).toEqual({ verb: 'onstopaudio' })
    expect(parseMediaCommand('call', 'onstopaudio')).toBeUndefined()
    expect(parseMediaCommand('evt', 'play')).toBeUndefined()
  })

  it('canonical text re-parses to itself for every whitelisted command and playFrom in [0, 1800]', () => {
    const fixed: MediaCommand[] = [{ verb: 'onstopaudio' }, ...(['play', 'resume', 'pause', 'stop', 'togglePause'] as const).map((verb) => ({ verb }))]
    for (const c of fixed) {
      const text = canonicalMediaCommand(c)
      expect(parseMediaCommand(c.verb === 'onstopaudio' ? 'evt' : 'call', text)).toEqual(c)
    }
    // Seeded: integers, short decimals, full-precision doubles and sub-1e-6 values (String() gives an exponent there).
    let s = 0x9e3779b9
    const rnd = () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 2 ** 32 }
    const seconds = [0, 1, 1800, 0.5, 1799.999, 1e-7, 1.234e-7, 5e-324, 0.000001, 0.0000012345, 123.456789012345]
    for (let i = 0; i < 2000; i++) seconds.push(rnd() * 1800, Math.round(rnd() * 1800), Math.round(rnd() * 18000) / 10, rnd() * 1e-6)
    for (const n of seconds) {
      const text = canonicalMediaCommand({ verb: 'playFrom', seconds: n })
      expect(text).not.toMatch(/e/)
      const back = parseMediaCommand('call', text)
      expect(back).toEqual({ verb: 'playFrom', seconds: n })
      expect(canonicalMediaCommand(back!)).toBe(text)
    }
    expect(() => canonicalMediaCommand({ verb: 'playFrom', seconds: -1 })).toThrow(RangeError)
    expect(() => canonicalMediaCommand({ verb: 'playFrom', seconds: Infinity })).toThrow(RangeError)
  })
})
