/**
 * S10.2 reproduced in Chrome: wrapped-text scrolling spent ~4.7 s of a 10 s
 * profile constructing a fresh Intl.Segmenter for every graphemes() call
 * (text layout appends run by run), so tile prefetch became 80 ms long tasks.
 * One segmenter per host constructor gives byte-identical clusters.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { graphemes } from '../src/core/text-recording'

type SegmenterCtor = new (locale?: string, options?: { granularity: string }) => { segment(text: string): Iterable<{ segment: string; index: number }> }
const intl = Intl as unknown as { Segmenter?: SegmenterCtor }
const original = intl.Segmenter

afterEach(() => { intl.Segmenter = original })

describe('graphemes', () => {
  it('reuses one host segmenter across calls and keeps the same clusters', () => {
    if (!original) return
    let constructed = 0
    const Base = original
    intl.Segmenter = class extends (Base as unknown as new (locale?: string, options?: { granularity: string }) => object) {
      constructor(locale?: string, options?: { granularity: string }) { super(locale, options); constructed++ }
    } as unknown as SegmenterCtor
    const samples = ['été', '👩‍👩‍👧 family', '🇹🇼🇯🇵', 'plain words wrap']
    const first = samples.map(graphemes)
    const second = samples.map(graphemes)
    expect(constructed).toBe(1)
    expect(second).toEqual(first)
    expect(first[0].map(g => g.text)).toEqual(['é', 't', 'é'])
    expect(first[2].map(g => g.text)).toEqual(['🇹🇼', '🇯🇵'])
  })

  it('printable ASCII takes a per-character fast path identical to the host segmenter', () => {
    if (!original) return
    const host = new original(undefined, { granularity: 'grapheme' })
    let ascii = ''
    for (let c = 0x20; c <= 0x7e; c++) ascii += String.fromCharCode(c)
    for (const text of [ascii, `R12C3 ${ascii.split('').reverse().join('')} lorem ipsum`, '']) {
      expect(graphemes(text)).toEqual([...host.segment(text)].map(s => ({ text: s.segment, start: s.index })))
    }
    // Not printable ASCII: CR LF is one cluster, so it still goes to the segmenter.
    expect(graphemes('a\r\nb').map(g => g.text)).toEqual(['a', '\r\n', 'b'])
  })
})
