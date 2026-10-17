import { describe, expect, test } from 'vitest'
import { createLocalTextMeasurer } from '../src/core/text-metrics'

/** A host may serialize effective zero as empty while rejecting empty setters. */
test('cached local metrics reset both spacing properties when caller getters serialize zero as empty', () => {
  const contexts: FakeContext[] = []
  class FakeCanvas {
    width = 4
    height = 4
    ownerDocument = { createElement: (_tag: string) => new FakeCanvas() }
    constructor(_width = 4, _height = 4) {}
    getContext(_type: string) {
      const context = new FakeContext(this)
      contexts.push(context)
      return context
    }
  }
  class FakeContext {
    font = '14px Arial'
    lang = 'en'
    direction = 'ltr'
    fontKerning = 'auto'
    fontStretch = 'normal'
    fontVariantCaps = 'normal'
    textRendering = 'auto'
    zeroSerialization = ''
    private letter = 0
    private word = 0
    constructor(readonly canvas: FakeCanvas) {}
    get letterSpacing() { return this.letter ? `${this.letter}px` : this.zeroSerialization }
    set letterSpacing(value: string) { if (/^-?\d+(?:\.\d+)?px$/.test(value)) this.letter = parseFloat(value) }
    get wordSpacing() { return this.word ? `${this.word}px` : this.zeroSerialization }
    set wordSpacing(value: string) { if (/^-?\d+(?:\.\d+)?px$/.test(value)) this.word = parseFloat(value) }
    measureText(text: string) { return { width: text.length * (10 + this.letter) + (text.match(/ /g)?.length ?? 0) * this.word } as TextMetrics }
  }
  const base = new FakeContext(new FakeCanvas())
  const local = createLocalTextMeasurer(base as unknown as CanvasRenderingContext2D)
  base.letterSpacing = '1.4px'
  base.wordSpacing = '2px'
  expect(local('a b').width).toBe(36.2)
  base.letterSpacing = '0px'
  base.wordSpacing = '0px'
  expect(base.letterSpacing).toBe('')
  expect(base.wordSpacing).toBe('')
  expect(local('a b').width).toBe(base.measureText('a b').width)
  base.zeroSerialization = 'normal'
  base.letterSpacing = '1.4px'; base.wordSpacing = '2px'
  expect(local('a b').width).toBe(36.2)
  base.letterSpacing = '0px'; base.wordSpacing = '0px'
  base.lang = 'ja'; base.direction = 'rtl'; base.fontKerning = 'none'
  expect(local('a b').width).toBe(30)
  expect(contexts[0]).toMatchObject({ lang: 'ja', direction: 'rtl', fontKerning: 'none' })
  expect(contexts).toHaveLength(1)
  expect(base.font).toBe('14px Arial')
})

describe('visual review #16/NEW-3: measuring canvases follow the painting language', () => {
  test('inheritedCanvasLang resolves the nearest lang, the document lang, or nothing', async () => {
    const { inheritedCanvasLang } = await import('../src/core/text-metrics')
    const html = document.documentElement
    const before = html.getAttribute('lang')
    try {
      html.setAttribute('lang', 'en')
      expect(inheritedCanvasLang()).toBe('en')
      const host = document.createElement('div')
      host.setAttribute('lang', 'ja')
      const canvas = document.createElement('canvas')
      host.appendChild(canvas)
      document.body.appendChild(host)
      expect(inheritedCanvasLang(canvas)).toBe('ja')
      host.remove()
      // A detached canvas inherits nothing: leave its context at "inherit".
      expect(inheritedCanvasLang(document.createElement('canvas'))).toBeNull()
      expect(inheritedCanvasLang(null)).toBeNull()
      html.removeAttribute('lang')
      expect(inheritedCanvasLang()).toBeNull()
    } finally {
      if (before === null) html.removeAttribute('lang')
      else html.setAttribute('lang', before)
    }
  })
})
