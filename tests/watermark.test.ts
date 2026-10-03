import { describe, test, expect } from 'vitest'
import {
  WATERMARK_HEADER,
  WATERMARK_DEFAULTS,
  normalizeWatermark,
  mergeWatermark,
  watermarkFromHeaders,
  watermarkStamps,
  fontSizePx,
  paintWatermark,
  type ResolvedWatermark,
  type WatermarkOptions,
} from '../src/core/watermark'
import { createCanvas } from 'canvas'
import { createMeasurer, layoutDocx, renderPages } from '../src/docx/layout'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { buildDocx } from '../src/testdata/ooxml-builders'

const LETTER = { widthPx: 816, heightPx: 1056 }
const headers = (init: Record<string, string>) => new Headers(init)

describe('normalizeWatermark', () => {
  test('fills in defaults for a bare text', () => {
    const mark = normalizeWatermark({ text: 'DRAFT' })
    expect(mark).toBeDefined()
    expect(mark!.text).toBe('DRAFT')
    expect(mark!.placement).toBe(WATERMARK_DEFAULTS.placement)
    expect(mark!.rotate).toBe(WATERMARK_DEFAULTS.rotate)
    expect(mark!.opacity).toBe(WATERMARK_DEFAULTS.opacity)
    expect(mark!.color).toBe(WATERMARK_DEFAULTS.color)
  })

  test('trims the text and disables a blank watermark', () => {
    expect(normalizeWatermark({ text: '  spaced  ' })?.text).toBe('spaced')
    expect(normalizeWatermark({ text: '' })).toBeUndefined()
    expect(normalizeWatermark({ text: '   ' })).toBeUndefined()
    expect(normalizeWatermark(undefined)).toBeUndefined()
    expect(normalizeWatermark(null)).toBeUndefined()
  })

  test('clamps opacity into 0..1 rather than trusting it', () => {
    expect(normalizeWatermark({ text: 'x', opacity: 5 })?.opacity).toBe(1)
    expect(normalizeWatermark({ text: 'x', opacity: -3 })?.opacity).toBe(0)
    expect(normalizeWatermark({ text: 'x', opacity: Number.NaN })?.opacity).toBe(WATERMARK_DEFAULTS.opacity)
  })

  test('an unknown placement falls back to the default', () => {
    expect(normalizeWatermark({ text: 'x', placement: 'sideways' as never })?.placement).toBe('tile')
    expect(normalizeWatermark({ text: 'x', placement: 'center' })?.placement).toBe('center')
  })

  test('a non-numeric rotate falls back instead of becoming NaN', () => {
    expect(normalizeWatermark({ text: 'x', rotate: Number.NaN })?.rotate).toBe(WATERMARK_DEFAULTS.rotate)
    expect(normalizeWatermark({ text: 'x', rotate: 90 })?.rotate).toBe(90)
  })
})

describe('mergeWatermark', () => {
  test('the header wins over the prop', () => {
    const merged = mergeWatermark(
      { text: 'FROM-PROP' },
      { text: 'FROM-HEADER', placement: 'footer' },
    )
    expect(merged?.text).toBe('FROM-HEADER')
    expect(merged?.placement).toBe('footer')
  })

  test('the prop is used when there is no header', () => {
    expect(mergeWatermark({ text: 'FROM-PROP' }, undefined)?.text).toBe('FROM-PROP')
    expect(mergeWatermark(undefined, undefined)).toBeUndefined()
  })
})

describe('watermarkFromHeaders', () => {
  test('parses the documented token list', () => {
    const mark = watermarkFromHeaders(
      headers({ [WATERMARK_HEADER]: 'text=DRAFT; rotate=-45; opacity=0.15' }),
    )
    expect(mark).toBeDefined()
    expect(mark!.text).toBe('DRAFT')
    expect(mark!.rotate).toBe(-45)
    expect(mark!.opacity).toBe(0.15)
    expect(mark!.placement).toBeUndefined()
  })

  test('accepts a comma separator, like the protection header', () => {
    expect(watermarkFromHeaders(headers({ [WATERMARK_HEADER]: 'text=A, text=B' }))?.text).toBe('B')
  })

  test('percent-decodes non-ASCII text', () => {
    expect(watermarkFromHeaders(headers({ [WATERMARK_HEADER]: 'text=%D0%9F%D0%BE%D0%B2%D0%B5%D1%80' }))?.text).toBe(
      'Повер',
    )
  })

  test('text with spaces survives because only the key=value split is special', () => {
    expect(watermarkFromHeaders(headers({ [WATERMARK_HEADER]: 'text=HIGH CONFIDENTIAL' }))?.text).toBe(
      'HIGH CONFIDENTIAL',
    )
  })

  test('absent header means no watermark', () => {
    expect(watermarkFromHeaders(headers({}))).toBeUndefined()
  })

  test('a header without text is ignored', () => {
    expect(watermarkFromHeaders(headers({ [WATERMARK_HEADER]: 'rotate=-45' }))).toBeUndefined()
    expect(watermarkFromHeaders(headers({ [WATERMARK_HEADER]: 'text=   ' }))).toBeUndefined()
  })

  test('malformed input never throws', () => {
    // A bad header must not be able to break a document load.
    expect(() => watermarkFromHeaders(headers({ [WATERMARK_HEADER]: '%%%' }))).not.toThrow()
    expect(() => watermarkFromHeaders(headers({ [WATERMARK_HEADER]: 'text=%E0%A4%A' }))).not.toThrow()
    expect(() => watermarkFromHeaders(headers({ [WATERMARK_HEADER]: ';;;===' }))).not.toThrow()
    expect(() => watermarkFromHeaders(headers({ [WATERMARK_HEADER]: 'text=A; opacity=abc' }))).not.toThrow()
  })

  test('a truncated percent escape keeps the text instead of throwing', () => {
    expect(watermarkFromHeaders(headers({ [WATERMARK_HEADER]: 'text=100%A5' }))?.text).toBe('100%A5')
  })

  test('a non-numeric opacity is dropped, not turned into NaN', () => {
    const mark = watermarkFromHeaders(headers({ [WATERMARK_HEADER]: 'text=X; opacity=abc' }))
    expect(mark?.text).toBe('X')
    expect(mark?.opacity).toBeUndefined()
    expect(normalizeWatermark(mark)?.opacity).toBe(WATERMARK_DEFAULTS.opacity)
  })
})

describe('watermark geometry', () => {
  test('center places one mark at the middle', () => {
    const marks = watermarkStamps(LETTER, normalizeWatermark({ text: 'X', placement: 'center' })!)
    expect(marks).toHaveLength(1)
    expect(marks[0]).toEqual({ xPx: LETTER.widthPx / 2, yPx: LETTER.heightPx / 2 })
  })

  test('header sits near the top and footer near the bottom', () => {
    const head = watermarkStamps(LETTER, normalizeWatermark({ text: 'X', placement: 'header' })!)
    const foot = watermarkStamps(LETTER, normalizeWatermark({ text: 'X', placement: 'footer' })!)
    expect(head[0].yPx).toBeLessThan(LETTER.heightPx / 4)
    expect(foot[0].yPx).toBeGreaterThan((LETTER.heightPx * 3) / 4)
    expect(head[0].xPx).toBe(foot[0].xPx)
  })

  test('tile covers the page rather than drawing a single mark', () => {
    const marks = watermarkStamps(LETTER, normalizeWatermark({ text: 'CONFIDENTIAL', placement: 'tile' })!)
    expect(marks.length).toBeGreaterThan(4)
    // tiles must reach the page area, not just hug one corner
    expect(marks.some((m) => m.yPx > LETTER.heightPx / 2)).toBe(true)
    expect(marks.some((m) => m.yPx < LETTER.heightPx / 2)).toBe(true)
  })

  test('a larger font produces fewer, bigger tiles', () => {
    // Inverse relationship: bigger marks mean fewer of them fit on the page.
    const small = watermarkStamps(LETTER, normalizeWatermark({ text: 'X', fontSizePt: 12 })!)
    const large = watermarkStamps(LETTER, normalizeWatermark({ text: 'X', fontSizePt: 72 })!)
    expect(small.length).toBeGreaterThan(large.length)
    expect(large.length).toBeGreaterThan(0)
  })

  test('font size scales with the page when not given', () => {
    const letter = fontSizePx(LETTER, normalizeWatermark({ text: 'X' })!)
    const a4 = fontSizePx({ widthPx: 794, heightPx: 1123 }, normalizeWatermark({ text: 'X' })!)
    // close, because the sizes are close, but derived from the page not fixed
    expect(Math.abs(letter - a4)).toBeLessThan(5)
    expect(letter).toBeGreaterThan(12)
  })

  test('an explicit pt size wins over the page-derived one', () => {
    expect(fontSizePx(LETTER, normalizeWatermark({ text: 'X', fontSizePt: 36 })!)).toBe(48) // 36pt @96dpi
  })
})

  /** Paint a document and count non-white pixels, which is the only assertion
   *  that proves the watermark actually reached the canvas. */
  async function inkWithWatermark(
    watermark?: WatermarkOptions | ResolvedWatermark,
  ): Promise<number> {
    const doc = await parseDocx(await OfficePackage.load(await buildDocx([{ runs: [{ text: 'body' }] }])))
    const measure = createMeasurer(createCanvas(10, 10).getContext('2d') as unknown as CanvasRenderingContext2D)
    const pages = layoutDocx(doc, measure)
    const c = createCanvas(Math.ceil(pages[0].widthPx), Math.ceil(pages[0].heightPx))
    const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, c.width, c.height)
    renderPages(pages, ctx, undefined, watermark ? { watermark } : undefined)
    const data = ctx.getImageData(0, 0, c.width, c.height).data
    let dark = 0
    for (let i = 0; i < data.length; i += 4) if (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250) dark++
    return dark
  }
describe('watermark reaches the canvas', () => {
  test('a watermark adds ink that the document alone does not have', async () => {
    const plain = await inkWithWatermark()
    const marked = await inkWithWatermark({ text: 'CONFIDENTIAL', placement: 'tile', opacity: 0.4 })
    expect(marked).toBeGreaterThan(plain)
  })

  test('every placement paints something', async () => {
    const plain = await inkWithWatermark()
    for (const placement of ['center', 'tile', 'header', 'footer'] as const) {
      const ink = await inkWithWatermark({ text: 'MARK', placement, opacity: 0.5 })
      expect(ink, `${placement} should add ink`).toBeGreaterThan(plain)
    }
  })

  test('an empty text paints nothing extra', async () => {
    const plain = await inkWithWatermark()
    const blank = await inkWithWatermark({ text: '   ' })
    expect(blank).toBe(plain)
  })
})

describe('paintWatermark', () => {
  const ctxOf = () => {
    const c = createCanvas(816, 1056)
    const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
    ctx.fillStyle = '#fff'
    ctx.fillRect(0, 0, c.width, c.height)
    return { c, ctx }
  }
  const darkCount = (c: ReturnType<typeof createCanvas>) => {
    const d = (c.getContext('2d') as unknown as { getImageData: (a: number, b: number, w: number, h: number) => { data: Uint8ClampedArray } })
      .getImageData(0, 0, c.width, c.height).data
    let n = 0
    for (let i = 0; i < d.length; i += 4) if (d[i] < 250) n++
    return n
  }

  test('paints ink and leaves the context state restored', () => {
    const { c, ctx } = ctxOf()
    const before = { alpha: ctx.globalAlpha, fill: ctx.fillStyle, font: ctx.font }
    paintWatermark(ctx, LETTER, { text: 'DRAFT', placement: 'center' })
    expect(darkCount(c)).toBeGreaterThan(0)
    // the painter must not leak alpha/font/fill to whatever draws next
    expect(ctx.globalAlpha).toBe(before.alpha)
    expect(ctx.fillStyle).toBe(before.fill)
    expect(ctx.font).toBe(before.font)
  })

  test('a zero-opacity watermark is invisible', () => {
    const visible = ctxOf()
    paintWatermark(visible.ctx, LETTER, { text: 'DRAFT', placement: 'center', opacity: 0.5 })
    const hidden = ctxOf()
    paintWatermark(hidden.ctx, LETTER, { text: 'DRAFT', placement: 'center', opacity: 0 })
    expect(darkCount(hidden.c)).toBeLessThan(darkCount(visible.c))
  })

  test('does not throw on unusable input', () => {
    const { ctx } = ctxOf()
    expect(() => paintWatermark(ctx, LETTER, undefined)).not.toThrow()
    expect(() => paintWatermark(ctx, LETTER, { text: '' })).not.toThrow()
  })
})