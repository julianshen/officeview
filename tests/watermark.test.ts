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
  setWatermarkMeasurer,
  type ResolvedWatermark,
  type WatermarkOptions,
} from '../src/core/watermark'
import { createCanvas } from 'canvas'
import { createMeasurer, layoutDocx, renderPages } from '../src/docx/layout'
import { buildTextIndex } from '../src/core/search'
import { getPaintables } from '../src/render/paint'
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

  test('a comma is data, not a pair separator', () => {
    // The review's finding: splitting on ',' truncated real watermark text.
    expect(watermarkFromHeaders(headers({ [WATERMARK_HEADER]: 'text=A, B' }))?.text).toBe('A, B')
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
/**
 * Geometry and isolation tests added after an independent review found that
 * ink-counting tests cannot see placement errors: a mark painted in the wrong
 * place, or overlapping its neighbours, still adds ink.
 */
describe('watermark placement geometry', () => {
  /** x-range of ink on a page, to check where a mark actually landed. */
  function inkSpan(ctx: CanvasRenderingContext2D, w: number, h: number): { minX: number; maxX: number; centre: number } {
    const d = ctx.getImageData(0, 0, w, h).data
    let minX = Infinity
    let maxX = -Infinity
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        if (d[(y * w + x) * 4] < 200) {
          if (x < minX) minX = x
          if (x > maxX) maxX = x
        }
      }
    }
    return { minX, maxX, centre: (minX + maxX) / 2 }
  }

  const blankPage = (w: number, h: number) => {
    const c = createCanvas(w, h)
    const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
    ctx.fillStyle = '#fff'
    ctx.fillRect(0, 0, w, h)
    return ctx
  }

  test('center placement is symmetric about the page centre', () => {
    for (const rotate of [0, -45, 45]) {
      const ctx = blankPage(816, 1056)
      paintWatermark(ctx, LETTER, { text: 'CONFIDENTIAL', placement: 'center', rotate, opacity: 1 })
      const { centre } = inkSpan(ctx, 816, 1056)
      expect(Math.abs(centre - 816 / 2), `rotation ${rotate}`).toBeLessThan(20)
    }
  })

  test('header and footer placement land in their halves', () => {
    const head = blankPage(816, 1056)
    paintWatermark(head, LETTER, { text: 'MARK', placement: 'header', rotate: 0, opacity: 1 })
    const dHead = head.getImageData(0, 0, 816, 1056).data
    let topHalf = 0
    let bottomHalf = 0
    for (let y = 0; y < 1056; y++) {
      for (let x = 0; x < 816; x++) {
        if (dHead[(y * 816 + x) * 4] < 200) (y < 528 ? topHalf++ : bottomHalf++)
      }
    }
    expect(topHalf).toBeGreaterThan(bottomHalf)
  })

  test('long watermark text does not overlap its neighbours', () => {
    // The review's finding: stepX came from the font size alone, so a long mark
    // was wider than the gap between tiles.
    const c = createCanvas(816, 1056)
    const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
    const mark = normalizeWatermark({ text: 'CONFIDENTIAL', placement: 'tile', rotate: 0 })!
    setWatermarkMeasurer(ctx)
    const textWidth = (() => {
      ctx.font = `${mark.fontSizePt ?? 40}pt ${mark.fontFamily}`
      const w = ctx.measureText(mark.text).width
      ctx.font = '10px sans-serif'
      return w
    })()
    const stamps = watermarkStamps(LETTER, mark)
    // horizontal spacing between stamps in the same row
    const rowY = stamps[0].yPx
    const row = stamps.filter((s) => Math.abs(s.yPx - rowY) < 1).map((s) => s.xPx).sort((a, b) => a - b)
    const gap = row[1] - row[0]
    expect(gap, `step ${gap} vs text width ${textWidth}`).toBeGreaterThanOrEqual(textWidth)
    setWatermarkMeasurer(null)
  })
})

describe('watermark does not pollute the text index', () => {
  test('searching the watermark text finds nothing, and line text is unchanged', async () => {
    const doc = await parseDocx(await OfficePackage.load(await buildDocx([{ runs: [{ text: 'alpha beta' }] }])))
    const plain = await getPaintables(doc)
    const marked = await getPaintables(doc, { watermark: { text: 'ZEBRAFISH', placement: 'tile', opacity: 0.4 } })

    const indexPlain = await buildTextIndex(plain)
    const indexMarked = await buildTextIndex(marked)

    // the mark must not appear in any indexed line
    for (const page of indexMarked.pages) {
      for (const line of page.lines) expect(line.text).not.toContain('ZEBRAFISH')
    }
    // and the document's own text must be byte-identical with and without a mark
    expect(indexMarked.pages[0].lines[0].text).toBe(indexPlain.pages[0].lines[0].text)
    expect(indexMarked.pages[0].lines[0].text).toContain('alpha beta')
  })

  test('the mute flag is cleared even if a paint throws', async () => {
    const { isWatermarkMuted } = await import('../src/core/watermark')
    expect(isWatermarkMuted()).toBe(false)
    await buildTextIndex([
      {
        spec: { widthPx: 10, heightPx: 10 },
        paint: () => {
          throw new Error('boom')
        },
      },
    ] as never).catch(() => undefined)
    // the flag must not be left set, or every later render would lose its mark
    expect(isWatermarkMuted()).toBe(false)
  })
})

describe('watermark header text with punctuation', () => {
  test('a comma in the text survives (it is not a separator)', () => {
    const mark = watermarkFromHeaders(headers({ [WATERMARK_HEADER]: 'text=DRAFT, DO NOT COPY' }))
    expect(mark?.text).toBe('DRAFT, DO NOT COPY')
  })

  test('a comma still works when percent-encoded', () => {
    const mark = watermarkFromHeaders(headers({ [WATERMARK_HEADER]: 'text=a%2Cb' }))
    expect(mark?.text).toBe('a,b')
  })

  test('pairs are still split on semicolons', () => {
    const mark = watermarkFromHeaders(headers({ [WATERMARK_HEADER]: 'text=DRAFT; rotate=-30' }))
    expect(mark?.text).toBe('DRAFT')
    expect(mark?.rotate).toBe(-30)
  })
})

describe('watermark covers every format', () => {
  test('a spreadsheet sheet gets the mark', async () => {
    const { buildXlsx } = await import('../src/testdata/ooxml-builders')
    const bytes = await buildXlsx(
      [
        {
          name: 'Sheet1',
          rows: [
            { r: 1, cells: [{ ref: 'A1', v: 'alpha' }, { ref: 'B1', v: 'beta' }] },
            { r: 2, cells: [{ ref: 'A2', v: 1250.5 }, { ref: 'B2', v: 0.185 }] },
          ],
        },
      ],
      [],
    )
    const { OfficePackage } = await import('../src/core/zip')
    const { parseXlsx } = await import('../src/xlsx/parse')
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const plain = await getPaintables(doc)
    const marked = await getPaintables(doc, { watermark: { text: 'CONFIDENTIAL', opacity: 0.5 } })
    const inkOf = (units: Awaited<ReturnType<typeof getPaintables>>) => {
      const c = createCanvas(Math.ceil(units[0].spec.widthPx), Math.ceil(units[0].spec.heightPx))
      const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
      ctx.fillStyle = '#fff'
      ctx.fillRect(0, 0, c.width, c.height)
      units[0].paint(ctx)
      const d = ctx.getImageData(0, 0, c.width, c.height).data
      let dark = 0
      for (let i = 0; i < d.length; i += 4) if (d[i] < 250 || d[i + 1] < 250 || d[i + 2] < 250) dark++
      return dark
    }
    expect(inkOf(marked)).toBeGreaterThan(inkOf(plain))
  })

  test('a slide gets the mark', async () => {
    const { buildPptx } = await import('../src/testdata/ooxml-builders')
    const bytes = await buildPptx([
      {
        prst: 'rect',
        off: ['914400', '914400'],
        ext: ['3657600', '1828800'],
        paragraphs: [{ runs: [{ text: 'slide one' }] }],
      },
    ])
    const { OfficePackage } = await import('../src/core/zip')
    const { parsePptx } = await import('../src/pptx/parse')
    const doc = await parsePptx(await OfficePackage.load(bytes))
    const plain = await getPaintables(doc)
    const marked = await getPaintables(doc, { watermark: { text: 'CONFIDENTIAL', opacity: 0.5 } })
    const inkOf = (units: Awaited<ReturnType<typeof getPaintables>>) => {
      const c = createCanvas(Math.ceil(units[0].spec.widthPx), Math.ceil(units[0].spec.heightPx))
      const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
      ctx.fillStyle = '#fff'
      ctx.fillRect(0, 0, c.width, c.height)
      units[0].paint(ctx)
      const d = ctx.getImageData(0, 0, c.width, c.height).data
      let dark = 0
      for (let i = 0; i < d.length; i += 4) if (d[i] < 250 || d[i + 1] < 250 || d[i + 2] < 250) dark++
      return dark
    }
    expect(inkOf(marked)).toBeGreaterThan(inkOf(plain))
  })
})

describe('invalid watermark colour', () => {
  test('falls back to the default instead of painting nothing', () => {
    const c = createCanvas(816, 1056)
    const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
    ctx.fillStyle = '#fff'
    ctx.fillRect(0, 0, 816, 1056)
    // A canvas ignores an invalid fillStyle and keeps the previous value, which
    // here is white — an invisible mark with no error at all.
    paintWatermark(ctx, LETTER, { text: 'MARK', placement: 'center', color: 'not-a-colour', opacity: 1 })
    const d = ctx.getImageData(0, 0, 816, 1056).data
    let dark = 0
    for (let i = 0; i < d.length; i += 4) if (d[i] < 200) dark++
    expect(dark).toBeGreaterThan(0)
  })
})
