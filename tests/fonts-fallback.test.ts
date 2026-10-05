import { describe, expect, test } from 'vitest'
import { withFallbackFonts, REGIONAL_FALLBACKS, type FallbackFontsOptions } from '../src/core/fonts/fallback'

const identity = (family: string): string => family

describe('fallback font resolver composition', () => {
  test('no options preserves current behavior exactly', () => {
    const resolve = withFallbackFonts(identity)
    expect(resolve('Calibri')).toBe('Calibri')
    expect(resolve('Liter')).toBe('Liter')
  })
  test('explicit chain appends before the generic without quoting the stack', () => {
    const resolve = withFallbackFonts(identity, { fallbackChain: ['Inter', 'Helvetica'] })
    expect(resolve('Calibri')).toBe('"Calibri", "Inter", "Helvetica", sans-serif')
    expect(resolve('Cambria')).toBe('"Cambria", "Inter", "Helvetica", serif')
  })
  test('stack inputs splice the chain before the final generic', () => {
    const stack = (family: string): string => `"${family}", sans-serif`
    const resolve = withFallbackFonts(stack, { fallbackChain: ['Inter'] })
    expect(resolve('Liter')).toBe('"Liter", "Inter", sans-serif')
  })
  test('embedded aliases keep working with the chain appended harmlessly', () => {
    const resolve = withFallbackFonts(identity, { fallbackChain: ['Inter'] })
    expect(resolve('OfficeviewFont_1_0')).toBe('"OfficeviewFont_1_0", "Inter", sans-serif')
  })
  test('cjkFallback region heads the chain with regional families', () => {
    const opts: FallbackFontsOptions = { cjkFallback: 'sc' }
    const resolve = withFallbackFonts(identity, opts)
    const out = resolve('微軟正黑體')
    expect(out.startsWith('"微軟正黑體", "PingFang SC"')).toBe(true)
    expect(out.endsWith('sans-serif')).toBe(true)
    expect(REGIONAL_FALLBACKS.jp[0]).toBe('Hiragino Kaku Gothic ProN')
  })
  test('explicit chain wins over regional defaults when both given', () => {
    const resolve = withFallbackFonts(identity, { cjkFallback: 'sc', fallbackChain: ['MyCJK'] })
    expect(resolve('微軟正黑體')).toBe('"微軟正黑體", "MyCJK", sans-serif')
  })
  test('composed resolver drives identical paint font strings', async () => {
    const { createCanvas } = await import('canvas')
    const { paintTextBody } = await import('../src/drawing/text-paint')
    const body = { direction: 'horz' as const, paragraphs: [{ runs: [{ text: 'Hi', fontSizePt: 12, fontFamily: 'Liter' }], align: 'left' as const, level: 0 }], anchor: 't' as const, wrap: true, insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0 }
    const fonts: string[] = []
    const ctx = createCanvas(200, 50).getContext('2d')
    const proxy = new Proxy(ctx, {
      get(t, p) {
        const v = Reflect.get(t, p, t)
        return typeof v === 'function' ? (v as (...a: never[]) => unknown).bind(t) : v
      },
      set(t, p, v) {
        if (p === 'font') fonts.push(String(v))
        return Reflect.set(t, p, v)
      },
    })
    paintTextBody(body, proxy as unknown as CanvasRenderingContext2D, 0, 0, 200, 50, withFallbackFonts(identity, { fallbackChain: ['Inter'] }))
    expect(fonts.length).toBeGreaterThan(0)
    for (const f of fonts.filter(f => f.includes('Liter'))) expect(f).toBe('12pt "Liter", "Inter", sans-serif')
  })

  test('getPaintables threads the chain into docx textbox paint', async () => {
    const JSZip = (await import('jszip')).default
    const { OfficePackage } = await import('../src/core/zip')
    const { parseDocx } = await import('../src/docx/parse')
    const { getPaintables } = await import('../src/render/paint')
    const { CT_TYPES, ROOT_RELS } = await import('../src/testdata/ooxml-builders')
    const WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
    const inner = `<w:p><w:r><w:rPr><w:rFonts w:ascii="Liter" w:hAnsi="Liter"/><w:sz w:val="24"/></w:rPr><w:t>Hi</w:t></w:r></w:p>`
    const box = `<wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>`
    const zip = new JSZip()
    zip.file('[Content_Types].xml', CT_TYPES)
    zip.file('_rels/.rels', ROOT_RELS)
    zip.file('word/document.xml', `<w:document ${WNS}><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="${200 * 9525}" cy="${100 * 9525}"/><wp:docPr id="1" name="b1"/><a:graphic><a:graphicData>${box}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`)
    zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const { createCanvas } = await import('canvas')
    for (const options of [undefined, { fallbackChain: ['Inter'] }]) {
      const paintables = await getPaintables(doc, options ? { fallbackFonts: options } : undefined)
      const ctx = createCanvas(300, 200).getContext('2d')
      const fonts: string[] = []
      const proxy = new Proxy(ctx, {
        get(t, p) {
          const v = Reflect.get(t, p, t)
          return typeof v === 'function' ? (v as (...a: never[]) => unknown).bind(t) : v
        },
        set(t, p, v) {
          if (p === 'font') fonts.push(String(v))
          return Reflect.set(t, p, v)
        },
      })
      paintables[0].paint(proxy as never)
      const liter = fonts.filter(f => f.includes('Liter'))
      expect(liter.length).toBeGreaterThan(0)
      if (options) for (const f of liter) expect(f).toContain('"Inter", sans-serif')
      else for (const f of liter) expect(f).not.toContain('Inter')
      paintables.dispose()
    }
  })
})
