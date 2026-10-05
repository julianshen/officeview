/**
 * Stage 3 Task 4 — SVG selection, self-contained preflight, raster fallback and
 * inline icons across DOCX/PPTX/XLSX.
 *
 * RED-first suite: it encodes the required behaviour before the implementation.
 */
import { describe, expect, test, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { parsePptx } from '../src/pptx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import { collectDocImages, layoutDocx, createMeasurer, type MeasureFn } from '../src/docx/layout'
import { getPaintables } from '../src/render/paint'
import { decodeImage, decodeSvgImage, decodeImageAsset } from '../src/core/images'
import { buildTextIndex, findMatches } from '../src/core/search'
import { scanSelfContainedSvg, findSvgBlip, looksLikeSvg, guardSvgBytes, SVG_BLIP_EXT_URI, type ImageSelection } from '../src/core/svg'
import { parseXml, getChildren } from '../src/core/xml'
import {
  MALICIOUS_MATRIX, CANARY_ORIGIN, ICONS, PNG, ASVG_NS, svgBytes,
  distinctDocx, narrowDocx, fallbackDocx, rasterOnlyDocx, primarySvgDocx, aliasDocx, headerDocx, maliciousDocx, nestedTextboxSvgDocx, justifyInlineDocx, wrappedInlineDocx,
  distinctPptx, rasterOnlyPptx, fallbackPptx, maliciousPptx,
  distinctXlsx, rasterOnlyXlsx, fallbackXlsx, maliciousXlsx,
  buildSvgDocx, buildSvgPptx, buildSvgXlsx, buildInlineDocx,
} from '../src/testdata/svg-fixtures'

const measureFixed: MeasureFn = (text, style) => text.length * style.fontSizePt * 0.6 * (96 / 72)

async function load(bytes: Uint8Array): Promise<OfficePackage> {
  return OfficePackage.load(bytes)
}

function colorCount(ctx: CanvasRenderingContext2D, width: number, height: number, rgb: [number, number, number], tolerance = 0): number {
  const data = ctx.getImageData(0, 0, width, height).data
  let count = 0
  for (let i = 0; i < data.length; i += 4) {
    if (Math.abs(data[i] - rgb[0]) <= tolerance && Math.abs(data[i + 1] - rgb[1]) <= tolerance && Math.abs(data[i + 2] - rgb[2]) <= tolerance) count++
  }
  return count
}

function parseRgb(hex: string): [number, number, number] {
  return [parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16)]
}

/** Near-black ink count inside a rectangle (text proof, independent of colour art). */
function darkCount(ctx: CanvasRenderingContext2D, x0: number, y0: number, x1: number, y1: number): number {
  const w = Math.max(0, Math.round(x1 - x0)), h = Math.max(0, Math.round(y1 - y0))
  if (!w || !h) return 0
  const data = ctx.getImageData(Math.round(x0), Math.round(y0), w, h).data
  let count = 0
  for (let i = 0; i < data.length; i += 4) if (data[i] < 96 && data[i + 1] < 96 && data[i + 2] < 96) count++
  return count
}

interface DecodeSpy {
  calls: Array<{ mime?: string; length: number }>
  svg: number
  raster: number
  decode: (bytes: Uint8Array, mime?: string) => Promise<CanvasImageSource | undefined>
}
function decodeSpy(impl: (bytes: Uint8Array, mime?: string) => Promise<CanvasImageSource | undefined> = (b, m) => decodeImage(b, m)): DecodeSpy {
  const spy: DecodeSpy = { calls: [], svg: 0, raster: 0, decode: async (bytes, mime) => { spy.calls.push({ mime, length: bytes.length }); if (mime === 'image/svg+xml') spy.svg++; else spy.raster++; return impl(bytes, mime) } }
  return spy
}

function selections(doc: { drawingCoverage?: Array<{ imageSelection?: ImageSelection }> }): ImageSelection[] {
  return (doc.drawingCoverage ?? []).map(e => e.imageSelection).filter((s): s is ImageSelection => !!s)
}

describe('svg self-contained preflight', () => {
  test.each(MALICIOUS_MATRIX.map(row => [row.name, row.verdict, row.svg] as const))('%s is %s', (_name, verdict, svg) => {
    const result = scanSelfContainedSvg(svg)
    expect(result.ok).toBe(verdict === 'keep')
  })

  test('reports deterministic reasons for the core vectors', () => {
    const reason = (svg: Uint8Array) => { const r = scanSelfContainedSvg(svg); return r.ok ? '' : r.reason }
    expect(reason(MALICIOUS_MATRIX.find(r => r.name === 'm25_svgz')!.svg)).toBe('svgz')
    expect(reason(MALICIOUS_MATRIX.find(r => r.name === 'm17_entity_ext')!.svg)).toMatch(/doctype|entity/)
    expect(reason(MALICIOUS_MATRIX.find(r => r.name === 'm20_script')!.svg)).toMatch(/active-content:script|script/)
    expect(reason(MALICIOUS_MATRIX.find(r => r.name === 'm21_event')!.svg)).toMatch(/event-handler/)
    expect(reason(MALICIOUS_MATRIX.find(r => r.name === 'm01_image_href')!.svg)).toMatch(/external/)
    expect(reason(MALICIOUS_MATRIX.find(r => r.name === 'm11_css_import')!.svg)).toMatch(/css-import/)
    expect(reason(MALICIOUS_MATRIX.find(r => r.name === 'm12_css_url')!.svg)).toMatch(/css-url/)
    expect(reason(MALICIOUS_MATRIX.find(r => r.name === 'm23_data_svg')!.svg)).toMatch(/external|data/)
  })

  test('keeps internal fragments and paint servers', () => {
    for (const row of MALICIOUS_MATRIX.filter(r => r.verdict === 'keep')) expect(scanSelfContainedSvg(row.svg).ok).toBe(true)
  })

  test('rejects an external xml:base that would re-resolve internal fragments', () => {
    const externalBase = svgBytes(`<svg xmlns="${'http://www.w3.org/2000/svg'}" xml:base="${CANARY_ORIGIN}/"><use href="#local"/><rect width="32" height="32" fill="url(#grad)"/></svg>`)
    expect(scanSelfContainedSvg(externalBase)).toMatchObject({ ok: false, reason: 'external-base' })
    // The same document with a fragment-only base stays valid.
    const internalBase = svgBytes(`<svg xmlns="${'http://www.w3.org/2000/svg'}" xml:base="#self"><use href="#local"/></svg>`)
    expect(scanSelfContainedSvg(internalBase).ok).toBe(true)
  })

  test('rejects any data: URL and every non-fragment URL form', () => {
    for (const raw of ['data:image/svg+xml;base64,PHN2Zz4=', 'data:text/html,<b/>', 'https://x/y', '//host/y', 'file:///tmp/x', 'javascript:alert(1)', 'media/x.svg']) {
      expect(scanSelfContainedSvg(svgBytes(`<svg xmlns="${'http://www.w3.org/2000/svg'}"><image href="${raw}"/></svg>`))).toMatchObject({ ok: false })
    }
    expect(scanSelfContainedSvg(svgBytes(`<svg xmlns="${'http://www.w3.org/2000/svg'}"><image href="#frag"/><rect fill="url(#grad)"/></svg>`)).ok).toBe(true)
  })

  test('recognizes a namespace-aliased SVG root and aliased href attribute', () => {
    expect(looksLikeSvg(svgBytes(`<s:svg xmlns:s="${'http://www.w3.org/2000/svg'}" width="32" height="32"><s:rect width="32" height="32"/></s:svg>`))).toBe(true)
    // An aliased external href is still rejected; an aliased internal fragment survives.
    expect(scanSelfContainedSvg(svgBytes(`<s:svg xmlns:s="${'http://www.w3.org/2000/svg'}" xmlns:xl="http://www.w3.org/1999/xlink"><s:use xl:href="${CANARY_ORIGIN}/x"/></s:svg>`))).toMatchObject({ ok: false })
    expect(scanSelfContainedSvg(svgBytes(`<s:svg xmlns:s="${'http://www.w3.org/2000/svg'}" xmlns:xl="http://www.w3.org/1999/xlink"><s:use xl:href="#local"/></s:svg>`)).ok).toBe(true)
  })

  test('looksLikeSvg detects svg/svgz by bytes and mime, not raster', () => {
    expect(looksLikeSvg(ICONS[0].svg)).toBe(true)
    expect(looksLikeSvg(PNG.redX)).toBe(false)
    expect(looksLikeSvg(PNG.redX, 'image/svg+xml')).toBe(true)
    expect(looksLikeSvg(MALICIOUS_MATRIX.find(r => r.name === 'm25_svgz')!.svg)).toBe(true)
    expect(looksLikeSvg(undefined)).toBe(false)
  })

  test('findSvgBlip locates the standard extension by GUID regardless of prefix', () => {
    const withPrefix = (prefix: string) => parseXml(`<a:blip xmlns:a="x" xmlns:r="r" xmlns:${prefix}="${ASVG_NS}" r:embed="png0"><a:extLst><a:ext uri="${SVG_BLIP_EXT_URI}"><${prefix}:svgBlip r:embed="svg0"/></a:ext></a:extLst></a:blip>`)
    for (const prefix of ['asvg', 'svg', 'x']) {
      const blip = getChildren(withPrefix(prefix), 'blip')[0] ?? withPrefix(prefix)
      expect(findSvgBlip(blip)).toEqual({ embed: 'svg0' })
    }
    const linked = parseXml(`<blip xmlns:r="r" xmlns:asvg="${ASVG_NS}"><extLst><ext uri="${SVG_BLIP_EXT_URI}"><asvg:svgBlip r:link="svg0"/></ext></extLst></blip>`)
    expect(findSvgBlip(linked)).toEqual({ link: 'svg0' })
  })
})

describe('docx svg representation selection', () => {
  test('selects the distinct SVG art and never the PNG fallback', async () => {
    const doc = await parseDocx(await load(await distinctDocx()))
    const spy = decodeSpy()
    const paintables = await getPaintables(doc, { decodeImage: spy.decode })
    const canvas = createCanvas(paintables[0].spec.widthPx, paintables[0].spec.heightPx)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    paintables[0].paint(ctx)
    for (const icon of ICONS) {
      expect(colorCount(ctx, canvas.width, canvas.height, parseRgb(icon.svgColor))).toBeGreaterThan(5)
      expect(colorCount(ctx, canvas.width, canvas.height, parseRgb(icon.pngColor))).toBe(0)
    }
    // Body icons 1-4 are unique; the three narrow icons reuse icons 1-3, so identical
    // candidate pairs decode once (4 SVG decodes for 7 selections).
    expect(spy.svg).toBe(4)
    const sel = selections(doc)
    expect(sel).toHaveLength(7)
    expect(sel.every(s => s.phase === 'decoded' && s.representation === 'svg')).toBe(true)
    paintables.dispose()
  })

  test('keeps one image identity per svgBlip+blip pair and the source order', async () => {
    const doc = await parseDocx(await load(await distinctDocx()))
    expect(collectDocImages(doc)).toHaveLength(7)
    const text = doc.sections[0].paragraphs.map(p => p.runs.map(r => r.text).join('')).join('')
    expect(text).toContain('Before')
    expect(text).toContain('After')
    expect(text).not.toContain('  ')
    const sel = selections(doc)
    expect(sel.every(s => s.phase === 'pending')).toBe(true)
  })

  test('falls back to the raster once with a recorded reason', async () => {
    const doc = await parseDocx(await load(await fallbackDocx()))
    const spy = decodeSpy()
    const paintables = await getPaintables(doc, { decodeImage: spy.decode })
    const sel = selections(doc)
    expect(sel).toHaveLength(4)
    expect(sel[0].representation).toBe('raster')
    expect(sel[0].reason).toMatch(/missing|unresolved|relationship/)
    expect(sel[1].representation).toBe('raster')
    expect(sel[1].reason).toMatch(/decode-failed/)
    expect(sel[2].representation).toBe('raster')
    expect(sel[2].reason).toMatch(/external/)
    expect(sel[3].representation).toBe('raster')
    expect(sel[3].reason).toBeUndefined()
    const canvas = createCanvas(paintables[0].spec.widthPx, paintables[0].spec.heightPx)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    paintables[0].paint(ctx)
    expect(colorCount(ctx, canvas.width, canvas.height, parseRgb('#800080'))).toBeGreaterThan(5)
    paintables.dispose()
  })

  test('raster-only documents keep their prior behaviour', async () => {
    const doc = await parseDocx(await load(await rasterOnlyDocx()))
    const paintables = await getPaintables(doc)
    const canvas = createCanvas(paintables[0].spec.widthPx, paintables[0].spec.heightPx)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    paintables[0].paint(ctx)
    expect(colorCount(ctx, canvas.width, canvas.height, parseRgb('#800080'))).toBeGreaterThan(5)
    expect(selections(doc)[0]).toMatchObject({ phase: 'decoded', representation: 'raster' })
    paintables.dispose()
  })

  test('a bare a:blip SVG part is a first-class SVG candidate', async () => {
    const doc = await parseDocx(await load(await primarySvgDocx()))
    const spy = decodeSpy()
    const paintables = await getPaintables(doc, { decodeImage: spy.decode })
    const canvas = createCanvas(paintables[0].spec.widthPx, paintables[0].spec.heightPx)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    paintables[0].paint(ctx)
    expect(colorCount(ctx, canvas.width, canvas.height, parseRgb('#5B9BD5'))).toBeGreaterThan(5)
    expect(selections(doc)[0]).toMatchObject({ phase: 'decoded', representation: 'svg' })
    expect(spy.svg).toBe(1)
    paintables.dispose()
  })

  test('a bare malicious SVG part is preflighted before the decoder', async () => {
    const doc = await parseDocx(await load(await primarySvgDocx(MALICIOUS_MATRIX.find(r => r.name === 'm20_script')!.svg)))
    const spy = decodeSpy()
    const paintables = await getPaintables(doc, { decodeImage: spy.decode })
    expect(spy.svg).toBe(0)
    expect(selections(doc)[0].representation).toBe('none')
    expect(selections(doc)[0].reason).toMatch(/script|active-content/)
    paintables.dispose()
  })

  test('a bare SVG with an external xml:base is rejected before the decoder', async () => {
    const hostile = svgBytes(`<svg xmlns="${'http://www.w3.org/2000/svg'}" xml:base="${CANARY_ORIGIN}/"><use href="#local"/></svg>`)
    const doc = await parseDocx(await load(await primarySvgDocx(hostile)))
    const seen: number[] = []
    const stub = async (bytes: Uint8Array) => { seen.push(bytes.length); return undefined }
    const paintables = await getPaintables(doc, { decodeImage: stub })
    expect(seen).not.toContain(hostile.length)
    expect(selections(doc)[0].representation).toBe('none')
    expect(selections(doc)[0].reason).toBe('external-base')
    paintables.dispose()
  })

  test('valid SVG with a missing raster still paints the SVG', async () => {
    const doc = await parseDocx(await load(await buildValidSvgMissingRaster()))
    const paintables = await getPaintables(doc)
    const canvas = createCanvas(paintables[0].spec.widthPx, paintables[0].spec.heightPx)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    paintables[0].paint(ctx)
    expect(colorCount(ctx, canvas.width, canvas.height, parseRgb('#5B9BD5'))).toBeGreaterThan(5)
    expect(selections(doc)[0].representation).toBe('svg')
    paintables.dispose()
  })

  test('invalid SVG with a missing raster reports an explicit none', async () => {
    const doc = await parseDocx(await load(await buildInvalidSvgMissingRaster()))
    const paintables = await getPaintables(doc)
    const sel = selections(doc)[0]
    expect(sel.representation).toBe('none')
    expect(sel.reason).toMatch(/decode-failed|no-image/)
    paintables.dispose()
  })
})

describe('docx candidate-pair dedupe and aliases', () => {
  test('same PNG with different SVG does not share a bitmap; identical pairs decode once', async () => {
    const doc = await parseDocx(await load(await aliasDocx()))
    const spy = decodeSpy()
    const paintables = await getPaintables(doc, { decodeImage: spy.decode })
    expect(spy.svg).toBe(3)
    const canvas = createCanvas(paintables[0].spec.widthPx, paintables[0].spec.heightPx)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    paintables[0].paint(ctx)
    expect(colorCount(ctx, canvas.width, canvas.height, parseRgb('#5B9BD5'))).toBeGreaterThan(2)
    expect(colorCount(ctx, canvas.width, canvas.height, parseRgb('#70AD47'))).toBeGreaterThan(2)
    expect(colorCount(ctx, canvas.width, canvas.height, parseRgb('#FFC000'))).toBeGreaterThan(2)
    const sel = selections(doc)
    expect(sel.every(s => s.representation === 'svg')).toBe(true)
    paintables.dispose()
  })

  test('decode failure propagates to every alias', async () => {
    const doc = await parseDocx(await load(await aliasFailureDocx()))
    const spy = decodeSpy()
    const paintables = await getPaintables(doc, { decodeImage: spy.decode })
    const sel = selections(doc)
    expect(sel).toHaveLength(2)
    expect(sel.every(s => s.representation === 'raster' && /decode-failed/.test(s.reason ?? ''))).toBe(true)
    paintables.dispose()
  })

  test('repeat rendering retries SVG from the retained candidates after a decoder failure', async () => {
    const doc = await parseDocx(await load(await primarySvgDocx()))
    let fail = true
    const flaky = async (bytes: Uint8Array, mime?: string) => { if (fail) return undefined; return decodeImage(bytes, mime) }
    const first = await getPaintables(doc, { decodeImage: flaky })
    expect(selections(doc)[0]).toMatchObject({ representation: 'none', reason: 'svg-decode-failed' })
    first.dispose()
    fail = false
    const second = await getPaintables(doc, { decodeImage: flaky })
    expect(selections(doc)[0]).toMatchObject({ phase: 'decoded', representation: 'svg' })
    expect(selections(doc)[0].reason).toBeUndefined()
    second.dispose()
  })
})

describe('docx layout, search and coverage preservation', () => {
  test('one paragraph: Before + 4 inline icons + After keeps order and advances the text pen past the last icon', async () => {
    const doc = await parseDocx(await load(await distinctDocx()))
    expect(collectDocImages(doc)).toHaveLength(7)
    const text = doc.sections[0].paragraphs.map(p => p.runs.map(r => r.text).join('')).join('')
    expect(text).toContain('Before')
    expect(text.indexOf('After')).toBeGreaterThan(text.indexOf('Before'))
    const pages = layoutDocx(doc, measureFixed)
    const images = pages[0].images
    expect(images).toHaveLength(7)
    const body = images.slice(0, 4)
    expect(body.every(i => Math.round(i.widthPx) === 48 && Math.round(i.heightPx) === 48)).toBe(true)
    // One inline line: same baseline, strictly increasing x in source order.
    expect(new Set(body.map(i => Math.round(i.yPx))).size).toBe(1)
    const xs = body.map(i => Math.round(i.xPx))
    expect(xs).toEqual([...xs].sort((a, b) => a - b))
    // The text pen must advance past the last inline icon (image widths contribute).
    const paintables = await getPaintables(doc)
    const canvas = createCanvas(paintables[0].spec.widthPx, paintables[0].spec.heightPx)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    paintables[0].paint(ctx)
    const last = body[body.length - 1]
    const bandX = Math.round(last.xPx + last.widthPx)
    expect(darkCount(ctx, bandX, last.yPx, bandX + 80, last.yPx + last.heightPx)).toBeGreaterThan(0)
    paintables.dispose()
  })

  test('narrow one paragraph: 3x48px icons wrap within ~133px and keep source order', async () => {
    const doc = await parseDocx(await load(await narrowDocx()))
    const images = layoutDocx(doc, measureFixed)[0].images
    expect(images).toHaveLength(3)
    expect(new Set(images.map(i => Math.round(i.yPx))).size).toBeGreaterThan(1)
    const [a, b, c] = images
    expect(Math.round(a.yPx)).toBe(Math.round(b.yPx))
    expect(Math.round(c.yPx)).toBeGreaterThan(Math.round(b.yPx))
    expect(Math.round(a.xPx)).toBeLessThan(Math.round(b.xPx))
    expect(Math.round(c.xPx)).toBeLessThanOrEqual(Math.round(b.xPx))
    expect(images.map(i => i.imageIndex)).toEqual([...images.map(i => i.imageIndex)].sort((x, y) => x - y))
  })

  test('body/header/footer pictures keep distinct identities and stable indices', async () => {
    const doc = await parseDocx(await load(await headerDocx()))
    const images = collectDocImages(doc)
    expect(images).toHaveLength(3)
    const pages = layoutDocx(doc, measureFixed)
    const bodyImage = pages[0].images.find(i => !i.repeated)!
    const headerImage = pages[0].images.find(i => i.repeated === 'header')!
    expect(bodyImage.imageIndex).not.toBe(headerImage.imageIndex)
  })

  test('a nested textbox picture keeps its own identity and selects SVG', async () => {
    const doc = await parseDocx(await load(await nestedTextboxSvgDocx()))
    const images = collectDocImages(doc)
    // The textbox vector plus its nested raster/SVG picture are two identities.
    expect(images).toHaveLength(2)
    expect(images.filter(i => i.drawing)).toHaveLength(1)
    const nested = images.find(i => !i.drawing)!
    expect(nested.mime).toBe('image/png')
    const spy = decodeSpy()
    const paintables = await getPaintables(doc, { decodeImage: spy.decode })
    expect(spy.svg).toBe(1)
    expect(nested.imageSelection?.representation).toBe('svg')
    paintables.dispose()
  })
})

describe('malicious SVG cannot reach a decoder or the network', () => {
  test('25 reject rows never decode SVG; 2 positive controls do; zero network contacts', async () => {
    const contacts: string[] = []
    const originalFetch = globalThis.fetch
    globalThis.fetch = (async (input: RequestInfo | URL) => { contacts.push(`fetch:${String(input)}`); throw new Error('network blocked') }) as typeof fetch
    // Patch node HTTP entry points where the runtime allows it (bun may freeze them).
    const restore: Array<() => void> = []
    const patch = (target: Record<string, unknown>, key: string, label: string) => {
      try {
        const original = target[key]
        target[key] = (...args: unknown[]) => { void args; contacts.push(label); throw new Error('network blocked') }
        restore.push(() => { try { target[key] = original } catch { /* frozen */ } })
      } catch { /* non-configurable in this runtime */ }
    }
    const http = await import('node:http')
    const https = await import('node:https')
    patch(http as unknown as Record<string, unknown>, 'request', 'http.request')
    patch(http as unknown as Record<string, unknown>, 'get', 'http.get')
    patch(https as unknown as Record<string, unknown>, 'request', 'https.request')
    patch(https as unknown as Record<string, unknown>, 'get', 'https.get')
    try {
      const doc = await parseDocx(await load(await maliciousDocx()))
      const spy = decodeSpy()
      const paintables = await getPaintables(doc, { decodeImage: spy.decode })
      const sel = selections(doc)
      expect(sel).toHaveLength(27)
      const rejectRows = MALICIOUS_MATRIX.filter(r => r.verdict === 'reject').length
      expect(spy.svg).toBe(MALICIOUS_MATRIX.length - rejectRows)
      expect(sel.filter(s => s.representation === 'raster').length).toBe(rejectRows)
      expect(sel.filter(s => s.representation === 'svg').length).toBe(MALICIOUS_MATRIX.length - rejectRows)
      expect(contacts).toEqual([])
      paintables.dispose()
    } finally {
      globalThis.fetch = originalFetch
      for (const undo of restore) undo()
    }
  })
})

describe('pptx and xlsx adapters', () => {
  test('pptx selects SVG and falls back with reasons', async () => {
    const distinct = await parsePptx(await load(await distinctPptx()))
    const spy = decodeSpy()
    const paintables = await getPaintables(distinct, { decodeImage: spy.decode })
    expect(spy.svg).toBe(4)
    expect(selections(distinct).every(s => s.representation === 'svg')).toBe(true)
    paintables.dispose()
    const fallback = await parsePptx(await load(await fallbackPptx()))
    await getPaintables(fallback)
    const sel = selections(fallback)
    expect(sel.map(s => s.representation)).toEqual(['raster', 'raster', 'raster', 'raster'])
    expect(sel[1].reason).toMatch(/decode-failed/)
    expect(sel[2].reason).toMatch(/external/)
  })

  test('pptx and xlsx paint the selected SVG art (exact distinct colours, no fallback colour)', async () => {
    const pptx = await parsePptx(await load(await distinctPptx()))
    const pp = await getPaintables(pptx)
    const pc = createCanvas(pp[0].spec.widthPx, pp[0].spec.heightPx)
    const pctx = pc.getContext('2d') as unknown as CanvasRenderingContext2D
    pp[0].paint(pctx)
    for (const icon of ICONS) {
      expect(colorCount(pctx, pc.width, pc.height, parseRgb(icon.svgColor))).toBeGreaterThan(2)
      expect(colorCount(pctx, pc.width, pc.height, parseRgb(icon.pngColor))).toBe(0)
    }
    pp.dispose()
    const xlsx = await parseXlsx(await load(await distinctXlsx()))
    const xp = await getPaintables(xlsx)
    const xc = createCanvas(xp[0].spec.widthPx, xp[0].spec.heightPx)
    const xctx = xc.getContext('2d') as unknown as CanvasRenderingContext2D
    xp[0].paint(xctx)
    for (const icon of ICONS) {
      expect(colorCount(xctx, xc.width, xc.height, parseRgb(icon.svgColor))).toBeGreaterThan(2)
      expect(colorCount(xctx, xc.width, xc.height, parseRgb(icon.pngColor))).toBe(0)
    }
    xp.dispose()
  })

  test('pptx raster-only behaviour is unchanged', async () => {
    const doc = await parsePptx(await load(await rasterOnlyPptx()))
    const paintables = await getPaintables(doc)
    const canvas = createCanvas(paintables[0].spec.widthPx, paintables[0].spec.heightPx)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    paintables[0].paint(ctx)
    expect(colorCount(ctx, canvas.width, canvas.height, parseRgb('#800080'))).toBeGreaterThan(5)
    expect(selections(doc)[0]).toMatchObject({ representation: 'raster' })
    paintables.dispose()
  })

  test('xlsx selects SVG and falls back with reasons', async () => {
    const distinct = await parseXlsx(await load(await distinctXlsx()))
    const spy = decodeSpy()
    const paintables = await getPaintables(distinct, { decodeImage: spy.decode })
    expect(spy.svg).toBe(4)
    expect(selections(distinct).every(s => s.representation === 'svg')).toBe(true)
    paintables.dispose()
    const fallback = await parseXlsx(await load(await fallbackXlsx()))
    await getPaintables(fallback)
    const sel = selections(fallback)
    expect(sel[1].reason).toMatch(/decode-failed/)
    expect(sel[2].reason).toMatch(/external/)
  })

  test('xlsx raster-only behaviour is unchanged', async () => {
    const doc = await parseXlsx(await load(await rasterOnlyXlsx()))
    const paintables = await getPaintables(doc)
    expect(selections(doc)[0]).toMatchObject({ representation: 'raster' })
    paintables.dispose()
  })

  test('pptx/xlsx malicious rows never decode SVG', async () => {
    const pptx = await parsePptx(await load(await maliciousPptx()))
    const pptxSpy = decodeSpy()
    ;(await getPaintables(pptx, { decodeImage: pptxSpy.decode })).dispose()
    const rejectRows = MALICIOUS_MATRIX.filter(r => r.verdict === 'reject').length
    expect(pptxSpy.svg).toBe(MALICIOUS_MATRIX.length - rejectRows)
    const xlsx = await parseXlsx(await load(await maliciousXlsx()))
    const xlsxSpy = decodeSpy()
    ;(await getPaintables(xlsx, { decodeImage: xlsxSpy.decode })).dispose()
    expect(xlsxSpy.svg).toBe(MALICIOUS_MATRIX.length - rejectRows)
  })
})

describe('R1 preflight semantics', () => {
  const SVG = 'http://www.w3.org/2000/svg'
  const wrap = (inner: string) => svgBytes(`<svg xmlns="${SVG}" xmlns:s="${SVG}" width="32" height="32">${inner}</svg>`)

  test('rejects namespaced <s:style> external url and @import', () => {
    expect(scanSelfContainedSvg(wrap(`<s:style>rect{fill:url(${CANARY_ORIGIN}/a.svg#g)}</s:style><rect width="32" height="32"/>`))).toMatchObject({ ok: false, reason: 'css-url' })
    expect(scanSelfContainedSvg(wrap(`<s:style>@import '${CANARY_ORIGIN}/a.css';</s:style>`))).toMatchObject({ ok: false, reason: 'css-import' })
  })

  test('normalizes XML character references hiding CSS url in attributes and style text', () => {
    expect(scanSelfContainedSvg(wrap(`<rect width="32" height="32" style="fill:&#117;rl(${CANARY_ORIGIN}/a.svg#g)"/>`))).toMatchObject({ ok: false, reason: 'css-url' })
    expect(scanSelfContainedSvg(wrap(`<style>rect{fill:&#117;rl(${CANARY_ORIGIN}/a.svg#g)}</style>`))).toMatchObject({ ok: false, reason: 'css-url' })
  })

  test('rejects resource-affecting SMIL set/animate href mutation', () => {
    expect(scanSelfContainedSvg(wrap(`<image id="i" href="#local"/><set href="#i" attributeName="href" to="${CANARY_ORIGIN}/x"/>`))).toMatchObject({ ok: false, reason: 'smil-resource' })
    expect(scanSelfContainedSvg(wrap(`<image id="i" href="#local"/><animate href="#i" attributeName="href" values="#local;${CANARY_ORIGIN}/x"/>`))).toMatchObject({ ok: false, reason: 'smil-resource' })
    // An internal-fragment-only mutation stays valid.
    expect(scanSelfContainedSvg(wrap(`<image id="i" href="#a"/><set href="#i" attributeName="href" to="#b"/>`)).ok).toBe(true)
  })

  test('keeps internal fragments and paint servers after normalization', () => {
    expect(scanSelfContainedSvg(wrap(`<defs><linearGradient id="g"/></defs><rect fill="url(#g)"/>`)).ok).toBe(true)
    expect(scanSelfContainedSvg(wrap(`<use href="#local"/>`)).ok).toBe(true)
  })

  test('full-root detection sees an SVG behind a 2200-char comment prefix', () => {
    const hidden = svgBytes(`<!--${'x'.repeat(2200)}--><svg xmlns="${SVG}"><script>alert(1)</script></svg>`)
    expect(looksLikeSvg(hidden)).toBe(true)
    expect(guardSvgBytes(hidden)).toMatchObject({ ok: false, reason: 'active-content:script' })
  })
})

describe('R1 every decoder ingress', () => {
  const SVG = 'http://www.w3.org/2000/svg'
  const hostile = svgBytes(`<svg xmlns="${SVG}"><script>alert(1)</script></svg>`)

  test('low-level decodeImage/decodeSvgImage preflight before the browser decoder', async () => {
    const g = globalThis as { createImageBitmap?: unknown }
    const prior = g.createImageBitmap
    let calls = 0
    g.createImageBitmap = async () => { calls++; return { width: 32, height: 32 } }
    try {
      await expect(decodeImage(hostile, 'image/svg+xml')).rejects.toThrow(/svg-rejected/)
      await expect(decodeSvgImage(hostile)).rejects.toThrow(/svg-rejected/)
      expect(calls).toBe(0)
    } finally { g.createImageBitmap = prior }
  })

  test('a hostile primary a:blip never reaches the decoder when the preferred svgBlip fails', async () => {
    const safe = svgBytes(`<svg xmlns="${SVG}" width="32" height="32"><rect width="32" height="32"/></svg>`)
    const doc = await parseDocx(await load(await buildSvgDocx([{ pic: { id: 1, png: { name: 'primary.svg', data: hostile }, svg: { name: 'safe.svg', data: safe } } }])))
    const seen: Array<{ mime?: string; hostile: boolean }> = []
    const spy = async (bytes: Uint8Array, mime?: string) => { seen.push({ mime, hostile: new TextDecoder().decode(bytes).includes('<script>') }); return mime === 'image/svg+xml' ? undefined : undefined }
    const paintables = await getPaintables(doc, { decodeImage: spy })
    expect(seen.some(c => c.hostile)).toBe(false)
    paintables.dispose()
  })
})

describe('R1 candidate retention and coverage', () => {
  const SVG = 'http://www.w3.org/2000/svg'
  const script = svgBytes(`<svg xmlns="${SVG}"><script>alert(1)</script></svg>`)
  const fileHref = svgBytes(`<svg xmlns="${SVG}"><image href="file:///tmp/x"/></svg>`)
  const shared = { name: 'shared.png', data: PNG.greenSolid }
  const parse = async (format: 'docx' | 'pptx' | 'xlsx', pics: Array<{ id: number; png?: { name: string; data: Uint8Array }; svg?: { name: string; data: Uint8Array } }>) => {
    if (format === 'docx') return parseDocx(await load(await buildSvgDocx(pics.map(pic => ({ pic })))))
    if (format === 'pptx') return parsePptx(await load(await buildSvgPptx(pics)))
    return parseXlsx(await load(await buildSvgXlsx(pics)))
  }

  test('an ordinary raster sharing a rejected source PNG keeps no reason; rejected keeps its own', async () => {
    for (const format of ['docx', 'pptx', 'xlsx'] as const) {
      const doc = await parse(format, [{ id: 1, png: shared, svg: { name: 'script.svg', data: script } }, { id: 2, png: shared }])
      await getPaintables(doc)
      const sel = selections(doc)
      expect(sel[0]).toMatchObject({ phase: 'decoded', representation: 'raster', reason: 'active-content:script' })
      expect(sel[1].representation).toBe('raster')
      expect(sel[1].reason).toBeUndefined()
      expect(sel[1].phase).toBe('decoded')
    }
  })

  test('different rejected candidates sharing one PNG keep distinct reasons', async () => {
    for (const format of ['docx', 'pptx', 'xlsx'] as const) {
      const doc = await parse(format, [{ id: 1, png: shared, svg: { name: 'script.svg', data: script } }, { id: 2, png: shared, svg: { name: 'file.svg', data: fileHref } }])
      await getPaintables(doc)
      const sel = selections(doc)
      expect(sel[0].reason).toBe('active-content:script')
      expect(sel[1]).toMatchObject({ phase: 'decoded', representation: 'raster', reason: 'external-href' })
    }
  })

  test('a rejected SVG-only picture is none, never a pending raster', async () => {
    for (const format of ['docx', 'pptx', 'xlsx'] as const) {
      const doc = await parse(format, [{ id: 1, svg: { name: 'script.svg', data: script } }])
      expect(selections(doc)[0].representation).toBe('none')
      await getPaintables(doc)
      expect(selections(doc)[0]).toMatchObject({ phase: 'decoded', representation: 'none', reason: 'active-content:script' })
    }
  })
})

describe('R1 fixture portability', () => {
  test('fixture module initializes without a global Buffer (browser)', async () => {
    const g = globalThis as { Buffer?: unknown }
    const saved = g.Buffer
    vi.resetModules()
    try {
      g.Buffer = undefined
      // Module initialization runs the base64 asset decode; it must not need Buffer.
      const mod = await import('../src/testdata/svg-fixtures')
      expect(mod.PNG.redX.length).toBeGreaterThan(0)
      expect(mod.PNG.purpleSolid[0]).toBe(0x89)
    } finally {
      g.Buffer = saved
      vi.resetModules()
    }
  })
})

describe('R1 inline pen geometry', () => {
  test('exact segment and image origins: text after the last icon starts at its right edge', async () => {
    const doc = await parseDocx(await load(await distinctDocx()))
    const page = layoutDocx(doc, measureFixed)[0]
    const line = page.lines.find(l => l.segs.some(s => s.text === 'After'))!
    const after = line.segs.find(s => s.text === 'After')!
    const body = page.images.slice(0, 4)
    expect(after.penOffset).toBeDefined()
    // Fixed-measure: no source-specific page coordinates.
    expect(line.xPx + after.penOffset!).toBeCloseTo(body[3].xPx + body[3].widthPx, 5)
    for (let i = 1; i < body.length; i++) expect(body[i].xPx - body[i - 1].xPx).toBeCloseTo(48, 5)
    // Segment pen offsets are monotonic and bounded by the line width.
    let acc = -1
    for (const seg of line.segs) { expect(seg.penOffset!).toBeGreaterThanOrEqual(acc); acc = seg.penOffset! }
  })

  test('recording/search: After is indexed in source order and hit right of the last icon', async () => {
    const doc = await parseDocx(await load(await distinctDocx()))
    // Use the same real measurer as getPaintables so the expected icon edge matches the recording.
    const measure = createMeasurer(createCanvas(4, 4).getContext('2d') as unknown as CanvasRenderingContext2D)
    const last = layoutDocx(doc, measure)[0].images[3]
    const paintables = await getPaintables(doc)
    const index = await buildTextIndex(paintables as never)
    const allText = index.pages.flatMap(p => p.lines.map(l => l.text)).join('')
    expect(allText.indexOf('Before')).toBeGreaterThanOrEqual(0)
    expect(allText.indexOf('After')).toBeGreaterThan(allText.indexOf('Before'))
    const matches = findMatches(index, 'After')
    expect(matches.length).toBeGreaterThan(0)
    expect(matches[0].rects[0].x).toBeGreaterThanOrEqual(last.xPx + last.widthPx - 1)
    paintables.dispose()
  })

  test('justified tab line keeps the tab advance and text after the icon', async () => {
    const doc = await parseDocx(await load(await justifyInlineDocx()))
    const line = layoutDocx(doc, measureFixed)[0].lines.find(l => l.segs.length)!
    const image = line.inlineImages![0]
    const right = line.segs.find(s => s.text.includes('Right'))!
    expect(right.penOffset!).toBeGreaterThan(0)
    expect(line.xPx + right.penOffset!).toBeGreaterThanOrEqual(image.xPx + image.widthPx - 1)
    // The tab segment itself advanced the pen.
    expect(line.segs.some(s => s.text === ' ' && (s.penOffset ?? 0) > 0)).toBe(true)
  })

  test('wrapped continuation line restarts at the line origin', async () => {
    const doc = await parseDocx(await load(await wrappedInlineDocx()))
    const lines = layoutDocx(doc, measureFixed)[0].lines.filter(l => l.segs.length || l.inlineImages?.length)
    expect(lines.length).toBeGreaterThan(1)
    const continuation = lines[lines.length - 1]
    expect(continuation.segs.length).toBeGreaterThan(0)
    expect(continuation.segs[0].penOffset ?? 0).toBe(0)
    expect(continuation.inlineImages?.length ?? 0).toBe(0)
  })

  test('image-free lines keep the exact accumulated pen (ordinary behaviour preserved)', async () => {
    const doc = await parseDocx(await load(await buildSvgDocx([{ text: 'Hello world again' }])))
    const line = layoutDocx(doc, measureFixed)[0].lines[0]
    expect(line.inlineImages ?? []).toHaveLength(0)
    let acc = 0
    for (const seg of line.segs) { expect(seg.penOffset!).toBeCloseTo(acc, 5); acc += seg.widthPx }
  })
})

describe('R2 namespace, prefix and byte-ingress safety', () => {
  const wrap = (inner: string) => svgBytes(`<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32">${inner}</svg>`)
  const hostileForms = [
    wrap(`<image href="${CANARY_ORIGIN}/x" xlink:href="#local" xmlns:xlink="http://www.w3.org/1999/xlink"/>`),
    wrap(`<g xml:base="${CANARY_ORIGIN}/" x:base="#local" xmlns:x="urn:ignored"><use href="#local"/></g>`),
    wrap(`<s.dot:style xmlns:s.dot="http://www.w3.org/2000/svg">rect{fill:url(${CANARY_ORIGIN}/x#g)}</s.dot:style>`),
    wrap(`<s-dash:style xmlns:s-dash="http://www.w3.org/2000/svg">@import '${CANARY_ORIGIN}/x';</s-dash:style>`),
  ]

  test('colliding qualified attributes are inspected independently', () => {
    expect(scanSelfContainedSvg(hostileForms[0])).toMatchObject({ ok: false, reason: 'external-href' })
    expect(scanSelfContainedSvg(hostileForms[1])).toMatchObject({ ok: false, reason: 'external-base' })
    // Safe fragments/gradients remain positive.
    expect(scanSelfContainedSvg(wrap(`<image href="#local"/><rect fill="url(#g)"/>`)).ok).toBe(true)
  })

  test('legal dotted and hyphenated style prefixes are scanned', () => {
    expect(scanSelfContainedSvg(hostileForms[2])).toMatchObject({ ok: false, reason: 'css-url' })
    expect(scanSelfContainedSvg(hostileForms[3])).toMatchObject({ ok: false, reason: 'css-import' })
  })

  test('namespace/prefix hostiles never reach any adapter or primary decoder', async () => {
    for (const svg of hostileForms) {
      for (const format of ['docx', 'pptx', 'xlsx'] as const) {
        const pic = { id: 1, png: { name: 'fallback.png', data: PNG.greenSolid }, svg: { name: 'hostile.svg', data: svg } }
        const doc = format === 'docx' ? await parseDocx(await load(await buildSvgDocx([{ pic }]))) : format === 'pptx' ? await parsePptx(await load(await buildSvgPptx([pic]))) : await parseXlsx(await load(await buildSvgXlsx([pic])))
        const hostile: boolean[] = []
        const paintables = await getPaintables(doc, { decodeImage: async (b: Uint8Array) => { hostile.push(b.length === svg.length && b.every((v, i) => v === svg[i])); return undefined } })
        expect(hostile.some(Boolean)).toBe(false)
        paintables.dispose()
      }
      let primaryHostile = false
      await decodeImageAsset({ data: svg, imageSelection: { phase: 'pending', representation: 'none' } }, async (b: Uint8Array) => { if (b.length === svg.length && b.every((v, i) => v === svg[i])) primaryHostile = true; return undefined })
      expect(primaryHostile).toBe(false)
    }
  })

  test('UTF-16 BOM SVG fails closed at every byte ingress', () => {
    const text = '<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>'
    const utf16 = (be: boolean) => { const b = new Uint8Array(2 + text.length * 2); b.set(be ? [0xfe, 0xff] : [0xff, 0xfe]); for (let i = 0; i < text.length; i++) b[2 + i * 2 + (be ? 1 : 0)] = text.charCodeAt(i); return b }
    for (const be of [false, true]) {
      const bytes = utf16(be)
      expect(looksLikeSvg(bytes)).toBe(true)
      expect(guardSvgBytes(bytes)).toMatchObject({ ok: false, reason: 'xml-encoding-unsupported' })
      expect(scanSelfContainedSvg(bytes)).toMatchObject({ ok: false, reason: 'xml-encoding-unsupported' })
    }
  })

  test('a >1MiB comment prefix cannot hide a hostile root', () => {
    const hidden = svgBytes(`<!--${'x'.repeat((1 << 20) + 1)}--><svg xmlns="http://www.w3.org/2000/svg"><script>1</script></svg>`)
    expect(guardSvgBytes(hidden)).toMatchObject({ ok: false, reason: 'active-content:script' })
  })

  test('low-level decoders reject unsupported encodings before the fake browser decoder', async () => {
    const g = globalThis as { createImageBitmap?: unknown }
    const prior = g.createImageBitmap
    let calls = 0
    g.createImageBitmap = async () => { calls++; return { width: 32, height: 32 } }
    try {
      const text = '<svg xmlns="http://www.w3.org/2000/svg"><script>1</script></svg>'
      const utf16 = new Uint8Array(2 + text.length * 2); utf16.set([0xff, 0xfe]); for (let i = 0; i < text.length; i++) utf16[2 + i * 2] = text.charCodeAt(i)
      await expect(decodeImage(utf16)).rejects.toThrow(/svg-rejected/)
      expect(calls).toBe(0)
    } finally { g.createImageBitmap = prior }
  })

  test('normal raster bytes still pass the guard', () => {
    expect(guardSvgBytes(PNG.redX)).toEqual({ ok: true })
    expect(looksLikeSvg(PNG.redX)).toBe(false)
  })
})

describe('R2 complete-pair cache and source facts', () => {
  const shared = { name: 'shared.png', data: PNG.greenSolid }
  const safe = svgBytes('<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32"/>')
  interface CachePic { id: number; png?: { name: string; data: Uint8Array }; svg?: { name: string; data: Uint8Array; target?: string; external?: boolean } }
  const parse = async (format: 'docx' | 'pptx' | 'xlsx', pics: CachePic[]) => {
    if (format === 'docx') return parseDocx(await load(await buildSvgDocx(pics.map(pic => ({ pic })))))
    if (format === 'pptx') return parsePptx(await load(await buildSvgPptx(pics)))
    return parseXlsx(await load(await buildSvgXlsx(pics)))
  }

  test('a missing/external SVG candidate does not contaminate an ordinary raster', async () => {
    for (const format of ['docx', 'pptx', 'xlsx'] as const) {
      const variants: Array<[string, CachePic['svg'], RegExp]> = [
        ['missing', { name: 'missing.svg', data: safe, target: 'absent.svg' }, /unresolved/],
        ['external', { name: 'external.svg', data: safe, external: true, target: `${CANARY_ORIGIN}/no` }, /external/],
      ]
      for (const [name, svg, reason] of variants) {
        const doc = await parse(format, [{ id: 1, png: shared, svg }, { id: 2, png: shared }])
        await getPaintables(doc)
        const sel = selections(doc)
        expect(sel[0].reason, `${format}/${name}`).toMatch(reason)
        expect(sel[1].representation, `${format}/${name}`).toBe('raster')
        expect(sel[1].reason, `${format}/${name}`).toBeUndefined()
      }
    }
  })

  test('SVG-only and SVG+primary-SVG-fallback stay distinct complete pairs', async () => {
    for (const format of ['docx', 'pptx', 'xlsx'] as const) {
      const doc = await parse(format, [{ id: 1, svg: { name: 'same.svg', data: safe } }, { id: 2, png: { name: 'same.svg', data: safe }, svg: { name: 'same.svg', data: safe } }])
      let calls = 0
      const paintables = await getPaintables(doc, { decodeImage: async () => { calls++; return calls === 1 ? undefined : ({ width: 32, height: 32 } as unknown as CanvasImageSource) } })
      expect(calls, format).toBeGreaterThanOrEqual(2)
      expect(selections(doc)[1].representation, format).toBe('svg')
      paintables.dispose()
    }
  })

  test('a successful repeated raster render clears a stale failure reason', async () => {
    for (const format of ['docx', 'pptx', 'xlsx'] as const) {
      const doc = await parse(format, [{ id: 1, png: shared }])
      ;(await getPaintables(doc, { decodeImage: async () => undefined })).dispose()
      expect(selections(doc)[0].reason, format).toBe('raster-decode-failed')
      ;(await getPaintables(doc, { decodeImage: async () => ({ width: 32, height: 32 } as unknown as CanvasImageSource) })).dispose()
      expect(selections(doc)[0].representation, format).toBe('raster')
      expect(selections(doc)[0].reason, format).toBeUndefined()
    }
  })
})

describe('R2 non-final justification with inline images', () => {
  test('non-final justified line expands text and image with the same pen', async () => {
    const doc = await parseDocx(await load(await buildInlineDocx(
      [{ text: 'AA BB ' }, { pic: { id: 1, png: { name: 'icon.png', data: ICONS[0].png }, svg: { name: 'icon.svg', data: ICONS[0].svg } } }, { text: ' abcdefghijklmnopqrst morewords that continue to wrap' }],
      { align: 'both' },
    )))
    doc.sections[0].pageSize.widthTwips = 5880 // 200px usable content
    const line = layoutDocx(doc, measureFixed)[0].lines.find(l => l.inlineImages?.length)!
    expect(line.isParagraphEnd).toBe(false)
    expect(line.align).toBe('justify')
    const image = line.inlineImages![0]
    const [aa, sp1, bb, sp2] = line.segs
    // Independent finite origins from the fixture structure (no page coordinates).
    const rawIconOffset = aa.widthPx + sp1.widthPx + bb.widthPx + sp2.widthPx
    const unexpanded = line.segs.reduce((n, s) => n + s.widthPx, 0) + (line.inlineImages ?? []).reduce((n, i) => n + i.widthPx, 0)
    const extra = (line.contentWidthPx - unexpanded) / line.segs.filter(s => s.text === ' ').length
    const expectedIconX = line.xPx + rawIconOffset + extra * 2
    expect(image.xPx).toBeCloseTo(expectedIconX, 4)
    // Text after the two preceding gaps no longer overlaps the icon.
    expect(line.xPx + bb.penOffset! + bb.widthPx).toBeLessThanOrEqual(image.xPx + 1e-6)
    expect(bb.penOffset!).toBeGreaterThan(aa.widthPx + sp1.widthPx)
  })

  test('justified inline logical text, search and copy remain intact', async () => {
    const doc = await parseDocx(await load(await buildInlineDocx(
      [{ text: 'AA BB ' }, { pic: { id: 1, png: { name: 'icon.png', data: ICONS[0].png }, svg: { name: 'icon.svg', data: ICONS[0].svg } } }, { text: ' abcdefghijklmnopqrst morewords that continue to wrap' }],
      { align: 'both' },
    )))
    doc.sections[0].pageSize.widthTwips = 5880
    const paintables = await getPaintables(doc)
    const index = await buildTextIndex(paintables as never)
    const all = index.pages.flatMap(p => p.lines.map(l => l.text)).join('')
    expect(all).toContain('AA BB')
    expect(findMatches(index, 'AA BB').length).toBeGreaterThan(0)
    paintables.dispose()
  })
})

describe('R3 XML lexical and namespace grammar', () => {
  const SVG = 'http://www.w3.org/2000/svg'
  const wrap = (inner: string) => svgBytes(`<svg xmlns="${SVG}" width="32" height="32">${inner}</svg>`)
  const r3Hostiles: Array<[string, Uint8Array]> = [
    ['cdata-comment-boundary', svgBytes(`<svg xmlns="${SVG}"><style><![CDATA[/* <!-- */ rect{fill:url(${CANARY_ORIGIN}/x#g)} /* --> */]]></style><rect width="32" height="32"/></svg>`)],
    ['unicode-aliased-root', svgBytes(`<é:svg xmlns:é="${SVG}"><é:script>alert(1)</é:script></é:svg>`)],
    ['legal-ncname-style', svgBytes(`<svg xmlns="${SVG}"><\u200C:style xmlns:\u200C="${SVG}">@import '${CANARY_ORIGIN}/x';</\u200C:style></svg>`)],
    ['greek-prefix-style', svgBytes(`<svg xmlns="${SVG}"><α:style xmlns:α="${SVG}">rect{fill:url(${CANARY_ORIGIN}/x#g)}</α:style></svg>`)],
    ['unterminated-style', svgBytes(`<svg xmlns="${SVG}"><style>rect{fill:url(${CANARY_ORIGIN}/x#g)}`)],
    ['unterminated-cdata', svgBytes(`<svg xmlns="${SVG}"><style><![CDATA[rect{fill:url(${CANARY_ORIGIN}/x#g)}`)],
  ]

  test('XML comments are removed in markup state only; CDATA payload is scanned', () => {
    expect(scanSelfContainedSvg(r3Hostiles[0][1])).toMatchObject({ ok: false, reason: 'css-url' })
    // Ordinary fragment/gradient CDATA stays a positive control.
    expect(scanSelfContainedSvg(svgBytes(`<svg xmlns="${SVG}"><style><![CDATA[rect{fill:url(#g)}]]></style><defs><linearGradient id="g"/></defs><use href="#l"/></svg>`)).ok).toBe(true)
    // A plain XML comment mentioning a URL is still ignored.
    expect(scanSelfContainedSvg(wrap(`<!-- url(${CANARY_ORIGIN}/x) --><rect fill="url(#g)"/>`)).ok).toBe(true)
  })

  test('legal Unicode / dotted / hyphenated prefixes are recognised consistently', () => {
    expect(looksLikeSvg(r3Hostiles[1][1])).toBe(true)
    expect(scanSelfContainedSvg(r3Hostiles[1][1])).toMatchObject({ ok: false, reason: 'active-content:script' })
    for (const [, bytes] of r3Hostiles.slice(2)) expect(scanSelfContainedSvg(bytes).ok).toBe(false)
  })

  test('malformed transitions fail closed', () => {
    expect(scanSelfContainedSvg(r3Hostiles[4][1])).toMatchObject({ ok: false })
    expect(scanSelfContainedSvg(r3Hostiles[5][1])).toMatchObject({ ok: false })
  })

  test('R3 hostiles never reach any adapter, primary or public decoder', async () => {
    for (const [, svg] of r3Hostiles) {
      for (const format of ['docx', 'pptx', 'xlsx'] as const) {
        const pic = { id: 1, png: { name: 'fallback.png', data: PNG.greenSolid }, svg: { name: 'hostile.svg', data: svg } }
        const doc = format === 'docx' ? await parseDocx(await load(await buildSvgDocx([{ pic }]))) : format === 'pptx' ? await parsePptx(await load(await buildSvgPptx([pic]))) : await parseXlsx(await load(await buildSvgXlsx([pic])))
        const hostile: boolean[] = []
        const paintables = await getPaintables(doc, { decodeImage: async (b: Uint8Array) => { hostile.push(b.length === svg.length && b.every((v, i) => v === svg[i])); return undefined } })
        expect(hostile.some(Boolean)).toBe(false)
        paintables.dispose()
      }
      let primaryHostile = false
      await decodeImageAsset({ data: svg, imageSelection: { phase: 'pending', representation: 'none' } }, async (b: Uint8Array) => { if (b.length === svg.length && b.every((v, i) => v === svg[i])) primaryHostile = true; return undefined })
      expect(primaryHostile).toBe(false)
    }
  })

  test('public low-level decoder rejects an R3 hostile before the fake browser decoder', async () => {
    const g = globalThis as { createImageBitmap?: unknown }
    const prior = g.createImageBitmap
    let calls = 0
    g.createImageBitmap = async () => { calls++; return { width: 32, height: 32 } }
    try {
      await expect(decodeImage(r3Hostiles[0][1])).rejects.toThrow(/svg-rejected/)
      expect(calls).toBe(0)
    } finally { g.createImageBitmap = prior }
  })
})

describe('R4 XML tokenization boundaries', () => {
  const SVG = 'http://www.w3.org/2000/svg'
  const entityQuote = svgBytes(`<svg xmlns="${SVG}"><image title="&quot; data-x=&apos;" href="${CANARY_ORIGIN}/forbidden.svg" data-end="&apos;"/></svg>`)
  const numericQuote = svgBytes(`<svg xmlns="${SVG}"><image title="&#34; data-x=&#39;" href="file:///tmp/forbidden.svg" data-end="&#39;"/></svg>`)
  const literalCdata = svgBytes(`<svg xmlns="${SVG}"><text><![CDATA[Example <script> tag]]></text></svg>`)
  const literalCharRefs = svgBytes(`<svg xmlns="${SVG}"><text>Example &lt;script&gt; tag</text></svg>`)
  const commentCdataGradient = svgBytes(`<svg xmlns="${SVG}"><!-- <script> harmless comment --><style><![CDATA[/* <!-- */ rect{fill:url(#g)} /* --> */]]></style><defs><linearGradient id="g"/></defs><rect width="32" height="32"/></svg>`)
  const realScript = svgBytes(`<svg xmlns="${SVG}"><script>alert(1)</script></svg>`)

  test('references are decoded within attribute values, not into markup boundaries', () => {
    expect(scanSelfContainedSvg(entityQuote)).toMatchObject({ ok: false, reason: 'external-href' })
    expect(scanSelfContainedSvg(numericQuote)).toMatchObject({ ok: false, reason: 'external-href' })
    // The real external href is a separate attribute from the synthesized data-x text.
    expect(scanSelfContainedSvg(svgBytes(`<svg xmlns="${SVG}"><image title="&quot; x=&apos;" href="#local" data-end="&apos;"/></svg>`)).ok).toBe(true)
  })

  test('literal script text (raw, char-ref or CDATA) is text, not an element', () => {
    expect(scanSelfContainedSvg(literalCdata).ok).toBe(true)
    expect(scanSelfContainedSvg(literalCharRefs).ok).toBe(true)
    // A real script element is still rejected.
    expect(scanSelfContainedSvg(realScript)).toMatchObject({ ok: false, reason: 'active-content:script' })
  })

  test('comment + CDATA style payload with internal gradient stays a positive', () => {
    expect(scanSelfContainedSvg(commentCdataGradient).ok).toBe(true)
  })

  test('R4 hostiles never reach any adapter, primary or public decoder', async () => {
    for (const svg of [entityQuote, numericQuote]) {
      for (const format of ['docx', 'pptx', 'xlsx'] as const) {
        const pic = { id: 1, png: { name: 'fallback.png', data: PNG.greenSolid }, svg: { name: 'hostile.svg', data: svg } }
        const doc = format === 'docx' ? await parseDocx(await load(await buildSvgDocx([{ pic }]))) : format === 'pptx' ? await parsePptx(await load(await buildSvgPptx([pic]))) : await parseXlsx(await load(await buildSvgXlsx([pic])))
        const hostile: boolean[] = []
        const paintables = await getPaintables(doc, { decodeImage: async (b: Uint8Array) => { hostile.push(b.length === svg.length && b.every((v, i) => v === svg[i])); return undefined } })
        expect(hostile.some(Boolean)).toBe(false)
        paintables.dispose()
      }
      let primaryHostile = false
      await decodeImageAsset({ data: svg, imageSelection: { phase: 'pending', representation: 'none' } }, async (b: Uint8Array) => { if (b.length === svg.length && b.every((v, i) => v === svg[i])) primaryHostile = true; return undefined })
      expect(primaryHostile).toBe(false)
    }
  })

  test('public low-level decoder rejects an R4 hostile before the fake browser decoder', async () => {
    const g = globalThis as { createImageBitmap?: unknown }
    const prior = g.createImageBitmap
    let calls = 0
    g.createImageBitmap = async () => { calls++; return { width: 32, height: 32 } }
    try {
      await expect(decodeImage(entityQuote)).rejects.toThrow(/svg-rejected/)
      expect(calls).toBe(0)
    } finally { g.createImageBitmap = prior }
  })
})

describe('R5 encoded-XML byte classification', () => {
  const SVG = 'http://www.w3.org/2000/svg'
  const utf16 = (text: string, le: boolean, bom = false): Uint8Array => {
    const b = new Uint8Array((bom ? 2 : 0) + text.length * 2); let o = 0
    if (bom) { b[0] = le ? 0xff : 0xfe; b[1] = le ? 0xfe : 0xff; o = 2 }
    for (let i = 0; i < text.length; i++) { const c = text.charCodeAt(i); b[o + i * 2 + (le ? 0 : 1)] = c & 0xff; b[o + i * 2 + (le ? 1 : 0)] = (c >> 8) & 0xff }
    return b
  }
  const utf32 = (text: string, le: boolean, bom = false): Uint8Array => {
    const b = new Uint8Array((bom ? 4 : 0) + text.length * 4); let o = 0
    if (bom) { if (le) { b[0] = 0xff; b[1] = 0xfe } else { b[2] = 0xfe; b[3] = 0xff } o = 4 }
    for (let i = 0; i < text.length; i++) { const c = text.charCodeAt(i); const bs = le ? [c & 0xff, (c >> 8) & 0xff, (c >> 16) & 0xff, (c >> 24) & 0xff] : [(c >> 24) & 0xff, (c >> 16) & 0xff, (c >> 8) & 0xff, c & 0xff]; for (let k = 0; k < 4; k++) b[o + i * 4 + k] = bs[k] }
    return b
  }
  const decl = (enc: string) => `<?xml version="1.0" encoding="${enc}"?>`
  const hostileText = (enc: string) => `${decl(enc)}<svg xmlns="${SVG}"><image href="${CANARY_ORIGIN}/forbidden.svg"/></svg>`
  const safeText = (enc: string) => `${decl(enc)}<svg xmlns="${SVG}"><rect width="1" height="1"/></svg>`
  const reject = (bytes: Uint8Array) => {
    expect(looksLikeSvg(bytes), 'looksLike').toBe(true)
    expect(guardSvgBytes(bytes)).toMatchObject({ ok: false, reason: 'xml-encoding-unsupported' })
    expect(scanSelfContainedSvg(bytes)).toMatchObject({ ok: false, reason: 'xml-encoding-unsupported' })
  }

  test('UTF-16 LE/BE without BOM (declaration + external href) fails closed', () => {
    reject(utf16(hostileText('UTF-16LE'), true))
    reject(utf16(hostileText('UTF-16BE'), false))
    reject(utf16(safeText('UTF-16LE'), true))
    reject(utf16(safeText('UTF-16BE'), false))
  })

  test('UTF-16/UTF-32 with BOM and no-declaration roots fail closed', () => {
    reject(utf16(hostileText('UTF-16'), true, true))
    reject(utf16(hostileText('UTF-16'), false, true))
    reject(utf32(safeText('UTF-32LE'), true, true))
    reject(utf32(safeText('UTF-32BE'), false, true))
    // No declaration, no BOM: a bare encoded root/comment/whitespace prefix.
    reject(utf16(`<svg xmlns="${SVG}"><image href="${CANARY_ORIGIN}/x"/></svg>`, true))
    reject(utf16(`<!-- lead --><svg xmlns="${SVG}"><image href="${CANARY_ORIGIN}/x"/></svg>`, true))
    reject(utf16(`   \t\n<svg xmlns="${SVG}"><image href="${CANARY_ORIGIN}/x"/></svg>`, false))
  })

  test('long (>128 byte) encoded whitespace prefix and all-whitespace undecided fail closed', () => {
    reject(utf16(' '.repeat(200) + hostileText('UTF-16LE'), true))
    reject(utf16(' '.repeat(200) + `<svg xmlns="${SVG}"><image href="${CANARY_ORIGIN}/x"/></svg>`, false))
    // Entire stream is legal encoded whitespace: UNDECIDED, never raster.
    reject(utf16(' '.repeat(64), true))
  })

  test('known raster magic and legitimate custom binary are preserved', () => {
    expect(looksLikeSvg(PNG.redX)).toBe(false)
    expect(guardSvgBytes(PNG.redX)).toEqual({ ok: true })
    const arbitrary = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10])
    expect(looksLikeSvg(arbitrary)).toBe(false)
    expect(guardSvgBytes(arbitrary)).toEqual({ ok: true })
    expect(looksLikeSvg(svgBytes('   \n  '))).toBe(false)
    expect(guardSvgBytes(svgBytes('   \n  '))).toEqual({ ok: true })
    // Safe UTF-8 SVG stays accepted.
    expect(scanSelfContainedSvg(svgBytes(`<svg xmlns="${SVG}"><defs><linearGradient id="g"/></defs><rect fill="url(#g)"/></svg>`)).ok).toBe(true)
  })

  test('encoded XML never reaches any adapter, primary, fallback or public decoder', async () => {
    for (const bytes of [utf16(hostileText('UTF-16LE'), true), utf16(hostileText('UTF-16BE'), false), utf32(hostileText('UTF-32LE'), true, true)]) {
      for (const format of ['docx', 'pptx', 'xlsx'] as const) {
        const pic = { id: 1, png: { name: 'opaque.bin', data: bytes } }
        const doc = format === 'docx' ? await parseDocx(await load(await buildSvgDocx([{ pic }]))) : format === 'pptx' ? await parsePptx(await load(await buildSvgPptx([pic]))) : await parseXlsx(await load(await buildSvgXlsx([pic])))
        const hostile: boolean[] = []
        const paintables = await getPaintables(doc, { decodeImage: async (b: Uint8Array) => { hostile.push(b.length === bytes.length && b.every((v, i) => v === bytes[i])); return undefined } })
        expect(hostile.some(Boolean)).toBe(false)
        paintables.dispose()
      }
      for (const mime of [undefined, 'image/png']) {
        let primaryHostile = false
        await decodeImageAsset({ data: bytes, mime, pathHint: mime ? 'media/opaque.png' : undefined, imageSelection: { phase: 'pending', representation: 'none' } }, async (b: Uint8Array) => { if (b.length === bytes.length && b.every((v, i) => v === bytes[i])) primaryHostile = true; return undefined })
        expect(primaryHostile).toBe(false)
      }
      // Failed safe-SVG candidate followed by the hostile encoded primary fallback.
      const safe = svgBytes(`<svg xmlns="${SVG}" width="1" height="1"><rect width="1" height="1"/></svg>`)
      let fallbackHostile = false
      await decodeImageAsset({ data: bytes, svg: { bytes: safe, verdict: { ok: true } }, hasRaster: true, imageSelection: { phase: 'pending', representation: 'none' } }, async (b: Uint8Array) => { if (b.length === bytes.length && b.every((v, i) => v === bytes[i])) fallbackHostile = true; return b.length === safe.length && b.every((v, i) => v === safe[i]) ? undefined : undefined })
      expect(fallbackHostile).toBe(false)
    }
  })

  test('public low-level decoders reject encoded XML before the fake browser decoder', async () => {
    const g = globalThis as { createImageBitmap?: unknown }
    const prior = g.createImageBitmap
    let calls = 0
    g.createImageBitmap = async () => { calls++; return { width: 1, height: 1 } }
    try {
      const bytes = utf16(hostileText('UTF-16LE'), true)
      await expect(decodeImage(bytes)).rejects.toThrow(/svg-rejected/)
      await expect(decodeSvgImage(bytes)).rejects.toThrow(/svg-rejected/)
      expect(calls).toBe(0)
    } finally { g.createImageBitmap = prior }
  })
})

// Local helpers that build the two structural edge fixtures inline.

async function buildValidSvgMissingRaster(): Promise<Uint8Array> {
  const { buildSvgDocx, svgBytes } = await import('../src/testdata/svg-fixtures')
  return buildSvgDocx([{ pic: { id: 1, svg: { name: 'only.svg', data: svgBytes('<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32"><rect width="32" height="32" fill="#5B9BD5"/></svg>') } } }])
}
async function buildInvalidSvgMissingRaster(): Promise<Uint8Array> {
  const { buildSvgDocx, svgBytes } = await import('../src/testdata/svg-fixtures')
  return buildSvgDocx([{ pic: { id: 1, svg: { name: 'bad.svg', data: svgBytes('not an svg') } } }])
}
async function aliasFailureDocx(): Promise<Uint8Array> {
  const { buildSvgDocx, svgBytes } = await import('../src/testdata/svg-fixtures')
  return buildSvgDocx([
    { pic: { id: 1, png: { name: 'shared.png', data: PNG.redX }, svg: { name: 'bad.svg', data: svgBytes('garbage') } } },
    { pic: { id: 2, png: { name: 'shared.png', data: PNG.redX }, svg: { name: 'bad.svg', data: svgBytes('garbage') } } },
  ])
}

// Keep vi referenced for future spying without an unused-import lint error.
void vi
void CANARY_ORIGIN
