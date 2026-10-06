import { describe, expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseTextBody } from '../src/drawing/text-parse'
import { paintTextBody } from '../src/drawing/text-paint'
import { layoutTextBody } from '../src/drawing/text-layout'
import type { DrawingTextBody } from '../src/drawing/text'

const identity = (family: string): string => family
// Mirrors the retained a11 VVVVXXXX run (54pt bold, 1pt outline, dkUpDiag
// pattern, hard offset shadow) with srgbClr colors so no theme is needed.
const WORDART_RPR =
  `<a:rPr sz="5400" b="1"><a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>` +
  `<a:pattFill prst="dkUpDiag"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr><a:bgClr><a:srgbClr val="00FF00"/></a:bgClr></a:pattFill>` +
  `<a:effectLst><a:outerShdw dist="38100" dir="2640000"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst></a:rPr>`
const txBody = (rpr: string, text: string) =>
  parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:p><a:r>${rpr}<a:t>${text}</a:t></a:r></a:p></a:txBody>`)
const bodyOf = (runs: Array<{ text: string; fontSizePt?: number; fontFamily?: string; [k: string]: unknown }>): DrawingTextBody => ({
  direction: 'horz', anchor: 't', wrap: true, insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0,
  paragraphs: [{ runs: runs.map(r => ({ fontSizePt: 24, fontFamily: 'Calibri', ...r })), align: 'left', level: 0 }],
})
function paint(body: DrawingTextBody, w = 300, h = 100) {
  const canvas = createCanvas(w, h)
  const ctx = canvas.getContext('2d')
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, w, h)
  paintTextBody(body, ctx as never, 0, 0, w, h, identity)
  return ctx.getImageData(0, 0, w, h).data
}
const inked = (data: Uint8ClampedArray): number => {
  let n = 0
  for (let i = 0; i < data.length; i += 4) if (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250) n++
  return n
}

describe('wordart run appearance parse', () => {
  test('pattern, outline and shadow land on the model with resolved colors', () => {
    const body = parseTextBody(txBody(WORDART_RPR, 'W'))
    const run = body.paragraphs[0].runs[0]
    expect(run.fontSizePt).toBe(54)
    expect(run.bold).toBe(true)
    expect(run.textFill).toMatchObject({ kind: 'pattern', preset: 'dkUpDiag', fg: '#FF0000', bg: '#00FF00' })
    expect(run.textOutline).toMatchObject({ color: '#0000FF' })
    expect(run.textOutline?.widthPx).toBeCloseTo((12700 * 96) / (12700 * 72), 6)
    expect(run.textShadow).toMatchObject({ color: '#000000', blurPx: 0 })
    // dist 38100 EMU = 4px at 44 degrees (dir 2640000/60000): exact offsets.
    expect(run.textShadow?.offsetX).toBeCloseTo(Math.cos((2640000 * Math.PI) / 10800000) * 4, 6)
    expect(run.textShadow?.offsetY).toBeCloseTo(Math.sin((2640000 * Math.PI) / 10800000) * 4, 6)
  })
  test('blur radius converts from EMU', () => {
    const rpr = `<a:rPr><a:effectLst><a:outerShdw dist="0" dir="0" bluRad="19050"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst></a:rPr>`
    const run = parseTextBody(txBody(rpr, 'B')).paragraphs[0].runs[0]
    expect(run.textShadow?.blurPx).toBe(2)
  })
  test('linear gradient stops resolve with angle', () => {
    const rpr = `<a:rPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill></a:rPr>`
    const run = parseTextBody(txBody(rpr, 'G')).paragraphs[0].runs[0]
    expect(run.textFill?.kind).toBe('gradient')
    expect(run.textFill).toMatchObject({ stops: [{ position: 0, color: '#FF0000' }, { position: 1, color: '#0000FF' }] })
    if (run.textFill?.kind === 'gradient') expect(run.textFill.angle).toBeCloseTo(Math.PI / 2, 9)
  })
  test('unsupported pattern preset and picture fill diagnose and fall back', () => {
    const rpr = `<a:rPr><a:pattFill prst="noSuchPattern"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr></a:pattFill></a:rPr>`
    const body = parseTextBody(txBody(rpr, 'P'))
    expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-appearance')).toBe(true)
    expect(body.paragraphs[0].runs[0].textFill).toBeUndefined()
    const blip = `<a:rPr><a:blipFill><a:blip r:embed="rId1"/></a:blipFill></a:rPr>`
    const blipBody = parseTextBody(txBody(blip, 'B'))
    expect(blipBody.diagnostics?.some(d => d.kind === 'unsupported-text-appearance' && d.feature === 'blipFill')).toBe(true)
    expect(blipBody.paragraphs[0].runs[0].textFill).toBeUndefined()
  })
  test('unsupported path gradient diagnoses and falls back to first stop color', () => {
    const rpr = `<a:rPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:path path="circle"/></a:gradFill></a:rPr>`
    const body = parseTextBody(txBody(rpr, 'P'))
    expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-appearance' && d.feature === 'gradFill:path')).toBe(true)
    expect(body.paragraphs[0].runs[0].textFill).toBeUndefined()
    expect(body.paragraphs[0].runs[0].color).toBe('#FF0000')
  })
  test('outline-only run keeps noFill with a stroked outline', () => {
    const rpr = `<a:rPr><a:noFill/><a:ln w="25400"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></a:rPr>`
    const run = parseTextBody(txBody(rpr, 'O')).paragraphs[0].runs[0]
    expect(run.noFill).toBe(true)
    expect(run.textOutline).toMatchObject({ color: '#00FF00' })
  })
  test('run gradFill clears style-layer noFill (most specific wins)', () => {
    const xml = parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:noFill/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></a:rPr><a:t>G</a:t></a:r></a:p></a:txBody>`)
    const run = parseTextBody(xml).paragraphs[0].runs[0]
    expect(run.noFill).toBe(false)
    expect(run.textFill?.kind).toBe('gradient')
  })
  test('appearance colors resolve through theme scheme colors', async () => {
    const { parseThemeContext } = await import('../src/drawing/style')
    const theme = parseThemeContext(
      `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>` +
      `<a:clrScheme><a:dk2><a:srgbClr val="FF0000"/></a:dk2></a:clrScheme></a:themeElements></a:theme>`)
    const rpr =
      `<a:rPr><a:pattFill prst="dkUpDiag"><a:fgClr><a:schemeClr val="tx2"/></a:fgClr>` +
      `<a:bgClr><a:schemeClr val="tx2"><a:lumMod val="50000"/></a:schemeClr></a:bgClr></a:pattFill>` +
      `<a:ln w="12700"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:ln></a:rPr>`
    const body = parseTextBody(txBody(rpr, 'T'), theme)
    const run = body.paragraphs[0].runs[0]
    expect(run.textFill).toMatchObject({ kind: 'pattern', fg: '#FF0000' })
    expect(run.textFill && run.textFill.kind === 'pattern' && run.textFill.bg).not.toBe('#FF0000')
    expect(run.textOutline).toMatchObject({ color: '#FF0000' })
  })
  test('run solidFill clears style-layer noFill (most specific wins)', () => {
    const xml = parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:noFill/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:rPr><a:t>G</a:t></a:r></a:p></a:txBody>`)
    const run = parseTextBody(xml).paragraphs[0].runs[0]
    expect(run.noFill).toBe(false)
    expect(run.color).toBe('#00FF00')
    expect(run.textFill).toBeUndefined()
  })
  test('run ln/noFill clears an inherited outline silently', () => {
    const xml = parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr><a:ln><a:noFill/></a:ln></a:rPr><a:t>N</a:t></a:r></a:p></a:txBody>`)
    const body = parseTextBody(xml)
    expect(body.paragraphs[0].runs[0].textOutline).toBeUndefined()
    expect(body.diagnostics ?? []).toEqual([])
  })
  test('width-less ln diagnoses and clears the inherited outline', () => {
    const xml = parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr><a:ln><a:prstDash val="solid"/></a:ln></a:rPr><a:t>N</a:t></a:r></a:p></a:txBody>`)
    const body = parseTextBody(xml)
    expect(body.paragraphs[0].runs[0].textOutline).toBeUndefined()
    expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-appearance' && d.feature === 'ln')).toBe(true)
  })
  test('malformed ln width diagnoses and clears instead of coercing to zero', () => {
    const xml = parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr><a:ln w="NaN"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></a:rPr><a:t>N</a:t></a:r></a:p></a:txBody>`)
    const body = parseTextBody(xml)
    expect(body.paragraphs[0].runs[0].textOutline).toBeUndefined()
    expect(body.diagnostics?.some(d => d.kind === 'unsupported-text-appearance' && d.feature === 'ln')).toBe(true)
  })
})

describe('wordart paint', () => {
  test('gradient run spans red to blue', () => {
    const data = paint(bodyOf([{ text: 'GG', textFill: { kind: 'gradient', stops: [{ position: 0, color: '#FF0000' }, { position: 1, color: '#0000FF' }], angle: 0 } }]))
    let minX = 300, maxX = 0
    for (let y = 0; y < 100; y++) {
      for (let x = 0; x < 300; x++) {
        const i = (y * 300 + x) * 4
        if (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250) {
          if (x < minX) minX = x
          if (x > maxX) maxX = x
        }
      }
    }
    expect(maxX - minX).toBeGreaterThan(10)
    const mid = (minX + maxX) / 2
    let leftR = 0, leftB = 0, rightR = 0, rightB = 0, leftN = 0, rightN = 0
    for (let y = 0; y < 100; y++) {
      for (let x = 0; x < 300; x++) {
        const i = (y * 300 + x) * 4
        if (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250) {
          if (x < mid) { leftR += data[i]; leftB += data[i + 2]; leftN++ }
          else { rightR += data[i]; rightB += data[i + 2]; rightN++ }
        }
      }
    }
    expect(leftN).toBeGreaterThan(5)
    expect(rightN).toBeGreaterThan(5)
    expect(leftR / leftN).toBeGreaterThan(rightR / rightN)
    expect(rightB / rightN).toBeGreaterThan(leftB / leftN)
  })
  test('outline-only run inks pixels with no fill', () => {
    const stroked = paint(bodyOf([{ text: 'O', noFill: true, textOutline: { color: '#00FF00', widthPx: 2 } }]))
    const filled = paint(bodyOf([{ text: 'O', color: '#FF0000' }]))
    const green = (d: Uint8ClampedArray): number => {
      let n = 0
      for (let i = 0; i < d.length; i += 4) if (d[i] < 100 && d[i + 1] > 150 && d[i + 2] < 100) n++
      return n
    }
    const red = (d: Uint8ClampedArray): number => {
      let n = 0
      for (let i = 0; i < d.length; i += 4) if (d[i] > 150 && d[i + 1] < 100 && d[i + 2] < 100) n++
      return n
    }
    // Stroked green present, fill red absent: outline with no fill.
    expect(green(stroked)).toBeGreaterThan(10)
    expect(red(stroked)).toBe(0)
    expect(red(filled)).toBeGreaterThan(10)
  })
  test('shadow displaces ink from the unshadowed position', () => {
    const bbox = (data: Uint8ClampedArray): [number, number, number, number] => {
      let l = 300, t = 100, r = 0, b = 0
      for (let y = 0; y < 100; y++) {
        for (let x = 0; x < 300; x++) {
          const i = (y * 300 + x) * 4
          if (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250) {
            if (x < l) l = x
            if (x > r) r = x
            if (y < t) t = y
            if (y > b) b = y
          }
        }
      }
      return [l, t, r, b]
    }
    const plain = bbox(paint(bodyOf([{ text: 'S' }])))
    const shadowed = bbox(paint(bodyOf([{ text: 'S', textShadow: { color: '#000000', blurPx: 0, offsetX: 3, offsetY: 3 } }])))
    expect(shadowed[0]).toBe(plain[0])
    expect(shadowed[1]).toBe(plain[1])
    expect(shadowed[2]).toBe(plain[2] + 3)
    expect(shadowed[3]).toBe(plain[3] + 3)
  })
  test('shadow state never leaks into the following run', async () => {
    const canvas = createCanvas(300, 100)
    const ctx = canvas.getContext('2d')
    const shadows: string[] = []
    const proxy = new Proxy(ctx, {
      get(t, p) {
        const v = Reflect.get(t, p, t)
        return typeof v === 'function' ? (v as (...a: never[]) => unknown).bind(t) : v
      },
      set(t, p, v) {
        if (p === 'shadowColor') shadows.push(String(v))
        return Reflect.set(t, p, v)
      },
    })
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 300, 100)
    paintTextBody(
      bodyOf([
        { text: 'S', textShadow: { color: '#000000', blurPx: 0, offsetX: 6, offsetY: 6 } },
        { text: 'P' },
      ]),
      proxy as never, 0, 0, 300, 100, identity,
    )
    expect(shadows.length).toBeGreaterThan(0)
    expect(shadows).toContain('#000000')
    // The plain run resets shadow state: nothing after it carries a shadow.
    expect(shadows[shadows.length - 1]).not.toBe('#000000')
  })
  test('bullet after a shadowed run paints with no shadow', async () => {
    const canvas = createCanvas(300, 100)
    const ctx = canvas.getContext('2d')
    const shadows: string[] = []
    const proxy = new Proxy(ctx, {
      get(t, p) {
        const v = Reflect.get(t, p, t)
        return typeof v === 'function' ? (v as (...a: never[]) => unknown).bind(t) : v
      },
      set(t, p, v) {
        if (p === 'shadowColor') shadows.push(String(v))
        return Reflect.set(t, p, v)
      },
    })
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 300, 100)
    const bulletBody: DrawingTextBody = {
      direction: 'horz', anchor: 't', wrap: true, insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0,
      paragraphs: [
        { runs: [{ text: 'S', fontSizePt: 24, fontFamily: 'Calibri', textShadow: { color: '#000000', blurPx: 0, offsetX: 6, offsetY: 6 } }], align: 'left', level: 0 },
        { runs: [{ text: 'B', fontSizePt: 24, fontFamily: 'Calibri' }], align: 'left', level: 0, bullet: true, bulletCharacter: '•' },
      ],
    }
    paintTextBody(bulletBody, proxy as never, 0, 0, 300, 100, identity)
    expect(shadows).toContain('#000000')
    expect(shadows[shadows.length - 1]).not.toBe('#000000')
  })
  test('diagonal pattern tile draws both colors', async () => {
    const { paintPatternTile } = await import('../src/drawing/text-paint')
    const canvas = createCanvas(16, 16)
    const ctx = canvas.getContext('2d')
    paintPatternTile(ctx as never, 'dkUpDiag', '#FF0000', '#00FF00', 16)
    const data = ctx.getImageData(0, 0, 16, 16).data
    let red = 0, green = 0
    for (let i = 0; i < data.length; i += 4) {
      if (data[i] > 200 && data[i + 1] < 100 && data[i + 2] < 100) red++
      if (data[i] < 100 && data[i + 1] > 200 && data[i + 2] < 100) green++
    }
    expect(red).toBeGreaterThan(10)
    expect(green).toBeGreaterThan(10)
  })
  test('pattern run paints (fg fallback where tiles are unavailable)', () => {
    const data = paint(bodyOf([{ text: 'P', textFill: { kind: 'pattern', preset: 'dkUpDiag', fg: '#FF0000', bg: '#00FF00' } }]))
    expect(inked(data)).toBeGreaterThan(10)
  })
  test('appearance runs keep exact search text with no extra records', async () => {
    const { buildTextIndex, findMatches } = await import('../src/core/search')
    const body = bodyOf([{ text: 'Hi', textFill: { kind: 'gradient', stops: [{ position: 0, color: '#FF0000' }, { position: 1, color: '#0000FF' }], angle: 0 }, textOutline: { color: '#000000', widthPx: 1 }, textShadow: { color: '#000000', blurPx: 0, offsetX: 2, offsetY: 2 } }])
    const index = await buildTextIndex([{ spec: { widthPx: 300, heightPx: 100 }, paint: (ctx) => paintTextBody(body, ctx as never, 0, 0, 300, 100, identity) }])
    expect(index.pages[0].lines.map(l => l.text).join('')).toBe('Hi')
    expect(findMatches(index, 'Hi').length).toBe(1)
  })
  test('appearance never changes layout advances', () => {
    const styled = bodyOf([{ text: 'Hi', textFill: { kind: 'gradient', stops: [{ position: 0, color: '#FF0000' }, { position: 1, color: '#0000FF' }], angle: 0 }, textOutline: { color: '#000000', widthPx: 2 }, textShadow: { color: '#000000', blurPx: 0, offsetX: 2, offsetY: 2 } }])
    const plain = bodyOf([{ text: 'Hi' }])
    const measure = (_text: string) => ({ width: 10, ascent: 8, descent: 2 })
    const a = layoutTextBody(styled, 300, 100, measure as never)
    const b = layoutTextBody(plain, 300, 100, measure as never)
    expect(a.lines[0].segments.map(s => s.width)).toEqual(b.lines[0].segments.map(s => s.width))
  })
  test('run with both noFill and textFill suppresses fill ink (noFill wins)', () => {
    const data = paint(bodyOf([{ text: 'X', noFill: true, textFill: { kind: 'pattern', preset: 'dkUpDiag', fg: '#FF0000', bg: '#00FF00' } }]))
    expect(inked(data)).toBe(0)
  })
})
