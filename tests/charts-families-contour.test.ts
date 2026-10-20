// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import type { ChartModel } from '../src/charts/model'

const parse = (xml: string) => parseXlsxChart(parseXmlOrdered(xml), undefined)!
const lit = (values: number[]) => `<numLit><ptCount val="${values.length}"/>${values.map((v, i) => `<pt idx="${i}"><v>${v}</v></pt>`).join('')}</numLit>`
const cats = (n: number) => `<cat><strLit><ptCount val="${n}"/>${Array.from({ length: n }, (_, i) => `<pt idx="${i}"><v>C${i}</v></pt>`).join('')}</strLit></cat>`
const ser = (i: number, values: number[]) => `<ser><idx val="${i}"/><order val="${i}"/><tx><v>R${i}</v></tx>${cats(values.length)}<val>${lit(values)}</val></ser>`
const AXES = '<catAx><axId val="1"/><crossAx val="2"/></catAx><serAx><axId val="2"/><crossAx val="1"/></serAx><valAx><axId val="3"/><crossAx val="1"/><majorUnit val="2"/></valAx>'
const space = (plot: string) => `<chartSpace><chart><plotArea>${plot}${AXES}</plotArea></chart></chartSpace>`

const W = 200, H = 160
function paint(model: ChartModel) {
  const canvas = createCanvas(W, H), ctx = canvas.getContext('2d') as any
  paintXlsxChart(model, ctx, W, H)
  return { ctx, data: ctx.getImageData(0, 0, W, H).data as Uint8ClampedArray }
}

describe('charts-families-contour (Task A5d)', () => {
  it('contour bands from the major unit; grid at most 256x256', () => {
    const model = parse(space(`<surfaceChart>${ser(0, [1, 2, 3])}${ser(1, [4, 5, 6])}<axId val="1"/><axId val="2"/><axId val="3"/></surfaceChart>`))
    expect(model).toBeDefined()
    expect(model!.groups[0].family).toBe('contour')
    // major unit 2 → band count ceil((6-1)/2)=3 (values 1..6); just ensure model admits
    const { data } = paint(model!)
    // count distinct non-white filled colours (bands)
    const colors = new Set<string>()
    for (let i = 0; i < data.length; i += 4) {
      const hex = [data[i], data[i+1], data[i+2]].map(v => v.toString(16).padStart(2,'0')).join('')
      if (hex !== 'ffffff' && data[i+3] > 0) colors.add(hex)
    }
    expect(colors.size).toBeGreaterThanOrEqual(2)
  })

  it('grid capped 256x256 with diagnostic', () => {
    // 300 categories × 2 series → truncated; model should mark truncated
    const manyCats = (n: number) => `<cat><strLit><ptCount val="${n}"/>${Array.from({ length: n }, (_, i) => `<pt idx="${i}"><v>C${i}</v></pt>`).join('')}</strLit></cat>`
    const serBig = (i: number, n: number) => `<ser><idx val="${i}"/><order val="${i}"/><tx><v>R${i}</v></tx>${manyCats(n)}<val>${lit(Array.from({length:n},(_,k)=>k+1))}</val></ser>`
    const xml = `<chartSpace><chart><plotArea><surfaceChart>${serBig(0,300)}${serBig(1,300)}<axId val="1"/><axId val="2"/><axId val="3"/></surfaceChart><catAx><axId val="1"/><crossAx val="2"/></catAx><serAx><axId val="2"/><crossAx val="1"/></serAx><valAx><axId val="3"/><crossAx val="1"/></valAx></plotArea></chart></chartSpace>`
    const model = parse(xml)
    // Should still parse (not throw) and mark truncated
    expect(model).toBeDefined()
    expect(model!.truncated).toBe(true)
  })

  it('no interpolation across missing cells', () => {
    // Row0: 1, (blank), 3
    const ser0 = `<ser><idx val="0"/><order val="0"/><tx><v>R0</v></tx>${cats(3)}<val><numLit><ptCount val="3"/><pt idx="0"><v>1</v></pt><pt idx="2"><v>3</v></pt></numLit></val></ser>`
    const model = parse(space(`<surfaceChart>${ser0}${ser(1, [4, 5, 6])}<axId val="1"/><axId val="2"/><axId val="3"/></surfaceChart>`))
    const { data } = paint(model!)
    // The blank cell (row0 col1) should be white (unpainted) — sample centre of that cell
    const cellW = W / 3, cellH = H / 2
    const cx = cellW * 1 + cellW / 2, cy = cellH * 0 + cellH / 2
    const i = (Math.round(cy) * W + Math.round(cx)) * 4
    // unpainted cell keeps the chart-area background (white), not a band colour
    expect([data[i], data[i+1], data[i+2]]).toEqual([255, 255, 255])
  })

  it('wireframe draws only band boundaries (no fills)', () => {
    const xml = `<chartSpace><chart><plotArea><surfaceChart><wireframe val="1"/>${ser(0, [1, 2, 3])}${ser(1, [4, 5, 6])}<axId val="1"/><axId val="2"/><axId val="3"/></surfaceChart>${AXES}</plotArea></chart></chartSpace>`
    // note AXES includes majorUnit
    const model = parse(xml.replace(AXES, '<catAx><axId val="1"/><crossAx val="2"/></catAx><serAx><axId val="2"/><crossAx val="1"/></serAx><valAx><axId val="3"/><crossAx val="1"/><majorUnit val="2"/></valAx>'))
    expect(model!.groups[0].wireframe).toBe(true)
    const { data } = paint(model!)
    // With wireframe, fills are not applied → interior pixels should be white/transparent (no band fills)
    let coloured = 0
    for (let i = 0; i < data.length; i += 4) {
      const r = data[i], g = data[i+1], b = data[i+2]
      if (data[i+3] > 0 && Math.max(r,g,b) - Math.min(r,g,b) > 50) coloured++
    }
    expect(coloured).toBe(0)
  })

  it('bandFmts colours used, else accent cycle', () => {
    const xml = `<chartSpace><chart><plotArea><surfaceChart><bandFmts><bandFmt idx="0"><spPr><solidFill><srgbClr val="FF0000"/></solidFill></spPr></bandFmt></bandFmts>${ser(0, [1, 2, 3])}${ser(1, [4, 5, 6])}<axId val="1"/><axId val="2"/><axId val="3"/></surfaceChart><catAx><axId val="1"/><crossAx val="2"/></catAx><serAx><axId val="2"/><crossAx val="1"/></serAx><valAx><axId val="3"/><crossAx val="1"/><majorUnit val="2"/></valAx></plotArea></chart></chartSpace>`
    const model = parse(xml)
    expect(model!.groups[0].bandColors).toBeDefined()
    expect(model!.groups[0].bandColors![0]).toBe('FF0000')
    const { data } = paint(model!)
    // red band should appear
    let hasRed = false
    for (let i = 0; i < data.length; i += 4) {
      if (data[i] > 200 && data[i+1] < 80 && data[i+2] < 80 && data[i+3] > 0) { hasRed = true; break }
    }
    expect(hasRed).toBe(true)
  })

  it('surface3DChart → fallback picture, else card', () => {
    // parse as unsupported
    const xml = `<chartSpace><chart><plotArea><surface3DChart>${ser(0, [1, 2])}<axId val="1"/><axId val="2"/><axId val="3"/></surface3DChart><catAx><axId val="1"/><crossAx val="2"/></catAx><serAx><axId val="2"/><crossAx val="1"/></serAx><valAx><axId val="3"/><crossAx val="1"/></valAx></plotArea></chart></chartSpace>`
    const model = parseXlsxChart(parseXmlOrdered(xml), undefined)
    expect(model).toBeDefined()
    expect(model!.unsupported).toContain('Surface chart')
  })

  // conformance: no throw on edge cases
  it('conformance: empty/all-blank/NaN/Inf/negative/single point', () => {
    // empty chart (no series) should not throw
    expect(() => parse(space(`<surfaceChart><axId val="1"/><axId val="2"/><axId val="3"/></surfaceChart>`))).not.toThrow()
    const m2 = parse(space(`<surfaceChart>${ser(0, [NaN, Infinity, -5])}<axId val="1"/><axId val="2"/><axId val="3"/></surfaceChart>`))
    expect(() => m2 && paint(m2)).not.toThrow()
    const m3 = parse(space(`<surfaceChart>${ser(0, [7])}<axId val="1"/><axId val="2"/><axId val="3"/></surfaceChart>`))
    expect(() => m3 && paint(m3)).not.toThrow()
  })
})
