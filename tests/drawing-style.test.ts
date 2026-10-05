import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseDrawingColor, parseThemeContext, resolveDrawingColor, resolveDrawingStyle, type DrawingIssue, type DrawingStyle } from '../src/drawing/style'
import { paintGeometry } from '../src/drawing/paint'
import type { ResolvedPath } from '../src/drawing/geometry'

const xml = parseXmlOrdered
const themeXml = `<a:theme xmlns:a="urn:a"><a:themeElements><a:clrScheme name="test">
<a:dk1><a:sysClr val="windowText" lastClr="112233"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
<a:accent1><a:srgbClr val="5B9BD5"/></a:accent1><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="4472C4"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
</a:clrScheme><a:fontScheme name="test"><a:majorFont><a:latin typeface="Major"/><a:ea typeface="East"/><a:cs typeface="Complex"/><a:font script="Jpan" typeface="Supplement"/></a:majorFont><a:minorFont><a:latin typeface="Minor"/></a:minorFont></a:fontScheme>
<a:fmtScheme name="test"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350" cap="flat"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:headEnd type="diamond" w="lg" len="sm"/></a:ln></a:lnStyleLst><a:bgFillStyleLst><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:bgFillStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw/></a:effectLst></a:effectStyle></a:effectStyleLst></a:fmtScheme></a:themeElements></a:theme>`
const theme = () => parseThemeContext(themeXml)
const color = (body: string, context = theme()) => resolveDrawingColor(parseDrawingColor(xml(`<c>${body}</c>`)), context)!
const rgba = (r: number, g: number, b: number, a = 1) => ({ r, g, b, a })
const rectangle = (x = 0, y = 0, w = 100, h = 100, fill = 'norm', stroke = true): ResolvedPath => ({ fill, stroke, commands: [['moveTo', x, y], ['lnTo', x + w, y], ['lnTo', x + w, y + h], ['lnTo', x, y + h], ['close']] })
const line = (commands: ResolvedPath['commands'] = [['moveTo', 20, 40], ['lnTo', 100, 40]]): ResolvedPath => ({ fill: 'none', stroke: true, commands })
const surface = (w = 140, h = 100) => {
  const canvas = createCanvas(w, h)
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  const pixel = (x: number, y: number) => Array.from(ctx.getImageData(x, y, 1, 1).data)
  return { canvas, ctx, pixel }
}
const blackLine = (extra: Partial<NonNullable<DrawingStyle['line']>> = {}): DrawingStyle => ({ issues: [], line: { width: 2, fill: { kind: 'solid', color: rgba(0, 0, 0) }, ...extra } })
const fillStyle = (r = 128, g = 128, b = 128): DrawingStyle => ({ issues: [], fill: { kind: 'solid', color: rgba(r, g, b) } })

// Expected values are defined independently from the resolver/painter.
describe('DrawingML theme and style resolution', () => {
  it('retains palette, font matrix and serializable style definitions', () => {
    const t = theme()
    expect(t.palette.accent1).toBe('#5b9bd5')
    expect(t.palette.dk1).toBe('#112233')
    expect(t.fonts.major).toEqual({ latin: 'Major', eastAsian: 'East', complexScript: 'Complex', supplemental: { Jpan: 'Supplement' } })
    expect(t.fonts.minor.latin).toBe('Minor')
    expect(t.fillStyles).toHaveLength(2)
    expect(t.lineStyles).toHaveLength(2)
    expect(JSON.parse(JSON.stringify(t))).toEqual(t)
  })
  it('resolves caller color map aliases without mutating the theme or XML', () => {
    const input = xml(themeXml)
    const before = JSON.stringify(input)
    const t = parseThemeContext(input, { bg1: 'accent4', tx1: 'accent6' })
    expect(color('<a:schemeClr val="bg1"/>', t)).toEqual(rgba(255, 192, 0))
    expect(color('<a:schemeClr val="tx1"/>', t)).toEqual(rgba(112, 173, 71))
    expect(color('<a:schemeClr val="bg2"/>', t)).toEqual(rgba(238, 238, 238))
    expect(JSON.stringify(input)).toBe(before)
  })
  it('applies repeated/interleaved transforms in source order', () => {
    expect(color('<a:srgbClr val="FF0000"><a:alpha val="80000"/><a:alphaMod val="50000"/><a:alphaOff val="10000"/><a:alphaMod val="50000"/></a:srgbClr>').a).toBeCloseTo(.25)
    expect(color('<a:srgbClr val="FF0000"><a:alphaMod val="50000"/><a:alpha val="80000"/></a:srgbClr>').a).toBe(.8)
    expect(color('<a:srgbClr val="FF0000"><a:hueOff val="5400000"/><a:hue val="10800000"/><a:hueOff val="5400000"/></a:srgbClr>')).toEqual(rgba(128, 0, 255))
  })
  it('mixes DrawingML tint and shade in linear RGB', () => {
    expect(color('<a:srgbClr val="808080"><a:shade val="50000"/></a:srgbClr>')).toEqual(rgba(92, 92, 92))
    expect(color('<a:srgbClr val="808080"><a:tint val="50000"/></a:srgbClr>')).toEqual(rgba(205, 205, 205))
    expect(color('<a:srgbClr val="000000"><a:tint val="50000"/></a:srgbClr>')).toEqual(rgba(188, 188, 188))
  })
  it('supports HSL set/mod/off operations with clamping and wrapping', () => {
    expect(color('<a:hslClr hue="0" sat="100000" lum="50000"><a:hueMod val="50000"/><a:hueOff val="7200000"/><a:satMod val="50000"/><a:satOff val="50000"/><a:lumMod val="50000"/><a:lumOff val="25000"/></a:hslClr>')).toEqual(rgba(0, 255, 0))
    expect(color('<a:srgbClr val="FF0000"><a:sat val="0"/><a:lum val="25000"/></a:srgbClr>')).toEqual(rgba(64, 64, 64))
  })
  it('preserves hue through consecutive achromatic HSL transformations', () => {
    expect(color('<a:srgbClr val="0000FF"><a:sat val="0"/><a:hueOff val="5400000"/><a:sat val="100000"/></a:srgbClr>')).toEqual(rgba(255, 0, 128))
    expect(color('<a:hslClr hue="14400000" sat="100000" lum="0"><a:lum val="50000"/></a:hslClr>')).toEqual(rgba(0, 0, 255))
  })
  it('supports system fallback, presets, scRGB and linear channel transforms', () => {
    expect(color('<a:sysClr val="windowText" lastClr="AABBCC"/>')).toEqual(rgba(170, 187, 204))
    expect(color('<a:prstClr val="dkBlue"/>')).toEqual(rgba(0, 0, 139))
    expect(color('<a:scrgbClr r="50000" g="0" b="100000"/>')).toEqual(rgba(188, 0, 255))
    expect(color('<a:srgbClr val="000000"><a:red val="50000"/><a:redMod val="50000"/><a:redOff val="25000"/><a:green val="100000"/><a:greenMod val="0"/><a:greenOff val="100000"/><a:blue val="100000"/><a:blueMod val="50000"/><a:blueOff val="-50000"/></a:srgbClr>')).toEqual(rgba(188, 255, 0))
  })
  it('reports invalid colors and deferred transforms without returning NaN', () => {
    const issues: DrawingIssue[] = []
    const c = resolveDrawingColor(parseDrawingColor(xml('<c><a:srgbClr val="808080"><a:alpha val="oops"/><a:gamma/><a:redOff val="NaN"/></a:srgbClr></c>')), theme(), undefined, issues)
    expect(c).toEqual(rgba(128, 128, 128))
    expect(issues.map(i => i.kind)).toContain('unsupported-color-transform')
    expect(issues.map(i => i.kind)).toContain('invalid-color-transform')
    expect(resolveDrawingColor(parseDrawingColor(xml('<c><a:schemeClr val="missing"/></c>')), theme(), undefined, issues)).toBeUndefined()
    expect(resolveDrawingColor(parseDrawingColor(xml('<c><a:srgbClr val="xyz"/></c>')), theme(), undefined, issues)).toBeUndefined()
  })
  it('rejects malformed theme XML and cyclic palette colors with issues', () => {
    expect(parseThemeContext('<bad>').issues.some(i => i.kind === 'invalid-theme')).toBe(true)
    const cyclic = parseThemeContext('<theme><themeElements><clrScheme><accent1><schemeClr val="accent2"/></accent1><accent2><schemeClr val="accent1"/></accent2></clrScheme></themeElements></theme>')
    expect(cyclic.palette).toEqual({})
    expect(cyclic.issues.some(i => i.kind === 'invalid-color')).toBe(true)
  })
  it('resolves style references and retains inherited line fields with direct arrows', () => {
    const style = resolveDrawingStyle(xml('<spPr><a:ln><a:tailEnd type="triangle" w="sm" len="lg"/></a:ln></spPr>'), xml('<style><a:fillRef idx="1"><a:schemeClr val="lt1"/></a:fillRef><a:lnRef idx="2"><a:schemeClr val="accent6"/></a:lnRef></style>'), theme())
    expect(style.fill).toEqual({ kind: 'solid', color: rgba(255, 255, 255) })
    expect(style.line?.width).toBeCloseTo(12700 / 9525)
    expect(style.line?.fill).toEqual({ kind: 'solid', color: rgba(112, 173, 71) })
    expect(style.line?.headEnd).toEqual({ type: 'diamond', width: 'lg', length: 'sm' })
    expect(style.line?.tailEnd).toEqual({ type: 'triangle', width: 'sm', length: 'lg' })
  })
  it('direct noFill overrides theme fill and direct width retains theme color', () => {
    const style = resolveDrawingStyle(xml('<spPr><a:noFill/><a:ln w="38100"><a:headEnd type="none"/></a:ln></spPr>'), xml('<style><a:fillRef idx="2"><a:schemeClr val="accent5"/></a:fillRef><a:lnRef idx="2"><a:schemeClr val="accent4"/></a:lnRef></style>'), theme())
    expect(style.fill).toEqual({ kind: 'none' })
    expect(style.line?.width).toBe(4)
    expect(style.line?.fill).toEqual({ kind: 'solid', color: rgba(255, 192, 0) })
    expect(style.line?.headEnd).toEqual({ type: 'none', width: 'lg', length: 'sm' })
  })
  it('selects background fill indices and handles zero/missing refs explicitly', () => {
    expect(resolveDrawingStyle(undefined, xml('<style><a:fillRef idx="1001"/></style>'), theme()).fill).toEqual({ kind: 'solid', color: rgba(18, 52, 86) })
    expect(resolveDrawingStyle(undefined, xml('<style><a:fillRef idx="0"/><a:lnRef idx="0"/></style>'), theme())).toMatchObject({ fill: { kind: 'none' }, line: { fill: { kind: 'none' } } })
    const invalid = resolveDrawingStyle(undefined, xml('<style><a:fillRef idx="999"/></style>'), theme())
    expect(invalid.issues.some(i => i.kind === 'missing-theme-style')).toBe(true)
  })
  it('uses transformed reference color as phClr, preserving gradient definitions', () => {
    const style = resolveDrawingStyle(undefined, xml('<style><a:fillRef idx="2"><a:schemeClr val="accent5"/></a:fillRef><a:lnRef idx="1"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef></style>'), theme())
    expect(style.fill).toMatchObject({ kind: 'gradient', gradient: 'linear', angle: Math.PI / 2, scaled: false })
    if (style.fill?.kind === 'gradient') {
      expect(style.fill.stops.map(s => s.position)).toEqual([0, .5, 1])
      expect(style.fill.stops[0].color.r).toBeGreaterThan(68)
      expect(style.fill.stops[0].color.b).toBeGreaterThan(196)
    }
    expect(style.line?.fill).toEqual({ kind: 'solid', color: rgba(34, 63, 89) })
  })
  it('merges direct gradient fields and arrow fields independently', () => {
    const style = resolveDrawingStyle(xml('<spPr><a:gradFill><a:lin ang="0"/></a:gradFill><a:ln><a:headEnd len="lg"/></a:ln></spPr>'), xml('<style><a:fillRef idx="2"><a:schemeClr val="accent5"/></a:fillRef><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef></style>'), theme())
    expect(style.fill).toMatchObject({ kind: 'gradient', angle: 0, scaled: false })
    if (style.fill?.kind === 'gradient') expect(style.fill.stops).toHaveLength(3)
    expect(style.line?.headEnd).toEqual({ type: 'diamond', width: 'lg', length: 'lg' })
  })
  it('documents rectangular/shape gradient first-stop and deferred effects', () => {
    const style = resolveDrawingStyle(xml('<spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:path path="rect"/></a:gradFill><a:effectLst><a:outerShdw/></a:effectLst><a:sp3d/></spPr>'), xml('<style><a:effectRef idx="1"/></style>'), theme())
    expect(style.fill).toEqual({ kind: 'solid', color: rgba(255, 0, 0) })
    expect(style.issues.some(i => i.kind === 'unsupported-gradient' && i.feature === 'rect')).toBe(true)
    expect(style.issues.some(i => i.kind === 'unsupported-effect')).toBe(true)
    expect(style.issues.some(i => i.kind === 'unsupported-3d')).toBe(true)
  })
})

describe('geometry painting', () => {
  it('paints linear gradient angle over local extents', () => {
    const style = resolveDrawingStyle(xml('<spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill></spPr>'))
    const { ctx, pixel } = surface()
    paintGeometry(ctx, [rectangle()], style, 100, 100)
    expect(pixel(5, 50)[0]).toBeGreaterThan(230)
    expect(pixel(95, 50)[2]).toBeGreaterThan(230)
    expect(pixel(50, 5)).toEqual(pixel(50, 95))
  })
  it('scales a linear gradient vector with the extent when requested', () => {
    const paint = (scaled: boolean) => {
      const s = surface(440, 220)
      paintGeometry(s.ctx, [rectangle(0, 0, 400, 200)], { issues: [], fill: { kind: 'gradient', gradient: 'linear', angle: Math.PI / 4, scaled, stops: [{ position: 0, color: rgba(255, 0, 0) }, { position: 1, color: rgba(0, 0, 255) }] } }, 400, 200)
      return s.pixel
    }
    const unscaled = paint(false), scaled = paint(true)
    // With scaled=false, points with the same x+y have the same color.
    expect(unscaled(80, 20)[2]).toBeCloseTo(unscaled(20, 80)[2], 0)
    // Office's 400x200 native fixture is bluer at (20,80) than (80,20).
    // Corrected scaled geometry gives t=(x+2y)/800, sampled at pixel centers.
    expect(Math.abs(scaled(80, 20)[2] - 39)).toBeLessThanOrEqual(1)
    expect(Math.abs(scaled(20, 80)[2] - 58)).toBeLessThanOrEqual(1)
  })
  it('paints circular radial gradients from stop zero at the point focus', () => {
    const style = resolveDrawingStyle(xml('<spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="25000" t="50000" r="75000" b="50000"/></a:path></a:gradFill></spPr>'))
    const { ctx, pixel } = surface(120, 120)
    paintGeometry(ctx, [rectangle()], style, 100, 100)
    expect(pixel(25, 50)[0]).toBeGreaterThan(245)
    expect(Math.abs(pixel(98, 50)[2] - 196)).toBeLessThanOrEqual(1)
    expect(pixel(25, 75)[2]).toBeGreaterThan(pixel(25, 50)[2] + 90)
  })
  it('matches native circular shading geometry for non-square extents and asymmetric focus', () => {
    const { ctx, pixel } = surface(320, 170)
    const style: DrawingStyle = { issues: [], fill: { kind: 'gradient', gradient: 'circle', angle: 0, scaled: false, focus: { left: .2, top: .3, right: .4, bottom: .5 }, stops: [{ position: 0, color: rgba(255, 0, 0) }, { position: 1, color: rgba(0, 0, 255) }] } }
    const issues = paintGeometry(ctx, [rectangle(0, 0, 300, 150)], style, 300, 150)
    expect(issues).toEqual([])
    // Native PDF: focus=(100,56.25), outer center=(150,75), radius=167.705.
    // The old focus-rectangle midpoint (120,60) must not be a flat red region.
    expect(pixel(100, 56)[2]).toBeLessThan(3)
    expect(Math.abs(pixel(120, 60)[2] - 24)).toBeLessThanOrEqual(1)
    expect(Math.abs(pixel(150, 56)[2] - 59)).toBeLessThanOrEqual(1)
    expect(Math.abs(pixel(100, 106)[2] - 71)).toBeLessThanOrEqual(1)
    expect(Math.abs(pixel(298, 56)[2] - 234)).toBeLessThanOrEqual(1)
  })
  it('uses the same centered circumcircle for omitted and explicit default focus', () => {
    const omitted = surface(220, 120), explicit = surface(220, 120)
    const fill: NonNullable<DrawingStyle['fill']> = { kind: 'gradient', gradient: 'circle', angle: 0, scaled: false, stops: [{ position: 0, color: rgba(255, 0, 0) }, { position: 1, color: rgba(0, 0, 255) }] }
    paintGeometry(omitted.ctx, [rectangle(0, 0, 200, 100)], { issues: [], fill }, 200, 100)
    paintGeometry(explicit.ctx, [rectangle(0, 0, 200, 100)], { issues: [], fill: { ...fill, focus: { left: .5, top: .5, right: .5, bottom: .5 } } }, 200, 100)
    // Native PDF radius is 111.803px: a physical circle through all four corners.
    expect(Math.abs(omitted.pixel(150, 50)[2] - 115)).toBeLessThanOrEqual(1)
    expect(Math.abs(omitted.pixel(100, 90)[2] - 92)).toBeLessThanOrEqual(1)
    expect(omitted.pixel(199, 99)[2]).toBeGreaterThan(250)
    expect(omitted.ctx.getImageData(0, 0, 200, 100).data).toEqual(explicit.ctx.getImageData(0, 0, 200, 100).data)
  })
  it('paints a translucent radial texture exactly once, under a caller transform', () => {
    const { ctx, pixel } = surface(240, 160)
    ctx.translate(10, 20)
    const style: DrawingStyle = { issues: [], fill: { kind: 'gradient', gradient: 'circle', angle: 0, scaled: false, focus: { left: .4, top: .5, right: .6, bottom: .5 }, stops: [{ position: 0, color: rgba(255, 0, 0, .5) }, { position: 1, color: rgba(0, 0, 255, .5) }] } }
    paintGeometry(ctx, [rectangle(0, 0, 200, 100)], style, 200, 100)
    expect(pixel(90, 70)[0]).toBeGreaterThan(245)
    expect(pixel(90, 75)[3]).toBeGreaterThanOrEqual(127)
    expect(pixel(170, 70)[3]).toBeGreaterThanOrEqual(127)
    expect(pixel(110, 118)[3]).toBeGreaterThanOrEqual(127)
    expect([pixel(90, 75)[3], pixel(170, 70)[3], pixel(110, 118)[3]].every(a => a <= 128)).toBe(true)
    const data = ctx.getImageData(11, 21, 198, 98).data
    for (let i = 3; i < data.length; i += 4) { expect(data[i]).toBeGreaterThanOrEqual(127); expect(data[i]).toBeLessThanOrEqual(128) }
  })
  it('draws line preset and custom dash gaps', () => {
    for (const dash of ['dash', [2, 3]] as const) {
      const { ctx, pixel } = surface()
      paintGeometry(ctx, [line([['moveTo', 10, 40], ['lnTo', 120, 40]])], blackLine({ dash: typeof dash === 'string' ? dash : [...dash] }), 130, 80)
      expect(pixel(11, 40)[3]).toBeGreaterThan(200)
      expect(pixel(dash === 'dash' ? 19 : 15, 40)[3]).toBe(0)
    }
  })
  it('keeps gradient colors aligned on arrow decorations', () => {
    const { ctx, pixel } = surface()
    const style = blackLine({ width: 4, fill: { kind: 'gradient', gradient: 'linear', angle: 0, scaled: false, stops: [{ position: 0, color: rgba(255, 0, 0) }, { position: 1, color: rgba(0, 0, 255) }] }, headEnd: { type: 'triangle', width: 'lg', length: 'lg' }, tailEnd: { type: 'triangle', width: 'lg', length: 'lg' } })
    paintGeometry(ctx, [line()], style, 120, 80)
    expect(pixel(27, 37)[0]).toBeGreaterThan(190)
    expect(pixel(93, 37)[2]).toBeGreaterThan(190)
  })
  it('paints radial line gradients with the same local brush as the fill', () => {
    const { ctx, pixel } = surface(220, 120)
    const style = blackLine({ width: 4, fill: { kind: 'gradient', gradient: 'circle', angle: 0, scaled: false, focus: { left: .4, top: .5, right: .6, bottom: .5 }, stops: [{ position: 0, color: rgba(255, 0, 0) }, { position: 1, color: rgba(0, 0, 255) }] } })
    expect(paintGeometry(ctx, [line([['moveTo', 1, 50], ['lnTo', 199, 50]])], style, 200, 100)).toEqual([])
    expect(pixel(80, 50)[0]).toBeGreaterThan(245)
    expect(Math.abs(pixel(198, 50)[2] - 229)).toBeLessThanOrEqual(1)
  })
  it('uses an HTML canvas owner document when OffscreenCanvas is unavailable', () => {
    const canvas = document.createElement('canvas'); canvas.width = 120; canvas.height = 120
    const ctx = canvas.getContext('2d')!
    const style: DrawingStyle = { issues: [], fill: { kind: 'gradient', gradient: 'circle', angle: 0, scaled: false, stops: [{ position: 0, color: rgba(255, 0, 0) }, { position: 1, color: rgba(0, 0, 255) }] } }
    expect(paintGeometry(ctx, [rectangle()], style, 100, 100)).toEqual([])
    expect(ctx.getImageData(50, 50, 1, 1).data[0]).toBeGreaterThan(245)
  })
  it('bounds radial scratch resolution for huge finite geometry', () => {
    const { ctx, pixel } = surface()
    const style: DrawingStyle = { issues: [], fill: { kind: 'gradient', gradient: 'circle', angle: 0, scaled: false, stops: [{ position: 0, color: rgba(255, 0, 0) }, { position: 1, color: rgba(0, 0, 255) }] } }
    expect(paintGeometry(ctx, [rectangle(0, 0, 1e6, 1e6)], style, 1e6, 1e6)).toEqual([])
    expect(pixel(50, 50)[2]).toBeGreaterThan(230)
  })
  it('parses custom dashes, cap/join and miter limit in line units', () => {
    const style = resolveDrawingStyle(xml('<spPr><a:ln w="19050" cap="rnd"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:custDash><a:ds d="200000" sp="300000"/></a:custDash><a:miter lim="500000"/></a:ln></spPr>'))
    expect(style.line).toMatchObject({ width: 2, dash: [2, 3], cap: 'round', join: 'miter', miterLimit: 5 })
  })
  it.each(['constructor', '__proto__', 'toString'])('reports inherited-object names as invalid line caps (%s)', cap => {
    const style = resolveDrawingStyle(xml(`<spPr><a:ln cap="${cap}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></spPr>`))
    expect(style.line?.cap).toBeUndefined()
    expect(style.issues.some(i => i.kind === 'invalid-line' && i.feature === 'cap')).toBe(true)
  })
  it.each(['constructor', '__proto__', 'toString'])('falls back to solid for malformed preset dash names (%s)', dash => {
    const { ctx, pixel } = surface(120, 80)
    const style = resolveDrawingStyle(xml(`<spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="${dash}"/></a:ln></spPr>`))
    ctx.lineWidth = 7; ctx.lineCap = 'round'; ctx.setLineDash([7, 9]); ctx.lineDashOffset = 5; ctx.translate(3, 2)
    const before = ctx.getTransform()
    let issues: DrawingIssue[] = []
    expect(() => { issues = paintGeometry(ctx, [line([['moveTo', 10, 20], ['lnTo', 100, 20]]), line([['moveTo', 10, 60], ['lnTo', 100, 60]])], style, 120, 80) }).not.toThrow()
    expect(issues.some(i => i.kind === 'unsupported-line' && i.feature === dash)).toBe(true)
    expect(pixel(20, 22)[3]).toBeGreaterThan(200)
    expect(pixel(20, 62)[3]).toBeGreaterThan(200)
    expect(ctx.lineWidth).toBe(7); expect(ctx.lineCap).toBe('round'); expect(ctx.getLineDash()).toEqual([7, 9]); expect(ctx.lineDashOffset).toBe(5); expect(ctx.getTransform()).toEqual(before)
  })
  it('covers acute miter joins with radial paint using the effective Canvas miter limit', () => {
    const path = line([['moveTo', 100, 200], ['lnTo', 102, 100], ['lnTo', 104, 200]])
    for (const miterLimit of [100, 1000, undefined, NaN, 0, Infinity]) {
      const solid = surface(240, 260), radial = surface(240, 260)
      const stroke = { width: 4, join: 'miter' as const, miterLimit }
      for (const ctx of [solid.ctx, radial.ctx]) { ctx.lineWidth = 9; ctx.miterLimit = 3; ctx.translate(3, 5) }
      const before = radial.ctx.getTransform()
      const gradient = { kind: 'gradient' as const, gradient: 'circle' as const, angle: 0, scaled: false, stops: [{ position: 0, color: rgba(0, 0, 0) }, { position: 1, color: rgba(0, 0, 0) }] }
      expect(paintGeometry(solid.ctx, [path], blackLine(stroke), 220, 240)).toEqual([])
      expect(paintGeometry(radial.ctx, [path], blackLine({ ...stroke, fill: gradient }), 220, 240)).toEqual([])
      if (miterLimit === 100 || miterLimit === 1000) expect(solid.pixel(105, 65)[3]).toBe(255)
      expect(radial.pixel(105, 65)).toEqual(solid.pixel(105, 65))
      expect(radial.pixel(105, 115)).toEqual(solid.pixel(105, 115))
      expect(radial.ctx.lineWidth).toBe(9); expect(radial.ctx.miterLimit).toBe(3); expect(radial.ctx.getTransform()).toEqual(before)
    }
  })
  it.each(['triangle', 'stealth', 'diamond', 'oval', 'arrow'] as const)('draws native %s endpoint regions with a gradient under a caller transform', type => {
    for (const reverse of [false, true]) {
      const { ctx, pixel } = surface()
      const commands: ResolvedPath['commands'] = reverse ? [['moveTo', 100, 40], ['lnTo', 20, 40]] : [['moveTo', 20, 40], ['lnTo', 100, 40]]
      ctx.translate(4, 7)
      const before = ctx.getTransform()
      const fill = { kind: 'gradient' as const, gradient: 'linear' as const, angle: 0, scaled: false, stops: [{ position: 0, color: rgba(255, 0, 0) }, { position: 1, color: rgba(0, 0, 255) }] }
      paintGeometry(ctx, [line(commands)], blackLine({ width: 4, fill, headEnd: { type, width: 'lg', length: 'lg' }, tailEnd: { type, width: 'lg', length: 'lg' } }), 120, 80)
      if (type === 'diamond' || type === 'oval') {
        // Office centers these decorations on each endpoint; both extend outward.
        expect(pixel(18, 46)[3]).toBeGreaterThan(100)
        expect(pixel(109, 46)[3]).toBeGreaterThan(100)
        expect(pixel(18, 46)[0]).toBeGreaterThan(200)
        expect(pixel(109, 46)[2]).toBeGreaterThan(200)
        expect(pixel(31, 44)[3]).toBe(0)
        expect(pixel(96, 44)[3]).toBe(0)
      } else {
        // Triangle, stealth and open arrows put their tip exactly on the endpoint.
        const y = type === 'arrow' ? 43 : 44
        expect(pixel(31, y)[3]).toBeGreaterThan(100)
        expect(pixel(97, y)[3]).toBeGreaterThan(100)
        expect(pixel(31, y)[0]).toBeGreaterThan(190)
        expect(pixel(97, y)[2]).toBeGreaterThan(190)
        expect(pixel(18, 46)[3]).toBe(0)
        expect(pixel(109, 46)[3]).toBe(0)
      }
      expect(pixel(14, 47)[3]).toBe(0)
      expect(pixel(114, 47)[3]).toBe(0)
      expect(ctx.getTransform()).toEqual(before)
    }
  })
  it('places headEnd at path start and tailEnd at path end', () => {
    const { ctx, pixel } = surface()
    paintGeometry(ctx, [line()], blackLine({ width: 4, headEnd: { type: 'triangle', width: 'lg', length: 'lg' } }), 120, 80)
    expect(pixel(27, 37)[3]).toBeGreaterThan(100)
    expect(pixel(93, 37)[3]).toBe(0)
  })
  it.each(([
    [['moveTo', 30, 30], ['lnTo', 30, 30], ['lnTo', 30, 80]],
    [['moveTo', 30, 30], ['quadBezTo', 30, 30, 30, 80]],
    [['moveTo', 30, 30], ['cubicBezTo', 30, 30, 30, 30, 30, 80]],
    [['moveTo', 30, 30], ['arcTo', 0, 30, 30, 30, 0, Math.PI / 2, false]],
  ] as ResolvedPath['commands'][]).map(commands => ({ commands })))('finds first nonzero tangent for endpoint decoration (%j)', ({ commands }) => {
    const { ctx, pixel } = surface()
    paintGeometry(ctx, [line(commands)], blackLine({ width: 4, headEnd: { type: 'triangle', width: 'lg', length: 'lg' } }), 100, 100)
    expect(pixel(27, 37)[3]).toBeGreaterThan(100)
    expect(pixel(37, 27)[3]).toBe(0)
  })
  it('uses the final nonzero curve tangent despite trailing zero segments', () => {
    const { ctx, pixel } = surface()
    paintGeometry(ctx, [line([['moveTo', 20, 20], ['cubicBezTo', 20, 60, 80, 60, 80, 60], ['lnTo', 80, 60]])], blackLine({ width: 4, tailEnd: { type: 'triangle', width: 'lg', length: 'lg' } }), 100, 100)
    expect(pixel(73, 57)[3]).toBeGreaterThan(100)
    expect(pixel(83, 53)[3]).toBe(0)
  })
  it('honors per-path fill/stroke flags and lighten/darken shades', () => {
    const { ctx, pixel } = surface(180, 100)
    const modes = ['none', 'norm', 'lighten', 'lightenLess', 'darken', 'darkenLess']
    const paths = modes.map((mode, i) => rectangle(i * 25 + 2, 2, 20, 20, mode, false))
    paintGeometry(ctx, { paths, issues: [] }, fillStyle(), 180, 100)
    expect(pixel(10, 10)[3]).toBe(0)
    expect(pixel(35, 10)[0]).toBe(128)
    expect(pixel(60, 10)[0]).toBeGreaterThan(pixel(85, 10)[0])
    expect(pixel(85, 10)[0]).toBeGreaterThan(128)
    expect(pixel(110, 10)[0]).toBeLessThan(pixel(135, 10)[0])
    expect(pixel(135, 10)[0]).toBeLessThan(128)
    const absent = surface()
    paintGeometry(absent.ctx, [rectangle(10, 10, 50, 50, 'none', false)], { ...fillStyle(), line: blackLine().line }, 100, 100)
    expect(absent.pixel(10, 30)[3]).toBe(0)
  })
  it('uses independent arrow width and length enums', () => {
    const small = surface(), wide = surface(), long = surface()
    paintGeometry(small.ctx, [line()], blackLine({ width: 4, tailEnd: { type: 'triangle', width: 'sm', length: 'sm' } }), 120, 80)
    paintGeometry(wide.ctx, [line()], blackLine({ width: 4, tailEnd: { type: 'triangle', width: 'lg', length: 'sm' } }), 120, 80)
    paintGeometry(long.ctx, [line()], blackLine({ width: 4, tailEnd: { type: 'triangle', width: 'sm', length: 'lg' } }), 120, 80)
    expect(small.pixel(94, 35)[3]).toBe(0)
    expect(wide.pixel(94, 35)[3]).toBeGreaterThan(100)
    expect(small.pixel(88, 37)[3]).toBe(0)
    expect(long.pixel(88, 37)[3]).toBeGreaterThan(100)
  })
  it('keeps native medium endpoint decorations visible on a thin theme line', () => {
    const { ctx, pixel } = surface()
    paintGeometry(ctx, [line()], blackLine({ width: 6350 / 9525, headEnd: { type: 'triangle' }, tailEnd: { type: 'triangle' } }), 120, 80)
    // Native a11 decorations are approximately 7px: 2.5pt * 1.5^2 = 7.5px.
    expect(pixel(26, 38)[3]).toBeGreaterThan(180)
    expect(pixel(93, 38)[3]).toBeGreaterThan(180)
    expect(pixel(90, 38)[3]).toBe(0)
  })
  it('covers thin large arrows with the radial brush without clipping their sides', () => {
    const { ctx, pixel } = surface()
    const style = blackLine({ width: 6350 / 9525, fill: { kind: 'gradient', gradient: 'circle', angle: 0, scaled: false, stops: [{ position: 0, color: rgba(255, 0, 0) }, { position: 1, color: rgba(0, 0, 255) }] }, tailEnd: { type: 'triangle', width: 'lg', length: 'lg' } })
    expect(paintGeometry(ctx, [line()], style, 120, 80)).toEqual([])
    // Large size is 2.5pt * 1.5^3 = 11.25px, independent of the thin pen.
    expect(pixel(90, 35)[3]).toBeGreaterThan(180)
    expect(pixel(90, 44)[3]).toBeGreaterThan(180)
    expect(pixel(87, 35)[3]).toBe(0)
  })
  it('keeps fills aligned under caller transforms and respects an existing clip', () => {
    const { ctx, pixel } = surface()
    ctx.beginPath(); ctx.rect(10, 10, 50, 50); ctx.clip(); ctx.translate(10, 10)
    const style = resolveDrawingStyle(xml('<spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></spPr>'))
    paintGeometry(ctx, [rectangle()], style, 100, 100)
    expect(pixel(15, 30)[0]).toBeGreaterThan(230)
    expect(pixel(65, 30)[3]).toBe(0)
    ctx.fillStyle = 'green'; ctx.fillRect(80, 0, 10, 10)
    expect(pixel(95, 15)[3]).toBe(0)
  })
  it('detects nonfinite derived gradient coordinates before invoking Canvas', () => {
    const { ctx } = surface()
    const issues = paintGeometry(ctx, [rectangle()], { issues: [], fill: { kind: 'gradient', gradient: 'linear', angle: Math.PI / 4, scaled: true, stops: [{ position: 0, color: rgba(0, 0, 0) }, { position: 1, color: rgba(255, 255, 255) }] } }, Number.MAX_VALUE, Number.MAX_VALUE)
    expect(issues.some(i => i.kind === 'invalid-paint')).toBe(true)
  })
  it('preserves the caller canvas state, transform and clip', () => {
    const { ctx, pixel } = surface()
    ctx.fillStyle = '#abcdef'; ctx.strokeStyle = '#123456'; ctx.lineWidth = 7; ctx.lineCap = 'square'; ctx.lineJoin = 'bevel'; ctx.miterLimit = 3; ctx.globalAlpha = .7; ctx.setLineDash([7, 9]); ctx.lineDashOffset = 4; ctx.translate(3, 5)
    const before = ctx.getTransform()
    paintGeometry(ctx, [rectangle()], { ...fillStyle(255, 0, 0), line: blackLine({ headEnd: { type: 'oval' } }).line }, 100, 100)
    expect(ctx.fillStyle).toBe('#abcdef'); expect(ctx.strokeStyle).toBe('#123456'); expect(ctx.lineWidth).toBe(7); expect(ctx.lineCap).toBe('square'); expect(ctx.lineJoin).toBe('bevel'); expect(ctx.miterLimit).toBe(3); expect(ctx.globalAlpha).toBeCloseTo(.7); expect(ctx.getLineDash()).toEqual([7, 9]); expect(ctx.lineDashOffset).toBe(4)
    expect(ctx.getTransform()).toEqual(before)
    expect(pixel(0, 0)[3]).toBe(0)
  })
  it('rejects nonnumeric command operands before Canvas can coerce them to NaN', () => {
    const { ctx, pixel } = surface()
    const commands = [['moveTo', 'bad', 0], ['lnTo', 100, 50]] as unknown as ResolvedPath['commands']
    expect(paintGeometry(ctx, [line(commands)], blackLine(), 100, 100).some(i => i.kind === 'invalid-paint')).toBe(true)
    expect(pixel(50, 25)[3]).toBe(0)
  })
  it('paints a large finite custom path and continues with a valid neighboring path', () => {
    const { ctx, pixel } = surface(140, 100)
    const commands: ResolvedPath['commands'] = [['moveTo', 10, 20]]
    for (let i = 1; i <= 150000; i++) commands.push(['lnTo', 10 + i / 1500, 20])
    ctx.lineWidth = 7; ctx.translate(3, 2)
    const before = ctx.getTransform()
    let issues: DrawingIssue[] = []
    expect(() => { issues = paintGeometry(ctx, [line(commands), line([['moveTo', 10, 60], ['lnTo', 110, 60]])], blackLine(), 140, 100) }).not.toThrow()
    expect(issues).toEqual([])
    expect(pixel(63, 22)[3]).toBeGreaterThan(200)
    expect(pixel(63, 62)[3]).toBeGreaterThan(200)
    expect(ctx.lineWidth).toBe(7)
    expect(ctx.getTransform()).toEqual(before)
  })
  it('uses the straight segment tangent when an arc has zero radii', () => {
    const { ctx, pixel } = surface()
    // The formal ellipse derivative points down, but the zero-radius segment is drawn up.
    paintGeometry(ctx, [line([['moveTo', 30, 30 + 30 / Math.sqrt(2)], ['arcTo', 30, 30, 0, 30, Math.PI / 4, Math.PI * 2, false]])], blackLine({ width: 4, tailEnd: { type: 'triangle', width: 'lg', length: 'lg' } }), 100, 100)
    expect(pixel(27, 37)[3]).toBeGreaterThan(100)
    expect(pixel(37, 23)[3]).toBe(0)
  })
  it('skips invalid paths/extents and malformed style values without throwing or state leaks', () => {
    const { ctx, pixel } = surface()
    ctx.lineWidth = 5
    const bad = resolveDrawingStyle(xml('<spPr><a:gradFill><a:gsLst><a:gs pos="NaN"><a:srgbClr val="FF0000"/></a:gs></a:gsLst><a:lin ang="NaN"/></a:gradFill><a:ln w="NaN"><a:custDash><a:ds d="NaN" sp="-1"/></a:custDash></a:ln></spPr>'))
    expect(bad.issues.length).toBeGreaterThan(0)
    expect(() => paintGeometry(ctx, [line([['moveTo', NaN, 0], ['lnTo', 100, Infinity]])], bad, 100, 100)).not.toThrow()
    expect(paintGeometry(ctx, [rectangle()], fillStyle(), NaN, 100).some(i => i.kind === 'invalid-paint')).toBe(true)
    expect(pixel(50, 50)[3]).toBe(0)
    expect(ctx.lineWidth).toBe(5)
  })
})
