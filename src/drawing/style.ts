/** Serializable DrawingML color, theme and paint models shared by Office adapters. */
import { attrs, getChild, getChildren, orderedChildren, parseXmlOrdered, type XmlNode } from '../core/xml'
export interface DrawingIssue {
  kind: 'invalid-color' | 'invalid-color-transform' | 'unsupported-color-transform' | 'invalid-fill' | 'unsupported-fill' | 'unsupported-gradient' | 'invalid-line' | 'unsupported-line' | 'missing-theme-style' | 'unsupported-effect' | 'unsupported-3d' | 'invalid-theme' | 'invalid-paint'
  message: string
  feature?: string
  pathIndex?: number
}
export interface DrawingColor { r: number; g: number; b: number; a: number }
export interface ColorDefinition {
  kind: 'srgb' | 'scheme' | 'system' | 'preset' | 'scrgb' | 'hsl'
  value?: string
  lastColor?: string
  channels?: [string, string, string]
  transforms: Array<{ name: string; value?: string }>
}
export interface RelativeRect { left: number; top: number; right: number; bottom: number }
export type FillDefinition =
  | { kind: 'none' }
  | { kind: 'solid'; color?: ColorDefinition }
  | { kind: 'gradient'; stops?: Array<{ position: number; color?: ColorDefinition }>; gradient?: string; angle?: number; scaled?: boolean; focus?: Partial<RelativeRect>; rotateWithShape?: boolean }
export type DrawingFill =
  | { kind: 'none' }
  | { kind: 'solid'; color: DrawingColor }
  | { kind: 'gradient'; gradient: 'linear' | 'circle'; stops: Array<{ position: number; color: DrawingColor }>; angle: number; scaled: boolean; focus?: RelativeRect; rotateWithShape?: boolean }
export interface ArrowEnd { type?: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow'; width?: 'sm' | 'med' | 'lg'; length?: 'sm' | 'med' | 'lg' }
export interface DrawingLine {
  width?: number
  fill?: DrawingFill
  /** Preset name or alternating dash/gap lengths, expressed in line-width units. */
  dash?: string | number[]
  cap?: CanvasLineCap
  join?: CanvasLineJoin
  miterLimit?: number
  headEnd?: ArrowEnd
  tailEnd?: ArrowEnd
}
export interface LineDefinition extends Omit<DrawingLine, 'fill'> { fill?: FillDefinition }
export interface ThemeFonts { latin?: string; eastAsian?: string; complexScript?: string; supplemental: Record<string, string> }
export interface ThemeContext {
  colors: Record<string, ColorDefinition>
  palette: Record<string, string>
  colorMap: Record<string, string>
  fonts: { major: ThemeFonts; minor: ThemeFonts }
  fillStyles: FillDefinition[]
  bgFillStyles: FillDefinition[]
  lineStyles: LineDefinition[]
  effectStyles: string[][]
  issues: DrawingIssue[]
}
export interface DrawingStyle { fill?: DrawingFill; line?: DrawingLine; issues: DrawingIssue[] }

const colorKinds = { srgbClr: 'srgb', schemeClr: 'scheme', sysClr: 'system', prstClr: 'preset', scrgbClr: 'scrgb', hslClr: 'hsl' } as const
const clamp = (value: number): number => Math.min(1, Math.max(0, value))
const wrap = (value: number): number => ((value % 1) + 1) % 1
const numeric = (value: string | undefined): number | undefined => {
  if (value === undefined || value.trim() === '') return undefined
  const n = Number(value)
  return Number.isFinite(n) ? n : undefined
}
const percentage = (value: string | undefined): number | undefined => {
  if (value?.endsWith('%')) { const n = numeric(value.slice(0, -1)); return n === undefined ? undefined : n / 100 }
  const n = numeric(value)
  return n === undefined ? undefined : n / 100000
}
const issue = (issues: DrawingIssue[], kind: DrawingIssue['kind'], message: string, feature?: string): void => { issues.push({ kind, message, ...(feature ? { feature } : {}) }) }
const has = <T>(map: Record<string, T>, key: string): boolean => Object.prototype.hasOwnProperty.call(map, key)

/** Supply a color-containing node, e.g. solidFill, gs, fillRef, or palette slot. */
export function parseDrawingColor(node: XmlNode | undefined): ColorDefinition | undefined {
  for (const [name, child] of orderedChildren(node)) {
    if (!has(colorKinds, name)) continue
    const kind = colorKinds[name as keyof typeof colorKinds]
    const a = attrs(child)
    return {
      kind,
      ...(a.val !== undefined ? { value: a.val } : {}),
      ...(a.lastClr !== undefined ? { lastColor: a.lastClr } : {}),
      ...(kind === 'scrgb' ? { channels: [a.r ?? '', a.g ?? '', a.b ?? ''] as [string, string, string] } : {}),
      ...(kind === 'hsl' ? { channels: [a.hue ?? '', a.sat ?? '', a.lum ?? ''] as [string, string, string] } : {}),
      transforms: orderedChildren(child).filter(([n]) => n !== '#text').map(([name, n]) => ({ name, ...(attrs(n).val !== undefined ? { value: attrs(n).val } : {}) })),
    }
  }
  return undefined
}

// DrawingML preset names use dk/lt/med abbreviations for the CSS/X11 colors.
const presetColors: Record<string, string> = Object.fromEntries(`aliceBlue:F0F8FF antiqueWhite:FAEBD7 aqua:00FFFF aquamarine:7FFFD4 azure:F0FFFF beige:F5F5DC bisque:FFE4C4 black:000000 blanchedAlmond:FFEBCD blue:0000FF blueViolet:8A2BE2 brown:A52A2A burlyWood:DEB887 cadetBlue:5F9EA0 chartreuse:7FFF00 chocolate:D2691E coral:FF7F50 cornflowerBlue:6495ED cornsilk:FFF8DC crimson:DC143C cyan:00FFFF dkBlue:00008B dkCyan:008B8B dkGoldenrod:B8860B dkGray:A9A9A9 dkGreen:006400 dkKhaki:BDB76B dkMagenta:8B008B dkOliveGreen:556B2F dkOrange:FF8C00 dkOrchid:9932CC dkRed:8B0000 dkSalmon:E9967A dkSeaGreen:8FBC8F dkSlateBlue:483D8B dkSlateGray:2F4F4F dkTurquoise:00CED1 dkViolet:9400D3 deepPink:FF1493 deepSkyBlue:00BFFF dimGray:696969 dodgerBlue:1E90FF firebrick:B22222 floralWhite:FFFAF0 forestGreen:228B22 fuchsia:FF00FF gainsboro:DCDCDC ghostWhite:F8F8FF gold:FFD700 goldenrod:DAA520 gray:808080 green:008000 greenYellow:ADFF2F honeydew:F0FFF0 hotPink:FF69B4 indianRed:CD5C5C indigo:4B0082 ivory:FFFFF0 khaki:F0E68C lavender:E6E6FA lavenderBlush:FFF0F5 lawnGreen:7CFC00 lemonChiffon:FFFACD ltBlue:ADD8E6 ltCoral:F08080 ltCyan:E0FFFF ltGoldenrodYellow:FAFAD2 ltGray:D3D3D3 ltGreen:90EE90 ltPink:FFB6C1 ltSalmon:FFA07A ltSeaGreen:20B2AA ltSkyBlue:87CEFA ltSlateGray:778899 ltSteelBlue:B0C4DE ltYellow:FFFFE0 lime:00FF00 limeGreen:32CD32 linen:FAF0E6 magenta:FF00FF maroon:800000 medAquamarine:66CDAA medBlue:0000CD medOrchid:BA55D3 medPurple:9370DB medSeaGreen:3CB371 medSlateBlue:7B68EE medSpringGreen:00FA9A medTurquoise:48D1CC medVioletRed:C71585 midnightBlue:191970 mintCream:F5FFFA mistyRose:FFE4E1 moccasin:FFE4B5 navajoWhite:FFDEAD navy:000080 oldLace:FDF5E6 olive:808000 oliveDrab:6B8E23 orange:FFA500 orangeRed:FF4500 orchid:DA70D6 paleGoldenrod:EEE8AA paleGreen:98FB98 paleTurquoise:AFEEEE paleVioletRed:DB7093 papayaWhip:FFEFD5 peachPuff:FFDAB9 peru:CD853F pink:FFC0CB plum:DDA0DD powderBlue:B0E0E6 purple:800080 red:FF0000 rosyBrown:BC8F8F royalBlue:4169E1 saddleBrown:8B4513 salmon:FA8072 sandyBrown:F4A460 seaGreen:2E8B57 seaShell:FFF5EE sienna:A0522D silver:C0C0C0 skyBlue:87CEEB slateBlue:6A5ACD slateGray:708090 snow:FFFAFA springGreen:00FF7F steelBlue:4682B4 tan:D2B48C teal:008080 thistle:D8BFD8 tomato:FF6347 turquoise:40E0D0 violet:EE82EE wheat:F5DEB3 white:FFFFFF whiteSmoke:F5F5F5 yellow:FFFF00 yellowGreen:9ACD32`.split(' ').map(pair => pair.split(':')))

// IEC sRGB transfer, also published in W3C CSS Color 4 §19:
// https://www.w3.org/TR/css-color-4/#color-conversion-code
const linear = (c: number): number => c <= .04045 ? c / 12.92 : ((c + .055) / 1.055) ** 2.4
const encoded = (c: number): number => c <= .0031308 ? 12.92 * c : 1.055 * c ** (1 / 2.4) - .055
function hexColor(hex: string | undefined): number[] | undefined {
  if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) return undefined
  return [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16) / 255)
}
function toHsl(rgb: number[]): number[] {
  const [r, g, b] = rgb, max = Math.max(r, g, b), min = Math.min(r, g, b), d = max - min, l = (max + min) / 2
  if (d < 1e-12) return [0, 0, l]
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4
  return [h / 6, d / (1 - Math.abs(2 * l - 1)), l]
}
function fromHsl(hsl: number[]): number[] {
  const [h, s, l] = [wrap(hsl[0]), clamp(hsl[1]), clamp(hsl[2])]
  const c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs((h * 6) % 2 - 1)), m = l - c / 2
  const rgb = h < 1 / 6 ? [c, x, 0] : h < 2 / 6 ? [x, c, 0] : h < 3 / 6 ? [0, c, x] : h < 4 / 6 ? [0, x, c] : h < 5 / 6 ? [x, 0, c] : [c, 0, x]
  return rgb.map(c => c + m)
}

/** Resolve at use time: phClr is the already transformed fillRef/lnRef color. */
export function resolveDrawingColor(color: ColorDefinition | undefined, theme?: ThemeContext, placeholder?: DrawingColor, issues: DrawingIssue[] = []): DrawingColor | undefined {
  return resolveColor(color, theme, placeholder, issues, new Set())
}
function resolveColor(color: ColorDefinition | undefined, theme: ThemeContext | undefined, placeholder: DrawingColor | undefined, issues: DrawingIssue[], visiting: Set<string>): DrawingColor | undefined {
  if (!color) return undefined
  let rgb: number[] | undefined, alpha = 1, hsl: number[] | undefined
  if (color.kind === 'srgb') rgb = hexColor(color.value)
  else if (color.kind === 'system') rgb = hexColor(color.lastColor) ?? hexColor(({ windowText: '000000', window: 'FFFFFF', btnText: '000000', btnFace: 'F0F0F0' } as Record<string, string>)[color.value ?? ''])
  else if (color.kind === 'preset') rgb = hexColor(presetColors[color.value ?? ''])
  else if (color.kind === 'scrgb') {
    const channels = color.channels?.map(percentage)
    if (channels?.every(c => c !== undefined)) rgb = (channels as number[]).map(c => encoded(clamp(c)))
  } else if (color.kind === 'hsl') {
    const h = numeric(color.channels?.[0]), s = percentage(color.channels?.[1]), l = percentage(color.channels?.[2])
    if (h !== undefined && s !== undefined && l !== undefined) { hsl = [wrap(h / 21600000), clamp(s), clamp(l)]; rgb = fromHsl(hsl) }
  } else if (color.kind === 'scheme') {
    const name = color.value ?? ''
    if (name === 'phClr') { if (placeholder) { rgb = [placeholder.r, placeholder.g, placeholder.b].map(c => c / 255); alpha = placeholder.a } }
    else {
      const mapped = theme?.colorMap[name] ?? name
      if (!visiting.has(mapped) && theme?.colors[mapped]) {
        visiting.add(mapped)
        const base = resolveColor(theme.colors[mapped], theme, placeholder, issues, visiting)
        visiting.delete(mapped)
        if (base) { rgb = [base.r, base.g, base.b].map(c => c / 255); alpha = base.a }
      }
    }
  }
  if (!rgb || rgb.some(c => !Number.isFinite(c)) || !Number.isFinite(alpha)) { issue(issues, 'invalid-color', `Cannot resolve ${color.kind} color ${color.value ?? ''}`, color.value); return undefined }
  for (const transform of color.transforms) {
    const { name, value } = transform
    const channel = /^(red|green|blue)(Mod|Off)?$/.exec(name)
    const hslChannel = /^(hue|sat|lum)(Mod|Off)?$/.exec(name)
    if (!channel && !hslChannel && !['alpha', 'alphaMod', 'alphaOff', 'tint', 'shade'].includes(name)) {
      issue(issues, 'unsupported-color-transform', `Color transform ${name} is deferred`, name)
      continue
    }
    const n = hslChannel?.[1] === 'hue' && hslChannel[2] !== 'Mod' ? numeric(value) : percentage(value)
    if (n === undefined) { issue(issues, 'invalid-color-transform', `Invalid ${name} value ${value ?? ''}`, name); continue }
    if (name === 'alpha') alpha = clamp(n)
    else if (name === 'alphaMod') alpha = clamp(alpha * n)
    else if (name === 'alphaOff') alpha = clamp(alpha + n)
    // DrawingML tint is the proportion of source retained (10% -> 90% white).
    // ISO 29500 §20.1.2.3.34/31, Microsoft OpenXML Tint/Shade documentation.
    else if (name === 'tint' || name === 'shade') { hsl = undefined; rgb = rgb.map(c => encoded(clamp(name === 'shade' ? linear(c) * clamp(n) : linear(c) * clamp(n) + 1 - clamp(n)))) }
    else if (channel) {
      hsl = undefined
      const i = { red: 0, green: 1, blue: 2 }[channel[1] as 'red' | 'green' | 'blue'], before = linear(rgb[i])
      rgb[i] = encoded(clamp(channel[2] === 'Mod' ? before * n : channel[2] === 'Off' ? before + n : n))
    } else if (hslChannel) {
      hsl ??= toHsl(rgb)
      const i = { hue: 0, sat: 1, lum: 2 }[hslChannel[1] as 'hue' | 'sat' | 'lum']
      const v = i === 0 && hslChannel[2] !== 'Mod' ? n / 21600000 : n
      const result = hslChannel[2] === 'Mod' ? hsl[i] * v : hslChannel[2] === 'Off' ? hsl[i] + v : v
      hsl[i] = i === 0 ? wrap(result) : clamp(result)
      rgb = fromHsl(hsl)
    }
  }
  const [r, g, b] = rgb.map(c => Math.round(clamp(c) * 255 + 1e-8))
  return { r, g, b, a: alpha }
}

function parseRect(node: XmlNode | undefined, issues: DrawingIssue[]): Partial<RelativeRect> | undefined {
  if (!node) return undefined
  const rect: Partial<RelativeRect> = {}
  for (const [xmlName, name] of [['l', 'left'], ['t', 'top'], ['r', 'right'], ['b', 'bottom']] as const) {
    const raw = attrs(node)[xmlName]
    if (raw === undefined) continue
    const n = percentage(raw)
    if (n === undefined) issue(issues, 'invalid-fill', `Invalid gradient ${xmlName} offset`, xmlName)
    else rect[name] = n
  }
  return rect
}
/** Parse fields without filling defaults, allowing property-by-property inheritance. */
export function parseFillDefinition(node: XmlNode | undefined, issues: DrawingIssue[] = []): FillDefinition | undefined {
  for (const [name, child] of orderedChildren(node)) {
    if (name === 'noFill') return { kind: 'none' }
    if (name === 'solidFill') { const color = parseDrawingColor(child); return { kind: 'solid', ...(color ? { color } : {}) } }
    if (name === 'gradFill') {
      const fill: Extract<FillDefinition, { kind: 'gradient' }> = { kind: 'gradient' }
      const list = getChild(child, 'gsLst')
      if (list) {
        fill.stops = []
        for (const stop of getChildren(list, 'gs')) {
          const position = percentage(attrs(stop).pos), color = parseDrawingColor(stop)
          if (position === undefined || !color) issue(issues, 'invalid-fill', 'Invalid gradient stop', 'gs')
          else fill.stops.push({ position: clamp(position), color })
        }
        fill.stops.sort((a, b) => a.position - b.position)
      }
      const lin = getChild(child, 'lin'), path = getChild(child, 'path')
      if (lin) {
        fill.gradient = 'linear'
        const angle = attrs(lin).ang
        if (angle !== undefined) { const n = numeric(angle); if (n !== undefined) fill.angle = n * Math.PI / 10800000; else issue(issues, 'invalid-fill', 'Invalid linear gradient angle', 'ang') }
        if (attrs(lin).scaled !== undefined) fill.scaled = ['true', '1'].includes(attrs(lin).scaled)
      } else if (path) {
        fill.gradient = attrs(path).path ?? 'shape'
        const focus = parseRect(getChild(path, 'fillToRect'), issues)
        if (focus) fill.focus = focus
      }
      if (attrs(child).rotWithShape !== undefined) fill.rotateWithShape = ['true', '1'].includes(attrs(child).rotWithShape)
      if (getChild(child, 'tileRect')) issue(issues, 'unsupported-gradient', 'Gradient tiling is deferred; use shape extents', 'tileRect')
      if (attrs(child).flip && attrs(child).flip !== 'none') issue(issues, 'unsupported-gradient', 'Gradient tile flipping is deferred', 'flip')
      return fill
    }
    if (['blipFill', 'pattFill', 'grpFill'].includes(name)) { issue(issues, 'unsupported-fill', `Fill ${name} is deferred`, name); return { kind: 'none' } }
  }
  return undefined
}
function parseEnd(node: XmlNode | undefined, issues: DrawingIssue[]): ArrowEnd | undefined {
  if (!node) return undefined
  const end: ArrowEnd = {}, a = attrs(node)
  if (a.type !== undefined) {
    if (['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'].includes(a.type)) end.type = a.type as ArrowEnd['type']
    else issue(issues, 'unsupported-line', `Unknown arrow type ${a.type}`, a.type)
  }
  for (const [attribute, field] of [['w', 'width'], ['len', 'length']] as const) {
    if (a[attribute] === undefined) continue
    if (['sm', 'med', 'lg'].includes(a[attribute])) end[field] = a[attribute] as 'sm' | 'med' | 'lg'
    else issue(issues, 'invalid-line', `Invalid arrow ${attribute}`, attribute)
  }
  return end
}
export function parseLineDefinition(node: XmlNode | undefined, issues: DrawingIssue[] = []): LineDefinition | undefined {
  if (!node) return undefined
  const line: LineDefinition = {}, a = attrs(node)
  if (a.w !== undefined) { const w = numeric(a.w); if (w !== undefined && w >= 0) line.width = w / 9525; else issue(issues, 'invalid-line', 'Invalid line width', 'w') }
  if (a.cap !== undefined) {
    const caps: Record<string, CanvasLineCap> = { flat: 'butt', rnd: 'round', sq: 'square' }
    if (has(caps, a.cap)) line.cap = caps[a.cap]
    else issue(issues, 'invalid-line', 'Invalid line cap', 'cap')
  }
  if (a.cmpd && a.cmpd !== 'sng') issue(issues, 'unsupported-line', 'Compound strokes are deferred', 'cmpd')
  if (a.algn && a.algn !== 'ctr') issue(issues, 'unsupported-line', 'Inset strokes are deferred', 'algn')
  const fill = parseFillDefinition(node, issues)
  if (fill) line.fill = fill
  const dash = getChild(node, 'prstDash'), custom = getChild(node, 'custDash')
  if (dash) line.dash = attrs(dash).val ?? 'solid'
  if (custom) {
    line.dash = []
    for (const ds of getChildren(custom, 'ds')) {
      const d = percentage(attrs(ds).d), gap = percentage(attrs(ds).sp)
      if (d === undefined || gap === undefined || d < 0 || gap < 0 || d + gap <= 0) issue(issues, 'invalid-line', 'Invalid custom dash lengths', 'custDash')
      else line.dash.push(d, gap)
    }
  }
  if (getChild(node, 'round')) line.join = 'round'
  if (getChild(node, 'bevel')) line.join = 'bevel'
  const miter = getChild(node, 'miter')
  if (miter) {
    line.join = 'miter'
    if (attrs(miter).lim !== undefined) {
      const limit = percentage(attrs(miter).lim)
      if (limit !== undefined && limit > 0) line.miterLimit = limit
      else issue(issues, 'invalid-line', 'Invalid miter limit', 'miter')
    }
  }
  const head = parseEnd(getChild(node, 'headEnd'), issues), tail = parseEnd(getChild(node, 'tailEnd'), issues)
  if (head) line.headEnd = head
  if (tail) line.tailEnd = tail
  return line
}
function parseFonts(node: XmlNode | undefined): ThemeFonts {
  const fonts: ThemeFonts = { supplemental: {} }
  for (const [xmlName, field] of [['latin', 'latin'], ['ea', 'eastAsian'], ['cs', 'complexScript']] as const) {
    const value = attrs(getChild(node, xmlName)).typeface
    if (value !== undefined) fonts[field] = value
  }
  for (const font of getChildren(node, 'font')) { const { script, typeface } = attrs(font); if (script && typeface !== undefined) fonts.supplemental[script] = typeface }
  return fonts
}
/** Parse with ordered XML so repeated color transforms retain source order. */
export function parseThemeContext(input?: string | XmlNode, colorMap: Record<string, string> = {}): ThemeContext {
  const theme: ThemeContext = { colors: {}, palette: {}, colorMap: { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2', ...colorMap }, fonts: { major: { supplemental: {} }, minor: { supplemental: {} } }, fillStyles: [], bgFillStyles: [], lineStyles: [], effectStyles: [], issues: [] }
  let node: XmlNode | undefined
  try { node = typeof input === 'string' ? parseXmlOrdered(input) : input } catch { issue(theme.issues, 'invalid-theme', 'Malformed theme XML'); return theme }
  const elements = getChild(node, 'themeElements') ?? node
  for (const [slot, child] of orderedChildren(getChild(elements, 'clrScheme'))) {
    if (slot === '#text') continue
    const color = parseDrawingColor(child)
    if (color) theme.colors[slot] = color
  }
  for (const [slot, definition] of Object.entries(theme.colors)) {
    const c = resolveDrawingColor(definition, theme, undefined, theme.issues)
    if (c) theme.palette[slot] = `#${[c.r, c.g, c.b].map(v => v.toString(16).padStart(2, '0')).join('')}`
  }
  const fontScheme = getChild(elements, 'fontScheme')
  theme.fonts.major = parseFonts(getChild(fontScheme, 'majorFont'))
  theme.fonts.minor = parseFonts(getChild(fontScheme, 'minorFont'))
  const format = getChild(elements, 'fmtScheme')
  for (const [list, field] of [['fillStyleLst', 'fillStyles'], ['bgFillStyleLst', 'bgFillStyles']] as const) {
    for (const [name, child] of orderedChildren(getChild(format, list))) {
      if (name === '#text') continue
      const fill = parseFillDefinition({ [name]: child }, theme.issues)
      theme[field].push(fill ?? { kind: 'none' })
    }
  }
  for (const child of getChildren(getChild(format, 'lnStyleLst'), 'ln')) theme.lineStyles.push(parseLineDefinition(child, theme.issues) ?? {})
  for (const child of getChildren(getChild(format, 'effectStyleLst'), 'effectStyle')) {
    const effects = orderedChildren(getChild(child, 'effectLst')).filter(([n]) => n !== '#text').map(([n]) => n)
    if (getChild(child, 'effectDag')) effects.push('effectDag')
    if (getChild(child, 'scene3d')) effects.push('scene3d')
    if (getChild(child, 'sp3d')) effects.push('sp3d')
    theme.effectStyles.push(effects)
  }
  return theme
}
function mergeFill(base: FillDefinition | undefined, direct: FillDefinition | undefined): FillDefinition | undefined {
  if (!direct) return base
  if (!base || base.kind !== direct.kind || direct.kind === 'none') return direct
  if (base.kind === 'gradient' && direct.kind === 'gradient') return { ...base, ...direct, ...(base.focus || direct.focus ? { focus: { ...base.focus, ...direct.focus } } : {}) }
  return { ...base, ...direct } as FillDefinition
}
export function resolveFill(definition: FillDefinition | undefined, theme?: ThemeContext, placeholder?: DrawingColor, issues: DrawingIssue[] = []): DrawingFill | undefined {
  if (!definition || definition.kind === 'none') return definition
  if (definition.kind === 'solid') {
    const color = resolveDrawingColor(definition.color, theme, placeholder, issues)
    return color ? { kind: 'solid', color } : { kind: 'none' }
  }
  const stops = (definition.stops ?? []).flatMap(stop => {
    const color = resolveDrawingColor(stop.color, theme, placeholder, issues)
    return color ? [{ position: stop.position, color }] : []
  })
  if (!stops.length) { issue(issues, 'invalid-fill', 'Gradient has no usable stops', 'gsLst'); return { kind: 'none' } }
  const gradient = definition.gradient ?? 'linear'
  if (gradient !== 'linear' && gradient !== 'circle') {
    issue(issues, 'unsupported-gradient', `${gradient} gradients use the first stop in this foundation stage`, gradient)
    return { kind: 'solid', color: stops[0].color }
  }
  const focus = gradient === 'circle' ? { left: .5, top: .5, right: .5, bottom: .5, ...definition.focus } : undefined
  if (definition.rotateWithShape === false) issue(issues, 'unsupported-gradient', 'Nonrotating gradients require adapter transform compensation; local rotation used', 'rotWithShape')
  return { kind: 'gradient', gradient, stops, angle: definition.angle ?? 0, scaled: definition.scaled ?? false, ...(focus ? { focus } : {}), ...(definition.rotateWithShape !== undefined ? { rotateWithShape: definition.rotateWithShape } : {}) }
}
function referenceIndex(node: XmlNode, issues: DrawingIssue[]): number | undefined {
  const index = numeric(attrs(node).idx)
  if (index === undefined || index < 0 || !Number.isInteger(index)) issue(issues, 'missing-theme-style', 'Invalid theme style index', attrs(node).idx)
  return index
}
/** Resolve direct spPr over theme fillRef/lnRef. Absent fields remain inheritable. */
export function resolveDrawingStyle(spPr?: XmlNode, style?: XmlNode, theme?: ThemeContext, defaults: Partial<DrawingStyle> = {}): DrawingStyle {
  const issues: DrawingIssue[] = [...(defaults.issues ?? [])]
  let inheritedFill: FillDefinition | undefined, inheritedLine: LineDefinition | undefined, fillPlaceholder: DrawingColor | undefined, linePlaceholder: DrawingColor | undefined
  const fillRef = getChild(style, 'fillRef'), lineRef = getChild(style, 'lnRef')
  if (fillRef) {
    const index = referenceIndex(fillRef, issues)
    fillPlaceholder = resolveDrawingColor(parseDrawingColor(fillRef), theme, undefined, issues)
    if (index === 0) inheritedFill = { kind: 'none' }
    else if (index !== undefined) {
      inheritedFill = index >= 1001 && index <= 1003 ? theme?.bgFillStyles[index - 1001] : index >= 1 && index <= 3 ? theme?.fillStyles[index - 1] : undefined
      if (!inheritedFill) issue(issues, 'missing-theme-style', `Missing fill style ${index}`, String(index))
    }
  }
  if (lineRef) {
    const index = referenceIndex(lineRef, issues)
    linePlaceholder = resolveDrawingColor(parseDrawingColor(lineRef), theme, undefined, issues)
    if (index === 0) inheritedLine = { fill: { kind: 'none' } }
    else if (index !== undefined) {
      inheritedLine = index >= 1 && index <= 3 ? theme?.lineStyles[index - 1] : undefined
      if (!inheritedLine) issue(issues, 'missing-theme-style', `Missing line style ${index}`, String(index))
    }
  }
  const definition = mergeFill(inheritedFill, parseFillDefinition(spPr, issues))
  const fill = definition ? resolveFill(definition, theme, fillPlaceholder, issues) : defaults.fill
  const directLine = parseLineDefinition(getChild(spPr, 'ln'), issues)
  let line: DrawingLine | undefined = defaults.line ? { ...defaults.line } : undefined
  if (inheritedLine || directLine) {
    const merged = { ...inheritedLine, ...directLine }
    const mergedFill = mergeFill(inheritedLine?.fill, directLine?.fill)
    const resolvedFill = resolveFill(mergedFill, theme, linePlaceholder, issues)
    const { fill: _definition, ...fields } = merged
    line = { width: 1, ...defaults.line, ...fields, fill: resolvedFill ?? defaults.line?.fill }
    for (const end of ['headEnd', 'tailEnd'] as const) {
      if (inheritedLine?.[end] || directLine?.[end]) line[end] = { type: 'none', width: 'med', length: 'med', ...defaults.line?.[end], ...inheritedLine?.[end], ...directLine?.[end] }
    }
  }
  for (const name of ['effectLst', 'effectDag']) {
    const effects = getChild(spPr, name)
    if (effects && orderedChildren(effects).some(([n]) => n !== '#text')) issue(issues, 'unsupported-effect', 'Drawing effects are deferred', name)
  }
  for (const name of ['sp3d', 'scene3d']) if (getChild(spPr, name)) issue(issues, 'unsupported-3d', 'Drawing 3D effects are deferred', name)
  const effectRef = getChild(style, 'effectRef')
  if (effectRef) {
    const index = referenceIndex(effectRef, issues)
    if (index && theme?.effectStyles[index - 1]?.length) issue(issues, 'unsupported-effect', `Theme effect style ${index} is deferred`, 'effectRef')
  }
  return { ...(fill ? { fill } : {}), ...(line ? { line } : {}), issues }
}
