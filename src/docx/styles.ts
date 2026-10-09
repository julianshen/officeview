import { attrs, elementChildren, getChildren, orderedChildren, type XmlNode } from '../core/xml'
import type { DocxParagraph, DocxTextRun } from './types'

export interface DocxTheme {
  colors: Map<string, string>
  fonts: Map<string, string>
}
export interface DocxStyleContext {
  theme: DocxTheme
  styles: Map<string, XmlNode>
  defaultParagraph?: string
  defaultTable?: string
  paragraphDefaults?: XmlNode
  runDefaults?: XmlNode
}

/** Preserve the font's generic family when Office fonts are not installed. */
export function fontFamilyCss(family: string): string {
  const serif =
    /^(Cambria(?: Math)?|Constantia|Times New Roman|Georgia|Garamond|Palatino(?: Linotype)?|Book Antiqua)$/i.test(
      family
    )
  return `${JSON.stringify(family)}, ${serif ? 'serif' : 'sans-serif'}`
}

export function readTheme(root: XmlNode | undefined): DocxTheme {
  const theme: DocxTheme = { colors: new Map(), fonts: new Map() }
  const elements = getChildren(root, 'themeElements')[0]
  for (const [name, color] of elementChildren(getChildren(elements, 'clrScheme')[0])) {
    const value = attrs(getChildren(color, 'srgbClr')[0]).val ?? attrs(getChildren(color, 'sysClr')[0]).lastClr
    if (value) theme.colors.set(name, value)
  }
  for (const [alias, name] of [
    ['tx1', 'dk1'],
    ['tx2', 'dk2'],
    ['bg1', 'lt1'],
    ['bg2', 'lt2']
  ]) {
    const value = theme.colors.get(name)
    if (value) theme.colors.set(alias, value)
  }
  const scheme = getChildren(elements, 'fontScheme')[0]
  for (const family of ['major', 'minor']) {
    const font = getChildren(scheme, `${family}Font`)[0]
    const latin = attrs(getChildren(font, 'latin')[0]).typeface
    if (latin) {
      theme.fonts.set(`${family}HAnsi`, latin)
      theme.fonts.set(`${family}Ascii`, latin)
      theme.fonts.set(family, latin)
      theme.fonts.set(`+${family === 'major' ? 'mj' : 'mn'}-lt`, latin)
    }
    const ea = attrs(getChildren(font, 'ea')[0]).typeface
    if (ea) {
      theme.fonts.set(`${family}EastAsia`, ea)
      theme.fonts.set(`+${family === 'major' ? 'mj' : 'mn'}-ea`, ea)
    }
    const cs = attrs(getChildren(font, 'cs')[0]).typeface
    if (cs) {
      theme.fonts.set(`${family}Bidi`, cs)
      theme.fonts.set(`${family}Cs`, cs)
      theme.fonts.set(`+${family === 'major' ? 'mj' : 'mn'}-cs`, cs)
    }
  }
  return theme
}

export function styleContext(root: XmlNode | undefined, theme: DocxTheme): DocxStyleContext {
  const out: DocxStyleContext = { theme, styles: new Map() }
  const defaults = getChildren(root, 'docDefaults')[0]
  out.paragraphDefaults = getChildren(getChildren(defaults, 'pPrDefault')[0], 'pPr')[0]
  out.runDefaults = getChildren(getChildren(defaults, 'rPrDefault')[0], 'rPr')[0]
  for (const style of getChildren(root, 'style')) {
    const a = attrs(style)
    if (a.styleId) out.styles.set(a.styleId, style)
    if (a.default === '1' || a.default === 'true') {
      if (a.type === 'paragraph') out.defaultParagraph = a.styleId
      if (a.type === 'table') out.defaultTable = a.styleId
    }
  }
  return out
}

/** Base styles first; a broken/cyclic basedOn chain terminates at its first repeat. */
export function styleChain(id: string | undefined, context: DocxStyleContext): XmlNode[] {
  const out: XmlNode[] = [],
    visited = new Set<string>()
  while (id && !visited.has(id)) {
    visited.add(id)
    const style = context.styles.get(id)
    if (!style) break
    out.unshift(style)
    id = attrs(getChildren(style, 'basedOn')[0]).val
  }
  return out
}

import { parseDrawingColor, resolveDrawingColor, parseThemeContext, type ThemeContext, type DrawingColor } from '../drawing/style'
import type { TextAppearanceIssue } from '../drawing/text-parse'

const themeContextCache = new WeakMap<DocxTheme, ThemeContext>()

function getThemeContext(theme?: DocxTheme): ThemeContext | undefined {
  if (!theme) return undefined
  let ctx = themeContextCache.get(theme)
  if (!ctx) {
    ctx = parseThemeContext()
    for (const [name, value] of theme.colors) {
      if (/^[\da-f]{6}$/i.test(value)) {
        ctx.colors[name] = { kind: 'srgb', value, transforms: [] }
      }
    }
    themeContextCache.set(theme, ctx)
  }
  return ctx
}

function parseAngleAttribute(
  raw: string | undefined,
  issues?: TextAppearanceIssue[],
  feature?: string
): number | undefined {
  if (raw === undefined) return 0
  const n = Number(raw)
  if (!Number.isFinite(n) || n < -21600000 || n > 21600000) {
    if (issues && feature) {
      issues.push({
        kind: 'unsupported-text-appearance',
        feature,
        message: `WordArt angle ${raw} is out of range; using fallback`,
      })
    }
    return undefined
  }
  const rad = (n / 10800000) * Math.PI
  return Number.isFinite(rad) ? rad : 0
}

function textCssColor(color: DrawingColor): string {
  const rgb = [color.r, color.g, color.b]
  return color.a < 1 ? `rgba(${rgb.join(',')},${color.a})` : `#${rgb.map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase()}`
}

export function readRunProperties(rPr: XmlNode | undefined, theme?: DocxTheme, issues?: TextAppearanceIssue[]): Partial<DocxTextRun> {
  const out: Partial<DocxTextRun> = {}
  const fonts = attrs(getChildren(rPr, 'rFonts')[0])
  const asciiFont = fonts.ascii ?? theme?.fonts.get(fonts.asciiTheme)
  const hAnsiFont = fonts.hAnsi ?? theme?.fonts.get(fonts.hAnsiTheme)
  const eaFont = fonts.eastAsia ?? theme?.fonts.get(fonts.eastAsiaTheme)
  const csFont = fonts.cs ?? theme?.fonts.get(fonts.cstheme)
  const isRtl = getChildren(rPr, 'rtl').length > 0 || getChildren(rPr, 'cs').length > 0
  const preferred = (isRtl || fonts.hint === 'cs')
    ? (csFont ?? asciiFont ?? hAnsiFont ?? eaFont)
    : fonts.hint === 'eastAsia'
      ? (eaFont ?? asciiFont ?? hAnsiFont ?? csFont)
      : (asciiFont ?? hAnsiFont ?? eaFont ?? csFont)
  if (preferred) out.fontFamily = preferred
  const size = Number(attrs(getChildren(rPr, 'sz')[0]).val)
  if (Number.isFinite(size) && size > 0) out.fontSizePt = size / 2
  for (const [element, property] of [
    ['b', 'bold'],
    ['i', 'italic'],
    ['strike', 'strike']
  ] as const) {
    const node = getChildren(rPr, element)[0]
    if (node) out[property] = !['0', 'false', 'off'].includes(attrs(node).val)
  }
  const underline = getChildren(rPr, 'u')[0]
  if (underline) out.underline = attrs(underline).val !== 'none'
  const color = attrs(getChildren(rPr, 'color')[0])
  if (color.val !== undefined || color.themeColor) {
    out.color = color.val === 'auto' ? undefined : (color.val ?? theme?.colors.get(color.themeColor))
  }
  const highlight = attrs(getChildren(rPr, 'highlight')[0]).val
  if (highlight) out.highlight = highlight
  const vertAlign = attrs(getChildren(rPr, 'vertAlign')[0]).val
  if (vertAlign === 'superscript' || vertAlign === 'subscript') {
    out.vertAlign = vertAlign
  }

  const textFill = getChildren(rPr, 'textFill')[0]
  if (textFill) {
    const rawChildren = orderedChildren(textFill).filter(([name]) => name !== '#text')
    if (rawChildren.length === 0) {
      // Explicit truly empty <w14:textFill/> element: CT_FillTextEffect specification:
      // "If this element has no child elements, a default of solid black fill is applied."
      // Overrides inherited/direct nonblack color, clears inherited gradient and noFill.
      out.color = '000000'
      out.textFill = undefined
      out.noFill = false
    } else {
      const grad = getChildren(textFill, 'gradFill')[0]
      const patt = getChildren(textFill, 'pattFill')[0]
      const blip = getChildren(textFill, 'blipFill')[0]
      const solid = getChildren(textFill, 'solidFill')[0]
      const noFill = getChildren(textFill, 'noFill')[0]

      if (noFill) {
        out.noFill = true
        out.textFill = undefined
      } else if (solid) {
        out.noFill = false
        out.textFill = undefined
        const themeCtx = getThemeContext(theme)
        const c = resolveDrawingColor(parseDrawingColor(solid), themeCtx)
        if (c) out.color = textCssColor(c)
      } else if (grad) {
        const gsLst = getChildren(grad, 'gsLst')[0]
        if (!gsLst) {
          // CT_GradientFillProperties: absent gsLst defaults to solid black fill.
          // Overrides inherited/direct nonblack color, clears inherited gradient and noFill.
          out.color = '000000'
          out.textFill = undefined
          out.noFill = false
        } else {
          const themeCtx = getThemeContext(theme)
          const stops = getChildren(gsLst, 'gs')
            .map(g => {
              const c = resolveDrawingColor(parseDrawingColor(g), themeCtx)
              const pos = attrs(g).pos
              const posNum = pos !== undefined && Number.isFinite(Number(pos)) ? Number(pos) : undefined
              return c && posNum !== undefined ? { position: posNum / 100000, color: textCssColor(c) } : undefined
            })
            .filter((s): s is { position: number; color: string } => !!s)
          const lin = getChildren(grad, 'lin')[0]
          const pathGrad = getChildren(grad, 'path')[0]
          if (pathGrad) {
            if (issues) issues.push({ kind: 'unsupported-text-appearance', feature: 'gradFill:path', message: 'WordArt path gradient is deferred; using first stop color' })
            if (stops.length > 0 && !out.color) out.color = stops[0].color
            out.textFill = undefined
          } else if (stops.length >= 2) {
            // Absent lin/path defaults to linear angle 0: CT_GradientFillProperties
            const angRaw = lin ? attrs(lin).ang : undefined
            const angle = parseAngleAttribute(angRaw, issues, 'gradFill:ang')
            if (angle !== undefined) {
              out.textFill = { kind: 'gradient', stops, angle }
            } else {
              if (stops.length > 0 && !out.color) out.color = stops[0].color
              out.textFill = undefined
            }
          } else {
            if (issues) issues.push({ kind: 'unsupported-text-appearance', feature: 'gradFill', message: 'WordArt gradient needs 2+ stops; using flat color' })
            if (stops.length > 0 && !out.color) out.color = stops[0].color
            out.textFill = undefined
          }
          out.noFill = false
        }
      } else if (patt) {
        const themeCtx = getThemeContext(theme)
        const fg = resolveDrawingColor(parseDrawingColor(getChildren(patt, 'fgClr')[0]), themeCtx)
        if (issues) issues.push({ kind: 'unsupported-text-appearance', feature: 'pattFill', message: 'WordArt pattern fill is unsupported; using flat color' })
        out.color = fg ? textCssColor(fg) : (out.color ?? '#000000')
        out.textFill = undefined
        out.noFill = false
      } else if (blip) {
        if (issues) issues.push({ kind: 'unsupported-text-appearance', feature: 'blipFill', message: 'WordArt picture fill is deferred; using flat color' })
        out.textFill = undefined
        out.noFill = false
      } else {
        const childName = rawChildren[0]?.[0] ?? 'unknownFill'
        if (issues) issues.push({ kind: 'unsupported-text-appearance', feature: `textFill:${childName}`, message: `WordArt fill ${childName} is unsupported; using flat fallback` })
        out.textFill = undefined
        out.noFill = false
      }
    }
  }

  const textOutline = getChildren(rPr, 'textOutline')[0]
  if (textOutline) {
    const wRaw = attrs(textOutline).w
    const wNum = wRaw !== undefined && Number.isFinite(Number(wRaw)) ? Number(wRaw) : undefined
    const solidLine = getChildren(textOutline, 'solidFill')[0]
    const themeCtx = getThemeContext(theme)
    const lineColor = solidLine ? resolveDrawingColor(parseDrawingColor(solidLine), themeCtx) : undefined
    if (lineColor && wNum !== undefined && wNum > 0) {
      out.textOutline = { color: textCssColor(lineColor), widthPx: Math.min((wNum * 96) / (12700 * 72), 100) }
    } else if (getChildren(textOutline, 'noFill').length) {
      out.textOutline = undefined
    } else {
      if (issues) issues.push({ kind: 'unsupported-text-appearance', feature: 'textOutline', message: 'WordArt outline needs a finite positive width and solid color; skipped' })
      out.textOutline = undefined
    }
  }

  const effectLst = getChildren(rPr, 'effectLst')[0]
  const glow = (effectLst ? getChildren(effectLst, 'glow')[0] : undefined) ?? getChildren(rPr, 'glow')[0]
  if (glow && issues) {
    issues.push({ kind: 'unsupported-text-appearance', feature: 'glow', message: 'WordArt text glow effect is unsupported; falling back to plain run' })
  }
  const reflection = (effectLst ? (getChildren(effectLst, 'reflection')[0] ?? getChildren(effectLst, 'refl')[0]) : undefined) ?? getChildren(rPr, 'reflection')[0] ?? getChildren(rPr, 'refl')[0]
  if (reflection && issues) {
    issues.push({ kind: 'unsupported-text-appearance', feature: 'reflection', message: 'WordArt text reflection effect is unsupported; falling back to plain run' })
  }

  const shadow = getChildren(rPr, 'shadow')[0] ?? (effectLst ? getChildren(effectLst, 'outerShdw')[0] : undefined) ?? getChildren(rPr, 'outerShdw')[0]
  if (shadow) {
    const sa = attrs(shadow)
    const themeCtx = getThemeContext(theme)
    const shadowColor = resolveDrawingColor(parseDrawingColor(shadow), themeCtx)
    if (shadowColor) {
      const distRaw = sa.dist
      const distNum = distRaw !== undefined ? Number(distRaw) : 0
      const dirAngle = parseAngleAttribute(sa.dir, issues, 'shadow:dir')

      if (dirAngle === undefined || !Number.isFinite(distNum) || distNum < -100000000 || distNum > 100000000) {
        if (issues && dirAngle !== undefined) {
          issues.push({ kind: 'unsupported-text-appearance', feature: 'shadow:dist', message: 'WordArt shadow distance is out of range; skipped' })
        }
        out.textShadow = undefined
      } else {
        const distPx = distNum / 9525
        const blurRaw = sa.blurRad ?? sa.bluRad
        const blurNum = blurRaw !== undefined ? Number(blurRaw) : undefined
        const blurPx = blurNum !== undefined && Number.isFinite(blurNum) ? Math.min(Math.max(0, blurNum / 9525), 100) : 0
        const cos = Math.cos(dirAngle)
        const sin = Math.sin(dirAngle)
        const offsetX = Number.isFinite(cos * distPx) ? cos * distPx : 0
        const offsetY = Number.isFinite(sin * distPx) ? sin * distPx : 0
        out.textShadow = {
          color: textCssColor(shadowColor),
          blurPx,
          offsetX,
          offsetY,
        }
      }
    } else {
      if (issues) issues.push({ kind: 'unsupported-text-appearance', feature: 'shadow', message: 'WordArt shadow needs a resolvable color; skipped' })
    }
  }

  return out
}

export type AppearanceCategory = 'fill' | 'outline' | 'shadow' | 'glow' | 'reflection'

export function authoredCategories(rPr: XmlNode | undefined): Set<AppearanceCategory> {
  const set = new Set<AppearanceCategory>()
  if (!rPr) return set
  if (getChildren(rPr, 'textFill').length > 0) {
    set.add('fill')
  }
  if (getChildren(rPr, 'textOutline').length > 0) {
    set.add('outline')
  }
  const effectLst = getChildren(rPr, 'effectLst')[0]
  if (getChildren(rPr, 'shadow').length > 0 || getChildren(rPr, 'outerShdw').length > 0 || (effectLst && getChildren(effectLst, 'outerShdw').length > 0)) {
    set.add('shadow')
  }
  if (getChildren(rPr, 'glow').length > 0 || (effectLst && getChildren(effectLst, 'glow').length > 0)) {
    set.add('glow')
  }
  if (getChildren(rPr, 'reflection').length > 0 || getChildren(rPr, 'refl').length > 0 || (effectLst && (getChildren(effectLst, 'reflection').length > 0 || getChildren(effectLst, 'refl').length > 0))) {
    set.add('reflection')
  }
  return set
}

export function issueCategory(feature: string): 'fill' | 'outline' | 'shadow' | 'glow' | 'reflection' {
  if (feature.startsWith('gradFill') || feature.startsWith('pattFill') || feature.startsWith('blipFill') || feature.startsWith('textFill')) {
    return 'fill'
  }
  if (feature === 'textOutline' || feature.startsWith('textOutline:')) return 'outline'
  if (feature === 'shadow' || feature.startsWith('shadow:') || feature === 'outerShdw' || feature.startsWith('outerShdw:')) return 'shadow'
  if (feature === 'glow' || feature.startsWith('glow:')) return 'glow'
  if (feature === 'reflection' || feature.startsWith('reflection:') || feature === 'refl' || feature.startsWith('refl:')) return 'reflection'
  return 'fill'
}

/** Table layers sit below paragraph styles and above document defaults. */
export interface ParagraphStyleLayers { pPr: Array<XmlNode | undefined>; rPr: Array<XmlNode | undefined> }
export function paragraphRunDefaults(
  p: XmlNode,
  context?: DocxStyleContext,
  table?: ParagraphStyleLayers,
  inheritedIssues?: TextAppearanceIssue[]
): Partial<DocxTextRun> {
  if (!context) return {}
  const id = attrs(getChildren(getChildren(p, 'pPr')[0], 'pStyle')[0]).val ?? context.defaultParagraph
  const layerNodes: Array<XmlNode | undefined> = [
    context.runDefaults,
    ...(table?.rPr ?? []),
    ...styleChain(id, context).map(style => getChildren(style, 'rPr')[0])
  ]
  const activeIssues = new Map<'fill' | 'outline' | 'shadow' | 'glow' | 'reflection', TextAppearanceIssue>()
  const merged: Partial<DocxTextRun> = {}

  for (const node of layerNodes) {
    if (!node) continue
    const authored = authoredCategories(node)
    const layerIssues: TextAppearanceIssue[] = []
    const layerProps = readRunProperties(node, context.theme, layerIssues)
    Object.assign(merged, layerProps)

    for (const cat of ['fill', 'outline', 'shadow', 'glow', 'reflection'] as const) {
      if (authored.has(cat)) {
        const issue = layerIssues.find(i => issueCategory(i.feature) === cat)
        if (issue) {
          activeIssues.set(cat, issue)
        } else {
          // Explicit supported override or clear suppresses earlier issue in this category
          activeIssues.delete(cat)
        }
      }
    }
  }

  if (inheritedIssues) {
    for (const issue of activeIssues.values()) {
      if (!inheritedIssues.some(existing => existing.kind === issue.kind && existing.feature === issue.feature)) {
        inheritedIssues.push(issue)
      }
    }
  }

  return merged
}

export function applyParagraphDefaults(paragraph: DocxParagraph, p: XmlNode, context?: DocxStyleContext, table?: ParagraphStyleLayers): void {
  if (!context) return
  const direct = getChildren(p, 'pPr')[0]
  const id = attrs(getChildren(direct, 'pStyle')[0]).val ?? context.defaultParagraph
  let lineValue: number | undefined, lineRule: 'auto' | 'exact' | 'atLeast' = 'auto'
  for (const pPr of [context.paragraphDefaults, ...(table?.pPr ?? []), ...styleChain(id, context).map(style => getChildren(style, 'pPr')[0]), direct]) {
    const spacing = attrs(getChildren(pPr, 'spacing')[0])
    if (spacing.before !== undefined) paragraph.spacingBeforeTwips = Number(spacing.before)
    if (spacing.after !== undefined) paragraph.spacingAfterTwips = Number(spacing.after)
    if (spacing.line !== undefined) lineValue = Number(spacing.line)
    if (spacing.lineRule !== undefined) lineRule = spacing.lineRule as typeof lineRule
    const jc = attrs(getChildren(pPr, 'jc')[0]).val
    if (jc !== undefined) paragraph.align = jc === 'both' ? 'justify' : jc === 'center' || jc === 'right' ? jc : 'left'
    const ind = attrs(getChildren(pPr, 'ind')[0])
    for (const [key, value] of [['indentLeftTwips', ind.left ?? ind.start], ['indentRightTwips', ind.right ?? ind.end], ['indentFirstLineTwips', ind.firstLine ?? (ind.hanging !== undefined ? -Number(ind.hanging) : undefined)]] as const)
      if (value !== undefined) paragraph[key] = Number(value)
    const outline = attrs(getChildren(pPr, 'outlineLvl')[0]).val
    if (outline !== undefined) paragraph.outlineLevel = Number(outline)
  }
  if (lineValue !== undefined) paragraph.lineSpacing = { value: lineValue, rule: lineRule }
}
