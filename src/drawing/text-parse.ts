/** DrawingML text metadata shared by slides, cached diagrams, and sheet drawings. */
import { attrs, getChildren, orderedChildren, textOf, type XmlNode } from '../core/xml'
import { parseDrawingColor, resolveDrawingColor, type DrawingColor, type ThemeContext } from './style'
import type { DrawingTextParagraph as PptxParagraph, DrawingTextBody as PptxTextBody, DrawingTextRun as PptxTextRun, DrawingTextSpacing as PptxTextSpacing, DrawingTextStyle as PptxTextStyle, DrawingTabStop as PptxTabStop, TextDirection, PatternPreset, TextWarpPreset } from './text'
import { SUPPORTED_PATTERN_PRESETS, SUPPORTED_TEXT_WARP_PRESETS, DEFAULT_WARP_ADJUSTMENTS } from './text'

const number = (v: string | undefined, fallback = 0): number => v !== undefined && Number.isFinite(Number(v)) ? Number(v) : fallback
function align(v: string | undefined): PptxParagraph['align'] {
  return v === 'ctr' ? 'center' : v === 'r' ? 'right' : v === 'just' ? 'justify' : 'left'
}
function textCssColor(color: DrawingColor): string {
  const rgb = [color.r, color.g, color.b]
  return color.a < 1 ? `rgba(${rgb.join(',')},${color.a})` : `#${rgb.map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase()}`
}
export interface TextAppearanceIssue { kind: 'unsupported-text-appearance'; feature: string; message: string }
function style(node: XmlNode | undefined, theme?: ThemeContext, issues: TextAppearanceIssue[] = []): PptxTextStyle {
  const a = attrs(node), out: PptxTextStyle = {}
  if (a.b !== undefined) out.bold = a.b === '1' || a.b === 'true'
  if (a.i !== undefined) out.italic = a.i === '1' || a.i === 'true'
  if (a.sz !== undefined) out.fontSizePt = number(a.sz) / 100
  if (a.spc !== undefined) out.characterSpacingPt = number(a.spc) / 100
  if (a.lang !== undefined) out.language = a.lang
  for (const [tag, field] of [['latin', 'fontFamily'], ['ea', 'fontFamilyEastAsia'], ['cs', 'fontFamilyComplexScript']] as const) {
    const face = attrs(getChildren(node, tag)[0]).typeface
    if (face !== undefined) out[field] = face
  }
  if (out.fontFamily === undefined && a.typeface !== undefined) out.fontFamily = a.typeface
  if (getChildren(node, 'noFill').length) out.noFill = true
  else if (getChildren(node, 'solidFill').length) out.noFill = false
  const color = resolveDrawingColor(parseDrawingColor(getChildren(node, 'solidFill')[0]), theme)
  if (color) out.color = textCssColor(color)
  // WordArt gradient fill: linear stops resolve like solid colors. Other
  // gradient forms fall back to the flat color with a diagnostic.
  const grad = getChildren(node, 'gradFill')[0]
  if (grad) {
    const stops = getChildren(getChildren(grad, 'gsLst')[0], 'gs')
      .map(g => {
        const c = resolveDrawingColor(parseDrawingColor(g), theme)
        const pos = attrs(g).pos
        return c && pos !== undefined ? { position: number(pos) / 100000, color: textCssColor(c) } : undefined
      })
      .filter((s): s is { position: number; color: string } => !!s)
    const lin = getChildren(grad, 'lin')[0]
    const pathGrad = getChildren(grad, 'path')[0]
    if (pathGrad) {
      issues.push({ kind: 'unsupported-text-appearance', feature: 'gradFill:path', message: 'WordArt path gradient is deferred; using first stop color' })
      if (stops.length > 0 && !out.color) out.color = stops[0].color
    } else if (stops.length >= 2 && lin) {
      const ang = attrs(lin).ang
      out.textFill = { kind: 'gradient', stops, angle: ang !== undefined ? (number(ang) * Math.PI) / 10800000 : 0 }
    } else {
      issues.push({ kind: 'unsupported-text-appearance', feature: lin ? 'gradFill' : 'gradFill:no-linear', message: 'WordArt gradient needs 2+ stops and a linear vector; using flat color' })
    }
  }
  // WordArt picture/pattern fill: supported presets tile in paint; anything
  // else (including blip picture fills) falls back to the flat color.
  const patt = getChildren(node, 'pattFill')[0]
  if (patt) {
    const preset = attrs(patt).prst ?? ''
    const fg = resolveDrawingColor(parseDrawingColor(getChildren(patt, 'fgClr')[0]), theme)
    const bg = resolveDrawingColor(parseDrawingColor(getChildren(patt, 'bgClr')[0]), theme)
    if (SUPPORTED_PATTERN_PRESETS.has(preset) && fg && bg) {
      out.textFill = { kind: 'pattern', preset: preset as PatternPreset, fg: textCssColor(fg), bg: textCssColor(bg) }
    } else {
      issues.push({ kind: 'unsupported-text-appearance', feature: `pattFill:${preset || 'missing'}`,
        message: 'WordArt pattern preset is deferred; using flat color' })
    }
  } else if (getChildren(node, 'blipFill').length) {
    issues.push({ kind: 'unsupported-text-appearance', feature: 'blipFill', message: 'WordArt picture fill is deferred; using flat color' })
  }
  // An explicit solid fill on this element clears any inherited gradient or
  // pattern (spread merge would otherwise keep the parent's textFill).
  if (getChildren(node, 'solidFill').length && !grad && !patt && !getChildren(node, 'blipFill').length) {
    out.textFill = undefined
  }
  // A direct fill of any kind clears an inherited noFill (most specific wins);
  // a same-element noFill still suppresses.
  if (!getChildren(node, 'noFill').length &&
    (grad || patt || getChildren(node, 'solidFill').length || getChildren(node, 'blipFill').length)) {
    out.noFill = false
  }
  // WordArt outline: solid-color stroke centered on the glyph edge. An
  // explicit ln replaces any inherited outline: noFill clears it silently,
  // malformed widths diagnose and clear, anything else unresolvable
  // diagnoses and clears.
  const ln = getChildren(node, 'ln')[0]
  if (ln) {
    const wRaw = attrs(ln).w
    const wNum = wRaw !== undefined ? number(wRaw, NaN) : undefined
    const lineColor = resolveDrawingColor(parseDrawingColor(getChildren(ln, 'solidFill')[0]), theme)
    if (lineColor && wNum !== undefined && Number.isFinite(wNum) && wNum > 0) {
      out.textOutline = { color: textCssColor(lineColor), widthPx: Math.min((wNum * 96) / (12700 * 72), 100) }
    } else if (!getChildren(ln, 'noFill').length) {
      issues.push({ kind: 'unsupported-text-appearance', feature: 'ln', message: 'WordArt outline needs a finite positive width and solid color; skipped' })
    }
    if (!out.textOutline) out.textOutline = undefined
  }
  // WordArt outer shadow: dist/dir offset, optional blur radius. Alignment,
  // rotate-with-shape, scale (@sx/@sy) and skew (@kx/@ky) are deferred
  // (documented limitations).
  const effectLst = getChildren(node, 'effectLst')[0]
  const shadow = getChildren(effectLst, 'outerShdw')[0]
  if (shadow) {
    const sa = attrs(shadow)
    const shadowColor = resolveDrawingColor(parseDrawingColor(shadow), theme)
    if (shadowColor) {
      const distPx = number(sa.dist) / 9525
      const dir = sa.dir !== undefined ? (number(sa.dir) * Math.PI) / 10800000 : 0
      out.textShadow = {
        color: textCssColor(shadowColor),
        // Negative radii are invalid: clamp to a hard shadow. Hostile huge radii
        // are capped to 100px to prevent browser rasterization hangs.
        blurPx: (sa.blurRad ?? sa.bluRad) !== undefined ? Math.min(Math.max(0, number(sa.blurRad ?? sa.bluRad) / 9525), 100) : 0,
        offsetX: Math.cos(dir) * distPx,
        offsetY: Math.sin(dir) * distPx,
      }
    } else {
      issues.push({ kind: 'unsupported-text-appearance', feature: 'outerShdw', message: 'WordArt shadow needs a resolvable color; skipped' })
    }
  }

  // Extended text appearance effects: glow and reflection
  const glow = getChildren(effectLst, 'glow')[0] ?? getChildren(node, 'glow')[0]
  if (glow) {
    issues.push({ kind: 'unsupported-text-appearance', feature: 'glow', message: 'WordArt text glow effect is unsupported; falling back to plain run' })
  }
  const reflection = getChildren(effectLst, 'reflection')[0] ?? getChildren(effectLst, 'refl')[0] ?? getChildren(node, 'reflection')[0] ?? getChildren(node, 'refl')[0]
  if (reflection) {
    issues.push({ kind: 'unsupported-text-appearance', feature: 'reflection', message: 'WordArt text reflection effect is unsupported; falling back to plain run' })
  }

  return out
}
function spacing(node: XmlNode | undefined): PptxTextSpacing | undefined {
  const points = attrs(getChildren(node, 'spcPts')[0]).val
  const percent = attrs(getChildren(node, 'spcPct')[0]).val
  return points !== undefined ? { kind: 'points', value: number(points) / 100 }
    : percent !== undefined ? { kind: 'percent', value: number(percent) / 100000 } : undefined
}
function paragraphProperties(node: XmlNode | undefined): Partial<PptxParagraph> {
  if (!node) return {}
  const a = attrs(node), out: Partial<PptxParagraph> = {}
  if (a.algn !== undefined) { out.align = align(a.algn); out.sourceAlign = a.algn }
  for (const [name, key] of [['marL', 'marginLeftEmu'], ['marR', 'marginRightEmu'], ['indent', 'indentEmu'], ['defTabSz', 'defaultTabSizeEmu']] as const) {
    if (a[name] !== undefined) out[key] = number(a[name])
  }
  for (const [name, key] of [['lnSpc', 'lineSpacing'], ['spcBef', 'spaceBefore'], ['spcAft', 'spaceAfter']] as const) {
    const s = spacing(getChildren(node, name)[0]); if (s) out[key] = s
  }
  const tabList = getChildren(node, 'tabLst')[0]
  if (tabList) out.tabStops = getChildren(tabList, 'tab').map((tab): PptxTabStop => {
    const a = attrs(tab)
    return { positionEmu: number(a.pos), align: a.algn === 'ctr' ? 'center' : a.algn === 'r' ? 'right' : a.algn === 'dec' ? 'decimal' : 'left' }
  }).sort((a, b) => a.positionEmu - b.positionEmu)
  const bullet = getChildren(node, 'buChar')[0]
  if (bullet) { out.bullet = true; out.bulletCharacter = attrs(bullet).char ?? '•' }
  if (getChildren(node, 'buNone').length) { out.bullet = false; out.bulletCharacter = undefined }
  return out
}

export interface InheritedTextLayer {
  /** Master text style or another level-based DrawingML style. */
  style?: XmlNode
  /** Placeholder txBody contributes bodyPr, list defaults and paragraph metadata, never its content. */
  body?: XmlNode
  origin: 'master' | 'layout' | 'placeholder'
}

const directions = new Set<TextDirection>(['horz', 'vert', 'vert270', 'wordArtVert', 'eaVert', 'mongolianVert', 'wordArtVertRtl'])

export interface TextWarpIssue { kind: 'unsupported-text-warp'; feature: string; message: string }
export interface ParsedDrawingTextBody extends PptxTextBody {
  diagnostics?: Array<{ kind: 'unsupported-text-alignment'; feature: string; message: string } | TextAppearanceIssue | TextWarpIssue>
}

export function parseTextBody(txBody: XmlNode, theme?: ThemeContext, defaults?: XmlNode, fontDefaults: PptxTextStyle = {}, inheritedLayers: readonly InheritedTextLayer[] = []): ParsedDrawingTextBody {
  const a = Object.assign({}, ...inheritedLayers.map(layer => attrs(getChildren(layer.body, 'bodyPr')[0])), attrs(getChildren(txBody, 'bodyPr')[0]))
  const body: ParsedDrawingTextBody = { paragraphs: [], anchor: a.anchor === 'ctr' ? 'ctr' : a.anchor === 'b' ? 'b' : 't',
    insetLeftEmu: number(a.lIns, 91440), insetRightEmu: number(a.rIns, 91440), insetTopEmu: number(a.tIns, 45720), insetBottomEmu: number(a.bIns, 45720), wrap: a.wrap !== 'none',
    ...(directions.has(a.vert as TextDirection) ? { direction: a.vert as TextDirection } : {}) }
  const bodyPrNodes = [...inheritedLayers.map(layer => getChildren(layer.body, 'bodyPr')[0]), getChildren(txBody, 'bodyPr')[0]].filter((n): n is XmlNode => !!n)
  let prstWarpNode: XmlNode | undefined
  for (let i = bodyPrNodes.length - 1; i >= 0; i--) {
    const warp = getChildren(bodyPrNodes[i], 'prstTxWarp')[0]
    if (warp) { prstWarpNode = warp; break }
  }
  if (prstWarpNode) {
    const warpPrst = attrs(prstWarpNode).prst
    if (warpPrst) {
      if (warpPrst === 'textNoShape' || warpPrst === 'textPlain') {
        // Handled as unwarped standard text
      } else if (SUPPORTED_TEXT_WARP_PRESETS.has(warpPrst)) {
        const avLst = getChildren(prstWarpNode, 'avLst')[0]
        const adjustments: Record<string, number> = { ...(DEFAULT_WARP_ADJUSTMENTS[warpPrst] ?? {}) }
        if (avLst) {
          for (const gd of getChildren(avLst, 'gd')) {
            const ga = attrs(gd)
            if (ga.name && ga.fmla) {
              const rawVal = ga.fmla.startsWith('val ') ? ga.fmla.slice(4).trim() : ga.fmla.trim()
              const valNum = Number(rawVal)
              if (Number.isFinite(valNum)) {
                adjustments[ga.name] = valNum
              }
            }
          }
        }
        body.textWarp = {
          preset: warpPrst as TextWarpPreset,
          adjustments: Object.keys(adjustments).length > 0 ? adjustments : undefined,
        }
      } else {
        body.diagnostics ??= []
        body.diagnostics.push({
          kind: 'unsupported-text-warp',
          feature: warpPrst,
          message: `WordArt warp preset ${warpPrst} is unsupported; using unwarped text`,
        })
      }
    }
  }
  const list = getChildren(txBody, 'lstStyle')[0]
  for (const p of getChildren(txBody, 'p')) {
    const pPr = getChildren(p, 'pPr')[0], level = number(attrs(pPr).lvl)
    const listNodes = (node: XmlNode | undefined) => [getChildren(node, 'defPPr')[0], getChildren(node, `lvl${level + 1}pPr`)[0]]
    const propertyLayers: Array<[XmlNode | undefined, 'default' | 'master' | 'layout' | 'placeholder' | 'list' | 'paragraph']> =
      listNodes(defaults).map(node => [node, 'default'])
    for (const layer of inheritedLayers) {
      for (const node of listNodes(layer.style)) propertyLayers.push([node, layer.origin])
      for (const node of listNodes(getChildren(layer.body, 'lstStyle')[0])) propertyLayers.push([node, layer.origin])
      // A placeholder paragraph can supply pPr even when its text is intentionally empty.
      const placeholderParagraph = getChildren(layer.body, 'p').find(candidate => number(attrs(getChildren(candidate, 'pPr')[0]).lvl) === level)
      if (placeholderParagraph) propertyLayers.push([getChildren(placeholderParagraph, 'pPr')[0], layer.origin])
    }
    propertyLayers.push(...listNodes(list).map(node => [node, 'list'] as [XmlNode | undefined, 'list']), [pPr, 'paragraph'])
    let inherited = { ...fontDefaults }
    const appearanceIssues: TextAppearanceIssue[] = []
    const propertySources: NonNullable<PptxTextRun['propertySources']> = {}
    for (const key of Object.keys(fontDefaults) as Array<keyof PptxTextStyle>) propertySources[key] = 'default'
    const para: PptxParagraph = { runs: [], align: 'left', level }
    for (const [node, origin] of propertyLayers) {
      Object.assign(para, paragraphProperties(node))
      const values = style(getChildren(node, 'defRPr')[0], theme, appearanceIssues)
      inherited = { ...inherited, ...values }
      for (const key of Object.keys(values) as Array<keyof PptxTextStyle>) propertySources[key] = origin
    }
    if (para.sourceAlign && !['l', 'ctr', 'r', 'just'].includes(para.sourceAlign)) {
      body.diagnostics ??= []
      body.diagnostics.push({ kind: 'unsupported-text-alignment', feature: para.sourceAlign, message: `DrawingML alignment ${para.sourceAlign} is deferred; using left alignment` })
    }
    para.defaultProperties = inherited
    para.endProperties = style(getChildren(p, 'endParaRPr')[0], theme, appearanceIssues)
    for (const [name, node] of orderedChildren(p)) {
      if (name !== 'r' && name !== 'br' && name !== 'fld') continue
      const directProperties = style(getChildren(node, 'rPr')[0], theme, appearanceIssues)
      const end = name === 'br' ? para.endProperties : {}
      const origins = { ...propertySources }
      for (const key of Object.keys(end) as Array<keyof PptxTextStyle>) origins[key] = 'end'
      for (const key of Object.keys(directProperties) as Array<keyof PptxTextStyle>) origins[key] = 'run'
      para.runs.push({ ...inherited, ...end, ...directProperties, text: name === 'br' ? '\n' : textOf(getChildren(node, 't')[0]), directProperties, propertySources: origins })
    }
    if (appearanceIssues.length) body.diagnostics = [...(body.diagnostics ?? []), ...appearanceIssues]
    body.paragraphs.push(para)
  }
  return body
}

/** Shape fontRef supplies theme face/color defaults; it never overrides explicit rPr. */
export function textFontDefaults(shapeStyle: XmlNode | undefined, theme?: ThemeContext): PptxTextStyle {
  const ref = getChildren(shapeStyle, 'fontRef')[0], index = attrs(ref).idx
  const out: PptxTextStyle = {}
  if (index === 'major' || index === 'minor') {
    const prefix = index === 'major' ? '+mj' : '+mn'
    Object.assign(out, { fontFamily: `${prefix}-lt`, fontFamilyEastAsia: `${prefix}-ea`, fontFamilyComplexScript: `${prefix}-cs` })
  }
  const color = resolveDrawingColor(parseDrawingColor(ref), theme)
  if (color) out.color = textCssColor(color)
  return out
}
