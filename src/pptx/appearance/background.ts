/** Lazy background resolution. Office-shaped sources are not Office visual evidence. */
import { HandlePackage, HANDLE_AGGREGATE_BYTES } from './handle-package'
import type { OfficePackage } from '../../core/zip'
import { PackageError } from '../../core/package-error'
import { attrs, getChild, getChildren, orderedChildren, parseXmlOrdered, type XmlNode } from '../../core/xml'
import { sniffImageMime } from '../../core/images'
import { resolvePartTarget } from '../../drawing/parts'
import { parseGeometry } from '../../drawing/geometry'
import { parseDrawingColor, parseDrawingColors, parseFillDefinition, resolveDrawingColor, resolveFill, type DrawingColor, type DrawingIssue, type DrawingFill, type ThemeContext } from '../../drawing/style'
import type { Duotone } from '../../core/images'
import type { SceneRectFractions } from '../../drawing/scene'
import type { PptxDiagnostic, PptxImageRef, PptxShape } from '../types'

export type SlideBackground =
  | { kind: 'none' }
  | { kind: 'fill'; fill: DrawingFill }
  | { kind: 'picture'; image: PptxImageRef; tile?: { sx: number; sy: number; tx: number; ty: number; algn?: string; flip?: string }; stretch?: SceneRectFractions }

const PART_BYTES = 4 * 1024 * 1024
// Pictures share S0-4's 32 MiB enrichment budget; this is not the 512 MiB slideshow media cap.
const IMAGE_BYTES = HANDLE_AGGREGATE_BYTES
async function xml(pkg: HandlePackage, path: string): Promise<XmlNode | undefined> {
  const text = await pkg.boundedText(path, PART_BYTES)
  return text === undefined ? undefined : parseXmlOrdered(text)
}
async function relationships(pkg: HandlePackage, owner: string): Promise<XmlNode[]> {
  const slash = owner.lastIndexOf('/')
  return getChildren(await xml(pkg, `${owner.slice(0, slash + 1)}_rels/${owner.slice(slash + 1)}.rels`), 'Relationship')
}
const none = (): SlideBackground => ({ kind: 'none' })
const fractionRect = (node: XmlNode | undefined): SceneRectFractions => {
  const a = attrs(node)
  const value = (key: string) => Number.isFinite(Number(a[key])) ? Number(a[key]) / 100000 : 0
  return { l: value('l'), t: value('t'), r: value('r'), b: value('b') }
}

async function related(pkg: HandlePackage, owner: string, type: string): Promise<string | undefined> {
  const rel = (await relationships(pkg, owner)).find(r => attrs(r).Type?.endsWith('/' + type) && attrs(r).TargetMode !== 'External')
  return rel && attrs(rel).Target ? resolvePartTarget(owner, attrs(rel).Target) : undefined
}

/** Find the theme's own package path, never resolve its image rIds through the slide. */
async function themePart(pkg: HandlePackage, slidePath: string): Promise<string | undefined> {
  let owner: string | undefined = slidePath
  for (const next of ['slideLayout', 'slideMaster', undefined]) {
    const theme = await related(pkg, owner, 'theme')
    if (theme) return theme
    owner = next ? await related(pkg, owner, next) : undefined
    if (!owner) break
  }
  return undefined
}

/** `placeholder` is the colour `phClr` stands for: the bgRef's colour for a theme background style. */
async function pictureBackground(pkg: HandlePackage, owner: string, blipFill: XmlNode, diagnostics: PptxDiagnostic[], theme: ThemeContext, placeholder?: DrawingColor): Promise<SlideBackground> {
  const blip = getChild(blipFill, 'blip')
  let duotone: Duotone | undefined
  for (const [name, effect] of orderedChildren(blip)) {
    if (name === '#text') continue
    if (name === 'duotone') {
      const issues: DrawingIssue[] = []
      const colours = parseDrawingColors(effect).map(color => resolveDrawingColor(color, theme, placeholder, issues))
      const [dark, light] = colours
      if (colours.length === 2 && dark && light) { duotone = { dark, light }; continue }
      diagnostics.push({ kind: 'unsupported-fill', message: 'Background picture duotone needs two resolvable colours; the picture is painted as is', feature: name })
      continue
    }
    // The existing painter cannot apply these transforms; preserve the picture with a visible diagnostic.
    diagnostics.push({ kind: 'unsupported-fill', message: `Background picture effect ${name} is not applied`, feature: name })
  }
  const id = attrs(blip)['r:embed'] ?? attrs(blip).embed
  const rel = (await relationships(pkg, owner)).find(r => attrs(r).Id === id)
  const a = attrs(rel)
  const path = a.Target && a.TargetMode !== 'External' && a.Type?.endsWith('/image') ? resolvePartTarget(owner, a.Target) : undefined
  let data: Uint8Array | undefined
  try { data = path ? await pkg.bytes(path, IMAGE_BYTES) : undefined }
  catch (error) {
    if (!isLimit(error)) throw error
    diagnostics.push({ kind: 'missing-image', message: 'Background picture refused by the part or enrichment byte budget', feature: 'resource-limit' })
    return none()
  }
  if (!data) {
    diagnostics.push({ kind: 'missing-image', message: `Background picture is missing or external: ${id ?? '(no embed)'}`, feature: 'blipFill' })
    return none()
  }
  const crop = getChild(blipFill, 'srcRect')
  const image: PptxImageRef = { data, partPath: path, mime: sniffImageMime(data), ...(crop ? { srcRect: fractionRect(crop) } : {}), ...(duotone ? { duotone } : {}) }
  const tile = getChild(blipFill, 'tile')
  if (tile) {
    const a = attrs(tile)
    const n = (key: string, fallback: number) => a[key] !== undefined && Number.isFinite(Number(a[key])) ? Number(a[key]) : fallback
    diagnostics.push({ kind: 'unsupported-fill', message: 'Picture background tiling is retained as metadata; the existing painter stretches the picture', feature: 'tile' })
    return { kind: 'picture', image, tile: { sx: n('sx', 100000) / 100000, sy: n('sy', 100000) / 100000, tx: n('tx', 0), ty: n('ty', 0), ...(a.algn ? { algn: a.algn } : {}), ...(a.flip ? { flip: a.flip } : {}) } }
  }
  return { kind: 'picture', image, stretch: fractionRect(getChild(getChild(blipFill, 'stretch'), 'fillRect')) }
}

function fillBackground(properties: XmlNode, theme: ThemeContext, diagnostics: PptxDiagnostic[]): SlideBackground {
  const issues: DrawingIssue[] = []
  const fill = resolveFill(parseFillDefinition(properties, issues), theme, undefined, issues)
  // Background nodes are never rotated, so no compensation is required.
  diagnostics.push(...issues.filter(issue => !(issue.kind === 'unsupported-gradient' && issue.feature === 'rotWithShape')))
  return fill ? { kind: 'fill', fill } : none()
}
const isLimit = (error: unknown) => error instanceof PackageError && error.code === 'resource-limit' || error instanceof Error && error.message.startsWith('Part byte limit exceeded')

/** Caller supplies the slide's effective colour map, including clrMapOvr. Reuse one HandlePackage
 * across an enrichment run to share its aggregate. Plain packages get a bounded view for this call.
 * Resource refusals degrade to none with diagnostics; unexpected errors still reject. */
export async function resolveSlideBackground(pkg: OfficePackage, slidePath: string, theme: ThemeContext): Promise<{ background: SlideBackground; diagnostics: PptxDiagnostic[] }> {
  const reader = pkg instanceof HandlePackage ? pkg : HandlePackage.over({ pkg, slideParts: [slidePath] }, { partBytes: IMAGE_BYTES })
  const diagnostics: PptxDiagnostic[] = []
  try {
    let owner = slidePath
    for (const next of ['slideLayout', 'slideMaster', undefined]) {
      const bg = getChild(getChild(await xml(reader, owner), 'cSld'), 'bg')
      if (bg) {
        const properties = getChild(bg, 'bgPr')
        if (properties) {
          const picture = getChild(properties, 'blipFill')
          return { background: picture ? await pictureBackground(reader, owner, picture, diagnostics, theme) : fillBackground(properties, theme, diagnostics), diagnostics }
        }
        const ref = getChild(bg, 'bgRef'), index = Number(attrs(ref).idx)
        if (index === 0 || index === 1000) return { background: none(), diagnostics }
        const valid = Number.isInteger(index) && index > 0
        const definition = valid && index < 1000 ? theme.fillStyles[index - 1] : valid && index >= 1001 ? theme.bgFillStyles[index - 1001] : undefined
        // ThemeContext retains slots for blipFill but cannot represent its image. Recover raw XML.
        if (definition?.kind === 'none') {
          const path = await themePart(reader, slidePath)
          if (path) {
            const format = getChild(getChild(await xml(reader, path), 'themeElements'), 'fmtScheme')
            const list = getChild(format, index >= 1001 ? 'bgFillStyleLst' : 'fillStyleLst')
            const slot = orderedChildren(list).filter(([name]) => name !== '#text')[index >= 1001 ? index - 1001 : index - 1]
            if (slot?.[0] === 'blipFill') {
              const issues: DrawingIssue[] = []
              const placeholder = resolveDrawingColor(parseDrawingColor(ref), theme, undefined, issues)
              diagnostics.push(...issues)
              return { background: await pictureBackground(reader, path, slot[1], diagnostics, theme, placeholder), diagnostics }
            }
            if (!slot || slot[0] !== 'noFill') diagnostics.push({ kind: 'unsupported-fill', message: 'Theme background style is not representable', feature: 'bgRef' })
          } else diagnostics.push({ kind: 'missing-theme-style', message: 'Theme background source is unavailable', feature: 'bgRef' })
        }
        const issues: DrawingIssue[] = []
        if (!definition) diagnostics.push({ kind: 'missing-theme-style', message: `Background style ${attrs(ref).idx ?? '(missing)'} is unavailable`, feature: 'bgRef' })
        const fill = resolveFill(definition, theme, resolveDrawingColor(parseDrawingColor(ref), theme, undefined, issues), issues)
        diagnostics.push(...issues.filter(issue => !(issue.kind === 'unsupported-gradient' && issue.feature === 'rotWithShape')))
        return { background: fill ? { kind: 'fill', fill } : none(), diagnostics }
      }
      const path = next ? await related(reader, owner, next) : undefined
      if (!path) break
      owner = path
    }
  } catch (error) {
    if (!isLimit(error)) throw error
    diagnostics.push({ kind: 'unsupported-fill', message: 'Background metadata refused by the part or enrichment byte budget', feature: 'resource-limit' })
  }
  return { background: none(), diagnostics }
}

export function backgroundNode(bg: SlideBackground, widthEmu: number, heightEmu: number): PptxShape | undefined {
  if (bg.kind === 'none') return undefined
  // The catalogue's rect (an empty path list would paint nothing).
  const node: PptxShape = { xEmu: 0, yEmu: 0, widthEmu, heightEmu, geometry: 'rect', drawingGeometry: parseGeometry(undefined) }
  if (bg.kind === 'fill') node.drawingStyle = { fill: bg.fill, issues: [] }
  else node.image = { ...bg.image, ...(bg.stretch ? { fillRect: bg.stretch } : {}) }
  return node
}
