/** Lazy background resolution. Office-shaped sources are not Office visual evidence. */
import type JSZip from 'jszip'
import type { OfficePackage } from '../../core/zip'
import { PackageError } from '../../core/package-error'
import { attrs, getChild, getChildren, parseXmlOrdered, type XmlNode } from '../../core/xml'
import { sniffImageMime } from '../../core/images'
import { resolvePartTarget } from '../../drawing/parts'
import { parseDrawingColor, parseFillDefinition, resolveDrawingColor, resolveFill, type DrawingIssue, type DrawingFill, type ThemeContext } from '../../drawing/style'
import type { SceneRectFractions } from '../../drawing/scene'
import type { PptxDiagnostic, PptxImageRef, PptxShape } from '../types'

export type SlideBackground =
  | { kind: 'none' }
  | { kind: 'fill'; fill: DrawingFill }
  | { kind: 'picture'; image: PptxImageRef; tile?: { sx: number; sy: number; tx: number; ty: number; algn?: string; flip?: string }; stretch?: SceneRectFractions }

const PART_BYTES = 4 * 1024 * 1024
const IMAGE_BYTES = 32 * 1024 * 1024

/** Uncached entry read; stop inflation at the crossing chunk, not after materialisation.
 * Kept private until S0-4's shared reader lands on this branch. */
async function read(pkg: OfficePackage, path: string, cap: number): Promise<Uint8Array | undefined> {
  const entry = pkg.entry(path)
  if (!entry || entry.dir) return undefined
  return new Promise((resolve, reject) => {
    const stream = (entry as unknown as { internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array> }).internalStream('uint8array')
    type Worker = { previous?: Worker; push: (...args: unknown[]) => unknown; end: () => unknown; error: (error: Error) => unknown }
    const last = (stream as unknown as { _worker: Worker })._worker
    let first = last
    while (first.previous) first = first.previous
    const chunks: Uint8Array[] = []
    let total = 0, failed = false
    const fail = (error: unknown) => {
      chunks.length = 0
      if (failed) return
      failed = true
      reject(error)
      last.error(error instanceof Error ? error : new Error(String(error)))
    }
    const push = first.push.bind(first), end = first.end.bind(first)
    first.push = (...args) => { try { return push(...args) } catch (error) { fail(error) } }
    first.end = () => { try { return end() } catch (error) { fail(error) } }
    stream.on('data', chunk => {
      if (failed) return
      if (total + chunk.length > cap) throw new PackageError('resource-limit', `Background part exceeds ${cap} bytes`, path)
      total += chunk.length
      chunks.push(chunk)
    }).on('error', fail).on('end', () => {
      if (failed) return
      const out = new Uint8Array(total)
      let offset = 0
      for (const chunk of chunks) { out.set(chunk, offset); offset += chunk.length }
      chunks.length = 0
      resolve(out)
    }).resume()
  })
}
async function xml(pkg: OfficePackage, path: string): Promise<XmlNode | undefined> {
  const bytes = await read(pkg, path, PART_BYTES)
  return bytes ? parseXmlOrdered(new TextDecoder().decode(bytes)) : undefined
}
async function relationships(pkg: OfficePackage, owner: string): Promise<XmlNode[]> {
  const slash = owner.lastIndexOf('/')
  return getChildren(await xml(pkg, `${owner.slice(0, slash + 1)}_rels/${owner.slice(slash + 1)}.rels`), 'Relationship')
}
const none = (): SlideBackground => ({ kind: 'none' })
const fractionRect = (node: XmlNode | undefined): SceneRectFractions => {
  const a = attrs(node)
  const value = (key: string) => Number.isFinite(Number(a[key])) ? Number(a[key]) / 100000 : 0
  return { l: value('l'), t: value('t'), r: value('r'), b: value('b') }
}

export async function resolveSlideBackground(pkg: OfficePackage, slidePath: string, theme: ThemeContext): Promise<{ background: SlideBackground; diagnostics: PptxDiagnostic[] }> {
  const diagnostics: PptxDiagnostic[] = []
  let owner = slidePath
  for (const next of ['slideLayout', 'slideMaster', undefined]) {
    const root = await xml(pkg, owner)
    const bg = getChild(getChild(root, 'cSld'), 'bg')
    if (bg) {
      const properties = getChild(bg, 'bgPr')
      const blipFill = getChild(properties, 'blipFill')
      if (blipFill) {
        const id = attrs(getChild(blipFill, 'blip'))['r:embed'] ?? attrs(getChild(blipFill, 'blip')).embed
        const rel = (await relationships(pkg, owner)).find(r => attrs(r).Id === id)
        const a = attrs(rel)
        const path = a.Target && a.TargetMode !== 'External' && a.Type?.endsWith('/image') ? resolvePartTarget(owner, a.Target) : undefined
        const data = path ? await read(pkg, path, IMAGE_BYTES) : undefined
        if (!data) {
          diagnostics.push({ kind: 'missing-image', message: `Background picture is missing or external: ${id ?? '(no embed)'}`, feature: 'blipFill' })
          return { background: none(), diagnostics }
        }
        const image: PptxImageRef = { data, partPath: path, mime: sniffImageMime(data), srcRect: fractionRect(getChild(blipFill, 'srcRect')) }
        const tile = getChild(blipFill, 'tile')
        if (tile) {
          const a = attrs(tile)
          const n = (key: string, fallback: number) => a[key] !== undefined && Number.isFinite(Number(a[key])) ? Number(a[key]) : fallback
          diagnostics.push({ kind: 'unsupported-fill', message: 'Picture background tiling is retained as metadata; the existing painter stretches the picture', feature: 'tile' })
          return { background: { kind: 'picture', image, tile: { sx: n('sx', 100000) / 100000, sy: n('sy', 100000) / 100000, tx: n('tx', 0), ty: n('ty', 0), ...(a.algn ? { algn: a.algn } : {}), ...(a.flip ? { flip: a.flip } : {}) } }, diagnostics }
        }
        return { background: { kind: 'picture', image, stretch: fractionRect(getChild(getChild(blipFill, 'stretch'), 'fillRect')) }, diagnostics }
      }
      const issues: DrawingIssue[] = []
      let fill: DrawingFill | undefined
      if (properties) fill = resolveFill(parseFillDefinition(properties, issues), theme, undefined, issues)
      else {
        const ref = getChild(bg, 'bgRef')
        const index = Number(attrs(ref).idx)
        if (index === 0 || index === 1000) return { background: none(), diagnostics }
        // ECMA-376 20.1.4.1 style matrix: 1-based fills; 1001-based background fills.
        const definition = Number.isInteger(index) && index >= 1 && index < 1000 ? theme.fillStyles[index - 1] : Number.isInteger(index) && index >= 1001 ? theme.bgFillStyles[index - 1001] : undefined
        if (!definition) diagnostics.push({ kind: 'missing-theme-style', message: `Background style ${attrs(ref).idx ?? '(missing)'} is unavailable`, feature: 'bgRef' })
        fill = resolveFill(definition, theme, resolveDrawingColor(parseDrawingColor(ref), theme, undefined, issues), issues)
      }
      diagnostics.push(...issues)
      return { background: fill ? { kind: 'fill', fill } : none(), diagnostics }
    }
    if (!next) break
    const rel = (await relationships(pkg, owner)).find(r => attrs(r).Type?.endsWith('/' + next) && attrs(r).TargetMode !== 'External')
    if (!rel || !attrs(rel).Target) break
    owner = resolvePartTarget(owner, attrs(rel).Target)
  }
  return { background: none(), diagnostics }
}

export function backgroundNode(bg: SlideBackground, widthEmu: number, heightEmu: number): PptxShape | undefined {
  if (bg.kind === 'none') return undefined
  const node: PptxShape = { xEmu: 0, yEmu: 0, widthEmu, heightEmu, geometry: 'rect', drawingGeometry: { preset: 'rect', adjustments: [], guides: [], paths: [] } }
  if (bg.kind === 'fill') node.drawingStyle = { fill: bg.fill, issues: [] }
  else node.image = { ...bg.image, ...(bg.stretch ? { fillRect: bg.stretch } : {}) }
  return node
}
