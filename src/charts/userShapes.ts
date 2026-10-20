/**
 * Chart user shapes (spec 5.11; lazy chunk): the drawing overlays a chart part names through `c:userShapes`.
 *
 * - **Resolve.** `c:userShapes@r:id` goes through the chart part's `.rels` with the bounded, admitted relationships
 *   reader (charts/embeddedZip.ts `resolveRelationshipIds`), never the package's cached XML. The part is read with
 *   the bounded chart-part reader (charts/read.ts), its tree charged and released once the overlay is prepared.
 * - **Prepare** (stage 1). Each `cdr:relSizeAnchor` (from/to as fractions of the chart) and `cdr:absSizeAnchor`
 *   (from as fractions, an EMU extent) keeps its objects as scene nodes prepared with the existing drawing
 *   preparation (geometry, style, group properties, text body), under the document's drawing node budget and the
 *   group depth bound. Pictures resolve through the user-shapes part's own `.rels`; their bytes are read uncached
 *   and bounded, their header probed against the existing image budget (`imageLimitsFor`) before any decode.
 * - **Cycles.** The user-shapes part, a picture or a nested chart reference that leads back to the chart part or to
 *   the user-shapes part is never followed (`chart-user-shapes-cycle`).
 * - **Ownership.** Everything retained is charged to the document's chart workspace before it is built, then
 *   transferred to the model's tag (refunded with the model). Decoded pictures live in a realm-local sidecar, each
 *   charged under its own tag until the model is released or collected. A refused reservation refuses the chart
 *   (spec 6.4), with everything refunded.
 * - **Failure.** A missing, external or unreadable user-shapes part leaves the chart native, with a diagnostic.
 * - **Paint** (stage 2). The anchors are placed over the whole chart area and painted through the scene painter.
 */
import type JSZip from 'jszip'
import { attrs, getChildren, orderedChildren, type XmlNode } from '../core/xml'
import { decodeImage, releaseDecodedImage, sniffImageMime } from '../core/images'
import { parseGeometry } from '../drawing/geometry'
import { DRAWING_GROUP_DEPTH, drawingPartContext, reserveDrawingNode } from '../drawing/parts'
import type { SceneImage, SceneNode } from '../drawing/scene'
import { paintScene } from '../drawing/scene-paint'
import { parseGroupShapeProperties, resolveDrawingStyle, type DrawingStyle } from '../drawing/style'
import { parseTextBody, textFontDefaults } from '../drawing/text-parse'
import { paintTextBody } from '../drawing/text-paint'
import type { PptxTextBody } from '../pptx/types'
import { imageLimitsFor, probeImageHeader } from '../xlsx/storage/imageAdmission'
import type { ChartParseContext } from './context'
import { resolveRelationshipIds, type RelationshipTarget } from './embeddedZip'
import type { ChartDiagnostic, ChartModel } from './model'
import { adoptChartModel, chartModelOwner, modelOwner, onChartModelRelease, succeedChartModel } from './own'
import { stringBytes } from './points'
import { CHART_CHUNK_LIMIT, CHART_PART_LIMIT, readChartPart, streamEntry } from './read'
import type { ChartWorkspace } from './workspace'

export interface ChartUserShapeImage extends SceneImage { data: Uint8Array; width: number; height: number }
export type ChartUserShapeNode = SceneNode<PptxTextBody, ChartUserShapeImage>
/** One anchor: fractions of the chart (`rel`: from/to; `abs`: from plus an EMU extent) and its objects. */
export type ChartUserShapeAnchor =
  | { kind: 'rel'; x: number; y: number; x2: number; y2: number; nodes: ChartUserShapeNode[] }
  | { kind: 'abs'; x: number; y: number; cx: number; cy: number; nodes: ChartUserShapeNode[] }
export interface ChartUserShapes { anchors: ChartUserShapeAnchor[]; images: ChartUserShapeImage[] }

/** Retained bytes priced per prepared node, beyond its source subtree (charged before the node exists). */
export const USER_SHAPE_NODE_BYTES = 512
const ANCHOR_BYTES = 128, IMAGE_BYTES = 256, ID_BYTES = 32, PLACED_BYTES = 640
const EMU = 9525
const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]
const finite = (raw: string | undefined): number | undefined => raw !== undefined && raw.trim() !== '' && Number.isFinite(Number(raw)) ? Number(raw) : undefined
const text = (node: XmlNode | undefined): string | undefined => { const t = node?.['#text']; return typeof t === 'string' ? t : undefined }
const bool = (raw: string | undefined) => raw === '1' || raw === 'true'
const OBJECTS = new Set(['sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame', 'AlternateContent'])
let reads = 0, paints = 0, decodes = 0

class Refused extends Error {}

/**
 * Stage 1: reads and prepares the chart's user shapes onto a successor model (never throws). A chart without
 * `c:userShapes` is returned as is; one whose overlay cannot be admitted becomes `refused`.
 */
export async function attachChartUserShapes(model: ChartModel, root: XmlNode, context: ChartParseContext): Promise<ChartModel> {
  const reference = child(root, 'userShapes')
  if (!reference || model.unsupported) return model
  const { workspace, pkg } = context
  let owner = chartModelOwner(model)
  if (!owner) { owner = modelOwner(workspace); adoptChartModel(model, owner) }
  const tag = `chart-user-shapes:${context.partPath}#${++reads}`, work = `${tag}:work`
  let staged = 0
  const charge = (bytes: number) => { if (!workspace.reserve(bytes, tag)) throw new Refused(); staged += bytes }
  const transient = (bytes: number) => { if (!workspace.reserve(bytes, work)) throw new Refused() }
  const diagnostics: ChartDiagnostic[] = [], seen = new Set<string>()
  const note = (code: string, reason: string) => { const key = `${code}\n${reason}`; if (!seen.has(key)) { seen.add(key); diagnostics.push({ code, reason }) } }
  let release: (() => void) | undefined
  try {
    const rId = attrs(reference).id
    const rels = rId ? await resolveRelationshipIds(context, [rId]) : new Map<string, RelationshipTarget>()
    if (rels === 'refused') throw new Refused()
    const target = rId ? rels.get(rId) : undefined
    if (target?.kind !== 'internal') {
      note('chart-user-shapes-missing', target?.kind === 'external' ? "The chart's shapes are linked outside the document" : target?.kind === 'outside' ? "The chart's shapes point outside the document package" : "The chart's shapes are missing")
      return finish(model, undefined, diagnostics)
    }
    const path = target.path
    if (path === context.partPath) { note('chart-user-shapes-cycle', "The chart's shapes refer back to the chart"); return finish(model, undefined, diagnostics) }
    const read = await readChartPart(pkg, path, context)
    if (read?.kind === 'refused') throw new Refused()
    if (!read) { note('chart-user-shapes-missing', "The chart's shapes are missing"); return finish(model, undefined, diagnostics) }
    if (read.kind !== 'tree') { note('chart-user-shapes-unreadable', "The chart's shapes could not be read"); return finish(model, undefined, diagnostics) }
    release = read.release
    const shapes = await prepare(read.root, { context, path, ancestry: [context.partPath, path], charge, transient, note })
    release(); release = undefined
    workspace.release(work)
    if (shapes) workspace.transfer(tag, owner.tag, staged)
    else workspace.release(tag)
    const next = finish(model, shapes, diagnostics)
    if (shapes?.images.length) await chartUserShapesReady(next)
    return next
  } catch (error) {
    release?.()
    workspace.release(work); workspace.release(tag)
    if (!(error instanceof Refused)) {
      note('chart-user-shapes-unreadable', "The chart's shapes could not be read")
      return finish(model, undefined, diagnostics)
    }
    return finish(model, undefined, [...diagnostics, { code: 'chart-refused', reason: 'This chart needs more memory than is available for charts' }], true)
  }
}

/** The successor model: takes over the old one's charge; its diagnostics are admitted (dropped if refused). */
function finish(model: ChartModel, userShapes: ChartUserShapes | undefined, diagnostics: ChartDiagnostic[], refused = false): ChartModel {
  if (!userShapes && !diagnostics.length && !refused) return model
  const owner = chartModelOwner(model)
  const admitted = diagnostics.filter(d => !owner || owner.workspace.reserve(stringBytes(d.code) + stringBytes(d.reason), owner.tag))
  const next: ChartModel = {
    ...model,
    ...(userShapes ? { userShapes } : {}),
    ...(admitted.length ? { diagnostics: [...model.diagnostics ?? [], ...admitted] } : {}),
    ...(refused ? { unsupported: 'This chart needs more memory than is available for charts', refused: true as const, preparedAxes: undefined } : {}),
  }
  succeedChartModel(model, next)
  return next
}

interface Preparation {
  context: ChartParseContext
  /** The user-shapes part. */
  path: string
  ancestry: readonly string[]
  /** Retained bytes, before the allocation (refusal throws). */
  charge(bytes: number): void
  /** Work released when preparation ends. */
  transient(bytes: number): void
  note(code: string, reason: string): void
}

/** The source subtree's retained size: elements, attributes and text (the prepared node is proportional to it). */
function sourceBytes(node: XmlNode, skip?: ReadonlySet<string>): number {
  let bytes = 0
  const pending: XmlNode[] = [node]
  while (pending.length) {
    const next = pending.pop()!
    bytes += 256
    for (const [name, value] of Object.entries(attrs(next))) bytes += 2 * (name.length + String(value).length)
    const t = text(next)
    if (t) bytes += 4 * t.length
    for (const [name, item] of orderedChildren(next)) if (name !== '#text' && !skip?.has(name)) pending.push(item)
  }
  return bytes
}

async function prepare(root: XmlNode, prep: Preparation): Promise<ChartUserShapes | undefined> {
  const { context, path } = prep
  // Every relationship the overlay names (pictures, nested charts), resolved once through the bounded reader.
  const ids = new Set<string>()
  const collect = (node: XmlNode) => {
    const pending = [node]
    while (pending.length) {
      const next = pending.pop()!
      for (const [name, item] of orderedChildren(next)) {
        if (name === '#text') continue
        const a = attrs(item)
        const id = name === 'blip' ? a.embed : name === 'chart' ? a.id : undefined
        if (id && !ids.has(id)) { prep.transient(ID_BYTES + 2 * id.length); ids.add(id) }
        pending.push(item)
      }
    }
  }
  collect(root)
  const rels = ids.size ? await resolveRelationshipIds({ pkg: context.pkg, partPath: path, workspace: context.workspace }, [...ids]) : new Map<string, RelationshipTarget>()
  if (rels === 'refused') throw new Refused()
  const images: ChartUserShapeImage[] = [], byPath = new Map<string, number>()
  const state = { exhausted: false }
  const limits = imageLimitsFor(context.profile === 'mobile' ? 'M' : 'D'), cap = CHART_PART_LIMIT[context.profile]

  const image = async (id: string | undefined): Promise<number | undefined> => {
    const target = id ? rels.get(id) : undefined
    if (target?.kind !== 'internal') { prep.note('chart-user-shapes-image-missing', target?.kind === 'external' ? 'A picture in the chart is linked outside the document' : 'A picture in the chart is missing'); return undefined }
    if (prep.ancestry.includes(target.path)) { prep.note('chart-user-shapes-cycle', 'A picture in the chart refers back to the chart'); return undefined }
    const known = byPath.get(target.path)
    if (known !== undefined) return known
    const data = await readImage(context, target.path, cap, prep)
    if (data === 'limit') { prep.note('chart-user-shapes-image-limit', 'A picture in the chart is larger than the viewer reads'); return undefined }
    if (!data) { prep.note('chart-user-shapes-image-missing', 'A picture in the chart is missing'); return undefined }
    const header = probeImageHeader(data)
    if ('error' in header) { prep.note('chart-user-shapes-image-unsupported', 'A picture in the chart is in a format the viewer does not draw'); return undefined }
    if (header.width * header.height > limits.maxPixels || header.width * header.height * 4 > limits.maxDecodedBytes) {
      prep.note('chart-user-shapes-image-limit', `A ${header.width}x${header.height} picture in the chart exceeds the image budget`)
      return undefined
    }
    prep.charge(IMAGE_BYTES)
    images.push({ data, mime: sniffImageMime(data), partPath: target.path, width: header.width, height: header.height })
    byPath.set(target.path, images.length - 1)
    return images.length - 1
  }

  const object = async (name: string, node: XmlNode, depth: number, top: boolean, groupStyle?: Partial<DrawingStyle>): Promise<ChartUserShapeNode | undefined> => {
    if (name === 'grpSp' && depth >= DRAWING_GROUP_DEPTH) { prep.note('chart-user-shapes-group-depth', 'A shape group in the chart is nested too deeply'); return undefined }
    if (name === 'graphicFrame') {
      const chart = descendant(node, 'chart'), target = chart ? rels.get(attrs(chart).id) : undefined
      if (target?.kind === 'internal' && prep.ancestry.includes(target.path)) prep.note('chart-user-shapes-cycle', 'A shape in the chart refers back to the chart')
      else prep.note('chart-user-shapes-unsupported', 'A graphic frame in the chart shapes is not drawn')
      return undefined
    }
    if (!reserveDrawingNode(drawingPartContext(context.pkg), path)) { state.exhausted = true; prep.note('chart-user-shapes-node-limit', 'The document has more drawing objects than the viewer draws'); return undefined }
    const spPr = child(node, name === 'grpSp' ? 'grpSpPr' : 'spPr')
    // A group is priced without its child objects (each is charged as it is prepared).
    prep.charge(USER_SHAPE_NODE_BYTES + sourceBytes(node, name === 'grpSp' ? OBJECTS : undefined))
    const xfrm = child(spPr, 'xfrm')
    const result: ChartUserShapeNode = transform(xfrm)
    // An anchor owns its object's box: a missing local transform is valid at the top level.
    if (top && !xfrm) result.transformValid = true
    const theme = context.theme
    if (name === 'grpSp') {
      const style = parseGroupShapeProperties(spPr, theme, groupStyle)
      result.drawingStyle = style
      const ca = attrs(child(xfrm, 'chOff')), ce = attrs(child(xfrm, 'chExt'))
      const cx = finite(ca.x), cy = finite(ca.y), cw = finite(ce.cx), ch = finite(ce.cy)
      if (cx === undefined || cy === undefined || cw === undefined || ch === undefined || cw <= 0 || ch <= 0) { result.transformValid = false; return result }
      result.group = { off: { x: result.xEmu, y: result.yEmu }, ext: { width: result.widthEmu, height: result.heightEmu }, chOff: { x: cx, y: cy }, chExt: { width: cw, height: ch } }
      result.children = await objects(node, depth + 1, false, style)
    } else if (name === 'sp' || name === 'cxnSp') {
      result.drawingGeometry = parseGeometry(spPr)
      result.drawingStyle = resolveDrawingStyle(spPr, child(node, 'style'), theme, name === 'cxnSp' ? { ...groupStyle, fill: { kind: 'none' } } : groupStyle)
      const body = child(node, 'txBody')
      if (body) result.textBody = parseTextBody(body, theme, undefined, textFontDefaults(child(node, 'style'), theme)) as unknown as PptxTextBody
    } else {
      const index = await image(attrs(child(child(node, 'blipFill'), 'blip')).embed)
      if (index === undefined) return undefined
      result.imageIndex = index
      result.image = images[index]
    }
    return result
  }
  const objects = async (container: XmlNode, depth: number, top: boolean, groupStyle?: Partial<DrawingStyle>): Promise<ChartUserShapeNode[]> => {
    const out: ChartUserShapeNode[] = []
    for (const [name, node] of orderedChildren(container)) {
      if (state.exhausted) break
      if (!OBJECTS.has(name)) continue
      // A compatibility wrapper draws its Fallback (the Choice requires features the viewer cannot vouch for).
      if (name === 'AlternateContent') { const fallback = child(node, 'Fallback'); if (fallback && depth < DRAWING_GROUP_DEPTH) out.push(...await objects(fallback, depth + 1, top, groupStyle)); continue }
      const prepared = await object(name, node, depth, top, groupStyle)
      if (prepared) out.push(prepared)
    }
    return out
  }

  const anchors: ChartUserShapeAnchor[] = []
  for (const [name, node] of orderedChildren(root)) {
    if (state.exhausted) break
    if (name !== 'relSizeAnchor' && name !== 'absSizeAnchor') continue
    const from = child(node, 'from'), x = finite(text(child(from, 'x'))), y = finite(text(child(from, 'y')))
    const to = child(node, 'to'), x2 = finite(text(child(to, 'x'))), y2 = finite(text(child(to, 'y')))
    const ext = attrs(child(node, 'ext')), cx = finite(ext.cx), cy = finite(ext.cy)
    const valid = x !== undefined && y !== undefined && (name === 'relSizeAnchor' ? x2 !== undefined && y2 !== undefined && x2 >= x && y2 >= y : cx !== undefined && cy !== undefined && cx >= 0 && cy >= 0)
    if (!valid) { prep.note('chart-user-shapes-anchor', 'A shape in the chart has an invalid anchor'); continue }
    prep.charge(ANCHOR_BYTES)
    const nodes = await objects(node, 0, true)
    if (!nodes.length) continue
    anchors.push(name === 'relSizeAnchor' ? { kind: 'rel', x: x!, y: y!, x2: x2!, y2: y2!, nodes } : { kind: 'abs', x: x!, y: y!, cx: cx!, cy: cy!, nodes })
  }
  return anchors.length ? { anchors, images } : undefined
}

function descendant(node: XmlNode, name: string): XmlNode | undefined {
  const pending = [node]
  while (pending.length) for (const [n, item] of orderedChildren(pending.pop()!)) { if (n === name) return item; if (n !== '#text') pending.push(item) }
  return undefined
}

function transform(node: XmlNode | undefined): ChartUserShapeNode {
  const oa = attrs(child(node, 'off')), ea = attrs(child(node, 'ext')), xa = attrs(node)
  const x = finite(oa.x), y = finite(oa.y), w = finite(ea.cx), h = finite(ea.cy), rot = xa.rot === undefined ? 0 : finite(xa.rot)
  const valid = x !== undefined && y !== undefined && w !== undefined && h !== undefined && rot !== undefined && w >= 0 && h >= 0
  return { xEmu: x ?? 0, yEmu: y ?? 0, widthEmu: w ?? 0, heightEmu: h ?? 0, rotationDeg: (rot ?? 0) / 60000, flipH: bool(xa.flipH), flipV: bool(xa.flipV), transformValid: valid }
}

/** A picture part's bytes, uncached and bounded by `cap` (actual inflated bytes), charged before they are kept. */
async function readImage(context: ChartParseContext, path: string, cap: number, prep: Preparation): Promise<Uint8Array | 'limit' | undefined> {
  const { pkg } = context
  const entry = (pkg as { entry?: (path: string) => JSZip.JSZipObject | null }).entry?.(path) ?? null
  if (!entry) {
    // A paged package reads the part itself, admitted against its own pool (bounded by the same cap).
    let bytes: Uint8Array | undefined
    try { bytes = await pkg.bytes(path, cap) } catch (error) { if (String((error as Error | undefined)?.message).startsWith('Part byte limit')) return 'limit'; return undefined }
    if (bytes) prep.charge(bytes.length)
    return bytes
  }
  if (entry.dir) return undefined
  const { workspace } = context, scratch = `chart-user-shapes-read:${path}#${++reads}`
  const chunks: Uint8Array[] = []
  let length = 0, over = false
  class Over extends Error {}
  try {
    // The in-flight inflate chunk, then each retained chunk, all refunded once the contiguous copy (retained) exists.
    if (!workspace.reserve(CHART_CHUNK_LIMIT, scratch)) throw new Refused()
    try {
      await streamEntry(entry, chunk => {
        if (chunk.length > CHART_CHUNK_LIMIT) throw new Refused()
        if (length + chunk.length > cap) { over = true; throw new Over() }
        if (!workspace.reserve(chunk.length, scratch)) throw new Refused()
        chunks.push(chunk)
        length += chunk.length
      })
    } catch (error) {
      if (error instanceof Refused) throw error
      return over ? 'limit' : undefined
    }
    prep.charge(length)
    const bytes = new Uint8Array(length)
    let at = 0
    for (const chunk of chunks) { bytes.set(chunk, at); at += chunk.length }
    return bytes
  } finally { chunks.length = 0; workspace.release(scratch) }
}

/** Decoded pictures (realm-local; never serialized with the model), each charged under its own tag. */
interface Decoded { image?: CanvasImageSource; pending?: Promise<void>; workspace?: ChartWorkspace; tag?: string; released?: boolean }
const decoded = new WeakMap<ChartUserShapeImage, Decoded>()
const collected = typeof FinalizationRegistry === 'function'
  ? new FinalizationRegistry<Decoded>(state => { if (state.tag) state.workspace?.release(state.tag); releaseDecodedImage(state.image) })
  : undefined
function releaseDecoded(image: ChartUserShapeImage): void {
  const state = decoded.get(image)
  if (!state) return
  decoded.delete(image)
  collected?.unregister(state)
  state.released = true
  if (state.tag) state.workspace?.release(state.tag)
  releaseDecodedImage(state.image)
  state.image = undefined
}
onChartModelRelease((model, next) => {
  if (!model.userShapes || next?.userShapes === model.userShapes) return
  for (const image of model.userShapes.images) releaseDecoded(image)
})

/** Decodes one picture into the sidecar: the surface is charged before the decode (and its transient beside it). */
function decodeOne(image: ChartUserShapeImage, workspace: ChartWorkspace): Promise<void> {
  const known = decoded.get(image)
  if (known) return known.pending ?? Promise.resolve()
  const state: Decoded = {}
  decoded.set(image, state)
  const tag = `chart-user-shapes-image#${++decodes}`, surface = 4 * image.width * image.height
  if (!workspace.reserve(surface, tag)) return Promise.resolve()
  if (!workspace.reserve(surface, `${tag}:decode`)) { workspace.release(tag); return Promise.resolve() }
  state.workspace = workspace; state.tag = tag
  collected?.register(image, state, state)
  state.pending = decodeImage(image.data, image.mime).then(
    bitmap => { if (state.released) releaseDecodedImage(bitmap); else state.image = bitmap },
    () => { workspace.release(tag); state.tag = undefined },
  ).finally(() => { workspace.release(`${tag}:decode`); state.pending = undefined })
  return state.pending
}

/** Resolves once every picture of the model's user shapes has been decoded (or failed) in this realm. */
export function chartUserShapesReady(model: ChartModel, workspace?: ChartWorkspace): Promise<void> {
  const images = model.userShapes?.images ?? []
  const owner = workspace ?? chartModelOwner(model)?.workspace
  if (!owner) return Promise.all(images.map(image => decoded.get(image)?.pending)).then(() => {})
  return Promise.all(images.map(image => decodeOne(image, owner))).then(() => {})
}

/** One anchor's box in chart pixels. */
function placement(anchor: ChartUserShapeAnchor, w: number, h: number): { x: number; y: number; w: number; h: number } | undefined {
  const x = anchor.x * w, y = anchor.y * h
  const box = anchor.kind === 'rel' ? { x, y, w: (anchor.x2 - anchor.x) * w, h: (anchor.y2 - anchor.y) * h } : { x, y, w: anchor.cx / EMU, h: anchor.cy / EMU }
  return [box.x, box.y, box.w, box.h].every(Number.isFinite) && box.w >= 0 && box.h >= 0 ? box : undefined
}

/**
 * Stage 2: paints the model's user shapes over the chart (the whole `w` x `h` chart area) through the scene
 * painter. The placed copies are paint-local, charged to `workspace` and refunded after the paint. A picture not
 * yet decoded in this realm is skipped; a `durable` workspace (one that outlives the paint) starts its decode.
 * A throwing overlay never affects the chart beneath or later shapes.
 */
export function paintChartUserShapes(model: ChartModel, ctx: CanvasRenderingContext2D, w: number, h: number, workspace: ChartWorkspace, durable: boolean): void {
  const shapes = model.userShapes
  if (!shapes?.anchors.length) return
  const tag = `chart-user-shapes-paint#${++paints}`
  if (!workspace.reserve(shapes.anchors.length * PLACED_BYTES + shapes.images.length * 16, tag)) return
  ctx.save()
  try {
    const images = shapes.images.map(image => {
      const state = decoded.get(image)
      if (!state && durable) void decodeOne(image, workspace)
      return state?.image
    })
    for (const anchor of shapes.anchors) {
      const box = placement(anchor, w, h)
      if (!box) continue
      const at = { x: box.x * EMU, y: box.y * EMU, width: box.w * EMU, height: box.h * EMU }
      const nodes = anchor.nodes.map(node => ({
        ...node, xEmu: at.x, yEmu: at.y, widthEmu: at.width, heightEmu: at.height,
        ...(node.group ? { group: { ...node.group, off: { x: at.x, y: at.y }, ext: { width: at.width, height: at.height } } } : {}),
      }))
      try {
        paintScene(nodes, ctx, { images, paintText(node, c, nw, nh) { if (node.textBody) paintTextBody(node.textBody, c, 0, 0, nw, nh, family => family) } })
      } catch { /* A malformed overlay cannot suppress the chart or later anchors. */ }
    }
  } finally { ctx.restore(); workspace.release(tag) }
}
