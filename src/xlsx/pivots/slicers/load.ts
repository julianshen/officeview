/**
 * A worksheet's slicers and timelines, resolved from the package at parse time
 * (Task 12, Ruling AD/AH). Lazy chunk `pivots`.
 *
 * The workbook-level parts — slicer caches, timeline caches and the pivot caches
 * they read, plus the custom styles — are read once per package (memoised); each
 * sheet then reads its own slicer and timeline parts through its relationships
 * (the x14 `slicerList` / x15 `timelineRefs` ids, Task 4's `controlIds`).
 *
 * Failure policy: a control that cannot be resolved is left out of the result, so
 * its drawing frame keeps the AlternateContent Fallback; the reason is a
 * diagnostic (`slicer-*` / `timeline-*`). Never throws.
 */
import type { XmlNode } from '../../../core/xml'
import { PackageError, isPackageError } from '../../../core/package-error'
import { parseStylesExtras } from '../../tableStyles/parseStylesExtras'
import type { StylesExtras } from '../../tableStyles/model'
import type { TableStyleSource } from '../../tableStyles/model'
import { parseCacheDefinition } from '../parseCacheDefinition'
import { relationships, workbookCaches, type PackageReader } from '../load'
import type { PivotCacheDefinition } from '../model'
import type { ControlDiagnostic, SheetControls, SheetControlsRequest } from './model'
import { buildSlicerModel, parseSlicerCache, parseSlicerPart, type SlicerCacheDef } from './parseSlicers'
import { buildTimelineModel, parseTimelineCache, parseTimelinePart, type TimelineCacheDef } from './parseTimelines'
import { slicerAppearance, timelineAppearance } from './style'

const WORKBOOK = 'xl/workbook.xml'
/** Bounds on the workbook-level lists (a slicer cache per slicer is the norm). */
export const MAX_CONTROL_CACHES = 1024

interface WorkbookControls {
  slicerCaches: Map<string, SlicerCacheDef>
  timelineCaches: Map<string, TimelineCacheDef>
  /** By x14 `pivotCacheId`, the id slicer and timeline caches name. */
  pivotCaches: Map<number, PivotCacheDefinition>
  extras: StylesExtras
}
const books = new WeakMap<PackageReader, Promise<WorkbookControls>>()
const preferredExtras = new WeakMap<PackageReader, WeakMap<TableStyleSource, Map<string, StylesExtras>>>()

function extrasForPreferredStyles(pkg: PackageReader, source: TableStyleSource, names: ReadonlySet<string>): StylesExtras {
  let sources = preferredExtras.get(pkg)
  if (!sources) { sources = new WeakMap(); preferredExtras.set(pkg, sources) }
  let byNames = sources.get(source)
  if (!byNames) { byNames = new Map(); sources.set(source, byNames) }
  const key = JSON.stringify([...names].sort())
  let extras = byNames.get(key)
  if (!extras) {
    extras = parseStylesExtras(source.root, source.colors, names, names)
    byNames.set(key, extras)
  }
  return extras
}

function cancellation(cause: unknown, signal?: AbortSignal): boolean {
  return Boolean(signal?.aborted || (isPackageError(cause) && cause.code === 'cancelled') || (cause instanceof Error && cause.name === 'AbortError'))
}
async function read(pkg: PackageReader, path: string, ordered = false, signal?: AbortSignal): Promise<XmlNode | undefined> {
  try { return ordered ? await pkg.xmlOrdered(path) : await pkg.xml(path) } catch (cause) { if (cancellation(cause, signal)) throw cause; return undefined }
}

async function workbookControls(pkg: PackageReader, request: SheetControlsRequest): Promise<WorkbookControls> {
  const slicerCaches = new Map<string, SlicerCacheDef>(), timelineCaches = new Map<string, TimelineCacheDef>(), pivotCaches = new Map<number, PivotCacheDefinition>()
  // Slicer and timeline cache parts are found by their workbook relationship types (the x14 slicerCaches /
  // x15 timelineCacheRefs lists name the same relationships).
  for (const path of [...(await relationships(pkg, WORKBOOK, '/slicerCache')).values()].slice(0, MAX_CONTROL_CACHES)) {
    const cache = parseSlicerCache(await read(pkg, path, false, request.signal))
    if (cache && !slicerCaches.has(cache.name)) slicerCaches.set(cache.name, cache)
  }
  for (const path of [...(await relationships(pkg, WORKBOOK, '/timelineCache')).values()].slice(0, MAX_CONTROL_CACHES)) {
    const cache = parseTimelineCache(await read(pkg, path, false, request.signal))
    if (cache && !timelineCaches.has(cache.name)) timelineCaches.set(cache.name, cache)
  }
  if (slicerCaches.size || timelineCaches.size) {
    for (const [cacheId, path] of [...(await workbookCaches(pkg, [])).entries()].slice(0, MAX_CONTROL_CACHES)) {
      const root = await read(pkg, path, true, request.signal)
      if (!root) continue
      const cache = parseCacheDefinition(root, cacheId, path, [])
      const id = cache.x14?.pivotCacheId
      if (id !== undefined && !pivotCaches.has(id)) pivotCaches.set(id, cache)
    }
  }
  let extras: StylesExtras = { dxfs: [], tableStyles: new Map(), slicerStyles: new Map() }
  if (request.stylesExtras) extras = request.stylesExtras
  else if (request.styles) { try { extras = parseStylesExtras(request.styles.root, request.styles.colors) } catch { /* built-in styles only */ } }
  if (request.signal?.aborted) throw new PackageError('cancelled', 'operation cancelled', request.sheetPath)
  return { slicerCaches, timelineCaches, pivotCaches, extras }
}

export async function loadSheetControls(pkg: PackageReader, request: SheetControlsRequest): Promise<SheetControls> {
  const out: SheetControls = { slicers: new Map(), timelines: new Map(), diagnostics: [] }
  const report = (diagnostic: ControlDiagnostic) => {
    if (!out.diagnostics.some(d => d.feature === diagnostic.feature && d.message === diagnostic.message)) out.diagnostics.push(diagnostic)
  }
  try {
    if (request.signal?.aborted) throw new PackageError('cancelled', 'operation cancelled', request.sheetPath)
    let book = books.get(pkg)
    if (!book) { book = workbookControls(pkg, request); books.set(pkg, book) }
    const workbook = await book
    if (request.signal?.aborted) throw new PackageError('cancelled', 'operation cancelled', request.sheetPath)
    const parts = async (ids: readonly string[], type: string, kind: 'slicer' | 'timeline') => {
      const rels = await relationships(pkg, request.sheetPath, type)
      const nodes: XmlNode[] = []
      for (const id of ids) {
        const path = rels.get(id), node = path ? await read(pkg, path, false, request.signal) : undefined
        if (node) nodes.push(node)
        else report({ kind: 'parse', feature: `${kind}-part-unreadable`, message: `The ${kind} part of relationship ${id} could not be read from ${request.sheetPath}` })
      }
      return nodes
    }
    const slicerParts = (await parts(request.slicerIds, '/slicer', 'slicer')).flatMap(parseSlicerPart)
    const timelineParts = (await parts(request.timelineIds, '/timeline', 'timeline')).flatMap(parseTimelinePart)
    // Match the paged importer: a referenced custom control style beyond the
    // ordinary style window is selected before parsing, while unused styles
    // remain capped. This keeps resident and paged control pixels identical.
    const preferredStyles = new Set([
      ...slicerParts.map(part => part.style ?? workbook.extras.defaultSlicerStyle ?? 'SlicerStyleLight1'),
      ...timelineParts.map(part => part.style ?? 'TimeSlicerStyleLight1'),
    ])
    const extras = request.styles
      ? extrasForPreferredStyles(pkg, request.styles, preferredStyles)
      : workbook.extras
    const controlsWorkbook = extras === workbook.extras ? workbook : { ...workbook, extras }
    for (const part of slicerParts) {
      if (out.slicers.has(part.name)) continue
      const style = slicerAppearance(part.style ?? controlsWorkbook.extras.defaultSlicerStyle, controlsWorkbook.extras, request.theme, report)
      const built = buildSlicerModel(part, controlsWorkbook.slicerCaches.get(part.cache), controlsWorkbook.pivotCaches, style)
      built.diagnostics.forEach(report)
      if (built.model) out.slicers.set(part.name, built.model)
    }
    for (const part of timelineParts) {
      if (out.timelines.has(part.name)) continue
      const style = timelineAppearance(part.style, controlsWorkbook.extras, request.theme, report)
      const built = buildTimelineModel(part, controlsWorkbook.timelineCaches.get(part.cache), controlsWorkbook.pivotCaches, style)
      built.diagnostics.forEach(report)
      if (built.model) out.timelines.set(part.name, built.model)
    }
  } catch (cause) {
    if (cancellation(cause, request.signal)) throw cause
    report({ kind: 'parse', feature: 'slicer-controls-unreadable', message: `Slicers and timelines of ${request.sheetPath} could not be read; their fallback shapes are shown` })
  }
  return out
}
