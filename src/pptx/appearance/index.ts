/**
 * The lazy slide appearance enrichment (slideshow spec 8, D2). Never imported statically from eager code:
 * `getPaintables` loads it through `import()`, and `src/index.ts` exports it behind one for direct `renderSlide`
 * callers. It attaches each slide's resolved background and its master and layout non-placeholder artwork as
 * `slide.underlay`, expressed as the scene nodes and fills the existing painter draws.
 */
import { pptxImageKey, pptxPackageHandle, slideTheme } from '../parse'
import type { ThemeContext } from '../../drawing/style'
import type { PptxDiagnostic, PptxDocument, PptxPackageHandle, PptxShape } from '../types'
import { HandlePackage, HANDLE_AGGREGATE_BYTES, HANDLE_PART_BYTES } from './handle-package'
import { validateSlideSize } from './size'
import { backgroundNode, resolveSlideBackground } from './background'
import { flagUseBgFill, masterArtwork, type ArtworkCache } from './master-art'
import { getChildren } from '../../core/xml'

/** One run per document: later and concurrent calls share it, so enrichment happens once. */
const runs = new WeakMap<PptxDocument, Promise<void>>()

const identity = (d: PptxDiagnostic) => `${d.kind}\u0000${('feature' in d ? d.feature : undefined) ?? ''}\u0000${d.message}`
/** Document-level diagnostics are recorded once per document. */
function documentDiagnostic(doc: PptxDocument, diagnostic: PptxDiagnostic): void {
  doc.diagnostics ??= []
  if (!doc.diagnostics.some(existing => identity(existing) === identity(diagnostic))) doc.diagnostics.push(diagnostic)
}

/**
 * Enrich every slide of a parsed document with its inherited appearance. Idempotent. A document without a
 * package handle (constructed, cloned or posted from another realm) or with an invalid slide size is left
 * unchanged with a `no-package` or `slide-size` diagnostic, and a later call (after a reopen) may still enrich it.
 * `slide.background` stays the solid approximation (cleared when the resolved background paints nothing); new pictures are appended to `doc.images`.
 */
export function enrichSlideAppearance(doc: PptxDocument): Promise<void> {
  const existing = runs.get(doc)
  if (existing) return existing
  const handle = pptxPackageHandle(doc)
  if (!handle) {
    documentDiagnostic(doc, { kind: 'no-package', message: 'The document has no package handle; slides paint without backgrounds and master artwork' })
    return Promise.resolve()
  }
  const size = validateSlideSize(doc.slideWidthEmu, doc.slideHeightEmu)
  if (!size.ok) {
    documentDiagnostic(doc, size.diagnostic)
    return Promise.resolve()
  }
  const run = enrich(doc, handle)
  runs.set(doc, run)
  return run
}

async function enrich(doc: PptxDocument, handle: PptxPackageHandle): Promise<void> {
  // One reader per run: one 32 MiB aggregate and one node budget, pictures up to the aggregate, XML parts 4 MiB.
  const pkg = HandlePackage.over(handle, { partBytes: HANDLE_AGGREGATE_BYTES, textBytes: HANDLE_PART_BYTES })
  const artwork: ArtworkCache = new Map()
  const themes = new Set<ThemeContext>()
  try {
    for (let i = 0; i < doc.slides.length; i++) {
      // slideParts is aligned with doc.slides (not slide.index, which counts skipped sldIds).
      const slide = doc.slides[i], path = handle.slideParts[i]
      if (!path) continue
      try {
        // slide.theme carries the slide's effective colour map (clrMap with clrMapOvr) for phClr and scheme colours.
        const theme = slide.theme ?? (await slideTheme(pkg, path)).context
        themes.add(theme)
        const { background, diagnostics } = await resolveSlideBackground(pkg, path, theme)
        // A failed artwork read loses only the artwork: the resolved background is still assigned.
        const art = await masterArtwork(pkg, path, theme, artwork).catch((error: unknown) => ({ master: [], layout: [], diagnostics: [unavailable('Master and layout artwork', error)] }))
        const node = backgroundNode(background, slide.widthEmu, slide.heightEmu)
        slide.underlay = [...(node ? [node] : []), ...art.master, ...art.layout]
        // useBgFill shapes on the slide itself also show the resolved background (the eager parse could only use a solid colour).
        flagUseBgFill(slide.shapes, getChildren(getChildren(await pkg.xmlOrdered(path), 'cSld')[0], 'spTree')[0], path)
        // No background anywhere: useBgFill shapes still show the default white slide, not their authored fill.
        // An explicit noFill paints nothing either. Only a resolved solid keeps the eager solid: any other background
        // (none, gradient, pattern, picture) stops inheritance, so the solid the eager parse inherited must not show through.
        const painted = node?.drawingStyle?.fill?.kind === 'none' ? undefined : node
        if (painted?.drawingStyle?.fill?.kind !== 'solid') delete slide.background
        slide.backdrop = painted ?? backgroundNode({ kind: 'fill', fill: { kind: 'solid', color: { r: 255, g: 255, b: 255, a: 1 } } }, slide.widthEmu, slide.heightEmu)
        slide.diagnostics ??= []
        slide.diagnostics.push(...diagnostics, ...art.diagnostics)
      } catch (error) {
        slide.diagnostics ??= []
        slide.diagnostics.push(unavailable('Slide background and master artwork', error))
      }
    }
  } finally {
    pkg.dropCache()
  }
  // Theme issues are shared by every slide on a master: surface each once per document.
  for (const theme of themes) for (const issue of theme.issues) documentDiagnostic(doc, issue)
  indexUnderlayImages(doc)
}

const unavailable = (what: string, error: unknown): PptxDiagnostic =>
  ({ kind: 'appearance-unavailable', message: `${what} could not be resolved: ${error instanceof Error ? error.message : String(error)}` })

/** Underlay pictures join `doc.images` under parsePptx's complete-pair key, so a picture shared with slide
 * content decodes once, and each node gets the `imageIndex` the painter reads. */
function indexUnderlayImages(doc: PptxDocument): void {
  const indices = new Map(doc.images.map((image, index) => [pptxImageKey(image), index]))
  const visit = (shapes: PptxShape[]) => {
    for (const shape of shapes) {
      if (shape.image) {
        const key = pptxImageKey(shape.image)
        let index = indices.get(key)
        if (index === undefined) {
          index = doc.images.length
          const { opacity: _opacity, srcRect: _srcRect, fillRect: _fillRect, ...part } = shape.image
          doc.images.push(part)
          indices.set(key, index)
        } else if (shape.image.imageSelection && doc.images[index].imageSelection !== shape.image.imageSelection) {
          shape.image.imageSelection = doc.images[index].imageSelection
        }
        shape.imageIndex = index
      }
      if (shape.children) visit(shape.children)
    }
  }
  for (const slide of doc.slides) visit(slide.underlay ?? [])
}
