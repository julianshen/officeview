/**
 * Target resolution for the timing reader (slideshow spec 6.1): spids are
 * matched against `shape.source.id` of the parsed slide, which already reflects
 * compatibility selection, so an id that exists only in a discarded
 * `mc:Choice` (or `mc:Fallback`) branch is missing. `underlay` is the slide's
 * layout and master artwork (`parseArtworkTree`); an id found only there is
 * `master-only`. Slide ids shadow underlay ids.
 */
import type { PptxShape, PptxSlide } from '../../pptx/types'
import type { TargetResolver } from '../types'

type Paragraphs = ReturnType<TargetResolver['paragraphs']>

function index(shapes: readonly PptxShape[], into: Map<string, PptxShape>, depth = 0): void {
  if (depth > 64) return
  for (const shape of shapes) {
    const id = shape.source?.id
    if (id !== undefined && !into.has(id)) into.set(id, shape)
    if (shape.children) index(shape.children, into, depth + 1)
  }
}

function paragraphsOf(shape: PptxShape): Paragraphs {
  const text = shape.textBody
  if (!text) return []
  return text.paragraphs.map(p => ({
    level: Number.isFinite(p.level) ? p.level : 0,
    runs: p.runs.length,
    empty: !p.runs.some(r => (r.text ?? '').length > 0),
  }))
}

export function shapeResolver(slide: PptxSlide, underlay: readonly PptxShape[] = []): TargetResolver {
  const own = new Map<string, PptxShape>()
  const under = new Map<string, PptxShape>()
  index(slide.shapes, own)
  index(underlay, under)
  const cache = new Map<string, Paragraphs>()
  return {
    shape: spid => own.has(spid) ? 'slide' : under.has(spid) ? 'master-only' : 'missing',
    paragraphs: spid => {
      const shape = own.get(spid)
      if (!shape) return []
      let p = cache.get(spid)
      if (!p) cache.set(spid, p = Object.freeze(paragraphsOf(shape)))
      return p
    },
  }
}
