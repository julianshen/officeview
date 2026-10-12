/**
 * Stable first-use media indexing over worksheet scenes. A leaf (types only) so
 * the viewer's sheet sources can index one sheet's pictures without pulling the
 * DrawingML parser into their chunk (S13.2).
 */
import type { XlsxDrawing, XlsxImage, XlsxSheet } from './types'

/** Stable first-use media indexing across all worksheet scenes. */
export function collectXlsxImages(sheets: XlsxSheet[]): XlsxImage[] {
  const images: XlsxImage[] = [], indices = new Map<string, number>()
  const walk = (nodes: XlsxDrawing[]) => {
    for (const node of nodes) {
      if (node.image && (node.image.data || node.image.svg) && node.image.partPath) {
        const svgState = node.image.svg ? (node.image.svg.verdict.ok ? 'ok' : `rej:${(node.image.svg.verdict as { reason: string }).reason}`) : 'none'
        const key = `${node.image.partPath}\u0000${node.image.svg?.partPath ?? ''}\u0000${svgState}\u0000${node.image.hasRaster === false ? 'svg-only' : 'raster'}`
        let index = indices.get(key)
        if (index === undefined) {
          index = images.length
          indices.set(key, index)
          images.push({ data: node.image.data ?? node.image.svg!.bytes, mime: node.image.mime, partPath: node.image.partPath, pathHint: node.image.pathHint ?? node.image.partPath,
            ...(node.image.svg ? { svg: node.image.svg } : {}),
            ...(node.image.primarySvgVerdict ? { primarySvgVerdict: node.image.primarySvgVerdict } : {}),
            hasRaster: node.image.hasRaster,
            ...(node.image.imageSelection ? { imageSelection: node.image.imageSelection } : {}) })
        } else if (node.image.imageSelection && images[index].imageSelection !== node.image.imageSelection) {
          node.image.imageSelection = images[index].imageSelection
        }
        node.imageIndex = index
      }
      if (node.children) walk(node.children)
    }
  }
  for (const sheet of sheets) walk(sheet.drawings ?? [])
  return images
}
