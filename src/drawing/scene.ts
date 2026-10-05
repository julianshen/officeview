import type { GeometryDefinition } from './geometry'
import type { DrawingStyle } from './style'
import type { ImageSelection, SvgCandidate, SvgVerdict } from '../core/svg'

/** DrawingML picture-use fractions; negative destination values are outsets. */
export interface SceneRectFractions { l: number; t: number; r: number; b: number }

/** Encoded assets are prepared by adapters; painting only uses imageIndex. */
export interface SceneImage {
  data?: Uint8Array
  mime?: string
  partPath?: string
  srcRect?: SceneRectFractions
  fillRect?: SceneRectFractions
  /** Alpha belongs to this use, even when several uses share a media part. */
  opacity?: number
  /** SVG candidate preferred over `data` when its preflight verdict allows. */
  svg?: SvgCandidate
  /** Preflight verdict when the primary `data` bytes are themselves SVG. */
  primarySvgVerdict?: SvgVerdict
  /** Explicitly false when an SVG candidate exists but no raster fallback does. */
  hasRaster?: boolean
  /** Owner-part media path hint for reliable SVG detection. */
  pathHint?: string
  /** Parse/decode selection record shared with the drawing coverage entry. */
  imageSelection?: ImageSelection
}

export interface SceneGroupTransform {
  off: { x: number; y: number }
  ext: { width: number; height: number }
  chOff: { x: number; y: number }
  chExt: { width: number; height: number }
}

/** Source-ordered static drawing. All placement and group coordinates are EMU.
 * Adapters own text/content models, parsing, placement and source diagnostics.
 */
export interface SceneNode<Text = unknown, Image extends SceneImage = SceneImage> {
  xEmu: number
  yEmu: number
  widthEmu: number
  heightEmu: number
  rotationDeg?: number
  flipH?: boolean
  flipV?: boolean
  transformValid?: boolean
  group?: SceneGroupTransform
  children?: SceneNode<Text, Image>[]
  drawingGeometry?: GeometryDefinition
  drawingStyle?: DrawingStyle
  /** Optional legacy simple geometry/style contract. */
  geometry?: 'rect' | 'ellipse' | 'roundRect' | 'other'
  fill?: string
  line?: { color: string; widthEmu?: number }
  image?: Image
  /** Stable adapter-assigned index into a predecoded asset list. */
  imageIndex?: number
  textBody?: Text
}
