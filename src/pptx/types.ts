/** PPTX document model. All geometry in EMU (rendered via emuToPx). */
export interface PptxTextRun {
  text: string
  bold?: boolean
  italic?: boolean
  fontSizePt?: number
  color?: string
  fontFamily?: string
}

export interface PptxParagraph {
  runs: PptxTextRun[]
  align: 'left' | 'center' | 'right' | 'justify'
  bullet?: boolean
  level: number
}

export interface PptxTextBody {
  paragraphs: PptxParagraph[]
  /** Anchor text vertically within shape: t/ctr/b. */
  anchor: 't' | 'ctr' | 'b'
  /** Inset in EMU (defaults 91440 l/r, 45720 t/b). */
  insetLeftEmu: number
  insetRightEmu: number
  insetTopEmu: number
  insetBottomEmu: number
  wrap: boolean
}

export interface PptxShape {
  /** Geometry in EMU, relative to slide origin. */
  xEmu: number
  yEmu: number
  widthEmu: number
  heightEmu: number
  geometry: 'rect' | 'ellipse' | 'roundRect' | 'other'
  fill?: string
  line?: { color: string; widthEmu?: number }
  textBody?: PptxTextBody
  rotationDeg?: number
}

export interface PptxSlide {
  index: number
  widthEmu: number
  heightEmu: number
  shapes: PptxShape[]
}

export interface PptxDocument {
  slideWidthEmu: number
  slideHeightEmu: number
  slides: PptxSlide[]
}
