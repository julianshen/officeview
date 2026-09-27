/**
 * DOCX document model — a flattened, renderer-friendly view of the OOXML
 * document part. Layout (line breaking, page flow) operates on this.
 */

export interface DocxTextRun {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  fontFamily?: string
  fontSizePt?: number
  color?: string
  highlight?: string
  tabWidths?: number[]
  breakBefore?: boolean
}

export type ParagraphAlign = 'left' | 'center' | 'right' | 'justify'

export interface DocxParagraph {
  runs: DocxTextRun[]
  align: ParagraphAlign
  /** Indents in twips. */
  indentLeftTwips?: number
  indentRightTwips?: number
  indentFirstLineTwips?: number
  /** Spacing in twips. */
  spacingBeforeTwips?: number
  spacingAfterTwips?: number
  lineSpacing?: { rule: 'auto' | 'exact' | 'atLeast'; value: number }
  outlineLevel?: number
}

export interface DocxPageMargins {
  topTwips: number
  rightTwips: number
  bottomTwips: number
  leftTwips: number
  headerTwips: number
  footerTwips: number
  gutterTwips: number
}

export interface DocxSection {
  margins: DocxPageMargins
  pageSize: { widthTwips: number; heightTwips: number; orientation: 'portrait' | 'landscape' }
  paragraphs: DocxParagraph[]
}

export interface DocxDocument {
  sections: DocxSection[]
  defaultFontFamily: string
  defaultFontSizePt: number
  styleDefaults: Map<string, { fontFamily?: string; fontSizePt?: number }>
}
