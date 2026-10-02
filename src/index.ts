/**
 * officeview — React components that render Office documents (Word, Excel,
 * PowerPoint) directly on an HTML <canvas>. Pixel-faithful rendering, not
 * HTML translation. Designed for desktop and mobile browsers.
 */
// High-level components
export { OfficeDoc, default as OfficeDocDefault } from './components/OfficeDoc'
export { OfficeFile, useOfficeFile, loadOfficeFile } from './components/OfficeFile'
export type { OfficeFileProps } from './components/OfficeFile'
export type { OfficeDocProps, OfficeDocSource } from './components/OfficeDoc'

// Low-level access
export { OfficePackage } from './core/zip'
// Streaming byte ingestion (used by <OfficeFile> for fetch/Response/Blob/streams)
export { readByteStream, readSource, protectionFromHeaders, PROTECTION_HEADER } from './core/stream'
export type { ByteSource, HttpSource, Progress, ProgressListener, ProtectionPolicy, ProtectionListener } from './core/stream'
// Text search (also used internally by <OfficeDoc>)
export { buildTextIndex, findMatches, stepMatch } from './core/search'
export type { TextIndex, TextIndexPage, IndexLine, TextSpan, SearchMatch, SearchOptions } from './core/search'
export { parseDocx } from './docx/parse'
export { parseOdt } from './odt/parse'
export { parseRtf, isRtf } from './rtf/parse'
export { layoutDocx, renderPages, createMeasurer } from './docx/layout'
export { parseXlsx } from './xlsx/parse'
export { computeMetrics, renderSheet, formatValue } from './xlsx/render'
export { parsePptx } from './pptx/parse'
export { slideMetrics, renderSlide } from './pptx/render'

// Types
export type { DocxDocument, DocxParagraph, DocxSection, DocxTextRun } from './docx/types'
export type { PageLayout, LineBox, MeasureFn } from './docx/layout'
export type { XlsxDocument, XlsxSheet, XlsxCell, XlsxCellStyle } from './xlsx/types'
export type { PptxDocument, PptxSlide, PptxShape } from './pptx/types'

// Text selection helpers (used internally by <OfficeDoc>)
export {
  hitTest,
  normalizeRange,
  rectsForSelectionOnPage,
  selectionSlices,
  textForRange,
  wordRangeAt,
} from './core/selection'
export type { CaretPos, SelectionRange } from './core/selection'
