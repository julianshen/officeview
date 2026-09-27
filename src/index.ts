/**
 * officeview — React components that render Office documents (Word, PowerPoint,
 * Excel) directly on an HTML <canvas>. Pixel-faithful rendering, not HTML
 * translation. Designed for desktop and mobile browsers.
 */
export { OfficePackage } from './core/zip'
export { parseDocx } from './docx/parse'
export { layoutDocx, renderPages, createMeasurer } from './docx/layout'
export type { DocxDocument, DocxParagraph, DocxTextRun, DocxSection } from './docx/types'
export type { PageLayout, LineBox, MeasureFn } from './docx/layout'
