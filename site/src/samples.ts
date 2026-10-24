import type { WatermarkOptions } from 'officeview'
import { href } from './layout'

export type Format = 'docx' | 'xlsx' | 'pptx' | 'odt' | 'rtf'

export interface Sample {
  id: string
  title: string
  format: Format
  file: string
  summary: string
  /** What to try in the viewer. */
  tryThis: string[]
  props?: { interactiveSheets?: boolean; watermark?: WatermarkOptions; allowCopy?: boolean; allowPrint?: boolean }
  /** The JSX shown under the demo. */
  code: string
}

export const sampleUrl = (s: Sample) => href(`samples/${s.file}`)

export const SAMPLES: Sample[] = [
  {
    id: 'report',
    title: 'Quarterly business review',
    format: 'docx',
    file: 'quarterly-review.docx',
    summary: 'A four-page Word report: paragraph and heading styles, KPI cards, banded tables with a repeating header row, an inline figure, bulleted and numbered lists, and a header and footer with live page numbers.',
    tryThis: ['Press Ctrl/⌘ + F and search for "revenue".', 'Drag across text to select it, then copy.', 'Pinch, Ctrl + scroll or use the toolbar to zoom.'],
    code: `// Created once, outside the component: a new object would reload the file.
const source = { url: '/files/quarterly-review.docx' }

<OfficeFile data={source} />`,
  },
  {
    id: 'launch',
    title: 'Product launch deck',
    format: 'pptx',
    file: 'product-launch.pptx',
    summary: 'A six-slide PowerPoint deck with master and layout backgrounds, a table, click builds, speaker notes and a sequence of transitions: Push, Vortex and Prism (3-D, drawn with WebGL2), Morph and Curtains.',
    tryThis: ['Click Present in the toolbar (or press F5 with the viewer focused).', 'Advance with → or a click; Esc ends the show.', 'Choose Presenter View from the Present menu to see the notes.'],
    code: `// Created once, outside the component: a new object would reload the file.
const source = { url: '/files/product-launch.pptx' }

const viewer = useRef<OfficeViewerHandle>(null)

<button onClick={() => viewer.current?.present({ from: 'beginning' })}>
  Present
</button>
<OfficeFile ref={viewer} data={source} />`,
  },
  {
    id: 'ledger',
    title: 'Interactive sales ledger',
    format: 'xlsx',
    file: 'sales-ledger.xlsx',
    summary: 'A 2,500-row order ledger opened in the interactive spreadsheet viewer: frozen header row, tiled rendering, absolute zoom, sorting and filtering by column, and workbook-wide find. Sorting and filtering are view-only.',
    tryThis: ['Open a column header menu to sort or filter (try Region).', 'Select a range and copy it into another app.', 'Press Ctrl/⌘ + F to search the whole sheet.'],
    props: { interactiveSheets: true },
    code: `// Created once, outside the component: a new object would reload the file.
const source = { url: '/files/sales-ledger.xlsx' }

<OfficeFile
  data={source}
  interactiveSheets
  viewerWorker={() => new Worker(new URL('./viewer-worker.ts', import.meta.url), { type: 'module' })}
/>`,
  },
  {
    id: 'budget',
    title: 'Operating budget',
    format: 'xlsx',
    file: 'operating-budget.xlsx',
    summary: 'A two-sheet workbook in page mode: merged title cells, styled header rows, currency and percentage number formats, borders, and formulas evaluated by the built-in engine.',
    tryThis: ['Scroll down to reach the Headcount sheet.', 'Compare the totals: they are formulas, recalculated on load.'],
    code: `// Created once, outside the component: a new object would reload the file.
const source = { url: '/files/operating-budget.xlsx' }

<OfficeFile data={source} />`,
  },
  {
    id: 'charts',
    title: 'Chart gallery',
    format: 'docx',
    file: 'chart-gallery.docx',
    summary: 'Every chart family the engine draws natively, in one Word document: column, bar, line, area, pie, doughnut, scatter, bubble, radar, stock and combination charts, plus the Office 2016 waterfall, funnel, histogram, Pareto, box-and-whisker, treemap and sunburst.',
    tryThis: ['Scroll through the pages; each chart names its family in its title.', 'Zoom in: charts are redrawn sharply at every zoom level.'],
    code: `// Created once, outside the component: a new object would reload the file.
const source = { url: '/files/chart-gallery.docx' }

<OfficeFile data={source} />`,
  },
  {
    id: 'protected',
    title: 'Watermarked and protected',
    format: 'docx',
    file: 'quarterly-review.docx',
    summary: 'The same report with a tiled CONFIDENTIAL watermark painted into each page, and with copying and printing turned off. These are deterrents, not DRM. A server can impose both with response headers.',
    tryThis: ['Try to select text or press Ctrl/⌘ + C: nothing is copied.', 'Notice there is no Copy or Print button in the toolbar.'],
    props: { watermark: { text: 'CONFIDENTIAL', placement: 'tile', rotate: -45, opacity: 0.12, color: '#c50f1f' }, allowCopy: false, allowPrint: false },
    code: `// Created once, outside the component: a new object would reload the file.
const source = { url: '/files/quarterly-review.docx' }

<OfficeFile
  data={source}
  watermark={{ text: 'CONFIDENTIAL', placement: 'tile', opacity: 0.12, color: '#c50f1f' }}
  allowCopy={false}
  allowPrint={false}
/>`,
  },
  {
    id: 'odt',
    title: 'Writer meeting notes',
    format: 'odt',
    file: 'design-review.odt',
    summary: 'An OpenDocument Text file with headings, styled runs, a bulleted list, a table with a header row, and a footer with a page-number field. ODT is laid out by the same engine as Word.',
    tryThis: ['Search and selection work exactly as they do for Word files.'],
    code: `// Created once, outside the component: a new object would reload the file.
const source = { url: '/files/design-review.odt' }

<OfficeFile data={source} />`,
  },
  {
    id: 'rtf',
    title: 'RTF letter',
    format: 'rtf',
    file: 'beta-welcome.rtf',
    summary: 'A Rich Text Format letter with an embedded PNG logo, font and colour tables, paragraph spacing, a shaded table, and a Unicode escape, parsed straight from the RTF byte stream.',
    tryThis: ['RTF is detected from its bytes: the same component opens it with no extra options.'],
    code: `// Created once, outside the component: a new object would reload the file.
const source = { url: '/files/beta-welcome.rtf' }

<OfficeFile data={source} />`,
  },
]

export const FORMAT_LABEL: Record<Format, string> = { docx: 'Word', xlsx: 'Excel', pptx: 'PowerPoint', odt: 'Writer', rtf: 'RTF' }
