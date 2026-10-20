/**
 * The paged viewer's offscreen chart layer (spec 6.6): one list per page or slide, in reading order, whose items
 * are `role="img"` elements named by each chart's description (labelled with the page's true chart count; a
 * `role="note"` entry says when some charts are not listed) (the canvas carries the pixels; the same pattern
 * as the sheet viewer's sheetCharts.tsx). With copying disallowed the descriptions are still read aloud, but the
 * layer carries no selectable text.
 */
import type { ReactElement } from 'react'
import type { Paintable } from '../../render/paint'

const HIDDEN = { position: 'absolute', width: 1, height: 1, margin: -1, padding: 0, overflow: 'hidden', clipPath: 'inset(50%)', whiteSpace: 'nowrap', border: 0 } as const

export function PageCharts({ pages, unit, selectable }: { pages: readonly Paintable[]; unit: 'Page' | 'Slide' | 'Sheet'; selectable: boolean }): ReactElement | null {
  if (!pages.some(page => page.charts?.length)) return null
  return (
    <div data-testid="officeview-page-charts" style={{ ...HIDDEN, ...(selectable ? {} : { userSelect: 'none', WebkitUserSelect: 'none' }) }}>
      {pages.map((page, index) => page.charts?.length ? (
        <ul key={index} aria-label={`Charts on ${unit.toLowerCase()} ${index + 1} (${page.chartTotal ?? page.charts.length})`} data-page={index + 1}>
          {/* A notice (charts not listed, or descriptions refused) is a note, never presented as a chart. */}
          {page.charts.map((chart, i) => <li key={i} role={chart.notice ? 'note' : 'img'} aria-label={chart.description}>{selectable ? chart.description : null}</li>)}
        </ul>
      ) : null)}
    </div>
  )
}
