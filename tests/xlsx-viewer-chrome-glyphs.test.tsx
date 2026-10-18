/**
 * Stream F item 1 (iOS Simulator run): iOS draws emoji-capable symbols
 * (Unicode Emoji=Yes, e.g. U+2195 ↕, U+2196 ↖) with Apple Color Emoji unless a
 * text-presentation selector (U+FE0E) follows, so the header sort glyph showed
 * as a blue emoji tile. Every glyph the viewer chrome and headers draw as text
 * must either be non-emoji or carry U+FE0E.
 */
import { describe, expect, it, vi } from 'vitest'
import { fireEvent, render } from '@testing-library/react'
import { GridHeaders } from '../src/components/xlsx/GridHeaders'
import { SearchBar, ZoomControls } from '../src/components/viewerControls'
import { PagedToolbar } from '../src/components/paged/PagedToolbar'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createScrollWindow } from '../src/xlsx/view/scrollWindow'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { emptyGridSelection, selectCell } from '../src/xlsx/view/selection'
import type { SortSpec } from '../src/xlsx/view/contracts'

/** Emoji-capable code points (ASCII #, *, digits excluded) not followed by U+FE0E. */
function emojiPresentation(text: string): string[] {
  const chars = [...text], bad: string[] = []
  chars.forEach((char, i) => {
    if (char.codePointAt(0)! > 0x7f && /\p{Emoji}/u.test(char) && char !== '︎' && chars[i + 1] !== '︎') bad.push(`U+${char.codePointAt(0)!.toString(16).toUpperCase()} ${char}`)
  })
  return bad
}
const texts = (root: HTMLElement) => [...root.querySelectorAll('button, span')].map(el => el.textContent ?? '')

describe('viewer chrome glyphs use text presentation', () => {
  const projection = (width: number) => createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: width }) })
  const headers = (options: { touch: boolean; sort: SortSpec | null; pending?: number; width?: number }) => {
    const p = projection(options.width ?? 160), viewport = { width: 480, height: 200 }
    return render(<GridHeaders projection={p} window={createScrollWindow(p, viewport)} viewport={viewport} columns={[0, 1, 2]}
      selection={selectCell(emptyGridSelection(), p, { row: 0, col: 1 })} sort={options.sort} touchControls={options.touch} controlMode
      pendingSortColumn={options.pending ?? null} onSelect={vi.fn()} onSelectAll={vi.fn()} onSortCycle={vi.fn()} onFilter={vi.fn()} onSortDialog={vi.fn()} />)
  }
  const sorted: SortSpec = { keys: [{ col: 0, order: 'asc' }, { col: 1, order: 'desc' }] } as unknown as SortSpec
  for (const touch of [false, true]) {
    it(`grid headers (${touch ? 'touch' : 'desktop'}): unsorted, sorted, pending`, () => {
      for (const ui of [headers({ touch, sort: null }), headers({ touch, sort: sorted }), headers({ touch, sort: null, pending: 2 })]) {
        const all = texts(ui.container)
        expect(all.join('')).toMatch(/[↕↖]/u)
        expect(all.flatMap(emojiPresentation)).toEqual([])
        ui.unmount()
      }
    })
  }
  it('the narrow-column touch lane', () => {
    const ui = headers({ touch: true, sort: null, width: 20 })
    fireEvent.click(ui.getByRole('button', { name: 'Actions for column B' }))
    expect(texts(ui.container).flatMap(emojiPresentation)).toEqual([])
  })
  // Spec §6.3: search, zoom and copy controls draw Lucide icons, never glyph text.
  const GLYPHS = /[↑↓+−✓⌘]/u
  const iconControls = (root: HTMLElement, selector: string) => [...root.querySelectorAll<HTMLElement>(selector)]
  it('search bar and zoom controls', () => {
    const ui = render(<div>
      <SearchBar inputRef={{ current: null }} query="x" onQuery={vi.fn()} count={2} active={0} indexing={false} onPrev={vi.fn()} onNext={vi.fn()} />
      <ZoomControls zoom={1} onIn={vi.fn()} onOut={vi.fn()} onReset={vi.fn()} atFit={false} />
    </div>)
    expect(texts(ui.container).flatMap(emojiPresentation)).toEqual([])
    const buttons = iconControls(ui.container, '[aria-label="Previous match"], [aria-label="Next match"], [aria-label="Zoom in"], [aria-label="Zoom out"]')
    expect(buttons).toHaveLength(4)
    for (const button of buttons) {
      expect(button.querySelector('svg[data-icon][aria-hidden="true"]'), button.getAttribute('aria-label')!).toBeTruthy()
      expect(button.textContent).not.toMatch(GLYPHS)
    }
    expect(ui.container.textContent).not.toMatch(GLYPHS)
  })
  it('paged toolbar copy button in every state', () => {
    for (const state of ['idle', 'done', 'failed'] as const) {
      const ui = render(<PagedToolbar kind="document" current={1} total={2} onPrev={vi.fn()} onNext={vi.fn()} zoom={1} onZoomIn={vi.fn()} onZoomOut={vi.fn()} onResetZoom={vi.fn()} onFit={vi.fn()} fitActive={false} onCopy={vi.fn()} copyEnabled copyState={state} onPrint={vi.fn()} compact={false} showSearch={false} showZoom focusSearchRequest={0} onEscapeToDocument={vi.fn()} />)
      const pill = ui.getByTestId('officeview-copy-selection')
      expect(pill.querySelector('svg[data-icon][aria-hidden="true"]'), state).toBeTruthy()
      expect(pill.textContent, state).not.toMatch(GLYPHS)
      expect(texts(ui.container).flatMap(emojiPresentation)).toEqual([])
      ui.unmount()
    }
  })
})
