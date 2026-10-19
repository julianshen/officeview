/**
 * Task 8 fix round 1: Lucide icons are outline icons. Icon.tsx once passed
 * fill={undefined}, which replaced lucide's default fill="none" and painted
 * every chrome icon solid black. jsdom checks the fill attribute; the computed
 * fill is asserted in tests/browser/ui-icon-fill.spec.ts (Chromium + WebKit).
 * The one filled icon is the active filter funnel (currentColor).
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render } from '@testing-library/react'
import { Filter, Printer } from 'lucide-react'
import { Icon } from '../src/components/Icon'
import { GridHeaders } from '../src/components/xlsx/GridHeaders'
import { SearchBar, ZoomControls } from '../src/components/viewerControls'
import { PagedToolbar } from '../src/components/paged/PagedToolbar'
import { gridFixture } from './helpers/xlsx-grid-fixture'

afterEach(cleanup)

const fills = (root: ParentNode) => [...root.querySelectorAll('svg[data-icon]')].map(svg => `${svg.getAttribute('data-icon')}=${svg.getAttribute('fill')}`)

describe('chrome icons are outlines (fill none) except the active filter funnel', () => {
  it('Icon keeps lucide fill="none" unless a fill is given', () => {
    const ui = render(<div><Icon icon={Printer} name="printer" /><Icon icon={Filter} name="filter" fill="currentColor" /></div>)
    expect(fills(ui.container)).toEqual(['printer=none', 'filter=currentColor'])
  })

  for (const compact of [false, true]) {
    it(`paged toolbar (${compact ? 'compact, More menu open' : 'desktop'})`, () => {
      const ui = render(<PagedToolbar kind="document" current={1} total={2} onPrev={vi.fn()} onNext={vi.fn()} zoom={1} onZoomIn={vi.fn()} onZoomOut={vi.fn()} onResetZoom={vi.fn()} onFit={vi.fn()} fitActive={false} onCopy={vi.fn()} copyEnabled copyState="idle" onPrint={vi.fn()} compact={compact} showSearch showZoom focusSearchRequest={0} onEscapeToDocument={vi.fn()} />)
      if (compact) fireEvent.click(ui.getAllByRole('button', { name: 'More options' })[0])
      const all = fills(ui.container)
      expect(all.length).toBeGreaterThan(4)
      expect(all.filter(f => !f.endsWith('=none'))).toEqual([])
    })
  }

  it('shared search and zoom controls', () => {
    const ui = render(<div>
      <SearchBar inputRef={{ current: null }} query="x" onQuery={vi.fn()} count={2} active={0} indexing={false} onPrev={vi.fn()} onNext={vi.fn()} />
      <ZoomControls zoom={1} onIn={vi.fn()} onOut={vi.fn()} onReset={vi.fn()} atFit={false} />
    </div>)
    const all = fills(ui.container)
    expect(all.length).toBe(4)
    expect(all.filter(f => !f.endsWith('=none'))).toEqual([])
  })

  it('sheet headers: only the active filter funnel is filled (currentColor)', () => {
    const p = gridFixture()
    const ui = render(<div role="grid"><GridHeaders projection={p.projection} window={p.window} viewport={p.viewport} columns={[0, 1]} selection={p.selection} sort={null} controlMode
      filter={{ range: { r0: 0, r1: 3, c0: 0, c1: 1, headerRows: 1, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }}
      onSelect={vi.fn()} onSelectAll={vi.fn()} onSortCycle={vi.fn()} onFilter={vi.fn()} onSortDialog={vi.fn()} /></div>)
    const active = ui.getByRole('button', { name: /^Filter column A/ }).querySelector('svg[data-icon="filter"]')
    const idle = ui.getByRole('button', { name: /^Filter column B/ }).querySelector('svg[data-icon="filter"]')
    expect(active?.getAttribute('fill')).toBe('currentColor')
    expect(idle?.getAttribute('fill')).toBe('none')
    const others = [...ui.container.querySelectorAll('svg[data-icon]')].filter(svg => svg !== active)
    expect(others.length).toBeGreaterThan(3)
    expect(others.map(svg => `${svg.getAttribute('data-icon')}=${svg.getAttribute('fill')}`).filter(f => !f.endsWith('=none'))).toEqual([])
  })
})
