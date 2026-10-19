/**
 * Task 6 (spec §6.4, owner decision A): motion is chrome-only. Menus,
 * dialogs, the mobile filter sheet, the sheet picker, the status line, the
 * zoom readouts and the Copied confirmation animate; the grid and page
 * canvases and the content transform change in one step, and no exiting
 * canvas is ever retained.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor, within } from '@testing-library/react'
import { createElement, forwardRef, type ComponentType } from 'react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { FilterDropdown } from '../src/components/xlsx/FilterDropdown'
import { controlsFixture } from './helpers/xlsx-controls-fixture'
import { workbookOf } from './helpers/xlsx-find'
import { docxModel, mockViewerLayout } from './helpers/paged-viewer'
import * as presets from '../src/components/motion/presets'

// Every `m.<tag>` render records its motion props by data-ov-motion name, so
// the tests read the resolved variants and the first-mount `initial` value
// without depending on jsdom animation timing.
const rec = vi.hoisted(() => ({ props: [] as { name: string; tick?: string; initial: unknown; variants: Record<string, Record<string, unknown>> | undefined }[] }))
vi.mock('motion/react', async (importOriginal) => {
  const mod = await importOriginal<typeof import('motion/react')>()
  const cache = new Map<string, ComponentType>()
  const m = new Proxy(mod.m, {
    get(target, tag: string) {
      const Inner = (target as unknown as Record<string, ComponentType>)[tag]
      if (typeof Inner !== 'function' && typeof Inner !== 'object') return Inner
      if (!cache.has(tag)) {
        cache.set(tag, forwardRef((props: Record<string, unknown>, ref) => {
          if (typeof props['data-ov-motion'] === 'string') rec.props.push({ name: props['data-ov-motion'], tick: props['data-ov-tick'] === undefined ? undefined : String(props['data-ov-tick']), initial: props.initial, variants: props.variants as never })
          return createElement(Inner, { ...props, ref } as never)
        }) as unknown as ComponentType)
      }
      return cache.get(tag)
    },
  })
  return { ...mod, m }
})

const reducedMotion = (reduce: boolean) => vi.stubGlobal('matchMedia', vi.fn((query: string) => ({
  matches: reduce && query.includes('prefers-reduced-motion'), media: query,
  addEventListener: () => {}, removeEventListener: () => {}, addListener: () => {}, removeListener: () => {},
})))

beforeEach(() => { rec.props = [] })
afterEach(() => { cleanup(); vi.restoreAllMocks(); vi.unstubAllGlobals() })

const settled = () => act(async () => { for (let i = 0; i < 4; i++) await new Promise(resolve => setTimeout(resolve, 0)) })
const sheetOf = (sheetId: string, rows: number) => ({ sheetId, cells: [
  { row: 0, col: 0, value: `Head ${sheetId}` },
  ...Array.from({ length: rows }, (_, i) => ({ row: i + 1, col: 0, value: `${sheetId}-${i}` })),
] })
async function sheetViewer() {
  const ui = render(<InteractiveSheetDoc doc={workbookOf([sheetOf('A', 8), sheetOf('B', 5)])} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))
  await settled()
  const screen = ui.container.querySelector('[data-officeview-sheet-screen]') as HTMLElement
  return { ui, screen, canvases: () => [...screen.querySelectorAll('canvas')] }
}
const readout = (root: HTMLElement) => root.querySelector('[aria-label="Reset zoom"] [data-ov-motion="tick"]') as HTMLElement | null

describe('chrome-only motion (spec §6.4, decision A)', () => {
  it('sheet switch swaps the grid canvas in one commit and never retains an exiting canvas', async () => {
    const { ui, screen, canvases } = await sheetViewer()
    const before = canvases()
    expect(before.length).toBeGreaterThan(0)
    // No canvas ever sits inside an animated element.
    expect(screen.querySelectorAll('[data-ov-motion] canvas')).toHaveLength(0)
    act(() => { fireEvent.change(ui.getByRole('combobox', { name: 'Worksheet' }), { target: { value: 'B' } }) })
    // The picker's tick is running: the canvases are the same nodes, none exits.
    const during = canvases()
    expect(during).toHaveLength(before.length)
    during.forEach((canvas, i) => expect(canvas).toBe(before[i]))
    // The new sheet's grid is committed together with the picker change.
    expect(ui.getByRole('grid').getAttribute('aria-label')).toContain('B')
    await settled()
    const after = canvases()
    expect(after).toHaveLength(before.length)
    after.forEach((canvas, i) => expect(canvas).toBe(before[i]))
    expect(screen.querySelectorAll('[data-ov-motion] canvas')).toHaveLength(0)
  })

  it('zoom change applies the content transform in one step', async () => {
    const restore = mockViewerLayout({ width: 1000, height: 800 })
    try {
      const { OfficeDoc } = await import('../src/components/OfficeDoc')
      const ui = render(<OfficeDoc document={docxModel(['one two three'])} />)
      await waitFor(() => expect(ui.container.querySelectorAll('canvas').length).toBeGreaterThan(0))
      const bar = await ui.findByRole('toolbar', { name: 'Document toolbar' })
      const content = ui.getByTestId('officeview-content')
      expect(content.closest('[data-ov-motion]')).toBeNull()
      const values: string[] = []
      const spy = new MutationObserver(records => { for (const r of records) if (r.attributeName === 'style') values.push((r.target as HTMLElement).style.transform) })
      spy.observe(content, { attributes: true, attributeFilter: ['style'] })
      for (const label of ['Zoom in', 'Zoom in', 'Zoom out']) {
        const start = content.style.transform
        values.length = 0
        act(() => { fireEvent.click(within(bar).getByLabelText(label)) })
        await act(async () => { await new Promise(resolve => setTimeout(resolve, 50)) })
        spy.takeRecords().forEach(r => values.push((r.target as HTMLElement).style.transform))
        const distinct = [...new Set(values)].filter(value => value !== start)
        expect(distinct, `${label}: one transform value per action`).toHaveLength(1)
        expect(content.style.transition).toBe('')
      }
      spy.disconnect()
    } finally { restore() }
  })

  it('chrome animates on sheet switch and zoom: picker label, status line and zoom readout are m elements keyed by sheet/zoom', async () => {
    const { ui } = await sheetViewer()
    const picker = ui.getByRole('combobox', { name: 'Worksheet' })
    const pickerTick = picker.closest('[data-ov-motion="tick"]') as HTMLElement
    expect(pickerTick?.getAttribute('data-ov-tick')).toBe('A')
    const status = () => ui.getByRole('status').querySelector('[data-ov-motion]') as HTMLElement
    const statusBefore = status()
    expect(statusBefore.getAttribute('data-ov-tick')).toBe('A')
    picker.focus()
    act(() => { fireEvent.change(picker, { target: { value: 'B' } }) })
    await settled()
    // The native select is never remounted (keyboard focus stays on it).
    expect(ui.getByRole('combobox', { name: 'Worksheet' })).toBe(picker)
    expect(document.activeElement).toBe(picker)
    expect(pickerTick.getAttribute('data-ov-tick')).toBe('B')
    // The status run is a new element keyed by the sheet; the live region stays.
    expect(status()).not.toBe(statusBefore)
    expect(status().getAttribute('data-ov-tick')).toBe('B')
    const zoomBefore = readout(ui.container)!
    expect(zoomBefore.getAttribute('data-ov-tick')).toBe('100')
    act(() => { fireEvent.click(ui.getByRole('button', { name: 'Zoom in' })) })
    const zoomAfter = readout(ui.container)!
    expect(zoomAfter).not.toBe(zoomBefore)
    expect(zoomAfter.getAttribute('data-ov-tick')).toBe('125')
    expect(zoomAfter.textContent).toBe('125%')
    // The keyed remount animates in from the tick's initial state.
    expect(rec.props.filter(p => p.name === 'tick' && String(p.tick) === '125').at(-1)?.initial).toBe('initial')
  })

  it('reduced motion uses zero-duration variants', async () => {
    const durations = (props: typeof rec.props) => props.flatMap(p => Object.values(p.variants ?? {}).map(target => (target.transition as { duration?: number } | undefined)?.duration))
    const mount = async () => {
      const f = controlsFixture(20)
      const ui = render(<FilterDropdown owner={f.snapshot.owner} client={f.client} col={0} applied={null} onClose={() => {}} onNotice={() => {}} onApplied={() => {}} invoker={null} />)
      await settled()
      const own = rec.props.filter(p => p.name === 'popover' || p.name === 'sheet')
      ui.unmount()
      return own
    }
    reducedMotion(false)
    const normal = await mount()
    expect(normal.length).toBeGreaterThan(0)
    const timed = durations(normal).filter(d => d !== undefined)
    expect(timed.length).toBeGreaterThan(0)
    expect(timed.every(d => typeof d === 'number' && d > 0)).toBe(true)
    rec.props = []
    reducedMotion(true)
    const reduced = await mount()
    expect(reduced.length).toBeGreaterThan(0)
    expect(durations(reduced).every(d => d === 0)).toBe(true)
    expect(Object.values(presets.reduce(presets.popover)).every(t => (t as { transition: { duration: number } }).transition.duration === 0)).toBe(true)
  })

  it('variants animate only transform and opacity', () => {
    const allowed = new Set(['opacity', 'scale', 'x', 'y', 'transition'])
    for (const variants of [presets.popover, presets.dialog, presets.sheet, presets.fade, presets.tick]) {
      for (const [label, target] of Object.entries(variants)) {
        expect(Object.keys(target as object).every(key => allowed.has(key)), `${label}: ${Object.keys(target as object)}`).toBe(true)
        const transition = (target as { transition?: { duration: number; ease?: unknown } }).transition
        if (transition) {
          expect(transition.duration).toBeGreaterThanOrEqual(0.12)
          expect(transition.duration).toBeLessThanOrEqual(0.22)
          expect(transition.ease).toEqual([0.1, 0.9, 0.2, 1])
        }
      }
    }
  })

  it('no animation on first mount', async () => {
    const { ui } = await sheetViewer()
    const ticks = rec.props.filter(p => p.name === 'tick')
    expect(ticks.length).toBeGreaterThan(0)
    expect(ticks.every(p => p.initial === false), JSON.stringify(ticks.map(p => p.initial))).toBe(true)
    for (const el of ui.container.querySelectorAll<HTMLElement>('[data-ov-motion]')) {
      expect(el.style.opacity === '' || el.style.opacity === '1').toBe(true)
      expect(el.style.transform === '' || el.style.transform === 'none').toBe(true)
    }
  })
})
