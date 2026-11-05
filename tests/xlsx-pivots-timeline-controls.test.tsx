import { cleanup, fireEvent, render, screen, waitFor, within } from '@testing-library/react'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { TimelineController } from '../src/xlsx/pivots/slicers/controller'
import { TimelineLayer } from '../src/components/xlsx/pivot/TimelineLayer'
import type { TimelineModel } from '../src/xlsx/pivots/slicers/model'
import { timelineLayout } from '../src/xlsx/pivots/slicers/paint'
import { buildTimelineModel, timelineLabelLocale, type TimelineCacheDef } from '../src/xlsx/pivots/slicers/parseTimelines'
import type { ScrollWindow } from '../src/xlsx/view/contracts'
import { readFileSync } from 'node:fs'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { XlsxDocument } from '../src/xlsx/types'
import { openPagedWorkbook } from '../src/components/xlsx/pagedWorkbook'
import { pagedWorkbookOf } from '../src/components/pagedDocument'
import { FORCE_PAGED_LIMITS } from '../src/debug/forcePaged'
import { ViewerKernel } from '../src/worker/viewerKernel'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'

const MARCH = Date.UTC(2024, 2, 1)
const APRIL = Date.UTC(2024, 3, 1)
const JANUARY = Date.UTC(2024, 0, 1)
const model: TimelineModel = {
  name: 'NativeTimeline_Date', cacheId: 1, drives: [{ pivotTable: 'Pivot1' }], caption: 'Date', style: 'TimeSlicerStyleLight1', level: 'months',
  range: { start: Date.UTC(2024, 0, 1), end: Date.UTC(2025, 0, 1) }, show: { header: true, selectionLabel: true, timeLevel: true, scrollbar: true },
}
const window: ScrollWindow = { logicalOrigin: { x: 0, y: 0 }, physicalPosition: { x: 0, y: 0 }, physicalExtent: { width: 1000, height: 400 }, zoom: 1, deviceScale: 1 }
const box = { x: 0, y: 0, width: 900, height: 120 }
afterEach(cleanup)

async function controlsDocument(): Promise<XlsxDocument> {
  return parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(`${__dirname}/fixtures/xlsx-pivots/controls/controls-seed.xlsx`))))
}

function workerHarness() {
  const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), importLimits: FORCE_PAGED_LIMITS }, { yield: async () => {} })
  const listeners = new Set<(event: MessageEvent) => void>(), posted: WorkerCommand[] = []
  const port = {
    postMessage(message: unknown, transfer: Transferable[]) {
      const open = message as { kind: string; source?: { kind: string; blob?: Blob } }
      const blob = open.kind === 'open' && open.source?.kind === 'blob' ? open.source.blob : undefined
      const command = structuredClone(blob ? { ...open, source: { kind: 'blob' } } : message, { transfer: transfer as ArrayBuffer[] }) as WorkerCommand
      if (blob && command.kind === 'open') command.source = { kind: 'blob', blob }
      posted.push(command)
      void kernel.handle(command, (reply: WorkerReply, replyTransfer) => {
        const data = structuredClone(reply, { transfer: replyTransfer })
        for (const listener of [...listeners]) listener({ data } as MessageEvent)
      })
    },
    addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
    terminate: vi.fn(),
  }
  return { port, posted }
}
async function renderPagedControls(file: string) {
  const h = workerHarness(), blob = new Blob([new Uint8Array(readFileSync(file))])
  const doc = await openPagedWorkbook(blob, { port: h.port as never, profile: 'D', signal: new AbortController().signal })
  const handle = doc && pagedWorkbookOf(doc)
  if (!doc || !handle) throw new Error('the control workbook did not open paged')
  return { ...render(<InteractiveSheetDoc doc={doc} source={handle.source} />), h }
}
function cellText(container: HTMLElement, row: number, col: number): string | undefined {
  return container.querySelector(`[role=row][aria-rowindex="${row + 2}"]`)?.querySelector(`[role=gridcell][aria-colindex="${col + 2}"]`)?.textContent ?? undefined
}

function setup(hostOverrides: Record<string, unknown> = {}) {
  const apply = vi.fn(async (_timeline: TimelineModel, _action: unknown, _signal: AbortSignal) => ({ kind: 'committed' as const, revision: 1 }))
  const announce = vi.fn()
  const controller = new TimelineController([model], { apply, announce, ...hostOverrides })
  return { apply, announce, controller }
}

describe('timeline control state', () => {
  it('uses measured zh-TW labels for a workbook identified by its linked pivot name', async () => {
    const doc = await controlsDocument()
    const drawing = doc.sheets.flatMap(sheet => sheet.drawings ?? []).find(item => item.content?.kind === 'timeline')
    expect(drawing?.content?.kind).toBe('timeline')
    if (drawing?.content?.kind !== 'timeline') throw new Error('timeline was not resolved')
    const timeline = drawing.content.timeline
    const layout = timelineLayout(timeline, { x: 0, y: 0, width: 900, height: 120 }, 1)
    expect(timeline.labelLocale).toBe('zh-TW')
    expect(layout.selectionLabel?.text).toBe('所有週期')
    expect(layout.level?.text).toBe('月')
    expect(layout.periods[0]?.label).toBe('3月')
    const selected = timelineLayout({ ...timeline, selection: { start: MARCH, end: Date.UTC(2024, 4, 31) } }, { x: 0, y: 0, width: 900, height: 120 }, 1)
    expect(selected.selectionLabel?.text).toBe('2024 3月 - 5月')
  })

  it('keeps the source fallback when linked pivots do not identify a measured label language', async () => {
    expect(timelineLabelLocale({ pivotTables: ['Pivot Sales Summary'] })).toBeUndefined()
    expect(timelineLabelLocale({ pivotTables: ['樞紐分析表1', 'PivotTable1'] })).toBeUndefined()
    const doc = await controlsDocument()
    const known = doc.sheets.flatMap(sheet => sheet.drawings ?? []).find(item => item.content?.kind === 'timeline')
    if (known?.content?.kind !== 'timeline' || !known.content.timeline.appearance) throw new Error('timeline appearance missing')
    const cache: TimelineCacheDef = { name: 'cache', sourceName: 'Date', pivotTables: ['Pivot Sales Summary'], bounds: { start: JANUARY, end: APRIL } }
    const result = buildTimelineModel({ name: 'unknown', cache: 'cache', level: 'months', show: { header: true, selectionLabel: true, timeLevel: true, scrollbar: true } }, cache, new Map(), { style: known.content.timeline.style, appearance: known.content.timeline.appearance })
    expect(result.model).toBeUndefined()
    expect(result.diagnostics[0]?.feature).toBe('timeline-label-locale')
  })

  it('selects a visible month as its complete UTC calendar interval, with inclusive final day', async () => {
    const { apply, controller } = setup()
    expect(await controller.selectTimeline(model.name, Date.UTC(2024, 2, 18), Date.UTC(2024, 2, 18), new AbortController().signal)).toBe(true)
    expect(apply.mock.calls[0]?.[1]).toEqual({ kind: 'selection', start: MARCH, end: APRIL - 86_400_000 })
    expect(controller.model(model.name)?.selection).toEqual({ start: MARCH, end: APRIL - 86_400_000 })
    expect(controller.model(model.name)?.revision).toBe(1)
  })

  it('changes the display level without recomputing pivots and preserves the committed range', async () => {
    const { apply, controller } = setup()
    await controller.selectTimeline(model.name, JANUARY, Date.UTC(2024, 2, 1), new AbortController().signal)
    const selection = controller.model(model.name)?.selection
    expect(await controller.setTimelineLevel(model.name, 'quarters', new AbortController().signal)).toBe(true)
    expect(controller.model(model.name)?.level).toBe('quarters')
    expect(controller.model(model.name)?.selection).toEqual(selection)
    expect(apply).toHaveBeenCalledTimes(1)
  })

  it('clears only after a committed reply and retains the old range on refusal or cancellation', async () => {
    let outcome: unknown = { kind: 'refused', message: 'test refusal' }
    const apply = vi.fn(async () => outcome as { kind: 'committed'; revision: number } | { kind: 'refused'; message: string } | { kind: 'stale' | 'cancelled' })
    const announce = vi.fn()
    const seeded = { ...model, selection: { start: MARCH, end: APRIL - 86_400_000 } }
    const controller = new TimelineController([seeded], { apply, announce })
    expect(await controller.clearTimeline(model.name, new AbortController().signal)).toBe(false)
    expect(controller.model(model.name)?.selection).toEqual(seeded.selection)
    expect(announce).toHaveBeenCalled()
    outcome = { kind: 'committed', revision: 3 }
    expect(await controller.clearTimeline(model.name, new AbortController().signal)).toBe(true)
    expect(controller.model(model.name)?.selection).toBeUndefined()
    controller.dispose()
  })

  it('does not allocate a replacement or call the host when memory admission refuses', async () => {
    const announce = vi.fn(), apply = vi.fn(async () => ({ kind: 'committed' as const, revision: 1 }))
    const controller = new TimelineController([{ ...model, selection: { start: MARCH, end: APRIL - 86_400_000 } }], { apply, announce, admit: () => undefined })
    expect(await controller.clearTimeline(model.name, new AbortController().signal)).toBe(false)
    expect(apply).not.toHaveBeenCalled()
    expect(controller.model(model.name)?.selection).toEqual({ start: MARCH, end: APRIL - 86_400_000 })
    expect(announce).toHaveBeenCalledWith(expect.stringMatching(/memory budget/i))
    controller.dispose()
  })
})

describe('interactive timeline layer', () => {
  it('exposes named periods and a labelled level picker, and period click selects March', async () => {
    const { apply, controller } = setup()
    render(<TimelineLayer model={model} box={box} window={window} viewport={{ width: 1000, height: 400 }} controller={controller} />)
    const march = screen.getByRole('button', { name: 'Select MAR' })
    expect(march.getAttribute('tabindex')).toBe('-1')
    fireEvent.click(march)
    await waitFor(() => expect(apply).toHaveBeenCalledTimes(1))
    expect(apply.mock.calls[0]?.[1]).toEqual({ kind: 'selection', start: MARCH, end: APRIL - 86_400_000 })
    expect(screen.getByRole('button', { name: 'Clear timeline selection' }).hasAttribute('disabled')).toBe(false)
    const picker = screen.getByRole('combobox', { name: 'Timeline level' })
    fireEvent.change(picker, { target: { value: 'quarters' } })
    await waitFor(() => expect(controller.model(model.name)?.level).toBe('quarters'))
    expect(await screen.findByRole('button', { name: 'Select Q1 2024' })).toBeTruthy()
    fireEvent.click(screen.getByRole('button', { name: 'Clear timeline selection' }))
    await waitFor(() => expect(controller.model(model.name)?.selection).toBeUndefined())
    expect(apply).toHaveBeenCalledTimes(2)
  })

  it('fails closed on drag, scroll and cancelled pointer input without changing the committed selection', async () => {
    const { apply, announce, controller } = setup()
    const { rerender } = render(<TimelineLayer model={model} box={box} window={window} viewport={{ width: 1000, height: 400 }} controller={controller} />)
    expect(screen.getByRole('button', { name: /Timeline scrolling unavailable/ }).hasAttribute('disabled')).toBe(true)
    fireEvent.click(screen.getByRole('button', { name: 'Select MAR' }))
    await waitFor(() => expect(controller.model(model.name)?.selection).toEqual({ start: MARCH, end: APRIL - 86_400_000 }))
    const before = controller.model(model.name)?.selection
    const jan = screen.getByRole('button', { name: 'Select JAN 2024' })
    const feb = screen.getByRole('button', { name: 'Select FEB' })
    fireEvent.pointerDown(jan, { pointerId: 1 })
    fireEvent.pointerMove(feb, { pointerId: 1 })
    fireEvent.pointerUp(feb, { pointerId: 1 })
    fireEvent.click(feb)
    fireEvent.pointerCancel(jan, { pointerId: 2 })
    expect(apply).toHaveBeenCalledTimes(1)
    expect(controller.model(model.name)?.selection).toEqual(before)
    expect(announce).toHaveBeenCalledWith(expect.stringMatching(/not been measured/i))
    rerender(<TimelineLayer model={model} box={box} window={window} viewport={{ width: 1000, height: 400 }} controller={controller} />)
  })

  it('supports keyboard range draft and Escape restores the committed range', async () => {
    const { apply, controller } = setup()
    render(<TimelineLayer model={model} box={box} window={window} viewport={{ width: 1000, height: 400 }} controller={controller} />)
    const jan = screen.getByRole('button', { name: 'Select JAN 2024' })
    jan.focus()
    fireEvent.keyDown(jan, { key: 'Enter' })
    fireEvent.keyDown(jan, { key: 'ArrowRight', shiftKey: true })
    fireEvent.keyDown(jan, { key: 'Escape' })
    expect(apply).not.toHaveBeenCalled()
    expect(controller.model(model.name)?.selection).toBeUndefined()
    fireEvent.keyDown(jan, { key: 'Enter' })
    fireEvent.keyDown(jan, { key: 'ArrowRight', shiftKey: true })
    const feb = screen.getByRole('button', { name: 'Select FEB' })
    expect(document.activeElement).toBe(feb)
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowRight', shiftKey: true })
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Enter' })
    await waitFor(() => expect(apply).toHaveBeenCalledTimes(1))
    expect(apply.mock.calls[0]?.[1]).toEqual({ kind: 'selection', start: JANUARY, end: Date.UTC(2024, 2, 31) })
  })

  it('keeps the grid as the only tab stop while Tab cycles through level and clear controls', () => {
    const { controller } = setup()
    const invoker = document.createElement('button')
    document.body.appendChild(invoker)
    render(<TimelineLayer model={model} box={box} window={window} viewport={{ width: 1000, height: 400 }} controller={controller} invoker={invoker} />)
    const period = screen.getByRole('button', { name: 'Select JAN 2024' })
    const level = screen.getByRole('combobox', { name: 'Timeline level' })
    const clear = screen.getByRole('button', { name: 'Clear timeline selection' })
    expect(period.getAttribute('tabindex')).toBe('-1')
    expect(level.getAttribute('tabindex')).toBe('-1')
    expect(clear.getAttribute('tabindex')).toBe('-1')
    period.focus(); fireEvent.keyDown(period, { key: 'Tab' })
    expect(document.activeElement).toBe(level)
    fireEvent.keyDown(level, { key: 'Tab' })
    expect(document.activeElement).toBe(clear)
    fireEvent.keyDown(clear, { key: 'Tab' })
    expect(document.activeElement).toBe(invoker)
  })

  it('mounts the timeline layer on a sheet with no local pivot', async () => {
    const doc = await controlsDocument()
    const pivotSheet = doc.sheets.find(sheet => sheet.pivotTables?.length)!
    const dataSheet = doc.sheets.find(sheet => sheet.name === 'Data')!
    const timeline = pivotSheet.drawings?.find(drawing => drawing.content?.kind === 'timeline')
    expect(timeline).toBeDefined()
    pivotSheet.drawings = pivotSheet.drawings?.filter(drawing => drawing !== timeline)
    timeline!.xEmu = 0
    timeline!.yEmu = 0
    if (timeline!.content?.kind === 'timeline') timeline!.content.timeline.scrollPosition = MARCH
    dataSheet.drawings = [timeline!]
    expect(dataSheet.pivotTables).toBeUndefined()
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    const navigator = await ui.findByRole('combobox', { name: 'Worksheet' })
    const dataOption = [...(navigator as HTMLSelectElement).options].find(option => option.textContent === 'Data')!
    fireEvent.change(navigator, { target: { value: dataOption.value } })
    const layer = await ui.findByRole('group', { name: 'Date timeline' })
    const march = (await within(layer).findAllByRole('button', { name: /^Select / }))[0]!
    fireEvent.click(march)
    await waitFor(() => expect(screen.getByRole('button', { name: 'Clear timeline selection' }).hasAttribute('disabled')).toBe(false))
    expect(ui.queryByRole('button', { name: /Filter and sort/ })).toBeNull()
    fireEvent.change(navigator, { target: { value: [...(navigator as HTMLSelectElement).options].find(option => option.textContent === '工作表1')!.value } })
    await waitFor(() => expect(cellText(ui.container, 4, 1)).toContain('40'))
    expect(screen.queryByRole('group', { name: 'Date timeline' })).toBeNull()
  }, 30_000)

  it('worker/paged timeline selection uses the linked-control transaction', async () => {
    const file = `${__dirname}/fixtures/xlsx-pivots/controls/controls-seed.xlsx`, ui = await renderPagedControls(file)
    const layer = await ui.findByRole('group', { name: 'Date timeline' })
    const first = (await within(layer).findAllByRole('button', { name: /^Select / }))[0]!
    fireEvent.click(first)
    await waitFor(() => expect(ui.h.posted.map(command => command.kind)).toContain('controlCommit'))
    await waitFor(() => expect(within(layer).getByRole('button', { name: 'Clear timeline selection' }).hasAttribute('disabled')).toBe(false))
    expect(ui.h.posted.map(command => command.kind)).toContain('controlApply')
    expect(ui.h.posted.map(command => command.kind)).toContain('controlCommit')
    ui.unmount()
  }, 60_000)

  it('covers painted periods at 200% zoom without double scaling', async () => {
    const { PX_PER_PT } = await import('../src/xlsx/pivots/buttons')
    const { controller } = setup()
    const zoomed: ScrollWindow = { ...window, zoom: 2 }
    render(<TimelineLayer model={model} box={box} window={zoomed} viewport={{ width: 2000, height: 800 }} controller={controller} />)
    // the layout is worksheet-logical; the viewport applies zoom exactly once
    const logical = timelineLayout(model, box, PX_PER_PT).periods[0]!
    const first = screen.getAllByRole('button', { name: /^Select / })[0] as HTMLElement
    expect(parseFloat(first.style.left)).toBeCloseTo(2 * logical.x, 5)
    expect(parseFloat(first.style.width)).toBeCloseTo(2 * logical.width, 5)
    const track = screen.getByRole('button', { name: /Timeline scrolling unavailable/ }) as HTMLElement
    const trackBox = timelineLayout(model, box, PX_PER_PT).track!
    expect(parseFloat(track.style.left)).toBeCloseTo(2 * trackBox.x, 5)
  })
})
