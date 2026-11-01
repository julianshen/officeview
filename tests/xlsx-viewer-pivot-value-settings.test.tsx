// @vitest-environment jsdom
import { readFileSync } from 'node:fs'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, screen, waitFor, within } from '@testing-library/react'
import { PivotFieldList } from '../src/components/xlsx/pivot/PivotFieldList'
import { PivotFieldMenu } from '../src/components/xlsx/pivot/PivotFieldMenu'
import { PivotValueSettingsDialog } from '../src/components/xlsx/pivot/PivotValueSettingsDialog'
import type { PivotMenuModel } from '../src/components/xlsx/pivot/pivotClient'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import type { XlsxDocument } from '../src/xlsx/types'
import { FIRST_PAINT_WAIT } from './helpers/viewer-timing'
import { openPagedWorkbook } from '../src/components/xlsx/pagedWorkbook'
import { pagedWorkbookOf } from '../src/components/pagedDocument'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { FORCE_PAGED_LIMITS } from '../src/debug/forcePaged'
import { ViewerKernel } from '../src/worker/viewerKernel'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'

afterEach(() => { cleanup(); vi.unstubAllGlobals() })

const FIXTURES = `${__dirname}/fixtures/xlsx-pivots/`
const BASELINE = `${FIXTURES}engine/baseline.xlsx`

async function load(file: string): Promise<XlsxDocument> {
  return parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(file))))
}
const pivotSheet = (doc: XlsxDocument) => doc.sheets.findIndex(sheet => sheet.pivotTables?.length)
const signal = () => new AbortController().signal

function workerHarness(env?: ConstructorParameters<typeof ViewerKernel>[0], intercept?: (command: WorkerCommand) => WorkerReply | undefined) {
  const kernel = new ViewerKernel({ parse: async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)), importLimits: FORCE_PAGED_LIMITS, ...env }, { yield: async () => {} })
  const listeners = new Set<(event: MessageEvent) => void>()
  const posted: WorkerCommand[] = []
  const port = {
    postMessage(message: unknown, transfer: Transferable[]) {
      const open = message as { kind: string; source?: { kind: string; blob?: Blob } }
      const blob = open.kind === 'open' && open.source?.kind === 'blob' ? open.source.blob : undefined
      const command = structuredClone(blob ? { ...open, source: { kind: 'blob' } } : message, { transfer: transfer as ArrayBuffer[] }) as WorkerCommand
      if (blob && command.kind === 'open') command.source = { kind: 'blob', blob }
      posted.push(command)
      const answered = intercept?.(command)
      if (answered) { queueMicrotask(() => { for (const listener of [...listeners]) listener({ data: answered } as MessageEvent) }); return }
      void kernel.handle(command, (out: WorkerReply, replyTransfer) => {
        const data = structuredClone(out, { transfer: replyTransfer })
        for (const listener of [...listeners]) listener({ data } as MessageEvent)
      })
    },
    addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
    terminate: vi.fn(),
  }
  return { port, posted }
}

const pagedFile = (file: string) => new Blob([new Uint8Array(readFileSync(file))])
async function renderPaged(h: ReturnType<typeof workerHarness>, file: string) {
  const doc = await openPagedWorkbook(pagedFile(file), { port: h.port as never, profile: 'D', signal: new AbortController().signal })
  const handle = doc && pagedWorkbookOf(doc)
  if (!doc || !handle) throw new Error('the workbook did not open paged')
  return render(<InteractiveSheetDoc doc={doc} source={handle.source} />)
}

const makeField = (index: number, name: string, axis: 'row' | 'col' | 'page' | 'none' = 'none', calculated = false) => ({
  index, sourceOrder: index, name, axis,
  items: [{ index: 0, label: 'Item 1', hidden: false, collapsed: false }, { index: 1, label: 'Item 2', hidden: false, collapsed: false }],
  sort: 'manual' as const, filters: [], values: 'text' as const,
  expandable: false, defaultDataFn: 'sum' as const, calculated,
  permissions: { dragToRow: true, dragToCol: true, dragToPage: true, dragToData: true, dragOff: true },
})

const makeModel = (overrides?: Partial<PivotMenuModel>): PivotMenuModel => ({
  name: 'TestPivot',
  rows: [0],
  cols: [1],
  pages: [],
  valuesCaption: 'Values',
  nextFilterId: 1,
  date1904: false,
  fields: [
    makeField(0, 'Region', 'row'),
    makeField(1, 'Category', 'col'),
    makeField(2, 'Sales', 'none'),
  ],
  data: [
    { index: 0, field: 2, name: 'Sum of Sales', fn: 'sum', showAs: 'normal', baseField: -1, baseItem: -1 },
  ],
  ...overrides,
})

describe('Pivot value field settings UI', () => {
  it('opens from Values chip in field list and from value header cell, traps Tab and restores focus on Escape', async () => {
    const onApply = vi.fn().mockResolvedValue(true)
    const onClose = vi.fn()
    const model = makeModel()

    // 1. From Values chip in field list:
    render(<PivotFieldList model={model} pivot={0} verified={new Set(['valueSettings'])} onApply={onApply} onClose={onClose} />)
    const actionSelect = screen.getByRole('combobox', { name: 'Actions for Sum of Sales' }) as HTMLSelectElement
    expect([...actionSelect.querySelectorAll('option')].map(o => o.textContent)).toContain('Value field settings…')

    actionSelect.focus()
    fireEvent.change(actionSelect, { target: { value: 'valueSettings' } })

    const dialog = screen.getByRole('dialog', { name: 'Value Field Settings' })
    expect(dialog).toBeTruthy()
    expect(dialog.contains(document.activeElement)).toBe(true)
    expect(document.activeElement).toBe(screen.getByLabelText('Custom Name'))

    // Tab trap:
    const focusable = dialog.querySelectorAll<HTMLElement>('button:not(:disabled),input:not(:disabled),select:not(:disabled)')
    const first = focusable.item(0)!
    const last = focusable.item(focusable.length - 1)!
    const focusFirst = vi.spyOn(first, 'focus')
    const focusLast = vi.spyOn(last, 'focus')

    last.focus()
    const forwardTab = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true })
    last.dispatchEvent(forwardTab)
    expect(forwardTab.defaultPrevented).toBe(true)
    expect(focusFirst).toHaveBeenCalled()

    first.focus()
    const backTab = new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true, cancelable: true })
    first.dispatchEvent(backTab)
    expect(backTab.defaultPrevented).toBe(true)
    expect(focusLast).toHaveBeenCalled()

    // Escape restores focus to invoker
    fireEvent.keyDown(dialog, { key: 'Escape' })
    expect(screen.queryByRole('dialog', { name: 'Value Field Settings' })).toBeNull()
    expect(document.activeElement).toBe(actionSelect)

    cleanup()

    // 2. From Values header / menu:
    render(<PivotFieldMenu owner={{ docId: 'doc', sheetId: '1', valueGen: 0, fontGen: 0, viewGen: 0 }}
      model={model} verified={new Set(['valueSettings'])} field={-2} onApply={onApply} onClose={onClose} onNotice={vi.fn()} />)

    const vsBtn = screen.getByRole('button', { name: /Value field settings/i })
    expect(vsBtn).toBeTruthy()
    vsBtn.focus()
    fireEvent.click(vsBtn)

    const dialogFromMenu = screen.getByRole('dialog', { name: 'Value Field Settings' })
    expect(dialogFromMenu).toBeTruthy()
    expect(dialogFromMenu.contains(document.activeElement)).toBe(true)
    expect(document.activeElement).toBe(within(dialogFromMenu).getByLabelText('Custom Name'))
    fireEvent.keyDown(dialogFromMenu, { key: 'Escape' })
    expect(screen.queryByRole('dialog', { name: 'Value Field Settings' })).toBeNull()
    expect(document.activeElement).toBe(vsBtn)
  })

  it('preserves automatic caption when summary function changes and user has not touched name', async () => {
    const onApply = vi.fn().mockResolvedValue(true)
    const onClose = vi.fn()
    const zhModel = makeModel({
      name: '樞紐分析表1',
      data: [
        { index: 0, field: 2, name: '加總 - Sales', fn: 'sum', showAs: 'normal', baseField: -1, baseItem: -1 },
      ],
    })

    render(<PivotValueSettingsDialog model={zhModel} dataIndex={0} onApply={onApply} onClose={onClose} />)

    const customNameInput = screen.getByLabelText('Custom Name') as HTMLInputElement
    expect(customNameInput.value).toBe('加總 - Sales')

    // Change summary function from Sum to Average
    const avgRadio = screen.getByRole('radio', { name: 'Average' })
    fireEvent.click(avgRadio)

    // The displayed default updates to the zh-TW auto-caption
    expect(customNameInput.value).toBe('平均值 - Sales')

    // Click OK without editing Custom Name
    await act(async () => {
      fireEvent.click(screen.getByRole('button', { name: 'OK' }))
    })

    // onApply receives delta without explicit custom name, letting delta.ts generate it
    expect(onApply).toHaveBeenCalledTimes(1)
    const delta = onApply.mock.calls[0]![0]
    expect(delta.kind).toBe('valueSettings')
    expect(delta.fn).toBe('average')
    expect(delta.name).toBeUndefined()
    expect(onClose).toHaveBeenCalled()
  })

  it('sends explicit custom name when edited, even if summary function changes afterwards', async () => {
    const onApply = vi.fn().mockResolvedValue(true)
    const onClose = vi.fn()
    const model = makeModel()

    render(<PivotValueSettingsDialog model={model} dataIndex={0} onApply={onApply} onClose={onClose} />)

    const customNameInput = screen.getByLabelText('Custom Name') as HTMLInputElement
    fireEvent.change(customNameInput, { target: { value: 'My Revenue' } })

    // Change function to Average
    fireEvent.click(screen.getByRole('radio', { name: 'Average' }))

    // User's custom name is preserved, not overwritten by default
    expect(customNameInput.value).toBe('My Revenue')

    await act(async () => {
      fireEvent.click(screen.getByRole('button', { name: 'OK' }))
    })

    expect(onApply).toHaveBeenCalledTimes(1)
    const delta = onApply.mock.calls[0]![0]
    expect(delta.name).toBe('My Revenue')
    expect(delta.fn).toBe('average')
  })

  it('switching showAs to % Difference From enables Base field and Base item with (previous)/(next) and field items', async () => {
    const onApply = vi.fn().mockResolvedValue(true)
    const onClose = vi.fn()
    const model = makeModel()

    render(<PivotValueSettingsDialog model={model} dataIndex={0} onApply={onApply} onClose={onClose} />)

    // Switch to "Show Values As" tab
    const showAsTab = screen.getByRole('tab', { name: 'Show Values As' })
    fireEvent.click(showAsTab)

    // Initially "normal" (No Calculation) - no base field / base item selectors
    expect(screen.queryByLabelText('Base field')).toBeNull()
    expect(screen.queryByLabelText('Base item')).toBeNull()

    // Switch to % Difference From
    const showAsSelect = screen.getByLabelText('Show values as') as HTMLSelectElement
    fireEvent.change(showAsSelect, { target: { value: 'percentDiff' } })

    // Base field and Base item appear
    const baseFieldSelect = screen.getByLabelText('Base field') as HTMLSelectElement
    const baseItemSelect = screen.getByLabelText('Base item') as HTMLSelectElement
    expect(baseFieldSelect).toBeTruthy()
    expect(baseItemSelect).toBeTruthy()

    // Check base items include (previous), (next) and items from Region (Item 1, Item 2)
    const itemOptions = [...baseItemSelect.querySelectorAll('option')].map(o => o.textContent)
    expect(itemOptions).toContain('(previous)')
    expect(itemOptions).toContain('(next)')
    expect(itemOptions).toContain('Item 1')
    expect(itemOptions).toContain('Item 2')

    // Click OK and verify onApply delta
    await act(async () => {
      fireEvent.click(screen.getByRole('button', { name: 'OK' }))
    })
    expect(onApply).toHaveBeenCalledWith(expect.objectContaining({
      kind: 'valueSettings',
      dataIndex: 0,
      showAs: 'percentDiff',
      baseField: 0,
    }))
    expect(onClose).toHaveBeenCalled()
  })

  it('calculated data field restricts summary function and showAs options', () => {
    const onApply = vi.fn().mockResolvedValue(true)
    const onClose = vi.fn()
    const model = makeModel({
      fields: [
        makeField(0, 'Region', 'row'),
        makeField(1, 'Category', 'col'),
        makeField(2, 'Bonus', 'none', true), // calculated: true
      ],
      data: [
        { index: 0, field: 2, name: 'Sum of Bonus', fn: 'sum', showAs: 'normal', baseField: -1, baseItem: -1 },
      ],
    })

    render(<PivotValueSettingsDialog model={model} dataIndex={0} onApply={onApply} onClose={onClose} />)

    // On summary tab: only Sum is offered
    const summaryRadios = screen.getAllByRole('radio')
    expect(summaryRadios.length).toBe(1)
    expect((summaryRadios[0] as HTMLInputElement).value).toBe('sum')

    // Switch to Show Values As tab: only No Calculation and % of Grand Total are offered
    fireEvent.click(screen.getByRole('tab', { name: 'Show Values As' }))
    const showAsSelect = screen.getByLabelText('Show values as') as HTMLSelectElement
    const options = [...showAsSelect.querySelectorAll('option')].map(o => o.value)
    expect(options).toEqual(['normal', 'percentOfTotal'])
  })

  it('number format preset updates formatted value cells on resident and paged paths', async () => {
    const { createResidentPivotClient } = await import('../src/components/xlsx/pivot/pivotClient')
    const doc = await load(BASELINE), sheet = pivotSheet(doc), source = createObjectWorkbookSource(doc, 'pivot-client')
    const store: any = source.store(doc.sheets[sheet]!.sheetId ?? String(sheet + 1))
    const client = createResidentPivotClient({ doc, ledger: new MainMemoryLedger(), install: commit => { store.setValueOverlay(commit.overlay) } })

    const opened = await client.open(sheet, 0, signal())
    expect(opened.ok).toBe(true)

    // Apply numFmtId 9 (0%)
    const applied = await client.apply(sheet, 0, { kind: 'valueSettings', dataIndex: 0, numFmtId: 9 })
    expect(applied.ok).toBe(true)
    if (applied.ok) {
      expect(applied.model.data[0]?.numFmtId).toBe(9)
    }

    // Now on paged path:
    const h = workerHarness()
    const ui = await renderPaged(h, BASELINE)
    await ui.findByRole('button', { name: /^Collapse East/ }, FIRST_PAINT_WAIT)

    // Open field list
    const fieldListBtn = await ui.findByRole('button', { name: 'Pivot fields' })
    fireEvent.click(fieldListBtn)

    const actionSelect = await ui.findByRole('combobox', { name: 'Actions for Sum of Sales' })
    fireEvent.change(actionSelect, { target: { value: 'valueSettings' } })

    const dialog = await ui.findByRole('dialog', { name: 'Value Field Settings' })
    const numFmtSelect = within(dialog).getByLabelText('Number Format') as HTMLSelectElement
    fireEvent.change(numFmtSelect, { target: { value: '9' } })

    await act(async () => {
      fireEvent.click(within(dialog).getByRole('button', { name: 'OK' }))
    })

    await waitFor(() => {
      expect(h.posted.some(cmd => cmd.kind === 'pivotApply' && (cmd.delta as any).kind === 'valueSettings' && (cmd.delta as any).numFmtId === 9)).toBe(true)
    })
    ui.unmount()
  }, 60_000)

  it('refused apply keeps dialog open with inline role="alert" and leaves committed cells unchanged', async () => {
    const onApply = vi.fn().mockResolvedValue(false)
    const onClose = vi.fn()
    const model = makeModel()

    render(<PivotValueSettingsDialog model={model} dataIndex={0} onApply={onApply} onClose={onClose} />)

    await act(async () => {
      fireEvent.click(screen.getByRole('button', { name: 'OK' }))
    })

    // Dialog stays open
    expect(screen.getByRole('dialog', { name: 'Value Field Settings' })).toBeTruthy()
    expect(onClose).not.toHaveBeenCalled()

    // Inline role="alert"
    const alert = screen.getByRole('alert')
    expect(alert).toBeTruthy()
    expect(alert.textContent).toContain('The value field settings could not be applied.')

    // Also test validation error for blank custom name
    const customNameInput = screen.getByLabelText('Custom Name')
    fireEvent.change(customNameInput, { target: { value: '   ' } })
    await act(async () => {
      fireEvent.click(screen.getByRole('button', { name: 'OK' }))
    })
    expect(screen.getByRole('alert').textContent).toContain('The custom name cannot be blank.')
  })
})
