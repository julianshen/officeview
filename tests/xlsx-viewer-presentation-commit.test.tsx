/**
 * AC-LIFE-02 (spec §12.4): the semantic presentation cache under concurrent
 * rendering. A render that suspends or is abandoned must not change what the
 * committed view's callbacks read, and must not leave explicit charges behind.
 * Collection callbacks never run here, so only explicit releases end a charge.
 * (Adapted from the codex review probes of 69bf1e3.)
 */
import React, { StrictMode, Suspense, startTransition } from 'react'
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { semanticsFor } from '../src/xlsx/view/semanticsFor'
import { workbookOf } from './helpers/xlsx-find'

type Presentation = (at: { row: number; col: number }) => { text: string }
const gate = vi.hoisted(() => ({ suspend: false, committedA: undefined as undefined | ((at: { row: number; col: number }) => { text: string }), wait: new Promise(() => {}) }))
vi.mock('../src/components/xlsx/GridSemantics', async original => {
  const real = await original<typeof import('../src/components/xlsx/GridSemantics')>()
  return { ...real, GridSemantics: (props: Parameters<typeof real.GridSemantics>[0]) => {
    if (props.owner.sheetId === 'A') gate.committedA = props.cellPresentation as Presentation
    // Sheet B formats a label during render, then suspends forever.
    if (props.owner.sheetId === 'B' && gate.suspend) { props.cellPresentation?.({ row: 0, col: 0 }); throw gate.wait }
    return React.createElement(real.GridSemantics, props)
  } }
})

const RealRegistry = globalThis.FinalizationRegistry
beforeEach(() => {
  globalThis.FinalizationRegistry = class { register() {} unregister() { return false } } as unknown as typeof FinalizationRegistry
})
afterEach(() => { cleanup(); vi.restoreAllMocks(); gate.suspend = false; gate.committedA = undefined; globalThis.FinalizationRegistry = RealRegistry })

function setup(strict = false) {
  // Distinct date systems make a cache mix-up visible in the label itself.
  const doc = workbookOf(['A', 'B'].map(sheetId => ({ sheetId, cells: [{ row: 0, col: 0, value: 1 }] })))
  doc.sheets.forEach((sheet, i) => { sheet.semantics = semanticsFor(i ? '1904' : '1900'); sheet.rows[0].cells[0].style = { numFmtId: 164, formatCode: 'yyyy-mm-dd' } })
  const tokens = new Set<string>()
  let ledger!: MainMemoryLedger
  const adopt = MainMemoryLedger.prototype.adoptWhileReachable
  vi.spyOn(MainMemoryLedger.prototype, 'adoptWhileReachable').mockImplementation(function (this: MainMemoryLedger, ...args: Parameters<typeof adopt>) {
    const token = adopt.apply(this, args)
    if (args[0] === 'layout' && args[1].startsWith('format-') && args[2] === 'document') tokens.add(token)
    return token
  })
  const tree = <Suspense fallback={<div>Suspended</div>}><InteractiveSheetDoc doc={doc} onMainLedger={value => { ledger = value }} /></Suspense>
  const ui = render(strict ? <StrictMode>{tree}</StrictMode> : tree)
  const live = () => [...tokens].reduce((sum, token) => sum + ledger.ownerBytes('document', token), 0)
  const suspendOnB = () => act(async () => {
    gate.suspend = true
    startTransition(() => { fireEvent.change(ui.getByRole('combobox', { name: 'Worksheet' }), { target: { value: 'B' } }) })
  })
  return { ui, live, suspendOnB }
}

it('a suspended sheet render does not redirect the committed view to its cache', async () => {
  const { suspendOnB } = setup()
  await waitFor(() => expect(gate.committedA).toBeTypeOf('function'))
  const committed = gate.committedA!
  expect(committed({ row: 0, col: 0 }).text).toBe('1900-01-01')
  await suspendOnB()
  expect(committed({ row: 0, col: 0 }).text, 'A stays committed while B suspends').toBe('1900-01-01')
})

it('a cache whose render never commits is released with the viewer, without GC', async () => {
  const { ui, live, suspendOnB } = setup()
  await waitFor(() => expect(live()).toBeGreaterThan(0))
  await suspendOnB()
  ui.unmount()
  expect(live()).toBe(0)
})

it('StrictMode: the remounted cache refills and releases on final unmount', async () => {
  const { ui, live } = setup(true)
  await waitFor(() => expect(gate.committedA).toBeTypeOf('function'))
  expect(gate.committedA!({ row: 0, col: 0 }).text).toBe('1900-01-01')
  expect(live()).toBeGreaterThan(0)
  ui.unmount()
  expect(live()).toBe(0)
})

it('the hook disposes a superseded committed cache when a new sheet commits', async () => {
  const { useSemanticPresentations } = await import('../src/components/xlsx/InteractiveSheetDoc')
  const { mainFormatBudget } = await import('../src/components/xlsx/mainFormatBudget')
  const ledger = new MainMemoryLedger({ register() {} }), budget = mainFormatBudget(ledger, 64 * 1024 * 1024)
  const sheets = workbookOf(['A', 'B'].map(sheetId => ({ sheetId, cells: [{ row: 0, col: 0, value: 3 }, { row: 1, col: 0, value: `${sheetId} text` }] }))).sheets
  const labels: string[] = []
  function Harness({ sheet }: { sheet: (typeof sheets)[number] }) {
    const present = useSemanticPresentations(sheet, 0, budget)
    for (const row of sheet.rows) for (const cell of row.cells) labels.push(present(cell).text)
    return null
  }
  const ui = render(<Harness sheet={sheets[0]} />)
  const one = ledger.snapshot().byPool.document
  expect(one).toBeGreaterThan(0)
  ui.rerender(<Harness sheet={sheets[1]} />)
  expect(ledger.snapshot().byPool.document, 'only the committed sheet B cache stays charged').toBe(one)
  ui.unmount()
  expect(ledger.snapshot().byPool.document).toBe(0)
  expect(labels).toContain('B text')
})

it('a cache from an initial render that never commits is released by the finalization backstop, exactly once', async () => {
  const { useSemanticPresentations } = await import('../src/components/xlsx/InteractiveSheetDoc')
  const { mainFormatBudget } = await import('../src/components/xlsx/mainFormatBudget')
  // Collection is driven by hand: the ledger's watcher records registrations.
  const registered = new Map<string, object[]>()
  const ledger = new MainMemoryLedger({ register(target, token) { registered.set(token, [...registered.get(token) ?? [], target]) } })
  const budget = mainFormatBudget(ledger, 64 * 1024 * 1024)
  ledger.set('document', 'sentinel', 100)
  const sheet = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 3 }] }]).sheets[0]
  const presented: object[] = []
  function NeverCommits(): null {
    const present = useSemanticPresentations(sheet, 0, budget)
    presented.push(present(sheet.rows[0].cells[0]))
    throw gate.wait
  }
  const ui = render(<Suspense fallback={<div>Suspended</div>}><NeverCommits /></Suspense>)
  ui.unmount()
  const charged = ledger.snapshot().byPool.document
  expect(charged, 'React ran no cleanup for the abandoned render: the backstop still holds it').toBeGreaterThan(100)
  const tokens = [...registered].filter(([, targets]) => targets.some(target => presented.includes(target))).map(([token]) => token)
  expect(tokens.length).toBeGreaterThan(0)
  for (const token of tokens) ledger.collected(token)
  expect(ledger.snapshot().byPool.document, 'collection ends the abandoned charges').toBe(100)
  for (const token of tokens) ledger.collected(token)
  expect(ledger.snapshot().byPool.document, 'a duplicate finalizer releases nothing more').toBe(100)
})
