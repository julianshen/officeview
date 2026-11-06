/**
 * Round-1 review item 1: the retained slicer charge has viewer-level coverage.
 * Reverting `InteractiveSheetDoc.tsx` to `db5ce24a` while keeping the new
 * controller makes these tests fail (the transferred hold is ignored and
 * leaks: replace doubles the charge, sheet removal and unmount never release
 * it). Reverting the controller instead also fails them (no hold is ever
 * transferred, so nothing is retained).
 */
import { readFileSync } from 'node:fs'
import { afterEach, describe, expect, it } from 'vitest'
import { cleanup, fireEvent, render, waitFor, within } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import type { XlsxDocument } from '../src/xlsx/types'
import { FIRST_PAINT_WAIT } from './helpers/viewer-timing'

afterEach(cleanup)

const SEED = `${__dirname}/fixtures/xlsx-pivots/controls/controls-seed.xlsx`
async function seed(): Promise<XlsxDocument> {
  return parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(SEED))))
}
/** Ledger charges this viewer admitted for slicer selections (`slicer-selection:` owners). */
function slicerCharges(ledger: MainMemoryLedger): Array<{ owner: string; bytes: number }> {
  const charges = (ledger as unknown as { charges: Map<string, { pool: string; bytes: number }> }).charges
  const out: Array<{ owner: string; bytes: number }> = []
  for (const [key, charge] of charges) {
    if (key.startsWith('controls\u0000slicer-selection:')) out.push({ owner: key, bytes: charge.bytes })
  }
  return out
}
const slicerBytes = (ledger: MainMemoryLedger): number => slicerCharges(ledger).reduce((sum, charge) => sum + charge.bytes, 0)

async function apples(ui: ReturnType<typeof render>) {
  const group = await ui.findByRole('group', { name: /Product slicer/i }, FIRST_PAINT_WAIT)
  fireEvent.click(within(group).getByRole('button', { name: 'Apples, selected' }))
  await waitFor(() => expect(within(group).getByRole('button', { name: 'Bananas, not selected' })).toBeTruthy(), FIRST_PAINT_WAIT)
  return group
}

describe('viewer-level retained slicer charge (review round 1, item 1)', () => {
  it('resident: replace keeps one charge, sheet removal and unmount release it', async () => {
    const doc = await seed()
    let ledger!: MainMemoryLedger
    const ui = render(<InteractiveSheetDoc doc={doc} onMainLedger={value => { ledger = value }} />)
    const group = await apples(ui)
    expect(ledger).toBeDefined()
    const first = slicerBytes(ledger)
    expect(first).toBeGreaterThan(0)
    expect(slicerCharges(ledger)).toHaveLength(1)

    // Replace: Clear Filter commits a new copy for the same cache. The previous
    // hold is released exactly once, so the pool holds one charge, not two.
    fireEvent.click(within(group).getByRole('button', { name: 'Clear Filter' }))
    await waitFor(() => expect(within(group).getByRole('button', { name: 'Apples, selected' }).getAttribute('aria-pressed')).toBe('true'), FIRST_PAINT_WAIT)
    expect(slicerBytes(ledger)).toBe(first)
    expect(slicerCharges(ledger)).toHaveLength(1)

    // Sheet removed: the viewer remounts for the new document and the old
    // instance's unmount releases its retained charge.
    const removed = { ...doc, sheets: doc.sheets.filter((_, index) => index !== 0) }
    const oldLedger = ledger
    ui.rerender(<InteractiveSheetDoc doc={removed} onMainLedger={value => { ledger = value }} />)
    await waitFor(() => expect(ledger).not.toBe(oldLedger))
    await waitFor(() => expect(slicerCharges(oldLedger)).toHaveLength(0))
    expect(slicerBytes(oldLedger)).toBe(0)
    expect(slicerCharges(ledger)).toHaveLength(0)

    // Unmount: nothing retained remains.
    ui.unmount()
    expect(slicerBytes(ledger)).toBe(0)
  }, 60_000)

  it('resident: retained ledger owners are unique and never stringify an object identity', async () => {
    const doc = await seed()
    let ledger!: MainMemoryLedger
    const ui = render(<InteractiveSheetDoc doc={doc} onMainLedger={value => { ledger = value }} />)
    const group = await apples(ui)
    const charges = slicerCharges(ledger)
    expect(charges).toHaveLength(1)
    // The P1 collision: `${identity}` stringifies the CacheIdentity object and
    // `holdSerial` restarts per controller, so a second controller reuses the
    // first owner's ledger key and overwrites its charge.
    for (const charge of charges) expect(charge.owner).not.toContain('[object Object]')
    const owners = charges.map(charge => charge.owner)
    expect(new Set(owners).size).toBe(owners.length)
    // Replace admits a second hold on the same controller: owners stay unique.
    fireEvent.click(within(group).getByRole('button', { name: 'Clear Filter' }))
    await waitFor(() => expect(within(group).getByRole('button', { name: 'Apples, selected' }).getAttribute('aria-pressed')).toBe('true'), FIRST_PAINT_WAIT)
    const after = slicerCharges(ledger)
    expect(after).toHaveLength(1)
    for (const charge of after) expect(charge.owner).not.toContain('[object Object]')
    ui.unmount()
    expect(slicerBytes(ledger)).toBe(0)
  }, 60_000)

  it('resident: unmounting with a live selection releases its charge', async () => {
    const doc = await seed()
    let ledger!: MainMemoryLedger
    const ui = render(<InteractiveSheetDoc doc={doc} onMainLedger={value => { ledger = value }} />)
    await apples(ui)
    const charged = slicerBytes(ledger)
    expect(charged).toBeGreaterThan(0)
    expect(slicerCharges(ledger)).toHaveLength(1)
    ui.unmount()
    expect(slicerBytes(ledger)).toBe(0)
    expect(slicerCharges(ledger)).toHaveLength(0)
  }, 60_000)
})
