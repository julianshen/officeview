/**
 * Spec §§10.1, 12.4 (overlap accounting): a worker view's buffers stay charged
 * exactly while a real owner holds them. The workbook holds its current view;
 * every display, copy or pending operation that keeps a view takes its own
 * hold. One viewer's revision is never proof that the others let go, and a
 * replaced view with no owner ends at once without waiting for a render.
 *
 * Collection callbacks never run here: only explicit releases end a charge.
 */
import { describe, expect, it } from 'vitest'
import { MainMemoryLedger, type LedgerHold } from '../src/xlsx/view/mainLedger'
import { DisplayHolds } from '../src/components/xlsx/displayHolds'
import { RemoteWorkbook, type RemoteSheetView } from '../src/worker/viewerRemoteSource'
import type { ViewerClient } from '../src/worker/viewerClient'
import type { ViewCommitDelta } from '../src/xlsx/view/contracts'
import { initialIdentity, initialWorkbookOwner } from '../src/worker/viewerProtocol'

function fixture() {
  const remote = new RemoteWorkbook({} as ViewerClient, { docId: 'd', valuePolicy: 'cached', sheets: [{ sheetId: 's', name: 'S' }], diagnostics: [], layouts: new Map() })
  const internals = remote as unknown as { adopt(sheetId: string, delta: ViewCommitDelta, spec: object): RemoteSheetView; mainLedger: MainMemoryLedger }
  // No collection callback ever runs (the default registry would be timing-dependent).
  ;(internals as { mainLedger: MainMemoryLedger }).mainLedger = new MainMemoryLedger({ register() {} })
  const adopt = (revision: number, bands = true) => {
    const owner = { ...initialIdentity('d', 's'), viewGen: revision }
    const delta: ViewCommitDelta = {
      owner, workbookOwner: { ...initialWorkbookOwner('d', [{ sheetId: 's' }]), globalViewGen: revision, sheetVector: [{ sheetId: 's', revision, valueGen: 0 }] },
      bands: bands ? [{ r0: 0, r1: 3, sourceAtOffset: Uint32Array.of(3, 2, 1, 0), displayOffsetOfSource: Uint32Array.of(3, 2, 1, 0) }] : [], visibility: [], metrics: { rows: [], cols: [] }, counts: { visible: 4, total: 4 },
      selection: { parts: [], active: null, anchor: null } as unknown as ViewCommitDelta['selection'],
    }
    return internals.adopt('s', delta, {})
  }
  const pool = () => remote.mainLedger.snapshot().byPool.viewBuffers
  return { remote, adopt, pool }
}

describe('RemoteWorkbook view ownership', () => {
  it('MAJOR-1: an old order a sibling or continuation still holds stays charged after another view is adopted', () => {
    const { remote, adopt, pool } = fixture()
    const old = adopt(1)
    const one = pool()
    // A sibling viewer's committed display (or a copy/find continuation) keeps revision 1.
    const sibling = remote.holdView(old)!
    // Another viewer advances: the workbook adopts revision 2 and that viewer
    // drops everything it held of revision 1.
    adopt(2)
    const advancing = remote.holdView(remote.view('s'))!
    expect(old.order.displayToSource(0)).toBe(3)
    expect(remote.mainLedger.hasLifetime(old.order), 'the sibling has not released its old order').toBe(true)
    expect(pool(), 'both views are live, so both are charged').toBe(2 * one)
    sibling.release()
    expect(remote.mainLedger.hasLifetime(old.order)).toBe(false)
    expect(pool()).toBe(one)
    advancing.release()
    expect(pool(), 'the workbook still holds its current view').toBe(one)
  })
  it('MAJOR-2: adoptions no viewer ever renders settle each replaced view at once and retain nothing', () => {
    const { remote, adopt, pool } = fixture()
    const one = (adopt(1), pool())
    const replaced: WeakRef<RemoteSheetView>[] = []
    for (let revision = 2; revision <= 12; revision++) {
      replaced.push(new WeakRef(remote.view('s')))
      adopt(revision)
      expect(pool(), `revision ${revision}`).toBe(one)
    }
    // Nothing in the workbook keeps an obsolete view: its only own reference is the current one.
    const fields = Object.values(remote as unknown as Record<string, unknown>)
    const reachable = (value: unknown, seen = new Set<unknown>()): boolean => {
      if (!value || typeof value !== 'object' || seen.has(value)) return false
      seen.add(value)
      if (replaced.some(ref => ref.deref() === value)) return true
      if (value instanceof Map) return [...value.values()].some(item => reachable(item, seen))
      if (Array.isArray(value)) return value.some(item => reachable(item, seen))
      return false
    }
    expect(fields.some(field => reachable(field))).toBe(false)
  })
  it('a hold taken on the identity-order base view or an identity-order adoption never touches the singleton', () => {
    const { remote, adopt, pool } = fixture()
    expect(remote.holdView(remote.view('s')), 'the base view owns no buffers').toBeNull()
    const first = adopt(1, false)
    expect(first.order.kind).toBe('identity')
    expect(remote.mainLedger.hold(first.order), 'the shared identity order names no charge').toBeNull()
    const visibilityOnly = pool()
    const held = remote.holdView(first)!
    adopt(2, false)
    expect(pool()).toBe(2 * visibilityOnly)
    held.release(); held.release()
    expect(pool()).toBe(visibilityOnly)
  })
  it('close releases the workbook holds but not a display still holding its view', async () => {
    const { remote, adopt, pool } = fixture()
    const view = adopt(1)
    const display = remote.holdView(view)!
    ;(remote as unknown as { client: ViewerClient }).client = { send: () => new Promise(() => {}), dispose() {}, nextRequestId: () => 'x' } as unknown as ViewerClient
    void remote.close()
    expect(pool()).toBeGreaterThan(0)
    display.release()
    expect(pool()).toBe(0)
  })
})

describe('DisplayHolds (one viewer\'s committed-display ownership)', () => {
  const setup = () => {
    const ledger = new MainMemoryLedger({ register() {} })
    const entry = (bytes: number) => { const buffer = {}; ledger.chargeWhileHeld('viewBuffers', [buffer], bytes); return { buffer } }
    const holds = new DisplayHolds<{ buffer: object }>(held => [ledger.hold(held.buffer)].filter((hold): hold is LedgerHold => hold !== null))
    return { ledger, entry, holds }
  }
  it('a replaced entry is released only by the commit of the display that replaced it (or a later one)', () => {
    const { ledger, entry, holds } = setup()
    const a = entry(10), b = entry(20)
    holds.own(a)
    holds.own(b); holds.retire(a, 5)
    holds.committed(4)
    expect(ledger.bytes, 'display 4 (still showing a) committed: a stays').toBe(30)
    holds.committed(5)
    expect(ledger.bytes).toBe(20)
    holds.committed(6)
    expect(ledger.bytes).toBe(20)
  })
  it('a continuation keeps a retired entry until it ends; dispose and restore follow unmount and remount', () => {
    const { ledger, entry, holds } = setup()
    const a = entry(10), b = entry(20)
    holds.own(a)
    const copy = holds.extra(a)
    holds.own(b); holds.retire(a, 1); holds.committed(1)
    expect(ledger.bytes, 'the copy still reads a').toBe(30)
    copy(); copy()
    expect(ledger.bytes).toBe(20)
    holds.dispose()
    expect(ledger.bytes).toBe(0)
    holds.dispose()
    holds.restore()
    expect(ledger.bytes, 'remounted: b is shown again').toBe(20)
    holds.restore()
    expect(ledger.bytes).toBe(20)
    holds.dispose()
    expect(ledger.bytes).toBe(0)
  })
})
