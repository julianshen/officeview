import { test } from 'vitest'
import { strict as assert } from 'node:assert'
import { setFlagsFromString } from 'node:v8'
import { runInNewContext } from 'node:vm'
import { viewerControls, createViewerControlResources } from '../src/components/xlsx/viewerControlsClient'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { DocumentSession } from '../src/worker/viewerSession'
import { semanticsFor } from '../src/xlsx/view/semanticsFor'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { emptyGridSelection } from '../src/xlsx/view/selection'

const WeakRef = (globalThis as any).WeakRef as new <T extends object>(target: T) => { deref(): T | undefined }
setFlagsFromString('--expose_gc')
const gc = runInNewContext('gc') as () => void
const task = () => new Promise<void>(resolve => setTimeout(resolve, 0))
async function collect() { for (let i = 0; i < 8; i++) { await task(); gc() } await task() }
function gate() {
  let open!: () => void, mark!: () => void
  return { wait: new Promise<void>(resolve => open = resolve), reached: new Promise<void>(resolve => mark = resolve), open: () => open(), mark: () => mark() }
}
async function fixture(profile: 'D' | 'M') {
  const n = 16, rows = Array.from({ length: n + 1 }, (_, row) => ({ index: row, cells: [{ row, col: 0, ref: '', value: row ? row : 'Header', styleIndex: 0, style: { numFmtId: 0 } }] }))
  const doc: any = { sheets: [{ name: 'S', sheetId: 's', workbookIndex: 0, rows, cols: [], merges: [], mergeRanges: [], semantics: semanticsFor('1900') }], tables: [] }
  const source = createObjectWorkbookSource(doc, 'd'), store = source.store('s'), ledger = new MainMemoryLedger()
  const resources = createViewerControlResources(profile, undefined, ledger), hidden = createIntervalSet([])
  let gen = 0, publications = 0
  let view: any = { sheetId: 's', revision: 0, order: createIdentityOrder(), baseHidden: hidden, visibility: visibilityFromHidden(hidden), sort: null, filter: null, selection: { ...emptyGridSelection(), active: { row: 9, col: 0 } } }
  const initial = view, sourceJSON = JSON.stringify(doc), selectionJSON = JSON.stringify(view.selection), stamp = { ...store.stamp }
  const owner = () => ({ ...store.stamp, fontGen: 0, viewGen: gen })
  const snapshot = () => ({ owner: owner(), filter: view.filter, sort: view.sort, records: { visible: n + 1, total: n + 1, completeness: 'complete' as const } })
  const controls = await viewerControls({ doc, source, sheetId: 's', resources, ledger, snapshot, view: () => view, publish: next => { publications++; gen++; view = { ...view, ...next }; return snapshot() } }, new AbortController().signal)
  const spec = { range: { r0: 0, r1: n, c0: 0, c1: 0, headerRows: 1 as const, totalsRows: 0 as const }, columns: new Map([[0, { kind: 'number' as const, op: 'gt' as const, first: 8 }]]) }
  return { source, store, ledger, resources, controls, owner, spec, view: () => view, publications: () => publications,
    unchanged() { assert.equal(view, initial); assert.equal(publications, 0); assert.equal(JSON.stringify(doc), sourceJSON); assert.equal(JSON.stringify(view.selection), selectionJSON); assert.deepEqual(store.stamp, stamp) },
    async close() { controls.dispose(); resources.dispose(); await source.release() } }
}
function trace(ledger: MainMemoryLedger) {
  let flagsOwner = '', metadataOwner = '', metadataToken = ''
  let flags: { deref(): Uint8Array | undefined } | undefined, state: { deref(): object | undefined } | undefined
  const native = globalThis.Uint8Array, set = ledger.set.bind(ledger), adopt = ledger.adoptWhileReachable.bind(ledger)
  const charges: number[] = []
  ledger.set = (pool, owner, bytes) => { if (bytes === 145) flagsOwner = owner; if (bytes === 4608) metadataOwner = owner; if (owner === flagsOwner) charges.push(bytes); set(pool, owner, bytes) }
  ledger.adoptWhileReachable = (fromPool, fromOwner, pool, targets, bytes, commit) => {
    const token = adopt(fromPool, fromOwner, pool, targets, bytes, commit)
    if (fromOwner === metadataOwner) { assert.equal(targets.length, 3); metadataToken = token; state = new WeakRef(targets[2]) }
    return token
  }
  globalThis.Uint8Array = new Proxy(native, { construct(target, args) {
    const value = Reflect.construct(target, args, target)
    if (args[0] === 17 && new Error().stack?.includes('viewerControlsClient')) flags = new WeakRef(value)
    return value
  } })
  return { flags: () => flags, flagsBytes: () => ledger.ownerBytes('controls', flagsOwner), charges,
    metadata() { assert.ok(metadataToken); assert.ok(state?.deref()); assert.equal(ledger.ownerBytes('controls', metadataToken), 4608) },
    restore() { globalThis.Uint8Array = native; ledger.set = set; ledger.adoptWhileReachable = adopt } }
}
for (const profile of ['D', 'M'] as const) {
  for (const mode of ['caller-cancel', 'port-close', 'resources-close', 'late-error'] as const) {
    test(`${profile} retains the one flags owner through the actual held producer after ${mode === 'resources-close' ? 'resources-close plus caller cancel' : mode}`, async () => {
      const f = await fixture(profile), t = trace(f.ledger), held = gate(), caller = new AbortController()
      const query = f.store.queryColumn.bind(f.store)
      let intercepted = false, settled = false
      f.store.queryColumn = async (...args) => {
        const page = await query(...args)
        if (!intercepted && t.flags()) { intercepted = true; held.mark(); await held.wait; if (mode === 'late-error') throw new Error('actual late provider failure') }
        return page
      }
      const pending = f.controls.client.filter(f.owner(), f.spec, caller.signal).then(value => { settled = true; return value }, error => { settled = true; throw error })
      // Observe a possible baseline rejection while checking the independent owner defect.
      void pending.catch(() => {})
      try {
        await held.reached
        assert.equal(t.flagsBytes(), 145)
        if (mode === 'port-close') f.controls.dispose()
        else if (mode === 'resources-close') { f.resources.dispose(); caller.abort() }
        else caller.abort()
        await collect()
        assert.equal(t.flags()!.deref()!.byteLength, 17)
        assert.equal(t.flagsBytes(), 145, 'Actual held producer has live flags but its original owner was refunded')
        assert.equal(settled, true, 'Cancellation must settle while the actual producer stays held')
        assert.equal((await pending).kind, 'cancelled')
        f.unchanged()
        t.metadata()
        held.open()
        await collect()
        assert.equal(t.flagsBytes(), 0)
        assert.deepEqual(t.charges, [145, 0])
        f.unchanged()
      } finally { held.open(); await pending.catch(() => {}); await task(); t.restore(); await f.close() }
    }, 10000)
  }
  test(`${profile} keeps flags through actual successful projection then revokes an escaped matcher`, async () => {
    const f = await fixture(profile), t = trace(f.ledger), held = gate()
    const prototype = DocumentSession.prototype as any, project = prototype.projectDraft
    let matcher!: (row: number) => boolean, settled = false
    prototype.projectDraft = async function(...args: any[]) { matcher = args[3]; held.mark(); await held.wait; return project.apply(this, args) }
    const pending = f.controls.client.filter(f.owner(), f.spec, new AbortController().signal).then(value => { settled = true; return value })
    try {
      await held.reached
      await collect()
      assert.equal(settled, false)
      assert.equal(t.flags()!.deref()!.byteLength, 17)
      assert.equal(t.flagsBytes(), 145, 'Producer completion must not refund flags while the actual projection still consumes them')
      assert.equal(matcher(1), false); assert.equal(matcher(9), true)
      f.unchanged()
      held.open()
      assert.equal((await pending).kind, 'committed')
      assert.equal(f.publications(), 1)
      assert.equal(f.view().visibility.visible(1), false); assert.equal(f.view().visibility.visible(9), true)
      assert.equal(t.flagsBytes(), 0)
      assert.deepEqual(t.charges, [145, 0])
      assert.throws(() => matcher(9), /released/)
      await collect()
      assert.equal(t.flags()!.deref(), undefined, 'An escaped matcher must not retain an uncharged flags alias')
      t.metadata()
    } finally { held.open(); await pending.catch(() => {}); prototype.projectDraft = project; t.restore(); await f.close() }
  }, 10000)
}

// This checks the existing composition helper only. The standalone S135 kernel
// has no S133 query-deadline wiring and is not certified by this scratch tree.
for (const late of ['resolve', 'reject'] as const) {
  test(`the copied deadline helper observes late ${late} without another public result`, async () => {
    const { awaitQueryDeadline, bindQueryDeadline } = await import('../src/worker/viewerAbort')
    const caller = new AbortController(), deadline = new AbortController()
    bindQueryDeadline(caller.signal, deadline.signal)
    let resolve!: (value: number) => void, reject!: (error: Error) => void, terminals = 0
    const work = new Promise<number>((yes, no) => { resolve = yes; reject = no })
    const result = awaitQueryDeadline(work, caller.signal).then(
      () => { terminals++; return 'complete' },
      error => { terminals++; assert.equal(error.name, 'AbortError'); return 'cancelled' },
    )
    deadline.abort()
    assert.equal(await result, 'cancelled')
    assert.equal(caller.signal.aborted, false)
    if (late === 'resolve') resolve(9)
    else reject(new Error('late actual evaluator rejection'))
    await task()
    assert.equal(terminals, 1)
  })
}
