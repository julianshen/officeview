import { test } from 'vitest'
const WeakRef = (globalThis as any).WeakRef as new <T extends object>(value: T) => {
  deref(): T | undefined
}
type WeakRef<T extends object> = {
  deref(): T | undefined
}
import { strict as assert } from 'node:assert'
import { viewerControls, createViewerControlResources } from '../src/components/xlsx/viewerControlsClient'
import { MainMemoryLedger, estimatePayloadBytes } from '../src/xlsx/view/mainLedger'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { DocumentSession } from '../src/worker/viewerSession'
import { semanticsFor } from '../src/xlsx/view/semanticsFor'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { emptyGridSelection } from '../src/xlsx/view/selection'
for (const profile of ['D', 'M'] as const)
  test(`charges raw lanes alongside the one display owner on ${profile}`, async () => {
    const signal = () => new AbortController().signal
    const freshTask = () => new Promise<void>(r => setTimeout(r, 0))
    function gate() { let open!: () => void, hit!: () => void; return { wait: new Promise<void>(r => open = r), hit: new Promise<void>(r => hit = r), open: () => open(), mark: () => hit() } }
    const gates = [gate(), gate(), gate()]
    const originals = { Uint8Array, Float64Array, Uint32Array, Int32Array }
    const sessionOriginals = { sort: DocumentSession.prototype.sort, filter: DocumentSession.prototype.filter }
    try {
      const allocations: {
        name: string
        bytes: number
        ref: WeakRef<any>
        stack: string
      }[] = []
      for (const name of ['Uint8Array', 'Float64Array', 'Uint32Array', 'Int32Array'] as const) {
        const real = globalThis[name];
        (globalThis as any)[name] = new Proxy(real, {
          construct(target, args) {
            const value = Reflect.construct(target, args, target); const stack = new Error().stack ?? ''; if (stack.split('\n')[2]?.includes('/xlsx/view/sort.ts:'))
              allocations.push({ name, bytes: value.byteLength, ref: new WeakRef(value), stack }); return value
          }
        })
      }
      const n = 4096, styles = [{ numFmtId: 0 }, { numFmtId: 0, formatCode: '"' + 'z'.repeat(3000) + '"0' }]
      const rows = Array.from({ length: n + 1 }, (_, r) => ({ index: r, cells: Array.from({ length: 3 }, (_, col) => ({ row: r, col, ref: '', value: r === 0 ? 'H' + col : col === 2 && r === n ? 'label' : col === 2 ? 0 : n - r, styleIndex: col === 2 ? 1 : 0, style: styles[col === 2 ? 1 : 0] })) }))
      const doc: any = { sheets: [{ name: 'S', sheetId: 's', workbookIndex: 0, rows, cols: [], merges: [], mergeRanges: [], semantics: semanticsFor('1900') }], tables: [] }
      const originalDoc = JSON.stringify(doc), source = createObjectWorkbookSource(doc, 'd'), store = source.store('s'), stamp = { ...store.stamp }
      const ledger = new MainMemoryLedger()
      ledger.chargeWhileReachable('document', [doc], estimatePayloadBytes(doc))
      const kernelOwners = new Set<string>(), seenOwners = new Set<string>()
    const setCharge = ledger.set.bind(ledger)
    ledger.set = (pool, owner, bytes) => {
      if (bytes > 0 && !seenOwners.has(owner)) {
        seenOwners.add(owner)
        if (new Error().stack?.includes('Budget.charge')) kernelOwners.add(owner)
      }
      setCharge(pool, owner, bytes)
    }
    const resources = createViewerControlResources(profile, undefined, ledger)
      const hidden = createIntervalSet([]), selection = { ...emptyGridSelection(), active: { row: 17, col: 1 } }
      let gen = 0, published = 0, view: any = { sheetId: 's', revision: 0, order: createIdentityOrder(), baseHidden: hidden, visibility: visibilityFromHidden(hidden), sort: null, filter: null, selection }
      const originalView = view, originalSelection = JSON.stringify(selection)
      const owner = () => ({ ...store.stamp, fontGen: 0, viewGen: gen }), snapshot = () => ({ owner: owner(), filter: view.filter, sort: view.sort, records: { visible: n + 1, total: n + 1, completeness: 'complete' as const } })
      const sessions: any[] = []
      for (const method of ['sort', 'filter'] as const) {
        const original = (DocumentSession.prototype as any)[method];
        (DocumentSession.prototype as any)[method] = function(...args: any[]) { sessions.push({ kind: method, session: this })
          if (method === 'sort') {
            const prepared = this.preparedStore.bind(this)
            // Hold the actual sort continuation above the interruptible source.
            this.preparedStore = (sheetId: string) => {
              const target = prepared(sheetId)
              return new Proxy(target, { get(t: any, prop) {
                if (prop === 'queryColumn') return async (req: any, opts: any) => {
                  const page = await t.queryColumn(req, opts)
                  if (req.col < 2 && !heldPages.includes(page)) {
                    heldPages.push(page)
                    gates[req.col].mark()
                    await gates[req.col].wait
                  }
                  return page
                }
                const value = Reflect.get(t, prop)
                return typeof value === 'function' ? value.bind(t) : value
              } })
            }
          }
          return original.apply(this, args) }
      }
      const controls = await viewerControls({ doc, source, sheetId: 's', resources, ledger, snapshot, view: () => view, publish: next => { published++; view = { ...view, ...next }; gen++; return snapshot() } }, signal())
      const counts = [0, 0, 0], heldPages: any[] = []
      const query = store.queryColumn.bind(store)
      store.queryColumn = async (req: any, opts: any) => {
        const page = await query(req, opts); counts[req.col]++; if (req.col === 2 && counts[2] === 3) {
          heldPages.push(page)
          gates[req.col].mark()
          await gates[req.col].wait
        } return page
      }
      const range: any = { r0: 0, r1: n, c0: 0, c1: 2, headerRows: 1, totalsRows: 0 }
      const aborts = [new AbortController(), new AbortController(), new AbortController()]
      const before = ledger.snapshot()
      const pending = [controls.client.sort(owner(), { range, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }, aborts[0].signal), controls.client.sort(owner(), { range, keys: [{ col: 1, order: 'asc' }], caseSensitive: false }, aborts[1].signal), controls.client.filter(owner(), { range, columns: new Map([[2, { kind: 'text', op: 'contains', first: '0', wildcards: false }]]) }, aborts[2].signal)]
      await Promise.all(gates.map(g => g.hit))
      await freshTask()
      await freshTask()
      const live = allocations.filter(a => a.ref.deref() && a.bytes > 0 && (a.name === 'Uint8Array' && a.bytes === n || a.name === 'Float64Array' && a.bytes === n * 8)), rawBytes = live.reduce((s, a) => s + a.bytes, 0)
      const during = ledger.snapshot(), privateDuring = sessions.map(x => ({ kind: x.kind, memory: x.session.memoryStats(), prepared: x.session.preparedStats() }))
      assert.equal(rawBytes, 2 * n * 9)
      assert.equal(resources.cache.usedBytes, 22690)
    assert.equal(ledger.ownerBytes('controls', resources.cacheOwner), 22690)
    assert.equal(kernelOwners.size, 2)
    assert.ok([...kernelOwners].reduce((total, owner) => total + ledger.ownerBytes('controls', owner), 0) >= rawBytes)
      assert.ok(during.byPool.controls >= rawBytes + 22690, `Observed controls ${during.byPool.controls} B < raw sort lanes ${rawBytes} B + display lane 22690 B`)
      assert.ok(during.bytes - before.bytes >= rawBytes + 22690)
      assert.equal(privateDuring.length, 3)
      assert.ok(privateDuring.every(x => x.prepared.bytes === 0 && x.memory.limitBytes === (profile === 'D' ? 64 : 32) * 1024 * 1024))
      assert.ok(heldPages.length === 3)
      assert.equal(view, originalView)
      assert.equal(published, 0)
      for (const a of aborts)
        a.abort()
      controls.dispose()
      resources.dispose()
      await freshTask()
      await freshTask()
      const afterClose = ledger.snapshot(), liveAfterClose = allocations.filter(a => a.ref.deref() && a.bytes > 0 && (a.name === 'Uint8Array' && a.bytes === n || a.name === 'Float64Array' && a.bytes === n * 8)).reduce((s, a) => s + a.bytes, 0)
      assert.equal(liveAfterClose, rawBytes)
      // Disposal ends only the port's own explicit owners (AC-LIFE-02): its lookup
      // metadata (2048 B + 48 B per style, two styles) and its structural index
      // (1696 B). Held raw lanes and the display lane stay charged until drained.
      assert.equal(afterClose.byPool.controls, during.byPool.controls - (2048 + 2 * 48) - 1696)
      assert.equal(ledger.ownerBytes('controls', resources.cacheOwner), 22690)
      assert.ok([...kernelOwners].reduce((total, owner) => total + ledger.ownerBytes('controls', owner), 0) >= rawBytes)
      assert.equal(view, originalView)
      assert.equal(JSON.stringify(view.selection), originalSelection)
      assert.equal(JSON.stringify(doc), originalDoc)
      assert.deepEqual(store.stamp, stamp)
      for (const g of gates)
        g.open()
      const settled = await Promise.allSettled(pending), outcomes = settled.map(x => x.status === 'fulfilled' ? x.value : { kind: x.reason.code, rejected: true })
      assert.deepEqual(outcomes.map(x => x.kind), ['cancelled', 'cancelled', 'cancelled'])
    for (const owner of kernelOwners) assert.equal(ledger.ownerBytes('controls', owner), 0)
      assert.ok(afterClose.byPool.controls - ledger.snapshot().byPool.controls >= rawBytes + 22690)
      assert.ok(sessions.every(x => x.session.memoryStats().totalBytes === 0))
      assert.equal(view, originalView)
      assert.equal(published, 0)
      await source.release()
    }
    finally {
      Object.assign(globalThis, originals)
      Object.assign(DocumentSession.prototype, sessionOriginals)
      for (const g of gates)
        g.open()
    }
  }, 30000)
test('keeps membership and independent reader copies charged through close', async () => {
  const signal = () => new AbortController().signal, freshTask = () => new Promise<void>(r => setTimeout(r, 0))
  let trace = false, open!: () => void, hit!: () => void
  const wait = new Promise<void>(r => open = r), reached = new Promise<void>(r => hit = r)
  const originals = { Uint8Array, Float64Array, Uint32Array, Int32Array }
  const sessionOriginals = { sort: DocumentSession.prototype.sort, filter: DocumentSession.prototype.filter }
  try {
    const allocations: {
      name: string
      bytes: number
      ref: WeakRef<any>
      stack: string
    }[] = []
    for (const name of ['Uint8Array', 'Float64Array', 'Int32Array'] as const) {
      const real = globalThis[name];
      (globalThis as any)[name] = new Proxy(real, {
        construct(target, args) {
          const value = Reflect.construct(target, args, target), stack = new Error().stack ?? ''; if (trace && stack.includes('ValueKeySet') && stack.includes('compileValuesMembership'))
            allocations.push({ name, bytes: value.byteLength, ref: new WeakRef(value), stack }); return value
        }
      })
    }
    const n = 64, doc: any = { sheets: [{ name: 'S', sheetId: 's', workbookIndex: 0, rows: Array.from({ length: n + 1 }, (_, r) => ({ index: r, cells: [{ ref: '', row: r, col: 0, value: r === 0 ? 'Header' : r, styleIndex: 0 }] })), cols: [], merges: [], mergeRanges: [], semantics: semanticsFor('1900') }], tables: [] }
    const frozenSource = JSON.stringify(doc), source = createObjectWorkbookSource(doc, 'd'), store = source.store('s'), ledger = new MainMemoryLedger()
    ledger.chargeWhileReachable('document', [doc], estimatePayloadBytes(doc))
    const resources = createViewerControlResources('D', undefined, ledger)
    let gen = 0, published = 0
    const hidden = createIntervalSet([])
    let view: any = { sheetId: 's', revision: 0, order: createIdentityOrder(), baseHidden: hidden, visibility: visibilityFromHidden(hidden), sort: null, filter: null, selection: { ...emptyGridSelection(), active: { row: 2, col: 0 } } }
    const owner = () => ({ ...store.stamp, fontGen: 0, viewGen: gen }), snapshot = () => ({ owner: owner(), filter: view.filter, sort: view.sort, records: { visible: n + 1, total: n + 1, completeness: 'complete' as const } })
    const controls = await viewerControls({ doc, source, sheetId: 's', resources, ledger, snapshot, view: () => view, publish: next => { view = { ...view, ...next }; gen++; published++; return snapshot() } }, signal())
    const range: any = { r0: 0, r1: n, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 }
    const draft = await controls.client.values(owner(), range, 0, null, signal())
    const page = await draft.read({ text: '', offset: 0 }, signal())
    assert.equal(page.entries.length, n)
    await draft.change({ kind: 'all', selected: false, includeBlanks: false }, { text: '', offset: 0 }, signal())
    await draft.change({ kind: 'toggle', ordinal: 0, selected: true }, { text: '', offset: 0 }, signal())
    const rule: any = await draft.candidate(signal())
    assert.equal(rule.selection.keys.count, 1)
    const committed = await controls.client.filter(owner(), { range, columns: new Map([[0, rule]]) }, signal(), undefined, draft)
    assert.equal(committed.kind, 'committed')
    const beforeView = view, beforeSelection = JSON.stringify(view.selection), beforeOwner = owner(), before = ledger.snapshot(), cacheBefore = resources.cache.usedBytes
    let readRefs: WeakRef<any>[] = []
    const read = resources.runs.read.bind(resources.runs)
    resources.runs.read = async (...args: any[]) => { const run = await (read as any)(...args); readRefs = [new WeakRef(run.kinds), new WeakRef(run.numbers)]; hit(); await wait; return run }
    trace = true
    const abort = new AbortController(), pending = controls.client.filter(owner(), { range, columns: new Map([[0, { ...rule, showBlanks: !rule.showBlanks }]]) }, abort.signal)
    await reached
    await freshTask()
    await freshTask()
    const during = ledger.snapshot(), liveBytes = allocations.filter(a => a.ref.deref()).reduce((n, a) => n + a.bytes, 0), readBytes = readRefs.reduce((n, r) => n + (r.deref()?.byteLength ?? 0), 0)
    assert.equal(liveBytes, 1538)
    assert.equal(readBytes, 9)
    assert.ok(resources.cache.usedBytes - cacheBefore >= liveBytes)
    assert.ok(during.bytes - before.bytes >= liveBytes + readBytes, `Observed main increment ${during.bytes - before.bytes} B < membership ${liveBytes} B + run copy ${readBytes} B`)
    abort.abort()
    controls.dispose()
    resources.dispose()
    await freshTask()
    await freshTask()
    const afterDispose = ledger.snapshot(), heldMembership = allocations.filter(a => a.ref.deref()).reduce((n, a) => n + a.bytes, 0), heldRead = readRefs.reduce((n, r) => n + (r.deref()?.byteLength ?? 0), 0)
    assert.ok(afterDispose.byPool.controls >= heldMembership + heldRead)
    assert.equal(resources.cache.size, 0)
    assert.equal(heldMembership, 1538)
    assert.equal(heldRead, 9)
    const stale = await draft.read({ text: '', offset: 0 }, signal()).then(() => null, e => e.code)
    assert.equal(stale, 'stale-source')
    draft.cancel()
    assert.ok(resources.cache.usedBytes >= heldMembership)
    const afterDraftClose = ledger.snapshot()
    assert.ok(afterDraftClose.byPool.controls >= heldMembership + heldRead, `Observed controls ${afterDraftClose.byPool.controls} B < held membership ${heldMembership} B + run reader copy ${heldRead} B`)
    await freshTask()
    await freshTask()
    const heldAfterDraftClose = allocations.filter(a => a.ref.deref()).reduce((n, a) => n + a.bytes, 0) + readRefs.reduce((n, r) => n + (r.deref()?.byteLength ?? 0), 0)
    assert.equal(heldAfterDraftClose, 1547)
    open()
    const outcome = await pending.then(x => ({ status: 'fulfilled', kind: x.kind }), e => ({ status: 'rejected', kind: e.code }))
    assert.equal(outcome.kind, 'cancelled')
    assert.equal(view, beforeView)
    assert.deepEqual(owner(), beforeOwner)
    assert.equal(JSON.stringify(view.selection), beforeSelection)
    assert.equal(JSON.stringify(doc), frozenSource)
    assert.equal(published, 1)
    await source.release()
  }
  finally {
    Object.assign(globalThis, originals)
    Object.assign(DocumentSession.prototype, sessionOriginals)
    open?.()
  }
}, 30000)
test('canonical formatted text refuses before constructing its output', async () => {
  const { cellDisplayText } = await import('../src/xlsx/render')
  const sheet = { semantics: semanticsFor('1900') } as any
  const call = cellDisplayText as (...args: any[]) => string
  assert.throws(() => call(sheet, { value: 12, style: { numFmtId: 0, formatCode: '"' + 'z'.repeat(3000) + '"0' } }, {
    allocate() { throw new Error('allocation refused') }, release() { },
  }), /allocation refused/)
})
test('resource observers remain called while the authoritative cache has one owner', async () => {
  const ledger = new MainMemoryLedger(), calls: number[] = []
  const resources = createViewerControlResources('D', bytes => calls.push(bytes), ledger)
  const reserve = resources.cache.reserve()
  reserve.charge(77)
  assert.equal(ledger.ownerBytes('controls', resources.cacheOwner), 77)
  assert.equal(ledger.snapshot().byPool.controls, 1024 + 77)
  assert.deepEqual(calls, [77])
  resources.dispose()
  assert.equal(ledger.ownerBytes('controls', resources.cacheOwner), 77)
  reserve.close()
  assert.equal(ledger.ownerBytes('controls', resources.cacheOwner), 0)
  assert.deepEqual(calls, [77, 0])
})
for (const profile of ['D', 'M'] as const)
  test(`real ${profile} source output refuses before typed constructors against the complete main aggregate`, async () => {
    const { controlSourceStore, MainControlMemory } = await import('../src/xlsx/view/mainControlMemory')
    const { MAIN_THREAD_BUDGETS } = await import('../src/xlsx/view/tileCache')
    const doc: any = {
      sheets: [{
        name: 'S', sheetId: 's', workbookIndex: 0,
        rows: [{ index: 0, cells: [{ row: 0, col: 0, ref: 'A1', value: 3, styleIndex: 0 }] }],
        cols: [], merges: [], mergeRanges: [], semantics: semanticsFor('1900')
      }], tables: []
    }
    const source = createObjectWorkbookSource(doc, 'd'), original = source.store('s')
    const ledger = new MainMemoryLedger()
    ledger.chargeWhileReachable('document', [doc], estimatePayloadBytes(doc))
    const memory = new MainControlMemory(ledger, profile), store = controlSourceStore(original, memory)
    const copyBytes = MAIN_THREAD_BUDGETS[profile] - ledger.bytes - 1672
    ledger.set('copy', 'actual-copy', copyBytes)
    let heldCopy: Uint8Array | undefined = new Uint8Array(copyBytes)
    const native = globalThis.Uint16Array
    let copies = 0
    globalThis.Uint16Array = new Proxy(native, { construct(target, args) { copies++; return Reflect.construct(target, args, target) } })
    const input = JSON.stringify(doc), request = { col: 0, rows: { kind: 'intervals' as const, intervals: [{ r0: 0, r1: 0 }] }, limit: 1, encoding: 'lanes' as const }
    try {
      await assert.rejects(store.queryColumn(request, { stamp: store.stamp, signal: new AbortController().signal }), (error: any) => error.code === 'quota')
      assert.equal(copies, 0)
      assert.equal(heldCopy.byteLength, copyBytes)
      assert.equal(JSON.stringify(doc), input)
      assert.ok(ledger.peakBytes <= MAIN_THREAD_BUDGETS[profile])
      heldCopy = undefined
      ledger.release('copy', 'actual-copy')
      ledger.set('copy', 'actual-copy', copyBytes - 4096)
      heldCopy = new Uint8Array(copyBytes - 4096)
      const result = await store.queryColumn(request, { stamp: store.stamp, signal: new AbortController().signal })
      assert.equal(result.items[0].kind, 'lanes')
      assert.equal(copies, 1)
      assert.ok(ledger.peakBytes <= MAIN_THREAD_BUDGETS[profile])
    }
    finally {
      globalThis.Uint16Array = native
      ledger.release('copy', 'actual-copy')
      await source.release()
    }
  })
test('a local session pool admits and drains distinct retained and transient owners', async () => {
  const { MainControlMemory } = await import('../src/xlsx/view/mainControlMemory')
  const { MemoryPool } = await import('../src/xlsx/view/tileCache')
  const ledger = new MainMemoryLedger(), memory = new MainControlMemory(ledger)
  const pool = new MemoryPool({ limitBytes: memory.limitBytes, allocation: memory } as any)
  const entry = pool.retain('new-lanes', { tier: 'rebuildable', bytes: 77 })!
  const transient = pool.charge()
  assert.equal(transient.set(11), true)
  assert.equal(ledger.snapshot().byPool.controls, 88)
  pool.close()
  assert.equal(entry.live, false)
  assert.equal(ledger.snapshot().byPool.controls, 11)
  transient.release()
  assert.equal(ledger.snapshot().byPool.controls, 0)
})
test('membership admits its two-byte boolean lane before its constructor', async () => {
  const { ValueKeySet } = await import('../src/xlsx/view/valueRuns')
  let calls = 0
  const native = globalThis.Uint8Array
  globalThis.Uint8Array = new Proxy(native, {
    construct(target, args) {
      if (args[0] === 2)
        calls++; return Reflect.construct(target, args, target)
    }
  })
  try {
    assert.throws(() => new ValueKeySet({ charge(bytes) { if (bytes === 2) throw new Error('refused') }, release() { } }), /refused/)
    assert.equal(calls, 0)
  }
  finally {
    globalThis.Uint8Array = native
  }
})
test('released membership revokes an escaped matcher before refunding its buffers', async () => {
  const { compileValuesMembership, ValueIndexCache, valueScope } = await import('../src/xlsx/view/valueIndex')
  const { MemoryValueRunStore } = await import('../src/xlsx/view/valueRuns')
  const identity: any = {
    docId: 'd', sheetId: 's', valueGen: 0, col: 0, classificationGen: 0,
    classification: { kind: 'number', hasNumeric: true }, range: { r0: 0, r1: 1, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }
  }
  const cache = new ValueIndexCache(), store = new MemoryValueRunStore()
  const ref = await store.put({ count: 1, kinds: new Uint8Array([1]), numbers: new Float64Array([1]), texts: [] }, identity, valueScope(identity))
  const membership = await compileValuesMembership({ kind: 'values', selection: { mode: 'include', keys: ref }, showBlanks: false }, identity, store, { cache })
  const matches = membership.matches
  assert.equal(matches({ kind: 'number', value: 1 }), true)
  membership.release()
  assert.equal(cache.usedBytes, 0)
  assert.throws(() => matches({ kind: 'number', value: 1 }), (error: any) => error.code === 'closed')
  await store.release(ref)
})
test('an actual run reader refuses before either slice and holds escaped buffers after store release', async () => {
  const { MainControlMemory } = await import('../src/xlsx/view/mainControlMemory')
  const { MemoryValueRunStore } = await import('../src/xlsx/view/valueRuns')
  const ledger = new MainMemoryLedger(), memory = new MainControlMemory(ledger), storedAccount = memory.reserve()
  const run = { count: 1, kinds: new Uint8Array([1]), numbers: new Float64Array([12]), texts: [] }
  ledger.chargeWhileReachable('document', [run], estimatePayloadBytes(run))
  const store = new MemoryValueRunStore({ account: storedAccount, readerAllocation: memory })
  const owner = { docId: 'd', sheetId: 's', valueGen: 0 }
  const ref = await store.put(run, owner, 'scope')
  const copyBytes = memory.limitBytes - ledger.bytes - 112
  ledger.set('copy', 'held', copyBytes)
  let held: Uint8Array | undefined = new Uint8Array(copyBytes)
  const slice8 = Uint8Array.prototype.slice, slice64 = Float64Array.prototype.slice
  let slices = 0
  Uint8Array.prototype.slice = function(...args) { slices++; return slice8.apply(this, args) }
  Float64Array.prototype.slice = function(...args) { slices++; return slice64.apply(this, args) }
  try {
    await assert.rejects(store.read(ref, 0, 1, 'scope'), (error: any) => error.code === 'quota')
    assert.equal(slices, 0)
    assert.equal(held.byteLength, copyBytes)
    assert.equal(store.refCount, 1)
    held = undefined
    ledger.release('copy', 'held')
    const page = await store.read(ref, 0, 1, 'scope')
    assert.equal(slices, 2)
    const escaped = page.numbers.buffer
    await store.release(ref)
    assert.equal(storedAccount.bytes, 0)
    assert.equal(escaped.byteLength, 8)
    assert.ok(ledger.snapshot().byPool.controls >= 785)
    assert.equal(memory.pendingBytes, 0)
  }
  finally {
    Uint8Array.prototype.slice = slice8
    Float64Array.prototype.slice = slice64
    held = undefined
    ledger.release('copy', 'held')
    await store.release(ref)
  }
})
function localFixture(profile: 'D' | 'M' = 'D', count = 16) {
  const doc: any = {
    sheets: [{
      name: 'S', sheetId: 's', workbookIndex: 0,
      rows: Array.from({ length: count + 1 }, (_, row) => ({ index: row, cells: [{ row, col: 0, ref: '', value: row ? count - row : 'Header', styleIndex: 0 }] })),
      cols: [], merges: [], mergeRanges: [], semantics: semanticsFor('1900')
    }], tables: []
  }
  const source = createObjectWorkbookSource(doc, 'd'), store = source.store('s'), ledger = new MainMemoryLedger()
  ledger.chargeWhileReachable('document', [doc], estimatePayloadBytes(doc))
  const hidden = createIntervalSet([])
  let gen = 0, published = 0
  let view: any = { sheetId: 's', revision: 0, order: createIdentityOrder(), baseHidden: hidden, visibility: visibilityFromHidden(hidden), sort: null, filter: null, selection: { ...emptyGridSelection(), active: { row: 3, col: 0 } } }
  ledger.chargeWhileReachable('viewBuffers', [view.visibility, view.visibility.words, view.visibility.words.buffer, view.visibility.summaries, view.visibility.summaries.buffer], view.visibility.words.byteLength + view.visibility.summaries.byteLength)
  const owner = () => ({ ...store.stamp, fontGen: 0, viewGen: gen })
  const snapshot = () => ({ owner: owner(), sort: view.sort, filter: view.filter, records: { visible: count + 1, total: count + 1, completeness: 'complete' as const } })
  const open = (resources = createViewerControlResources(profile, undefined, ledger)) => viewerControls({ doc, source, sheetId: 's', resources, ledger, snapshot, view: () => view, publish: next => { view = { ...view, ...next }; gen++; published++; return snapshot() } }, new AbortController().signal)
  return {
    doc, source, store, ledger, owner, snapshot, open, view: () => view, published: () => published,
    range: { r0: 0, r1: count, c0: 0, c1: 0, headerRows: 1 as const, totalsRows: 0 as const }
  }
}
for (const profile of ['D', 'M'] as const)
  test(`20 actual ${profile} cancel/reopen/close cycles drain explicit owners without eager header columns`, async () => {
    const f = localFixture(profile), input = JSON.stringify(f.doc), initial = f.view(), stamp = { ...f.store.stamp }
    const query = f.store.queryColumn.bind(f.store)
    let queries = 0, structures = 0
    const structural = f.store.queryStructural.bind(f.store)
    f.store.queryStructural = async (...args) => { structures++; return structural(...args) }
    for (let cycle = 0;cycle < 20;cycle++) {
      const resources = createViewerControlResources(profile, undefined, f.ledger), memory = resources.memory
      const controls = await f.open(resources)
      assert.equal(queries, cycle)
      assert.equal(resources.cache.size, 0)
      assert.equal(resources.cache.usedBytes, 0)
      assert.equal(structures, cycle * 2 + 1)
      let hit!: () => void, open!: () => void
      const reached = new Promise<void>(resolve => hit = resolve), wait = new Promise<void>(resolve => open = resolve)
      f.store.queryColumn = async (...args) => { const page = await query(...args); queries++; hit(); await wait; return page }
      const abort = new AbortController(), pending = controls.client.sort(f.owner(), { range: f.range, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }, abort.signal)
      await reached
      const held = memory.pendingBytes
      assert.ok(held >= 16 * 9)
      abort.abort()
      controls.dispose()
      resources.dispose()
      assert.equal(memory.pendingBytes, held)
      open()
      const result = await pending
      assert.equal(result.kind, 'cancelled')
      assert.equal(memory.pendingBytes, 0)
      assert.equal(f.ledger.ownerBytes('controls', resources.cacheOwner), 0)
      assert.equal(f.view(), initial)
      assert.equal(f.published(), 0)
      assert.equal(JSON.stringify(f.doc), input)
      assert.deepEqual(f.store.stamp, stamp)
      assert.ok(f.ledger.peakBytes <= memory.limitBytes)
      assert.equal((await controls.client.sort(f.owner(), { range: f.range, keys: [{ col: 0, order: 'desc' }], caseSensitive: false }, new AbortController().signal)).kind, 'cancelled')
    }
    await f.source.release()
  }, 30000)
for (const profile of ['D', 'M'] as const)
  test(`actual ${profile} kernel refuses before raw lanes when another main pool consumes the final byte`, async () => {
    const { sortRowOrder } = await import('../src/xlsx/view/sort')
    const { MainControlMemory, controlSourceStore } = await import('../src/xlsx/view/mainControlMemory')
    const f = localFixture(profile, 64), memory = new MainControlMemory(f.ledger, profile)
    const needed = 1024 + 128 + 1024 + 64 * 9 + 512
    const copyBytes = memory.limitBytes - f.ledger.bytes - needed + 1
    f.ledger.set('copy', 'held', copyBytes)
    let held: Uint8Array | undefined = new Uint8Array(copyBytes)
    const native = globalThis.Uint8Array
    let raw = 0
    globalThis.Uint8Array = new Proxy(native, {
      construct(target, args) {
        if (args[0] === 64)
          raw++; return Reflect.construct(target, args, target)
      }
    })
    try {
      const result = await sortRowOrder(controlSourceStore(f.store, memory), { range: f.range, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }, { allocation: memory, maxWorkerBytes: memory.limitBytes })
      assert.equal(result.kind, 'resource-limit')
      assert.equal(raw, 0)
      assert.equal(memory.pendingBytes, 0)
      assert.equal(held.byteLength, copyBytes)
      assert.ok(f.ledger.peakBytes <= memory.limitBytes)
    }
    finally {
      globalThis.Uint8Array = native
      held = undefined
      f.ledger.release('copy', 'held')
      await f.source.release()
    }
  })
test('legacy observer binding is explicit and preserves notifications and one authoritative owner', () => {
  const calls: number[] = [], resources = createViewerControlResources('D', bytes => calls.push(bytes)), ledger = new MainMemoryLedger()
  assert.throws(() => resources.bindLedger(ledger), /resourceUsage/)
  resources.bindLedger(ledger, 'observer')
  const account = resources.cache.reserve()
  account.charge(17)
  assert.deepEqual(calls, [17])
  assert.equal(ledger.ownerBytes('controls', resources.cacheOwner), 17)
  account.close()
  assert.deepEqual(calls, [17, 0])
  assert.equal(ledger.ownerBytes('controls', resources.cacheOwner), 0)
  resources.dispose()
})
test('adoption transfers a reserved owner without a trough or a doubled peak', async () => {
  const { MainControlMemory } = await import('../src/xlsx/view/mainControlMemory')
  const registered: string[] = []
  const ledger = new MainMemoryLedger({ register(_target, token) { registered.push(token) } }), memory = new MainControlMemory(ledger)
  const account = memory.reserve()
  account.charge(1000)
  const target = { buffer: new Uint8Array(600) }
  account.adopt([target, target.buffer, target.buffer.buffer], 600)
  assert.equal(account.bytes, 400)
  assert.equal(ledger.bytes, 1000)
  assert.equal(ledger.peakBytes, 1000)
  account.release()
  assert.equal(ledger.bytes, 600)
  ledger.collected(registered[0])
  ledger.collected(registered[1])
  assert.equal(ledger.bytes, 600)
  ledger.collected(registered[2])
  assert.equal(ledger.bytes, 0)
  assert.equal(memory.pendingBytes, 0)
})
test('an imported replacement charges only its newly distinct backing store', async () => {
  const { MainControlMemory, controlView } = await import('../src/xlsx/view/mainControlMemory')
  const { createRowOrder } = await import('../src/xlsx/view/order')
  const memory = new MainControlMemory(new MainMemoryLedger())
  const order = createRowOrder([{ r0: 0, r1: 1, sourceAtOffset: new Uint32Array([1, 0]) }])
  const visibility = visibilityFromHidden(createIntervalSet([]))
  controlView(memory, order, visibility)
  const before = memory.ledger.bytes
  const replacement: any = { ...order, bands: [{ ...order.bands[0], displayOffsetOfSource: new Uint32Array([1, 0]) }] }
  controlView(memory, replacement, visibility)
  assert.equal(memory.ledger.bytes - before, 128 + 8)
  controlView(memory, replacement, visibility)
  assert.equal(memory.ledger.bytes - before, 128 + 8)
  assert.equal(memory.pendingBytes, 0)
})
test('a real escaped reader backing store remains charged until its final collection', async () => {
  const v8 = await import('node:v8'), vm = await import('node:vm')
  v8.setFlagsFromString('--expose_gc')
  const gc = vm.runInNewContext('gc') as () => void
  const collect = async () => {
    for (let i = 0;i < 8;i++) {
      gc()
      await new Promise(resolve => setTimeout(resolve, 10))
    }
  }
  const { MainControlMemory } = await import('../src/xlsx/view/mainControlMemory')
  const { MemoryValueRunStore } = await import('../src/xlsx/view/valueRuns')
  const ledger = new MainMemoryLedger(), memory = new MainControlMemory(ledger), account = memory.reserve()
  const store = new MemoryValueRunStore({ account, readerAllocation: memory })
  let escaped: ArrayBufferLike | undefined
  await (async () => {
    const ref = await store.put({ count: 1, kinds: new Uint8Array([1]), numbers: new Float64Array([12]), texts: [] }, { docId: 'd', sheetId: 's', valueGen: 0 }, 'scope')
    const page = await store.read(ref, 0, 1, 'scope')
    escaped = page.numbers.buffer
    await store.release(ref)
  })()
  const ref = new WeakRef(escaped!)
  await collect()
  assert.ok(ref.deref())
  assert.equal(ledger.snapshot().byPool.controls, 785)
  escaped = undefined
  await collect()
  assert.equal(ref.deref(), undefined)
  assert.equal(ledger.snapshot().byPool.controls, 0)
}, 30000)
test('closing a port aborts its actual held producer without requiring the caller to abort', async () => {
  const f = localFixture(), resources = createViewerControlResources('D', undefined, f.ledger), controls = await f.open(resources)
  const query = f.store.queryColumn.bind(f.store)
  let hit!: () => void, open!: () => void, producerSignal: AbortSignal | undefined
  const reached = new Promise<void>(resolve => hit = resolve), wait = new Promise<void>(resolve => open = resolve)
  f.store.queryColumn = async (...args) => { const page = await query(...args); producerSignal = args[1].signal; hit(); await wait; return page }
  const caller = new AbortController(), initial = f.view()
  const pending = controls.client.sort(f.owner(), { range: f.range, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }, caller.signal)
  try {
    await reached
    controls.dispose()
    assert.equal(caller.signal.aborted, false)
    assert.equal(producerSignal!.aborted, true)
    open()
    assert.equal((await pending).kind, 'cancelled')
    assert.equal(f.view(), initial)
    assert.equal(f.published(), 0)
  }
  finally {
    open()
    await pending
    resources.dispose()
    await f.source.release()
  }
})
test('a real values search paused at its pacing yield keeps its bitmap charged through cache close', async () => {
  const { ValueIndex } = await import('../src/xlsx/view/valueIndex')
  const f = localFixture('D', 512), resources = createViewerControlResources('D', undefined, f.ledger), controls = await f.open(resources)
  const draft = await controls.client.values(f.owner(), f.range, 0, null, new AbortController().signal)
  const search = ValueIndex.prototype.search, native = globalThis.Uint32Array
  let hit!: () => void, open!: () => void, clock = 0, bitmap: WeakRef<Uint32Array> | undefined
  const reached = new Promise<void>(resolve => hit = resolve), wait = new Promise<void>(resolve => open = resolve)
  ValueIndex.prototype.search = function(query, options) { return search.call(this, query, { ...options, now: () => clock += 10, yield: async () => { hit(); await wait } }) }
  globalThis.Uint32Array = new Proxy(native, {
    construct(target, args) {
      const value = Reflect.construct(target, args, target)
      if (args[0] === 16 && new Error().stack?.includes('ValueIndex.search'))
        bitmap = new WeakRef(value)
      return value
    }
  })
  const input = JSON.stringify(f.doc), initial = f.view()
  const pending = draft.read({ text: '1', offset: 0 }, new AbortController().signal)
  try {
    await reached
    assert.equal(bitmap!.deref()!.byteLength, 64)
    resources.dispose()
    assert.equal(resources.cache.size, 0)
    assert.ok(resources.cache.usedBytes >= bitmap!.deref()!.byteLength, 'Cache close refunded a bitmap still held by the actual search continuation')
    draft.cancel()
    open()
    await assert.rejects(pending)
    assert.equal(resources.cache.usedBytes, 0)
    assert.equal(resources.memory.pendingBytes, 0)
    assert.equal(JSON.stringify(f.doc), input)
    assert.equal(f.view(), initial)
    assert.equal(f.published(), 0)
  }
  finally {
    open()
    await pending.catch(() => { })
    ValueIndex.prototype.search = search
    globalThis.Uint32Array = native
    draft.cancel()
    controls.dispose()
    resources.dispose()
    await f.source.release()
  }
})
for (const profile of ['D', 'M'] as const)
  test(`eight actual local ${profile} raw keys preserve sort semantics and the main limit`, async () => {
    const f = localFixture(profile), resources = createViewerControlResources(profile, undefined, f.ledger), controls = await f.open(resources)
    const input = JSON.stringify(f.doc), selected = JSON.stringify(f.view().selection)
    const result = await controls.client.sort(f.owner(), { range: { ...f.range, c1: 7 }, keys: Array.from({ length: 8 }, (_, col) => ({ col, order: 'asc' as const })), caseSensitive: false }, new AbortController().signal)
    assert.equal(result.kind, 'committed')
    assert.equal(f.view().order.displayToSource(1), 16)
    assert.equal(JSON.stringify(f.doc), input)
    assert.equal(JSON.stringify(f.view().selection), selected)
    assert.equal(resources.memory.pendingBytes, 0)
    assert.ok(f.ledger.peakBytes <= resources.memory.limitBytes)
    controls.dispose()
    resources.dispose()
    await f.source.release()
  })
test('one main ledger cannot acquire a second profile allowance', () => {
  const ledger = new MainMemoryLedger(), mobile = createViewerControlResources('M', undefined, ledger)
  assert.throws(() => createViewerControlResources('D', undefined, ledger), /profile/)
  mobile.dispose()
})
test('actual raw text sort admits its decode builders and argument scratch before String.fromCharCode', async () => {
  const text = 'x'.repeat(10000)
  const doc: any = {
    sheets: [{
      name: 'S', sheetId: 's', workbookIndex: 0, rows: [
        { index: 0, cells: [{ row: 0, col: 0, ref: '', value: 'Header', styleIndex: 0 }] },
        { index: 1, cells: [{ row: 1, col: 0, ref: '', value: text, styleIndex: 0 }] },
      ], cols: [], merges: [], mergeRanges: [], semantics: semanticsFor('1900')
    }], tables: []
  }
  const source = createObjectWorkbookSource(doc, 'd'), store = source.store('s'), ledger = new MainMemoryLedger()
  ledger.chargeWhileReachable('document', [doc], estimatePayloadBytes(doc))
  const resources = createViewerControlResources('D', undefined, ledger), hidden = createIntervalSet([])
  let gen = 0, view: any = { sheetId: 's', revision: 0, order: createIdentityOrder(), baseHidden: hidden, visibility: visibilityFromHidden(hidden), sort: null, filter: null, selection: emptyGridSelection() }
  const owner = () => ({ ...store.stamp, fontGen: 0, viewGen: gen }), snapshot = () => ({ owner: owner(), filter: view.filter, sort: view.sort, records: { visible: 2, total: 2, completeness: 'complete' as const } })
  const controls = await viewerControls({ doc, source, sheetId: 's', resources, ledger, snapshot, view: () => view, publish: next => { view = { ...view, ...next }; gen++; return snapshot() } }, new AbortController().signal)
  const canonical = String.fromCharCode
  let held = 0
  String.fromCharCode = (...args) => {
    if (!held && args.length === 256 && new Error().stack?.includes('/xlsx/view/sort.ts'))
      held = resources.memory.pendingBytes; return canonical(...args)
  }
  try {
    const result = await controls.client.sort(owner(), { range: { r0: 0, r1: 1, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }, new AbortController().signal)
    assert.equal(result.kind, 'committed')
    assert.ok(held >= text.length * 6 + 256 * 16 + 96, `Decode had only ${held} B admitted before its actual builders/arguments`)
  }
  finally {
    String.fromCharCode = canonical
    controls.dispose()
    resources.dispose()
    await source.release()
  }
})
test('direct local main sessions do not retarget or release their borrowed source', async () => {
  const { MainControlMemory } = await import('../src/xlsx/view/mainControlMemory')
  const f = localFixture(), memory = new MainControlMemory(f.ledger)
  let bindings = 0, releases = 0
  const source = { ...f.source, bindMemory() { bindings++ }, async release() { releases++ } }
  const session = new DocumentSession('d', source, { allocation: memory })
  assert.equal(session.memoryStats().limitBytes, memory.limitBytes)
  await session.release()
  assert.equal(bindings, 0)
  assert.equal(releases, 0)
  const page = await f.store.queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 1, r1: 1 }] }, limit: 1 }, { stamp: f.store.stamp, signal: new AbortController().signal })
  assert.equal(page.completeness, 'complete')
  await f.source.release()
})

test('a source scan releases traversal scratch before yielding an escapable page', async () => {
  const { MainControlMemory, controlSourceStore } = await import('../src/xlsx/view/mainControlMemory')
  const f = localFixture(), memory = new MainControlMemory(f.ledger), store = controlSourceStore(f.store, memory)
  const iterator = store.scanPopulated({ range: f.range, limit: 1 }, { stamp: store.stamp, signal: new AbortController().signal })[Symbol.asyncIterator]()
  const first = await iterator.next()
  assert.equal(first.done, false)
  assert.equal(memory.pendingBytes, 0)
  assert.ok(f.ledger.snapshot().byPool.controls > 0)
  await iterator.return?.()
  await f.source.release()
})

test('an actual source IO failure drains explicit main preparation and preserves source/view/selection', async () => {
  const f = localFixture(), resources = createViewerControlResources('D', undefined, f.ledger), controls = await f.open(resources)
  const before = f.view(), sourceJSON = JSON.stringify(f.doc), selected = JSON.stringify(before.selection)
  const original = f.store.queryColumn.bind(f.store)
  f.store.queryColumn = async (...args) => { await original(...args); throw new Error('actual producer IO failure') }
  try {
    const result = await controls.client.sort(f.owner(), { range: f.range, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }, new AbortController().signal)
    assert.equal(result.kind, 'refused')
    if (result.kind === 'refused') assert.equal(result.refusal.reason, 'source-error')
    assert.equal(resources.memory.pendingBytes, 0)
    assert.equal(f.view(), before)
    assert.equal(f.published(), 0)
    assert.equal(JSON.stringify(f.doc), sourceJSON)
    assert.equal(JSON.stringify(f.view().selection), selected)
  } finally { controls.dispose(); resources.dispose(); await f.source.release() }
})
