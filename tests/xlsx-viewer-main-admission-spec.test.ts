import { expect, test } from 'vitest'
import { DocumentSession } from '../src/worker/viewerSession'
import { MainControlMemory } from '../src/xlsx/view/mainControlMemory'
import { MainMemoryLedger, estimatePayloadBytes } from '../src/xlsx/view/mainLedger'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { classifyColumn } from '../src/xlsx/view/filter'
import { semanticsFor } from '../src/xlsx/view/semanticsFor'
import { createViewerControlResources, viewerControls } from '../src/components/xlsx/viewerControlsClient'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { emptyGridSelection } from '../src/xlsx/view/selection'
import { ValueIndex, ValueIndexCache, valueScope, valueSourceFromStore } from '../src/xlsx/view/valueIndex'

function workbook() {
  return { sheets: ['s', 'other'].map(sheetId => ({ name: sheetId, sheetId, workbookIndex: sheetId === 's' ? 0 : 1,
    rows: Array.from({ length: 5 }, (_, row) => ({ index: row, cells: [{ row, col: 0, ref: '', value: row ? 5 - row : 'Header', styleIndex: 0 }] })),
    cols: [], merges: [], mergeRanges: [], semantics: semanticsFor('1900') })), tables: [] } as any
}
const signal = () => new AbortController().signal
const range = { r0: 0, r1: 4, c0: 0, c1: 0, headerRows: 1 as const, totalsRows: 0 as const }
function pressure(memory: MainControlMemory, headroom = 0) {
  // Synthetic aggregate pressure backed by a held allocation, not a clipboard/UI operation.
  const bytes = memory.limitBytes - memory.ledger.bytes - headroom
  memory.ledger.set('copy', 'synthetic-held-buffer', bytes)
  const buffer = new Uint8Array(bytes)
  return { buffer, release() { memory.ledger.release('copy', 'synthetic-held-buffer') } }
}

for (const profile of ['D', 'M'] as const) for (const route of ['direct', 'identity-rebuild', 'sort-rebuild'] as const) {
  test(`${profile} ${route} refuses before visibility constructors and slices, then admits the unchanged view`, async () => {
    const doc = workbook(), originalDoc = JSON.stringify(doc), source = createObjectWorkbookSource(doc, 'd')
    const ledger = new MainMemoryLedger(), memory = new MainControlMemory(ledger, profile)
    ledger.chargeWhileReachable('document', [doc], estimatePayloadBytes(doc))
    for (const id of ['s', 'other']) { const geometry = source.geometry(id); ledger.chargeWhileReachable('geometry', [geometry], estimatePayloadBytes(geometry)) }
    const session = new DocumentSession('d', source, { allocation: memory, profile })
    if (route !== 'direct') {
      if (route === 'sort-rebuild') {
        const result = await session.sort(session.identity('s'), { range, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }, signal())
        expect(result.kind).toBe('committed')
      } else session.view('s')
      session.activate('other')
      session.memory.trim(0)
      expect(session.isResident('s')).toBe(false)
    }
    const identity = session.identity('s'), gen = session.globalViewGen, held = pressure(memory), before = ledger.bytes
    const Native = Uint32Array, slice = Native.prototype.slice
    let constructors = 0, slices = 0
    globalThis.Uint32Array = new Proxy(Native, { construct(target, args) {
      if ((new Error().stack ?? '').includes('visibility.ts')) constructors++
      return Reflect.construct(target, args, target)
    } })
    Native.prototype.slice = function (...args) { if ((new Error().stack ?? '').includes('visibility.ts')) slices++; return slice.apply(this, args) }
    try {
      if (route === 'direct') {
        let error: any; try { session.view('s') } catch (cause) { error = cause }
        expect(error?.code).toBe('quota')
      } else expect((await session.ensureView('s', signal())).kind).toBe('quota')
      expect(constructors + slices).toBe(0)
      expect(session.isResident('s')).toBe(false)
      expect(session.identity('s')).toEqual(identity)
      expect(session.globalViewGen).toBe(gen)
      expect(ledger.bytes).toBe(before)
      expect(memory.pendingBytes).toBe(0)
    } finally { globalThis.Uint32Array = Native; Native.prototype.slice = slice; held.release() }
    const next = route === 'direct' ? session.view('s') : await session.ensureView('s', signal())
    expect(route === 'direct' ? (next as any).visibility.logicalCount : (next as any).kind).toBe(route === 'direct' ? 1048576 : 'resident')
    expect(session.identity('s')).toEqual(identity)
    expect(session.globalViewGen).toBe(gen)
    const view = session.view('s'), bytes = ledger.bytes
    expect(memory.owns(view.visibility)).toBe(true)
    expect(memory.owns(view.visibility.words.buffer)).toBe(true)
    expect(memory.owns(view.visibility.summaries.buffer)).toBe(true)
    expect(session.view('s').visibility).toBe(view.visibility)
    expect(ledger.bytes).toBe(bytes)
    expect(ledger.peakBytes).toBeLessThanOrEqual(memory.limitBytes)
    expect(memory.pendingBytes).toBe(0)
    expect(JSON.stringify(doc)).toBe(originalDoc)
    await session.release(); await source.release()
  })
}

test('default worker mask construction and 512/256 MiB policies remain unchanged', async () => {
  for (const profile of ['D', 'M'] as const) {
    const source = createObjectWorkbookSource(workbook(), 'd'), session = new DocumentSession('d', source, { profile })
    expect(session.view('s').visibility.logicalCount).toBe(1048576)
    expect(session.memoryStats().limitBytes).toBe((profile === 'D' ? 512 : 256) * 1024 * 1024)
    await session.release()
  }
})

async function actualValues(profile: 'D' | 'M') {
  const doc = workbook(), source = createObjectWorkbookSource(doc, 'd'), store = source.store('s'), ledger = new MainMemoryLedger()
  ledger.chargeWhileReachable('document', [doc], estimatePayloadBytes(doc))
  const resources = createViewerControlResources(profile, undefined, ledger), hidden = createIntervalSet([])
  const view: any = { sheetId: 's', revision: 0, order: createIdentityOrder(), baseHidden: hidden, visibility: visibilityFromHidden(hidden), sort: null, filter: null, selection: emptyGridSelection() }
  const owner = () => ({ ...store.stamp, fontGen: 0, viewGen: 0 })
  const snapshot = () => ({ owner: owner(), sort: null, filter: null, records: { visible: 5, total: 5, completeness: 'complete' as const } })
  const controls = await viewerControls({ doc, source, sheetId: 's', resources, ledger, snapshot, view: () => view, publish: () => null as any }, signal())
  return { doc, source, resources, ledger, controls, owner, view }
}
for (const profile of ['D', 'M'] as const) test(`${profile} actual warm-classification values reopen refuses before Boolean construction at main limit`, async () => {
  const { doc, source, resources, ledger, controls, owner, view } = await actualValues(profile)
  const originalDoc = JSON.stringify(doc), selection = JSON.stringify(view.selection)
  const first = await controls.client.values(owner(), range, 0, null, signal())
  first.cancel(); resources.cache.invalidate()
  expect(resources.cache.usedBytes).toBe(0)
  const held = pressure(resources.memory), before = ledger.bytes, Native = Int32Array
  let boolBytes = 0
  globalThis.Int32Array = new Proxy(Native, { construct(target, args) {
    const out = Reflect.construct(target, args, target)
    if ((new Error().stack ?? '').includes('new ValueIndex')) boolBytes += out.byteLength
    return out
  } })
  try {
    const result = await controls.client.values(owner(), range, 0, null, signal()).then(draft => { draft.cancel(); return undefined }, e => e)
    expect(boolBytes).toBe(0)
    expect(result?.code).toBe('quota')
    expect(resources.cache.size).toBe(0)
    expect(resources.cache.usedBytes).toBe(0)
    expect(ledger.bytes).toBe(before)
  } finally { globalThis.Int32Array = Native; held.release() }
  for (let i = 0; i < 5; i++) {
    const draft = await controls.client.values(owner(), range, 0, null, signal())
    expect(resources.cache.size).toBe(1)
    expect(resources.cache.usedBytes).toBeGreaterThanOrEqual(8)
    expect(ledger.ownerBytes('controls', resources.cacheOwner)).toBe(resources.cache.usedBytes)
    draft.cancel(); resources.cache.invalidate()
    expect(resources.cache.size).toBe(0); expect(resources.cache.usedBytes).toBe(0)
  }
  expect(JSON.stringify(doc)).toBe(originalDoc); expect(JSON.stringify(view.selection)).toBe(selection)
  controls.dispose(); resources.dispose(); expect(resources.cache.usedBytes).toBe(0); await source.release()
})

function domain() {
  const source = createObjectWorkbookSource(workbook(), 'd'), store = source.store('s')
  return { source, identity: { ...store.stamp, range, col: 0, classification: { kind: 'number' as const, hasNumeric: true }, classificationGen: 0 }, values: valueSourceFromStore(store, 0, range) }
}
test('new index construction is admitted by cache before Boolean/container publication and rolls back constructor faults', async () => {
  const { source, identity, values } = domain(), observer: number[] = []
  const cache = new ValueIndexCache({ onUsage: bytes => observer.push(bytes) })
  const index = cache.open(identity, values), initial = cache.usedBytes
  expect(initial).toBeGreaterThan(8)
  expect(index.ledger.chargedBytes).toBe(initial)
  expect(cache.open(identity, values)).toBe(index)
  expect(cache.usedBytes).toBe(initial)
  index.dispose(); expect(cache.usedBytes).toBe(0)
  const tiny = new ValueIndexCache({ limitBytes: initial - 1 }), exact = new ValueIndexCache({ limitBytes: initial })
  const Native = Int32Array; let constructed = 0
  globalThis.Int32Array = new Proxy(Native, { construct(target, args) { constructed++; return Reflect.construct(target, args, target) } })
  try { expect(() => tiny.open(identity, values)).toThrow(); expect(constructed).toBe(0) }
  finally { globalThis.Int32Array = Native }
  expect(tiny.size).toBe(0); expect(tiny.usedBytes).toBe(0)
  const admitted = exact.open(identity, values); expect(exact.usedBytes).toBe(initial); admitted.pin(1)
  expect(() => exact.open({ ...identity, classificationGen: 1 }, values)).toThrow()
  expect(exact.size).toBe(1); expect(exact.usedBytes).toBe(initial); admitted.pin(-1)
  for (let i = 0; i < 10; i++) { exact.open({ ...identity, classificationGen: i }, values); expect(exact.size).toBe(1); expect(exact.usedBytes).toBe(initial) }
  exact.invalidate(); expect(exact.size).toBe(0); expect(exact.usedBytes).toBe(0)
  globalThis.Int32Array = new Proxy(Native, { construct() { throw new Error('actual Boolean constructor fault') } })
  try { expect(() => cache.open(identity, values)).toThrow('actual Boolean constructor fault') }
  finally { globalThis.Int32Array = Native }
  expect(cache.size).toBe(0); expect(cache.usedBytes).toBe(0); expect(observer.at(-1)).toBe(0)
  const last = cache.open(identity, values); expect(cache.usedBytes).toBe(initial); last.dispose(); last.dispose(); expect(cache.usedBytes).toBe(0)
  await source.release()
})

for (const profile of ['D', 'M'] as const) test(`${profile} cache construction uses exact main headroom and one authoritative owner`, async () => {
  const { source, identity, values } = domain(), ledger = new MainMemoryLedger()
  const resources = createViewerControlResources(profile, undefined, ledger), cache = resources.cache
  const first = cache.open(identity, values), initial = cache.usedBytes
  first.dispose()
  let held = pressure(resources.memory, initial - 1)
  expect(() => cache.open(identity, values)).toThrow()
  expect(cache.size).toBe(0); expect(cache.usedBytes).toBe(0); held.release()
  held = pressure(resources.memory, initial)
  const Native = Int32Array; let atConstructor = 0
  globalThis.Int32Array = new Proxy(Native, { construct(target, args) {
    atConstructor = ledger.bytes
    expect(cache.usedBytes).toBe(initial)
    expect(ledger.ownerBytes('controls', resources.cacheOwner)).toBe(initial)
    return Reflect.construct(target, args, target)
  } })
  let admitted
  try { admitted = cache.open(identity, values) } finally { globalThis.Int32Array = Native }
  expect(atConstructor).toBe(resources.memory.limitBytes)
  expect(ledger.bytes).toBe(resources.memory.limitBytes)
  admitted.dispose(); expect(cache.usedBytes).toBe(0); expect(cache.size).toBe(0)
  held.release(); resources.dispose(); await source.release()
})

test('local mask constructor faults refund construction and successful handoff charges one output', async () => {
  const source = createObjectWorkbookSource(workbook(), 'd'), geometry = source.geometry('s'), ledger = new MainMemoryLedger()
  ledger.chargeWhileReachable('geometry', [geometry], estimatePayloadBytes(geometry))
  const memory = new MainControlMemory(ledger), session = new DocumentSession('d', source, { allocation: memory })
  const before = ledger.bytes, Native = Uint32Array
  globalThis.Uint32Array = new Proxy(Native, { construct() { throw new Error('actual mask constructor fault') } })
  try { expect(() => session.view('s')).toThrow('actual mask constructor fault') } finally { globalThis.Uint32Array = Native }
  expect(ledger.bytes).toBe(before); expect(memory.pendingBytes).toBe(0); expect(session.isResident('s')).toBe(false)
  const construction = 2 * 131072 + 4100 + 1024, slice = Native.prototype.slice
  const observed: number[] = []
  globalThis.Uint32Array = new Proxy(Native, { construct(target, args) {
    observed.push(ledger.bytes - before); return Reflect.construct(target, args, target)
  } })
  Native.prototype.slice = function (...args) { observed.push(ledger.bytes - before); return slice.apply(this, args) }
  let view
  try { view = session.view('s') } finally { globalThis.Uint32Array = Native; Native.prototype.slice = slice }
  expect(observed).toEqual([construction, construction, construction])
  expect(ledger.bytes - before).toBe(135172 + 1024)
  expect(memory.pendingBytes).toBe(0)
  await session.release()
  expect(memory.owns(view.visibility.words.buffer)).toBe(true)
  expect(ledger.bytes - before).toBe(135172 + 1024)
  await source.release()
})

test('index disposal drops its Boolean lane before the release observer runs', async () => {
  const { source, identity, values } = domain()
  let index: any, observed = false
  const cache = new ValueIndexCache({ onUsage: bytes => {
    if (bytes === 0 && index?.disposed) { observed = true; expect(index.booleanIds).toBeUndefined() }
  } })
  index = cache.open(identity, values)
  index.dispose()
  expect(observed).toBe(true)
  await source.release()
})

async function observerDomain() {
  const doc: any = { sheets: [{ name: 'S', sheetId: 's', workbookIndex: 0,
    rows: [{ index: 0, cells: [{ row: 0, col: 0, ref: 'A1', value: true, styleIndex: 0 }] }],
    cols: [], merges: [], mergeRanges: [], semantics: semanticsFor('1900') }], tables: [] }
  const source = createObjectWorkbookSource(doc, 'd'), store = source.store('s')
  const body = { r0: 0, r1: 1, c0: 0, c1: 0, headerRows: 0 as const, totalsRows: 0 as const }
  const read = await store.readCell({ row: 0, col: 0 }, { stamp: store.stamp, signal: signal() })
  if (read.kind !== 'cell' || typeof read.cell.value !== 'boolean') throw new Error('Expected actual Boolean source cell')
  const classification = classifyColumn([{ kind: 'boolean', value: read.cell.value }], true)
  const identity = { ...store.stamp, range: body, col: 0, classification, classificationGen: 0 }
  return { doc, source, identity, values: valueSourceFromStore(store, 0, body) }
}
for (const profile of ['D', 'M'] as const) for (const route of ['open', 'constructor'] as const) {
  test(`${profile} ${route} initial usage observer failures refund the same main owner and permit retry`, async () => {
    const { doc, source, identity, values } = await observerDomain(), original = JSON.stringify(doc), ledger = new MainMemoryLedger()
    const primary = new Error('initial observer fault'), secondary = new Error('rollback observer fault'), calls: number[] = []
    let enabled = true
    const resources = createViewerControlResources(profile, bytes => { calls.push(bytes); if (enabled) throw bytes ? primary : secondary }, ledger)
    const cache = resources.cache, baseline = ledger.bytes, initial = 1032 + valueScope(identity).length * 2
    const create = () => route === 'open' ? cache.open(identity, values) : new ValueIndex(identity, values, cache, valueScope(identity))
    try {
      for (let attempt = 0; attempt < 3; attempt++) {
        let error: unknown; try { create() } catch (cause) { error = cause }
        expect(error).toBe(primary)
        expect(cache.usedBytes).toBe(0); expect(cache.size).toBe(0)
        expect(ledger.ownerBytes('controls', resources.cacheOwner)).toBe(0)
        expect(ledger.bytes).toBe(baseline)
        expect(calls.slice(-2)).toEqual([initial, 0])
      }
      enabled = false
      const index = create()
      expect(cache.usedBytes).toBe(1112); expect(index.ledger.chargedBytes).toBe(1112)
      expect(ledger.ownerBytes('controls', resources.cacheOwner)).toBe(1112)
      if (route === 'open') {
        enabled = true
        const count = calls.length
        expect(create()).toBe(index); expect(calls.length).toBe(count)
        enabled = false
      }
      index.dispose()
      expect(cache.usedBytes).toBe(0); expect(cache.size).toBe(0); expect(ledger.bytes).toBe(baseline)
      expect(JSON.stringify(doc)).toBe(original)
    } finally { enabled = false; resources.dispose(); await source.release() }
    expect(ledger.ownerBytes('controls', resources.cacheOwner)).toBe(0)
  })
}

for (const route of ['open', 'constructor'] as const) test(`${route} preserves a later constructor error when its rollback observer also throws`, async () => {
  const { source, identity, values } = await observerDomain(), ledger = new MainMemoryLedger()
  const primary = new Error('later result registry constructor fault'), secondary = new Error('refund observer fault')
  let refundFault = true
  const resources = createViewerControlResources('D', bytes => { if (!bytes && refundFault) throw secondary }, ledger)
  const cache = resources.cache, baseline = ledger.bytes, Native = WeakMap
  globalThis.WeakMap = new Proxy(Native, { construct() { throw primary } })
  try {
    let error: unknown
    try { if (route === 'open') cache.open(identity, values); else new ValueIndex(identity, values, cache, valueScope(identity)) } catch (cause) { error = cause }
    expect(error).toBe(primary)
    expect(cache.usedBytes).toBe(0); expect(cache.size).toBe(0)
    expect(ledger.ownerBytes('controls', resources.cacheOwner)).toBe(0); expect(ledger.bytes).toBe(baseline)
  } finally { globalThis.WeakMap = Native; refundFault = false }
  const retry = cache.open(identity, values); expect(cache.usedBytes).toBe(1112); retry.dispose()
  resources.dispose(); await source.release()
})

test('initial admission rejects quota and false or throwing growth before mutation without refunding a pinned live index', async () => {
  const { source, identity, values } = await observerDomain(), calls: number[] = [], growthError = new Error('growth fault')
  let mode = 'allow'
  const cache = new ValueIndexCache({ onUsage: bytes => calls.push(bytes), beforeGrowth: () => {
    if (mode === 'throw') throw growthError
    return mode !== 'refuse'
  } })
  const first = cache.open(identity, values), initial = cache.usedBytes
  first.pin(1)
  for (const failure of ['refuse', 'throw']) {
    mode = failure
    let error: any; try { cache.open({ ...identity, classificationGen: 1 }, values) } catch (cause) { error = cause }
    expect(failure === 'throw' ? error : error?.code).toBe(failure === 'throw' ? growthError : 'quota')
    expect(cache.usedBytes).toBe(initial); expect(cache.size).toBe(1); expect(first.disposed).toBe(false)
    expect(first.ledger.chargedBytes).toBe(initial); expect(calls).toEqual([initial])
  }
  const bounded = new ValueIndexCache({ limitBytes: initial, onUsage: bytes => calls.push(bytes) }), pinned = bounded.open(identity, values)
  pinned.pin(1)
  const notificationCount = calls.length
  expect(() => bounded.open({ ...identity, classificationGen: 1 }, values)).toThrow()
  expect(bounded.usedBytes).toBe(initial); expect(bounded.size).toBe(1); expect(calls.length).toBe(notificationCount)
  pinned.pin(-1); bounded.invalidate()
  mode = 'allow'
  const second = cache.open({ ...identity, classificationGen: 1 }, values)
  expect(cache.usedBytes).toBe(2 * initial); expect(first.disposed).toBe(false)
  second.dispose(); expect(cache.usedBytes).toBe(initial); first.pin(-1); first.dispose(); expect(cache.usedBytes).toBe(0)
  await source.release()
})

test('initial observer rollback preserves actual idle eviction and its separately held cache reservation', async () => {
  const { source, identity, values } = await observerDomain(), ledger = new MainMemoryLedger(), calls: number[] = [], fault = new Error('new index observer fault')
  let failAt = -1
  const resources = createViewerControlResources('D', bytes => { calls.push(bytes); if (bytes === failAt) throw fault }, ledger), cache = resources.cache
  const first = cache.open(identity, values), initial = cache.usedBytes, held = cache.reserve()
  // Synthetic cache pressure backed by an actual held buffer, not UI work.
  held.charge(cache.limitBytes - initial)
  let buffer: Uint8Array | undefined = new Uint8Array(cache.limitBytes - initial)
  expect(buffer.byteLength).toBe(held.bytes)
  failAt = cache.limitBytes
  let error: unknown; try { cache.open({ ...identity, classificationGen: 1 }, values) } catch (cause) { error = cause }
  expect(error).toBe(fault)
  expect(first.disposed).toBe(true); expect(cache.size).toBe(0)
  expect(cache.usedBytes).toBe(held.bytes)
  expect(ledger.ownerBytes('controls', resources.cacheOwner)).toBe(held.bytes)
  expect(calls.slice(-3)).toEqual([held.bytes, cache.limitBytes, held.bytes])
  failAt = -1
  const retry = cache.open({ ...identity, classificationGen: 1 }, values)
  expect(cache.usedBytes).toBe(cache.limitBytes); retry.dispose()
  buffer = undefined; held.close(); expect(cache.usedBytes).toBe(0)
  resources.dispose(); await source.release()
})

test('observer and later constructor failures preserve a prior pinned index and its authoritative main charge', async () => {
  const { source, identity, values } = await observerDomain(), ledger = new MainMemoryLedger()
  const observerFault = new Error('additional index observer fault'), constructorFault = new Error('additional index constructor fault')
  let failAt = -1
  const resources = createViewerControlResources('D', bytes => { if (bytes === failAt) throw observerFault }, ledger), cache = resources.cache
  const first = cache.open(identity, values), initial = cache.usedBytes, baseline = ledger.bytes
  first.pin(1)
  failAt = 2 * initial
  let error: unknown; try { cache.open({ ...identity, classificationGen: 1 }, values) } catch (cause) { error = cause }
  expect(error).toBe(observerFault)
  expect(cache.size).toBe(1); expect(cache.usedBytes).toBe(initial); expect(first.disposed).toBe(false)
  expect(ledger.bytes).toBe(baseline); expect(ledger.ownerBytes('controls', resources.cacheOwner)).toBe(initial)
  failAt = initial
  const Native = WeakMap
  globalThis.WeakMap = new Proxy(Native, { construct() { throw constructorFault } })
  try { cache.open({ ...identity, classificationGen: 2 }, values) } catch (cause) { error = cause }
  finally { globalThis.WeakMap = Native }
  expect(error).toBe(constructorFault)
  expect(cache.size).toBe(1); expect(cache.usedBytes).toBe(initial); expect(first.disposed).toBe(false)
  expect(ledger.bytes).toBe(baseline); expect(ledger.ownerBytes('controls', resources.cacheOwner)).toBe(initial)
  expect(cache.open(identity, values)).toBe(first)
  failAt = -1
  const retry = cache.open({ ...identity, classificationGen: 1 }, values)
  expect(cache.usedBytes).toBe(2 * initial); retry.dispose()
  first.pin(-1); first.dispose(); expect(cache.usedBytes).toBe(0)
  resources.dispose(); await source.release()
})
