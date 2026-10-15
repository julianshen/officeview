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
async function fixture(profile: 'D' | 'M', ledger = new MainMemoryLedger(), merged = false) {
  const n = 16, rows = Array.from({ length: n + 1 }, (_, row) => ({ index: row, cells: [0, 1].map(col => ({ row, col, ref: '', value: row ? row : 'Header', styleIndex: 0, style: { numFmtId: 0 } })) }))
  const doc: any = { sheets: [{ name: 'S', sheetId: 's', workbookIndex: 0, rows, cols: [], merges: [], mergeRanges: merged ? [{ minRow: 2, maxRow: 3, minCol: 0, maxCol: 1 }] : [], semantics: semanticsFor('1900') }], tables: [] }
  const source = createObjectWorkbookSource(doc, 'd'), store = source.store('s')
  const resources = createViewerControlResources(profile, undefined, ledger), hidden = createIntervalSet([])
  let gen = 0, valueGen = 0, publications = 0
  let view: any = { sheetId: 's', revision: 0, order: createIdentityOrder(), baseHidden: hidden, visibility: visibilityFromHidden(hidden), sort: null, filter: null, selection: { ...emptyGridSelection(), active: { row: 9, col: 0 } } }
  const initial = view, sourceJSON = JSON.stringify(doc), selectionJSON = JSON.stringify(view.selection), stamp = { ...store.stamp }
  const owner = () => ({ ...store.stamp, valueGen: store.stamp.valueGen + valueGen, fontGen: 0, viewGen: gen })
  const snapshot = () => ({ owner: owner(), filter: view.filter, sort: view.sort, records: { visible: n + 1, total: n + 1, completeness: 'complete' as const } })
  const controls = await viewerControls({ doc, source, sheetId: 's', resources, ledger, snapshot, view: () => view, publish: next => { publications++; gen++; view = { ...view, ...next }; return snapshot() } }, new AbortController().signal)
  const spec = { range: { r0: 0, r1: n, c0: 0, c1: 0, headerRows: 1 as const, totalsRows: 0 as const }, columns: new Map([[0, { kind: 'number' as const, op: 'gt' as const, first: 8 }]]) }
  return { source, store, ledger, resources, controls, owner, spec, bump: () => gen++, bumpValue: () => valueGen++, view: () => view, publications: () => publications,
    unchanged() { assert.equal(view, initial); assert.equal(publications, 0); assert.equal(JSON.stringify(doc), sourceJSON); assert.equal(JSON.stringify(view.selection), selectionJSON); assert.deepEqual(store.stamp, stamp) },
    async close() { controls.dispose(); resources.dispose(); await source.release() } }
}
// Private-seam observables track actual admission/adoption, not estimated heap.
function observe(ledger: MainMemoryLedger) {
  const set = ledger.set.bind(ledger), adopt = ledger.adoptWhileReachable.bind(ledger), native = globalThis.Uint8Array
  const flagsOwners: string[] = [], metadataOwners: string[] = [], states: any[] = [], tokens: string[] = [], engineRefs: any[] = [], memoryRefs: any[] = [], totals: number[][] = []
  let constructions = 0
  ledger.set = (pool, owner, bytes) => { if (pool === 'controls' && bytes === 145 && !flagsOwners.includes(owner)) flagsOwners.push(owner); if (pool === 'controls' && bytes === 4608) metadataOwners.push(owner); set(pool, owner, bytes) }
  ledger.adoptWhileReachable = (fp, fo, pool, targets, bytes, commit) => {
    const before = [ledger.bytes, ledger.peakBytes]
    const token = adopt(fp, fo, pool, targets, bytes, commit)
    if (metadataOwners.includes(fo)) { states.push(targets[2] ? new WeakRef(targets[2]) : undefined); engineRefs.push(new WeakRef(targets[0])); memoryRefs.push(new WeakRef(targets[1])); tokens.push(token); totals.push([...before, ledger.bytes, ledger.peakBytes]) }
    return token
  }
  globalThis.Uint8Array = new Proxy(native, { construct(target, args) { if (args[0] === 17 && new Error().stack?.includes('viewerControlsClient')) constructions++; return Reflect.construct(target, args, target) } })
  return { flagsOwners, metadataOwners, states, tokens, engineRefs, memoryRefs, totals, constructions: () => constructions,
    bytes: (i = 0) => ledger.ownerBytes('controls', flagsOwners[i]),
    restore() { ledger.set = set; ledger.adoptWhileReachable = adopt; globalThis.Uint8Array = native } }
}
for (const profile of ['D', 'M'] as const) {
  for (const allowance of [4608, 4607]) test(`${profile} metadata admission at ${allowance} remaining bytes precedes engine and flags`, async () => {
    const f = await fixture(profile), t = observe(f.ledger), memory = f.resources.memory, reserve = memory.reserve.bind(memory)
    const filter = DocumentSession.prototype.filter
    let engines = 0, admitted = 0
    DocumentSession.prototype.filter = async function(...args) { engines++; return filter.apply(this, args) }
    memory.reserve = () => {
      const account = reserve(), charge = account.charge.bind(account)
      account.charge = bytes => {
        if (bytes !== 4608) return charge(bytes)
        f.ledger.set('copy', 'actual-parent-pressure', memory.limitBytes - f.ledger.bytes - allowance)
        try { charge(bytes); admitted++ } finally { f.ledger.release('copy', 'actual-parent-pressure') }
      }
      return account
    }
    try {
      if (allowance === 4608) { assert.equal((await f.controls.client.filter(f.owner(), f.spec, new AbortController().signal)).kind, 'committed'); assert.equal(admitted, 1); assert.equal(engines, 1); assert.equal(t.constructions(), 1); assert.equal(t.metadataOwners.length, 1); assert.ok(t.totals.every(([a,b,c,d]) => a === c && b === d)) }
      else { await assert.rejects(f.controls.client.filter(f.owner(), f.spec, new AbortController().signal), (e: any) => e.code === 'quota'); assert.equal(admitted, 0); assert.equal(engines, 0); assert.equal(t.constructions(), 0); assert.equal(t.metadataOwners.length, 0); f.unchanged() }
      assert.equal(memory.pendingBytes, 0)
    } finally { memory.reserve = reserve; DocumentSession.prototype.filter = filter; t.restore(); await f.close() }
  })
  for (const fault of ['register-third', 'publication', 'release'] as const) test(`${profile} ${fault} cleanup retains reservation and original primary source error`, async () => {
    const primary = new Error('primary actual source'), cleanup = new Error('cleanup fault'), registrations: string[] = []
    let armed = false, count = 0
    const ledger = new MainMemoryLedger({ register(_target, token) { registrations.push(token); if (armed && ++count === 3) throw cleanup } })
    const f = await fixture(profile, ledger), t = observe(ledger), adopt = ledger.adoptWhileReachable.bind(ledger), release = DocumentSession.prototype.release
    const query = f.store.queryColumn.bind(f.store)
    f.store.queryColumn = async (...args) => { if (t.constructions()) throw primary; return query(...args) }
    let before: number[] = [], after: number[] = []
    ledger.adoptWhileReachable = (fp, fo, pool, targets, bytes, commit) => {
      if (!t.metadataOwners.includes(fo)) return adopt(fp, fo, pool, targets, bytes, commit)
      before = [ledger.bytes, ledger.peakBytes]
      armed = fault === 'register-third'; count = 0
      try { return adopt(fp, fo, pool, targets, bytes, fault === 'publication' ? () => { throw cleanup } : commit) }
      finally { armed = false; after = [ledger.bytes, ledger.peakBytes] }
    }
    DocumentSession.prototype.release = async function() { await release.call(this); if (fault === 'release') throw cleanup }
    try {
      await assert.rejects(f.controls.client.filter(f.owner(), f.spec, new AbortController().signal), e => e === primary)
      assert.equal(t.bytes(), 0); assert.equal(t.constructions(), 1); f.unchanged()
      assert.deepEqual(after, before, 'adoption cannot create a total or peak trough')
      assert.equal(ledger.ownerBytes('controls', t.metadataOwners[0]), fault === 'release' ? 0 : 4608)
      assert.equal(f.resources.memory.pendingBytes, fault === 'release' ? 0 : 4608)
      for (const token of registrations) ledger.collected(token)
      if (fault !== 'release') assert.equal(ledger.ownerBytes('controls', t.metadataOwners[0]), 4608, 'failed registration callbacks cannot refund explicit reservation')
    } finally { DocumentSession.prototype.release = release; ledger.adoptWhileReachable = adopt; t.restore(); await f.close() }
  })
}
for (const profile of ['D', 'M'] as const) {
  test(`${profile} resources close alone stays pending until actual provider exits`, async () => {
    const f = await fixture(profile), t = observe(f.ledger), held = gate(), query = f.store.queryColumn.bind(f.store)
    let hit = false, settled = false
    f.store.queryColumn = async (...args) => { const page = await query(...args); if (!hit && t.constructions()) { hit = true; held.mark(); await held.wait }; return page }
    const pending = f.controls.client.filter(f.owner(), f.spec, new AbortController().signal).then(value => { settled = true; return value })
    try { await held.reached; f.resources.dispose(); f.resources.dispose(); await collect(); assert.equal(settled, false); assert.equal(t.bytes(), 145); f.unchanged(); held.open(); assert.equal((await pending).kind, 'cancelled'); assert.equal(t.bytes(), 0); f.unchanged() }
    finally { held.open(); await pending.catch(() => {}); t.restore(); await f.close() }
  }, 10000)
  for (const mode of ['abort', 'unchanged', 'sort', 'clear', 'invalid-range'] as const) test(`${profile} ${mode} starts no producer`, async () => {
    const f = await fixture(profile)
    if (mode === 'clear') assert.equal((await f.controls.client.filter(f.owner(), f.spec, new AbortController().signal)).kind, 'committed')
    const t = observe(f.ledger), caller = new AbortController()
    if (mode === 'abort') caller.abort()
    try {
      const result = mode === 'sort' ? await f.controls.client.sort(f.owner(), { range: { ...f.spec.range, c1: 1 }, keys: [{ col: 0, order: 'desc' }], caseSensitive: false }, caller.signal)
        : await f.controls.client.filter(f.owner(), mode === 'unchanged' || mode === 'clear' ? null : mode === 'invalid-range' ? { ...f.spec, range: { ...f.spec.range, r1: -1 } } : f.spec, caller.signal)
      assert.equal(result.kind, mode === 'abort' ? 'cancelled' : mode === 'unchanged' ? 'unchanged' : mode === 'invalid-range' ? 'refused' : 'committed')
      assert.equal(t.constructions(), 0); assert.equal(t.flagsOwners.length, 0); assert.equal(f.resources.memory.pendingBytes, 0)
      for (const ref of t.states) { const state = ref.deref(); assert.equal(state.producer, 'new'); assert.equal(state.drained, true); assert.equal(state.consumerDone, true) }
      if (mode === 'abort' || mode === 'unchanged' || mode === 'invalid-range') f.unchanged()
    } finally { t.restore(); await f.close() }
  })
  for (const fault of ['flags-constructor', 'flags-quota', 'source', 'engine-constructor', 'release-only', 'adoption-only'] as const) test(`${profile} ${fault} exits without partial publication`, async () => {
    const f = await fixture(profile), t = observe(f.ledger), primary = new Error(`actual ${fault}`), native = globalThis.Uint8Array
    const reserve = f.resources.memory.reserve.bind(f.resources.memory), query = f.store.queryColumn.bind(f.store), release = DocumentSession.prototype.release
    const { MemoryPool } = await import('../src/xlsx/view/tileCache'), protect = MemoryPool.prototype.protect, adopt = f.ledger.adoptWhileReachable.bind(f.ledger)
    if (fault === 'engine-constructor') MemoryPool.prototype.protect = () => { throw primary }
    if (fault === 'flags-constructor') globalThis.Uint8Array = new Proxy(native, { construct(target, args) { if (args[0] === 17 && new Error().stack?.includes('viewerControlsClient')) throw primary; return Reflect.construct(target, args, target) } })
    if (fault === 'flags-quota') f.resources.memory.reserve = () => { const a = reserve(), charge = a.charge.bind(a); a.charge = bytes => { if (bytes !== 145) return charge(bytes); f.ledger.set('copy', 'actual-pressure', f.resources.memory.limitBytes - f.ledger.bytes - 144); try { charge(bytes) } finally { f.ledger.release('copy', 'actual-pressure') } }; return a }
    if (fault === 'source') f.store.queryColumn = async (...args) => { if (t.constructions()) throw primary; return query(...args) }
    if (fault === 'release-only') DocumentSession.prototype.release = async function() { await release.call(this); throw primary }
    if (fault === 'adoption-only') f.ledger.adoptWhileReachable = (fp, fo, pool, targets, bytes, commit) => { if (t.metadataOwners.includes(fo)) throw primary; return adopt(fp, fo, pool, targets, bytes, commit) }
    try {
      if (fault === 'flags-quota') assert.equal((await f.controls.client.filter(f.owner(), f.spec, new AbortController().signal)).kind, 'refused')
      else await assert.rejects(f.controls.client.filter(f.owner(), f.spec, new AbortController().signal), e => e === primary)
      assert.equal(t.bytes(), 0); f.unchanged()
      assert.equal(f.resources.memory.pendingBytes, fault === 'adoption-only' ? 4608 : 0)
      if (fault === 'flags-quota' || fault === 'engine-constructor') assert.equal(t.constructions(), 0)
    } finally { MemoryPool.prototype.protect = protect; DocumentSession.prototype.release = release; f.resources.memory.reserve = reserve; globalThis.Uint8Array = native; t.restore(); await f.close() }
  })
}
for (const profile of ['D', 'M'] as const) {
  test(`${profile} metadata watcher retains charge until the third state target is reported collected`, async () => {
    const ledger = new MainMemoryLedger({ register() {} }), f = await fixture(profile, ledger), t = observe(ledger)
    try {
      assert.equal((await f.controls.client.filter(f.owner(), f.spec, new AbortController().signal)).kind, 'committed')
      assert.equal(t.tokens.length, 1); const token = t.tokens[0]
      assert.equal(ledger.ownerBytes('controls', token), 4608)
      ledger.collected(token); ledger.collected(token)
      assert.equal(ledger.ownerBytes('controls', token), 4608, 'engine and memory collection cannot refund state metadata')
      ledger.collected(token); assert.equal(ledger.ownerBytes('controls', token), 0)
      ledger.collected(token); assert.equal(ledger.ownerBytes('controls', token), 0)
      assert.ok(t.totals.every(([a,b,c,d]) => a === c && b === d))
    } finally { t.restore(); await f.close() }
  })
  test(`${profile} actual GC leaves only revoked matcher state after engine collection`, async () => {
    const f = await fixture(profile), t = observe(f.ledger), prototype = DocumentSession.prototype as any, project = prototype.projectDraft
    let matcher: ((row: number) => boolean) | undefined
    prototype.projectDraft = function(...args: any[]) { matcher = args[3]; return project.apply(this, args) }
    try {
      assert.equal((await f.controls.client.filter(f.owner(), f.spec, new AbortController().signal)).kind, 'committed')
      prototype.projectDraft = project
      await collect()
      assert.equal(t.engineRefs[0].deref(), undefined, 'real engine must collect while matcher survives')
      assert.equal(t.memoryRefs[0].deref(), undefined, 'state-only evidence also requires its old memory target to collect')
      assert.ok(t.states[0].deref()); assert.equal(f.ledger.ownerBytes('controls', t.tokens[0]), 4608)
      assert.throws(() => matcher!(9), /released/); assert.equal(t.bytes(), 0)
      matcher = undefined; await collect(); assert.equal(t.states[0].deref(), undefined); assert.equal(f.ledger.ownerBytes('controls', t.tokens[0]), 0)
    } finally { prototype.projectDraft = project; matcher = undefined; t.restore(); await f.close() }
  }, 10000)
  test(`${profile} actual held producer retains engine and state metadata until continuation exits`, async () => {
    const f = await fixture(profile), t = observe(f.ledger), held = gate(), caller = new AbortController(), query = f.store.queryColumn.bind(f.store)
    let hit = false
    f.store.queryColumn = async (...args) => { const page = await query(...args); if (!hit && t.constructions()) { hit = true; held.mark(); await held.wait }; return page }
    const pending = f.controls.client.filter(f.owner(), f.spec, caller.signal)
    try {
      await held.reached; caller.abort(); assert.equal((await pending).kind, 'cancelled'); await collect()
      assert.ok(t.engineRefs[0].deref(), 'Node retains the existing cooperative request engine while this producer is held')
      assert.ok(t.states[0].deref()); assert.equal(f.ledger.ownerBytes('controls', t.tokens[0]), 4608); assert.equal(t.bytes(), 145)
      held.open(); await collect(); assert.equal(t.bytes(), 0); f.unchanged()
    } finally { held.open(); await pending.catch(() => {}); t.restore(); await f.close() }
  }, 10000)
}
for (const profile of ['D', 'M'] as const) {
  for (const columns of ['same', 'different'] as const) for (const mode of ['reverse', 'old-cancel', 'value-stale', 'view-stale'] as const) test(`${profile} ${columns} concurrent stages ${mode} keep independent owners`, async () => {
    const f = await fixture(profile), t = observe(f.ledger), held = [gate(), gate()], caller = [new AbortController(), new AbortController()]
    const prototype = DocumentSession.prototype as any, filter = prototype.filter, query = f.store.queryColumn.bind(f.store), ids = new WeakMap<AbortSignal, number>()
    let next = 0; const hit = [false, false]
    prototype.filter = function(...args: any[]) { ids.set(args[3], next++); return filter.apply(this, args) }
    f.store.queryColumn = async (...args) => { const page = await query(...args), id = ids.get(args[1].signal!)!; if (id !== undefined && !hit[id] && t.constructions() >= id + 1) { hit[id] = true; held[id].mark(); await held[id].wait }; return page }
    const spec2 = columns === 'same' ? { ...f.spec, columns: new Map([[0, { kind: 'number' as const, op: 'gt' as const, first: 10 }]]) } : { range: { ...f.spec.range, c1: 1 }, columns: new Map([[1, { kind: 'number' as const, op: 'gt' as const, first: 10 }]]) }
    const first = f.controls.client.filter(f.owner(), f.spec, caller[0].signal); void first.catch(() => {})
    await held[0].reached
    const second = f.controls.client.filter(f.owner(), spec2, caller[1].signal); void second.catch(() => {})
    try {
      await held[1].reached; assert.equal(t.bytes(0), 145); assert.equal(t.bytes(1), 145); assert.equal(t.flagsOwners.length, 2)
      if (mode === 'old-cancel') { caller[0].abort(); assert.equal((await first).kind, 'cancelled'); assert.equal(t.bytes(0), 145); assert.equal(t.bytes(1), 145) }
      if (mode === 'value-stale') f.bumpValue()
      if (mode === 'view-stale') f.bump()
      held[1].open(); const newer = mode === 'value-stale' || mode === 'view-stale' ? (await assert.rejects(second, (e: any) => e.code === 'stale-source'), { kind: 'stale' }) : await second
      assert.equal(newer.kind, mode === 'value-stale' || mode === 'view-stale' ? 'stale' : 'committed'); assert.equal(t.bytes(1), 0); assert.equal(t.bytes(0), 145)
      if (newer.kind === 'committed') { assert.equal(f.publications(), 1); assert.equal(f.view().visibility.visible(9), false); assert.equal(f.view().visibility.visible(11), true) }
      held[0].open(); const older = mode === 'old-cancel' ? await first : (await assert.rejects(first, (e: any) => e.code === 'stale-source'), { kind: 'stale' })
      await task(); await task()
      assert.equal(older.kind, mode === 'old-cancel' ? 'cancelled' : 'stale'); assert.equal(t.bytes(0), 0); assert.equal(t.bytes(1), 0)
      assert.equal(f.publications(), newer.kind === 'committed' ? 1 : 0)
      f.controls.dispose(); f.controls.dispose(); f.resources.dispose(); f.resources.dispose(); assert.equal(t.bytes(0), 0); assert.equal(t.bytes(1), 0)
    } finally { held.forEach(g => g.open()); await Promise.allSettled([first, second]); prototype.filter = filter; t.restore(); await f.close() }
  }, 10000)
  test(`${profile} duplicate running and late private evaluator starts cannot allocate or revive`, async () => {
    const f = await fixture(profile), t = observe(f.ledger), held = gate(), caller = new AbortController(), prototype = DocumentSession.prototype as any, filter = prototype.filter, query = f.store.queryColumn.bind(f.store)
    let evaluate: any, request: any, hit = false
    prototype.filter = function(...args: any[]) { const original = this.options.filterEvaluator; evaluate = original; this.options.filterEvaluator = (input: any) => { request = input; return original(input) }; return filter.apply(this, args) }
    f.store.queryColumn = async (...args) => { const page = await query(...args); if (!hit && t.constructions()) { hit = true; held.mark(); await held.wait }; return page }
    const pending = f.controls.client.filter(f.owner(), f.spec, caller.signal)
    try {
      await held.reached; assert.equal(t.bytes(), 145)
      await assert.rejects(evaluate(request), (e: any) => e.code === 'cancelled'); assert.equal(t.constructions(), 1); assert.equal(t.bytes(), 145)
      caller.abort(); assert.equal((await pending).kind, 'cancelled'); assert.equal(t.bytes(), 145)
      held.open(); await task(); await task(); assert.equal(t.bytes(), 0)
      await assert.rejects(evaluate({ ...request, signal: new AbortController().signal }), (e: any) => e.code === 'cancelled')
      assert.equal(t.constructions(), 1); assert.equal(t.bytes(), 0); f.unchanged()
    } finally { held.open(); await pending.catch(() => {}); prototype.filter = filter; t.restore(); await f.close() }
  }, 10000)
}
for (const profile of ['D', 'M'] as const) test(`${profile} evaluator never used by sort cannot start after its consumer closes`, async () => {
  const f = await fixture(profile), t = observe(f.ledger), prototype = DocumentSession.prototype as any, sort = prototype.sort
  let evaluate: any
  prototype.sort = function(...args: any[]) { evaluate = this.options.filterEvaluator; return sort.apply(this, args) }
  try {
    assert.equal((await f.controls.client.sort(f.owner(), { range: { ...f.spec.range, c1: 1 }, keys: [{ col: 0, order: 'desc' }], caseSensitive: false }, new AbortController().signal)).kind, 'committed')
    assert.ok(evaluate); assert.equal(t.states[0].deref().producer, 'new'); assert.equal(t.states[0].deref().drained, true)
    let lateError: any
    try { await evaluate({ spec: f.spec, signal: new AbortController().signal, store: f.store, nowMs: Date.now() }) } catch (error) { lateError = error }
    assert.equal(t.constructions(), 0, 'late unused evaluator cannot construct flags after consumer drain')
    assert.equal(t.flagsOwners.length, 0); assert.equal(lateError?.code, 'cancelled')
  } finally { prototype.sort = sort; t.restore(); await f.close() }
})

for (const profile of ['D', 'M'] as const) test(`${profile} actual structural child-only range refusal creates no stage or producer`, async () => {
  const f = await fixture(profile, new MainMemoryLedger(), true), t = observe(f.ledger)
  f.view().selection = { ...emptyGridSelection(), active: { row: 3, col: 1 }, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 2, r1: 3 }] }, cols: [{ c0: 1, c1: 1 }] }] }
  const selection = f.view().selection, before = f.view(), stamp = { ...f.store.stamp }
  try {
    const result = await f.controls.client.resolve(f.owner(), 'filter', new AbortController().signal, 0)
    assert.equal(result.kind, 'refused')
    if (result.kind === 'refused') assert.equal(result.refusal.reason, 'merge')
    assert.equal(t.metadataOwners.length, 0); assert.equal(t.constructions(), 0); assert.equal(t.flagsOwners.length, 0); assert.equal(f.publications(), 0)
    assert.equal(f.view(), before); assert.equal(f.view().selection, selection); assert.deepEqual(f.store.stamp, stamp)
  } finally { t.restore(); await f.close() }
})
