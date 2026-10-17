// @vitest-environment node
// S13.2 drawing capabilities must survive S13.3 cancellation and deadline adapters.
import { afterEach, expect, it, vi } from 'vitest'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { initialIdentity } from '../src/worker/viewerProtocol'
import type { SheetDrawings, WorkerReply } from '../src/xlsx/view/contracts'
import { workbookOf } from './helpers/xlsx-find'

afterEach(() => vi.useRealTimers())
const drawings: SheetDrawings = { completeness: 'complete', drawings: [], images: [], bytes: 0, diagnostics: [{ code: 'drawings-probe', message: 'Drawing provider result' }] }
const fixture = () => {
  const document = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 5 }] }])
  const source = createObjectWorkbookSource(document, 'doc')
  const kernel = new ViewerKernel({ parse: async () => document, openSource: async () => source }, { yield: async () => {} })
  return { document, source, kernel }
}
const flush = async () => { for (let i = 0; i < 30; i++) await Promise.resolve() }

it('forwards the existing drawing capability with its receiver, profile and memory share', async () => {
  const { document, source, kernel } = fixture(), store = source.store('s')
  let during = 0
  const provider = vi.spyOn(store, 'readDrawings').mockImplementation(async function (this: typeof store, options, profile, memory) {
    expect(this).toBe(store)
    expect(options.stamp).toMatchObject({ docId: 'doc', sheetId: 's', valueGen: 0 })
    expect(profile).toBe('M')
    expect(memory?.set(128)).toBe(true)
    during = kernel.lifecycleStats().transientBytes
    return drawings
  })
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'doc', source: { kind: 'object', document }, profile: 'M' }, () => {})
  const replies: WorkerReply[] = []
  try {
    await kernel.handle({ kind: 'read', requestId: 'drawings', owner: initialIdentity('doc', 's'), query: { op: 'drawings' } }, out => replies.push(out))
    expect(replies).toMatchObject([{ kind: 'complete', result: { kind: 'read', result: { op: 'drawings', result: drawings } } }])
    expect(provider).toHaveBeenCalledTimes(1)
    expect(during).toBeGreaterThanOrEqual(128)
    expect(kernel.lifecycleStats().transientBytes).toBe(0)
  } finally { await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {}) }
})

it.each(['timeout', 'cancel'] as const)('interrupts a signal-ignoring drawing provider on %s and ignores its late result', async mode => {
  vi.useFakeTimers()
  const { document, source, kernel } = fixture()
  let release!: () => void
  const held = new Promise<void>(resolve => { release = resolve })
  let providerSignal: AbortSignal | undefined, lateAccepted: boolean | undefined
  const provider = vi.spyOn(source.store('s'), 'readDrawings').mockImplementation(async (options, _profile, memory) => {
    providerSignal = options.signal
    memory?.set(128)
    await held
    lateAccepted = memory?.set(256)
    return drawings
  })
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
  const replies: WorkerReply[] = []
  const pending = kernel.handle({ kind: 'read', requestId: 'drawings', owner: initialIdentity('doc', 's'), query: { op: 'drawings' } }, out => replies.push(out))
  try {
    await flush()
    expect(provider, 'the adapter must actually invoke the available provider').toHaveBeenCalledTimes(1)
    expect(kernel.lifecycleStats().transientBytes).toBeGreaterThanOrEqual(128)
    if (mode === 'timeout') await vi.advanceTimersByTimeAsync(120_001)
    else {
      await kernel.handle({ kind: 'cancel', requestId: 'cancel', targetRequestId: 'drawings' }, () => {})
      await flush()
      expect(replies, 'cancellation completes before the ignored provider resumes').toMatchObject([{ kind: 'cancelled' }])
    }
    await pending
    expect(replies).toMatchObject([{ kind: 'cancelled', diagnostics: [{ code: mode === 'timeout' ? 'query-timeout' : 'cancel-requested' }] }])
    // FC-T2: the request settled, but the provider still holds its working set
    // until it actually returns; its share stays charged until then.
    expect(kernel.lifecycleStats()).toMatchObject({ inflight: 0 })
    expect(kernel.lifecycleStats().transientBytes).toBeGreaterThanOrEqual(128)
    expect(providerSignal?.aborted).toBe(true)
    expect(lateAccepted).toBeUndefined()
    release()
    await flush()
    expect(replies).toHaveLength(1)
    expect(kernel.lifecycleStats()).toMatchObject({ inflight: 0, transientBytes: 0 })
    expect(lateAccepted).toBe(true)
    provider.mockResolvedValueOnce(drawings)
    const next: WorkerReply[] = []
    await kernel.handle({ kind: 'read', requestId: 'next', owner: initialIdentity('doc', 's'), query: { op: 'drawings' } }, out => next.push(out))
    expect(next).toMatchObject([{ kind: 'complete' }])
    expect(kernel.lifecycleStats().transientBytes).toBe(0)
  } finally { release(); await pending; await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {}) }
})

it('preserves the diagnosed unavailable result when a source has no drawing capability', async () => {
  const { document, source, kernel } = fixture()
  const store = source.store('s')
  const withoutDrawings = new Proxy(store, {
    get(target, key) { if (key === 'readDrawings') return undefined; const value = Reflect.get(target, key); return typeof value === 'function' ? value.bind(target) : value },
  })
  vi.spyOn(source, 'store').mockReturnValue(withoutDrawings)
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
  const replies: WorkerReply[] = []
  try {
    await kernel.handle({ kind: 'read', requestId: 'drawings', owner: initialIdentity('doc', 's'), query: { op: 'drawings' } }, out => replies.push(out))
    expect(replies).toMatchObject([{ kind: 'complete', result: { result: { result: { completeness: 'incomplete', diagnostics: [{ code: 'drawings-unavailable', message: 'This source cannot serve drawing content; drawings are not shown' }] } } } }])
  } finally { await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {}) }
})
