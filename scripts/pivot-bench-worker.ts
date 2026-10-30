/**
 * Task 20: Worker for pivot recomputation benchmark.
 *
 * Receives the cache definition, pivot definition, and raw XML record chunks,
 * streams decodeRecords to build the columnar RecordStore, runs 2 warm-up iterations
 * of computeAxes + computeValues, and then runs 7 timed iterations.
 */
import { initialViewState } from '../src/xlsx/pivots/engine/view'
import { computeAxes } from '../src/xlsx/pivots/engine/axes'
import { computeValues } from '../src/xlsx/pivots/engine/values'
import { RECORD_LIMITS, recordEvents } from '../src/xlsx/pivots/engine/records'
import { tokenizeBytes } from '../src/core/xml-stream'
import type { PivotCacheDefinition, PivotTableDefinition } from '../src/xlsx/pivots/model'

interface BenchWorkerPayload {
  cache: PivotCacheDefinition
  def: PivotTableDefinition
  chunks: Uint8Array[]
  warmupRuns?: number
  timedRuns?: number
}

export interface BenchSample {
  run: number
  axesMs: number
  valuesMs: number
  totalMs: number
}

export interface BenchWorkerResponse {
  ok: boolean
  error?: string
  decodeMs: number
  storeBytes: number
  recordCount: number
  samples: BenchSample[]
}

declare const self: Worker

self.onmessage = async (e: MessageEvent<BenchWorkerPayload>) => {
  const { cache, def, chunks, warmupRuns = 2, timedRuns = 7 } = e.data

  async function* source(): AsyncIterable<Uint8Array> {
    for (const chunk of chunks) yield chunk
  }

  const tDecodeStart = performance.now()
  const ev = recordEvents(cache, RECORD_LIMITS.D, false)
  try {
    await tokenizeBytes(source(), ev.events)
  } catch (err) {
    self.postMessage({
      ok: false,
      error: `tokenizeBytes error: ${err instanceof Error ? err.message : String(err)}`,
      decodeMs: 0,
      storeBytes: 0,
      recordCount: 0,
      samples: [],
    } satisfies BenchWorkerResponse)
    return
  }

  const storeRes = ev.finish()
  const decodeMs = performance.now() - tDecodeStart

  if (!storeRes.ok) {
    self.postMessage({
      ok: false,
      error: `record decode failed: ${storeRes.diagnostic.message}`,
      decodeMs,
      storeBytes: 0,
      recordCount: 0,
      samples: [],
    } satisfies BenchWorkerResponse)
    return
  }

  const store = storeRes.store
  const view = initialViewState(def, cache)

  // Warm-up runs
  for (let w = 0; w < warmupRuns; w++) {
    const axesRes = computeAxes(store, cache, def, view)
    if (!axesRes.ok) {
      self.postMessage({
        ok: false,
        error: `warm-up axes error: ${axesRes.diagnostic.message}`,
        decodeMs,
        storeBytes: store.bytes,
        recordCount: store.count,
        samples: [],
      } satisfies BenchWorkerResponse)
      return
    }
    const valuesRes = computeValues(store, cache, def, view, axesRes.value)
    if (!valuesRes.ok) {
      self.postMessage({
        ok: false,
        error: `warm-up values error: ${valuesRes.diagnostic.message}`,
        decodeMs,
        storeBytes: store.bytes,
        recordCount: store.count,
        samples: [],
      } satisfies BenchWorkerResponse)
      return
    }
  }

  // Timed runs
  const samples: BenchSample[] = []
  for (let r = 0; r < timedRuns; r++) {
    const tStart = performance.now()
    const tAxes0 = performance.now()
    const axesRes = computeAxes(store, cache, def, view)
    const axesMs = performance.now() - tAxes0
    if (!axesRes.ok) {
      self.postMessage({
        ok: false,
        error: `timed axes error on run ${r + 1}: ${axesRes.diagnostic.message}`,
        decodeMs,
        storeBytes: store.bytes,
        recordCount: store.count,
        samples,
      } satisfies BenchWorkerResponse)
      return
    }

    const tValues0 = performance.now()
    const valuesRes = computeValues(store, cache, def, view, axesRes.value)
    const valuesMs = performance.now() - tValues0
    if (!valuesRes.ok) {
      self.postMessage({
        ok: false,
        error: `timed values error on run ${r + 1}: ${valuesRes.diagnostic.message}`,
        decodeMs,
        storeBytes: store.bytes,
        recordCount: store.count,
        samples,
      } satisfies BenchWorkerResponse)
      return
    }

    const totalMs = performance.now() - tStart
    samples.push({
      run: r + 1,
      axesMs,
      valuesMs,
      totalMs,
    })
  }

  self.postMessage({
    ok: true,
    decodeMs,
    storeBytes: store.bytes,
    recordCount: store.count,
    samples,
  } satisfies BenchWorkerResponse)
}
