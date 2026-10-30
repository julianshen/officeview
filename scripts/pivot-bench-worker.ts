/**
 * Task 20: Worker for pivot recomputation benchmark.
 *
 * Receives the cache definition, pivot definition, and raw XML record chunks,
 * streams decodeRecords to build the columnar RecordStore, runs 2 warm-up iterations
 * of computeAxes + computeValues, and then runs 7 timed iterations.
 *
 * Untimed Bun.gc(true) is invoked before each sample.
 * Result shape metrics (rows, cols, non-empty cells, grand totals checksum) are returned.
 */
import { initialViewState } from '../src/xlsx/pivots/engine/view'
import { computeAxes } from '../src/xlsx/pivots/engine/axes'
import { computeValues } from '../src/xlsx/pivots/engine/values'
import { RECORD_LIMITS, recordEvents } from '../src/xlsx/pivots/engine/records'
import { tokenizeBytes } from '../src/core/xml-stream'
import type { PivotCacheDefinition, PivotDiagnostic, PivotTableDefinition } from '../src/xlsx/pivots/model'

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

export interface BenchShapeReport {
  rowLines: number
  colLines: number
  nonEmptyCells: number
  grandTotalsChecksum: number
}

export interface BenchWorkerResponse {
  ok: boolean
  diagnostic?: PivotDiagnostic
  error?: string
  decodeMs: number
  storeBytes: number
  recordCount: number
  samples: BenchSample[]
  shape?: BenchShapeReport
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
      diagnostic: storeRes.diagnostic,
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
    if (typeof Bun !== 'undefined' && typeof Bun.gc === 'function') Bun.gc(true)
    const axesRes = computeAxes(store, cache, def, view)
    if (!axesRes.ok) {
      self.postMessage({
        ok: false,
        diagnostic: axesRes.diagnostic,
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
        diagnostic: valuesRes.diagnostic,
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
  let lastAxes: ReturnType<typeof computeAxes> | undefined
  let lastValues: ReturnType<typeof computeValues> | undefined

  for (let r = 0; r < timedRuns; r++) {
    if (typeof Bun !== 'undefined' && typeof Bun.gc === 'function') Bun.gc(true)

    const tAxes0 = performance.now()
    const axesRes = computeAxes(store, cache, def, view)
    const axesMs = performance.now() - tAxes0
    if (!axesRes.ok) {
      self.postMessage({
        ok: false,
        diagnostic: axesRes.diagnostic,
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
        diagnostic: valuesRes.diagnostic,
        decodeMs,
        storeBytes: store.bytes,
        recordCount: store.count,
        samples,
      } satisfies BenchWorkerResponse)
      return
    }

    lastAxes = axesRes
    lastValues = valuesRes
    samples.push({
      run: r + 1,
      axesMs,
      valuesMs,
      totalMs: axesMs + valuesMs,
    })
  }

  // Inspect shape of final result
  let shape: BenchShapeReport | undefined
  if (lastAxes && lastAxes.ok && lastValues && lastValues.ok) {
    const axes = lastAxes.value
    const grid = lastValues.value
    let nonEmptyCells = 0
    let checksum = 0

    const rowCount = axes.rows.length
    const colCount = axes.cols.length
    const dataCount = view.data.length

    for (let r = 0; r < rowCount; r++) {
      for (let c = 0; c < colCount; c++) {
        for (let d = 0; d < dataCount; d++) {
          const cell = grid.get(r, c, d)
          if (cell.kind === 'number') {
            nonEmptyCells++
            // Add grand total line contributions or cell value to checksum
            checksum += cell.value
          }
        }
      }
    }

    shape = {
      rowLines: rowCount,
      colLines: colCount,
      nonEmptyCells,
      grandTotalsChecksum: checksum,
    }
  }

  self.postMessage({
    ok: true,
    decodeMs,
    storeBytes: store.bytes,
    recordCount: store.count,
    samples,
    shape,
  } satisfies BenchWorkerResponse)
}
