/**
 * Task 20: 1,000,000-record pivot benchmark script (Stage 2 gate).
 *
 * Spec §6.3:
 * - 1,000,000 records;
 * - row fields with cardinalities 10, 50 and 200;
 * - column field with cardinality 12;
 * - two numeric data fields.
 *
 * Fields are drawn independently using seeded PRNGs (mulberry32).
 * Default subtotals and grand totals enabled per Excel defaults.
 *
 * Runs decodeRecords (timed separately), then inside a Bun Worker:
 * - 2 warm-up runs of computeAxes + computeValues;
 * - 7 timed runs of computeAxes + computeValues;
 * - untimed Bun.gc(true) before each timed sample;
 * - reports median, min, max, each sample, store.bytes, shape metrics.
 *
 * Spawns the worker N times sequentially when `--runs N` (default 3) is given.
 * Prints GATE PASS|FAIL median-of-medians=<ms> (≤ 200) or GATE ERROR <feature>.
 * Exits 0 either way.
 *
 * Usage:
 *   bun scripts/pivot-bench.ts [--runs N]
 */
import { parseXmlOrdered } from '../src/core/xml'
import { parseCacheDefinition } from '../src/xlsx/pivots/parseCacheDefinition'
import { parsePivotTable } from '../src/xlsx/pivots/parsePivotTable'
import type { BenchWorkerResponse } from './pivot-bench-worker'

const RECORD_COUNT = 1_000_000
const WARMUP_RUNS = 2
const TIMED_RUNS = 7
const TARGET_MEDIAN_MS = 200
const BATCH_SIZE = 10_000

function parseArgs(): { runs: number } {
  let runs = 3
  for (let i = 2; i < process.argv.length; i++) {
    if (process.argv[i] === '--runs' && i + 1 < process.argv.length) {
      const n = parseInt(process.argv[++i]!, 10)
      if (Number.isSafeInteger(n) && n > 0) runs = n
    }
  }
  return { runs }
}

function mulberry32(a: number) {
  return function() {
    let t = a += 0x6D2B79F5
    t = Math.imul(t ^ (t >>> 15), t | 1)
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

function buildCacheXml(recordCount: number): string {
  const r1Items = Array.from({ length: 10 }, (_, i) => `<s v="R1_${i}"/>`).join('')
  const r2Items = Array.from({ length: 50 }, (_, i) => `<s v="R2_${i}"/>`).join('')
  const r3Items = Array.from({ length: 200 }, (_, i) => `<s v="R3_${i}"/>`).join('')
  const c1Items = Array.from({ length: 12 }, (_, i) => `<s v="C1_${i}"/>`).join('')

  return `<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="${recordCount}">
  <cacheSource type="worksheet"/>
  <cacheFields count="6">
    <cacheField name="R1" databaseField="1"><sharedItems count="10">${r1Items}</sharedItems></cacheField>
    <cacheField name="R2" databaseField="1"><sharedItems count="50">${r2Items}</sharedItems></cacheField>
    <cacheField name="R3" databaseField="1"><sharedItems count="200">${r3Items}</sharedItems></cacheField>
    <cacheField name="C1" databaseField="1"><sharedItems count="12">${c1Items}</sharedItems></cacheField>
    <cacheField name="D1" databaseField="1"><sharedItems count="0"/></cacheField>
    <cacheField name="D2" databaseField="1"><sharedItems count="0"/></cacheField>
  </cacheFields>
</pivotCacheDefinition>`
}

function buildPivotXml(): string {
  // Excel defaults: defaultSubtotal on, grand totals on
  return `<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="BenchPivot1M" cacheId="1" dataCaption="Data">
  <location ref="A3:F100" firstHeaderRow="1" firstDataRow="2" firstDataCol="3"/>
  <pivotFields count="6">
    <pivotField axis="axisRow" showAll="0"/>
    <pivotField axis="axisRow" showAll="0"/>
    <pivotField axis="axisRow" showAll="0"/>
    <pivotField axis="axisCol" showAll="0"/>
    <pivotField dataField="1" showAll="0"/>
    <pivotField dataField="1" showAll="0"/>
  </pivotFields>
  <rowFields count="3"><field x="0"/><field x="1"/><field x="2"/></rowFields>
  <colFields count="1"><field x="3"/></colFields>
  <dataFields count="2">
    <dataField name="Sum of D1" fld="4" subtotal="sum" baseField="0" baseItem="0"/>
    <dataField name="Sum of D2" fld="5" subtotal="sum" baseField="0" baseItem="0"/>
  </dataFields>
</pivotTableDefinition>`
}

function generateRecordChunks(totalRecords: number, batchSize: number): Uint8Array[] {
  const chunks: Uint8Array[] = []
  const encoder = new TextEncoder()
  chunks.push(encoder.encode(`<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${totalRecords}">`))

  const rngR1 = mulberry32(1001)
  const rngR2 = mulberry32(2002)
  const rngR3 = mulberry32(3003)
  const rngC1 = mulberry32(4004)
  const rngD1 = mulberry32(5005)
  const rngD2 = mulberry32(6006)

  let buffer = ''
  for (let i = 0; i < totalRecords; i++) {
    const r1 = Math.floor(rngR1() * 10)
    const r2 = Math.floor(rngR2() * 50)
    const r3 = Math.floor(rngR3() * 200)
    const c1 = Math.floor(rngC1() * 12)
    const d1 = (Math.floor(rngD1() * 10000) / 10) + 0.5
    const d2 = (Math.floor(rngD2() * 5000) / 10) + 1.25

    buffer += `<r><x v="${r1}"/><x v="${r2}"/><x v="${r3}"/><x v="${c1}"/><n v="${d1}"/><n v="${d2}"/></r>`
    if ((i + 1) % batchSize === 0) {
      chunks.push(encoder.encode(buffer))
      buffer = ''
    }
  }
  if (buffer.length > 0) chunks.push(encoder.encode(buffer))
  chunks.push(encoder.encode('</pivotCacheRecords>'))
  return chunks
}

async function runSingleWorker(cache: any, def: any, chunks: Uint8Array[]): Promise<BenchWorkerResponse> {
  const workerUrl = new URL('./pivot-bench-worker.ts', import.meta.url).href
  const worker = new Worker(workerUrl)

  const response = await new Promise<BenchWorkerResponse>((resolve, reject) => {
    worker.onmessage = (e: MessageEvent<BenchWorkerResponse>) => resolve(e.data)
    worker.onerror = (err) => reject(err)
    worker.postMessage({
      cache,
      def,
      chunks,
      warmupRuns: WARMUP_RUNS,
      timedRuns: TIMED_RUNS,
    })
  })

  await worker.terminate()
  return response
}

async function main(): Promise<void> {
  const { runs } = parseArgs()
  const cacheXml = buildCacheXml(RECORD_COUNT)
  const pivotXml = buildPivotXml()

  const cache = parseCacheDefinition(parseXmlOrdered(cacheXml), 1, 'xl/pivotCache/pivotCacheDefinition1.xml', [])
  const def = parsePivotTable(parseXmlOrdered(pivotXml), {
    sheet: 'Sheet1',
    region: { sheet: 'Sheet1', partPath: 'xl/pivotTables/pivotTable1.xml', targetRect: { minRow: 2, maxRow: 99, minCol: 0, maxCol: 5 } },
  }, [])
  if (!def) {
    console.error('Failed to parse synthetic pivot definition')
    process.exit(1)
  }

  const chunks = generateRecordChunks(RECORD_COUNT, BATCH_SIZE)

  const processRuns: Array<{
    run: number
    decodeMs: number
    samples: BenchWorkerResponse['samples']
    medianTotalMs: number
    minTotalMs: number
    maxTotalMs: number
    shape?: BenchWorkerResponse['shape']
  }> = []

  let firstDiagnostic: BenchWorkerResponse['diagnostic'] | undefined
  let firstError: string | undefined

  for (let r = 0; r < runs; r++) {
    const response = await runSingleWorker(cache, def, chunks)
    if (!response.ok) {
      firstDiagnostic = response.diagnostic
      firstError = response.error ?? response.diagnostic?.message ?? 'unknown worker error'
      break
    }

    const totals = response.samples.map(s => s.totalMs).sort((a, b) => a - b)
    const min = totals[0]!
    const max = totals[totals.length - 1]!
    const median = totals[Math.floor(totals.length / 2)]!

    processRuns.push({
      run: r + 1,
      decodeMs: Number(response.decodeMs.toFixed(2)),
      samples: response.samples.map(s => ({
        run: s.run,
        axesMs: Number(s.axesMs.toFixed(2)),
        valuesMs: Number(s.valuesMs.toFixed(2)),
        totalMs: Number(s.totalMs.toFixed(2)),
      })),
      medianTotalMs: Number(median.toFixed(2)),
      minTotalMs: Number(min.toFixed(2)),
      maxTotalMs: Number(max.toFixed(2)),
      shape: response.shape,
    })
  }

  if (firstDiagnostic || firstError) {
    const feature = firstDiagnostic?.feature ?? 'error'
    const result = {
      benchmark: 'pivot-1m-records',
      targetMedianMs: TARGET_MEDIAN_MS,
      records: RECORD_COUNT,
      runsRequested: runs,
      runsCompleted: processRuns.length,
      status: 'error',
      feature,
      error: firstError,
      diagnostic: firstDiagnostic,
      environment: {
        runtime: 'bun',
        bunVersion: Bun.version,
        platform: process.platform,
        arch: process.arch,
      },
    }
    console.log(JSON.stringify(result, null, 2))
    console.log(`GATE ERROR ${feature}`)
    process.exit(0)
  }

  const runMedians = processRuns.map(r => r.medianTotalMs).sort((a, b) => a - b)
  const medianOfMedians = runMedians[Math.floor(runMedians.length / 2)]!

  const result = {
    benchmark: 'pivot-1m-records',
    targetMedianMs: TARGET_MEDIAN_MS,
    records: RECORD_COUNT,
    runsRequested: runs,
    runsCompleted: processRuns.length,
    processRuns,
    summary: {
      medianOfMediansMs: Number(medianOfMedians.toFixed(2)),
      runMedians,
    },
    environment: {
      runtime: 'bun',
      bunVersion: Bun.version,
      platform: process.platform,
      arch: process.arch,
    },
  }

  console.log(JSON.stringify(result, null, 2))

  const passed = medianOfMedians <= TARGET_MEDIAN_MS
  const gateLine = `GATE ${passed ? 'PASS' : 'FAIL'} median-of-medians=${medianOfMedians.toFixed(2)} (≤ ${TARGET_MEDIAN_MS})`
  console.log(gateLine)
  process.exit(0)
}

await main()
