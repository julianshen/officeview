/**
 * S10.2: worker-side operation range detection. A paged header sort used to
 * admit its range on the main thread through ~4,000 worker strip reads (12 s
 * on 1M rows). The `detectRange` read op runs the same detectors next to the
 * worker's store. Differential: on resident and paged sources the worker
 * returns exactly what the main-thread detectors return over the same data —
 * including gaps, header rows, hidden rows, merged cells and refusals.
 */
import { describe, expect, it } from 'vitest'
import type { CacheIdentity, GridSelection, WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { initialIdentity } from '../src/worker/viewerProtocol'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { admitSortRange, detectOperationRange } from '../src/xlsx/view/rangeDetection'
import { StructuralIndex } from '../src/xlsx/view/structure'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { colName, xlsxBytes } from './helpers/xlsx-package'

const quiet = { yield: async () => {} }
let nextId = 0
const rid = () => `rd-${++nextId}`
async function run<K extends WorkerCommand['kind']>(kernel: ViewerKernel, command: Extract<WorkerCommand, { kind: K }>): Promise<WorkerReply<K>> {
  const replies: WorkerReply[] = []
  await kernel.handle(command, out => { replies.push(out) })
  return replies[replies.length - 1] as WorkerReply<K>
}
const PAGED = { object: { rows: 1 } }
const FULL = { r0: 0, r1: 1048575, c0: 0, c1: 16383 }

interface Case { name: string; xml: string; after?: string; active: { row: number; col: number }; ranges?: GridSelection['ranges'] }
const cell = (r: number, c: number, v: number | string) => typeof v === 'number' ? `<c r="${colName(c)}${r + 1}"><v>${v}</v></c>` : `<c r="${colName(c)}${r + 1}" t="inlineStr"><is><t>${v}</t></is></c>`
const table = (rows: number, cols: number, opts: { header?: boolean; gapRow?: number; hidden?: number[]; r0?: number } = {}) => {
  const out: string[] = []
  const r0 = opts.r0 ?? 0
  for (let r = r0; r < r0 + rows; r++) {
    if (r === opts.gapRow) continue
    const cells: string[] = []
    for (let c = 0; c < cols; c++) cells.push(opts.header && r === r0 ? cell(r, c, `H${c}`) : cell(r, c, (r * 7 + c * 3) % 11))
    out.push(`<row r="${r + 1}"${opts.hidden?.includes(r) ? ' hidden="1"' : ''}>${cells.join('')}</row>`)
  }
  return out.join('')
}
const CASES: Case[] = [
  { name: 'header row, current region', xml: table(40, 4, { header: true }), active: { row: 0, col: 2 } },
  { name: 'no header (numeric first row)', xml: table(40, 4), active: { row: 5, col: 1 } },
  { name: 'blank row gap splits regions', xml: table(40, 3, { header: true, gapRow: 20 }), active: { row: 3, col: 0 } },
  { name: 'below the gap', xml: table(40, 3, { header: true, gapRow: 20 }), active: { row: 30, col: 1 } },
  { name: 'hidden rows inside the region', xml: table(30, 3, { header: true, hidden: [4, 5, 17] }), active: { row: 0, col: 1 } },
  { name: 'offset region', xml: table(25, 5, { header: true, r0: 10 }), active: { row: 10, col: 4 } },
  { name: 'merged cells inside the band refuse sort', xml: table(20, 3, { header: true }), after: '<mergeCells count="1"><mergeCell ref="B5:C6"/></mergeCells>', active: { row: 0, col: 0 } },
  { name: 'explicit selection range', xml: table(30, 4, { header: true }), active: { row: 2, col: 1 }, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 2, r1: 12 }] }, cols: [{ c0: 1, c1: 2 }] }] },
  { name: 'blank active cell, single used region', xml: table(15, 2, { header: true }), active: { row: 100, col: 9 } },
]

async function openBoth(bytes: Uint8Array) {
  const parse = async (b: Uint8Array) => parseXlsx(await OfficePackage.load(b))
  const paged = new ViewerKernel({ parse, importLimits: PAGED }, quiet)
  await run(paged, { kind: 'open', requestId: rid(), docId: 'p', source: { kind: 'blob', blob: new Blob([bytes as BlobPart]) }, profile: 'D' })
  const doc = await parse(bytes)
  const resident = new ViewerKernel({ parse }, quiet)
  await run(resident, { kind: 'open', requestId: rid(), docId: 'r', source: { kind: 'object', document: doc }, profile: 'D' })
  return { paged, resident, doc }
}
async function workerDetect(kernel: ViewerKernel, docId: string, kind: 'sort' | 'filter', selection: GridSelection) {
  const owner: CacheIdentity = initialIdentity(docId, '1')
  const out = await run(kernel, { kind: 'read', requestId: rid(), owner, query: { op: 'detectRange', kind, selection } })
  if (out.kind !== 'complete' || out.result.kind !== 'read' || out.result.result.op !== 'detectRange') throw new Error(`detectRange failed: ${JSON.stringify(out).slice(0, 300)}`)
  return out.result.result.detection
}
async function mainDetect(store: Parameters<typeof admitSortRange>[0]['store'], kind: 'sort' | 'filter', selection: GridSelection) {
  const signal = new AbortController().signal
  const structure = new StructuralIndex(await store.queryStructural(FULL, { signal, stamp: { ...store.stamp } }))
  return (kind === 'sort' ? admitSortRange : detectOperationRange)({ store, structure, selection, options: { signal, stamp: { ...store.stamp } } })
}

/** Detections carry their source stamp (docId differs per opened copy); compare the rest. */
const norm = (detection: unknown) => { const { stamp: _stamp, ...rest } = detection as { stamp?: unknown }; return rest }
const stampOf = (detection: unknown) => (detection as { stamp?: { docId: string } }).stamp?.docId

describe('worker detectRange equals the main-thread detectors', () => {
  for (const c of CASES) {
    it(c.name, async () => {
      const bytes = xlsxBytes({ sheetData: c.xml, ...(c.after ? { afterSheetData: c.after } : {}) })
      const { paged, resident, doc } = await openBoth(bytes)
      const selection: GridSelection = { active: c.active, anchor: c.active, ranges: c.ranges ?? [] }
      for (const kind of ['sort', 'filter'] as const) {
        // Main thread over the resident document (today's small-file path) …
        const objectStore = createObjectWorkbookSource(doc, 'main').store(doc.sheets[0].sheetId ?? '1')
        const expected = await mainDetect(objectStore, kind, selection)
        // … and over the paged worker store (what the remote client scanned).
        const pagedStore = paged.session('p')!.source.store('1')
        expect(norm(await mainDetect(pagedStore, kind, selection)), `${kind}: paged main == resident main`).toEqual(norm(expected))
        const fromPaged = await workerDetect(paged, 'p', kind, selection), fromResident = await workerDetect(resident, 'r', kind, selection)
        expect(norm(fromPaged), `${kind}: paged worker`).toEqual(norm(expected))
        expect(norm(fromResident), `${kind}: resident worker`).toEqual(norm(expected))
        if (stampOf(expected)) { expect(stampOf(fromPaged)).toBe('p'); expect(stampOf(fromResident)).toBe('r') }
      }
    })
  }

  it('is owner-checked: a stale owner is refused, not answered', async () => {
    const { paged } = await openBoth(xlsxBytes({ sheetData: table(10, 2, { header: true }) }))
    const stale: CacheIdentity = { ...initialIdentity('p', '1'), viewGen: 99 }
    const out = await run(paged, { kind: 'read', requestId: rid(), owner: stale, query: { op: 'detectRange', kind: 'sort', selection: { active: { row: 0, col: 0 }, anchor: null, ranges: [] } } })
    expect(out.kind).toBe('error')
  })

  it('is cancellable: an aborted read answers cancelled', async () => {
    const { paged } = await openBoth(xlsxBytes({ sheetData: table(200, 3, { header: true }) }))
    const owner = initialIdentity('p', '1'), requestId = rid()
    const replies: WorkerReply[] = []
    const pending = paged.handle({ kind: 'read', requestId, owner, query: { op: 'detectRange', kind: 'sort', selection: { active: { row: 0, col: 0 }, anchor: null, ranges: [] } } }, out => { replies.push(out) })
    await paged.handle({ kind: 'cancel', requestId: rid(), targetRequestId: requestId } as WorkerCommand, () => {})
    await pending
    expect(replies.at(-1)?.kind).toBe('cancelled')
  })
})
