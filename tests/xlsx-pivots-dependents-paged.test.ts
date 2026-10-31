// F4: on the paged path InteractiveSheetDoc passes the shell sheet (pagedWorkbook.ts:76, rows: []) to
// markPagedDependentsStale, which scans sheet.rows for formulas: the notice can never appear.
import { expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { openPagedWorkbook } from '../src/components/xlsx/pagedWorkbook'
import { FORCE_PAGED_LIMITS } from '../src/debug/forcePaged'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { markPagedDependentsStale } from '../src/xlsx/formula/workbook'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import { SEED, load } from './helpers/xlsx-pivots-dependent-fixtures'

function workerHarness() {
  const kernel = new ViewerKernel({ parse: async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)), importLimits: FORCE_PAGED_LIMITS }, { yield: async () => {} })
  const listeners = new Set<(event: MessageEvent) => void>()
  return { port: {
    postMessage(message: unknown, transfer: Transferable[]) {
      const open = message as { kind: string; source?: { kind: string; blob?: Blob } }
      const blob = open.kind === 'open' && open.source?.kind === 'blob' ? open.source.blob : undefined
      const command = structuredClone(blob ? { ...open, source: { kind: 'blob' } } : message, { transfer: transfer as ArrayBuffer[] }) as WorkerCommand
      if (blob && command.kind === 'open') (command as { source: unknown }).source = { kind: 'blob', blob }
      void kernel.handle(command, (out: WorkerReply, replyTransfer) => {
        const data = structuredClone(out, { transfer: replyTransfer })
        for (const listener of [...listeners]) listener({ data } as MessageEvent)
      })
    },
    addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
    terminate: vi.fn(),
  } }
}

it('F4 (Critical): the paged shell sheet the viewer hands markPagedDependentsStale can show the notice', async () => {
  const resident = await load(SEED)
  const formulas = resident.sheets.findIndex(s => s.name === 'Formulas')
  // Resident sheet: the predicate fires (control).
  expect(markPagedDependentsStale(new Set(), formulas, resident.sheets[formulas]!, [])).toEqual(new Set([formulas]))
  const h = workerHarness()
  const doc = await openPagedWorkbook(new Blob([new Uint8Array(readFileSync(SEED))]), { port: h.port as never, profile: 'D', signal: new AbortController().signal })
  expect(doc).toBeTruthy()
  // Paged shell: the same sheet, with its 12 formulas, scans as formula-free.
  expect(doc!.hasFormulaCellsOutsidePivots).toBe(true)
  expect(markPagedDependentsStale(new Set(), formulas, doc!.sheets[formulas]!, [], doc!.hasFormulaCellsOutsidePivots)).toEqual(new Set([formulas]))
})
