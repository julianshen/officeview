import { afterEach, expect, it, vi } from 'vitest'
import { cleanup, render, waitFor } from '@testing-library/react'
import { OfficeFile } from '../src/components/OfficeFile'
import * as loader from '../src/components/loadOfficeSource'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { valueStatus } from '../src/xlsx/view/values'
import { workbookOf } from './helpers/xlsx-find'
import type { WorkerReply } from '../src/xlsx/view/contracts'
afterEach(() => { cleanup(); vi.restoreAllMocks() })
it('OfficeFile delegates interactive XLSX saved-value notices to the viewer exactly once', async () => {
  const doc = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 1 }] }]); doc.valuePolicy = 'cached'
  doc.diagnostics = [{ kind: 'large-source', feature: 'cached-value-policy', message: 'Showing saved values; formulas are not recalculated' }]
  vi.spyOn(loader, 'loadOfficeSource').mockResolvedValue(doc)
  const ui = render(<OfficeFile data={new Uint8Array([1])} interactiveSheets />)
  await waitFor(() => expect(ui.container.textContent).toMatch(/Showing saved values/))
  expect(ui.container.textContent?.match(/Showing saved values/g)).toHaveLength(1)
})
it('worker completion replies reread progressive source diagnostics instead of retaining the count-free open snapshot', async () => {
  const doc = workbookOf([{ sheetId: 'A', cells: [] }]); doc.valuePolicy = 'cached'
  const base = createObjectWorkbookSource(doc, 'notice'); let formulas: number | undefined
  const source = { ...base, valueStatus: () => valueStatus('cached', formulas), get openDiagnostics() { return [{ code: 'cached-value-policy', message: valueStatus('cached', formulas).message! }] } }
  const kernel = new ViewerKernel({ host: 'in-thread', parse: async () => doc, openSource: async () => source })
  const replies: WorkerReply[] = []
  await kernel.handle({ kind: 'open', docId: 'notice', requestId: 'open', profile: 'D', source: { kind: 'object', document: doc } }, r => replies.push(r))
  expect(replies.at(-1)?.diagnostics[0].message).toContain('formulas are not recalculated')
  formulas = 7
  await kernel.handle({ kind: 'activateSheet', requestId: 'activate', owner: kernel.session('notice')!.workbookOwner(), sheetId: 'A' }, r => replies.push(r))
  expect(replies.at(-1)?.diagnostics).toContainEqual({ code: 'cached-value-policy', message: 'Showing saved values — 7 formulas not recalculated (file too large)' })
  await kernel.session('notice')?.release()
})
