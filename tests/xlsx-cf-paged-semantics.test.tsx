/**
 * Task 9 fix round 2, finding 1: paged (worker-backed) semantic text must use
 * prepared CF formats, with localized descriptors and unavailable states.
 * Mounts the real viewer against a real worker kernel (same harness shape as
 * the unmount-GC test): no hand-built scopes, no resident fallbacks.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, render } from '@testing-library/react'
import { OfficeFile } from '../src/components/OfficeFile'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { xlsxBytes } from './helpers/xlsx-package'

vi.mock('../src/xlsx/storage/importSession', async original => {
  const actual = await original<typeof import('../src/xlsx/storage/importSession')>()
  return { ...actual, importLimitsFor: (profile: 'D' | 'M', overrides?: Parameters<typeof actual.importLimitsFor>[1]) => actual.importLimitsFor(profile, { ...overrides, object: { ...overrides?.object, rows: 1 } }) }
})
afterEach(() => { cleanup(); vi.restoreAllMocks() })

const CF_SHEET_DATA = '<row r="1"><c r="A1" s="0"><v>1.25</v></c><c r="B1" s="0"><v>2</v></c></row>'
const CF_STYLES = '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0"/></numFmts><cellXfs count="1"><xf numFmtId="164"/></cellXfs><dxfs count="2"><dxf><numFmt numFmtId="164" formatCode="0.00"/></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>'
const CF_AFTER = '<conditionalFormatting sqref="A1:A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="B1:B1"><cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>'

function workerPort() {
  const kernel = new ViewerKernel({ parse: async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)) }, { yield: async () => {} })
  const listeners = new Set<(event: MessageEvent) => void>()
  return {
    postMessage(message: unknown, transfer: Transferable[]) {
      const command = structuredClone(message, { transfer: transfer as ArrayBuffer[] }) as { kind: string; source?: unknown }
      void kernel.handle(command as never, (out: unknown, replyTransfer) => {
        const data = structuredClone(out, { transfer: replyTransfer as ArrayBuffer[] })
        for (const listener of [...listeners]) listener({ data } as MessageEvent)
      })
    },
    addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
    terminate: vi.fn(),
  }
}

async function mountCf(extra: { afterSheetData?: string } = {}) {
  const bytes = xlsxBytes({ sheetData: CF_SHEET_DATA, styles: CF_STYLES, afterSheetData: extra.afterSheetData ?? CF_AFTER })
  const ui = render(<OfficeFile data={new Blob([new Uint8Array(bytes).buffer])} interactiveSheets viewerWorker={() => workerPort()} />)
  return ui
}

describe('worker-backed semantic text follows CF', () => {
  it('paged CF numFmt formats the accessible name and fill gains a localized descriptor', async () => {
    const ui = await mountCf()
    try {
      const formatted = await ui.findByRole('gridcell', { name: /A1, value 1\.25/ }, { timeout: 30_000 })
      expect(formatted.textContent).toContain('1.25')
      const filled = await ui.findByRole('gridcell', { name: /Conditional formatting applied: fill/ }, { timeout: 30_000 })
      expect(filled.textContent).toContain('2')
    } finally {
      ui.unmount()
    }
  }, 60_000)

  it('a paged unresolved rule announces unavailable, never raw internals', async () => {
    const afterSheetData = '<conditionalFormatting sqref="A1:A1"><cfRule type="expression" dxfId="0" priority="1"><formula>LEN(A1)&gt;100000</formula></cfRule></conditionalFormatting>'
    const ui = await mountCf({ afterSheetData })
    try {
      const cell = await ui.findByRole('gridcell', { name: /A1, value .*unavailable/i }, { timeout: 30_000 })
      expect(cell.textContent).not.toContain('LEN(A1)')
      expect(cell.textContent).not.toContain('expression')
    } finally {
      ui.unmount()
    }
  }, 60_000)
})
