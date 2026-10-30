// @vitest-environment node
/**
 * Task 5 fix round 1: a paged (worker-backed) workbook indents by its own
 * maximum digit width. The paged shell sheet is what the main-thread layout
 * host lays cells out against, so it must carry the workbook MDW.
 */
import { createCanvas } from 'canvas'
import { describe, expect, it, vi } from 'vitest'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { openPagedWorkbook } from '../src/components/xlsx/pagedWorkbook'
import { pagedWorkbookOf } from '../src/components/pagedDocument'
import { layoutCellBody, paintCellLayout, textMeasurer } from '../src/xlsx/render'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import type { XlsxCell, XlsxSheet } from '../src/xlsx/types'
import { xlsxBytes } from './helpers/xlsx-package'

// One object row forces the paged route for any real sheet.
vi.mock('../src/xlsx/storage/importSession', async original => {
  const actual = await original<typeof import('../src/xlsx/storage/importSession')>()
  return { ...actual, importLimitsFor: (profile: 'D' | 'M', overrides?: Parameters<typeof actual.importLimitsFor>[1]) => actual.importLimitsFor(profile, { ...overrides, object: { ...overrides?.object, rows: 1 } }) }
})

function port() {
  const kernel = new ViewerKernel({ parse: async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)) }, { yield: async () => {} })
  const listeners = new Set<(event: MessageEvent) => void>()
  return {
    postMessage(message: unknown, transfer: Transferable[]) {
      const open = message as { kind: string; source?: { kind: string; blob?: Blob } }
      const blob = open.kind === 'open' && open.source?.kind === 'blob' ? open.source.blob : undefined
      const command = structuredClone(blob ? { ...open, source: { kind: 'blob' } } : message, { transfer: transfer as ArrayBuffer[] }) as WorkerCommand
      if (blob && command.kind === 'open') command.source = { kind: 'blob', blob }
      void kernel.handle(command, (out: WorkerReply, replyTransfer) => {
        const data = structuredClone(out, { transfer: replyTransfer })
        for (const listener of [...listeners]) listener({ data } as MessageEvent)
      })
    },
    addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
    terminate: vi.fn(),
  }
}

const W = 160, H = 30
function leftmostInk(sheet: XlsxSheet, indent: number): number {
  const canvas = createCanvas(W, H), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.fillStyle = '#FFFFFF'
  ctx.fillRect(0, 0, W, H)
  const cell: XlsxCell = { ref: 'A1', row: 0, col: 0, value: 'Hamburg', styleIndex: 0, style: { numFmtId: 0, fontSizePt: 11, horizontal: 'left', ...(indent ? { indent } : {}) } }
  paintCellLayout(ctx, layoutCellBody(ctx, sheet, cell, 0, 0, W, H, textMeasurer(ctx)), { width: W, height: H })
  const data = ctx.getImageData(0, 0, W, H).data
  for (let x = 0; x < W; x++) for (let y = 0; y < H; y++) { const i = (y * W + x) * 4; if (data[i] + data[i + 1] + data[i + 2] < 384) return x }
  return -1
}

describe('paged indent', () => {
  it('a paged workbook whose default font has MDW 8 indents level 2 by 16 px', async () => {
    const rows = Array.from({ length: 5 }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}" s="1"><v>${r}</v></c></row>`).join('')
    // Calibri 12 at 96 dpi: maximum digit width 8 px.
    const styles = '<fonts count="1"><font><sz val="12"/><name val="Calibri"/></font></fonts><cellXfs count="2"><xf numFmtId="0" fontId="0"/><xf numFmtId="0" fontId="0"><alignment horizontal="left" indent="2"/></xf></cellXfs>'
    const doc = await openPagedWorkbook(new Blob([xlsxBytes({ sheetData: rows, styles }) as BlobPart]), { port: port(), profile: 'D', signal: new AbortController().signal })
    expect(doc && pagedWorkbookOf(doc)).toBeTruthy()
    try {
      const shell = doc!.sheets[0]
      expect(leftmostInk(shell, 2) - leftmostInk(shell, 0)).toBe(16)
    } finally { await pagedWorkbookOf(doc!)!.release() }
  })
})
