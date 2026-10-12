/**
 * S13.2 item 2 through the product path: <OfficeFile interactiveSheets
 * viewerWorker> opens a paged workbook in the (in-process) resident worker; the
 * sheet's drawing content arrives after its structure, the picture lease gets
 * the drawings and their sheet-local pictures, the painted sheet carries the
 * workbook theme, and what the large-file view cannot carry is a visible notice.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, render, waitFor } from '@testing-library/react'
import { OfficeFile } from '../src/components/OfficeFile'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import type { XlsxDocument } from '../src/xlsx/types'
import { prepareSheetAssets } from '../src/components/xlsx/sheetAssets'
import { buildZip } from './helpers/zip-builder'
import { xlsxEntries } from './helpers/xlsx-package'
import { EMU_PER_PX } from './helpers/xlsx-structure-fixtures'

// Paged on both sides (object path admits one row), as in xlsx-viewer-paged-viewer.
vi.mock('../src/xlsx/storage/importSession', async original => {
  const actual = await original<typeof import('../src/xlsx/storage/importSession')>()
  return { ...actual, importLimitsFor: (profile: 'D' | 'M', overrides?: Parameters<typeof actual.importLimitsFor>[1]) => actual.importLimitsFor(profile, { ...overrides, object: { ...overrides?.object, rows: 1 } }) }
})
const budget = vi.hoisted(() => ({ bytes: undefined as number | undefined }))
vi.mock('../src/xlsx/view/drawingPayload', async original => {
  const actual = await original<typeof import('../src/xlsx/view/drawingPayload')>()
  return { ...actual, drawingPayloadBudget: (profile: 'D' | 'M') => budget.bytes ?? actual.drawingPayloadBudget(profile) }
})
afterEach(() => { cleanup(); vi.restoreAllMocks(); budget.bytes = undefined })

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='), c => c.charCodeAt(0))
const from = (col: number, row: number) => `<xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>`
const ext = (px: number) => `<xdr:ext cx="${px * EMU_PER_PX}" cy="${px * EMU_PER_PX}"/>`
const DRAWING = `<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`
  + `<xdr:oneCellAnchor>${from(0, 1)}${ext(40)}<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Box"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></xdr:spPr></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>`
  + `<xdr:oneCellAnchor>${from(2, 1)}${ext(20)}<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="Picture"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="${R_NS}" r:embed="rImg"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`
  + '</xdr:wsDr>'
function workbook(): Blob {
  const entries = xlsxEntries({
    sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>alpha</t></is></c></row><row r="2"><c r="A2"><v>2</v></c></row>',
    afterSheetData: `<drawing xmlns:r="${R_NS}" r:id="rD1"/>`,
    extra: [
      { name: 'xl/worksheets/_rels/sheet1.xml.rels', data: `<?xml version="1.0"?><Relationships xmlns="${PKG_RELS}"><Relationship Id="rD1" Type="${R_NS}/drawing" Target="../drawings/drawing1.xml"/></Relationships>` },
      { name: 'xl/drawings/drawing1.xml', data: DRAWING },
      { name: 'xl/drawings/_rels/drawing1.xml.rels', data: `<?xml version="1.0"?><Relationships xmlns="${PKG_RELS}"><Relationship Id="rImg" Type="${R_NS}/image" Target="../media/image1.png"/></Relationships>` },
      { name: 'xl/media/image1.png', data: PNG },
    ],
  })
  return new Blob([buildZip(entries).bytes as BlobPart])
}

/** A port with dedicated-worker semantics over a real kernel (structured clone + transfer). */
function workerPort() {
  const kernel = new ViewerKernel({ parse: async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)) }, { yield: async () => {} })
  const listeners = new Set<(event: MessageEvent) => void>()
  const posted: WorkerCommand[] = []
  return {
    posted,
    port: {
      postMessage(message: unknown, transfer: Transferable[]) {
        const open = message as { kind: string; source?: { kind: string; blob?: Blob } }
        const blob = open.kind === 'open' && open.source?.kind === 'blob' ? open.source.blob : undefined
        const command = structuredClone(blob ? { ...open, source: { kind: 'blob' } } : message, { transfer: transfer as ArrayBuffer[] }) as WorkerCommand
        if (blob && command.kind === 'open') command.source = { kind: 'blob', blob }
        posted.push(command)
        void kernel.handle(command, (out: WorkerReply, replyTransfer) => {
          const data = structuredClone(out, { transfer: replyTransfer })
          for (const listener of [...listeners]) listener({ data } as MessageEvent)
        })
      },
      addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
      removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
      terminate: vi.fn(),
    },
  }
}

describe('paged drawings in the worker-backed viewer', () => {
  it('leases the arrived drawings and their pictures for the active sheet', async () => {
    const h = workerPort()
    const docs: XlsxDocument[] = []
    const assetLoader = vi.fn((doc: XlsxDocument, options?: Parameters<typeof prepareSheetAssets>[1]) => { docs.push(doc); return prepareSheetAssets(doc, { ...options, decodeImage: async () => undefined }) })
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} {...({ assetLoader } as object)} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy())
    await waitFor(() => expect(docs.some(doc => doc.sheets[0].drawings?.length === 2)).toBe(true))
    expect(h.posted.filter(command => command.kind === 'read' && command.query.op === 'drawings')).toHaveLength(1)
    const leased = docs.find(doc => doc.sheets[0].drawings?.length === 2)!
    const pic = leased.sheets[0].drawings!.find(node => node.source.element === 'pic')!
    expect([...leased.images![pic.imageIndex!].data]).toEqual([...PNG])
    expect(leased.sheets[0].drawingTheme).toBeDefined()
    expect(ui.container.textContent).not.toMatch(/not shown/)
  })

  it('shows what the large-file view could not carry as a status notice', async () => {
    budget.bytes = 16
    const h = workerPort()
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy())
    await waitFor(() => expect(ui.container.textContent).toMatch(/Drawing part xl\/media\/image1\.png is not shown in the large-file view/))
  })
})
