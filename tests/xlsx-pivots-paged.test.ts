// @vitest-environment node
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { describe, expect, it, vi } from 'vitest'
import type { CacheIdentity, RemotePivotDisplay, SourceStore, WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import { PIVOT_FORMAT_MEMO_BYTES } from '../src/xlsx/view/contracts'
import type { XlsxSheet } from '../src/xlsx/types'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { RemoteRegionStyleCache } from '../src/worker/viewerRemoteRegionStyles'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, renderSheet } from '../src/xlsx/render'
import { ensurePivotEngine } from '../src/xlsx/pivotLoader'
import { ensureTableStyleEngine } from '../src/xlsx/tableStyleLoader'
import { sheetWithRegionStyles } from '../src/worker/viewerRemoteRegionStyles'
import { sheetWithDrawings } from '../src/worker/viewerRemoteDrawings'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { ViewerClient } from '../src/worker/viewerClient'
import { openRemoteWorkbook, RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { pagedSheetGeometry } from '../src/components/xlsx/pagedGeometry'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { validateOverflowProof } from '../src/xlsx/view/viewportOverflow'
import { formatA1 } from '../src/xlsx/view/coordinates'
import { emptyWork, type ViewportPlan } from '../src/xlsx/view/viewport'
import { interruptibleStore } from '../src/worker/viewerAbort'
import { resolveSheetControls } from '../src/xlsx/pivotLoader'
import { PackageError } from '../src/core/package-error'
import * as pagedFormatting from '../src/xlsx/storage/pagedFormatting'
import type { PackageReader } from '../src/xlsx/pivots/load'
import { parseThemeContext } from '../src/drawing/style'
import { semanticsFor } from '../src/xlsx/view/semanticsFor'
import * as stylesExtrasParser from '../src/xlsx/tableStyles/parseStylesExtras'
import { controlIds } from '../src/xlsx/pivots/stub'

const bytes = (name: string) => new Uint8Array(readFileSync(resolve(process.cwd(), `tests/fixtures/xlsx-pivots/${name}.xlsx`)))
function pixels(sheet: Parameters<typeof computeMetrics>[0]): Buffer {
  const metrics = computeMetrics(sheet), canvas = createCanvas(metrics.widthPx, metrics.heightPx)
  renderSheet(sheet, canvas.getContext('2d') as unknown as CanvasRenderingContext2D, metrics)
  return Buffer.from(canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data.buffer)
}


/** The paged path end to end: worker import and pivot-display read, then the main-thread layer build. */
async function pagedPivotLayer(input: Uint8Array<ArrayBuffer>, docId: string): Promise<{ payload: RemotePivotDisplay; layer?: XlsxSheet['regionStyles']; diagnostics: readonly { feature: string }[] }> {
  const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId, profile: 'D' })
  await session.done
  const source = createPagedWorkbookSource(session)
  try {
    const sheet = source.sheets.find(item => item.name === 'Pivot') ?? source.sheets[0]!, store = source.store(sheet.sheetId)
    const payload = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 4 * 1024 * 1024, { set: () => true })
    const owner: CacheIdentity = { docId: store.stamp.docId, sheetId: store.stamp.sheetId, valueGen: store.stamp.valueGen, fontGen: 0, viewGen: 0 }
    const reply = { kind: 'complete', commandKind: 'read', requestId: 'r', owner, diagnostics: [], replyBytes: 64,
      result: { kind: 'read', result: { op: 'pivotDisplay', result: payload } } } as unknown as WorkerReply<'read'>
    const cache = new RemoteRegionStyleCache({
      identity: () => owner, hasRegionStyles: () => false, hasPivotDisplay: () => true, loadState: () => ({ completeness: 'complete' }),
      grant: () => 4 * 1024 * 1024, releaseGrant: () => {}, request: async () => reply, nextRequestId: () => 'r', release: () => {},
      admitLayer: () => () => {}, changed: () => {},
    }, 64 * 1024 * 1024, 'pivotDisplay')
    let state = cache.get(payload.sheetId)
    for (let i = 0; i < 500 && state.state !== 'ready' && state.state !== 'failed'; i++) { await new Promise(resolveTick => setTimeout(resolveTick, 2)); state = cache.get(payload.sheetId) }
    expect(state.state).toBe('ready')
    // The state is the live entry, which close() resets.
    const out = { payload, layer: state.layer, diagnostics: [...state.diagnostics] }
    cache.close()
    return out
  } finally { await source.release(); await session.release() }
}
async function residentPivotSheet(input: Uint8Array): Promise<XlsxSheet> {
  const doc = await parseXlsx(await OfficePackage.load(input))
  await ensureTableStyleEngine(doc.sheets)
  return doc.sheets.find(sheet => sheet.pivotTables?.length)!
}
async function editPackage(name: string, edits: Record<string, (xml: string) => string>): Promise<Uint8Array<ArrayBuffer>> {
  const archive = await JSZip.loadAsync(bytes(name))
  for (const [path, edit] of Object.entries(edits)) archive.file(path, edit(await archive.file(path)!.async('string')))
  return new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
}

describe('paged pivot display descriptors', () => {
  it('memoizes resident control style reparses per package and preferred-name set', async () => {
    const pkg = await OfficePackage.load(bytes('slicer-light1'))
    const sheetPath = 'xl/worksheets/sheet1.xml'
    const ids = controlIds(await pkg.xml(sheetPath))
    const stylesRoot = await pkg.xml('xl/styles.xml')
    if (!stylesRoot) throw new Error('slicer fixture has no styles.xml')
    const spy = vi.spyOn(stylesExtrasParser, 'parseStylesExtras')
    const base = { dxfs: [], tableStyles: new Map(), slicerStyles: new Map() }
    const request = { sheetPath, ...ids, theme: parseThemeContext(), styles: { root: stylesRoot, colors: {} }, stylesExtras: base }
    try {
      await resolveSheetControls(pkg as unknown as PackageReader, request)
      await resolveSheetControls(pkg as unknown as PackageReader, request)
      expect(spy).toHaveBeenCalledTimes(1)
      await resolveSheetControls(pkg as unknown as PackageReader, { ...request, slicerIds: [], timelineIds: [] })
      await resolveSheetControls(pkg as unknown as PackageReader, { ...request, slicerIds: [], timelineIds: [] })
      expect(spy).toHaveBeenCalledTimes(2)
    } finally { spy.mockRestore() }
  })

  it('ships a pivot custom style even when it is defined after the retained style cap', async () => {
    const archive = await JSZip.loadAsync(bytes('pivot-custom-style'))
    const stylesPath = 'xl/styles.xml', stylesText = await archive.file(stylesPath)!.async('string')
    const match = /<tableStyles\b([^>]*)>([\s\S]*?)<\/tableStyles>/.exec(stylesText)!
    const style = match[2]!.match(/<tableStyle\b[\s\S]*?<\/tableStyle>/)![0]
    const unused = Array.from({ length: 4096 }, (_, i) => `<tableStyle name="Unused${i}${'x'.repeat(300)}" pivot="1" table="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle>`).join('')
    archive.file(stylesPath, stylesText.replace(match[0], `<tableStyles${match[1]}>${unused}${style}</tableStyles>`))
    const input = new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: 'pivot-custom-style-cap', profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      const store = source.store(source.sheets[0]!.sheetId)
      const display = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 1024 * 1024, { set: () => true })
      expect((display as unknown as { pivotStyles?: Array<{ name: string }> }).pivotStyles?.map(item => item.name)).toContain('OfficeViewPivot')
    } finally { await source.release(); await session.release() }
  })

  it('prioritizes and applies a referenced custom slicer style on the paged drawings path', async () => {
    const archive = await JSZip.loadAsync(bytes('slicer-light1'))
    const stylesPath = 'xl/styles.xml', styles = await archive.file(stylesPath)!.async('string')
    const x14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
    const unused = Array.from({ length: 4097 }, (_, i) => `<x14:slicerStyle name="UnusedSlicer${i}"><x14:slicerStyleElements/></x14:slicerStyle>`).join('')
    const extension = `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"><x14:dxfs xmlns:x14="${x14}" count="1"><x14:dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></x14:dxf></x14:dxfs><x14:slicerStyles xmlns:x14="${x14}" count="4098" defaultSlicerStyle="LateSlicer">${unused}<x14:slicerStyle name="LateSlicer"><x14:slicerStyleElements><x14:slicerStyleElement type="unselectedItemWithData" dxfId="0"/></x14:slicerStyleElements></x14:slicerStyle></x14:slicerStyles></ext></extLst>`
    archive.file(stylesPath, styles.replace('</styleSheet>', `${extension}</styleSheet>`))
    const slicerPath = Object.keys(archive.files).find(path => /^xl\/slicers\/slicer[^/]*\.xml$/.test(path))!
    archive.file(slicerPath, (await archive.file(slicerPath)!.async('string')).replace('style="SlicerStyleLight1"', 'style="LateSlicer"'))
    const input = new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: 'custom-slicer-style', profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      const sheet = source.sheets[0]!
      const store = source.store(sheet.sheetId)
      expect(session.workbook.tableStyleExtras?.slicerStyles.has('LateSlicer')).toBe(true)
      const drawings = await store.readDrawings!({ signal: new AbortController().signal, stamp: store.stamp }, 'D', { set: () => true })
      expect(drawings.drawings.some(drawing => drawing.content?.kind === 'slicer')).toBe(true)
      const resident = await parseXlsx(await OfficePackage.load(input), { valuePolicy: 'cached' })
      await ensurePivotEngine(resident.sheets)
      const residentSheet = resident.sheets[0]!
      const pagedSheet = { ...residentSheet, drawings: [...drawings.drawings] }
      expect(pixels(pagedSheet)).toEqual(pixels(residentSheet))
      expect(pixels(pagedSheet)).not.toEqual(pixels({ ...residentSheet, drawings: [] }))
    } finally { await source.release(); await session.release() }
  })

  it('prioritizes and paints a custom timeline style beyond the cap with resident parity', async () => {
    const archive = await JSZip.loadAsync(bytes('timeline-light1-months'))
    const stylesPath = 'xl/styles.xml', styles = await archive.file(stylesPath)!.async('string')
    const x14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
    const unused = Array.from({ length: 4097 }, (_, i) => `<x14:slicerStyle name="UnusedTimeline${i}"><x14:slicerStyleElements/></x14:slicerStyle>`).join('')
    const extension = `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"><x14:dxfs xmlns:x14="${x14}" count="1"><x14:dxf><fill><patternFill patternType="solid"><fgColor rgb="FF00AA00"/></patternFill></fill></x14:dxf></x14:dxfs><x14:slicerStyles xmlns:x14="${x14}" count="4098" defaultSlicerStyle="LateTimeline">${unused}<x14:slicerStyle name="LateTimeline"><x14:slicerStyleElements><x14:slicerStyleElement type="selectedItemWithData" dxfId="0"/></x14:slicerStyleElements></x14:slicerStyle></x14:slicerStyles></ext></extLst>`
    const withDxfs = styles.replace(/<dxfs count="0"\s*\/>/, '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>')
    const withTable = withDxfs.replace(/<tableStyles\b([^>]*)\s*\/>/, (_tag, attrs: string) => `<tableStyles${attrs}><tableStyle name="LateTimeline" pivot="1" table="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle></tableStyles>`)
    archive.file(stylesPath, withTable.replace('</styleSheet>', `${extension}</styleSheet>`))
    const timelinePath = Object.keys(archive.files).find(path => /^xl\/timelines\/timeline[^/]*\.xml$/.test(path))!
    const timelineText = await archive.file(timelinePath)!.async('string')
    archive.file(timelinePath, timelineText.replace(/<timeline\b([^>]*?)(\/?)>/, (_tag, attrs: string, slash: string) => `<timeline${/\bstyle=/.test(attrs) ? attrs.replace(/\bstyle="[^"]*"/, 'style="LateTimeline"') : `${attrs} style="LateTimeline"`}${slash}>`))
    const input = new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: 'custom-timeline-style-cap', profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      expect(session.workbook.tableStyleExtras?.slicerStyles.has('LateTimeline')).toBe(true)
      const store = source.store(source.sheets[0]!.sheetId)
      const drawings = await store.readDrawings!({ signal: new AbortController().signal, stamp: store.stamp }, 'D', { set: () => true })
      expect(drawings.diagnostics).not.toContainEqual(expect.objectContaining({ feature: 'timeline-style-unknown' }))
      const resident = await parseXlsx(await OfficePackage.load(input), { valuePolicy: 'cached' })
      await ensurePivotEngine(resident.sheets)
      const residentSheet = resident.sheets[0]!
      const pagedSheet = { ...residentSheet, drawings: [...drawings.drawings] }
      expect(pixels(pagedSheet)).toEqual(pixels(residentSheet))
      expect(pixels(pagedSheet)).not.toEqual(pixels({ ...residentSheet, drawings: [] }))
    } finally { await source.release(); await session.release() }
  })

  it('forwards the active profile through the interruptible store wrapper', async () => {
    const stamp = { docId: 'profile-forward', sheetId: 'sheet-1', valueGen: 0 }
    let received: string | undefined
    const payload: RemotePivotDisplay = { sheetId: stamp.sheetId, stamp, pivotTables: [], pivotStyles: [], diagnostics: [] }
    const store = {
      stamp, extent: { rows: 1, cols: 1 }, valuePolicy: 'cached',
      readPivotDisplay: async (_options: unknown, _grantBytes: number, _memory: unknown, profile?: string) => { received = profile; return payload },
      release: async () => {},
    } as unknown as SourceStore
    const wrapped = interruptibleStore(store)
    await wrapped.readPivotDisplay!({ signal: new AbortController().signal, stamp }, 4096, undefined, 'M')
    expect(received).toBe('M')
  })

  it('exposes the receiver-granted pivot display read for a sheet with pivot or control metadata', async () => {
    const input = bytes('pivot-layout-compact')
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: 'pivot-display-contract', profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      const sheet = source.sheets[0]!
      expect(source.hasPivotDisplay?.(sheet.sheetId)).toBe(true)
      const store = source.store(sheet.sheetId)
      expect(store.readPivotDisplay).toBeTypeOf('function')
      const result = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 1024 * 1024, { set: () => true })
      expect(result.pivotTables.length).toBeGreaterThan(0)
      expect('buttons' in result).toBe(false)
      expect(result.pivotStyles).toEqual([])
      expect('slicerStyles' in result).toBe(false)
      expect(result.diagnostics).not.toContainEqual(expect.objectContaining({ feature: 'pivot-display-reply-limit' }))
    } finally {
      await source.release()
      await session.release()
    }
  })

  it('sizes pivot-display grants by pivot tables and ignores control count', async () => {
    const grants: number[] = []
    for (const name of ['pivot-layout-compact', 'slicer-light1'] as const) {
      const input = bytes(name)
      const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: `grant-count-${name}`, profile: 'D' })
      await session.done
      const source = createPagedWorkbookSource(session)
      const kernel = new ViewerKernel({ parse: async () => { throw new Error('expected supplied paged source') }, openSource: async () => source }, { yield: async () => {} })
      const inProcess = connectInProcess(kernel), posted: WorkerCommand[] = []
      const port = { ...inProcess, postMessage(message: unknown, transfer: Transferable[]) { posted.push(message as WorkerCommand); inProcess.postMessage(message, transfer) } }
      try {
        const client = new ViewerClient(port, `grant-count-${name}`)
        const opened = await openRemoteWorkbook(client, `grant-count-${name}`, new Blob([input]), 'D')
        if (!('layouts' in opened)) throw new Error('paged workbook did not open')
        const remote = new RemoteWorkbook(client, opened), sheetId = remote.sheets[0]!.sheetId
        for (let i = 0; i < 200 && remote.loadState(sheetId).completeness !== 'complete'; i++) await new Promise(resolveTick => setTimeout(resolveTick, 2))
        expect(remote.layout(sheetId).hasPivotDisplay).toBe(true)
        for (let i = 0; i < 200 && !posted.some(command => command.kind === 'read' && command.query.op === 'pivotDisplay'); i++) {
          remote.pivotDisplay(sheetId)
          await new Promise(resolveTick => setTimeout(resolveTick, 2))
        }
        const command = posted.find(command => command.kind === 'read' && command.query.op === 'pivotDisplay')
        if (command?.kind !== 'read' || command.query.op !== 'pivotDisplay') throw new Error('pivot display request was not sent')
        grants.push(command.query.grantBytes)
        await remote.close()
      } finally { await source.release(); await session.release() }
    }
    expect(grants).toHaveLength(2)
    expect(grants[0]).toBe(grants[1])
  })

  it.each(['slicer-light1', 'timeline-light1-months'] as const)('keeps %s style definitions out of pivot-display payloads', async name => {
    const input = bytes(name)
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: `pivot-display-${name}`, profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      const sheet = source.sheets[0]!
      expect(source.hasPivotDisplay?.(sheet.sheetId)).toBe(true) // This fixture also contains a pivot table.
      const store = source.store(sheet.sheetId)
      const result = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 1024 * 1024, { set: () => true })
      expect('slicerStyles' in result).toBe(false)
      expect(result.pivotStyles.map(style => style.name)).not.toContain('SlicerStyleLight1')
      expect(result.diagnostics).not.toContainEqual(expect.objectContaining({ feature: 'pivot-display-reply-limit' }))
    } finally { await source.release(); await session.release() }
  })

  it.each(['slicer-light1', 'table-medium-2'] as const)('does not schedule pivot-display reads for control/table-only sheet %s', async name => {
    const input = bytes(name)
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: `pivot-control-only-${name}`, profile: 'D' })
    await session.done
    if (name === 'slicer-light1') {
      // This fixture has a pivot table as well as controls; remove only that
      // metadata to model a controls-only sheet while preserving drawing parts.
      for (const sheet of session.workbook.sheets) Object.assign(sheet.metadata, { pivotRegions: [], pivotDiagnostics: [] })
    }
    const source = createPagedWorkbookSource(session)
    try {
      const sheet = source.sheets[0]!
      expect(source.hasPivotDisplay?.(sheet.sheetId)).toBe(false)
      const kernel = new ViewerKernel({ parse: async () => { throw new Error('expected supplied paged source') }, openSource: async () => source }, { yield: async () => {} })
      const inProcess = connectInProcess(kernel), posted: WorkerCommand[] = []
      const port = { ...inProcess, postMessage(message: unknown, transfer: Transferable[]) { posted.push(message as WorkerCommand); inProcess.postMessage(message, transfer) } }
      const client = new ViewerClient(port, `pivot-control-only-${name}`)
      const opened = await openRemoteWorkbook(client, `pivot-control-only-${name}`, new Blob([input]), 'D')
      if (!('layouts' in opened)) throw new Error('paged workbook did not open')
      const remote = new RemoteWorkbook(client, opened)
      expect(remote.layout(sheet.sheetId).hasRegionStyles).toBe(name === 'table-medium-2')
      for (let i = 0; i < 200 && remote.loadState(sheet.sheetId).completeness !== 'complete'; i++) await new Promise(resolveTick => setTimeout(resolveTick, 2))
      remote.pivotDisplay(sheet.sheetId)
      if (name === 'table-medium-2') {
        for (let i = 0; i < 200 && !posted.some(command => command.kind === 'read' && command.query.op === 'regionStyles'); i++) { remote.regionStyles(sheet.sheetId); await new Promise(resolveTick => setTimeout(resolveTick, 2)) }
        expect(posted.filter(command => command.kind === 'read' && command.query.op === 'regionStyles')).toHaveLength(1)
      } else await new Promise(resolveTick => setTimeout(resolveTick, 10))
      expect(posted.filter(command => command.kind === 'read' && command.query.op === 'pivotDisplay')).toHaveLength(0)
      await remote.close()
    } finally { await source.release(); await session.release() }
  })

  it('retains a referenced custom table style after more than 2 MiB of unused definitions', async () => {
    const archive = await JSZip.loadAsync(bytes('table-medium-2'))
    const tablePath = Object.keys(archive.files).find(path => /^xl\/tables\/table\d+\.xml$/.test(path))!
    const tableText = await archive.file(tablePath)!.async('string')
    archive.file(tablePath, tableText.replace('TableStyleMedium2', 'LateReferenced'))
    const stylesPath = 'xl/styles.xml', stylesText = await archive.file(stylesPath)!.async('string')
    const unused = Array.from({ length: 12_000 }, (_, i) => `<tableStyle name="UnusedStyle${i}" pivot="1" table="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle>`).join('')
    const used = '<tableStyle name="LateReferenced" pivot="1" table="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle>'
    const mutated = stylesText
      .replace('<dxfs count="0"/>', '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>')
      .replace(/<tableStyles count="0"([^>]*)\/>/, `<tableStyles count="12001"$1>${unused}${used}</tableStyles>`)
    archive.file(stylesPath, mutated)
    const input = new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: 'referenced-style-priority', profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      expect(session.workbook.tableStyleDefs?.has('LateReferenced')).toBe(true)
      expect(session.workbook.tableStyleDefs?.get('LateReferenced')?.elements.wholeTable?.dxf.fill?.fg).toEqual({ rgb: 'FFFF0000' })
      const sheet = source.sheets[0]!, store = source.store(sheet.sheetId)
      const display = await store.readRegionStyles!({ signal: new AbortController().signal, stamp: store.stamp }, 1024 * 1024, { set: () => true })
      expect(display.tables[0]?.def.name).toBe('LateReferenced')
      expect(display.tables[0]?.def.elements.wholeTable?.dxf.fill?.fg).toEqual({ rgb: 'FFFF0000' })
    } finally { await source.release(); await session.release() }
  })

  it('does not expose a completed sheet while referenced styles are being reprioritized', async () => {
    const archive = await JSZip.loadAsync(bytes('pivot-layout-compact'))
    const pivotPath = Object.keys(archive.files).find(path => /^xl\/pivotTables\/pivotTable[^/]*\.xml$/.test(path))!
    const pivotText = await archive.file(pivotPath)!.async('string')
    archive.file(pivotPath, pivotText.replace(/(<pivotTableStyleInfo\b[^>]*\bname=")PivotStyleLight16/, '$1LateReferenced'))
    const stylesPath = 'xl/styles.xml', stylesText = await archive.file(stylesPath)!.async('string')
    const unused = Array.from({ length: 12_000 }, (_, i) => `<tableStyle name="UnusedStyle${i}" pivot="1" table="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle>`).join('')
    const used = '<tableStyle name="LateReferenced" pivot="1" table="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle>'
    archive.file(stylesPath, stylesText
      .replace('<dxfs count="0"/>', '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>')
      .replace(/<tableStyles count="0"([^>]*)\/>/, `<tableStyles count="12001"$1>${unused}${used}</tableStyles>`))
    const input = new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
    let entered!: () => void, release!: () => void
    const priorityStarted = new Promise<void>(resolveStarted => { entered = resolveStarted })
    const priorityGate = new Promise<void>(resolveGate => { release = resolveGate })
    const original = pagedFormatting.workbookFormattingWithColorContext
    const spy = vi.spyOn(pagedFormatting, 'workbookFormattingWithColorContext').mockImplementation(async (...args) => {
      const result = await original(...args)
      if (args[3]?.has('LateReferenced')) { entered(); await priorityGate }
      return result
    })
    let session: Awaited<ReturnType<typeof openImportSession>> | undefined
    let source: ReturnType<typeof createPagedWorkbookSource> | undefined
    try {
      session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: 'priority-completion-race', profile: 'D' })
      await priorityStarted
      const sheet = session.workbook.sheets.find(item => item.name === 'Pivot')!
      source = createPagedWorkbookSource(session)
      const store = source.store(sheet.sheetId)
      const duringPriority = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 1024 * 1024, { set: () => true })
      expect(duringPriority.diagnostics).toContainEqual(expect.objectContaining({ feature: 'pivot-display-read-failed', message: expect.stringContaining('structure is complete') }))
      release()
      await session.done
      const afterPriority = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 1024 * 1024, { set: () => true })
      expect(afterPriority.pivotStyles.map(style => style.name)).toContain('LateReferenced')
    } finally {
      release()
      spy.mockRestore()
      if (source) await source.release()
      if (session) await session.release()
    }
  })

  it.each(['pivot-layout-compact', 'pivot-custom-style', 'slicer-light1', 'timeline-light1-months'])('matches resident pixels through paged worker descriptors and drawings for %s', async name => {
    const input = bytes(name)
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: `pivot-pixels-${name}`, profile: 'D' })
    await session.done
    const pagedSource = createPagedWorkbookSource(session)
    const kernel = new ViewerKernel({ parse: async () => { throw new Error('expected supplied paged source') }, openSource: async () => pagedSource }, { yield: async () => {} })
    const inProcess = connectInProcess(kernel), posted: WorkerCommand[] = []
    const port = { ...inProcess, postMessage(message: unknown, transfer: Transferable[]) { posted.push(message as WorkerCommand); inProcess.postMessage(message, transfer) } }
    const client = new ViewerClient(port, `pivot-pixels-${name}`)
    const opened = await openRemoteWorkbook(client, `pivot-pixels-${name}`, new Blob([input]), 'D')
    if (!('layouts' in opened)) throw new Error('paged workbook did not open')
    const remote = new RemoteWorkbook(client, opened), sheetId = remote.sheets[0]!.sheetId
    try {
      for (let i = 0; i < 100 && remote.loadState(sheetId).completeness !== 'complete'; i++) await new Promise(resolveTick => setTimeout(resolveTick, 2))
      remote.pivotDisplay(sheetId)
      for (let i = 0; i < 200 && !['ready', 'failed'].includes(remote.pivotDisplay(sheetId).state); i++) await new Promise(resolveTick => setTimeout(resolveTick, 2))
      const display = remote.pivotDisplay(sheetId)
      expect(display.state, 'worker descriptor cache').toBe('ready')
      expect(posted.filter(command => command.kind === 'read' && command.query.op === 'pivotDisplay')).toHaveLength(1)
      remote.drawings(sheetId)
      for (let i = 0; i < 200 && !['ready', 'failed'].includes(remote.drawings(sheetId).state); i++) await new Promise(resolveTick => setTimeout(resolveTick, 2))
      expect(remote.drawings(sheetId).state, 'worker drawing/control cache').toBe('ready')
      const resident = await parseXlsx(await OfficePackage.load(input), { valuePolicy: 'cached' })
      const residentSheet = resident.sheets[0]!
      await ensureTableStyleEngine([residentSheet])
      await ensurePivotEngine([residentSheet])
      // Match pagedWorkbook.ts: its sheet is a metadata shell with no resident
      // pivotRegions/cells; pivot tables arrive through the worker display read.
      const shell: XlsxSheet = { name: remote.sheets[0]!.name, sheetId, workbookIndex: 0, rows: [], cols: [], merges: [], mergeRanges: [], semantics: semanticsFor(opened.formatting?.dateSystem ?? '1900'), ...(opened.formatting?.maxDigitWidthPx !== undefined ? { maxDigitWidthPx: opened.formatting.maxDigitWidthPx } : {}) }
      const pagedSheet = sheetWithRegionStyles(sheetWithDrawings(shell, remote.drawings(sheetId)), display)
      expect(pagedSheet.pivotRegions).toBeUndefined()
      if (name === 'pivot-layout-compact' && pagedSheet.pivotTables?.length) {
        // Page/export painting consumes worker pivot definitions without
        // resident pivotRegions, just like the interactive shell does.
        const pagedPrintSheet = sheetWithRegionStyles({ ...residentSheet, pivotRegions: undefined, pivotTables: undefined } as XlsxSheet, display)
        expect(pagedPrintSheet.pivotRegions).toBeUndefined()
        for (const table of pagedPrintSheet.pivotTables ?? []) table.printDrill = true
        const withoutPrintButtons = { ...pagedPrintSheet, pivotTables: undefined }
        expect(pixels(pagedPrintSheet).equals(pixels(withoutPrintButtons)), 'printDrill button pixels require pivotTables even without pivotRegions').toBe(false)
      }
      // Compare a real worker viewport plan painted by the interactive compositor.
      const window = { box: { x: 0, y: 0, width: 480, height: 360 }, zoom: 1, deviceScale: 1 }
      const geometry = pagedSheetGeometry(remote, sheetId, createIdentityOrder(), visibilityFromHidden(createIntervalSet([])))
      if (geometry.kind !== 'ready') throw new Error('paged viewport geometry is not ready')
      const frame = remote.currentPageFrame, tile = `pivot-parity:${sheetId}`, pageLimit = remote.pageLimits.tile
      expect(remote.admitPages(frame, tile, pageLimit, false, false)).toBe(true)
      expect(remote.admitPages(frame, tile, pageLimit, false)).toBe(true)
      const viewport = await remote.viewport(remote.identity(sheetId), window, new AbortController().signal, 'tile')
      expect(typeof viewport).toBe('object')
      if (typeof viewport !== 'object' || viewport.kind !== 'complete') throw new Error('paged worker viewport did not complete')
      const result = viewport.result
      if (result.content.kind !== 'detail') throw new Error('paged worker viewport was not detailed')
      const proof = validateOverflowProof(result, remote.identity(sheetId), window, 'D', 'tile', remote.columns(sheetId))
      const cells = result.content.cells.flatMap(item => {
        if (item.kind !== 'cell' && item.kind !== 'unavailable') return []
        const overflow = proof.get(item.source.row * 16384 + item.source.col)!
        const style = remote.styleOf(item.cell.styleIndex)
        return [{ source: item.source, x: item.bounds.x, y: item.bounds.y, width: item.bounds.width, height: item.bounds.height,
          overflow: { left: overflow.left, right: overflow.right }, cell: { ...item.cell, ref: formatA1(item.source), ...(style ? { style } : {}) } }]
      })
      const axes = remote.geometry(sheetId), extent = { width: axes.cols.totalCss(), height: axes.rows.totalCss() }
      const plan: Extract<ViewportPlan, { kind: 'detail' }> = { kind: 'detail', cells, merges: (result.merges ?? []).map(merge => ({ range: merge.range, ...merge.bounds })), drawings: [...remote.drawings(sheetId).drawings], rowLines: result.rowLines ?? [], colLines: result.colLines ?? [], extent, diagnostics: [], coarse: result.coarse ?? [], work: result.work ?? emptyWork() }
      const paint = (sheet: typeof pagedSheet, drawings: typeof plan.drawings) => {
        const canvas = createCanvas(window.box.width, window.box.height)
        const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
        renderSheetViewport(sheet, ctx, geometry, window, undefined, { preparedPlan: { ...plan, drawings } })
        return Buffer.from(ctx.getImageData(0, 0, canvas.width, canvas.height).data.buffer)
      }
      const pagedPixels = paint(pagedSheet, remote.drawings(sheetId).drawings)
      const residentPixels = paint(residentSheet, residentSheet.drawings ?? [])
      expect(pagedPixels.equals(residentPixels), 'paged viewport raster matches resident').toBe(true)
      expect(pagedPixels).not.toEqual(Buffer.alloc(pagedPixels.length))
      const withoutLayer = { ...residentSheet, pivotRegions: undefined, pivotTables: undefined, regionStyles: undefined, drawings: [] }
      expect(pagedPixels.equals(paint(withoutLayer, []))).toBe(false)
      remote.mainLedger.settle(viewport)
    } finally { await remote.close(); await pagedSource.release(); await session.release() }
  })

  it('returns a typed pivot-display-reply-limit and exact retry size when the receiver grant is too small', async () => {
    const input = bytes('pivot-layout-compact')
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: 'pivot-display-small-grant', profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      const sheet = source.sheets[0]!, store = source.store(sheet.sheetId)
      const result = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 1, { set: () => true })
      expect(result.pivotTables).toEqual([])
      expect(result.requiredBytes).toBeGreaterThan(1)
      expect(result.diagnostics).toContainEqual(expect.objectContaining({ feature: 'pivot-display-reply-limit' }))
    } finally { await source.release(); await session.release() }
  })

  it('does not read or leak another sheet over-cap pivot parts into this sheet display', async () => {
    const input = bytes('pivot-layout-compact')
    const archive = await JSZip.loadAsync(input)
    const pivotPath = Object.keys(archive.files).find(path => /^xl\/pivotTables\/pivotTable[^/]*\.xml$/.test(path))!
    expect(archive.file(pivotPath)).toBeTruthy()
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: 'pivot-sheet-isolation', profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      const data = source.sheets.find(sheet => sheet.name === 'Data')!
      const pivot = source.sheets.find(sheet => sheet.name === 'Pivot')!
      const rangeSource = session.workbook.parts.src
      const entry = session.workbook.parts.inventory.find(pivotPath)!
      expect(entry.uncompressedSize).toBeLessThanOrEqual(64 * 1024)
      let touchedPivot = false
      const read = rangeSource.read.bind(rangeSource)
      rangeSource.read = async (start, end, signal) => {
        if (start < entry.rangeEnd && end > entry.localHeaderOffset) touchedPivot = true
        return read(start, end, signal)
      }
      const dataStore = source.store(data.sheetId)
      const result = await dataStore.readPivotDisplay!({ signal: new AbortController().signal, stamp: dataStore.stamp }, 4 * 1024 * 1024, { set: () => true })
      expect(result.pivotTables).toHaveLength(0)
      expect(touchedPivot).toBe(false)
      const pivotStore = source.store(pivot.sheetId)
      const pivotResult = await pivotStore.readPivotDisplay!({ signal: new AbortController().signal, stamp: pivotStore.stamp }, 4 * 1024 * 1024, { set: () => true })
      expect(pivotResult.pivotTables.length).toBeGreaterThan(0)
      expect(touchedPivot).toBe(true)
    } finally { await source.release(); await session.release() }
  })

  it('retains cache-level pivot diagnostics for this sheet', async () => {
    const archive = await JSZip.loadAsync(bytes('pivot-layout-compact'))
    const cachePath = Object.keys(archive.files).find(path => /^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/.test(path))!
    archive.file(cachePath, '<pivotCacheDefinition><broken></pivotCacheDefinition>')
    const input = new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
    const resident = await parseXlsx(await OfficePackage.load(input), { valuePolicy: 'cached' })
    await ensurePivotEngine(resident.sheets)
    const residentDiagnostics = resident.sheets[0]!.diagnostics?.filter(diagnostic => diagnostic.feature === 'pivot-cache-unreadable') ?? []
    expect(residentDiagnostics.some(diagnostic => diagnostic.message.includes(cachePath))).toBe(true)
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: 'pivot-cache-diagnostic', profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      const sheet = source.sheets.find(item => item.name === 'Pivot')!, store = source.store(sheet.sheetId)
      const result = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 1024 * 1024, { set: () => true })
      expect(result.diagnostics.some(diagnostic => diagnostic.feature === 'pivot-cache-unreadable' && diagnostic.message.includes(cachePath))).toBe(true)
    } finally { await source.release(); await session.release() }
  })

  it.each([
    ['pivot-layout-compact', /xl\/pivotTables\/pivotTable[^/]*\.xml$/i, /<pivotTableDefinition\b/],
    ['pivot-layout-compact', /xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/i, /<pivotCacheDefinition\b/],
    ['slicer-light1', /xl\/slicerCaches\/slicerCache[^/]*\.xml$/i, /<slicerCacheDefinition\b/],
    ['timeline-light1-months', /xl\/timelines\/timeline[^/]*\.xml$/i, /<timeline\b/],
  ] as const)('reports pivot-part-too-large for over-cap package parts (%s)', async (fixture, pathPattern, rootPattern) => {
    const archive = await JSZip.loadAsync(bytes(fixture))
    const path = Object.keys(archive.files).find(name => pathPattern.test(name))
    expect(path, 'fixture part').toBeTruthy()
    const original = await archive.file(path!)!.async('string')
    expect(original).toMatch(rootPattern)
    const rootEnd = original.indexOf('>')
    archive.file(path!, `${original.slice(0, rootEnd + 1)}<!--${' '.repeat(70 * 1024)}-->${original.slice(rootEnd + 1)}`)
    const inflated = new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
    // A small injected cap (Ruling AP: the profile cap is 8 MiB desktop, 4 MiB mobile).
    const session = await openImportSession({ kind: 'blob', blob: new Blob([inflated]) }, { docId: `pivot-cap-${fixture}`, profile: 'D', limits: { paged: { pivotPartBytes: 64 * 1024 } } })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      const sheet = source.sheets[0]!, store = source.store(sheet.sheetId)
      const options = { signal: new AbortController().signal, stamp: store.stamp }
      const result = /slicerCaches|timelines/.test(pathPattern.source)
        ? await store.readDrawings!(options, 'D', { set: () => true })
        : await store.readPivotDisplay!(options, 4 * 1024 * 1024, { set: () => true })
      const pivotDiagnostics = result.diagnostics.filter(diagnostic => ('feature' in diagnostic && diagnostic.feature === 'pivot-part-too-large' || 'code' in diagnostic && diagnostic.code === 'pivot-part-too-large')
        && (('partPath' in diagnostic && diagnostic.partPath?.endsWith(path!.split('/').pop()!)) || ('partId' in diagnostic && diagnostic.partId?.endsWith(path!.split('/').pop()!))))
      expect(pivotDiagnostics).toHaveLength(1)
    } finally { await source.release(); await session.release() }
  })

  it.each([
    ['pivot-layout-compact', /xl\/pivotTables\/pivotTable[^/]*\.xml$/i],
    ['pivot-layout-compact', /xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/i],
  ] as const)('reads pivot parts past 64 KiB under the profile cap (Ruling AP; %s)', async (fixture, pathPattern) => {
    const archive = await JSZip.loadAsync(bytes(fixture))
    const path = Object.keys(archive.files).find(name => pathPattern.test(name))!
    const original = await archive.file(path)!.async('string')
    // Grown at its end, as a long saved layout or a large shared-item list grows a part: the location stays near the start.
    const close = original.lastIndexOf('</')
    archive.file(path, `${original.slice(0, close)}<!--${' '.repeat(1024 * 1024)}-->${original.slice(close)}`)
    const inflated = new Uint8Array(await archive.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
    for (const profile of ['D', 'M'] as const) {
      const session = await openImportSession({ kind: 'blob', blob: new Blob([inflated]) }, { docId: `pivot-profile-cap-${profile}`, profile })
      await session.done
      const source = createPagedWorkbookSource(session)
      try {
        const sheet = source.sheets.find(item => item.name !== 'Data') ?? source.sheets[0]!, store = source.store(sheet.sheetId)
        const result = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 4 * 1024 * 1024, { set: () => true }, profile)
        expect(result.diagnostics.filter(item => item.feature === 'pivot-part-too-large')).toEqual([])
        expect(result.pivotTables).toHaveLength(1)
      } finally { await source.release(); await session.release() }
    }
  }, 60_000)

  it('keeps the receiver grant as the bound on what a large pivot part sends to main', async () => {
    const archive = await JSZip.loadAsync(bytes('pivot-layout-compact'))
    const path = Object.keys(archive.files).find(name => /pivotTables\/pivotTable1\.xml$/.test(name))!
    const original = await archive.file(path)!.async('string')
    const close = original.lastIndexOf('</')
    archive.file(path, `${original.slice(0, close)}<!--${' '.repeat(1024 * 1024)}-->${original.slice(close)}`)
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(await archive.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))]) }, { docId: 'pivot-profile-grant', profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      const store = source.store(source.sheets[0]!.sheetId)
      const result = await store.readPivotDisplay!({ signal: new AbortController().signal, stamp: store.stamp }, 512, { set: () => true })
      expect(result.pivotTables).toEqual([])
      expect(result.diagnostics.map(item => item.feature)).toContain('pivot-display-reply-limit')
      expect(result.requiredBytes).toBeGreaterThan(512)
    } finally { await source.release(); await session.release() }
  })


  it.each([
    ['table-medium-2', 0],
    ['pivot-layout-compact', 1],
    ['pivot-custom-style', 1],
    ['slicer-light1', 1],
  ] as const)('parses the styles extras in the paged worker only when a sheet needs them (I4; %s)', async (name, calls) => {
    const spy = vi.spyOn(stylesExtrasParser, 'parseStylesExtras')
    try {
      const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes(name)]) }, { docId: `extras-lazy-${name}`, profile: 'D' })
      await session.done
      expect(spy).toHaveBeenCalledTimes(calls)
      const source = createPagedWorkbookSource(session)
      try {
        const sheet = source.sheets.find(item => item.name !== 'Data') ?? source.sheets[0]!, store = source.store(sheet.sheetId)
        const options = { signal: new AbortController().signal, stamp: store.stamp }
        if (name === 'slicer-light1') {
          // A slicer sheet loads the extras, so a default slicer style in styles.xml applies as on the resident path.
          expect(session.workbook.tableStyleExtras).toBeDefined()
        } else if (name === 'table-medium-2') {
          // The built-in style still resolves, from the lazily loaded built-ins.
          const styles = await store.readRegionStyles!(options, 1024 * 1024, { set: () => true })
          expect(styles.tables.map(item => item.def.name)).toEqual(['TableStyleMedium2'])
        } else {
          const display = await store.readPivotDisplay!(options, 4 * 1024 * 1024, { set: () => true })
          expect(display.pivotTables).toHaveLength(1)
        }
      } finally { await source.release(); await session.release() }
    } finally { spy.mockRestore() }
  })

  describe('main-thread layer parity with the resident path (final review I1, I2)', () => {
    // pivot-formats-override: one <format> (dxf 0, #FFE699 fill) over B7, which Excel also writes into B7's own s=.
    // Removing that s= leaves the <format> as the only source of the fill, as a non-Excel writer saves it.
    const ownStyleRemoved = (xml: string) => xml.replace('<c r="B7" s="5">', '<c r="B7">')
    const border = (xml: string) => xml.replace('<dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFE699"/></patternFill></fill></dxf>',
      '<dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFE699"/></patternFill></fill><border><bottom style="thick"><color rgb="FFFF0000"/></bottom></border></dxf>')
    const cells = (sheet: XlsxSheet) => {
      const byKey = new Map(sheet.rows.flatMap(row => row.cells.map(cell => [`${cell.row}:${cell.col}`, cell.style] as const)))
      const out: Array<{ row: number; col: number; own: XlsxSheet['rows'][number]['cells'][number]['style'] }> = []
      for (let row = 2; row <= 22; row++) for (let col = 0; col <= 8; col++) out.push({ row, col, own: byKey.get(`${row}:${col}`) })
      return out
    }

    it.each([
      ['the format dxf only', { 'xl/worksheets/sheet1.xml': ownStyleRemoved }],
      ['a format dxf with a border', { 'xl/worksheets/sheet1.xml': ownStyleRemoved, 'xl/styles.xml': border }],
    ] as const)('paints %s the same paged as resident', async (name, edits) => {
      const input = await editPackage('pivot-formats-override', edits)
      const resident = await residentPivotSheet(input)
      const paged = await pagedPivotLayer(input, `pivot-formats-${name}`)
      const residentLayer = resident.regionStyles!, pagedLayer = paged.layer!
      expect(residentLayer.styleAt(6, 1, undefined, 'screen')?.fillColor).toBe('FFFFE699')
      for (const { row, col, own } of cells(resident)) {
        for (const own2 of [own, undefined]) expect(pagedLayer.styleAt(row, col, own2, 'screen'), `${row}:${col}`).toEqual(residentLayer.styleAt(row, col, own2, 'screen'))
      }
      expect(paged.payload.pivotDxfs?.map(item => item.id)).toEqual([0])
      if (name.includes('border')) {
        expect(pagedLayer.styleAt(6, 1, undefined, 'screen')?.borders?.bottom).toBe('thick')
        // The shared-edge rule: the cell below shows the format's line on its top edge.
        expect(pagedLayer.styleAt(7, 1, undefined, 'screen')?.borders?.top).toBe('thick')
      }
    }, 30_000)

    it('reports the layer diagnostics paged as resident (unknown style, no saved layout)', async () => {
      const input = await editPackage('pivot-formats-override', {
        'xl/pivotTables/pivotTable1.xml': xml => xml.replace(/<pivotTableStyleInfo name="[^"]*"/, '<pivotTableStyleInfo name="NoSuchStyle"').replace(/<rowItems\b[\s\S]*?<\/rowItems>/, ''),
      })
      const resident = await residentPivotSheet(input)
      const paged = await pagedPivotLayer(input, 'pivot-diagnostics-parity')
      const features = (list: readonly { feature: string }[]) => [...new Set(list.map(item => item.feature))].filter(feature => feature.startsWith('pivot-') || feature.startsWith('table-style-')).sort()
      expect(features(resident.diagnostics ?? [])).toEqual(expect.arrayContaining(['pivot-layout-unavailable', 'pivot-style-unknown']))
      expect(features(paged.diagnostics)).toEqual(features(resident.diagnostics ?? []))
    }, 30_000)
  })

  it('releases retained descriptors on sheet-owner change and close, and settles error replies', async () => {
    const owner: CacheIdentity = { docId: 'pivot-cache', sheetId: 's', valueGen: 1, fontGen: 0, viewGen: 0 }
    let currentOwner = { ...owner }
    let released = 0, resolve!: (value: WorkerReply<'read'>) => void
    const success = (stamp = currentOwner) => ({ kind: 'complete', commandKind: 'read', requestId: 'r', owner: stamp, diagnostics: [], replyBytes: 64,
      result: { kind: 'read', result: { op: 'pivotDisplay', result: { sheetId: 's', stamp, pivotTables: [], pivotStyles: [], diagnostics: [] } } } }) as unknown as WorkerReply<'read'>
    const cache = new RemoteRegionStyleCache({
      identity: () => currentOwner, hasRegionStyles: () => false, hasPivotDisplay: () => true, loadState: () => ({ completeness: 'complete' }),
      grant: () => 4096, releaseGrant: () => {}, request: () => new Promise<WorkerReply<'read'>>(done => { resolve = done }),
      nextRequestId: () => 'r', release: () => { released++ }, admitLayer: () => () => {}, changed: () => {},
    }, 4096, 'pivotDisplay')
    cache.get('s'); resolve(success())
    await new Promise(resolveTick => setTimeout(resolveTick, 0))
    expect(cache.get('s').state).toBe('ready')
    currentOwner = { ...currentOwner, valueGen: 2 }
    cache.get('s')
    expect(released).toBe(1)
    resolve(success(currentOwner))
    await new Promise(resolveTick => setTimeout(resolveTick, 0))
    cache.invalidate('s')
    expect(released).toBe(2)
    cache.close()

    const failing = new RemoteRegionStyleCache({
      identity: () => currentOwner, hasRegionStyles: () => false, hasPivotDisplay: () => true, loadState: () => ({ completeness: 'complete' }),
      grant: () => 4096, releaseGrant: () => {}, request: async () => ({ kind: 'error', commandKind: 'read', requestId: 'err', owner, code: 'parse-error', retryable: false, diagnostics: [] } as unknown as WorkerReply<'read'>),
      nextRequestId: () => 'err', release: () => { released++ }, admitLayer: () => () => {}, changed: () => {},
    }, 4096, 'pivotDisplay')
    failing.get('s')
    await new Promise(resolveTick => setTimeout(resolveTick, 0))
    expect(released).toBe(3)
    expect(failing.get('s').diagnostics).toContainEqual(expect.objectContaining({ feature: 'pivot-display-read-failed' }))
    failing.close()
  })

  it('charges each formatted pivot\'s overrides memo to the layer, and refuses what does not fit (codex re-review)', async () => {
    const resident = await residentPivotSheet(bytes('pivot-formats-override'))
    const formatted = resident.pivotTables![0]!
    expect(formatted.formats.length).toBeGreaterThan(0)
    const plain = { ...formatted, formats: [] }
    const admitted = async (pivotTables: RemotePivotDisplay['pivotTables'], headroom = Infinity) => {
      const owner: CacheIdentity = { docId: `memo-${pivotTables.length}-${headroom}`, sheetId: 's', valueGen: 1, fontGen: 0, viewGen: 0 }
      const payload: RemotePivotDisplay = { sheetId: 's', stamp: owner, pivotTables, pivotStyles: [], pivotDxfs: [{ id: 0, dxf: { fill: { pattern: 'solid', bg: { rgb: 'FFFFE699' } } } }], diagnostics: [] }
      const reply = { kind: 'complete', commandKind: 'read', requestId: 'r', owner, diagnostics: [], replyBytes: 64,
        result: { kind: 'read', result: { op: 'pivotDisplay', result: payload } } } as unknown as WorkerReply<'read'>
      let charged = 0
      const cache = new RemoteRegionStyleCache({
        identity: () => owner, hasRegionStyles: () => false, hasPivotDisplay: () => true, loadState: () => ({ completeness: 'complete' }),
        grant: () => 4096, releaseGrant: () => {}, request: async () => reply, nextRequestId: () => 'r', release: () => {},
        admitLayer: (_sheet, _request, bytes) => { if (bytes > headroom) return null; charged = bytes; return () => { charged = 0 } }, changed: () => {},
      }, 256 * 1024 * 1024, 'pivotDisplay')
      let state = cache.get('s')
      for (let i = 0; i < 500 && state.state !== 'ready' && state.state !== 'failed'; i++) { await new Promise(resolveTick => setTimeout(resolveTick, 2)); state = cache.get('s') }
      const out = { state: state.state, charged, retained: cache.retainedBytes, features: state.diagnostics.map(item => item.feature) }
      cache.close()
      return out
    }
    const one = await admitted([formatted]), eight = await admitted(Array(8).fill(formatted)), eightPlain = await admitted(Array(8).fill(plain))
    expect(one.state).toBe('ready')
    expect(eight.charged - eightPlain.charged).toBe(8 * PIVOT_FORMAT_MEMO_BYTES)
    expect(eight.charged - one.charged).toBeGreaterThanOrEqual(7 * PIVOT_FORMAT_MEMO_BYTES)
    expect(eight.retained).toBeGreaterThanOrEqual(eight.charged)
    // Headroom for the plain layer but not for the memos: refused, never over-retained.
    const refused = await admitted(Array(8).fill(formatted), eightPlain.charged + PIVOT_FORMAT_MEMO_BYTES)
    expect(refused.state).toBe('failed')
    expect(refused.features).toContain('pivot-display-reply-limit')
    expect(refused.charged).toBe(0)
  })

  it('uses pivot-display diagnostics when the receiver cannot admit its decoded layer', async () => {
    const owner: CacheIdentity = { docId: 'pivot-cache-budget', sheetId: 's', valueGen: 1, fontGen: 0, viewGen: 0 }
    const payload: RemotePivotDisplay = { sheetId: 's', stamp: owner, pivotTables: [], pivotStyles: [{ name: 'custom', table: false, pivot: true, elements: {} }], diagnostics: [] }
    const reply = { kind: 'complete', commandKind: 'read', requestId: 'r', owner, diagnostics: [], replyBytes: 64,
      result: { kind: 'read', result: { op: 'pivotDisplay', result: payload } } } as unknown as WorkerReply<'read'>
    const cache = new RemoteRegionStyleCache({
      identity: () => owner, hasRegionStyles: () => false, hasPivotDisplay: () => true, loadState: () => ({ completeness: 'complete' }),
      grant: () => 4096, releaseGrant: () => {}, request: async () => reply, nextRequestId: () => 'r', release: () => {},
      admitLayer: () => null, changed: () => {},
    }, 4096, 'pivotDisplay')
    cache.get('s')
    await new Promise(resolveTick => setTimeout(resolveTick, 0))
    expect(cache.get('s').diagnostics).toContainEqual(expect.objectContaining({ feature: 'pivot-display-reply-limit' }))
    cache.close()
  })

  it('propagates a control-part cancellation instead of converting it to missing controls', async () => {
    const controller = new AbortController()
    const pkg = { xml: async () => { controller.abort(); throw new PackageError('cancelled', 'cancelled') }, xmlOrdered: async () => undefined } as unknown as PackageReader
    await expect(resolveSheetControls(pkg, { sheetPath: 'xl/worksheets/sheet1.xml', slicerIds: ['rId1'], timelineIds: [], theme: parseThemeContext(), signal: controller.signal }))
      .rejects.toMatchObject({ code: 'cancelled' })
  })
})
