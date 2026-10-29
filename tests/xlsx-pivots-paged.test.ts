// @vitest-environment node
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { describe, expect, it, vi } from 'vitest'
import type { CacheIdentity, RemotePivotDisplay, SourceStore, WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
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

const bytes = (name: string) => new Uint8Array(readFileSync(resolve(process.cwd(), `tests/fixtures/xlsx-pivots/${name}.xlsx`)))
function pixels(sheet: Parameters<typeof computeMetrics>[0]): Buffer {
  const metrics = computeMetrics(sheet), canvas = createCanvas(metrics.widthPx, metrics.heightPx)
  renderSheet(sheet, canvas.getContext('2d') as unknown as CanvasRenderingContext2D, metrics)
  return Buffer.from(canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data.buffer)
}

describe('paged pivot display descriptors', () => {
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
    const payload: RemotePivotDisplay = { sheetId: stamp.sheetId, stamp, tables: [], pivotTables: [], pivotStyles: [], slicerStyles: [], diagnostics: [] }
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
      expect(result.slicerStyles).toEqual([])
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
      expect(result.slicerStyles).toEqual([])
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
      const pagedSheet = sheetWithRegionStyles(sheetWithDrawings(residentSheet, remote.drawings(sheetId)), display)
      expect(pixels(pagedSheet)).toEqual(pixels(residentSheet))
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
      expect(pagedPixels).toEqual(residentPixels)
      expect(pagedPixels).not.toEqual(Buffer.alloc(pagedPixels.length))
      const withoutLayer = { ...residentSheet, pivotRegions: undefined, pivotTables: undefined, regionStyles: undefined, drawings: [] }
      expect(pagedPixels).not.toEqual(paint(withoutLayer, []))
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
    const session = await openImportSession({ kind: 'blob', blob: new Blob([inflated]) }, { docId: `pivot-cap-${fixture}`, profile: 'D' })
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

  it('releases retained descriptors on sheet-owner change and close, and settles error replies', async () => {
    const owner: CacheIdentity = { docId: 'pivot-cache', sheetId: 's', valueGen: 1, fontGen: 0, viewGen: 0 }
    let currentOwner = { ...owner }
    let released = 0, resolve!: (value: WorkerReply<'read'>) => void
    const success = (stamp = currentOwner) => ({ kind: 'complete', commandKind: 'read', requestId: 'r', owner: stamp, diagnostics: [], replyBytes: 64,
      result: { kind: 'read', result: { op: 'pivotDisplay', result: { sheetId: 's', stamp, tables: [], pivotTables: [], pivotStyles: [], slicerStyles: [], diagnostics: [] } } } }) as unknown as WorkerReply<'read'>
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
    cache.closeSheet('s')
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

  it('uses pivot-display diagnostics when the receiver cannot admit its decoded layer', async () => {
    const owner: CacheIdentity = { docId: 'pivot-cache-budget', sheetId: 's', valueGen: 1, fontGen: 0, viewGen: 0 }
    const payload: RemotePivotDisplay = { sheetId: 's', stamp: owner, tables: [], pivotTables: [], pivotStyles: [{ name: 'custom', table: false, pivot: true, elements: {} }], slicerStyles: [], diagnostics: [] }
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
