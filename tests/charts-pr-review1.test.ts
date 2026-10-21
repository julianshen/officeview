/**
 * PR #15 review (codex), round 1:
 * P1 — a paged fill's live scratch (its cell maps and typed arrays) stays counted in every chart-ledger admission
 *      the fill makes while it holds that scratch (model storage built during the fill), until the fill releases it.
 * P2 — a discarded evaluation parse whose every chart gave way to an authored fallback still returns the sheet's
 *      retained chart allocations (a published formatter diagnostic) to the document's chart workspace.
 * Fixtures are Office-shaped synthetic workbooks.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { drawingPartContext } from '../src/drawing/parts'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { loadPagedSheetDrawings, pagedDrawingLimits } from '../src/xlsx/storage/pagedDrawings'
import { parseSpreadsheet } from '../src/components/largeSource'
import type { ChartParentLedger, ChartWorkspace } from '../src/charts/workspace'
import { CACHED_COLUMN, NS, addNumberFormat, chartWorkbookBytes, emu } from './helpers/xlsx-chart-fixture'

// P2: largeSource's evaluator-cap reopen, armed per test.
const cap = vi.hoisted(() => ({ armed: false }))
vi.mock('../src/xlsx/storage/admission', async original => {
  const actual = await original<typeof import('../src/xlsx/storage/admission')>()
  return { ...actual, evaluationHitCap: (doc: Parameters<typeof actual.evaluationHitCap>[0]) => cap.armed ? (cap.armed = false, 'evaluation-cap') : actual.evaluationHitCap(doc) }
})
// P2: once a label's formatter diagnostic has moved to the sheet's owner (charts/own retainWithTarget: a transfer
// from a `:label` tag to a `chart-model#n` tag), the next label reservation is refused: the chart is refused after
// its diagnostic was retained, and its authored fallback replaces it.
const refuseAfterRetain = vi.hoisted(() => ({ on: false, armed: false }))
vi.mock('../src/charts/workspace', async original => {
  const actual = await original<typeof import('../src/charts/workspace')>()
  return {
    ...actual,
    createChartWorkspace: (profile: Parameters<typeof actual.createChartWorkspace>[0], parent?: ChartParentLedger) => {
      const workspace = actual.createChartWorkspace(profile, parent)
      const reserve = workspace.reserve.bind(workspace), transfer = workspace.transfer.bind(workspace)
      workspace.reserve = (bytes, tag) => {
        if (refuseAfterRetain.armed && tag.endsWith(':label')) { refuseAfterRetain.armed = false; return false }
        return reserve(bytes, tag)
      }
      workspace.transfer = (from, to, bytes) => {
        transfer(from, to, bytes)
        if (refuseAfterRetain.on && from.endsWith(':label') && /^chart-model#\d+$/.test(to)) refuseAfterRetain.armed = true
      }
      return workspace
    },
  }
})

afterEach(() => { cap.armed = false; refuseAfterRetain.on = false; refuseAfterRetain.armed = false })

describe('PR #15 review round 1', () => {
  it('P1: a paged fill\'s live scratch stays counted in the chart-ledger admissions made while it is held', async () => {
    const bytes = await chartWorkbookBytes(CACHED_COLUMN)
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'P1', profile: 'D' })
    expect((await session.done).kind).toBe('complete')
    const source = createPagedWorkbookSource(session)
    try {
      const sheet = session.workbook.sheets[0]
      const SCRATCH = 100_000, MODEL = 500
      let budget = Infinity
      const sets: number[] = []
      const memory = { set: (value: number) => { sets.push(value); return value <= budget } }
      const seen: { base?: number; withScratch?: number; afterModel?: number; refused?: boolean; refusedAt?: number; afterRelease?: number } = {}
      await loadPagedSheetDrawings(session.workbook.parts, { name: sheet.name, partPath: sheet.partPath, drawingIds: sheet.metadata.drawingIds },
        source.geometry('1'), pagedDrawingLimits('D'), new AbortController().signal, {
          memory, chartDocument: {},
          fillCharts: async (_drawings, _owner, _names, admit, workspace) => {
            const scope = workspace as ChartWorkspace
            expect(scope).toBeDefined()
            // The fill reserves its scratch: the share is the read's working set plus the scratch.
            expect(admit(SCRATCH)).toBe(true)
            seen.withScratch = sets.at(-1)!
            seen.base = seen.withScratch - SCRATCH
            // A model allocation made while the scratch is live: both are counted.
            expect(scope.reserve(MODEL, 'review-model')).toBe(true)
            seen.afterModel = sets.at(-1)!
            // A model release while the scratch is live keeps the scratch counted too.
            scope.release('review-model')
            seen.afterRelease = sets.at(-1)!
            // The combined total over the budget is refused, though the allocation alone would fit
            // (a pool refusal refuses the whole read).
            expect(scope.reserve(MODEL, 'review-model')).toBe(true)
            budget = seen.base + SCRATCH + MODEL + MODEL - 1
            seen.refused = !scope.reserve(MODEL, 'review-model-2')
            seen.refusedAt = sets.at(-1)!
            return { pendingSheets: [], refused: true }
          },
        })
      expect(seen.afterModel).toBe(seen.base! + SCRATCH + MODEL)
      expect(seen.refused).toBe(true)
      expect(seen.refusedAt).toBe(seen.base! + SCRATCH + MODEL + MODEL)
      expect(seen.afterRelease).toBe(seen.base! + SCRATCH)
    } finally { await source.release() }
  })

  it('P1: a fill that ends releases its scratch before the read goes on (fallback parts and the payload)', async () => {
    const bytes = await chartWorkbookBytes(CACHED_COLUMN)
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'P1b', profile: 'D' })
    expect((await session.done).kind).toBe('complete')
    const source = createPagedWorkbookSource(session)
    try {
      const sheet = session.workbook.sheets[0]
      const sets: number[] = []
      let base = 0
      const result = await loadPagedSheetDrawings(session.workbook.parts, { name: sheet.name, partPath: sheet.partPath, drawingIds: sheet.metadata.drawingIds },
        source.geometry('1'), pagedDrawingLimits('D'), new AbortController().signal, {
          memory: { set: value => { sets.push(value); return true } }, chartDocument: {},
          fillCharts: async (_drawings, _owner, _names, admit, workspace) => {
            admit(100_000)
            base = sets.at(-1)! - 100_000
            ;(workspace as ChartWorkspace).reserve(500, 'review-model')
            ;(workspace as ChartWorkspace).release('review-model')
            return { pendingSheets: [] }
          },
        })
      // The read's final charge is its working set plus the payload: no scratch left in it.
      expect(sets.at(-1)).toBe(result.workingSet)
      expect(result.workingSet - (result.payloadBytes ?? 0)).toBe(base)
      result.chartScope?.releaseAll()
    } finally { await source.release() }
  })

  it('P2: a discarded evaluation parse whose charts all gave way to fallbacks returns the sheet\'s retained charge', async () => {
    const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
    const chart = `<c:chartSpace xmlns:c="${C}"><c:chart><c:plotArea><c:lineChart><c:ser><c:idx val="0"/><c:cat><c:strRef><c:f>Styled!$A$1:$A$2</c:f></c:strRef></c:cat><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>`
    const graphicFrame = `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="${C}"><c:chart r:id="rChart"/></a:graphicData></a:graphic></xdr:graphicFrame>`
    const fallback = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Fallback"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></xdr:spPr></xdr:sp>`
    const frame = `<xdr:oneCellAnchor><xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${emu(320)}" cy="${emu(200)}"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c">${graphicFrame}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:oneCellAnchor>`
    expect(NS).toContain('xmlns:c=')
    // A1:A2 carry an unsupported custom format: formatting a category label publishes a diagnostic on the sheet.
    const bytes = await chartWorkbookBytes(chart, {
      frame, cells: [{ ref: 'A1', v: 5, style: 3 }, { ref: 'A2', v: 6, style: 3 }],
      patch: async zip => {
        const styles = await zip.file('xl/styles.xml')!.async('string')
        zip.file('xl/styles.xml', styles.replace(/<fonts/, '<numFmts count="1"><numFmt numFmtId="164" formatCode="[FOO]0"/></numFmts><fonts'))
        await addNumberFormat(zip, 164)
      },
    })
    const pkg = await OfficePackage.load(bytes)
    const decided = { mode: 'object', decision: { evaluation: { maxArrayCells: 1000, maxArraySide: 1000, maxMaterializedCells: 1000 }, valuePolicy: 'engine', counters: { loadEvaluatedFormulas: 0, alternateContentData: [], highRatioParts: [] } } }
    const ctx = { profile: 'D', options: {}, limits: { object: { maxRatio: 100 } } }
    refuseAfterRetain.on = true
    cap.armed = true
    const doc = await parseSpreadsheet(pkg, bytes, ctx as never, decided as never)
    expect(cap.armed, 'the evaluation parse was discarded').toBe(false)
    const sheet = doc.sheets[0]
    // The reopened parse: the chart gave way to its fallback, after its diagnostic was published on the sheet.
    expect(sheet.drawings?.map(node => [node.source.element, node.source.representation, node.content?.kind])).toEqual([['sp', 'fallback', undefined]])
    expect(sheet.diagnostics?.length).toBeGreaterThan(0)
    const workspace = drawingPartContext(pkg).chart?.workspace as ChartWorkspace
    expect(workspace.used, 'the live sheet holds its retained charge').toBeGreaterThan(0)
    // The live document's release returns the workspace to empty: the discarded parse left nothing behind.
    const { releaseChartRetained } = await import('../src/charts')
    releaseChartRetained(sheet)
    expect(workspace.used).toBe(0)
  })
})
