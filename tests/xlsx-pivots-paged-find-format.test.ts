// @vitest-environment node
/**
 * PR #55 round 3 P2: worker-backed Find formats a styled table/pivot cell with the
 * number format its region supplies (a pivot `formats` dxf here), the text the viewer
 * paints, as resident Find already does (InteractiveSheetDoc textOf).
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cellDisplayText } from '../src/xlsx/render'
import { ensureTableStyleEngine } from '../src/xlsx/tableStyleLoader'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { DocumentSession } from '../src/worker/viewerSession'

/** pivot-formats-override with its `formats` dxf also carrying a percent number format. */
async function percentPivot(): Promise<Uint8Array<ArrayBuffer>> {
  const archive = await JSZip.loadAsync(readFileSync(resolve(process.cwd(), 'tests/fixtures/xlsx-pivots/pivot-formats-override.xlsx')))
  const styles = await archive.file('xl/styles.xml')!.async('string')
  const edited = styles.replace('<dxf><fill>', '<dxf><numFmt numFmtId="10" formatCode="0.00%"/><fill>')
  expect(edited).not.toBe(styles)
  archive.file('xl/styles.xml', edited)
  return new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
}

describe('worker-backed Find uses the region-resolved number format', () => {
  it('finds a pivot value by the percent text its `formats` dxf paints', async () => {
    const input = await percentPivot()
    // Resident reference: the cell the dxf formats, and the text the viewer paints there.
    const doc = await parseXlsx(await OfficePackage.load(input))
    await ensureTableStyleEngine(doc.sheets)
    const sheet = doc.sheets.find(item => item.pivotTables?.length)!
    let target: { row: number; col: number; painted: string; own: string } | undefined
    for (const row of sheet.rows) for (const cell of row.cells) {
      if (target || typeof cell.value !== 'number') continue
      const style = sheet.regionStyles?.styleAt(cell.row, cell.col, cell.style, 'screen')
      if (style?.formatCode !== '0.00%') continue
      target = { row: cell.row, col: cell.col, painted: cellDisplayText(sheet, { value: cell.value, style }), own: cellDisplayText(sheet, { value: cell.value, style: cell.style }) }
    }
    expect(target).toBeDefined()
    expect(target!.painted).toMatch(/%$/)
    expect(target!.own).not.toContain('%')

    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: 'paged-find-format', profile: 'D' })
    await session.done
    const worker = new DocumentSession('paged-find-format', createPagedWorkbookSource(session))
    try {
      const result = await worker.search(worker.workbookOwner(), { text: target!.painted, caseSensitive: false }, undefined, new AbortController().signal)
      expect(result).toMatchObject({ complete: true })
      if (!('hits' in result)) throw new Error('search did not complete')
      // 429f26da: no hit, the worker formatted the cell with its own General style.
      expect(result.hits.map(hit => ({ row: hit.source.row, col: hit.source.col, text: hit.text }))).toContainEqual({ row: target!.row, col: target!.col, text: target!.painted })
    } finally { await worker.release(); await session.release() }
    // The search's descriptor reads and decoded layer were charged to it and end with the session.
    expect(worker.memoryStats().transientBytes).toBe(0)
  }, 60_000)
})
