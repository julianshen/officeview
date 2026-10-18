/** Background priority (spec §10.2): the progressive import parks at its
 * source reads while the host has foreground work (the kernel passes
 * waitForForeground as `beforeRead`). */
import { describe, expect, it } from 'vitest'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, memoryPageBackend } from '../src/xlsx/storage/sourcePages'
import { buildZip } from './helpers/zip-builder'
import { xlsxEntries } from './helpers/xlsx-package'

function workbook(rows: number): Uint8Array {
  let sheetData = ''
  for (let row = 1; row <= rows; row++) sheetData += `<row r="${row}"><c r="A${row}"><v>${row}</v></c></row>`
  return buildZip(xlsxEntries({ sheetData })).bytes
}

describe('import source reads respect the host priority hook', () => {
  it('awaits beforeRead before reads, and makes no progress while it is held', async () => {
    const bytes = workbook(2000)
    let calls = 0, held: (() => void) | undefined, holding = false
    const beforeRead = () => {
      calls++
      if (!holding) return
      return new Promise<void>(resolve => { held = resolve })
    }
    const store = new SourcePageStore({ docId: 'p', backend: memoryPageBackend(), memoryBudgetBytes: 64 * 1024 * 1024 })
    holding = true
    const opening = openImportSession({ kind: 'blob', blob: new Blob([bytes as BlobPart]) }, { docId: 'p', profile: 'D', store, beforeRead })
    // The first read parks: the session cannot even open while the hook holds.
    let opened = false
    void opening.then(() => { opened = true })
    for (let i = 0; i < 20; i++) await new Promise(resolve => setTimeout(resolve, 0))
    expect(calls).toBe(1)
    expect(opened).toBe(false)
    holding = false
    held?.()
    const session = await opening
    await session.done
    expect(calls).toBeGreaterThan(1)
    await session.release()
  })
})
