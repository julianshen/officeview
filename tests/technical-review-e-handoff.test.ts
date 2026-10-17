// @vitest-environment node
import { expect, it } from 'vitest'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { DocumentSession } from '../src/worker/viewerSession'
import type { XlsxSheet } from '../src/xlsx/types'

it('a diagnostic already admitted by the formatter transfers without reserving its bytes twice', async () => {
  const sheet: XlsxSheet = { name: 'S', sheetId: '1', cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [{ ref: 'A1', row: 0, col: 0, styleIndex: 0, value: 3, style: { numFmtId: 164, formatCode: '[Purple]0' } }] }] }
  const source = createObjectWorkbookSource({ sheets: [sheet] }, 'd')
  const original = source.displayText!
  const session = new DocumentSession('d', source)
  let observed = false
  source.displayText = (cell, id, allocation) => original(cell, id, allocation && { ...allocation, retain(targets, bytes, commit, kind) {
    const pressure = session.memory.charge()
    const before = session.memory.totalBytes, room = bytes - 1
    expect(pressure.set(session.memory.limitBytes - before - room)).toBe(true)
    observed = true
    console.log(JSON.stringify({ before, diagnosticBytesAlreadyReserved: bytes, roomDuringTransfer: session.memory.limitBytes - session.memory.totalBytes }))
    try { allocation.retain!(targets, bytes, commit, kind) } finally { pressure.release() }
  } })
  try {
    const result = await session.search(session.workbookOwner(), { text: '3', caseSensitive: false }, undefined, new AbortController().signal)
    expect(observed).toBe(true)
    console.log(JSON.stringify(result))
    expect(result).toMatchObject({ complete: true, hits: [{ text: '3' }] })
  } finally { await session.release() }
})
