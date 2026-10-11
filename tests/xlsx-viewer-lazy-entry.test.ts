/** Orchestrator decision 1: DOCX/PPTX byte inputs never load the large-source chunk. */
import { describe, expect, it, vi } from 'vitest'
import { loadOfficeFile } from '../src/components/OfficeFile'
import { buildDocx, buildPptx, buildXlsx } from '../src/testdata/ooxml-builders'

vi.mock('../src/components/largeSource', () => {
  throw new Error('large-source chunk loaded')
})

describe('lazy large-source chunk', () => {
  it('DOCX and PPTX bytes load without it', async () => {
    expect('sections' in await loadOfficeFile(await buildDocx([{ runs: [{ text: 'no chunk' }] }]))).toBe(true)
    expect('slides' in await loadOfficeFile(await buildPptx([{ prst: 'rect' }]))).toBe(true)
  })

  it('DOCX and PPTX Blobs (local files) load without it (decision D)', async () => {
    const docx = new Blob([await buildDocx([{ runs: [{ text: 'blob' }] }]) as BlobPart])
    expect('sections' in await loadOfficeFile(docx)).toBe(true)
    const pptx = new Blob([await buildPptx([{ prst: 'rect' }]) as BlobPart])
    expect('slides' in await loadOfficeFile(pptx)).toBe(true)
  })

  it('a spreadsheet Blob loads it', async () => {
    await expect(loadOfficeFile(new Blob([await buildXlsx([{ name: 'S', rows: [] }]) as BlobPart]))).rejects.toThrow(/mock/)
  })

  it('a spreadsheet is the trigger that loads it', async () => {
    await expect(loadOfficeFile(await buildXlsx([{ name: 'S', rows: [] }]))).rejects.toThrow(/mock/)
  })
})
