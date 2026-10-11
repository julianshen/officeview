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

  it('a spreadsheet is the trigger that loads it', async () => {
    await expect(loadOfficeFile(await buildXlsx([{ name: 'S', rows: [] }]))).rejects.toThrow(/mock/)
  })
})
