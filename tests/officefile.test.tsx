import { describe, test, expect } from 'vitest'
import { render, waitFor } from '@testing-library/react'
import { OfficeFile, loadOfficeFile } from '../src/components/OfficeFile'
import { buildDocx, buildXlsx, buildPptx } from '../src/testdata/ooxml-builders'

describe('loadOfficeFile', () => {
  test('detects docx by content type', async () => {
    const doc = await loadOfficeFile(await buildDocx([{ runs: [{ text: 'hi' }] }]))
    expect('sections' in doc).toBe(true)
  })

  test('detects xlsx by content type', async () => {
    const doc = await loadOfficeFile(await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }] }]))
    expect('sheets' in doc).toBe(true)
  })

  test('detects pptx by content type', async () => {
    const doc = await loadOfficeFile(await buildPptx([{ prst: 'rect' }]))
    expect('slides' in doc).toBe(true)
  })
})

describe('<OfficeFile>', () => {
  test('renders async from raw bytes', async () => {
    const data = await buildDocx([{ runs: [{ text: 'async render' }] }])
    const { container } = render(<OfficeFile data={data} />)
    await waitFor(() => expect(container.querySelectorAll('canvas').length).toBeGreaterThanOrEqual(1))
  })

  test('shows error element on garbage input after promise settles', async () => {
    const { container } = render(<OfficeFile data={new Uint8Array([1, 2, 3])} />)
    for (let i = 0; i < 20; i++) {
      await new Promise((r) => setTimeout(r, 20))
      if (container.querySelector('[role="alert"]')) break
    }
    expect(container.querySelector('[role="alert"]')).toBeTruthy()
  })
})
