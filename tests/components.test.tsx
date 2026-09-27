import { describe, test, expect } from 'vitest'
import { render } from '@testing-library/react'
import { OfficeDoc } from '../src/components/OfficeDoc'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { buildDocx } from './ooxml-fixtures'

describe('<OfficeDoc>', () => {
  test('renders one canvas per page for a docx', async () => {
    const doc = await parseDocx(await OfficePackage.load(await buildDocx([
      { runs: [{ text: 'Hello component' }] },
    ])))
    const { container } = render(<OfficeDoc document={doc} />)
    const canvases = container.querySelectorAll('canvas')
    expect(canvases.length).toBeGreaterThanOrEqual(1)
    // canvas has non-zero backing size
    expect(canvases[0].width).toBeGreaterThan(0)
    expect(canvases[0].height).toBeGreaterThan(0)
  })

  test('renders canvas per sheet and per slide', async () => {
    const { parseXlsx } = await import('../src/xlsx/parse')
    const { buildXlsx } = await import('./ooxml-fixtures')
    const xlsx = await parseXlsx(await OfficePackage.load(await buildXlsx([
      { name: 'S1', rows: [{ r: 1, cells: [{ ref: 'A1', v: 5 }] }] },
    ])))
    const x = render(<OfficeDoc document={xlsx} />)
    expect(x.container.querySelectorAll('canvas')).toHaveLength(1)

    const { parsePptx } = await import('../src/pptx/parse')
    const { buildPptx } = await import('./ooxml-fixtures')
    const pptx = await parsePptx(await OfficePackage.load(await buildPptx([
      { prst: 'rect', off: ['0', '0'], ext: ['9144000', '6858000'] },
    ])))
    const p = render(<OfficeDoc document={pptx} />)
    expect(p.container.querySelectorAll('canvas')).toHaveLength(1)
  })
})
