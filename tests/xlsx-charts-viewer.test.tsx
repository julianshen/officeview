import { afterEach, describe, expect, it } from 'vitest'
import { cleanup, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { parseXlsx } from '../src/xlsx/parse'
import { chartWorkbook } from './helpers/xlsx-chart-fixture'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'

afterEach(cleanup)

const openpyxlChart = `<chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><chart><title><tx><rich><a:p><a:r><a:t>Bar chart</a:t></a:r></a:p></rich></tx></title><plotArea><barChart><barDir val="col"/><grouping val="clustered"/><ser><idx val="0"/><order val="0"/><cat><numRef><f>'Styled'!$K$20:$K$22</f></numRef></cat><val><numRef><f>'Styled'!$L$20:$L$22</f></numRef></val></ser></barChart><catAx/><valAx/></plotArea><legend/></chart></chartSpace>`
const cells = [0, 1, 2].flatMap(i => [{ ref: `K${20 + i}`, v: `Q${i + 1}` }, { ref: `L${20 + i}`, v: (i + 1) * 10 }])

describe('interactive viewer chart accessibility', () => {
  it('names every chart on the active sheet and gives a notice for a placeholder', FIRST_PAINT_TEST, async () => {
    const doc = await parseXlsx(await chartWorkbook(openpyxlChart, { cells }))
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    await waitFor(() => expect(ui.getAllByRole('gridcell').length).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    const list = ui.getByRole('list', { name: 'Charts on this sheet (1)' })
    expect(list.querySelector('[role="img"]')?.getAttribute('aria-label')).toBe('Chart: Bar chart. Column chart, 1 series; 3 categories (Q1, Q2, Q3); Series 1: 10, 20, 30.')
    cleanup()
    const radar = await parseXlsx(await chartWorkbook('<chartSpace><chart><title><tx><rich><p><r><t>Radar</t></r></p></rich></tx></title><plotArea><radarChart/></plotArea></chart></chartSpace>'))
    const second = render(<InteractiveSheetDoc doc={radar} />)
    await waitFor(() => expect(second.getAllByRole('gridcell').length).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    expect(second.getByRole('img', { name: 'Chart: Radar. Radar chart is not supported; it is not drawn.' })).toBeTruthy()
    expect(second.container.textContent).toContain('Chart “Radar” is shown as a placeholder: Radar chart is not supported.')
  })
})
