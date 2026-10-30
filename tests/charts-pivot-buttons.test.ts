// @vitest-environment node
/**
 * Task 11: Pivot chart link (ChartModel.pivot).
 * Excel for Mac draws NO field buttons for any flag combination (Ruling AI);
 * ChartModel.pivot preserves c:pivotSource and the dropZone / expand-collapse
 * options for Stage 3 while rendering pixel-identical to a standard chart.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { createChartWorkspace } from '../src/charts/workspace'
import { stringBytes } from '../src/charts/points'
import type { ChartParseContext } from '../src/charts/context'
import { paintXlsxChart } from '../src/charts/paint'
import { layoutChart } from '../src/charts/layout'
import { parseXlsx } from '../src/xlsx/parse'
import type { ChartModel, ChartPivotSource } from '../src/charts/model'

const FIXTURES_DIR = join(__dirname, 'fixtures/xlsx-pivots')
const parse = (xmlNode: ReturnType<typeof parseXmlOrdered>) => parseXlsxChart(xmlNode, undefined)

interface FixtureCase {
  file: string
  expected: ChartPivotSource
}

const CASES: FixtureCase[] = [
  {
    file: 'pivot-chart-buttons-all.xlsx',
    expected: {
      name: '[pivot-chart-buttons-all.xlsx]Pivot!樞紐分析表2',
      fmtId: 0,
      dropZones: {
        filter: true,
        categories: true,
        data: true,
        series: true,
        visible: true,
      },
      showExpandCollapse: true,
    },
  },
  {
    file: 'pivot-chart-buttons-hidden.xlsx',
    expected: {
      name: '[pivot-chart-buttons-hidden.xlsx]Pivot!樞紐分析表2',
      fmtId: 0,
      dropZones: {
        filter: true,
        categories: true,
        data: true,
        series: true,
        visible: false,
      },
      showExpandCollapse: true,
    },
  },
  {
    file: 'pivot-chart-buttons-filter.xlsx',
    expected: {
      name: '[pivot-chart-buttons-filter.xlsx]Pivot!樞紐分析表2',
      fmtId: 0,
      dropZones: {
        filter: true,
        categories: false,
        data: false,
        series: false,
        visible: true,
      },
      showExpandCollapse: true,
    },
  },
  {
    file: 'pivot-chart-buttons-categories.xlsx',
    expected: {
      name: '[pivot-chart-buttons-categories.xlsx]Pivot!樞紐分析表2',
      fmtId: 0,
      dropZones: {
        filter: false,
        categories: true,
        data: false,
        series: false,
        visible: true,
      },
      showExpandCollapse: true,
    },
  },
  {
    file: 'pivot-chart-buttons-data.xlsx',
    expected: {
      name: '[pivot-chart-buttons-data.xlsx]Pivot!樞紐分析表2',
      fmtId: 0,
      dropZones: {
        filter: false,
        categories: false,
        data: true,
        series: false,
        visible: true,
      },
      showExpandCollapse: true,
    },
  },
  {
    file: 'pivot-chart-buttons-series.xlsx',
    expected: {
      name: '[pivot-chart-buttons-series.xlsx]Pivot!樞紐分析表2',
      fmtId: 0,
      dropZones: {
        filter: false,
        categories: false,
        data: false,
        series: true,
        visible: true,
      },
      showExpandCollapse: true,
    },
  },
  {
    file: 'pivot-chart-buttons-no-expand.xlsx',
    expected: {
      name: '[pivot-chart-buttons-no-expand.xlsx]Pivot!樞紐分析表2',
      fmtId: 0,
      dropZones: {
        filter: true,
        categories: true,
        data: true,
        series: true,
        visible: true,
      },
      showExpandCollapse: false,
    },
  },
]

async function loadFixtureChartXml(fileName: string): Promise<string> {
  const buf = readFileSync(join(FIXTURES_DIR, fileName))
  const pkg = await OfficePackage.load(buf)
  const text = await pkg.text('xl/charts/chart1.xml')
  if (!text) throw new Error(`Missing xl/charts/chart1.xml in ${fileName}`)
  return text
}

function stripPivotMarkup(xml: string): string {
  // Strip c:pivotSource
  let stripped = xml.replace(/<c:pivotSource>[\s\S]*?<\/c:pivotSource>/g, '')
  // Strip x14 pivotOptions extension
  stripped = stripped.replace(/<c:ext\s+uri="\{781A3756-C4B2-4CAC-9D66-4F8BD8637D16\}"[\s\S]*?<\/c:ext>/g, '')
  // Strip c16 pivotOptions16 extension
  stripped = stripped.replace(/<c:ext\s+uri="\{E28EC0CA-F0BB-4C9C-879D-F8772B89E7AC\}"[\s\S]*?<\/c:ext>/g, '')
  return stripped
}

function renderChartToBuffer(model: ChartModel, w = 500, h = 350): Buffer {
  const canvas = createCanvas(w, h)
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  paintXlsxChart(model, ctx, w, h)
  const imageData = ctx.getImageData(0, 0, w, h)
  return Buffer.from(imageData.data)
}

describe('Pivot chart field buttons and ChartModel.pivot (Task 11)', () => {
  describe('Excel-saved fixture variants', () => {
    for (const { file, expected } of CASES) {
      it(`parses ChartModel.pivot correctly for ${file}`, async () => {
        const rawXml = await loadFixtureChartXml(file)
        const parsed = parseXmlOrdered(rawXml)
        const model = parse(parsed)
        expect(model).toBeDefined()
        expect(model?.pivot).toBeDefined()
        expect(model?.pivot).toEqual(expected)
      })
    }

    it('preserves localized default name with [file]Sheet! prefix', async () => {
      const rawXml = await loadFixtureChartXml('pivot-chart-buttons-all.xlsx')
      const model = parse(parseXmlOrdered(rawXml))
      expect(model?.pivot?.name).toBe('[pivot-chart-buttons-all.xlsx]Pivot!樞紐分析表2')
    })
  })

  describe('Synthetic edge cases', () => {
    const minimalChart = (extraRoot = '') => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  ${extraRoot}
  <c:chart>
    <c:autoTitleDeleted val="1"/>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:val>
            <c:numLit>
              <c:ptCount val="2"/>
              <c:pt idx="0"><c:v>10</c:v></c:pt>
              <c:pt idx="1"><c:v>20</c:v></c:pt>
            </c:numLit>
          </c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

    it('chart without pivotSource has pivot undefined and is unchanged', () => {
      const xml = minimalChart()
      const model = parse(parseXmlOrdered(xml))
      expect(model).toBeDefined()
      expect(model?.pivot).toBeUndefined()
    })

    it('chart with pivotSource but no pivotOptions element gets all zones false and showExpandCollapse false', () => {
      const xml = minimalChart(`
        <c:pivotSource>
          <c:name>TestPivotTable</c:name>
          <c:fmtId val="42"/>
        </c:pivotSource>
      `)
      const model = parse(parseXmlOrdered(xml))
      expect(model).toBeDefined()
      expect(model?.pivot).toEqual({
        name: 'TestPivotTable',
        fmtId: 42,
        dropZones: {
          filter: false,
          categories: false,
          data: false,
          series: false,
          visible: false,
        },
        showExpandCollapse: false,
      })
    })

    it('charges the pivot chart name to the chart workspace (final review M6)', () => {
      const used = (xml: string) => {
        const workspace = createChartWorkspace('desktop')
        const model = parseXlsxChart(parseXmlOrdered(xml), undefined, { workspace, profile: 'desktop', partPath: 'chart1.xml' } as unknown as ChartParseContext)!
        expect(model.refused).toBeFalsy()
        return workspace.used
      }
      const name = 'P'.repeat(100_000)
      const plain = used(minimalChart())
      expect(used(minimalChart(`<c:pivotSource><c:name>${name}</c:name><c:fmtId val="0"/></c:pivotSource>`)) - plain).toBeGreaterThanOrEqual(stringBytes(name))
    })

    it('defaults fmtId to 0 when omitted', () => {
      const xml = minimalChart(`
        <c:pivotSource>
          <c:name>DefaultFmtIdPivot</c:name>
        </c:pivotSource>
      `)
      const model = parse(parseXmlOrdered(xml))
      expect(model?.pivot?.name).toBe('DefaultFmtIdPivot')
      expect(model?.pivot?.fmtId).toBe(0)
    })
  })

  describe('Pixel-identical rendering (Ruling AI)', () => {
    for (const { file } of CASES) {
      it(`renders pixel-identical to the same chart with pivotSource stripped (${file})`, async () => {
        const rawXml = await loadFixtureChartXml(file)
        const strippedXml = stripPivotMarkup(rawXml)

        // Ensure stripped XML indeed has no pivotSource or pivot extensions
        expect(strippedXml).not.toContain('pivotSource')
        expect(strippedXml).not.toContain('781A3756-C4B2-4CAC-9D66-4F8BD8637D16')
        expect(strippedXml).not.toContain('E28EC0CA-F0BB-4C9C-879D-F8772B89E7AC')

        const modelWithPivot = parse(parseXmlOrdered(rawXml))!
        const modelStripped = parse(parseXmlOrdered(strippedXml))!

        expect(modelWithPivot.pivot).toBeDefined()
        expect(modelStripped.pivot).toBeUndefined()

        // Layout should be identical
        const dummyMeasure = {
          width: (s: string) => s.length * 7,
          lineHeight: () => 14,
        }
        const layoutWithPivot = layoutChart(modelWithPivot, 500, 350, 1, 0, dummyMeasure)
        const layoutStripped = layoutChart(modelStripped, 500, 350, 1, 0, dummyMeasure)
        expect(layoutWithPivot.plot).toEqual(layoutStripped.plot)

        // Pixel output must be byte-for-byte identical
        const bufferWithPivot = renderChartToBuffer(modelWithPivot, 500, 350)
        const bufferStripped = renderChartToBuffer(modelStripped, 500, 350)

        expect(bufferWithPivot.equals(bufferStripped)).toBe(true)
      })
    }
  })

  describe('Full XLSX workbook parse integration', () => {
    it('populates chart.pivot on drawing content when loaded through parseXlsx', async () => {
      const buf = readFileSync(join(FIXTURES_DIR, 'pivot-chart-buttons-all.xlsx'))
      const pkg = await OfficePackage.load(buf)
      const wb = await parseXlsx(pkg)
      const pivotSheet = wb.sheets.find(s => s.name === 'Pivot')
      expect(pivotSheet).toBeDefined()
      const chartDrawing = pivotSheet?.drawings?.find(d => d.content?.kind === 'xlsxChart')
      expect(chartDrawing).toBeDefined()
      const chartModel = (chartDrawing?.content as { kind: 'xlsxChart'; chart: ChartModel }).chart
      expect(chartModel.pivot).toBeDefined()
      expect(chartModel.pivot?.name).toBe('[pivot-chart-buttons-all.xlsx]Pivot!樞紐分析表2')
      expect(chartModel.pivot?.dropZones.visible).toBe(true)
    })
  })
})
