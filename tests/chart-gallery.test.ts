// @vitest-environment node
/**
 * The synthetic chart gallery (spec 8, Goldens; src/testdata/chart-gallery.ts) runs in CI too: its pixels are
 * Mac-only goldens (tests/golden.test.ts), but that every gallery chart is native, never a card or a fallback, is
 * checked here on every platform. Office-shaped synthetic, never Office evidence.
 */
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { parsePptx } from '../src/pptx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import { documentChartContext } from '../src/charts/context'
import { releaseChartModel } from '../src/charts/own'
import { CHARTEX_PAINTED } from '../src/charts/chartex/parse'
import { chartExGallery, chartGallery, classicGallery, type ChartDocFormat } from '../src/testdata/chart-gallery'
import type { ChartModel } from '../src/charts/model'

describe('chart gallery documents (spec 8)', () => {
  it('covers every painted ChartEx family', () => {
    expect(new Set(chartExGallery(false).map(c => c.family))).toEqual(CHARTEX_PAINTED)
  })
  it.each(['docx', 'pptx', 'xlsx'] as ChartDocFormat[])('%s: every gallery chart is native (no card, no fallback)', async format => {
    const pkg = await OfficePackage.load(await chartGallery(format)), context = documentChartContext('desktop')
    let charts: Array<ChartModel | undefined>, coverage
    if (format === 'docx') { const doc = await parseDocx(pkg, { chartContext: context }); charts = doc.sections.flatMap(s => s.paragraphs).flatMap(p => p.images ?? []).map(i => i.drawing?.kind === 'xlsxChart' ? i.drawing.chart : undefined); coverage = doc.drawingCoverage }
    else if (format === 'pptx') { const doc = await parsePptx(pkg, { chartContext: context }); charts = doc.slides.flatMap(s => s.shapes).map(s => s.content?.kind === 'xlsxChart' ? s.content.chart : undefined); coverage = doc.drawingCoverage }
    else { const doc = await parseXlsx(pkg, { chartContext: context }); charts = (doc.sheets[0].drawings ?? []).map(d => d.content?.kind === 'xlsxChart' ? d.content.chart : undefined); coverage = doc.drawingCoverage }
    const expected = classicGallery().length + chartExGallery(false).length
    expect(charts).toHaveLength(expected)
    for (const chart of charts) {
      expect(chart?.unsupported, chart?.title).toBeUndefined()
      expect(chart!.groups.length).toBeGreaterThan(0)
    }
    const original = (coverage ?? []).filter(e => e.scope === 'original')
    expect(original).toHaveLength(expected)
    expect(original.every(e => e.selectedRepresentation === 'chart')).toBe(true)
    for (const chart of charts) releaseChartModel(chart!)
    expect(context.workspace.used).toBe(0)
  })
})
