import { afterEach, describe, expect, it, vi } from 'vitest'
import { CHART_WORKSPACE_LIMIT, createChartWorkspace, type ChartParentLedger } from '../src/charts/workspace'
import { admitChartModel, validateAxisBinding } from '../src/charts/axisBinding'
import { documentChartContext, chartParseContext } from '../src/charts/context'
import type { ChartAxis, ChartAxisBinding, ChartGroup, ChartModel, ChartPoint, ChartResult, ChartSeries } from '../src/charts/model'
import type { ChartPoint as LegacyReexport, XlsxChartModel } from '../src/xlsx/chartModel'
import { parseXlsx } from '../src/xlsx/parse'
import { CACHED_COLUMN, chartWorkbook, openpyxlFrame } from './helpers/xlsx-chart-fixture'

// The lazy chart chunk is wrapped so the test sees the context each chart parse receives.
vi.mock('../src/charts', async importOriginal => {
  const original = await importOriginal<typeof import('../src/charts')>()
  return { ...original, parseXlsxChart: vi.fn(original.parseXlsxChart) }
})

const MiB = 1024 * 1024

/** A parent ledger with its own aggregate limit, recording every call. */
function parentLedger(limit: number) {
  const calls: string[] = []
  let used = 0
  const ledger: ChartParentLedger & { readonly used: number } = {
    get used() { return used },
    reserve(bytes, tag) { calls.push(`reserve ${tag} ${bytes}`); if (used + bytes > limit) return false; used += bytes; return true },
    release(tag, bytes) { calls.push(`release ${tag} ${bytes}`); used -= bytes },
  }
  return { ledger, calls }
}

describe('chart workspace (spec 6.4/6.5)', () => {
  it('workspace reserves, transfers and releases; refuses past the limit; never unlimited', () => {
    expect(CHART_WORKSPACE_LIMIT).toEqual({ desktop: 64 * MiB, mobile: 32 * MiB })
    const desktop = createChartWorkspace('desktop')
    const mobile = createChartWorkspace('mobile')
    expect(desktop.limit).toBe(64 * MiB)
    expect(mobile.limit).toBe(32 * MiB)
    expect(Number.isFinite(desktop.limit)).toBe(true)

    expect(mobile.reserve(20 * MiB, 'zip')).toBe(true)
    expect(mobile.used).toBe(20 * MiB)
    expect(mobile.reserve(13 * MiB, 'strings')).toBe(false)
    expect(mobile.used).toBe(20 * MiB)
    expect(mobile.reserve(12 * MiB, 'strings')).toBe(true)
    expect(mobile.used).toBe(32 * MiB)
    expect(mobile.reserve(1, 'one-more')).toBe(false)

    // Transfer moves ownership: the total is unchanged, the old tag no longer owns it.
    mobile.transfer('zip', 'model', 5 * MiB)
    expect(mobile.used).toBe(32 * MiB)
    mobile.release('zip')
    expect(mobile.used).toBe(17 * MiB)
    mobile.release('model')
    expect(mobile.used).toBe(12 * MiB)
    mobile.release('model') // idempotent
    mobile.release('never-reserved')
    expect(mobile.used).toBe(12 * MiB)
    expect(() => mobile.transfer('strings', 'model', 13 * MiB)).toThrow(RangeError)
    mobile.release('strings')
    expect(mobile.used).toBe(0)

    // Never unlimited: bad sizes are refused, a profile cannot be widened.
    for (const bad of [Infinity, NaN, -1]) expect(desktop.reserve(bad, 'bad')).toBe(false)
    expect(desktop.reserve(64 * MiB + 1, 'huge')).toBe(false)
    expect(desktop.used).toBe(0)
    expect(() => createChartWorkspace('unlimited' as never)).toThrow()
  })

  it('a reservation is refused when the parent ledger is near its limit even though the chart sublimit has room; a refusal rolls back both', () => {
    const { ledger, calls } = parentLedger(10 * MiB)
    expect(ledger.reserve(9 * MiB, 'document-pool')).toBe(true)
    const workspace = createChartWorkspace('desktop', ledger)
    // The chart sublimit (64 MiB) has room, the parent does not.
    expect(workspace.reserve(2 * MiB, 'chart')).toBe(false)
    expect(workspace.used).toBe(0)
    expect(ledger.used).toBe(9 * MiB)
    // Admitted by both: one allocation counts once in each.
    expect(workspace.reserve(MiB, 'chart')).toBe(true)
    expect(workspace.used).toBe(MiB)
    expect(ledger.used).toBe(10 * MiB)
    // Transfer stays inside the facet: the parent is untouched.
    calls.length = 0
    workspace.transfer('chart', 'model', MiB)
    expect(calls).toEqual([])
    expect(ledger.used).toBe(10 * MiB)
    // Release refunds the parent under the tag it was charged with.
    workspace.release('model')
    expect(calls).toEqual([`release chart ${MiB}`])
    expect(ledger.used).toBe(9 * MiB)

    // The chart sublimit refuses first: the parent is never asked, so nothing to roll back.
    const roomy = parentLedger(1024 * MiB)
    const small = createChartWorkspace('mobile', roomy.ledger)
    expect(small.reserve(33 * MiB, 'big')).toBe(false)
    expect(roomy.calls).toEqual([])
    expect(roomy.ledger.used).toBe(0)
  })

  it('hold keeps charges until a gated producer settles after abort', async () => {
    const { ledger } = parentLedger(100 * MiB)
    const workspace = createChartWorkspace('desktop', ledger)
    const abort = new AbortController()
    let open!: (value: string) => void
    const gate = new Promise<string>(resolve => { open = resolve })
    expect(workspace.reserve(4 * MiB, 'embedding')).toBe(true)
    const held = workspace.hold(gate, 'embedding')
    abort.abort()
    await Promise.resolve()
    await new Promise(resolve => setTimeout(resolve, 0))
    expect(abort.signal.aborted).toBe(true)
    expect(workspace.used).toBe(4 * MiB)
    expect(ledger.used).toBe(4 * MiB)
    open('inflated')
    await expect(held).resolves.toBe('inflated')
    expect(workspace.used).toBe(0)
    expect(ledger.used).toBe(0)

    // A rejected producer also refunds, and the rejection reaches the caller.
    expect(workspace.reserve(MiB, 'corrupt')).toBe(true)
    await expect(workspace.hold(Promise.reject(new Error('bad zip')), 'corrupt')).rejects.toThrow('bad zip')
    expect(workspace.used).toBe(0)
  })
})

const axis = (id: string, kind: ChartAxis['kind']): ChartAxis => ({ id, kind, deleted: false, gridlines: false })
const values = (...v: number[]): ChartPoint[] => v.map(n => ({ state: 'value', v: n }))
const series = (): ChartSeries => ({ name: 'S1', color: '4472C4', values: values(1, 2) })
const group = (family: ChartGroup['family'], binding: ChartAxisBinding): ChartGroup => ({ family, grouping: 'standard', series: [series()], binding, gapWidth: 150, overlap: 0 })
const model = (groups: ChartGroup[], axes: ChartAxis[]): ChartModel => ({ groups, axes, categories: [], fontFamily: 'Calibri', fontSizePt: 10 })

describe('chart model contracts (spec 5.1, 5.8, 6.3)', () => {
  it('point states are value, blank, error(code) and unavailable; #N/A keeps its identity', () => {
    const points: ChartPoint[] = [{ state: 'value', v: 1, fmt: '0.0' }, { state: 'blank' }, { state: 'error', code: '#N/A' }, { state: 'error', code: '#DIV/0!' }, { state: 'unavailable' }]
    expect(points.map(p => p.state)).toEqual(['value', 'blank', 'error', 'error', 'unavailable'])
    // D0: the legacy alias is the final shape (point states, not number | undefined).
    const legacy: XlsxChartModel['groups'][number]['series'][number]['values'] = [{ state: 'value', v: 1 }, { state: 'blank' }]
    expect(legacy.map(p => p.state)).toEqual(['value', 'blank'])
    const reexported: LegacyReexport = { state: 'blank' }
    expect(reexported.state).toBe('blank')
    const results: ChartResult[] = [{ kind: 'native', model: model([], []) }, { kind: 'unsupported', reason: 'x' }, { kind: 'refused', reason: 'y' }]
    expect(results.map(r => r.kind)).toEqual(['native', 'unsupported', 'refused'])
  })

  it('cartesian binding has exactly one of category or x; the validator enforces it', () => {
    const axes = [axis('10', 'category'), axis('20', 'value'), axis('30', 'value')]
    expect(validateAxisBinding('bar', { kind: 'cartesian', category: '10', value: '20' }, axes)).toBeUndefined()
    expect(validateAxisBinding('line', { kind: 'cartesian', x: '30', value: '20' }, axes)).toBeUndefined()
    expect(validateAxisBinding('bar', { kind: 'cartesian', category: '10', x: '30', value: '20' }, axes)).toMatch(/exactly one/)
    expect(validateAxisBinding('bar', { kind: 'cartesian', value: '20' }, axes)).toMatch(/exactly one/)
  })

  it('axis bindings: pie none, cartesian named roles, xy, radial, surface; invalid references rejected at parse', () => {
    const axes = [axis('1', 'category'), axis('2', 'value'), axis('3', 'value'), axis('4', 'series'), axis('5', 'date')]
    const ok: Array<[ChartGroup['family'], ChartAxisBinding]> = [
      ['pie', { kind: 'none' }],
      ['doughnut', { kind: 'none' }],
      ['bar', { kind: 'cartesian', category: '1', value: '2' }],
      ['area', { kind: 'cartesian', category: '5', value: '2' }],
      ['scatter', { kind: 'xy', x: '2', y: '3' }],
      ['bubble', { kind: 'xy', x: '3', y: '2' }],
      ['radar', { kind: 'radial', value: '2' }],
      ['contour', { kind: 'surface', category: '1', series: '4', value: '2' }],
    ]
    for (const [family, binding] of ok) expect(validateAxisBinding(family, binding, axes), `${family} ${binding.kind}`).toBeUndefined()

    const bad: Array<[ChartGroup['family'], ChartAxisBinding]> = [
      ['pie', { kind: 'cartesian', category: '1', value: '2' }], // wrong kind for the family
      ['scatter', { kind: 'cartesian', category: '1', value: '2' }],
      ['radar', { kind: 'none' }],
      ['bar', { kind: 'cartesian', category: '99', value: '2' }], // dangling reference
      ['bar', { kind: 'cartesian', category: '2', value: '3' }], // category on a value axis
      ['bar', { kind: 'cartesian', category: '1', value: '1' }], // value on a category axis
      ['scatter', { kind: 'xy', x: '2', y: '2' }], // one axis for both
      ['scatter', { kind: 'xy', x: '1', y: '2' }], // x on a category axis
      ['radar', { kind: 'radial', value: '4' }],
      ['contour', { kind: 'surface', category: '1', series: '2', value: '3' }], // series on a value axis
    ]
    for (const [family, binding] of bad) expect(validateAxisBinding(family, binding, axes), `${family} ${JSON.stringify(binding)}`).toBeTypeOf('string')

    // At parse the whole model is admitted or rejected (never drawn with a bad reference).
    const good = model([group('bar', { kind: 'cartesian', category: '1', value: '2' }), group('pie', { kind: 'none' })], axes)
    expect(admitChartModel(good)).toEqual({ kind: 'native', model: good })
    const dangling = model([group('bar', { kind: 'cartesian', category: '1', value: '2' }), group('line', { kind: 'cartesian', category: '1', value: '77' })], axes)
    const result = admitChartModel(dangling)
    expect(result.kind).toBe('unsupported')
    expect(result.kind === 'unsupported' && result.reason).toMatch(/77/)
    // Duplicate axis ids are an invalid reference too.
    expect(admitChartModel(model([group('pie', { kind: 'none' })], [axis('1', 'category'), axis('1', 'value')])).kind).toBe('unsupported')
  })
})

describe('document chart context (spec 6.4 ownership)', () => {
  afterEach(() => { vi.mocked(parseXlsxChartMock()).mockClear() })

  it('documentChartContext is created once per parse and reaches the XLSX adapter', async () => {
    const abort = new AbortController()
    const doc = documentChartContext('mobile', abort.signal)
    expect(doc.profile).toBe('mobile')
    expect(doc.signal).toBe(abort.signal)
    expect(doc.workspace.limit).toBe(32 * MiB)

    // Two charts in one workbook share one context; each gets its own part path.
    const frame = openpyxlFrame('Chart 1') + openpyxlFrame('Chart 2').replace('id="1"', 'id="2"').replace('r:id="rChart"', 'r:id="rChart2"')
    const twoCharts = () => chartWorkbook(CACHED_COLUMN, { frame, patch: zip => {
      zip.file('xl/drawings/_rels/drawing1.xml.rels', '<Relationships><Relationship Id="rChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/><Relationship Id="rChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/></Relationships>')
      zip.file('xl/charts/chart2.xml', CACHED_COLUMN)
    } })
    const contexts = async (options?: Parameters<typeof parseXlsx>[1], pkg?: Awaited<ReturnType<typeof twoCharts>>) => {
      vi.mocked(parseXlsxChartMock()).mockClear()
      const parsed = await parseXlsx(pkg ?? await twoCharts(), options)
      expect(parsed.sheets[0].drawings?.map(d => d.content?.kind)).toEqual(['xlsxChart', 'xlsxChart'])
      return vi.mocked(parseXlsxChartMock()).mock.calls.map(call => call[2]).filter(c => c !== undefined)
    }

    const supplied = await contexts({ chartContext: doc })
    expect(supplied).toHaveLength(2)
    expect(supplied.map(c => c!.workspace)).toEqual([doc.workspace, doc.workspace])
    expect(supplied[0]!.workspace).toBe(doc.workspace)
    expect(supplied.map(c => [c!.profile, c!.signal])).toEqual([['mobile', abort.signal], ['mobile', abort.signal]])
    expect(new Set(supplied.map(c => c!.partPath))).toEqual(new Set(['xl/charts/chart1.xml', 'xl/charts/chart2.xml']))

    // Without one, the lazy chunk creates exactly one per parse, on demand.
    const pkg = await twoCharts()
    const first = await contexts(undefined, pkg)
    expect(first).toHaveLength(2)
    expect(first[0]!.workspace).toBe(first[1]!.workspace)
    expect(first[0]!.workspace).not.toBe(doc.workspace)
    expect(first[0]!.profile).toBe('desktop')
    expect(first[0]!.pkg).toBe(pkg)
    // A second parse (even of the same package) gets a fresh context.
    const second = await contexts(undefined, pkg)
    expect(second[0]!.workspace).toBe(second[1]!.workspace)
    expect(second[0]!.workspace).not.toBe(first[0]!.workspace)

    // The per-chart context derives from the document one.
    const derived = chartParseContext(doc, pkg, 'xl/charts/chart9.xml', undefined, '1904')
    expect(derived).toMatchObject({ pkg, partPath: 'xl/charts/chart9.xml', profile: 'mobile', signal: abort.signal, workspace: doc.workspace, epoch: '1904' })
  })
})

let chartModule: typeof import('../src/charts') | undefined
const parseXlsxChartMock = () => chartModule!.parseXlsxChart
chartModule = await import('../src/charts')
