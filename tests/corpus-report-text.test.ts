import { describe, expect, it } from 'vitest'
import { mkdtempSync, mkdirSync, readFileSync, writeFileSync, symlinkSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { spawnSync } from 'node:child_process'
import { xlsxBytes } from './helpers/xlsx-package'
import { createCanvas } from 'canvas'
import { getPaintables, type Paintable } from '../src/render/paint'
import { buildTextIndex, type TextIndex } from '../src/core/search'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../src/xlsx/types'
import { RECORD_TEXT } from '../src/core/text-recording'
import { hasRecordingLoss, measureXlsxTextCoverage } from '../scripts/corpus-report-text'

const context = () => createCanvas(4, 4).getContext('2d') as unknown as CanvasRenderingContext2D
const cell = (row: number, col: number, value: XlsxCell['value'], style?: XlsxCell['style']): XlsxCell => ({ ref: `${String.fromCharCode(65 + col)}${row + 1}`, row, col, value, styleIndex: 0, ...(style ? { style } : {}) })
const sheet = (cells: XlsxCell[], extra: Partial<XlsxSheet> = {}): XlsxSheet => ({ name: 'generic', rows: [...new Set(cells.map(c => c.row))].map(index => ({ index, cells: cells.filter(c => c.row === index) })), cols: [], merges: [], mergeRanges: [], ...extra })
async function measured(doc: XlsxDocument, change?: (p: Paintable[]) => Paintable[], indexChange?: (i: TextIndex) => TextIndex) {
  const pages = await getPaintables(doc)
  try {
    const actualPages = change ? change(pages) : pages
    const index = await buildTextIndex(actualPages)
    return measureXlsxTextCoverage(doc, actualPages, indexChange ? indexChange(index) : index, context())
  } finally { pages.dispose() }
}

describe('XLSX corpus recording preservation', () => {
  it('compares visible recording while retaining long hidden source strings', async () => {
    const s = sheet([cell(0, 0, 'ok'), cell(1, 0, 'hidden'.repeat(30))])
    s.rows[1].hidden = true
    const result = await measured({ sheets: [s] })
    expect(result.allSourceStringChars).toBe(182)
    expect(result.eligibleSourceStringChars).toBe(2)
    expect(result.expectedSpanChars).toBe(2)
    expect(result.coverage.string).toEqual({ expected: 2, matched: 2 })
    expect(result.excluded).toContainEqual({ reason: 'collapsed', cells: 1, sourceStringChars: 180 })
    expect(result.clipping).toBe('not-assessed')
    expect(hasRecordingLoss(result)).toBe(false)
  })
})

function omit(pages: Paintable[], predicate: (text: string, x: number, y: number) => boolean, symbol = true, extra = ''): Paintable[] {
  return pages.map(p => ({ ...p, paint(ctx) {
    const proxy = new Proxy(ctx, {
      get(target, prop) {
        if (prop === 'fillText' || prop === 'strokeText' || symbol && prop === RECORD_TEXT) return (...args: unknown[]) => {
          if (predicate(String(args[0]), Number(args[1]), Number(args[2]))) return
          const fn = Reflect.get(target, prop) as (...args: unknown[]) => unknown
          return fn?.(...args)
        }
        const value = Reflect.get(target, prop)
        return typeof value === 'function' ? value.bind(target) : value
      },
      set(target, prop, value) { return Reflect.set(target, prop, value) },
    })
    p.paint(proxy)
    if (extra) { ctx.font = '10pt "Calibri"'; ctx.fillText(extra, 400, 100) }
  } }))
}

describe('exact correspondence and failure boundaries', () => {
  it('rejects a real missing fillText despite numbers and same-text extra at another position', async () => {
    const cells = [cell(0, 0, '123'), ...Array.from({ length: 12 }, (_, i) => cell(i + 1, 1, 123))]
    let omitted = 0
    const r = await measured({ sheets: [sheet(cells)] }, p => omit(p, text => { if (text === '123' && !omitted) { omitted++; return true }; return false }, true, '123'.repeat(80)))
    expect(omitted).toBe(1)
    expect(r.coverage.string).toEqual({ expected: 3, matched: 0 })
    expect(r.coverage.other.matched).toBe(r.coverage.other.expected)
    expect(r.unmatchedRecordedChars).toBe(240)
    expect(hasRecordingLoss(r)).toBe(true)
  })
  it('rejects wrapped text omitted through the real recording symbol path', async () => {
    const r = await measured({ sheets: [sheet([cell(0, 0, 'alpha beta gamma', { numFmtId: 0, wrapText: true }), cell(1, 0, 999999)])] }, p => omit(p, (_t, _x, y) => y < 20))
    expect(r.coverage.string.expected).toBeGreaterThan(0)
    expect(r.coverage.string.matched).toBe(0)
    expect(r.coverage.other.matched).toBeGreaterThan(0)
    expect(hasRecordingLoss(r)).toBe(true)
  })
  it('does not let another sheet cover a missing full page', async () => {
    const d = { sheets: [sheet([cell(0, 0, 'lost')]), sheet([cell(0, 0, 'lost')], { name: 'second' })] }
    const r = await measured(d, p => p.slice(0, 1))
    expect(r.sheets[1].coverage.string).toEqual({ expected: 4, matched: 0 })
    expect(r.status).toBe('unavailable')
    expect(hasRecordingLoss(r)).toBe(true)
  })
  it('cannot call missing all pages legitimate empty', async () => {
    const r = await measured({ sheets: [sheet([cell(0, 0, 'lost')])] }, () => [])
    expect(r.expectedSpanChars).toBe(4)
    expect(r.status).toBe('unavailable')
    expect(hasRecordingLoss(r)).toBe(true)
  })
  it('requires repeated-key multiplicity', async () => {
    let calls = 0
    const r = await measured({ sheets: [sheet([cell(0, 0, 'same'), cell(0, 0, 'same'), cell(0, 0, 'same')])] }, p => omit(p, () => ++calls > 1))
    expect(r.coverage.string).toEqual({ expected: 12, matched: 4 })
    expect(hasRecordingLoss(r)).toBe(true)
  })
  it('reports ambiguous mixed-stratum keys with insufficient actual multiplicity', async () => {
    const style = { numFmtId: 0, horizontal: 'left' }
    let calls = 0
    const r = await measured({ sheets: [sheet([cell(0, 0, '123', style), cell(0, 0, 123, style)])] }, p => omit(p, () => ++calls === 1))
    expect(r.status).toBe('unavailable')
    expect(r.diagnostics.join(' ')).toMatch(/ambiguous/)
    expect(hasRecordingLoss(r)).toBe(true)
  })
  it.each([['majority', 3, 2, true], ['exactly half', 2, 1, false], ['minority', 3, 1, false]] as const)('keeps strict fifty percent on numeric-only %s loss', async (_name, count, missing, degraded) => {
    const r = await measured({ sheets: [sheet(Array.from({ length: count }, (_, i) => cell(i, 0, 11)))] }, p => omit(p, (_t, _x, y) => y < missing * 20))
    expect(r.coverage.other).toEqual({ expected: count * 2, matched: (count - missing) * 2 })
    expect(hasRecordingLoss(r)).toBe(degraded)
    expect(r.diagnostics.join(' ')).toMatch(/unmatched expected/)
  })
  it('reports refusal instead of using its notice as coverage', async () => {
    const r = await measured({ sheets: [sheet([cell(20000, 0, 'outside')])] })
    expect(r.status).toBe('refused')
    expect(r.allSourceStringChars).toBe(7)
    expect(hasRecordingLoss(r)).toBe(true)
    expect(r.diagnostics.join(' ')).toMatch(/static|Row|refus/)
  })
  it('rejects inconsistent index page identity and nonfinite actual spans', async () => {
    const d = { sheets: [sheet([cell(0, 0, 'ok')])] }
    const wrong = await measured(d, undefined, i => ({ pages: i.pages.map(p => ({ ...p, index: 3 })) }))
    expect(wrong.status).toBe('unavailable')
    const infinite = await measured(d, undefined, i => ({ pages: i.pages.map(p => ({ ...p, lines: p.lines.map(l => ({ ...l, spans: l.spans.map(s => ({ ...s, x: NaN })) })) })) }))
    expect(infinite.status).toBe('unavailable')
  })
  it('reports nonblank expected text that cannot be recorded as unavailable', async () => {
    const d = { sheets: [sheet([cell(0, 0, 'ok')])] }, pages = await getPaintables(d)
    try {
      const ctx = context(), index = await buildTextIndex(pages)
      const bad = new Proxy(ctx, { get(target, prop) { if (prop === 'getTransform') return () => ({ a: NaN, b: 0, c: 0, d: 1, e: 0, f: 0 }); const v = Reflect.get(target, prop); return typeof v === 'function' ? v.bind(target) : v }, set(target, prop, value) { return Reflect.set(target, prop, value) } })
      const r = measureXlsxTextCoverage(d, pages, index, bad)
      expect(r.status).toBe('unavailable')
      expect(r.diagnostics.join(' ')).toMatch(/nonblank.*record/)
    } finally { pages.dispose() }
  })
})

describe('effective static layout eligibility', () => {
  it('uses merged bounds for a hidden anchor and excludes populated nonanchors', async () => {
    const s = sheet([cell(0, 0, 'anchor'), cell(1, 0, 'covered')], { mergeRanges: [{ minRow: 0, maxRow: 1, minCol: 0, maxCol: 0 }] })
    s.rows[0].hidden = true
    const r = await measured({ sheets: [s] })
    expect(r.eligibleSourceStringChars).toBe(6)
    expect(r.excluded).toContainEqual({ reason: 'merge-covered', cells: 1, sourceStringChars: 7 })
    expect(r.coverage.string).toEqual({ expected: 6, matched: 6 })
  })
  it('honors covering merges before anchors and the last anchored bounds', async () => {
    const s = sheet([cell(0, 0, 'first'), cell(0, 1, 'covered'), cell(1, 1, 'tail')], { mergeRanges: [{ minRow: 0, maxRow: 0, minCol: 0, maxCol: 2 }, { minRow: 0, maxRow: 1, minCol: 0, maxCol: 1 }, { minRow: 0, maxRow: 1, minCol: 1, maxCol: 1 }] })
    const r = await measured({ sheets: [s] })
    expect(r.eligibleCells).toBe(1)
    expect(r.coverage.string).toEqual({ expected: 5, matched: 5 })
  })
  it.each(['rows', 'columns', 'custom zero'] as const)('keeps authored-hidden source and zero expected for %s', async mode => {
    const s = sheet([cell(0, 0, 'hidden')])
    if (mode === 'rows') s.rows[0].hidden = true
    if (mode === 'columns') s.cols = [{ min: 0, max: 0, hidden: true }]
    if (mode === 'custom zero') Object.assign(s.rows[0], { heightPt: 0, customHeight: true })
    const r = await measured({ sheets: [s] })
    expect(r.allSourceStringChars).toBe(6)
    expect(r.expectedSpanChars).toBe(0)
    expect(r.sheets[0].collapsedGeometry).toBe(true)
    expect(hasRecordingLoss(r)).toBe(false)
  })
  it('follows ordered column overrides and duplicate row metadata conflict', async () => {
    const s = sheet([cell(0, 0, 'a')], { cols: [{ min: 0, max: 0, hidden: true }, { min: 0, max: 0, widthChars: 12 }] })
    s.rows.push({ index: 0, hidden: true, cells: [cell(0, 0, 'b')] })
    const r = await measured({ sheets: [s] })
    expect(r.coverage.string).toEqual({ expected: 2, matched: 2 })
    expect(r.diagnostics.join(' ')).toMatch(/conflicting-row-metadata/)
  })
  it('places explicit refs by their own source row and partitions overflow occupancy', async () => {
    const s = sheet([cell(0, 0, 'alpha beta gamma', { numFmtId: 0, wrapText: true }), cell(0, 1, 'occupied'), cell(1, 0, 'second')])
    s.rows = [{ index: 0, cells: s.rows.flatMap(r => r.cells) }]
    const r = await measured({ sheets: [s] })
    expect(r.coverage.string.matched).toBe(r.coverage.string.expected)
    expect(r.diagnostics.join(' ')).toMatch(/inconsistent-row-coordinate/)
  })
  it('matches mixed format/wrap/rotation units while separating formatted from per-span trimming', async () => {
    const values: Array<[XlsxCell['value'], Partial<NonNullable<XlsxCell['style']>>]> = [
      [644.24222, {}], [123456789, { formatCode: '0.000000000000' }], [644.24222, { shrinkToFit: true }], [12, { formatCode: '@' }], [true, {}], ['#DIV/0!', {}], ['blank', { formatCode: ';;;' }], ['alpha beta gamma', { wrapText: true }], ['alpha\nbeta', { wrapText: true }], ['rotated', { textRotation: 45 }], ['stacked', { textRotation: 255 }], ['a', { formatCode: '"literal"@' }], [null, { formatCode: '"literal"' }],
    ]
    const s = sheet(values.map(([value, style], row) => cell(row, 0, value, { numFmtId: 0, ...style })))
    s.rows.forEach(r => { r.heightPt = 45; r.customHeight = true })
    const r = await measured({ sheets: [s] })
    expect(r.formattedExpectedChars).toBeGreaterThan(r.expectedSpanChars)
    expect(r.expectedSpanChars).toBe(r.matchedSpanChars)
    expect(r.unmatchedRecordedChars).toBe(0)
    expect(hasRecordingLoss(r)).toBe(false)
  })
})

function cli(bytes: Uint8Array, alter?: (source: string) => string) {
  const root = process.cwd(), scratch = mkdtempSync(join(tmpdir(), 'iv-report-cli-'))
  try {
    mkdirSync(join(scratch, 'scripts')); mkdirSync(join(scratch, 'corpus'))
    symlinkSync(join(root, 'src'), join(scratch, 'src'), 'dir')
    symlinkSync(join(root, 'node_modules'), join(scratch, 'node_modules'), 'dir')
    const source = readFileSync(join(root, 'scripts/corpus-report.ts'), 'utf8')
    writeFileSync(join(scratch, 'scripts/corpus-report.ts'), alter ? alter(source) : source)
    writeFileSync(join(scratch, 'scripts/corpus-report-text.ts'), readFileSync(join(root, 'scripts/corpus-report-text.ts')))
    writeFileSync(join(scratch, 'corpus/generic.xlsx'), bytes)
    return spawnSync('bun', ['scripts/corpus-report.ts'], { cwd: scratch, encoding: 'utf8' })
  } finally { rmSync(scratch, { recursive: true, force: true }) }
}
const sourceCell = '<c r="A1" t="inlineStr"><is><t>visible</t></is></c>'
describe('real corpus CLI boundaries', () => {
  it('uses recording denominator with generic hidden source retained', () => {
    const run = cli(xlsxBytes({ sheetData: `<row r="1">${sourceCell}</row><row r="2" hidden="1"><c r="A2" t="inlineStr"><is><t>${'hidden'.repeat(30)}</t></is></c></row>` }))
    expect(run.status, run.stdout + run.stderr).toBe(0)
    expect(run.stdout).toContain('source-strings=187')
    expect(run.stdout).toContain('eligible-strings=7')
    expect(run.stdout).toContain('collapsed=1cells/180source-strings')
    expect(run.stdout).toContain('matched/expected=7/7')
    expect(run.stdout).toContain('clipping=not-assessed')
  })
  it.each(['rows', 'columns'] as const)('handles real all-hidden %s without division by zero', mode => {
    const run = cli(xlsxBytes({ beforeSheetData: mode === 'columns' ? '<cols><col min="1" max="1" hidden="1"/></cols>' : '', sheetData: `<row r="1"${mode === 'rows' ? ' hidden="1"' : ''}>${sourceCell}</row>` }))
    expect(run.status, run.stdout + run.stderr).toBe(0)
    expect(run.stdout).toContain('source-strings=7')
    expect(run.stdout).toContain('matched/expected=0/0')
    expect(run.stdout).toContain('no eligible cell text')
  })
  it('keeps real static refusal non-success', () => {
    const run = cli(xlsxBytes({ sheetData: '<row r="20001"><c r="A20001" t="inlineStr"><is><t>outside</t></is></c></row>' }))
    expect(run.status, run.stdout + run.stderr).toBe(1)
    expect(run.stdout).toMatch(/static export refused/)
  })
  it('keeps index construction errors non-success', () => {
    const run = cli(xlsxBytes({ sheetData: `<row r="1">${sourceCell}</row>` }), s => s.replace("import { buildTextIndex } from '../src/core/search'", "const buildTextIndex = async (_pages: unknown) => { throw Error('deliberate index construction failure') }"))
    expect(run.status, run.stdout + run.stderr).toBe(1)
    expect(run.stdout).toContain('deliberate index construction failure')
  })
  it('degrades all visible paint omitted despite unrelated recorded extras', () => {
    const run = cli(xlsxBytes({ sheetData: `<row r="1">${sourceCell}</row>` }), s => s.replace("import { buildTextIndex } from '../src/core/search'", "import { buildTextIndex as realIndex } from '../src/core/search'\nconst buildTextIndex = (pages: any[]) => realIndex(pages.map(p => ({ ...p, paint: (ctx: CanvasRenderingContext2D) => { ctx.fillText('unrelated'.repeat(30), 400, 100) } })))"))
    expect(run.status, run.stdout + run.stderr).toBe(1)
    expect(run.stdout).toContain('matched/expected=0/7')
    expect(run.stdout).toContain('recording coverage below 50%')
  })
})


describe('additional recording boundary controls', () => {
  it('keeps an absent text index unavailable with expected denominator intact', async () => {
    const d = { sheets: [sheet([cell(0, 0, 'known')])] }, p = await getPaintables(d)
    try {
      const r = measureXlsxTextCoverage(d, p, undefined, context())
      expect(r.expectedSpanChars).toBe(5)
      expect(r.status).toBe('unavailable')
      expect(hasRecordingLoss(r)).toBe(true)
    } finally { p.dispose() }
  })
  it('keeps a fully collapsed merge empty while preserving its source', async () => {
    const s = sheet([cell(0, 0, 'anchor'), cell(1, 0, 'covered')], { mergeRanges: [{ minRow: 0, maxRow: 1, minCol: 0, maxCol: 0 }] })
    s.rows.forEach(r => { r.hidden = true })
    const r = await measured({ sheets: [s] })
    expect(r.allSourceStringChars).toBe(13)
    expect(r.expectedSpanChars).toBe(0)
    expect(hasRecordingLoss(r)).toBe(false)
  })
  it('uses positive merged width even when its anchor column is hidden', async () => {
    const s = sheet([cell(0, 0, 'anchor'), cell(0, 1, 'covered')], { cols: [{ min: 0, max: 0, hidden: true }], mergeRanges: [{ minRow: 0, maxRow: 0, minCol: 0, maxCol: 1 }] })
    const r = await measured({ sheets: [s] })
    expect(r.coverage.string).toEqual({ expected: 6, matched: 6 })
    expect(r.eligibleCells).toBe(1)
  })
  it('counts visible drawing text over hidden cells only as unmatched recording', async () => {
    const s = sheet([cell(0, 0, 'hidden')])
    s.rows[0].hidden = true
    s.drawings = [{ xEmu: 0, yEmu: 0, widthEmu: 200 * 9525, heightEmu: 50 * 9525, source: { partPath: 'drawing', treePath: 'shape', element: 'sp', representation: 'native' }, textBody: { paragraphs: [{ runs: [{ text: 'DRAWING' }], align: 'left', level: 0 }], anchor: 't', insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0, wrap: true } }]
    const r = await measured({ sheets: [s] })
    expect(r.expectedSpanChars).toBe(0)
    expect(r.recordedSpanChars).toBe(7)
    expect(r.unmatchedRecordedChars).toBe(7)
    expect(hasRecordingLoss(r)).toBe(false)
  })
  it('keeps intentionally blank formatted values distinct from hidden eligibility', async () => {
    const r = await measured({ sheets: [sheet([cell(0, 0, 'blank', { numFmtId: 0, formatCode: ';;;' })])] })
    expect(r.allSourceStringChars).toBe(5)
    expect(r.eligibleSourceStringChars).toBe(5)
    expect(r.eligibleCells).toBe(1)
    expect(r.formattedExpectedChars).toBe(0)
    expect(r.expectedSpans).toBe(0)
    expect(r.status).toBe('comparable')
  })
  it('releases the font lease even if index construction throws', () => {
    const run = cli(xlsxBytes({ sheetData: `<row r="1">${sourceCell}</row>` }), s => s.replace("import { getPaintables } from '../src/render/paint'", "import { getPaintables as realPaintables } from '../src/render/paint'\nconst getPaintables = async (doc: any) => { const pages = await realPaintables(doc), release = pages.dispose; pages.dispose = () => { console.log('font lease disposed'); release() }; return pages }").replace("import { buildTextIndex } from '../src/core/search'", "const buildTextIndex = async (_pages: unknown) => { throw Error('index failure') }"))
    expect(run.status).toBe(1)
    expect(run.stdout).toContain('font lease disposed')
    expect(run.stdout).toContain('index failure')
  })
})


it('does not match different actual text at identical position and digit geometry', async () => {
  const d = { sheets: [sheet([cell(0, 0, '123')])] }
  const r = await measured(d, pages => pages.map(p => ({ ...p, paint(ctx) {
    const proxy = new Proxy(ctx, { get(target, prop) {
      if (prop === 'fillText') return (text: string, x: number, y: number) => target.fillText(text === '123' ? '321' : text, x, y)
      const value = Reflect.get(target, prop)
      return typeof value === 'function' ? value.bind(target) : value
    }, set(target, prop, value) { return Reflect.set(target, prop, value) } })
    p.paint(proxy)
  } })))
  expect(r.recordedSpanChars).toBe(3)
  expect(r.coverage.string).toEqual({ expected: 3, matched: 0 })
  expect(r.unmatchedRecordedChars).toBe(3)
  expect(hasRecordingLoss(r)).toBe(true)
})

import { buildDocx, buildPptx } from '../src/testdata/ooxml-builders'

describe('SPEC maintained correspondence controls', () => {
  it('degrades an existing empty-recording sheet even when another sheet covers most strings globally', async () => {
    const doc = { sheets: [sheet([cell(0, 0, 'a')], { name: 'lost' }), sheet(Array.from({ length: 30 }, (_, row) => cell(row, 0, 'a')), { name: 'retained' })] }
    const result = await measured(doc, pages => {
      expect(pages).toHaveLength(2)
      return pages.map((page, i) => i === 0 ? { ...page, paint: () => {} } : page)
    }, index => { expect(index.pages).toHaveLength(2); return index })
    expect(result.status).toBe('comparable')
    expect(result.sheets[0].coverage.string).toEqual({ expected: 1, matched: 0 })
    expect(result.sheets[1].coverage.string).toEqual({ expected: 30, matched: 30 })
    expect(hasRecordingLoss(result)).toBe(true)
  })

  it.each(['width', 'x', 'y', 'fontSize'] as const)('rejects a %s-only span mismatch of 0.0001 with all other key fields intact', async field => {
    const doc = { sheets: [sheet([cell(0, 0, 'same')])] }
    const valid = await measured(doc)
    expect(valid.coverage.string).toEqual({ expected: 4, matched: 4 })
    expect(hasRecordingLoss(valid)).toBe(false)
    const mismatch = await measured(doc, undefined, index => ({ pages: index.pages.map(page => ({ ...page, lines: page.lines.map(line => ({ ...line, spans: line.spans.map(span => ({ ...span, [field]: span[field] + 0.0001 })) })) })) }))
    expect(mismatch.status).toBe('comparable')
    expect(mismatch.coverage.string).toEqual({ expected: 4, matched: 0 })
    expect(mismatch.recordedSpanChars).toBe(4)
    expect(mismatch.unmatchedRecordedChars).toBe(4)
    expect(hasRecordingLoss(mismatch)).toBe(true)
  })
})

function crossFormatCli(bytes: Uint8Array, ext: 'docx' | 'pptx', omitRecording: boolean) {
  const root = process.cwd(), scratch = mkdtempSync(join(tmpdir(), 'iv-report-cross-format-'))
  try {
    mkdirSync(join(scratch, 'scripts')); mkdirSync(join(scratch, 'corpus'))
    symlinkSync(join(root, 'src'), join(scratch, 'src'), 'dir')
    symlinkSync(join(root, 'node_modules'), join(scratch, 'node_modules'), 'dir')
    let source = readFileSync(join(root, 'scripts/corpus-report.ts'), 'utf8')
    if (omitRecording) source = source.replace("import { buildTextIndex } from '../src/core/search'", "import { buildTextIndex as realIndex } from '../src/core/search'\nconst buildTextIndex = (pages: any[]) => realIndex(pages.map(p => ({ ...p, paint: (_ctx: CanvasRenderingContext2D) => {} })))")
    writeFileSync(join(scratch, 'scripts/corpus-report.ts'), source)
    writeFileSync(join(scratch, 'scripts/corpus-report-text.ts'), readFileSync(join(root, 'scripts/corpus-report-text.ts')))
    writeFileSync(join(scratch, `corpus/generic.${ext}`), bytes)
    return spawnSync('bun', ['scripts/corpus-report.ts'], { cwd: scratch, encoding: 'utf8' })
  } finally { rmSync(scratch, { recursive: true, force: true }) }
}

describe('SPEC maintained ordinary DOCX/PPTX CLI controls', () => {
  // Packages are self-contained: fixed authored counts are independent of the
  // actual index, and outer recording omission leaves their parsed models intact.
  it.each(['paragraph DOCX', 'table DOCX', 'text PPTX'] as const)('conserves ordinary counts and detects greater-than-half loss for %s', async kind => {
    const ext = kind === 'text PPTX' ? 'pptx' : 'docx'
    const expected = kind === 'paragraph DOCX' ? 46 : kind === 'table DOCX' ? 20 : 36
    const bytes = kind === 'paragraph DOCX'
      ? await buildDocx([{ runs: [{ text: 'Paragraph recording sentinel has visible text.' }] }])
      : kind === 'table DOCX'
        ? await buildDocx([], [{ gridCols: ['8640'], rows: [{ cells: [{ paragraphs: [{ runs: [{ text: 'Table recording text' }] }] }] }] }])
        : await buildPptx([{ off: ['0', '0'], ext: ['8229600', '914400'], paragraphs: [{ runs: [{ text: 'PowerPoint recording text is visible' }] }] }])
    const normal = crossFormatCli(bytes, ext, false)
    expect(normal.status, normal.stdout + normal.stderr).toBe(0)
    expect(normal.stdout).toMatch(new RegExp(`✓ generic\\.${ext}\\s+${ext}\\s+units=\\s+1\\s+text=\\s*${expected}/${expected}\\s`))
    expect(normal.stdout).toContain('1 ok · 0 degraded · 0 legitimately empty · 0 correctly rejected')
    const omitted = crossFormatCli(bytes, ext, true)
    expect(omitted.status, omitted.stdout + omitted.stderr).toBe(1)
    expect(omitted.stdout).toMatch(new RegExp(`! generic\\.${ext}\\s+${ext}\\s+units=\\s+1\\s+text=\\s*0/${expected}\\s`))
    expect(omitted.stdout).toContain(`dropped ${expected}/${expected} model chars`)
    expect(omitted.stdout).toContain('0 ok · 1 degraded · 0 legitimately empty · 0 correctly rejected')
  })
})

import { loadOfficeFile } from '../src/components/OfficeFile'

describe('QUALITY effective geometry controls', () => {
  it.each(['negative-height', 'negative-only'] as const)('rejects real %s OOXML without relabeling invalid source as hidden', async kind => {
    const rows = kind === 'negative-height'
      ? '<row r="1" ht="15" customHeight="1"><c r="A1" t="inlineStr"><is><t>VISIBLE</t></is></c></row><row r="2" ht="-15" customHeight="1"><c r="A2" t="inlineStr"><is><t>INVALID</t></is></c></row>'
      : '<row r="1" ht="-15" customHeight="1"><c r="A1" t="inlineStr"><is><t>VISIBLE</t></is></c></row>'
    const bytes = xlsxBytes({ sheetData: rows }), doc = await loadOfficeFile(bytes) as XlsxDocument
    expect(doc.sheets[0].rows.map(row => row.heightPt)).toEqual(kind === 'negative-height' ? [15, -15] : [-15])
    expect(doc.sheets[0].rows.flatMap(row => row.cells.map(c => c.value))).toEqual(kind === 'negative-height' ? ['VISIBLE', 'INVALID'] : ['VISIBLE'])
    const result = await measured(doc)
    expect(result.allSourceStringChars).toBe(kind === 'negative-height' ? 14 : 7)
    expect(result.status).toBe('unavailable')
    expect(result.sheets[0].collapsedGeometry).toBe(false)
    expect(result.excluded.some(e => e.reason === 'collapsed')).toBe(false)
    expect(result.diagnostics.join(' ')).toMatch(/invalid effective.*geometry/)
    expect(hasRecordingLoss(result)).toBe(true)
    const run = cli(bytes)
    expect(run.status, run.stdout + run.stderr).toBe(1)
    expect(run.stdout).toContain(`source-strings=${kind === 'negative-height' ? 14 : 7}`)
    expect(run.stdout).toMatch(/invalid effective.*geometry/)
    expect(run.stdout).not.toContain('no eligible cell text')
  })

  it.each(['hidden', 'zero height'] as const)('keeps legitimate %s actual CLI geometry successful', async kind => {
    const bytes = xlsxBytes({ sheetData: `<row r="1" ${kind === 'hidden' ? 'hidden="1"' : 'ht="0" customHeight="1"'}><c r="A1" t="inlineStr"><is><t>VISIBLE</t></is></c></row>` })
    const result = await measured(await loadOfficeFile(bytes) as XlsxDocument)
    expect(result.allSourceStringChars).toBe(7)
    expect(result.status).toBe('comparable')
    expect(result.sheets[0].collapsedGeometry).toBe(true)
    expect(hasRecordingLoss(result)).toBe(false)
    const run = cli(bytes)
    expect(run.status, run.stdout + run.stderr).toBe(0)
    expect(run.stdout).toContain('source-strings=7')
    expect(run.stdout).toContain('matched/expected=0/0')
    expect(run.stdout).toContain('no eligible cell text')
  })

  it.each(['positive-total negative row', 'nonfinite row', 'negative column', 'nonfinite column'] as const)('validates individual effective axes for %s', async kind => {
    const s = sheet([cell(0, 0, 'first'), cell(1, 0, 'second')])
    if (kind === 'positive-total negative row') {
      Object.assign(s.rows[0], { heightPt: 30, customHeight: true })
      Object.assign(s.rows[1], { heightPt: -15, customHeight: true })
    }
    if (kind === 'nonfinite row') Object.assign(s.rows[0], { heightPt: NaN, customHeight: true })
    if (kind === 'negative column') s.cols = [{ min: 0, max: 0, widthChars: -10 }]
    if (kind === 'nonfinite column') s.cols = [{ min: 0, max: 0, widthChars: NaN }]
    const result = await measured({ sheets: [s] })
    expect(result.allSourceStringChars).toBe(11)
    expect(result.status).toBe('unavailable')
    expect(result.sheets[0].collapsedGeometry).toBe(false)
    expect(result.excluded.some(e => e.reason === 'collapsed')).toBe(false)
    expect(result.diagnostics.join(' ')).toMatch(/invalid effective.*geometry/)
    expect(hasRecordingLoss(result)).toBe(true)
  })
})
