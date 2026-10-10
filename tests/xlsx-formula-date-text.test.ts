/**
 * C3 — DATE/TIME/TEXT evaluation + shared serial/format render integration.
 *
 * Expectation provenance per row (no fabricated native outcomes):
 *  - MEASURED   : exact accepted row from NATIVE-FAMILIES.json (1900) /
 *                 NATIVE-DATE1904.json (1904), verified:true, not in the
 *                 contract's excluded list (excluded: families id37/id41
 *                 locale-unknown slash dates, which are NOT universalized).
 *  - CONFIGURED : the documented en-US m/d/yyyy order applied ONLY when the
 *                 caller supplies locale 'en-US'; the native ledger's slash
 *                 outcomes are profile-specific observations, never a
 *                 universal order (so 2/29/1900 = #VALUE! below is the
 *                 configured en-US + real-calendar model choice).
 *  - DOC-DERIVED: follows from a measured row by the contract's fixed
 *                 parallel-serial model (SEMANTICS.md), not a new measurement.
 *  - MODEL      : explicitly unmeasured internal model behavior (kept honest).
 *
 * The 1904 fixture drives BOTH the date-function result AND the public
 * format/render text through the same resolved dateSystem.
 */
import { describe, it, expect, vi } from 'vitest'
import { evaluateFormula, evaluateFormulaInternal, formulaError } from '../src/xlsx/formula/evaluator'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import { evaluateArrayFormula } from '../src/xlsx/formula/arrays'
import { formatCommon } from '../src/xlsx/formula/format'
import { decodeSerial } from '../src/xlsx/formula/serial'
import type { EvaluationContext, ResolvedSemantics } from '../src/xlsx/formula/types'
import type { XlsxCell, XlsxDocument } from '../src/xlsx/types'

function sem(dateSystem: '1900' | '1904'): ResolvedSemantics {
  return {
    dateSystem,
    unicode: { version: 2, source: 'explicit' },
    locale: 'en-US',
    timeZone: 'UTC',
    epochNowMs: 1768478400000, // 2026-01-15T12:00:00Z (injected, not sampled)
  }
}

function ctxOf(dateSystem: '1900' | '1904', extra: Partial<EvaluationContext> = {}): EvaluationContext {
  return { semantics: sem(dateSystem), ...extra }
}

function ev(formula: string, dateSystem: '1900' | '1904' = '1900'): unknown {
  return evaluateFormula(formula, ctxOf(dateSystem))
}

// ---------------------------------------------------------------------------
// 1900 native deterministic ledger (NATIVE-FAMILIES.json, verified rows)
// ---------------------------------------------------------------------------
describe('1900 native measured DATE/TIME/TEXT ledger', () => {
  const measured: Array<[string, number | string]> = [
    ['DATE(1900,2,28)', 59], // id28
    ['DATE(1899,12,31)', 693962], // id29
    ['DATE(2024,13,0)', 45657], // id30
    ['DAY(59)', 28], // id31
    ['YEAR(-1)', '#NUM!'], // id32
    ['TIME(25,0,0)', 1 / 24], // id33
    ['TIME(-1,0,0)', '#NUM!'], // id34
    ['TIME(0,0,90)', 90 / 86400], // id35
    ['SECOND(0.999999)', 0], // id36
    ['EDATE(60,1)', 89], // id38
    ['EOMONTH(60,0)', 59], // id39
    ['DATEVALUE("2024-02-29")', 45351], // id40
    ['TIMEVALUE("12:34:56")', 0.524259259259], // id42
    ['TIMEVALUE("11:30 PM")', 0.979166666667], // id43
    ['WEEKDAY(1,1)', 1], // id44
    ['WEEKDAY(1,2)', 7], // id45
    ['WEEKDAY(1,3)', 6], // id46
    ['WEEKDAY(1,11)', 7], // id47
    ['WEEKDAY(1,12)', 6], // id48
    ['WEEKDAY(1,13)', 5], // id49
    ['WEEKDAY(1,14)', 4], // id50
    ['WEEKDAY(1,15)', 3], // id51
    ['WEEKDAY(1,16)', 2], // id52
    ['WEEKDAY(1,17)', 1], // id53
    ['TEXT(1234.567,"#,##0.00")', '1,234.57'], // id54
    ['TEXT(0.125,"0.0%")', '12.5%'], // id55
    ['TEXT(1234567,"0,,")', '1'], // id56
    ['TEXT(-12.3,"0.00;(0.00);zero")', '(12.30)'], // id57
    ['TEXT(0,"0.00;(0.00);\\z\\e\\r\\o")', 'zero'], // id58
    ['TEXT(0.5,"h:mm AM/PM")', '12:00 PM'], // id59
    ['TEXT(1.52425925925926,"yyyy-mm-dd hh:mm:ss")', '1900-01-01 12:34:56'], // id60
    ['TEXT(60,"ddd mmmm d, yyyy")', 'Wed February 29, 1900'], // id61
  ]
  for (const [formula, expected] of measured) {
    it(`${formula} = ${String(expected)} [MEASURED]`, () => {
      const actual = ev(formula)
      if (typeof expected === 'number') {
        expect(typeof actual).toBe('number')
        expect(actual as number).toBeCloseTo(expected, 9)
      } else {
        expect(actual).toBe(expected)
      }
    })
  }
})

// ---------------------------------------------------------------------------
// Contract boundary tokens (NATIVE-BOUNDARIES accepted rows + SEMANTICS table)
// ---------------------------------------------------------------------------
describe('1900 accepted boundary tokens', () => {
  it('DATE(1900,1,0)=0, DATE(1900,1,1)=1, DATE(1900,2,29)=60, DATE(1900,3,1)=61 [MEASURED]', () => {
    expect(ev('DATE(1900,1,0)')).toBe(0)
    expect(ev('DATE(1900,1,1)')).toBe(1)
    expect(ev('DATE(1900,2,29)')).toBe(60)
    expect(ev('DATE(1900,3,1)')).toBe(61)
  })

  it('YEAR/MONTH/DAY(0)=1900/1/0 and YEAR/MONTH/DAY(60)=1900/2/29 [MEASURED]', () => {
    expect(ev('YEAR(0)')).toBe(1900)
    expect(ev('MONTH(0)')).toBe(1)
    expect(ev('DAY(0)')).toBe(0)
    expect(ev('YEAR(60)')).toBe(1900)
    expect(ev('MONTH(60)')).toBe(2)
    expect(ev('DAY(60)')).toBe(29)
  })

  it('TEXT(0,"yyyy-mm-dd")="1900-01-00", TEXT(60,"yyyy-mm-dd")="1900-02-29" [MEASURED]', () => {
    expect(ev('TEXT(0,"yyyy-mm-dd")')).toBe('1900-01-00')
    expect(ev('TEXT(60,"yyyy-mm-dd")')).toBe('1900-02-29')
  })

  it('WEEKDAY(1)=1 and WEEKDAY(61)=5 with default type [MEASURED]', () => {
    expect(ev('WEEKDAY(1)')).toBe(1)
    expect(ev('WEEKDAY(61)')).toBe(5)
  })

  it('DATE(1900,2,30)=61 [MEASURED overflow]', () => {
    expect(ev('DATE(1900,2,30)')).toBe(61)
  })
})

// ---------------------------------------------------------------------------
// 1904 native ledger (NATIVE-DATE1904.json, verified rows)
// ---------------------------------------------------------------------------
describe('1904 native measured ledger', () => {
  const measured: Array<[string, number | string]> = [
    ['DATE(1904,1,1)', 0], // id0
    ['YEAR(0)', 1904], // id1
    ['MONTH(0)', 1], // id2
    ['DAY(0)', 1], // id3
    ['DATE(1904,2,29)', 59], // id4
    ['YEAR(59)', 1904], // id5
    ['MONTH(59)', 2], // id6
    ['DAY(59)', 29], // id7
    ['TEXT(0,"yyyy-mm-dd")', '1904-01-01'], // id8
    ['TEXT(59,"yyyy-mm-dd")', '1904-02-29'], // id9
    ['DATE(1900,1,1)', '#NUM!'], // id10
    ['YEAR(-1)', '#NUM!'], // id11
    ['TEXT(-1,"yyyy-mm-dd")', '-1904-01-02'], // id12
    ['EDATE(0,1)', 31], // id13
    ['EOMONTH(0,0)', 30], // id14
    ['WEEKDAY(0,1)', 6], // id15
    ['WEEKDAY(0,2)', 5], // id16
  ]
  for (const [formula, expected] of measured) {
    it(`${formula} = ${String(expected)} [MEASURED 1904]`, () => {
      const actual = ev(formula, '1904')
      if (typeof expected === 'number') {
        expect(typeof actual).toBe('number')
        expect(actual as number).toBeCloseTo(expected, 9)
      } else {
        expect(actual).toBe(expected)
      }
    })
  }
})

// ---------------------------------------------------------------------------
// Unsupported-locale text-date capability gate (NOT a universalized rule)
// ---------------------------------------------------------------------------
describe('text-date locale capability gate', () => {
  it('DATEVALUE("2/29/1900")=#VALUE! [CONFIGURED en-US order + real-calendar model choice; ledger id41 is profile-specific, NOT a universal native proof]', () => {
    expect(ev('DATEVALUE("2/29/1900")')).toBe('#VALUE!')
  })

  it('DAYS with non-en-US locale slash text is a marked NAME capability gate, not an invented order', () => {
    const ctx = ctxOf('1900', { semantics: { ...sem('1900'), locale: 'en-GB' } })
    expect(evaluateFormula('DAYS("3/1/1900","2/28/1900")', ctx)).toBe('#NAME?')
    expect(ctx.unsupportedFeatures?.has('date-text-locale-unsupported')).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// Shared serial/format renderer: parse -> real canvas render (public display)
// ---------------------------------------------------------------------------

interface RenderFixture {
  date1904: boolean
  /** numFmtId -> formatCode custom entries. */
  customNumFmts?: Record<number, string>
  /** cellXfs entries as [numFmtId]. */
  xfs: number[]
  cells: Array<{ ref: string; value?: number | string; style: number; type?: string; formula?: string }>
}

function escapeXmlAttr(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

async function buildDateFixture(f: RenderFixture): Promise<Uint8Array> {
  const JSZip = (await import('jszip')).default
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`)
  zip.file('xl/workbook.xml', `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="${f.date1904 ? '1' : '0'}"/><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`)
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`)
  const cellXml = f.cells.map((c) => `<c r="${c.ref}"${c.type ? ` t="${c.type}"` : ''} s="${c.style}">${c.formula !== undefined ? `<f>${c.formula}</f>` : ''}${c.value !== undefined ? `<v>${typeof c.value === 'string' ? escapeXmlAttr(c.value) : c.value}</v>` : ''}</c>`).join('')
  zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">${cellXml}</row></sheetData></worksheet>`)
  const numFmts = f.customNumFmts
    ? `<numFmts count="${Object.keys(f.customNumFmts).length}">${Object.entries(f.customNumFmts).map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${escapeXmlAttr(code)}"/>`).join('')}</numFmts>`
    : ''
  const xfs = f.xfs.map((id) => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`).join('')
  zip.file('xl/styles.xml', `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmts}<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${f.xfs.length}">${xfs}</cellXfs></styleSheet>`)
  return zip.generateAsync({ type: 'uint8array' })
}

async function renderFixtureTexts(f: RenderFixture): Promise<{ texts: string[]; sheetSemantics: ResolvedSemantics | undefined; sheet: import('../src/xlsx/types').XlsxSheet }> {
  const { OfficePackage } = await import('../src/core/zip')
  const { parseXlsx } = await import('../src/xlsx/parse')
  const { renderSheet } = await import('../src/xlsx/render')
  const { createCanvas } = await import('canvas')
  const doc = await parseXlsx(await OfficePackage.load(await buildDateFixture(f)))
  const sheet = doc.sheets[0]
  const canvas = createCanvas(400, 200)
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  const texts: string[] = []
  const original = ctx.fillText.bind(ctx)
  ;(ctx as unknown as { fillText: (t: string, x: number, y: number) => void }).fillText = (t, x, y) => {
    texts.push(String(t))
    original(t, x, y)
  }
  renderSheet(sheet, ctx)
  return { texts, sheetSemantics: sheet.semantics, sheet }
}

describe('parse -> shared renderer date display', () => {
  it('1900 builtin date style: serial 0 renders the day-0 token, serial 60 the fictitious day', async () => {
    const { texts, sheetSemantics } = await renderFixtureTexts({
      date1904: false,
      xfs: [0, 14],
      cells: [{ ref: 'A1', value: 0, style: 1 }, { ref: 'B1', value: 60, style: 1 }],
    })
    expect(sheetSemantics?.dateSystem).toBe('1900')
    expect(texts).toContain('1900-01-00')
    expect(texts).toContain('1900-02-29')
  })

  it('1904 fixture drives render text through the SAME resolved system (serial 0 -> 1904-01-01)', async () => {
    const { texts, sheetSemantics } = await renderFixtureTexts({
      date1904: true,
      xfs: [0, 14],
      cells: [{ ref: 'A1', value: 0, style: 1 }, { ref: 'B1', value: 59, style: 1 }],
    })
    expect(sheetSemantics?.dateSystem).toBe('1904')
    expect(texts).toContain('1904-01-01')
    expect(texts).toContain('1904-02-29')
  })

  it('authored custom numFmt formatCode is parsed and used by the renderer', async () => {
    const { texts } = await renderFixtureTexts({
      date1904: false,
      customNumFmts: { 164: 'yyyy-mm-dd' },
      xfs: [0, 164],
      cells: [{ ref: 'A1', value: 60, style: 1 }],
    })
    expect(texts).toContain('1900-02-29')
  })

  it.each([false, true])('parsed workbook -> ordinary canvas date parity, date1904=%s', async (system) => {
    const { texts, sheet } = await renderFixtureTexts({
      date1904: system,
      customNumFmts: { 164: 'yyyy-mm-dd' },
      xfs: [0, 164],
      cells: [
        { ref: 'A1', value: 0, style: 1 },
        { ref: 'B1', value: system ? 59 : 60, style: 1 },
        { ref: 'C1', style: 1, type: 'str', formula: 'TEXT(A1,"yyyy-mm-dd")' },
      ],
    })
    const first = system ? '1904-01-01' : '1900-01-00'
    const second = system ? '1904-02-29' : '1900-02-29'
    expect(sheet.rows[0].cells[2].value).toBe(first)
    expect(texts.join('')).toContain(first)
    expect(texts.join('')).toContain(second)
    expect(texts.join('')).not.toContain(system ? '1900-' : '1899-')
  })

  it('authored custom text section is used by the ordinary renderer (string value)', async () => {
    const { texts } = await renderFixtureTexts({
      date1904: false,
      customNumFmts: { 164: '0;0;0;"label "@' },
      xfs: [0, 164],
      cells: [{ ref: 'A1', value: 'hello', type: 'str', style: 1 }],
    })
    expect(texts.join('')).toContain('label hello')
  })

  it('unsupported authored custom format records a specific sheet diagnostic and keeps raw fallback', async () => {
    const { texts, sheet } = await renderFixtureTexts({
      date1904: false,
      customNumFmts: { 164: '[Red]0.00' },
      xfs: [0, 164],
      cells: [{ ref: 'A1', value: 3.14159, style: 1 }],
    })
    expect(texts.join('')).toContain('3.14')
    const diags = (sheet as unknown as { diagnostics?: Array<{ feature: string }> }).diagnostics ?? []
    expect(diags.some((d) => d.feature === 'format-unsupported')).toBe(true)
  })

  it('unrelated numeric builtin styles keep their existing product output', async () => {
    const { formatValue } = await import('../src/xlsx/render')
    expect(formatValue(3.14159, 2)).toBe('3.14')
    expect(formatValue(0.42, 9)).toBe('42%')
    expect(formatValue(4500.5, 3)).toBe('4,501')
  })
})

// ---------------------------------------------------------------------------
// Dispatch boundary controls (one-child order, typed errors, suspensions)
// ---------------------------------------------------------------------------
describe('date/text dispatch boundary controls', () => {
  it.each([
    ['DATE(2024,3,0)', 45351],
    ['DATE(2024,3,-1)', 45350],
    ['DATE(1900,3,0)', 60],
    ['DAYS(61.75,59.25)', 2.5],
    ['DATEVALUE("1899-12-31")', '#VALUE!'],
    ['DAYS(2958466,1)', '#NUM!'],
  ])('%s = %s', (formula, expected) => {
    const actual = ev(formula as string)
    if (typeof expected === 'number') {
      expect(typeof actual).toBe('number')
      expect(actual as number).toBeCloseTo(expected, 10)
    } else {
      expect(actual).toBe(expected)
    }
  })

  it.each(['YEAR(NA())', 'TIME(NA(),0,0)', 'DATE(NA(),1,1)', 'DAYS(NA(),1)', 'TEXT(NA(),"0")'])(
    'typed error propagates: %s',
    (formula) => {
      expect(evaluateFormulaInternal(formula, ctxOf('1900'))).toEqual(formulaError('#N/A'))
    },
  )

  it('lazy dead formats/dates emit no capability issue', () => {
    const unsupportedFeatures = new Set<string>()
    const ctx = ctxOf('1900', { unsupportedFeatures })
    expect(evaluateFormula('IF(FALSE,TEXT(1,"[Red]0"),DATE(1900,2,29))', ctx)).toBe(60)
    expect([...unsupportedFeatures]).toEqual([])
  })

  it('external suspension identity survives date dispatch (one read)', () => {
    const pending = { rootPending: true }
    let read = 0
    const ctx = ctxOf('1900', {
      getCellValue: () => {
        read++
        throw pending
      },
    })
    try {
      evaluateFormula('DATE(A1,1,1)', ctx)
      expect.fail('missing suspension')
    } catch (error) {
      expect(error).toBe(pending)
    }
    expect(read).toBe(1)
  })

  it('DAYS known first error never reads the later reference (one-child order)', () => {
    let reads = 0
    const ctx = ctxOf('1900', {
      getCellValue: () => {
        reads++
        throw new Error('unreachable later date argument')
      },
    })
    expect(evaluateFormula('DAYS(NA(),A1)', ctx)).toBe('#N/A')
    expect(reads).toBe(0)
  })

  it.each([
    ['TEXT(1000,"##0")', '1000'],
    ['TEXT(10000,"###0")', '10000'],
    ['TEXT(0,"0.00;(0.00);\\z\\e\\r\\o")', 'zero'],
  ])('formatter retention %s = %s', (formula, expected) => {
    expect(ev(formula)).toBe(expected)
  })
})

// ---------------------------------------------------------------------------
// Uppercase date-format grammar (APPLICATION INTEROPERABILITY PROFILE).
//
// Evidence qualification: MS-OI29500 §18.8.30 ABNF declares ASCII SMALL-letter
// date tokens, so the primary page alone does NOT prove uppercase-equivalence.
// The fixture corpus (poi-64508) authors "DD.MM.YYYY"; the finite rule below
// recognizes uppercase date-looking tokens as the same class so they never
// silently render as literal-success text. This is an application
// interoperability profile, NOT a claimed native parity rule; exact native
// fixtures remain required before any parity claim. Lowercase tokens and
// quoted literals stay distinct.
// ---------------------------------------------------------------------------
describe('uppercase date-format grammar (interoperability profile; UNVERIFIED native)', () => {
  it('formatCommon(43990,"DD.MM.YYYY") equals the lowercase result and the verified civil date 08.06.2020', () => {
    const parts = decodeSerial(43990, '1900')
    expect('civil' in parts).toBe(true)
    if (!('civil' in parts)) return
    const pad = (n: number) => String(n).padStart(2, '0')
    const verifiedCivil = `${pad(parts.civil.day)}.${pad(parts.civil.month)}.${pad(parts.civil.year)}`
    expect(verifiedCivil).toBe('08.06.2020')
    const lower = formatCommon(43990, 'dd.mm.yyyy', sem('1900'))
    const upper = formatCommon(43990, 'DD.MM.YYYY', sem('1900'))
    expect(lower.kind === 'text' && lower.text).toBe(verifiedCivil)
    expect(upper).toEqual(lower)
  })

  it('TEXT with uppercase date tokens renders the same date as lowercase', () => {
    expect(ev('TEXT(43990,"DD.MM.YYYY")')).toBe('08.06.2020')
    expect(ev('TEXT(43990,"dd.mm.yyyy")')).toBe('08.06.2020')
  })

  it('quoted uppercase date-looking text stays a literal, distinct from date tokens', () => {
    const quoted = formatCommon(43990, '0" DD.MM.YYYY"', sem('1900'))
    expect(quoted.kind === 'text' && quoted.text).toBe('43990 DD.MM.YYYY')
    const date = formatCommon(43990, 'DD.MM.YYYY', sem('1900'))
    expect(date.kind === 'text' && date.text).toBe('08.06.2020')
  })

  it('renderer honors an authored uppercase date formatCode (no literal format text)', async () => {
    const { texts } = await renderFixtureTexts({
      date1904: false,
      customNumFmts: { 164: 'DD.MM.YYYY' },
      xfs: [0, 164],
      cells: [{ ref: 'A1', value: 43990, style: 1 }],
    })
    expect(texts.join('')).toContain('08.06.2020')
    expect(texts.join('')).not.toContain('DD.MM.YYYY')
  })
})

// ---------------------------------------------------------------------------
// Workbook generation clock + manual-mode volatile suppression
// ---------------------------------------------------------------------------
function wbDoc(formula: string, value: number | string | boolean, error = false): XlsxDocument {
  const cell: XlsxCell = { ref: 'A1', col: 0, row: 0, value, formula, styleIndex: 0, hasCachedValue: true, valueIsError: error }
  return {
    sheets: [{ name: 'S', sheetId: 's', workbookIndex: 0, rows: [{ index: 0, cells: [cell] }], cols: [], merges: [], mergeRanges: [] }],
    definedNames: [],
    tables: [],
    semantics: { dateSystem: '1904', unicode: { version: 1, source: 'workbook-default' }, locale: 'en-US', timeZone: 'UTC', epochNowMs: 0 },
  }
}

describe('workbook generation clock + manual volatile suppression', () => {
  it('one injected sample shared by NOW()/NOW()/TODAY() per generation, refreshed once on the next force', () => {
    const d = wbDoc('NOW()', 77)
    d.sheets[0].rows[0].cells.push(
      { ref: 'B1', col: 1, row: 0, value: 88, styleIndex: 0, hasCachedValue: true, formula: 'NOW()' },
      { ref: 'C1', col: 2, row: 0, value: 99, styleIndex: 0, hasCachedValue: true, formula: 'TODAY()' },
    )
    let samples = 0
    const now = () => Date.UTC(2024, 0, 1, 12) + 86400000 * samples++
    const options = { forceRecalc: true, now, timeZone: 'UTC' }
    evaluateWorkbookFormulas(d, options)
    const [a, b, c] = d.sheets[0].rows[0].cells
    expect(samples).toBe(1)
    expect(a.value).toBe(b.value)
    expect(c.value).toBe(Math.floor(a.value as number))
    const previous = a.value as number
    evaluateWorkbookFormulas(d, options)
    expect(samples).toBe(2)
    expect(a.value).toBe(previous + 1)
    expect(a.value).toBe(b.value)
    expect(c.value).toBe(Math.floor(a.value as number))
  })

  it('manual cached volatile never invokes an injected clock', () => {
    const d = wbDoc('NOW()', false)
    d.calc = { calcMode: 'manual', fullCalcOnLoad: false, iterate: false, iterateCount: 100, iterateDelta: 0.001 }
    let calls = 0
    evaluateWorkbookFormulas(d, { now: () => { calls++; throw new Error('manual clock must not run') } })
    expect(calls).toBe(0)
    expect(d.sheets[0].rows[0].cells[0].value).toBe(false)
  })

  it('manual cached volatile does not sample the ambient Date.now', () => {
    const d = wbDoc('NOW()', '')
    d.calc = { calcMode: 'manual', fullCalcOnLoad: false, iterate: false, iterateCount: 100, iterateDelta: 0.001 }
    const spy = vi.spyOn(Date, 'now')
    try {
      evaluateWorkbookFormulas(d)
      expect(spy).not.toHaveBeenCalled()
      expect(d.sheets[0].rows[0].cells[0].value).toBe('')
    } finally {
      spy.mockRestore()
    }
  })
})

// ---------------------------------------------------------------------------
// C3 r2 — scalar argument slots gate an UNVERIFIED MATRIX origin BEFORE any
// type/coercion path. Ports root C3/r1/root-review/array-gates.test.ts (4 RED)
// plus the fresh-SPEC 19-slot inventory's remaining variants: DATEVALUE array
// text, TIMEVALUE array text, DAYS arg0/arg1 numeric matrix, TEXT value+format
// matrix. The gate shape is mark + attributable diagnostic + internal #NAME?
// (never a counterfeit #VALUE! that overwrites a valid cache). A GENUINE
// scalar #VALUE! (e.g. DATEVALUE(TRUE)) stays raw and distinct from the
// matrix-unsupported capability gate.
// ---------------------------------------------------------------------------
describe('scalar date/time/TEXT matrix-input gate (unverified profile)', () => {
  const GATED_SLOTS: Array<[string, string]> = [
    ['DATEVALUE array text', 'DATEVALUE({"2024-01-01","2024-01-02"})'],
    ['TIMEVALUE array text', 'TIMEVALUE({"12:00","13:00"})'],
    ['DAYS first (end) matrix', 'DAYS({61,62},59)'],
    ['DAYS second (start) matrix', 'DAYS(59,{61,62})'],
    ['TEXT format matrix', 'TEXT(1,{"0","0.0"})'],
    ['TEXT value+format matrix', 'TEXT({1},{"0"})'],
  ]

  function matrixDoc(formula: string, cached: number | string = 77): XlsxDocument {
    const cell: XlsxCell = { ref: 'A1', col: 0, row: 0, value: cached, styleIndex: 0, hasCachedValue: true, formula }
    return {
      sheets: [{ name: 'S', sheetId: 's', workbookIndex: 0, rows: [{ index: 0, cells: [cell] }], cols: [], merges: [], mergeRanges: [] }],
      definedNames: [],
      tables: [],
      semantics: sem('1900'),
    }
  }

  it.each(GATED_SLOTS)('%s: workbook retains cached 77 with attribution', (_label, formula) => {
    const d = matrixDoc(formula)
    evaluateWorkbookFormulas(d, { forceRecalc: true })
    expect(d.sheets[0].rows[0].cells[0].value).toBe(77)
    expect(d.diagnostics?.some((x) => x.feature.includes('array-input-unverified'))).toBe(true)
  })

  it.each(GATED_SLOTS)('%s: direct scalar evaluation is the internal #NAME? gate, never #VALUE!', (_label, formula) => {
    const unsupportedFeatures = new Set<string>()
    const ctx = ctxOf('1900', { unsupportedFeatures })
    expect(evaluateFormulaInternal(formula, ctx)).toEqual(formulaError('#NAME?'))
    expect([...unsupportedFeatures].some((f) => f.includes('array-input-unverified'))).toBe(true)
  })

  it.each(GATED_SLOTS)('%s: evaluateArrayFormula array mode gates identically', (_label, formula) => {
    const unsupportedFeatures = new Set<string>()
    const ctx = ctxOf('1900', { unsupportedFeatures })
    expect(evaluateArrayFormula(formula, ctx)).toEqual(formulaError('#NAME?'))
    expect([...unsupportedFeatures].some((f) => f.includes('array-input-unverified'))).toBe(true)
  })

  it('positive scalar controls keep evaluating (no gate, no mark)', () => {
    const unsupportedFeatures = new Set<string>()
    const ctx = ctxOf('1900', { unsupportedFeatures })
    expect(evaluateFormula('DATEVALUE("2024-02-29")', ctx)).toBe(45351)
    expect(evaluateFormula('TIMEVALUE("12:34:56")', ctx)).toBeCloseTo(0.524259259259, 9)
    expect(evaluateFormula('DAYS(61,59)', ctx)).toBe(2)
    expect(evaluateFormula('TEXT(1234.567,"#,##0.00")', ctx)).toBe('1,234.57')
    expect([...unsupportedFeatures]).toEqual([])
  })

  it('live branch gates; dead lazy branch records nothing', () => {
    const live = new Set<string>()
    expect(evaluateFormulaInternal('IF(TRUE,DATEVALUE({"2024-01-01","2024-01-02"}),0)', ctxOf('1900', { unsupportedFeatures: live }))).toEqual(formulaError('#NAME?'))
    expect([...live].some((f) => f.includes('array-input-unverified'))).toBe(true)

    const dead = new Set<string>()
    expect(evaluateFormula('IF(FALSE,DATEVALUE({"2024-01-01","2024-01-02"}),DATE(1900,2,29))', ctxOf('1900', { unsupportedFeatures: dead }))).toBe(60)
    expect([...dead]).toEqual([])
  })

  it('typed error passthrough wins and records no matrix gate', () => {
    const unsupportedFeatures = new Set<string>()
    const ctx = ctxOf('1900', { unsupportedFeatures })
    expect(evaluateFormulaInternal('DATEVALUE(NA())', ctx)).toEqual(formulaError('#N/A'))
    expect(evaluateFormulaInternal('TIMEVALUE(NA())', ctx)).toEqual(formulaError('#N/A'))
    expect(evaluateFormulaInternal('TEXT(NA(),"0")', ctx)).toEqual(formulaError('#N/A'))
    expect(evaluateFormulaInternal('DAYS(NA(),{1})', ctx)).toEqual(formulaError('#N/A'))
    expect([...unsupportedFeatures]).toEqual([])
  })

  it('DAYS matrix first arg never reads the later reference (one-child order preserved)', () => {
    let reads = 0
    const ctx = ctxOf('1900', {
      getCellValue: () => {
        reads++
        throw new Error('unreachable later date argument')
      },
    })
    expect(evaluateFormulaInternal('DAYS({61,62},A1)', ctx)).toEqual(formulaError('#NAME?'))
    expect(reads).toBe(0)
  })

  it('TEXT unsupported matrix VALUE resolves the capability before the later format callback is read', () => {
    let reads = 0
    const unsupportedFeatures = new Set<string>()
    const ctx = ctxOf('1900', {
      unsupportedFeatures,
      getCellValue: () => {
        reads++
        throw new Error('later format should not be read after unsupported matrix value')
      },
    })
    expect(evaluateFormula('TEXT({1},A1)', ctx)).toBe('#NAME?')
    expect(reads).toBe(0)
    expect(unsupportedFeatures.has('text-array-input-unverified')).toBe(true)
  })

  it('matrix consumed by an array-capable function is untouched by the scalar gate', () => {
    const unsupportedFeatures = new Set<string>()
    const ctx = ctxOf('1900', { unsupportedFeatures })
    expect(evaluateFormula('SUM({1,2,3})', ctx)).toBe(6)
    expect([...unsupportedFeatures]).toEqual([])
  })

  it('genuine scalar #VALUE! stays raw, distinct from the matrix capability gate', () => {
    const unsupportedFeatures = new Set<string>()
    const ctx = ctxOf('1900', { unsupportedFeatures })
    expect(evaluateFormulaInternal('DATEVALUE(TRUE)', ctx)).toEqual(formulaError('#VALUE!'))
    expect(evaluateFormulaInternal('TIMEVALUE(TRUE)', ctx)).toEqual(formulaError('#VALUE!'))
    expect(evaluateFormulaInternal('TEXT(1,TRUE)', ctx)).toEqual(formulaError('#VALUE!'))
    expect([...unsupportedFeatures]).toEqual([])

    const d = matrixDoc('DATEVALUE(TRUE)')
    evaluateWorkbookFormulas(d, { forceRecalc: true })
    expect(d.sheets[0].rows[0].cells[0].value).toBe('#VALUE!')
    expect(d.diagnostics ?? []).toEqual([])
  })
})
