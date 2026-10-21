import { describe, expect, it } from 'vitest'
import { list, pts } from './helpers/chart-points'
import { scanChartNumber } from '../src/charts'

const scan = (text: string) => scanChartNumber(text.length, i => text.charCodeAt(i))
// Excel-style numeric text (coordinator decision, r5): optional surrounding ASCII
// spaces, sign, digits with an optional point, optional exponent; nothing else.
const asNumber=(s:string)=>{const t=s.replace(/^ +| +$/g,'');if(!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t))return undefined;const n=Number(t);return Number.isFinite(n)?n:undefined}


describe('chart numeric scanner (streams UTF-16, never builds the text)', () => {
  it('reads Excel-style numeric text exactly; radix, Infinity, tabs/NBSP and control characters are not numbers', () => {
    const cases = ['5', ' 42 ', '\t-3.25e2\n', '+.5', '5.', '.', '', '  ', '1e', '1e+', '0x1F', '0b101', '0o17', '-0x1', 'Infinity', '-Infinity', '1e400', '1e-400', '00012.3400', '1_000', '12abc', ' 7　',
      '0'.repeat(300) + '5', '9'.repeat(400), '1.' + '0'.repeat(1000) + '1', '0.' + '0'.repeat(500) + '123e+510', '123456789012345678901234567890e-10', '-0', '0.000']
    for (const text of cases) expect(scan(text), JSON.stringify(text.slice(0, 40))).toBe(asNumber(text))
  })
  it('rounds long digit strings exactly (sticky digit) like Number()', () => {
    let seed = 7
    const random = () => (seed = (seed * 48271) % 2147483647) / 2147483647
    for (let k = 0; k < 300; k++) {
      const digits = Array.from({ length: 1 + Math.floor(random() * 1500) }, () => String(Math.floor(random() * 10))).join('')
      const text = `${random() < 0.5 ? '-' : ''}${digits.slice(0, 20)}.${digits.slice(20)}e${Math.floor(random() * 40) - 20}`
      expect(scan(text)).toBe(asNumber(text))
    }
    // A halfway case decided only by a far digit.
    const halfway = '9007199254740993.' + '0'.repeat(900) + '1'
    expect(scan(halfway)).toBe(Number(halfway))
    expect(scan(halfway)).toBe(9007199254740994)
    expect(scan('9007199254740993.' + '0'.repeat(900))).toBe(9007199254740992)
  })
  it('radix text is never a number (Excel), whatever its width', () => {
    for (const text of ['0x1F', '0x10000000000000', '0o100000000000000000', '0b101', '0x\u0011', '0x' + 'F'.repeat(14)]) expect(scan(text)).toBeUndefined()
    expect(scan('  12  ')).toBe(12)
    expect(scan('\t12')).toBeUndefined()
  })
})

describe('resident and paged charts agree on numeric text', async () => {
  const { chartWorkbookBytes } = await import('./helpers/xlsx-chart-fixture')
  const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
  const { createPagedWorkbookSource } = await import('../src/xlsx/view/pagedSource')
  const { parseXlsx } = await import('../src/xlsx/parse')
  const { OfficePackage } = await import('../src/core/zip')
  const xml = '<chartSpace><chart><plotArea><barChart><ser><val><numRef><f>Styled!A1:A3</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>'
  it('long decimal text, radix text and plain text agree (Excel: radix text is not a number)', async () => {
    const bytes = await chartWorkbookBytes(xml, { cells: [{ ref: 'A1', v: '0'.repeat(2000) + '7.5' }, { ref: 'A2', v: '0x' + 'F'.repeat(14) }, { ref: 'A3', v: 'n/a' }] })
    const resident = (await parseXlsx(await OfficePackage.load(bytes))).sheets[0].drawings![0].content as any
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'scan', profile: 'D' })
    expect((await session.done).kind).toBe('complete')
    const source = createPagedWorkbookSource(session), adapter = source.store('1') as any
    try {
      const paged = (await adapter.readDrawings({ stamp: adapter.stamp, signal: new AbortController().signal }, 'D')).drawings[0].content
      expect(list(resident.chart.groups[0].series[0].values)).toEqual(pts(7.5, undefined, undefined))
      expect(paged.chart.groups[0].series[0].values).toEqual(resident.chart.groups[0].series[0].values)
    } finally { await source.release() }
  })
})
