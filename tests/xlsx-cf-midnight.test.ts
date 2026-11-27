// @vitest-environment node
/**
 * The across-midnight follow-up (brief cf-midnight): once the live local day differs from the day a document was loaded
 * on, every cell of a clock-reading rule and every cell whose rule reads a clock-tainted cell is withheld
 * (`cf-unmeasured-across-midnight`) until the document is reloaded; cells that read neither are unchanged; the same day is
 * unchanged. Each pre-registered reading of `clock:across-midnight` (frozen, rules-tick, full-recalc), injected with its
 * evidence, replaces the withhold on every route: static paint, resident paint, text (copy, find, a11y), the worker, the
 * asset and print paths.
 */
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest'
import { computeDocumentClockTaint } from '../src/xlsx/formula/workbook'
import { documentLoadClockOf, documentClockTainted, markDocumentClockUnknown, pagedWorkbookClockTainted } from '../src/xlsx/formula/loadClock'
import { CF_ACROSS_MIDNIGHT_CAP, cfAcrossMidnight, cfClockDay, cfDateContextDay, injectCfAcrossMidnightReadings, type CfAcrossMidnightReading } from '../src/xlsx/cf/midnight'
import { cfEvaluationContextFor, cfDocumentEvaluationOf } from '../src/xlsx/cf/formula'
import { createCfMemory } from '../src/xlsx/cf/limits'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { GREEN, NOW, RED, REGISTERED, cellIsRule, cfBlock, exprRule, load, rcOf, refOfRc, residentLeg, staticLeg, textLeg, withPaint, workbookBytes, type Loaded } from './helpers/cf-views'

const previousTz = process.env.TZ
beforeAll(() => { process.env.TZ = 'UTC' })
afterAll(() => { if (previousTz === undefined) delete process.env.TZ; else process.env.TZ = previousTz })

const DAY = 86_400_000
const S = Math.floor(NOW / DAY) + 25569
const CAPS = REGISTERED.filter(id => !id.startsWith('formula:calc-manual-'))
const UNAVAILABLE = 'cf-unavailable|Conditional formatting unavailable'

/**
 * A2:A6 the dates S-2..S+2 under `timePeriod today` (green) and `=$A2=TODAY()` over C2:C6 (red): the clock rules.
 * E2 a `ca` TODAY() cell (saved value S-5), E3 `=E2+1` (a cached formula reading it), E4 the literal S, under the literal
 * `cellIs = S` (green): a clock-free rule over two tainted cells and a clean one. G2:G4 = 1..3 and H2 `=G2*2` under
 * `cellIs > 1` (red): no clock, no taint.
 */
const bytes = (): Uint8Array => workbookBytes({
  ...Object.fromEntries([-2, -1, 0, 1, 2].map((d, i) => [`A${i + 2}`, { date: S + d }])),
  ...Object.fromEntries([2, 3, 4, 5, 6].map(r => [`C${r}`, 0])),
  E2: { f: 'TODAY()', v: S - 5, ca: true }, E3: { f: 'E2+1', v: 0 }, E4: S,
  G2: 1, G3: 2, G4: 3, H2: { f: 'G2*2', v: 2 },
}, cfBlock('C2:C6', exprRule('$A2=TODAY()', 1, 0))
  + cfBlock('A2:A6', '<cfRule type="timePeriod" timePeriod="today" priority="2" dxfId="1"><formula>FLOOR(A2,1)=TODAY()</formula></cfRule>')
  + cfBlock('E2:E4', cellIsRule('equal', String(S), 3, 1))
  + cfBlock('G2:G4 H2', cellIsRule('greaterThan', '1', 4, 0)))
const CLOCK = ['A2', 'A3', 'A4', 'A5', 'A6', 'C2', 'C3', 'C4', 'C5', 'C6']
const TAINTED = ['E2', 'E3']
const CLEAN = ['E4', 'G2', 'G3', 'G4', 'H2']
const REFS = [...CLOCK, ...TAINTED, ...CLEAN]

/** The fill a reading paints at each ref on day `offset` after the load day (resident: E2 recalculated to S at load). */
function residentExpected(offset: number, world: 'withhold' | CfAcrossMidnightReading): Record<string, string | undefined> {
  const day = world === 'frozen' || offset === 0 ? 0 : offset
  const clockWithheld = offset !== 0 && world === 'withhold'
  const taintWithheld = offset !== 0 && (world === 'withhold' || world === 'full-recalc')
  const out: Record<string, string | undefined> = {}
  ;[-2, -1, 0, 1, 2].forEach((d, i) => {
    out[`A${i + 2}`] = clockWithheld ? undefined : d === day ? GREEN : undefined
    out[`C${i + 2}`] = clockWithheld ? undefined : d === day ? RED : undefined
  })
  out.E2 = taintWithheld ? undefined : GREEN
  out.E3 = undefined
  Object.assign(out, { E4: GREEN, G2: undefined, G3: RED, G4: RED, H2: RED })
  return out
}
/** The refs a world withholds on day `offset` (resident). */
function residentWithheld(offset: number, world: 'withhold' | CfAcrossMidnightReading): string[] {
  if (offset === 0) return []
  return world === 'withhold' ? [...CLOCK, ...TAINTED] : world === 'full-recalc' ? TAINTED : []
}

async function withReading<T>(world: 'withhold' | CfAcrossMidnightReading, fn: () => Promise<T>): Promise<T> {
  if (world === 'withhold') return withPaint(fn, CAPS)
  injectCfAcrossMidnightReadings({ [CF_ACROSS_MIDNIGHT_CAP]: world })
  try { return await withPaint(fn, [...CAPS, CF_ACROSS_MIDNIGHT_CAP]) } finally { injectCfAcrossMidnightReadings(null) }
}
const fills = (m: ReadonlyMap<string, { fill?: string }>): Record<string, string | undefined> => Object.fromEntries(REFS.map(r => [r, m.get(r)?.fill]))

describe('the load clock and the taint (formula/loadClock.ts)', () => {
  it('parse records its one clock sample and the clock-tainted cells: a volatile cell, its static readers, nothing else', async () => {
    const l = await load(bytes(), 'engine')
    const record = documentLoadClockOf(l.doc)!
    expect(record.epochMs).toBe(NOW)
    expect(documentLoadClockOf(l.sheet)).toBe(record)
    const cells = record.taint === 'all' ? 'all' : [...(record.taint.get(l.sheet.name.toLowerCase()) ?? [])].map(id => refOfRc(Math.floor(id / 16384), id % 16384)).sort()
    expect(cells).toEqual(['E2', 'E3'])
    expect(documentClockTainted(l.sheet)).toBe(true)
    // The load-time evaluation read the same sample: the `ca` TODAY() cell names the load day.
    expect(l.sheet.rows.flatMap(r => r.cells).find(c => c.ref === 'E2')!.value).toBe(S)
  })
  it('a workbook without a volatile formula has no taint; a saved-values-only one with one is tainted everywhere', async () => {
    const plain = await load(workbookBytes({ A1: 1, B1: { f: 'A1*2', v: 2 } }, cfBlock('A1:B1', cellIsRule('greaterThan', '0', 1, 0))), 'engine')
    expect(documentClockTainted(plain.doc)).toBe(false)
    expect(computeDocumentClockTaint(plain.doc)).toEqual(new Map())
    const cached = await load(bytes(), 'cached')
    expect(documentLoadClockOf(cached.doc)!.taint).toBe('all')
  })
  it('a defined name calling TODAY() taints its readers, and a later clock-reading recalculation taints every cell', async () => {
    const l = await load(bytes(), 'engine')
    l.doc.definedNames = [{ name: 'Stamp', source: 'TODAY()', baseProvenance: 'unknown' }]
    const row = l.sheet.rows.find(r => r.index === 1)!
    row.cells.push({ ref: 'J2', row: 1, col: 9, styleIndex: 0, value: 0, formula: 'Stamp+1' })
    const taint = computeDocumentClockTaint(l.doc)
    expect(taint !== 'all' && taint.get(l.sheet.name.toLowerCase())?.has(1 * 16384 + 9)).toBe(true)
    markDocumentClockUnknown(l.doc)
    expect(documentLoadClockOf(l.sheet)!.taint).toBe('all')
    expect(documentLoadClockOf(l.doc)!.epochMs).toBe(NOW)
  })
  it('the day helpers name the evaluation context\'s clockGen and the predicate context\'s day', () => {
    for (const zone of ['UTC', 'Asia/Taipei', 'America/Los_Angeles']) for (const t of [NOW, NOW + 0.4 * DAY, NOW - 0.6 * DAY]) {
      const ctx = cfEvaluationContextFor({ ...cfDocumentEvaluationOf({ calc: undefined, valuePolicy: 'cached', semantics: { dateSystem: '1900', timeZone: zone } } as never) }, () => t)
      expect(cfClockDay(t, zone, '1900')).toBe(ctx.clock.clockGen)
      expect(cfDateContextDay({ dateSystem: '1900', timeZone: zone, epochNowMs: t })).toBe(ctx.clock.clockGen)
    }
    expect(cfDateContextDay({ dateSystem: '1900', clockSerial: S + 0.25 })).toBe(Math.floor(NOW / DAY))
    expect(cfDateContextDay({ dateSystem: '1904', clockSerial: S - 1462 })).toBe(Math.floor(NOW / DAY))
  })
  it('the reading applies only with its evidence; unregistered it is the withhold', () => {
    expect(cfAcrossMidnight(() => true)).toBe('withhold')
    injectCfAcrossMidnightReadings({ [CF_ACROSS_MIDNIGHT_CAP]: 'frozen' })
    try {
      expect(cfAcrossMidnight(() => true)).toBe('frozen')
      expect(cfAcrossMidnight(() => false)).toBe('withhold')
    } finally { injectCfAcrossMidnightReadings(null) }
  })
})

describe('resident routes: static paint, resident paint, text', () => {
  for (const world of ['withhold', 'frozen', 'rules-tick', 'full-recalc'] as const) {
    it(`${world}: the load day is unchanged; the next day follows the reading`, async () => {
      const l = await load(bytes(), 'engine')
      await withReading(world, async () => {
        for (const offset of [0, 1]) {
          const options = { now: () => NOW + offset * DAY + 3_600_000 }
          const expected = residentExpected(offset, world)
          const legs = { static: await staticLeg(l, REFS, options), resident: await residentLeg(l, REFS, options), text: await textLeg(l, REFS, options), textResident: await textLeg(l, REFS, { ...options, resident: true }) }
          for (const [name, leg] of Object.entries(legs)) expect(fills(leg), `${name} day+${offset}`).toEqual(expected)
          const described = legs.text
          for (const ref of REFS) {
            const withheld = residentWithheld(offset, world).includes(ref)
            expect(described.get(ref)!.desc === UNAVAILABLE, `${ref} day+${offset} withheld`).toBe(withheld)
          }
        }
      })
    })
  }
  it('a document with no clock rule and no tainted cell is unchanged the next day', async () => {
    const l = await load(workbookBytes({ A1: 1, A2: 2, B1: { f: 'A1*2', v: 2 } }, cfBlock('A1:A2 B1', cellIsRule('greaterThan', '1', 1, 0))), 'engine')
    await withPaint(async () => {
      const day0 = fills(await residentLeg(l, ['A1', 'A2', 'B1'], { now: () => NOW }))
      const day3 = fills(await residentLeg(l, ['A1', 'A2', 'B1'], { now: () => NOW + 3 * DAY }))
      expect(day3).toEqual(day0)
      expect(day0['A2']).toBe(RED)
    }, CAPS)
  })
  it('a statistical rule over a tainted cell is withheld at every cell of its scope', async () => {
    const l = await load(workbookBytes({ A1: 1, A2: 5, A3: { f: 'TODAY()', v: S, ca: true } }, cfBlock('A1:A3', '<cfRule type="aboveAverage" priority="1" dxfId="0"/>')), 'engine')
    await withPaint(async () => {
      const day0 = fills(await residentLeg(l, ['A1', 'A2', 'A3'], { now: () => NOW }))
      expect(day0['A3']).toBe(RED)
      const text = await textLeg(l, ['A1', 'A2', 'A3'], { now: () => NOW + DAY })
      for (const r of ['A1', 'A2', 'A3']) expect(text.get(r)!.desc, r).toBe(UNAVAILABLE)
    }, CAPS)
  })
})

/** One paged session read at several instants (the worker's load clock is its first judged preparation). */
async function pagedReads(l: Loaded, instants: readonly number[]): Promise<Array<Map<string, { fill?: string; withheld: boolean }>>> {
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(l.bytes)]) }, { docId: `midnight-${Math.random()}`, profile: 'D' })
  try {
    await session.done
    const source = createPagedWorkbookSource(session), sheetId = source.sheets[0]!.sheetId, store = source.store(sheetId)
    const out: Array<Map<string, { fill?: string; withheld: boolean }>> = []
    for (const at of instants) {
      const wire = await store.readConditionalFormatting!({ rects: REFS.map(r => { const { row, col } = rcOf(r); return { r0: row, r1: row, c0: col, c1: col } }), stamp: { owner: 'm', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }, grantBytes: 8 * 1024 * 1024 }, createCfMemory(), { signal: new AbortController().signal, now: () => at } as never)
      expect(wire.kind).toBe('ready')
      const m = new Map<string, { fill?: string; withheld: boolean }>()
      // A withheld rule taints the slots it writes (the wire's `unresolvedWrites`), or the whole cell (`unresolvedReason`).
      if (wire.kind === 'ready') for (const s of wire.states) m.set(refOfRc(s.row, s.col), { fill: s.properties['fill.bg'] as string | undefined, withheld: (s.unresolvedWrites?.length ?? 0) > 0 || s.unresolvedReason !== undefined })
      out.push(m)
    }
    return out
  } finally { await session.release() }
}

describe('the worker route (paged: saved values, the first judged clock is the load clock)', () => {
  it('the paged workbook with a ca TODAY() formula is tainted everywhere; one without a volatile formula is not', async () => {
    for (const [b, tainted] of [[bytes(), true], [workbookBytes({ A1: 1, B1: { f: 'A1*2', v: 2 } }, cfBlock('A1', cellIsRule('greaterThan', '0', 1, 0))), false]] as const) {
      const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(b)]) }, { docId: `t-${Math.random()}`, profile: 'D' })
      try {
        await session.done
        expect(pagedWorkbookClockTainted(session.workbook as never)).toBe(tainted)
      } finally { await session.release() }
    }
  })
  for (const world of ['withhold', 'frozen', 'rules-tick', 'full-recalc'] as const) {
    it(`${world}: the first day paints, the next day follows the reading (every cell tainted: fail closed)`, async () => {
      const l = await load(bytes(), 'cached')
      await withReading(world, async () => {
        const [day0, day1] = await pagedReads(l, [NOW + 3_600_000, NOW + DAY + 3_600_000])
        // Saved values: E2 shows its saved S-5.
        const expected0 = { ...residentExpected(0, 'rules-tick'), E2: undefined }
        expect(Object.fromEntries(REFS.map(r => [r, day0!.get(r)?.fill])), 'day 0').toEqual(expected0)
        const expected1 = world === 'withhold' || world === 'full-recalc' ? Object.fromEntries(REFS.map(r => [r, undefined]))
          : { ...residentExpected(1, world), E2: undefined }
        expect(Object.fromEntries(REFS.map(r => [r, day1!.get(r)?.fill])), 'day 1').toEqual(expected1)
        for (const r of REFS) {
          const withheld = world === 'withhold' || (world === 'full-recalc')
          expect(day1!.get(r)?.withheld === true, `${r} withheld`).toBe(withheld)
          expect(day0!.get(r)?.withheld === true, `${r} day 0`).toBe(false)
        }
      })
    })
  }
})

describe('the asset and print paths read the wall clock (prepareSheetAssets, getPaintables)', () => {
  it('a document loaded yesterday withholds its clock and tainted cells today; its clean cells paint', async () => {
    const l = await load(bytes(), 'engine')
    const { prepareSheetAssets } = await import('../src/components/xlsx/sheetAssets')
    const at = async (): Promise<Record<string, string | undefined>> => {
      const lease = await prepareSheetAssets(l.doc, { sheetIndex: 0 })
      try { return Object.fromEntries(REFS.map(r => { const { row, col } = rcOf(r); return [r, lease.assets.conditionalFormatting?.layer.styleAt(row, col, undefined)?.fillColor] })) } finally { lease.dispose() }
    }
    vi.useFakeTimers({ toFake: ['Date'], now: NOW + 3_600_000 })
    try {
      await withPaint(async () => {
        expect(await at(), 'load day').toEqual(residentExpected(0, 'withhold'))
        vi.setSystemTime(NOW + DAY + 3_600_000)
        expect(await at(), 'next day').toEqual(residentExpected(1, 'withhold'))
      }, CAPS)
    } finally { vi.useRealTimers() }
  })
})

/* ------------------------------------------------------------------ */
/* Fix round 1 (P1): every reference form either resolves or taints.   */
/* ------------------------------------------------------------------ */

describe('the clock taint follows every reference form, or fails closed (fix round 1, P1)', () => {
  type C = { ref: string; value?: number | null; formula?: string }
  const rc = (ref: string) => rcOf(ref)
  const docOf = (sheets: Record<string, C[]>, extra: { tables?: unknown[]; definedNames?: unknown[] } = {}) => ({
    sheets: Object.entries(sheets).map(([name, cells], i) => ({
      name, sheetId: String(i + 1), cols: [], merges: [], mergeRanges: [],
      rows: [...new Set(cells.map(c => rc(c.ref).row))].sort((a, b) => a - b).map(index => ({ index, cells: cells.filter(c => rc(c.ref).row === index).map(c => ({ ref: c.ref, ...rc(c.ref), styleIndex: 0, value: c.value ?? 0, ...(c.formula ? { formula: c.formula } : {}) })) })),
    })),
    workbookSheets: Object.keys(sheets).map((name, i) => ({ sheetId: String(i + 1), name })),
    ...extra,
  }) as never
  const tainted = (doc: never, sheet = 'S'): string[] => {
    const t = computeDocumentClockTaint(doc)
    if (t === 'all') return ['all']
    return [...(t.get(sheet.toLowerCase()) ?? [])].map(id => refOfRc(Math.floor(id / 16384), id % 16384)).sort()
  }
  const table = { id: '1', name: 'T1', displayName: 'T1', sheetId: '1', partPath: 'xl/tables/table1.xml', extent: { sheetId: '1', firstCol: 0, firstRow: 0, cols: 1, rows: 3 }, headerRowCount: 1, totalsRowCount: 0, columns: [{ id: '1', name: 'Date', index: 0 }] }
  const base: C[] = [{ ref: 'A1', value: null }, { ref: 'A2', value: 5 }, { ref: 'A3', formula: 'TODAY()' }, { ref: 'A8', value: 2 }, { ref: 'B8', formula: 'A8*2' }]
  it('a qualified table reference the parser cannot read (MAX(T1[Date])) taints its formula', () => {
    expect(tainted(docOf({ S: [...base, { ref: 'B1', formula: 'MAX(T1[Date])' }] }, { tables: [table] }))).toEqual(['A3', 'B1'])
  })
  it('an unqualified table reference resolves to the table holding the cell; outside any table it taints', () => {
    // B2 sits outside the table (A1:A3): no table to resolve, so it taints; a reader of B2 follows.
    expect(tainted(docOf({ S: [...base, { ref: 'B2', formula: '[@Date]+1' }, { ref: 'C2', formula: 'B2+1' }] }, { tables: [table] }))).toEqual(['A3', 'B2', 'C2'])
    // The same reference inside a table spanning A1:B3 reads the whole table, which holds A3.
    const wide = { ...table, extent: { ...table.extent, cols: 2 } }
    expect(tainted(docOf({ S: [...base, { ref: 'B2', formula: '[@Date]+1' }] }, { tables: [wide] }))).toEqual(['A3', 'B2'])
  })
  it('a defined name over a tainted cell taints its readers; an unknown name taints', () => {
    expect(tainted(docOf({ S: [...base, { ref: 'B3', formula: 'R+1' }] }, { definedNames: [{ name: 'R', source: 'S!$A$3', baseProvenance: 'unknown' }] }))).toEqual(['A3', 'B3'])
    expect(tainted(docOf({ S: [...base, { ref: 'B4', formula: 'Missing+1' }] }))).toEqual(['A3', 'B4'])
  })
  it('a 3D reference over the tainted sheet taints; one naming unknown sheets taints', () => {
    expect(tainted(docOf({ S: base, U: [{ ref: 'B5', formula: 'SUM(S:U!A3)' }] }), 'U')).toEqual(['B5'])
    expect(tainted(docOf({ S: [...base, { ref: 'B5', formula: 'SUM(P:Q!A3)' }] }))).toEqual(['A3', 'B5'])
  })
  it('an external workbook reference taints', () => {
    expect(tainted(docOf({ S: [...base, { ref: 'B6', formula: "'[1]Other'!A1+1" }] }))).toEqual(['A3', 'B6'])
  })
  it('a spill reference to a tainted anchor taints; unions, intersections and unparsable text too; a clean reader stays clean', () => {
    expect(tainted(docOf({ S: [...base, { ref: 'B7', formula: 'SUM(A3#)' }] }))).toEqual(['A3', 'B7'])
    expect(tainted(docOf({ S: [...base, { ref: 'B9', formula: 'SUM((A2,A3))' }] }))).toEqual(['A3', 'B9'])
    expect(tainted(docOf({ S: [...base, { ref: 'B9', formula: 'SUM(A1:A3 A3:B3)' }] }))).toEqual(['A3', 'B9'])
    expect(tainted(docOf({ S: [...base, { ref: 'C9', formula: '(((' }] }))).toEqual(['A3', 'C9'])
    expect(tainted(docOf({ S: base }))).toEqual(['A3'])
  })
})
