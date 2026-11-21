// @vitest-environment node
/**
 * The cfvo-nonnumeric follow-up (brief 2026-11-21, Phase 1 Task 1): a per-kind reading of `formula:cfvo-nonnumeric`.
 *
 * stage5-cfvo-formulas L/M/N matched none of the uniform readings (`withheld`, `zero`, `skipped`): Excel painted nothing
 * for a text and an error bound and read a blank one as 0. A per-kind reading names one treatment for each result kind,
 * `blank-<t>+text-<t>+error-<t>`, each `<t>` one of the uniform treatments; the three all-equal combinations keep their
 * uniform names (one name per behaviour). Production implements it behind the registry: the registered entry is
 * unchanged, so a composite name is honoured only through the test seam `injectCfFormulaReadings`, never in production.
 * Every expectation is a literal twin production itself resolves, or withheld; nothing here predicts Excel.
 */
import { describe, expect, it } from 'vitest'
import { VERIFIED_CAPABILITIES } from '../src/xlsx/cf/evidence'
import { CF_FORMULA_PRODUCTION_READINGS, cfvoNonNumericTreatment, injectCfFormulaReadings } from '../src/xlsx/cf/formulaReadings'
import { parseRules, prepare, sheetOf, shown, type Cell } from './helpers/cf-formula-rules'

const CAP = 'formula:cfvo-nonnumeric'
const TREATMENTS = ['withheld', 'zero', 'skipped'] as const
const KINDS = ['blank', 'text', 'error'] as const
type Treatment = (typeof TREATMENTS)[number]
const composite = (b: Treatment, t: Treatment, e: Treatment): string => (b === t && t === e ? b : `blank-${b}+text-${t}+error-${e}`)
const ALL = TREATMENTS.flatMap(b => TREATMENTS.flatMap(t => TREATMENTS.map(e => [b, t, e] as const)))

describe('cfvoNonNumericTreatment: the reading name decides one treatment per kind', () => {
  it('a uniform name gives its treatment to every kind', () => {
    for (const t of TREATMENTS) for (const k of KINDS) expect(cfvoNonNumericTreatment(t, k), `${t} ${k}`).toBe(t)
  })
  it('each of the 24 mixed combinations gives each kind its own treatment', () => {
    let mixed = 0
    for (const [b, t, e] of ALL) {
      const name = composite(b, t, e)
      expect([cfvoNonNumericTreatment(name, 'blank'), cfvoNonNumericTreatment(name, 'text'), cfvoNonNumericTreatment(name, 'error')], name).toEqual([b, t, e])
      if (name.includes('+')) mixed++
    }
    expect(mixed).toBe(24)
  })
  it('an all-equal composite, a reordered, partial or unknown name is no reading (fail closed)', () => {
    for (const name of ['blank-zero+text-zero+error-zero', 'text-zero+blank-withheld+error-withheld', 'blank-zero+text-withheld', 'blank-zero+text-withheld+error-ignored', 'per-kind', '', 'Zero', undefined]) {
      for (const k of KINDS) expect(cfvoNonNumericTreatment(name, k), `${String(name)} ${k}`).toBeUndefined()
    }
  })
})

const C2 = '<color rgb="FFF8696B"/><color rgb="FF63BE7B"/>'
const C3 = '<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/>'
const scale = (sqref: string, priority: number, cfvo: string, colors = C2): string =>
  `<conditionalFormatting sqref="${sqref}"><cfRule type="colorScale" priority="${priority}"><colorScale>${cfvo}${colors}</colorScale></cfRule></conditionalFormatting>`
const V = [-32, -8, 8, 16, 32]
const rowsOf = (c: string): string[] => V.map((_, i) => `${c}${i + 2}`)
const colOf = (c: string): Record<string, Cell> => Object.fromEntries(V.map((v, i) => [`${c}${i + 2}`, v]))
/** Bound sources: A1 blank (absent), A2 text, A3 the empty text a `=""` formula caches, A4 #N/A. */
const SOURCES = { blank: '$A$1', text: '$A$2', 'empty-text': '$A$3', error: '$A$4' } as const
const SOURCE_KIND = { blank: 'blank', text: 'text', 'empty-text': 'text', error: 'error' } as const
const POSITIONS = ['first', 'last', 'middle'] as const
/** Rules: one formula-bound scale per source and position (columns C..N), then the literal twins (P..S). */
function fixture() {
  const f = (src: string): string => `<cfvo type="formula" val="${src}"/>`
  const parts: string[] = []
  const cols: Record<string, { source: keyof typeof SOURCES; position: (typeof POSITIONS)[number] }> = {}
  let p = 1, c = 2
  for (const source of Object.keys(SOURCES) as Array<keyof typeof SOURCES>) for (const position of POSITIONS) {
    const col = String.fromCharCode(65 + c++)
    cols[col] = { source, position }
    const s = SOURCES[source]
    parts.push(position === 'first' ? scale(`${col}2:${col}6`, p++, `${f(s)}<cfvo type="max"/>`)
      : position === 'last' ? scale(`${col}2:${col}6`, p++, `<cfvo type="min"/>${f(s)}`)
        : scale(`${col}2:${col}6`, p++, `<cfvo type="min"/>${f(s)}<cfvo type="max"/>`, C3))
  }
  // P [num 0, max], Q [min, max], R [min, num 0], S [min, num 0, max] (3 colours).
  parts.push(scale('P2:P6', p++, '<cfvo type="num" val="0"/><cfvo type="max"/>'), scale('Q2:Q6', p++, '<cfvo type="min"/><cfvo type="max"/>'),
    scale('R2:R6', p++, '<cfvo type="min"/><cfvo type="num" val="0"/>'), scale('S2:S6', p++, '<cfvo type="min"/><cfvo type="num" val="0"/><cfvo type="max"/>', C3))
  const cells: Record<string, Cell> = { A2: 'abc', A3: '', A4: { err: '#N/A' }, ...Object.assign({}, ...[...Object.keys(cols), 'P', 'Q', 'R', 'S'].map(colOf)) }
  return { xml: parts.join(''), cells, cols }
}
/** The literal twin each treatment predicts at each position (undefined: withheld). */
const twinOf = (t: Treatment, position: (typeof POSITIONS)[number]): string | undefined =>
  t === 'withheld' ? undefined : t === 'zero' ? { first: 'P', last: 'R', middle: 'S' }[position] : position === 'middle' ? undefined : 'Q'

async function run(reading: string, inject: boolean) {
  const { xml, cells, cols } = fixture()
  const caps = [...VERIFIED_CAPABILITIES, CAP]
  const refs = [...Object.keys(cols), 'P', 'Q', 'R', 'S'].flatMap(rowsOf)
  if (inject) injectCfFormulaReadings({ [CAP]: [...TREATMENTS, reading] })
  try {
    return { cols, out: await prepare(sheetOf(parseRules(xml, caps), cells), refs, caps, { selections: { [CAP]: reading } }) }
  } finally { if (inject) injectCfFormulaReadings(null) }
}

describe('substitutedBound: each result kind follows its own treatment (behind the test seam)', () => {
  it('the registry is unchanged: the uniform selection list, no composite', () => {
    expect(CF_FORMULA_PRODUCTION_READINGS[CAP]).toEqual(['withheld', 'zero', 'skipped'])
  })
  for (const [b, t, e] of ALL) {
    const name = composite(b, t, e)
    it(`${name}: every bound shows its kind's twin, or withholds (withheld; skipped at a midpoint)`, async () => {
      const { cols, out } = await run(name, true)
      const treatment = { blank: b, text: t, error: e }
      let painted = 0
      for (const [col, { source, position }] of Object.entries(cols)) {
        const twin = twinOf(treatment[SOURCE_KIND[source]], position)
        for (const r of rowsOf(col)) {
          const got = shown(out.get(r))
          if (twin === undefined) { expect(got.withheld, `${r} (${source} ${position})`).toBe(true); continue }
          const want = shown(out.get(`${twin}${r.slice(col.length)}`))
          expect([got.withheld, want.withheld, want.fill === 'none'], `${r} twin ${twin}`).toEqual([false, false, false])
          expect(got.fill, `${r} (${source} ${position}) vs ${twin}`).toBe(want.fill)
          painted++
        }
      }
      expect(painted).toBe(Object.values(cols).filter(x => twinOf(treatment[SOURCE_KIND[x.source]], x.position) !== undefined).length * V.length)
    })
  }
})

describe('production is unchanged without the seam', () => {
  it('a composite selection is no registered reading: every non-numeric bound withholds', async () => {
    const { cols, out } = await run('blank-zero+text-withheld+error-withheld', false)
    for (const col of Object.keys(cols)) for (const r of rowsOf(col)) expect(shown(out.get(r)).withheld, r).toBe(true)
  })
  it('a uniform selection behaves as before (zero: every bound reads 0)', async () => {
    const { cols, out } = await run('zero', false)
    for (const [col, { position }] of Object.entries(cols)) for (const r of rowsOf(col)) expect(shown(out.get(r)).fill, r).toBe(shown(out.get(`${twinOf('zero', position)}${r.slice(1)}`)).fill)
  })
})
