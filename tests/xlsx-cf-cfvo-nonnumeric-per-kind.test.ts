// @vitest-environment node
/**
 * The cfvo-nonnumeric follow-up (brief 2026-11-21, Phase 1 Task 1): a per-kind reading of `formula:cfvo-nonnumeric`.
 *
 * stage5-cfvo-formulas L/M/N matched none of the uniform readings (`withheld`, `zero`, `skipped`): Excel painted nothing
 * for a text and an error bound and read a blank one as 0. A per-kind reading names one treatment for each result kind,
 * `blank-<t>+text-<t>+error-<t>`, each `<t>` one of the uniform treatments; the three all-equal combinations keep their
 * uniform names (one name per behaviour). Pre-capture ruling: each kind is also evidenced on its own id
 * (`CF_CFVO_NONNUMERIC_KIND_IDS`: blank, text, the empty string of ="", error), which, once registered, decides its kind
 * alone (`cfvoNonNumericTreatmentOf`); the umbrella decides an unregistered blank, text or error kind, never the empty
 * string. The four per-kind ids are registered since stage5-cfvo-nonnumeric-v2 was captured (16822e70: blank zero; text,
 * the empty string and error withheld); the other combinations here run behind the test seam `injectCfFormulaReadings`.
 * Every expectation is a literal twin production itself resolves, or withheld; nothing here predicts Excel.
 */
import { describe, expect, it } from 'vitest'
import { VERIFIED_CAPABILITIES } from '../src/xlsx/cf/evidence'
import { VERIFIED_STAGE5_FORMULA_SELECTIONS } from '../src/xlsx/cf/verifiedCapabilities'
import { CF_CFVO_NONNUMERIC_KIND_IDS, CF_FORMULA_PRODUCTION_READINGS, cfvoNonNumericTreatment, injectCfFormulaReadings } from '../src/xlsx/cf/formulaReadings'
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

type BoundKind = keyof typeof CF_CFVO_NONNUMERIC_KIND_IDS
const BOUND_KIND = { blank: 'blank', text: 'text', 'empty-text': 'empty-string', error: 'error' } as const satisfies Record<keyof typeof SOURCES, BoundKind>
/**
 * Prepare the fixture under a proposed registration (the test seams): `umbrella` selects a formula:cfvo-nonnumeric reading
 * (registered as the uniform list plus it), `perKind` registers and selects a per-kind id's treatment (held unless listed in
 * `unheld`); every other per-kind id is proposed unregistered (the umbrella fallback). `inject: false` runs the committed
 * registry (the four per-kind ids registered since 16822e70) with the selection only.
 */
async function run(o: { umbrella?: string; perKind?: Partial<Record<BoundKind, Treatment>>; unheld?: readonly BoundKind[]; inject?: boolean }) {
  const { xml, cells, cols } = fixture()
  const perKind = Object.entries(o.perKind ?? {}) as Array<[BoundKind, Treatment]>
  const ids = perKind.map(([k]) => CF_CFVO_NONNUMERIC_KIND_IDS[k])
  const unheldIds = (o.unheld ?? []).map(k => CF_CFVO_NONNUMERIC_KIND_IDS[k])
  const caps = [...VERIFIED_CAPABILITIES, ...(o.umbrella !== undefined ? [CAP] : []), ...perKind.map(([k]) => CF_CFVO_NONNUMERIC_KIND_IDS[k])].filter(id => !unheldIds.includes(id))
  const refs = [...Object.keys(cols), 'P', 'Q', 'R', 'S'].flatMap(rowsOf)
  const inject = o.inject !== false
  // An unregistered per-kind id is proposed as `undefined` (cfFormulaReadingEntry reads the proposal first).
  const unregistered = Object.values(CF_CFVO_NONNUMERIC_KIND_IDS).filter(id => !ids.includes(id)).map(id => [id, undefined])
  if (inject) injectCfFormulaReadings({ ...(o.umbrella !== undefined ? { [CAP]: [...TREATMENTS, o.umbrella] } : {}), ...Object.fromEntries(unregistered), ...Object.fromEntries(ids.map(id => [id, TREATMENTS])) } as Record<string, readonly string[]>)
  const selections = { ...(o.umbrella !== undefined ? { [CAP]: o.umbrella } : {}), ...Object.fromEntries(perKind.map(([k, t]) => [CF_CFVO_NONNUMERIC_KIND_IDS[k], t])) }
  try {
    return { cols, out: await prepare(sheetOf(parseRules(xml, caps), cells), refs, caps, { selections }) }
  } finally { if (inject) injectCfFormulaReadings(null) }
}
/** Every bound shows its expected treatment's twin, or withholds; returns how many cells painted. */
function expectTreatments(r: Awaited<ReturnType<typeof run>>, treatmentOf: (source: keyof typeof SOURCES) => Treatment | undefined): number {
  let painted = 0
  for (const [col, { source, position }] of Object.entries(r.cols)) {
    const t = treatmentOf(source)
    const twin = t === undefined ? undefined : twinOf(t, position)
    for (const ref of rowsOf(col)) {
      const got = shown(r.out.get(ref))
      if (twin === undefined) { expect(got.withheld, `${ref} (${source} ${position})`).toBe(true); continue }
      const want = shown(r.out.get(`${twin}${ref.slice(col.length)}`))
      expect([got.withheld, want.withheld, want.fill === 'none'], `${ref} twin ${twin}`).toEqual([false, false, false])
      expect(got.fill, `${ref} (${source} ${position}) vs ${twin}`).toBe(want.fill)
      painted++
    }
  }
  return painted
}

describe('substitutedBound: each kind follows its own per-kind id (behind the test seams)', () => {
  it('the registry: no umbrella entry (superseded), and each per-kind id the selection list (registered after 16822e70)', () => {
    expect(CF_FORMULA_PRODUCTION_READINGS[CAP]).toBeUndefined()
    for (const id of Object.values(CF_CFVO_NONNUMERIC_KIND_IDS)) expect(CF_FORMULA_PRODUCTION_READINGS[id], id).toEqual(['withheld', 'zero', 'skipped'])
  })
  // Every one of the 81 combinations of one treatment per bound kind (blank, text, the empty string, error).
  const KIND4 = ['blank', 'text', 'empty-string', 'error'] as const
  const combos: Array<Record<BoundKind, Treatment>> = []
  for (const b of TREATMENTS) for (const t of TREATMENTS) for (const s of TREATMENTS) for (const e of TREATMENTS) combos.push({ blank: b, text: t, 'empty-string': s, error: e })
  for (const c of combos) {
    it(`per-kind ${KIND4.map(k => `${k}=${c[k]}`).join(' ')}: every bound shows its kind's twin, or withholds`, async () => {
      const r = await run({ perKind: c })
      expectTreatments(r, src => c[BOUND_KIND[src]])
    })
  }
  it('a kind whose id is registered but not held withholds, whatever the others and the umbrella read', async () => {
    const r = await run({ umbrella: 'zero', perKind: { blank: 'zero', text: 'zero', 'empty-string': 'zero', error: 'zero' }, unheld: ['text'] })
    expect(expectTreatments(r, src => (src === 'text' ? undefined : 'zero'))).toBe(9 * V.length)
  })
  it('an unregistered kind falls back to the umbrella reading; the empty string never does', async () => {
    const r = await run({ umbrella: 'blank-zero+text-skipped+error-withheld', perKind: { error: 'zero' } })
    expectTreatments(r, src => (src === 'blank' ? 'zero' : src === 'text' ? 'skipped' : src === 'error' ? 'zero' : undefined))
  })
})

describe('the umbrella reading alone (no per-kind id registered)', () => {
  for (const [b, t, e] of ALL) {
    const name = composite(b, t, e)
    it(`${name}: blank, text and error follow it; the empty string withholds`, async () => {
      const r = await run({ umbrella: name })
      const treatment = { blank: b, text: t, error: e }
      expectTreatments(r, src => (src === 'empty-text' ? undefined : treatment[BOUND_KIND[src] as 'blank' | 'text' | 'error']))
    })
  }
})

describe('the committed registry and published selections (no seams)', () => {
  it('the published selections (blank zero; text, the empty string and error withheld): blank paints as 0, the others withhold', async () => {
    expect(Object.fromEntries(Object.values(CF_CFVO_NONNUMERIC_KIND_IDS).map(id => [id, VERIFIED_STAGE5_FORMULA_SELECTIONS[id]]))).toEqual({
      'formula:cfvo-nonnumeric-blank': 'zero', 'formula:cfvo-nonnumeric-text': 'withheld', 'formula:cfvo-nonnumeric-empty-string': 'withheld', 'formula:cfvo-nonnumeric-error': 'withheld',
    })
    for (const id of Object.values(CF_CFVO_NONNUMERIC_KIND_IDS)) expect(VERIFIED_CAPABILITIES.has(id), id).toBe(true)
    expect(VERIFIED_CAPABILITIES.has(CAP)).toBe(false)
    const r = await run({ perKind: { blank: 'zero', text: 'withheld', 'empty-string': 'withheld', error: 'withheld' }, inject: false })
    expect(expectTreatments(r, src => (src === 'blank' ? 'zero' : 'withheld'))).toBe(3 * V.length)
  })
  it('the umbrella is no reading once the per-kind ids are registered: an umbrella selection alone withholds every bound', async () => {
    for (const umbrella of ['zero', 'blank-zero+text-withheld+error-withheld']) expectTreatments(await run({ umbrella, inject: false }), () => undefined)
  })
})
