/**
 * Stage 5, Task 33 (Ruling S5-G, round-1 #9): a `timePeriod` rule's `<formula>` is classified instead of trusted.
 *
 * A formula that matches Excel's own pattern for the period keeps the Stage 1 attribute predicate. A divergent one is
 * unresolved until `formula:predicate-timeperiod-divergent` records which reading Excel follows, and never reaches the
 * attribute predicate unevidenced. The canonical patterns are read from the Excel-resaved Stage 1 clock workbooks here,
 * never typed by hand.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { attrs, getChildren, parseXml, textOf } from '../src/core/xml'
import { tokenizeBytes } from '../src/core/xml-stream'
import { OfficePackage } from '../src/core/zip'
import { createCfStreamEvents } from '../src/xlsx/cf/capture'
import { withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { createCfMemory } from '../src/xlsx/cf/limits'
import { canonicalTimePeriodFormula, parseCfSheet, type CfStylesSource } from '../src/xlsx/cf/parse'
import type { CfRule } from '../src/xlsx/cf/model'

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REPO_ROOT = join(__dirname, '..')
const FIXTURES = ['stage1-timeperiod-today-clock', 'stage1-timeperiod-today-clock-2']
const PERIODS = ['today', 'yesterday', 'tomorrow', 'last7Days', 'thisWeek', 'lastWeek', 'nextWeek', 'thisMonth', 'lastMonth', 'nextMonth']
const DXF = '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>'

const styles = (): CfStylesSource => {
  const root = parseXml(`<styleSheet xmlns="${NS}"><dxfs count="1">${DXF}</dxfs></styleSheet>`)
  return { colors: {}, dxfNodes: getChildren(getChildren(root, 'dxfs')[0]!, 'dxf') }
}
const wrap = (inner: string): string => `<worksheet xmlns="${NS}">${inner}</worksheet>`
const rule = (period: string, sqref: string, formula: string | undefined, priority = 1): string =>
  `<conditionalFormatting sqref="${sqref}"><cfRule type="timePeriod" timePeriod="${period}" priority="${priority}" dxfId="0">${formula === undefined ? '' : `<formula>${formula}</formula>`}</cfRule></conditionalFormatting>`
const escapeXml = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
const resident = (xml: string) => { const lease = parseCfSheet(parseXml(xml), styles(), createCfMemory()); return lease.value }
async function streamed(xml: string, split = 1e9) {
  const handle = createCfStreamEvents(styles(), createCfMemory(), {})
  const bytes = new TextEncoder().encode(xml)
  await tokenizeBytes((async function* () { yield bytes.subarray(0, split); yield bytes.subarray(split) })(), handle.events)
  return handle.finish().value
}
const only = (rules: readonly CfRule[]): CfRule => { expect(rules).toHaveLength(1); return rules[0]! }

type Saved = { fixture: string; period: string; sqref: string; formula: string }
/** Every `timePeriod` rule Excel wrote into the resaved Stage 1 clock workbooks, read from the saved archive. */
async function savedRules(): Promise<Saved[]> {
  const out: Saved[] = []
  for (const fixture of FIXTURES) {
    const pkg = await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf', `${fixture}.xlsx`))))
    const sheet = parseXml((await pkg.text('xl/worksheets/sheet1.xml'))!)
    for (const block of getChildren(sheet, 'conditionalFormatting')) {
      const sqref = attrs(block)['sqref'] ?? ''
      for (const cf of getChildren(block, 'cfRule')) {
        const f = getChildren(cf, 'formula')[0]
        out.push({ fixture, period: attrs(cf)['timePeriod'] ?? '', sqref, formula: f ? textOf(f) : '' })
      }
    }
  }
  return out
}
const anchorOf = (sqref: string): string => sqref.split(':')[0]!

describe('canonical timePeriod formulas (S5-G)', () => {
  it('canonical timePeriod patterns come from Excel\'s own XML', async () => {
    const saved = await savedRules()
    // The two workbooks together carry every one of the ten periods, so no pattern is typed in this test.
    expect([...new Set(saved.map(s => s.period))].sort()).toEqual([...PERIODS].sort())
    for (const s of saved) {
      expect(s.formula, `${s.fixture} ${s.period}`).not.toBe('')
      expect(canonicalTimePeriodFormula(s.period, [s.formula], anchorOf(s.sqref)), `${s.fixture} ${s.period}`).toBe(true)
      // The whole parser agrees: the rule resolves from the attribute on the resident and the streamed path.
      const xml = wrap(rule(s.period, s.sqref, escapeXml(s.formula)))
      expect(only(resident(xml).rules).definition).toMatchObject({ kind: 'timePeriod', period: s.period })
      expect(only((await streamed(xml)).rules).definition).toMatchObject({ kind: 'timePeriod', period: s.period })
    }
  })

  it('no mutation of an Excel-written pattern classifies canonical', async () => {
    for (const s of await savedRules()) {
      const anchor = anchorOf(s.sqref)
      const other = anchor === 'Z9' ? 'Y9' : 'Z9'
      const mutants: Array<[string, string]> = [
        ['clock', s.formula.replace('TODAY()', 'NOW()')],
        ['anchor', s.formula.split(anchor).join(other)],
        ['tail', `${s.formula}+0`],
        ['pinned', s.formula.split(anchor).join(`$${anchor}`)],
        ['negated', `NOT(${s.formula})`],
      ]
      for (const [name, mutant] of mutants) {
        expect(mutant, `${s.period} ${name} must differ`).not.toBe(s.formula)
        expect(canonicalTimePeriodFormula(s.period, [mutant], anchor), `${s.period} ${name}`).toBe(false)
      }
      // A canonical formula for another period is divergent for this one.
      const wrongPeriod = PERIODS[(PERIODS.indexOf(s.period) + 1) % PERIODS.length]!
      expect(canonicalTimePeriodFormula(wrongPeriod, [s.formula], anchor), `${s.period} as ${wrongPeriod}`).toBe(false)
      // Not exactly one formula, an unknown anchor (multi-range) or an unknown period is never canonical.
      expect(canonicalTimePeriodFormula(s.period, [s.formula, s.formula], anchor)).toBe(false)
      expect(canonicalTimePeriodFormula(s.period, [], anchor)).toBe(false)
      expect(canonicalTimePeriodFormula(s.period, [s.formula], undefined)).toBe(false)
      expect(canonicalTimePeriodFormula('someday', [s.formula], anchor)).toBe(false)
    }
  })

  it('whitespace and letter case are ignored, any other spelling is divergent', async () => {
    const saved = (await savedRules()).find(s => s.period === 'thisWeek')!
    const anchor = anchorOf(saved.sqref)
    const spaced = saved.formula.replace(/,/g, ' , ').replace(/</g, ' < ')
    expect(canonicalTimePeriodFormula('thisWeek', [spaced], anchor)).toBe(true)
    expect(canonicalTimePeriodFormula('thisWeek', [saved.formula.toLowerCase()], anchor)).toBe(true)
    // Excel wrote no parentheses round thisWeek's operands in either saved workbook; an equivalent spelling is not measured.
    const parenthesised = saved.formula.replace(/<=(WEEKDAY\(TODAY\(\)\)-1)/, '<=($1)')
    expect(parenthesised).not.toBe(saved.formula)
    expect(canonicalTimePeriodFormula('thisWeek', [parenthesised], anchor)).toBe(false)
  })
})

describe('divergent timePeriod formulas never reach the attribute predicate (round-1 #9)', () => {
  const divergentTodayMinus3 = rule('today', 'A2:A9', 'A2=TODAY()-3')
  const divergentFloor = rule('today', 'B2:B9', 'FLOOR(B2,1)=TODAY()-2', 2)
  const xml = wrap(divergentTodayMinus3 + divergentFloor)
  const FORMULA_READING = ['formula:predicate-timeperiod-divergent', 'formula:dialect-base', 'formula:anchor-single', 'formula:fn-today', 'formula:truth-boolean']

  it('is unresolved on the resident and streamed paths, keeps its DXF mask and has one diagnostic', async () => {
    const reference = only(resident(wrap(rule('today', 'A2:A9', undefined))).rules)
    expect(reference.definition).toMatchObject({ kind: 'timePeriod', period: 'today' })
    expect(reference.writes.size).toBeGreaterThan(0)
    for (const formula of ['A2=TODAY()-3', 'FLOOR(B2,1)=TODAY()-2', 'FLOOR(B2,1)=TODAY()-2']) {
      const x = wrap(rule('today', formula.startsWith('A2') ? 'A2:A9' : 'B2:B9', formula))
      for (const value of [resident(x), await streamed(x)]) {
        const r = only(value.rules)
        expect(r.definition, formula).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
        expect([...r.writes].sort(), formula).toEqual([...reference.writes].sort())
        expect(value.diagnostics.filter(d => d.feature === 'cf-unsupported-rule'), formula).toHaveLength(1)
        expect(value.diagnostics, formula).toHaveLength(1)
      }
    }
    expect(await streamed(xml)).toEqual(resident(xml))
  })

  it('stays unresolved under every partial formula-reading set; the exact set compiles the expression reading', async () => {
    for (const drop of FORMULA_READING.filter(c => c !== 'formula:fn-today' && c !== 'formula:truth-boolean')) {
      const partial = FORMULA_READING.filter(c => c !== drop)
      withExactCfEvidence(partial, () => {
        expect(only(resident(wrap(divergentTodayMinus3)).rules).definition, `without ${drop}`).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
      })
    }
    // The bare capability id is the formula reading only: with no grammar evidence it neither compiles nor falls back to the attribute.
    withExactCfEvidence(['formula:predicate-timeperiod-divergent'], () => {
      expect(only(resident(wrap(divergentTodayMinus3)).rules).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    })
    await withExactCfEvidence(FORMULA_READING, async () => {
      const value = resident(xml)
      const [a, b] = [value.rules.find(r => r.priority === 1)!, value.rules.find(r => r.priority === 2)!]
      expect(a.definition).toMatchObject({ kind: 'expression', sourceType: 'timePeriod' })
      expect([...a.writes].sort()).toEqual([...only(resident(wrap(rule('today', 'A2:A9', undefined))).rules).writes].sort())
      // FLOOR is neither registered nor allow-listed (round-2 #6): it stays refused whatever the evidence.
      expect(b.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
      expect(await streamed(xml)).toEqual(value)
    })
  })

  it('under the attribute reading the divergent formula keeps the attribute predicate', async () => {
    await withExactCfEvidence(['formula:predicate-timeperiod-divergent:attribute'], async () => {
      const value = resident(xml)
      expect(value.rules.map(r => r.definition)).toEqual([
        expect.objectContaining({ kind: 'timePeriod', period: 'today' }),
        expect.objectContaining({ kind: 'timePeriod', period: 'today' }),
      ])
      expect(await streamed(xml)).toEqual(value)
    })
  })

  it('an absent formula and a foreign-namespace formula keep their main-branch outcomes', async () => {
    const absent = wrap(rule('today', 'B2:B9', undefined))
    const foreign = wrap(`<conditionalFormatting sqref="B2:B9"><cfRule type="timePeriod" timePeriod="today" priority="1" dxfId="0"><formula xmlns="urn:example:other">FLOOR(B2,1)=TODAY()</formula></cfRule></conditionalFormatting>`)
    for (const evidence of [[], FORMULA_READING, ['formula:predicate-timeperiod-divergent:attribute']]) {
      await withExactCfEvidence(evidence, async () => {
        expect(only(resident(absent).rules).definition).toMatchObject({ kind: 'timePeriod', period: 'today' })
        expect(await streamed(absent)).toEqual(resident(absent))
        // Pre-existing, untouched by Task 33: only the streaming reader resolves element namespaces, so it alone sees a
        // foreign-namespace formula (withheld). The resident reader (`pkg.xml`, plain `parseXml`) cannot tell it apart and reads
        // it as a worksheet formula, which is Excel's own pattern here and so keeps the attribute predicate, exactly as on main.
        expect(only((await streamed(foreign)).rules).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
        expect(only(resident(foreign).rules).definition).toMatchObject({ kind: 'timePeriod', period: 'today' })
      })
    }
  })

  it('a canonical formula off its anchor, a multi-range rule and two formulas are divergent', () => {
    const offAnchor = wrap(rule('today', 'B2:B9', 'FLOOR(C2,1)=TODAY()'))
    const multi = wrap(rule('today', 'B2:B9 D2:D9', 'FLOOR(B2,1)=TODAY()'))
    const two = wrap(`<conditionalFormatting sqref="B2:B9"><cfRule type="timePeriod" timePeriod="today" priority="1" dxfId="0"><formula>FLOOR(B2,1)=TODAY()</formula><formula>FLOOR(B2,1)=TODAY()</formula></cfRule></conditionalFormatting>`)
    for (const x of [offAnchor, multi, two]) {
      expect(only(resident(x).rules).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    }
  })
})
