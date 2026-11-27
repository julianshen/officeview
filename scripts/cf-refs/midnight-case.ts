/**
 * The across-midnight follow-up (brief cf-midnight): the layout of `stage1-clock-across-midnight` and every prediction,
 * fixed before any capture. One source of truth for the case template (cases.ts), the generator (generate.ts), the capture
 * (capture-midnight.ts) and its decision (`decideAcrossMidnight`).
 *
 * The workbook is built for one night: D is the local date (Asia/Taipei) of the evening the capture starts; every literal
 * date and every literal cellIs operand is D plus an offset. A prediction never names D: it is a function of the day the
 * cells name and the day the rules read, each as an offset from D (0 or 1), so it holds for whichever night is built.
 *
 * Instants (capture-midnight.ts): 1 at about 23:55 on D, 2 at about 00:05 on D+1 with no edit, recalculation, scroll or
 * window change, 3 after an explicit Calculate Now. The readings of `clock:across-midnight` at instant 2:
 * - `frozen`: cells D, rules D; `rules-tick`: cells D, rules D+1; `full-recalc`: cells D+1, rules D+1.
 * Instant 1 is cells D, rules D under every reading; instant 3 is cells D+1, rules D+1 under every reading.
 */
import type { CfDxfExpectation, CfPartMarker } from './cases'

export const MIDNIGHT_CASE_ID = 'stage1-clock-across-midnight'
export const MIDNIGHT_CAP = 'clock:across-midnight'
export const MIDNIGHT_READINGS = ['frozen', 'rules-tick', 'full-recalc'] as const
export type MidnightReading = (typeof MIDNIGHT_READINGS)[number]
export type MidnightToken = 'red' | 'green' | 'yellow' | 'none'
/** The day the cells name and the day the rules read, as offsets from D. */
export type MidnightWorld = Readonly<{ cellDay: 0 | 1; ruleDay: 0 | 1 }>
export const MIDNIGHT_WORLDS: Readonly<Record<MidnightReading, MidnightWorld>> = Object.freeze({
  frozen: { cellDay: 0, ruleDay: 0 },
  'rules-tick': { cellDay: 0, ruleDay: 1 },
  'full-recalc': { cellDay: 1, ruleDay: 1 },
})
export const MIDNIGHT_INSTANT1: MidnightWorld = Object.freeze({ cellDay: 0, ruleDay: 0 })
export const MIDNIGHT_INSTANT3: MidnightWorld = Object.freeze({ cellDay: 1, ruleDay: 1 })

/** The three dxfs of the Stage 1 default table (opaque fills), by index. */
export const MIDNIGHT_DXFS: readonly [CfDxfExpectation, CfDxfExpectation, CfDxfExpectation] = [
  { fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true },
  { fillRgb: 'FFC6EFCE', fontRgb: 'FF006100' },
  { fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500' },
]
const TOKEN_OF_DXF: readonly MidnightToken[] = ['red', 'green', 'yellow']
export const MIDNIGHT_TOKEN_FILL: Readonly<Record<Exclude<MidnightToken, 'none'>, string>> = { red: 'FFFFC7CE', green: 'FFC6EFCE', yellow: 'FFFFEB9C' }

/** A cell's content: a literal date (D + offset), a number, or a `TODAY()+k` formula, saved at its load-day value D + k (so a
 * route that shows saved values, and the viewer before any resave, name the load day; Excel recalculates it on open). */
export type MidnightValue = Readonly<{ date: number }> | Readonly<{ today: number }> | Readonly<{ number: number }>
export type MidnightRule =
  | Readonly<{ kind: 'timePeriod'; period: 'today' | 'yesterday' | 'tomorrow' }>
  | Readonly<{ kind: 'expression-today' }>
  | Readonly<{ kind: 'cellIs'; operator: 'equal' | 'greaterThan' | 'lessThan'; operand: Readonly<{ date: number }> | Readonly<{ number: number }> }>
/**
 * Groups: `literal-clock` (literal dates under clock rules: separates frozen from the others), `cell-literal` (`TODAY()`
 * cells under literal cellIs: separates full-recalc), `cell-clock` (`TODAY()` cells under clock rules: separates rules-tick),
 * `clock-control` (far literal dates under a clock rule: none in every world), `free` (no clock, no clock cell: the
 * clock-free control region).
 */
export type MidnightGroup = 'literal-clock' | 'cell-literal' | 'cell-clock' | 'clock-control' | 'free'
export interface MidnightCell {
  readonly ref: string
  readonly controlRef: string
  readonly group: MidnightGroup
  readonly value: MidnightValue
  /** Index into `MIDNIGHT_RULES`. */
  readonly rule: number
}
export interface MidnightRuleSpec {
  readonly sqref: string
  readonly rule: MidnightRule
  readonly dxfId: 0 | 1 | 2
  readonly capabilityId: string
}

const DATES = [-1, 0, 1, 2]
/** Rows 2..5 hold the designated cells; rows 8..11 their unformatted same-content controls. */
const CONTROL_OFFSET = 6
export const MIDNIGHT_RULES: readonly MidnightRuleSpec[] = [
  { sqref: 'B2:B5', rule: { kind: 'timePeriod', period: 'today' }, dxfId: 0, capabilityId: 'timePeriod:today' },
  { sqref: 'C2:C5', rule: { kind: 'timePeriod', period: 'yesterday' }, dxfId: 1, capabilityId: 'timePeriod:yesterday' },
  { sqref: 'D2:D5', rule: { kind: 'timePeriod', period: 'tomorrow' }, dxfId: 2, capabilityId: 'timePeriod:tomorrow' },
  { sqref: 'E2:E5', rule: { kind: 'expression-today' }, dxfId: 0, capabilityId: 'formula:fn-today' },
  { sqref: 'F2', rule: { kind: 'cellIs', operator: 'equal', operand: { date: 0 } }, dxfId: 1, capabilityId: 'cellIs:equal-numeric' },
  { sqref: 'F3', rule: { kind: 'cellIs', operator: 'equal', operand: { date: 1 } }, dxfId: 1, capabilityId: 'cellIs:equal-numeric' },
  { sqref: 'F4', rule: { kind: 'cellIs', operator: 'equal', operand: { date: 1 } }, dxfId: 2, capabilityId: 'cellIs:equal-numeric' },
  { sqref: 'F5', rule: { kind: 'cellIs', operator: 'equal', operand: { date: 0 } }, dxfId: 2, capabilityId: 'cellIs:equal-numeric' },
  { sqref: 'G2', rule: { kind: 'cellIs', operator: 'greaterThan', operand: { date: 0 } }, dxfId: 0, capabilityId: 'cellIs:greaterThan-numeric' },
  { sqref: 'G3', rule: { kind: 'cellIs', operator: 'lessThan', operand: { date: 1 } }, dxfId: 0, capabilityId: 'cellIs:lessThan-numeric' },
  { sqref: 'H2', rule: { kind: 'timePeriod', period: 'today' }, dxfId: 0, capabilityId: 'timePeriod:today' },
  { sqref: 'H3', rule: { kind: 'timePeriod', period: 'yesterday' }, dxfId: 1, capabilityId: 'timePeriod:yesterday' },
  { sqref: 'H4', rule: { kind: 'timePeriod', period: 'tomorrow' }, dxfId: 2, capabilityId: 'timePeriod:tomorrow' },
  { sqref: 'H5', rule: { kind: 'timePeriod', period: 'today' }, dxfId: 1, capabilityId: 'timePeriod:today' },
  { sqref: 'I2:I5', rule: { kind: 'timePeriod', period: 'today' }, dxfId: 2, capabilityId: 'timePeriod:today' },
  { sqref: 'J2:J5', rule: { kind: 'cellIs', operator: 'greaterThan', operand: { number: 2 } }, dxfId: 1, capabilityId: 'cellIs:greaterThan-numeric' },
]
const ctrl = (ref: string): string => ref.replace(/\d+$/, n => String(Number(n) + CONTROL_OFFSET))
const column = (col: string, group: MidnightGroup, rule: number, values: readonly MidnightValue[]): MidnightCell[] =>
  values.map((value, i) => ({ ref: `${col}${i + 2}`, controlRef: ctrl(`${col}${i + 2}`), group, value, rule }))
const one = (ref: string, group: MidnightGroup, rule: number, value: MidnightValue): MidnightCell => ({ ref, controlRef: ctrl(ref), group, value, rule })
export const MIDNIGHT_CELLS: readonly MidnightCell[] = [
  ...column('B', 'literal-clock', 0, DATES.map(date => ({ date }))),
  ...column('C', 'literal-clock', 1, DATES.map(date => ({ date }))),
  ...column('D', 'literal-clock', 2, DATES.map(date => ({ date }))),
  ...column('E', 'literal-clock', 3, DATES.map(date => ({ date }))),
  one('F2', 'cell-literal', 4, { today: 0 }), one('F3', 'cell-literal', 5, { today: 0 }), one('F4', 'cell-literal', 6, { today: 1 }), one('F5', 'cell-literal', 7, { today: -1 }),
  one('G2', 'cell-literal', 8, { today: 0 }), one('G3', 'cell-literal', 9, { today: 0 }),
  one('H2', 'cell-clock', 10, { today: 0 }), one('H3', 'cell-clock', 11, { today: 0 }), one('H4', 'cell-clock', 12, { today: 1 }), one('H5', 'cell-clock', 13, { today: 1 }),
  ...column('I', 'clock-control', 14, [-30, 30, -365, 365].map(date => ({ date }))),
  ...column('J', 'free', 15, [1, 5, 0, 7].map(number => ({ number }))),
]
/** The recalculation sentinels (no rule): a `NOW()` and a `TODAY()` cell, read as values at every step. */
export const MIDNIGHT_SENTINELS = { now: 'K2', today: 'K3' } as const
/** The static marker fill (a plain cell fill, no rule) at the range's corners: it locates the grid in a window screenshot. */
export const MIDNIGHT_MARKER = { fillRgb: 'FF13F7A9', refs: ['A1', 'K12'] } as const
export const MIDNIGHT_RANGE = 'A1:K12'

/** The day offset a cell names in a world (its literal date, or `TODAY()+k` on the cells' day). */
export function midnightCellValue(value: MidnightValue, world: MidnightWorld): number {
  return 'date' in value ? value.date : 'today' in value ? world.cellDay + value.today : value.number
}
/** A rule's outcome on a value in a world (an independent restatement of the Stage 1 semantics: day floors and plain comparisons). */
export function midnightRuleMatches(rule: MidnightRule, v: number, world: MidnightWorld): boolean {
  switch (rule.kind) {
    case 'timePeriod': return Math.floor(v) - world.ruleDay === (rule.period === 'today' ? 0 : rule.period === 'yesterday' ? -1 : 1)
    case 'expression-today': return v === world.ruleDay
    case 'cellIs': {
      const o = 'date' in rule.operand ? rule.operand.date : rule.operand.number
      return rule.operator === 'equal' ? v === o : rule.operator === 'greaterThan' ? v > o : v < o
    }
  }
}
export function midnightToken(cell: MidnightCell, world: MidnightWorld): MidnightToken {
  const spec = MIDNIGHT_RULES[cell.rule]!
  return midnightRuleMatches(spec.rule, midnightCellValue(cell.value, world), world) ? TOKEN_OF_DXF[spec.dxfId]! : 'none'
}
export function midnightPredictions(world: MidnightWorld): Readonly<Record<string, MidnightToken>> {
  return Object.freeze(Object.fromEntries(MIDNIGHT_CELLS.map(c => [c.ref, midnightToken(c, world)])))
}
/** A cell is decisive when two readings predict different tokens at instant 2; a control when every reading and both other instants agree. */
export function midnightDecisive(cell: MidnightCell): boolean {
  const tokens = MIDNIGHT_READINGS.map(r => midnightToken(cell, MIDNIGHT_WORLDS[r]))
  return new Set(tokens).size > 1
}
/** The refs that separate two readings at instant 2. */
export function midnightSeparating(a: MidnightReading, b: MidnightReading): string[] {
  return MIDNIGHT_CELLS.filter(c => midnightToken(c, MIDNIGHT_WORLDS[a]) !== midnightToken(c, MIDNIGHT_WORLDS[b])).map(c => c.ref)
}

/* ------------------------------------------------------------------ */
/* Workbook content (generate.ts) and the case's markers (cases.ts).   */
/* ------------------------------------------------------------------ */

/** The local date `iso` (YYYY-MM-DD) as a 1900 serial. */
export function midnightSerial(iso: string): number {
  const [y, m, d] = iso.split('-').map(Number)
  return Math.round(Date.UTC(y!, m! - 1, d!) / 86_400_000) + 25569
}
const ruleXml = (spec: MidnightRuleSpec, priority: number, D: number): string => {
  const first = spec.sqref.split(':')[0]!
  const r = spec.rule
  if (r.kind === 'timePeriod') {
    const formula = r.period === 'today' ? `FLOOR(${first},1)=TODAY()` : r.period === 'yesterday' ? `FLOOR(${first},1)=TODAY()-1` : `FLOOR(${first},1)=TODAY()+1`
    return `<cfRule type="timePeriod" dxfId="${spec.dxfId}" priority="${priority}" timePeriod="${r.period}"><formula>${formula}</formula></cfRule>`
  }
  if (r.kind === 'expression-today') return `<cfRule type="expression" dxfId="${spec.dxfId}" priority="${priority}"><formula>${first}=TODAY()</formula></cfRule>`
  const o = 'date' in r.operand ? D + r.operand.date : r.operand.number
  return `<cfRule type="cellIs" dxfId="${spec.dxfId}" priority="${priority}" operator="${r.operator}"><formula>${o}</formula></cfRule>`
}
export function midnightCfXml(anchorIso: string): string {
  const D = midnightSerial(anchorIso)
  return MIDNIGHT_RULES.map((spec, i) => `<conditionalFormatting sqref="${spec.sqref}">${ruleXml(spec, i + 1, D)}</conditionalFormatting>`).join('')
}
export function midnightMarkers(anchorIso: string): CfPartMarker[] {
  const D = midnightSerial(anchorIso)
  return MIDNIGHT_RULES.map((spec, i) => {
    const first = spec.sqref.split(':')[0]!
    const r = spec.rule
    const formula = r.kind === 'timePeriod' ? (r.period === 'today' ? `FLOOR(${first},1)=TODAY()` : r.period === 'yesterday' ? `FLOOR(${first},1)=TODAY()-1` : `FLOOR(${first},1)=TODAY()+1`)
      : r.kind === 'expression-today' ? `${first}=TODAY()` : String('date' in r.operand ? D + r.operand.date : r.operand.number)
    return {
      // One question per rule (capability ids are unique across the catalogue): `clock:across-midnight:<first cell>`; the
      // case's registration id is `acrossMidnight.capabilityId` (clock:across-midnight).
      part: 'xl/worksheets/sheet1.xml', extension: 'legacy', capabilityId: `${MIDNIGHT_CAP}:${first}`, sqref: spec.sqref,
      ruleType: r.kind === 'timePeriod' ? 'timePeriod' : r.kind === 'expression-today' ? 'expression' : 'cellIs',
      ...(r.kind === 'cellIs' ? { operator: r.operator } : {}), ...(r.kind === 'timePeriod' ? { timePeriod: r.period } : {}),
      priority: i + 1, formulas: [formula], dxf: MIDNIGHT_DXFS[spec.dxfId]!,
    } as CfPartMarker
  })
}
/** Sheet1's cells for the night of `anchorIso`: headers, the designated cells, their controls, the sentinels, the markers. */
export function midnightCells(anchorIso: string): Array<{ ref: string; v: number | string | null; f?: string; s?: number; marker?: true }> {
  const D = midnightSerial(anchorIso)
  const out: Array<{ ref: string; v: number | string | null; f?: string; s?: number; marker?: true }> = []
  const headers: Record<string, string> = { B: 'Today', C: 'Yesterday', D: 'Tomorrow', E: 'ExprToday', F: 'CellEq', G: 'CellCmp', H: 'CellClock', I: 'FarDate', J: 'Free', K: 'Sentinel' }
  for (const [col, h] of Object.entries(headers)) { out.push({ ref: `${col}1`, v: h }); out.push({ ref: `${col}${1 + CONTROL_OFFSET}`, v: `${h}_Ctrl` }) }
  for (const c of MIDNIGHT_CELLS) for (const ref of [c.ref, c.controlRef]) {
    const v = c.value
    if ('today' in v) out.push({ ref, f: v.today === 0 ? 'TODAY()' : `TODAY()${v.today > 0 ? '+' : ''}${v.today}`, v: D + v.today, s: 4 })
    else if ('date' in v) out.push({ ref, v: D + v.date, s: 4 })
    else out.push({ ref, v: v.number })
  }
  out.push({ ref: MIDNIGHT_SENTINELS.now, f: 'NOW()', v: D })
  out.push({ ref: MIDNIGHT_SENTINELS.today, f: 'TODAY()', v: D, s: 4 })
  // An empty text: a cell the generator writes (it drops a null cell), carrying only the static marker fill.
  for (const ref of MIDNIGHT_MARKER.refs) out.push({ ref, v: '', marker: true })
  return out.sort((a, b) => rowOf(a.ref) - rowOf(b.ref) || colOf(a.ref) - colOf(b.ref))
}
const rowOf = (ref: string): number => Number(/\d+/.exec(ref)![0])
const colOf = (ref: string): number => [...ref.replace(/\d+/g, '')].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)

/* ------------------------------------------------------------------ */
/* The decision (docs/validation/xlsx-cf-recipes-stage1.md §4.6).       */
/* ------------------------------------------------------------------ */

/** One instant's observation: per designated ref the DisplayFormat token and the screenshot token, and the sentinel values. */
export interface MidnightInstantObservation {
  /** Local date (YYYY-MM-DD, the case zone) at the start and the end of the instant's interactions. */
  readonly localDateBefore: string
  readonly localDateAfter: string
  readonly displayFormat: Readonly<Record<string, MidnightToken>>
  /** The screenshot taken before any Excel interaction of the instant (`undefined`: unmeasurable). */
  readonly screen?: Readonly<Record<string, MidnightToken>>
  /** The screenshot taken after the readout (recorded only). */
  readonly screenAfter?: Readonly<Record<string, MidnightToken>>
  /** The sentinel `NOW()` value read first and last in the readout, and the `TODAY()` sentinel and `TODAY()` cells' values. */
  readonly nowBefore: string
  readonly nowAfter: string
  readonly cellValues: Readonly<Record<string, string>>
}
export interface MidnightCapture {
  readonly anchorLocalDate: string
  readonly timeZone: string
  readonly instants: readonly [MidnightInstantObservation, MidnightInstantObservation, MidnightInstantObservation | undefined]
}
export type MidnightDecision =
  | Readonly<{ kind: 'register'; reading: MidnightReading; reason: string }>
  | Readonly<{ kind: 'nothing'; reason: string }>

const nextIso = (iso: string): string => {
  const [y, m, d] = iso.split('-').map(Number)
  return new Date(Date.UTC(y!, m! - 1, d! + 1)).toISOString().slice(0, 10)
}
const mismatches = (observed: Readonly<Record<string, MidnightToken>> | undefined, predicted: Readonly<Record<string, MidnightToken>>, refs: readonly string[]): string[] =>
  refs.filter(r => observed?.[r] !== predicted[r])

/**
 * The pre-registered decision. A reading registers only when, all at once:
 * 1. instant 1 ran on D and instant 2 on exactly D+1 (start and end of each);
 * 2. instant 1 matches the load-day predictions at every designated cell, DisplayFormat and screen alike;
 * 3. no recalculation happened before or during instant 2's readout: the `NOW()` sentinel read first and last in instant 2
 *    equals the one of instant 1 (a spontaneous recalculation before instant 2 is allowed only when the sentinel moved
 *    BEFORE instant 2's readout started, i.e. instant 1's last value differs from instant 2's first value and instant 2's
 *    first and last values agree);
 * 4. at instant 2 exactly one reading matches every decisive cell in DisplayFormat, and every control matches;
 * 5. the instant 2 screenshot (taken before any interaction) shows the same tokens as its DisplayFormat at every
 *    designated cell (a disagreement means Excel's paint depends on a redraw: nothing registers);
 * 6. instant 3 (after Calculate Now), when captured, matches the post-recalculation predictions (the cells and rules can
 *    move: a positive control); an instant 3 that does not is an unsound case (nothing registers). A missing instant 3
 *    does not block a registration.
 */
export function decideAcrossMidnight(capture: MidnightCapture): MidnightDecision {
  const [i1, i2, i3] = capture.instants
  const D = capture.anchorLocalDate, D1 = nextIso(D)
  if (i1.localDateBefore !== D || i1.localDateAfter !== D) return { kind: 'nothing', reason: `instant 1 ran on ${i1.localDateBefore}..${i1.localDateAfter}, not ${D}` }
  if (i2.localDateBefore !== D1 || i2.localDateAfter !== D1) return { kind: 'nothing', reason: `instant 2 ran on ${i2.localDateBefore}..${i2.localDateAfter}, not ${D1}` }
  const refs = MIDNIGHT_CELLS.map(c => c.ref)
  const load = midnightPredictions(MIDNIGHT_INSTANT1)
  const bad1 = [...mismatches(i1.displayFormat, load, refs), ...mismatches(i1.screen, load, refs)]
  if (bad1.length) return { kind: 'nothing', reason: `instant 1 differs from the load-day predictions at ${[...new Set(bad1)].join(', ')}` }
  if (i2.nowBefore !== i2.nowAfter) return { kind: 'nothing', reason: 'the NOW() sentinel moved during instant 2\'s readout: the readout recalculated' }
  const spontaneous = i2.nowBefore !== i1.nowAfter
  const controls = MIDNIGHT_CELLS.filter(c => !midnightDecisive(c)).map(c => c.ref)
  const decisive = MIDNIGHT_CELLS.filter(midnightDecisive).map(c => c.ref)
  const badControls = mismatches(i2.displayFormat, midnightPredictions(MIDNIGHT_WORLDS.frozen), controls)
  if (badControls.length) return { kind: 'nothing', reason: `instant 2 controls differ at ${badControls.join(', ')}` }
  const matching = MIDNIGHT_READINGS.filter(r => mismatches(i2.displayFormat, midnightPredictions(MIDNIGHT_WORLDS[r]), decisive).length === 0)
  if (matching.length !== 1) return { kind: 'nothing', reason: matching.length === 0 ? 'no reading matches every decisive cell at instant 2' : `several readings match: ${matching.join(', ')}` }
  const reading = matching[0]!
  // A recalculation before instant 2 is only consistent with full-recalc (the cells moved by themselves).
  if (spontaneous && reading !== 'full-recalc') return { kind: 'nothing', reason: `the NOW() sentinel moved before instant 2 but ${reading} matched` }
  if (!spontaneous && reading === 'full-recalc') return { kind: 'nothing', reason: 'full-recalc matched but the NOW() sentinel did not move: no recalculation explains the cells' }
  const screenBad = i2.screen === undefined ? refs : mismatches(i2.screen, i2.displayFormat, refs)
  if (screenBad.length) return { kind: 'nothing', reason: i2.screen === undefined ? 'the instant 2 screenshot is unmeasurable' : `the instant 2 screen differs from its DisplayFormat at ${screenBad.join(', ')} (redraw-dependent)` }
  if (i3) {
    const bad3 = mismatches(i3.displayFormat, midnightPredictions(MIDNIGHT_INSTANT3), refs)
    if (bad3.length) return { kind: 'nothing', reason: `instant 3 (after Calculate Now) differs from the post-recalculation predictions at ${bad3.join(', ')}` }
  }
  return { kind: 'register', reading, reason: `${reading} matches all ${decisive.length} decisive cells and ${controls.length} controls at instant 2; screen agrees` }
}
