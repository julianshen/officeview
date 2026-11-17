/**
 * Task 39 fix rounds 1-2 (rulings F1, R1, R2): a Stage 5 cell may leave a write slot unresolved only where an
 * independently recorded justification exists for THAT slot of THAT rule, keyed per cell from evidence, never from
 * production's own unavailable state. A justification names the rule (its marker: capability and priority) and the
 * slots it may leave unresolved: the rule's own slots, plus, for a stopIfTrue rule, the slots of the lower rules covering
 * the cell (the stop's downstream effect). Its reason is one of:
 * - a capability absent from the published set (or a reading production does not implement) that the rule needs: its
 *   marker capability, every parse and evaluation capability its formula text requires (compiled here from the
 *   manifest's marker), a formula bound's family, or a literal Stage 3-4 gate a twin meets;
 * - a record configuration fixed before capture ("recorded, never evidence") covering the cell;
 * - an enumerated refusal (`STAGE5_ENUMERATED_REFUSALS`), keyed by case, cell and exact formula text.
 * A compile refusal is NOT a justification by itself (R1): the compiler cannot vouch for its own refusal. Only an
 * enumerated expectation (`STAGE5_EXPECTED_COMPILE_REFUSALS`: case, formula text, refusal reason) can.
 */
import { compileCfFormula, formulaRequiredCapabilities, isCfFormulaRefusal, type CfFormulaUse } from '../../src/xlsx/cf/formula'
import { CF_FORMULA_PRODUCTION_READINGS, cfFormulaReading } from '../../src/xlsx/cf/formulaReadings'
import { createCfMemory } from '../../src/xlsx/cf/limits'
import { VERIFIED_CAPABILITIES } from '../../src/xlsx/cf/evidence'
import { isCellInSqref } from './verify'
import type { CfOracleCase } from './cases'

type Cfvo = Readonly<{ type: string; val?: string; f?: string }>
type Marker = Readonly<{ sqref: string; capabilityId: string; ruleType: string; priority?: number; stopIfTrue?: boolean; formulas?: readonly string[]
  formula?: Readonly<{ kind: string; anchor: string; dialect: 'base' | 'x14' }>
  dxf?: Readonly<{ fillRgb?: string; fontRgb?: string; bold?: boolean; italic?: boolean; underline?: unknown; strikethrough?: boolean; numFmt?: unknown; borderTop?: unknown; borderBottom?: unknown }>
  scale?: Readonly<{ cfvo?: readonly Cfvo[] }>; bar?: Readonly<{ cfvo?: readonly Cfvo[] }>; icon?: Readonly<{ cfvo?: readonly Cfvo[] }> }>
type CaseLike = Readonly<{ id: string; markers: readonly Marker[]; formulaConfigurations?: ReadonlyArray<Readonly<{ id: string; capabilityId?: string; refs: readonly string[]; record?: string }>> }>

/** One rule-and-slot-specific justification. */
export interface Stage5Justification { rule: string; slots: readonly string[]; reason: string }

/** The family capability a `type="formula"` bound needs (service.ts VISUAL_FORMULA_CAPABILITY, Ruling S5-H). */
const FORMULA_BOUND_FAMILY: Readonly<Record<string, string>> = { colorScale: 'scale:thresholds-formula', dataBar: 'bar:bounds-formula', iconSet: 'icon:thresholds-formula' }
const isLiteralNumber = (v: string | undefined): boolean => v === undefined || /^-?\d+(\.\d+)?$/.test(v.trim())

/**
 * Enumerated refusals a published configuration still takes: case, cell range, exact formula text and the ruling.
 * Checked against the workbook's own data in tests/xlsx-cf-stage5.test.ts.
 */
export const STAGE5_ENUMERATED_REFUSALS: ReadonlyArray<Readonly<{ caseId: string; refs: readonly string[]; formula: string; reason: string }>> = [{
  caseId: 'stage5-formula-functions', refs: Array.from({ length: 18 }, (_, i) => `H${i + 2}`), formula: 'COUNTIF($A:$A,A2)>1',
  reason: 'S5-HOIST policy (Task 34; recorded in stage5-formula-hoist hoist-refused-boolean): COUNTIF($A:$A,A2)>1 counts over a range holding a boolean (A6 = TRUE), which the COUNTIF frequency hoist refuses for the whole rule',
}]

/**
 * Ruling R1: compile refusals that are expected, each recorded independently of the compiler (case, exact formula text,
 * refusal code, and the ruling that predicts it). Any other compile refusal of a covering rule leaves its slots
 * unjustified. tests/xlsx-cf-stage5.test.ts checks each entry still refuses exactly so.
 */
export const STAGE5_EXPECTED_COMPILE_REFUSALS: ReadonlyArray<Readonly<{ caseId: string; formula: string; refused: 'cf-unsupported-rule' | 'cf-resource-limit'; reason: string }>> = [
  { caseId: 'stage5-formula-truthiness', formula: 'NA()', refused: 'cf-unsupported-rule', reason: 'OD-13 / Ruling S5-B: the closed tier-1 function allow-list holds no NA(), so the error-truth rule built on it is unresolved' },
  { caseId: 'stage5-composition', formula: 'INDIRECT("A"&ROW())>0', refused: 'cf-unsupported-rule', reason: 'Ruling S5-O / S5-B: the composition case\'s INDIRECT expression is the one production cannot evaluate (the case\'s question; INDIRECT is a recorded unresolved shape, stage5-unresolved-shapes-indirect)' },
]

const a1 = (ref: string): { row: number; col: number } => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!
  return { row: Number(m[2]) - 1, col: [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 }
}

/** Whether production holds `capability` with a reading it implements (a registered formula id) or as published evidence. */
function held(capability: string, published: (id: string) => boolean): boolean {
  return Object.prototype.hasOwnProperty.call(CF_FORMULA_PRODUCTION_READINGS, capability) ? cfFormulaReading(capability, published) !== undefined : published(capability)
}

/**
 * Fix round 3 (S2): the exact write paths a rule's marker writes, as compose.ts keeps them (parse.ts rule.writes): a
 * visual's own slot (a colour scale writes `colorScale` and `fill`), or each DXF property separately (`fill`,
 * `font.strike`, `font.bold`, ...). Never a whole-font grant.
 */
export function stage5MarkerSlots(m: Marker): string[] {
  if (m.ruleType === 'dataBar') return ['dataBar']
  if (m.ruleType === 'iconSet') return ['iconSet']
  if (m.ruleType === 'colorScale') return ['colorScale', 'fill']
  const d = m.dxf ?? {}
  const out: string[] = []
  if (d.fillRgb !== undefined) out.push('fill')
  if (d.fontRgb !== undefined) out.push('font.color')
  if (d.bold !== undefined) out.push('font.bold')
  if (d.italic !== undefined) out.push('font.italic')
  if (d.underline !== undefined) out.push('font.underline')
  if (d.strikethrough !== undefined) out.push('font.strike')
  if (d.numFmt !== undefined) out.push('numFmt')
  if (d.borderTop !== undefined) out.push('border.top')
  if (d.borderBottom !== undefined) out.push('border.bottom')
  return out
}

const label = (m: Marker): string => `${m.capabilityId}/${m.ruleType}${m.priority !== undefined ? `@p${m.priority}` : ''}`

/** A composition gate a marker's capability names (compose.ts): a stop gate of a lower kind, or the expression-vs-scale
 * fill arbitration. Undefined: the capability is the rule's own reading (its evaluation). */
function compositionGate(capabilityId: string): { gate: 'stop'; kind: 'dxf' | 'scale' | 'bar' | 'icon' } | { gate: 'fill-arbitration' } | undefined {
  const stop = /^formula:composition-(?:stop|indirect-ahead)-(dxf|scale|bar|icon)$/.exec(capabilityId)
  if (stop) return { gate: 'stop', kind: stop[1] as 'dxf' | 'scale' | 'bar' | 'icon' }
  if (capabilityId === 'formula:composition-expression-vs-scale') return { gate: 'fill-arbitration' }
  return undefined
}

/** The reasons a formula marker's RULE is unresolved (its own evaluation; empty: production holds everything it needs).
 * The marker's composition gate, if any, is not one of them (see `stage5Justifications`). */
function formulaMarkerReasons(m: Marker, caseId: string, published: (id: string) => boolean): string[] {
  const reasons: string[] = []
  if (!compositionGate(m.capabilityId) && !held(m.capabilityId, published)) reasons.push(`${m.capabilityId} not held`)
  const ranges = m.sqref.trim().split(/\s+/)
  const use = m.formula!.kind as CfFormulaUse
  const texts = new Set(m.formulas ?? [])
  // A visual's formula bounds (S5-H): each non-literal cfvo value is a cfvo formula; a `formula` bound needs its family
  // capability, a colour scale's reference `num` the Stage 2 reference reading (service.ts visualFormulaRefusal).
  if (use === 'cfvo') {
    for (const b of m.scale?.cfvo ?? m.bar?.cfvo ?? m.icon?.cfvo ?? []) {
      const text = b.f ?? b.val
      if (isLiteralNumber(text)) continue
      texts.add(text!)
      if (b.type === 'formula' && FORMULA_BOUND_FAMILY[m.ruleType] && !held(FORMULA_BOUND_FAMILY[m.ruleType]!, published)) reasons.push(`${FORMULA_BOUND_FAMILY[m.ruleType]} not held (a formula bound)`)
      if (m.ruleType === 'colorScale' && b.type === 'num' && !published('scale:thresholds-num-val-reference')) reasons.push('scale:thresholds-num-val-reference not held')
    }
  }
  for (const text of texts) {
    const out = compileCfFormula(text, use === 'cfvo' || ranges.length === 1 ? a1(m.formula!.anchor) : undefined, createCfMemory())
    if (isCfFormulaRefusal(out)) {
      // R1: only an independently recorded expectation of exactly this refusal justifies it.
      const expected = STAGE5_EXPECTED_COMPILE_REFUSALS.find(e => e.caseId === caseId && e.formula === text && e.refused === out.refused)
      if (expected) reasons.push(`${text}: expected compile refusal (${expected.reason})`)
      continue
    }
    try {
      const caps = formulaRequiredCapabilities(out.value, use, { dialect: m.formula!.dialect, rangeCount: ranges.length, sourceType: m.ruleType as never })
      const needed = [...caps.parse, ...caps.evaluation]
      if (use === 'cfvo') needed.push(out.value.ast.type === 'cell' ? 'formula:cfvo-reference' : 'formula:cfvo-expression')
      for (const cap of needed) if (!held(cap, published)) reasons.push(`${cap} not held (needed by ${text})`)
    } finally { out.release() }
  }
  return reasons
}

/**
 * The rule- and slot-specific justifications for `ref` of `c` (empty: no write of the cell may be left unresolved). Each
 * is scoped by its real effect in compose.ts (fix round 3, S1):
 * - the rule itself unresolved (its evaluation needs an unheld capability, a record configuration, an enumerated or an
 *   expected compile refusal): a stopIfTrue rule halts composition, `ALL`; otherwise its own writes, plus the lower icon or
 *   bar slot an unknown fill or bar leaves unknown (icon:composition-dxf-fill, bar:composition-dxf-fill,
 *   icon:composition-with-bar not held);
 * - a stop gate the rule's marker names (formula:composition-stop-<kind>, -indirect-ahead-<kind>) not held: applied AFTER
 *   the stopping rule paints, it taints only the writes of the lower rules of that kind;
 * - the expression-vs-scale arbitration not held: `fill` and `colorScale`;
 * - a literal Stage 3-4 gate a twin meets (bar:bounds-num, icon:composition-dxf-fill): that visual's own slot.
 */
export function stage5Justifications(c: CaseLike | CfOracleCase, ref: string, published: (id: string) => boolean = id => VERIFIED_CAPABILITIES.has(id)): Stage5Justification[] {
  const cs = c as unknown as CaseLike
  const covering = cs.markers.filter(m => isCellInSqref(ref, m.sqref))
  const lower = (m: Marker): Marker[] => m.priority === undefined ? [] : covering.filter(o => o.priority !== undefined && o.priority > m.priority!)
  const kindOf = (m: Marker): 'dxf' | 'scale' | 'bar' | 'icon' => (m.ruleType === 'colorScale' ? 'scale' : m.ruleType === 'dataBar' ? 'bar' : m.ruleType === 'iconSet' ? 'icon' : 'dxf')
  /** The writes an unresolved rule leaves unknown (compose.ts: its own writes, or ALL when it halts). */
  const unresolvedScope = (m: Marker): string[] => {
    if (m.stopIfTrue) return ['ALL']
    const own = stage5MarkerSlots(m)
    const out = new Set(own)
    for (const o of lower(m)) {
      if (o.ruleType === 'iconSet' && ((own.includes('fill') && !published('icon:composition-dxf-fill')) || (own.includes('dataBar') && !published('icon:composition-with-bar')))) out.add('iconSet')
      if (o.ruleType === 'dataBar' && own.includes('fill') && !published('bar:composition-dxf-fill')) out.add('dataBar')
    }
    return [...out]
  }
  const out: Stage5Justification[] = []
  for (const m of covering) {
    if (m.formula !== undefined) {
      for (const reason of formulaMarkerReasons(m, cs.id, published)) out.push({ rule: label(m), slots: unresolvedScope(m), reason })
      const gate = compositionGate(m.capabilityId)
      if (gate && !held(m.capabilityId, published)) {
        const slots = gate.gate === 'fill-arbitration' ? ['fill', 'colorScale'] : [...new Set(lower(m).filter(o => kindOf(o) === gate.kind).flatMap(stage5MarkerSlots))]
        if (slots.length > 0) out.push({ rule: label(m), slots, reason: `${m.capabilityId} not held (a composition gate)` })
      }
    } else {
      if (m.ruleType === 'dataBar' && (m.bar?.cfvo ?? []).some(b => b.type === 'num') && !published('bar:bounds-num')) out.push({ rule: label(m), slots: unresolvedScope(m), reason: 'bar:bounds-num not held (a literal bar with num bounds)' })
      if (m.ruleType === 'iconSet' && covering.some(o => o !== m && o.dxf?.fillRgb !== undefined) && !published('icon:composition-dxf-fill')) out.push({ rule: label(m), slots: ['iconSet'], reason: 'icon:composition-dxf-fill not held (an icon over a DXF fill)' })
    }
  }
  // A record configuration (fixed before capture: "recorded, never evidence") names a refusal of the rules of its capability.
  for (const cfg of cs.formulaConfigurations ?? []) {
    if (!cfg.record || !cfg.refs.includes(ref)) continue
    for (const m of covering.filter(x => x.formula !== undefined && (cfg.capabilityId === undefined || x.capabilityId === cfg.capabilityId))) out.push({ rule: label(m), slots: unresolvedScope(m), reason: `${cfg.id}: ${cfg.record}` })
  }
  for (const e of STAGE5_ENUMERATED_REFUSALS) {
    if (e.caseId !== cs.id || !e.refs.includes(ref)) continue
    for (const m of covering.filter(x => x.formulas?.includes(e.formula))) out.push({ rule: label(m), slots: unresolvedScope(m), reason: e.reason })
  }
  return out
}

/** Backward-compatible flat list of reasons (any slot); prefer `stage5Justifications`. */
export function stage5WithholdJustification(c: CaseLike | CfOracleCase, ref: string, published?: (id: string) => boolean): string[] {
  return [...new Set(stage5Justifications(c, ref, published).map(j => `${j.rule}: ${j.reason}`))].sort()
}

/** Whether the justifications excuse the unresolved write `path` (fix round 3: exact paths; `ALL` only by an `ALL` grant). */
export function stage5SlotJustified(justifications: readonly Stage5Justification[], path: string): boolean {
  return justifications.some(j => j.slots.some(s => s === 'ALL' || s === path || (path !== 'ALL' && path.startsWith(`${s}.`))))
}

/** compose.ts `conflicts`: whether an unresolved write `a` covers the component `b`. */
export const stage5Conflicts = (a: string, b: string): boolean => a === 'ALL' || b === 'ALL' || a === b || a.startsWith(`${b}.`) || b.startsWith(`${a}.`)

/** What production shows at a cell: the resolved fill/strike, the unresolved write slots, and the visual it draws. */
export interface Stage5Production { fill: string; strike: boolean; unresolved: string[]; visual: string }
/** What Excel showed: its DisplayFormat fill/strike, and the bar/icon measured from the PNG where measurable. */
export interface Stage5Excel { fill: string; strike: boolean; visual?: string }

const argbClose = (a: string, b: string): boolean => a === b || (a.length === 8 && b.length === 8 && [2, 4, 6].every(k => Math.abs(parseInt(a.slice(k, k + 2), 16) - parseInt(b.slice(k, k + 2), 16)) <= 8))
/**
 * Rulings F3/R2/S1/S2: one cell's verdict. Every unresolved write path must be excused by a justification naming it (an
 * `ALL` only by an `ALL` grant, which compose.ts produces only for a halting rule). Every component no unresolved write
 * covers is compared with Excel, property by property: the fill (within +-8, the Stage 2 scale tolerance), the
 * strikethrough, and a bar or icon Excel was measured to draw (or not).
 */
export function stage5CellAgreement(production: Stage5Production, excel: Stage5Excel | undefined, justifications: readonly Stage5Justification[]): { agrees: boolean; disagreement?: string } {
  for (const path of production.unresolved) if (!stage5SlotJustified(justifications, path)) return { agrees: false, disagreement: `${path} unresolved without a recorded justification for it` }
  const open = (component: string): boolean => !production.unresolved.some(u => stage5Conflicts(u, component))
  if (!excel) return { agrees: true }
  if (open('fill') && open('colorScale') && !argbClose(production.fill, excel.fill)) return { agrees: false, disagreement: `fill ${production.fill} vs Excel ${excel.fill}` }
  if (open('font.strike') && production.strike !== excel.strike) return { agrees: false, disagreement: `strikethrough ${production.strike} vs Excel ${excel.strike}` }
  if (excel.visual !== undefined) {
    const slot = excel.visual === 'bar' || production.visual === 'bar' ? 'dataBar' : 'iconSet'
    if (production.visual === 'withheld') {
      if (excel.visual !== 'none' && !stage5SlotJustified(justifications, slot)) return { agrees: false, disagreement: `visual withheld without a recorded justification (Excel ${excel.visual})` }
    } else if (open(slot) && production.visual !== excel.visual) return { agrees: false, disagreement: `visual ${production.visual} vs Excel ${excel.visual}` }
  }
  return { agrees: true }
}
