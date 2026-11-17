/**
 * Task 39 fix round 1 (F1, ruling): a Stage 5 cell may be withheld only where an independently recorded justification
 * exists, keyed per cell from evidence, never from production's own unavailable state:
 * - a capability absent from the published set (or a reading production does not implement) that a covering formula rule
 *   needs: the rule's own marker capability, and every parse and evaluation capability its formula text requires
 *   (compiled here from the manifest's marker, not taken from production's parse);
 * - a formula the CF compiler refuses outright (an unsupported shape: the refusal is the justification);
 * - an enumerated refusal (`STAGE5_ENUMERATED_REFUSALS`), each with its ruling.
 * An empty list means the cell has no justification: withholding it is a failure.
 */
import { compileCfFormula, formulaRequiredCapabilities, isCfFormulaRefusal, type CfFormulaUse } from '../../src/xlsx/cf/formula'
import { CF_FORMULA_PRODUCTION_READINGS, cfFormulaReading } from '../../src/xlsx/cf/formulaReadings'
import { createCfMemory } from '../../src/xlsx/cf/limits'
import { VERIFIED_CAPABILITIES } from '../../src/xlsx/cf/evidence'
import { isCellInSqref } from './verify'

type Cfvo = Readonly<{ type: string; val?: string; f?: string }>
type Marker = Readonly<{ sqref: string; capabilityId: string; ruleType: string; formulas?: readonly string[]; formula?: Readonly<{ kind: string; anchor: string; dialect: 'base' | 'x14' }>
  scale?: Readonly<{ cfvo?: readonly Cfvo[] }>; bar?: Readonly<{ cfvo?: readonly Cfvo[] }>; icon?: Readonly<{ cfvo?: readonly Cfvo[] }> }>
/** The family capability a `type="formula"` bound needs (service.ts VISUAL_FORMULA_CAPABILITY, Ruling S5-H). */
const FORMULA_BOUND_FAMILY: Readonly<Record<string, string>> = { colorScale: 'scale:thresholds-formula', dataBar: 'bar:bounds-formula', iconSet: 'icon:thresholds-formula' }
const isLiteralNumber = (v: string | undefined): boolean => v === undefined || /^-?\d+(\.\d+)?$/.test(v.trim())
type CaseLike = Readonly<{ id: string; markers: readonly Marker[]; formulaConfigurations?: ReadonlyArray<Readonly<{ id: string; refs: readonly string[]; record?: string }>> }>

/**
 * Per-cell refusals a published configuration still takes, each with the ruling that records it (`caseId:ref`). Each is
 * checked against the workbook's own data in tests/xlsx-cf-stage5.test.ts.
 */
export const STAGE5_ENUMERATED_REFUSALS: Readonly<Record<string, string>> = Object.fromEntries(
  Array.from({ length: 18 }, (_, i) => [`stage5-formula-functions:H${i + 2}`,
    'S5-HOIST policy (Task 34; recorded in stage5-formula-hoist hoist-refused-boolean): COUNTIF($A:$A,A2)>1 counts over a range holding a boolean (A6 = TRUE), which the COUNTIF frequency hoist refuses for the whole rule']))

const a1 = (ref: string): { row: number; col: number } => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!
  return { row: Number(m[2]) - 1, col: [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 }
}

/** Whether production holds `capability` with a reading it implements (a registered formula id) or as published evidence. */
function held(capability: string, published: (id: string) => boolean): boolean {
  return Object.prototype.hasOwnProperty.call(CF_FORMULA_PRODUCTION_READINGS, capability) ? cfFormulaReading(capability, published) !== undefined : published(capability)
}

/** The recorded reasons `ref` of `c` may be withheld (empty: none, so withholding it is a failure). */
export function stage5WithholdJustification(c: CaseLike, ref: string, published: (id: string) => boolean = id => VERIFIED_CAPABILITIES.has(id)): string[] {
  const reasons = new Set<string>()
  for (const m of c.markers) {
    if (m.formula === undefined || !isCellInSqref(ref, m.sqref)) continue
    if (!held(m.capabilityId, published)) reasons.add(`${m.capabilityId} not held`)
    const ranges = m.sqref.trim().split(/\s+/)
    const use = m.formula.kind as CfFormulaUse
    const texts = new Set(m.formulas ?? [])
    // A visual's formula bounds (S5-H): each non-literal cfvo value is a cfvo formula; a `formula` bound needs its family
    // capability, a colour scale's reference `num` the Stage 2 reference reading (service.ts visualFormulaRefusal).
    if (use === 'cfvo') {
      for (const b of m.scale?.cfvo ?? m.bar?.cfvo ?? m.icon?.cfvo ?? []) {
        const text = b.f ?? b.val
        if (isLiteralNumber(text)) continue
        texts.add(text!)
        if (b.type === 'formula' && FORMULA_BOUND_FAMILY[m.ruleType] && !held(FORMULA_BOUND_FAMILY[m.ruleType]!, published)) reasons.add(`${FORMULA_BOUND_FAMILY[m.ruleType]} not held (a formula bound)`)
        if (m.ruleType === 'colorScale' && b.type === 'num' && !published('scale:thresholds-num-val-reference')) reasons.add('scale:thresholds-num-val-reference not held')
      }
    }
    for (const text of texts) {
      const out = compileCfFormula(text, use === 'cfvo' || ranges.length === 1 ? a1(m.formula.anchor) : undefined, createCfMemory())
      if (isCfFormulaRefusal(out)) { reasons.add(`${m.capabilityId}: ${text} refused at compile (${out.refused}: ${out.reason})`); continue }
      try {
        const caps = formulaRequiredCapabilities(out.value, use, { dialect: m.formula.dialect, rangeCount: ranges.length, sourceType: m.ruleType as never })
        const needed = [...caps.parse, ...caps.evaluation]
        if (use === 'cfvo') needed.push(out.value.ast.type === 'cell' ? 'formula:cfvo-reference' : 'formula:cfvo-expression')
        for (const cap of needed) if (!held(cap, published)) reasons.add(`${cap} not held (needed by ${text})`)
      } finally { out.release() }
    }
  }
  // A record configuration (fixed before capture: "recorded, never evidence") names its refusal itself.
  for (const cfg of c.formulaConfigurations ?? []) if (cfg.record && cfg.refs.includes(ref)) reasons.add(`${cfg.id}: ${cfg.record}`)
  const enumerated = STAGE5_ENUMERATED_REFUSALS[`${c.id}:${ref}`]
  if (enumerated) reasons.add(enumerated)
  return [...reasons].sort()
}
