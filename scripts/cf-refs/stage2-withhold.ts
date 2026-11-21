/**
 * Stage 2 cells production withholds where Excel paints a fill, each with its justification, shared by the resident
 * exact-ARGB conformance (tests/xlsx-cf-oracle-conformance.test.ts) and the interactive/paged comparison gate
 * (scripts/cf-compare.ts). A withheld cell is accepted only when it carries no painted fill and is justified here:
 * - `STAGE2_WITHHELD_FINDINGS`: the typed, separately justified findings (Ruling F-2, S2-I, S2-J, Stage 5 stubs);
 * - `stage2NumSegmentShape`: the num-segment finding (capture e448d831, a87121e8), recomputed from the case's own marker
 *   and the cell's value, never a cell list.
 * A finding can never encode a painted (wrong) colour.
 */
import { isCellInSqref } from './verify'

export type Stage2Finding = { outcome: 'withheld'; reason: string }

export const STAGE2_WITHHELD_FINDINGS: Readonly<Record<string, Stage2Finding>> = {
    // Finding F2: Overlapping scales withheld under Task 15 ruling S2-I (base fill FFFFFFFF)
    'stage2-scale-scope:B2': { outcome: 'withheld', reason: 'Overlapping scale ranges B2:B4 and B2:B6 withhold fill under S2-I' },
    'stage2-scale-scope:B3': { outcome: 'withheld', reason: 'Overlapping scale ranges B2:B4 and B2:B6 withhold fill under S2-I' },
    'stage2-scale-scope:B5': { outcome: 'withheld', reason: 'Overlapping scale ranges B2:B6 and B2:B6 withhold fill under S2-I' },
    'stage2-scale-scope:B6': { outcome: 'withheld', reason: 'Overlapping scale ranges B2:B6 and B2:B6 withhold fill under S2-I' },
    'stage2-scale-scope:D2': { outcome: 'withheld', reason: 'Overlapping scale ranges D2:D4 and D2:D6 withhold fill under S2-I' },
    'stage2-scale-scope:D3': { outcome: 'withheld', reason: 'Overlapping scale ranges D2:D4 and D2:D6 withhold fill under S2-I' },

    // Finding F6 (final review I5 / Ruling F-2): the scope fixture cannot discriminate hidden-row, autofilter or
    // merged-cell participation, so scale:scope-hidden-row, scale:scope-autofilter-hidden-row and scale:scope-merged-cell
    // are unverified and these scales are WITHHELD until re-measured (scale:scope-disjoint is verified since
    // stage2-scale-scope-disjoint; these cells are withheld by their other gates).
    'stage2-scale-scope:C2': { outcome: 'withheld', reason: 'Hidden row 4 inside C2:C6: scale:scope-hidden-row is non-discriminating (hidden value 30 is not extremal) and unverified' },
    'stage2-scale-scope:C3': { outcome: 'withheld', reason: 'Hidden row 4 inside C2:C6: scale:scope-hidden-row is non-discriminating and unverified' },
    'stage2-scale-scope:C5': { outcome: 'withheld', reason: 'Hidden row 4 inside C2:C6: scale:scope-hidden-row is non-discriminating and unverified' },
    'stage2-scale-scope:C6': { outcome: 'withheld', reason: 'Hidden row 4 inside C2:C6: scale:scope-hidden-row is non-discriminating and unverified' },
    'stage2-scale-scope:D5': { outcome: 'withheld', reason: 'Hidden row 4 inside D2:D6 (the saved file has no autoFilter): scale:scope-autofilter-hidden-row is non-discriminating and unverified' },
    'stage2-scale-scope:D6': { outcome: 'withheld', reason: 'Hidden row 4 inside D2:D6: scale:scope-autofilter-hidden-row is non-discriminating and unverified' },
    'stage2-scale-scope:F2': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4 (F4 populated): hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F3': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F5': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F6': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F13': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F14': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F15': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F16': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:F17': { outcome: 'withheld', reason: 'Whole-column F:F scope contains hidden row 4: hidden-row scope shapes are unverified' },
    'stage2-scale-scope:B8': { outcome: 'withheld', reason: 'Merged B8:C9 inside the scope: scale:scope-merged-cell is non-discriminating (every merged value is 100) and unverified' },
    'stage2-scale-scope:C8': { outcome: 'withheld', reason: 'Merged B8:C9 inside the scope: scale:scope-merged-cell is non-discriminating and unverified' },
    'stage2-scale-scope:B9': { outcome: 'withheld', reason: 'Merged B8:C9 inside the scope: scale:scope-merged-cell is non-discriminating and unverified' },
    'stage2-scale-scope:C9': { outcome: 'withheld', reason: 'Merged B8:C9 inside the scope: scale:scope-merged-cell is non-discriminating and unverified' },

    // Finding F3: Overlapping scales on I2:I5 withheld under Task 15 ruling S2-I
    'stage2-scale-composition:I2': { outcome: 'withheld', reason: 'Two overlapping scales on I2:I5 withhold fill under S2-I' },
    'stage2-scale-composition:I3': { outcome: 'withheld', reason: 'Two overlapping scales on I2:I5 withhold fill under S2-I' },
    'stage2-scale-composition:I4': { outcome: 'withheld', reason: 'Two overlapping scales on I2:I5 withhold fill under S2-I' },
    'stage2-scale-composition:I5': { outcome: 'withheld', reason: 'Two overlapping scales on I2:I5 withhold fill under S2-I' },

    // Finding F4: Text cell J2 in scale range
    'stage2-scale-composition:J2': { outcome: 'withheld', reason: 'Text cell J2 lower dxf fill unmeasured in Stage 2' },

    // Finding F5: Expression rule 19 with stopIfTrue="1" withholds lower scale 20 under S2-J
    'stage2-scale-composition:L2': { outcome: 'withheld', reason: 'Stage 5 expression rule with stopIfTrue withholds lower scale' },
    'stage2-scale-composition:L3': { outcome: 'withheld', reason: 'Stage 5 expression rule with stopIfTrue withholds lower scale' },
    'stage2-scale-composition:L4': { outcome: 'withheld', reason: 'Stage 5 expression rule with stopIfTrue withholds lower scale' },
    'stage2-scale-composition:L5': { outcome: 'withheld', reason: 'Stage 5 expression rule with stopIfTrue withholds lower scale' },
}

type MarkerLike = { ruleType: string; sqref: string; scale?: { cfvo: readonly { type: string; val?: string; f?: string }[] } }

/**
 * The num-segment finding (production's `numSegmentUnmeasured`, recomputed independently from the marker): a 3-colour
 * scale cell whose value lies strictly inside a segment bounded by two literal num cfvo (after the midpoint-above-max
 * re-sort), where f*2^20 is not an integer. Excel 16.115 drew T4 and V3 of stage2-scale-scope-disjoint one unit HIGHER per
 * RGB channel than production (the descending step one unit smaller in magnitude), by a rule not yet explained.
 */
export function stage2NumSegmentShape(markers: readonly MarkerLike[], ref: string, value: unknown): boolean {
  if (typeof value !== 'number' || !Number.isFinite(value)) return false
  return markers.some(m => {
    if (m.ruleType !== 'colorScale' || m.scale?.cfvo.length !== 3 || !isCellInSqref(ref, m.sqref)) return false
    const literal = (x: { type: string; val?: string; f?: string }): number | undefined =>
      x.type === 'num' && x.f === undefined && x.val !== undefined && /^-?\d+(\.\d+)?$/.test(x.val) ? Number(x.val) : undefined
    const t = m.scale.cfvo.map(literal)
    // Positional thresholds: the midpoint above the maximum swaps the last two (as production re-sorts them).
    const idx = t[1] !== undefined && t[2] !== undefined && t[1] > t[2] ? [0, 2, 1] : [0, 1, 2]
    const [a, b, c] = idx.map(i => t[i])
    if (b === undefined) return false
    const [lo, hi] = value <= b ? [a, b] : [b, c]
    if (lo === undefined || hi === undefined || !(value > lo && value < hi)) return false
    return !Number.isInteger(((value - lo) / (hi - lo)) * 1048576)
  })
}
