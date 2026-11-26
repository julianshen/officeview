/**
 * The percent-order follow-up (docs/validation/xlsx-cf.md §14.z.24): the records of the captures before stage2-scale-percent-order
 * and stage5-icon-percent-order (ecf00906) are kept as of their capture. Production then computed every percent / percentile
 * threshold lerp-first and had no scale:formula-bounds-v2 reading; this selects that state through the registration seams
 * (the evidence is untouched). A test file that pins such a record calls it in `beforeAll` and `asRegistered` in `afterAll`.
 */
import { CF_FORMULA_BOUNDS_V2_CAP, CF_PERCENT_ORDER_CAPS, injectCfFormulaBoundsV2Readings, injectCfPercentOrderReadings, type CfFormulaBoundsV2Reading } from '../../src/xlsx/cf/colorScale'

export function asBeforePercentOrder(): void {
  injectCfPercentOrderReadings(Object.fromEntries(CF_PERCENT_ORDER_CAPS.map(id => [id, 'lerp-first' as const])))
  injectCfFormulaBoundsV2Readings({ [CF_FORMULA_BOUNDS_V2_CAP]: undefined as unknown as CfFormulaBoundsV2Reading })
}

/** Back to the registered tables (clears both seams). */
export function asRegistered(): void {
  injectCfPercentOrderReadings(null)
  injectCfFormulaBoundsV2Readings(null)
}
