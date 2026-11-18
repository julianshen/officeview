/**
 * Debug-only (src/debug is excluded from the build, tsconfig.build.json): the e2e page's Stage 5 evidence seam. No Stage 5
 * capability is verified yet, so production withholds every formula cell; `?cfev=1` makes the e2e page, and the viewer worker it
 * names (`cfev` Worker-name segment, since Playwright init scripts never reach a dedicated worker), treat every Stage 5 reading
 * production implements as held, so the browser can compare real paint across legs. The tests that use it say so; nothing here
 * is reachable from the shipped entries.
 */
import { injectCfEvidence, injectCfFormulaSelections } from '../xlsx/cf/evidence'
import { CF_FORMULA_PRODUCTION_READINGS } from '../xlsx/cf/formulaReadings'

export function installCfEvidence(): void {
  injectCfEvidence(Object.keys(CF_FORMULA_PRODUCTION_READINGS))
  const selections: Record<string, string> = {}
  for (const [id, reading] of Object.entries(CF_FORMULA_PRODUCTION_READINGS)) if (typeof reading !== 'string') selections[id] = reading[0]!
  injectCfFormulaSelections(selections)
}
