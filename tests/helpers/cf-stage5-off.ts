/**
 * Task 39: Stage 5 is published, so production evidence now holds formula capabilities. Tests whose subject is the
 * pre-publication (or never-evidenced) path run it under exactly production's evidence minus the published Stage 5 set.
 */
import { VERIFIED_CAPABILITIES, withExactCfEvidence } from '../../src/xlsx/cf/evidence'
import * as verified from '../../src/xlsx/cf/verifiedCapabilities'

const STAGE5: ReadonlySet<string> = (Reflect.get(verified, 'VERIFIED_STAGE5_CAPABILITIES') as ReadonlySet<string> | undefined) ?? new Set()
/** Production's evidence without any Stage 5 id (the Stage 1-4 sets). */
export const PRE_STAGE5_EVIDENCE: readonly string[] = [...VERIFIED_CAPABILITIES].filter(id => !STAGE5.has(id))
/** Runs `fn` as production did before `--publish-stage5`: Stage 1-4 evidence only. */
export function withoutStage5Evidence<T>(fn: () => T): T {
  return withExactCfEvidence(PRE_STAGE5_EVIDENCE, fn)
}
