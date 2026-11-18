// @vitest-environment node
/**
 * Stage 5 final review F1: a CF formula reads text under the workbook's own Unicode compatibility version, exactly as the
 * worksheet's formulas do. A parsed workbook with no verified metadata is version 1 (UTF-16 code units), so `LEN("😀")` is 2
 * on the sheet and must be 2 in a CF rule too: `LEN($A$1)=1` must not paint, `LEN($A$1)=2` paints. Every route (static,
 * resident, text, worker and its forced-paged scope) is checked, since the resident and paged policies are built separately.
 */
import { describe, expect, it } from 'vitest'
import { cfDocumentEvaluationOf, cfDocumentEvaluationOfPaged, cfFormulaSemantics, cfEvaluationContextFor } from '../src/xlsx/cf/formula'
import { RED, GREEN, cfBlock, exprRule, load, residentLeg, staticLeg, textLeg, workbookBytes, workerLeg } from './helpers/cf-views'

const EMOJI = '\u{1F600}'
const bytes = (): Uint8Array => workbookBytes({ A1: EMOJI, C1: { f: 'LEN(A1)' }, D1: 0 },
  cfBlock('B1', exprRule('LEN($A$1)=1', 1, 0)) + cfBlock('D1', exprRule('LEN($A$1)=2', 2, 1)))
// One row: the resident routes prepare the used extent (A1:D1 here), as a viewer does for this sheet.

describe('CF formulas read text under the workbook Unicode compatibility version (final review F1)', () => {
  it('the worksheet reads LEN of a non-BMP character as 2 (version 1, the parsed-workbook default)', async () => {
    const l = await load(bytes(), 'engine')
    const c1 = l.sheet.rows.flatMap(r => r.cells).find(c => c.row === 0 && c.col === 2)
    expect(c1?.value).toBe(2)
    expect(l.doc.semantics?.unicode.version).toBe(1)
  })

  it('the document policy and the CF semantics carry that version, resident and paged alike', async () => {
    const l = await load(bytes(), 'engine')
    const policy = cfDocumentEvaluationOf(l.doc)
    expect(policy.unicode).toEqual({ version: 1, source: 'workbook-default' })
    expect(cfFormulaSemantics(cfEvaluationContextFor(policy, () => 0)).unicode.version).toBe(1)
    expect(cfDocumentEvaluationOfPaged({}).unicode).toEqual({ version: 1, source: 'workbook-default' })
  })

  it('LEN($A$1)=1 does not paint and LEN($A$1)=2 paints on every route, as the worksheet LEN reads', async () => {
    for (const policy of ['engine', 'cached'] as const) {
      const l = await load(bytes(), policy)
      const refs = ['B1', 'D1']
      const worker = await workerLeg(l, refs)
      const legs = { static: await staticLeg(l, refs), resident: await residentLeg(l, refs), text: await textLeg(l, refs), textResident: await textLeg(l, refs, { resident: true }), forcedPaged: worker.scope }
      for (const [name, leg] of Object.entries(legs)) {
        expect(leg.get('B1')!.fill, `${policy}/${name}: LEN($A$1)=1 is false under version 1`).toBeUndefined()
        expect(leg.get('D1')!.fill, `${policy}/${name}: LEN($A$1)=2 is true under version 1`).toBe(GREEN)
      }
      expect(worker.states.get('B1')?.properties['fill.bg']).not.toBe(RED)
    }
  })
})
