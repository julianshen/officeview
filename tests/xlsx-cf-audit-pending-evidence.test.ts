// @vitest-environment node
/**
 * The Stage 5 publication gate (`audit.ts --write --publish-stage5`) judges the records under the evidence about to be
 * published. Review P2 (percent-order full review): the pending set is parsed from the rendered source, so a parse that
 * finds nothing, or a set missing committed ids, must refuse instead of judging under an empty or partial evidence set
 * (unresolved visuals are withheld, not compared, so an empty set could pass the gate trivially).
 */
import { describe, expect, it } from 'vitest'
import { pendingPublishedCapabilities } from '../scripts/cf-refs/audit'

const source = (ids: string[]): string =>
  `export const VERIFIED_STAGE5_CAPABILITIES: ReadonlySet<string> = new Set([\n  'x:a',\n])\n` +
  `export const VERIFIED_CAPABILITIES: ReadonlySet<string> = new Set([\n${ids.map(i => `  '${i}',`).join('\n')}\n])\n`

describe('audit: the pending Stage 5 evidence fails closed', () => {
  it('returns the published ids when they include every committed id', () => {
    expect(pendingPublishedCapabilities(source(['a:1', 'b:2', 'c:3']), new Set(['a:1', 'b:2']))).toEqual(['a:1', 'b:2', 'c:3'])
  })
  it('refuses a source without the VERIFIED_CAPABILITIES declaration', () => {
    expect(() => pendingPublishedCapabilities('export const OTHER = 1\n', new Set(['a:1']))).toThrow(/declaration not found/)
  })
  it('refuses an empty set', () => {
    expect(() => pendingPublishedCapabilities(source([]), new Set())).toThrow(/lists no id/)
  })
  it('refuses a set that drops a committed id', () => {
    expect(() => pendingPublishedCapabilities(source(['a:1']), new Set(['a:1', 'b:2']))).toThrow(/drops committed ids: b:2/)
  })
  it('accepts the committed source itself', async () => {
    const { readFileSync } = await import('node:fs')
    const { resolve } = await import('node:path')
    const { VERIFIED_CAPABILITIES } = await import('../src/xlsx/cf/verifiedCapabilities')
    const committed = readFileSync(resolve(__dirname, '../src/xlsx/cf/verifiedCapabilities.ts'), 'utf8')
    expect(new Set(pendingPublishedCapabilities(committed, VERIFIED_CAPABILITIES))).toEqual(VERIFIED_CAPABILITIES)
  })
})
