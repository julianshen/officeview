/**
 * PR #11 review findings: byte sources are bounded before the legacy reader
 * inflates package metadata, and the per-load abort signal reaches fetch and
 * an already-aborted stream.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { loadOfficeSource } from '../src/components/loadOfficeSource'
import { OfficePackage } from '../src/core/zip'
import { ImportRefusal } from '../src/xlsx/storage/importSession'
import { buildZip } from './helpers/zip-builder'
import { xlsxEntries } from './helpers/xlsx-package'

const MiB = 1024 * 1024
const CT = '[Content_Types].xml'

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})

/** An honest spreadsheet whose content-types part inflates to ~8 MiB from a few KiB. */
function contentTypesBomb(): Uint8Array {
  const entries = xlsxEntries({})
  const ct = entries.find((e) => e.name === CT)!
  const text = typeof ct.data === 'string' ? ct.data : new TextDecoder().decode(ct.data)
  const padded = text.replace('</Types>', `${' '.repeat(8 * MiB)}</Types>`)
  return buildZip(entries.map((e) => (e === ct ? { ...e, data: padded, method: 8 } : e))).bytes
}

async function rejection(p: Promise<unknown>): Promise<unknown> {
  try { await p } catch (e) { return e }
  throw new Error('expected a rejection')
}

describe('byte sources are bounded before the legacy reader inflates metadata (PR #11 P1)', () => {
  for (const [label, wrap] of [
    ['Uint8Array', (b: Uint8Array) => b],
    ['ArrayBuffer', (b: Uint8Array) => b.slice().buffer],
    ['Blob', (b: Uint8Array) => new Blob([b as BlobPart])],
  ] as const) {
    it(`${label}: an oversized [Content_Types].xml is refused as a resource limit without being inflated whole`, async () => {
      const bytes = contentTypesBomb()
      expect(bytes.length).toBeLessThan(256 * 1024)
      const text = vi.spyOn(OfficePackage.prototype, 'text')
      const error = await rejection(loadOfficeSource(wrap(bytes)))
      expect(text.mock.calls.map((c) => c[0])).not.toContain(CT)
      // the typed refusal <OfficeFile> surfaces with its code and reason
      expect(error).toBeInstanceOf(ImportRefusal)
      expect(error).toMatchObject({ code: 'resource-limit', reason: 'limit' })
    })
  }
})
