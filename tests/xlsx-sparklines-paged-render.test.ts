// @vitest-environment node
/**
 * Paged equals resident, as pixels (plan Task 13 Step 1). For every sparkline reference fixture and the
 * XlsxWriter-written workbook, the sheet is rendered once with the resident `resolvedSparklines` and once with the
 * `ResolvedSparkline`s the paged path produces (the worker-side `readSparklines` snapshot, rebuilt on the main thread
 * by `rebuildResolvedSparklines`), through the same static renderer at two zoom levels, and the two bitmaps must be
 * byte-identical. A render with no sparklines at all is the control that proves the comparison is not blank == blank.
 *
 * tests/xlsx-sparklines-paged.test.ts proves the same equality as data; this proves it where it matters, in the pixels.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { SPARKLINE_CASES } from '../scripts/sparkline-refs/cases'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, renderSheet } from '../src/xlsx/render'
import { preloadSparklines } from '../src/xlsx/sparklineLoader'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import type { ResolvedSparklines } from '../src/xlsx/sparklines/resolve'
import type { XlsxSheet } from '../src/xlsx/types'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { rebuildResolvedSparklines } from '../src/worker/viewerRemoteSparklines'

const fixtureBytes = (name: string) => new Uint8Array(readFileSync(resolve(process.cwd(), `tests/fixtures/xlsx-sparklines/${name}.xlsx`)))
/** Fixtures whose sparklines legitimately draw nothing at any zoom: a tiny host (the insets swallow the plot) and a dateAxis group with no date range. */
const BLANK = new Set(['line-tiny-cell', 'col-tiny', 'date-missing-range'])
const SCALES = [0.75, 1.5]

/** Every sheet's worker-resolved sparklines, rebuilt like the viewer's main thread does (the helper of the paged parity test, copied: that file is not imported). */
async function resolvePaged(bytes: Uint8Array, docId: string): Promise<Map<string, ResolvedSparklines | undefined>> {
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer as ArrayBuffer]) }, { docId, profile: 'D' })
  try {
    expect((await session.done).kind).toBe('complete')
    const source = createPagedWorkbookSource(session)
    const out = new Map<string, ResolvedSparklines | undefined>()
    for (const sheet of source.sheets) {
      const store = source.store(sheet.sheetId)
      const result = await store.readSparklines!({ signal: new AbortController().signal, stamp: store.stamp }, 64 * 1024 * 1024, { set: () => true })
      out.set(sheet.name, rebuildResolvedSparklines(result))
    }
    await source.release()
    return out
  } finally { await session.release() }
}

function pixels(sheet: XlsxSheet, scale: number): Buffer {
  const m = computeMetrics(sheet)
  const canvas = createCanvas(Math.max(1, Math.ceil(m.widthPx * scale)), Math.max(1, Math.ceil(m.heightPx * scale))), ctx = canvas.getContext('2d')
  ctx.scale(scale, scale)
  renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, m)
  return Buffer.from(ctx.getImageData(0, 0, canvas.width, canvas.height).data.buffer)
}

const differing = (a: Buffer, b: Buffer): number => {
  if (a.length !== b.length) return Number.POSITIVE_INFINITY
  let n = 0
  for (let i = 0; i < a.length; i += 4) if (a[i] !== b[i] || a[i + 1] !== b[i + 1] || a[i + 2] !== b[i + 2] || a[i + 3] !== b[i + 3]) n++
  return n
}

const fixtures = [...SPARKLINE_CASES.map(c => c.id), 'xlsxwriter-written']

describe('paged-path sparklines render pixel-identically to the resident path', () => {
  it('covers every reference fixture and the XlsxWriter workbook', () => {
    expect(fixtures).toHaveLength(SPARKLINE_CASES.length + 1)
    expect(new Set(fixtures).size).toBe(fixtures.length)
    expect(fixtures.length).toBeGreaterThanOrEqual(49)
  })

  it('every fixture, at every zoom, resident == paged, and the sparklines are really painted', async () => {
    await preloadSparklines()
    const report: string[] = []
    let inkedFixtures = 0
    for (const fixture of fixtures) {
      const bytes = fixtureBytes(fixture)
      const resident = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
      const paged = await resolvePaged(bytes, `spark-render-parity-${fixture}`)
      let inked = false, hostSheets = 0
      for (const sheet of resident.sheets) {
        const rebuilt = paged.get(sheet.name)
        // A sheet without resolved sparklines (no host, or nothing resolvable such as date-missing-range) resolves to nothing on both paths.
        if (!sheet.resolvedSparklines?.size) expect(rebuilt, `${fixture}/${sheet.name}: paged resolved nothing too`).toBeUndefined()
        else expect(rebuilt, `${fixture}/${sheet.name}: paged resolved sparklines`).toBeDefined()
        if (!sheet.sparklineGroups?.length) continue
        hostSheets++
        for (const scale of SCALES) {
          const a = pixels(sheet, scale)
          const b = pixels({ ...sheet, resolvedSparklines: rebuilt }, scale)
          const none = pixels({ ...sheet, resolvedSparklines: undefined }, scale)
          const diff = differing(a, b)
          if (diff !== 0) report.push(`${fixture}/${sheet.name} @${scale}: ${diff} differing pixels`)
          if (differing(a, none) > 0) inked = true
        }
      }
      expect(hostSheets, `${fixture} has a sparkline host sheet`).toBeGreaterThan(0)
      expect(inked, `${fixture}: ${BLANK.has(fixture) ? 'is expected to draw nothing' : 'must paint ink the no-sparkline render lacks'}`).toBe(!BLANK.has(fixture))
      if (inked) inkedFixtures++
    }
    expect(report).toEqual([])
    expect(inkedFixtures).toBe(fixtures.length - BLANK.size)
  }, 600_000)

  it('control: a one-value change in the paged resolution is visible in the pixels', async () => {
    await preloadSparklines()
    const bytes = fixtureBytes('line-basic')
    const resident = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
    const paged = await resolvePaged(bytes, 'spark-render-parity-control')
    const sheet = resident.sheets[0], rebuilt = paged.get(sheet.name)!
    const sp = [...rebuilt.values()].flatMap(row => [...row.values()])[0]
    ;(sp.values as number[])[3] += 3
    expect(differing(pixels(sheet, 0.75), pixels({ ...sheet, resolvedSparklines: rebuilt }, 0.75))).toBeGreaterThan(20)
  }, 60_000)
})
