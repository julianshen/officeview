// @vitest-environment node
/**
 * Task 11 fix round 1, finding 4: interactive/paged compare modes must be
 * actual viewport adapters (run the data path, bind outputs to workbook and
 * oracle hashes), not unconditional stubs. Pixel diff stays pending without
 * Excel viewport refs, but the adapter must prove the path runs and record
 * attributable hashes.
 */
import { describe, expect, it } from 'vitest'
import { createHash } from 'node:crypto'
import { resolve, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { compareCfCase } from '../scripts/cf-compare'
import type { CfManifest } from '../scripts/cf-refs/cases'
import { readFileSync } from 'node:fs'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')

function loadManifest(): CfManifest {
  return JSON.parse(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
}

function oracleHashOf(c: { expected: unknown; markers: unknown }): string {
  return createHash('sha256').update(JSON.stringify({ expected: c.expected, markers: c.markers }), 'utf8').digest('hex')
}

describe('viewport adapters run the data path and bind hashes', () => {
  it('interactive adapter prepares resident viewport states with hashes', async () => {
    const manifest = loadManifest()
    const c = manifest.cases.find(entry => entry.id === 'stage1-cellis-literals')!
    const res = await compareCfCase(c, 'interactive', { rootDir: REPO_ROOT })
    expect(res.outcome).toBe('pass')
    expect(res.outputHash, 'adapter records an output hash').toMatch(/^[0-9a-f]{64}$/)
    expect(res.workbookHash, 'adapter binds the workbook hash').toBe(c.expected.savedSha256)
    expect(res.oracleRef, 'adapter binds the oracle case').toContain('stage1-cellis-literals')
    expect(res.oracleHash, 'adapter binds oracle content').toBe(oracleHashOf(c))
    expect(res.reads, 'adapter ran real reads').toBeGreaterThan(0)
    // Deterministic: same inputs, same hash.
    const again = await compareCfCase(c, 'interactive', { rootDir: REPO_ROOT })
    expect(again.outputHash).toBe(res.outputHash)

    // A pending case returns recipe-pending
    const pendingCase = { ...c, status: 'recipe-pending' as const }
    const resPending = await compareCfCase(pendingCase, 'interactive', { rootDir: REPO_ROOT })
    expect(resPending.outcome).toBe('recipe-pending')
  }, 60000)

  it('interactive output hash is content-sensitive across cases', async () => {
    const manifest = loadManifest()
    const a = manifest.cases.find(entry => entry.id === 'stage1-cellis-literals')!
    const b = manifest.cases.find(entry => entry.id === 'stage1-numfmt-accessibility')!
    const ra = await compareCfCase(a, 'interactive', { rootDir: REPO_ROOT })
    const rb = await compareCfCase(b, 'interactive', { rootDir: REPO_ROOT })
    expect(ra.outputHash).toMatch(/^[0-9a-f]{64}$/)
    expect(rb.outputHash).toMatch(/^[0-9a-f]{64}$/)
    expect(rb.outputHash, 'different oracle content hashes differently').not.toBe(ra.outputHash)
  }, 60000)

  it('paged adapter reads worker viewport states with hashes', async () => {
    const manifest = loadManifest()
    const c = manifest.cases.find(entry => entry.id === 'stage1-cellis-literals')!
    const res = await compareCfCase(c, 'paged', { rootDir: REPO_ROOT })
    expect(res.outcome).toBe('pass')
    expect(res.outputHash, 'adapter records an output hash').toMatch(/^[0-9a-f]{64}$/)
    expect(res.workbookHash, 'adapter binds the workbook hash').toBe(c.expected.savedSha256)
    expect(res.oracleRef, 'adapter binds the oracle case').toContain('stage1-cellis-literals')
    expect(res.oracleHash, 'adapter binds oracle content').toBe(oracleHashOf(c))
    expect(res.reads, 'adapter ran real reads').toBeGreaterThan(0)
    const again = await compareCfCase(c, 'paged', { rootDir: REPO_ROOT })
    expect(again.outputHash).toBe(res.outputHash)

    // A pending case returns recipe-pending
    const pendingCase = { ...c, status: 'recipe-pending' as const }
    const resPending = await compareCfCase(pendingCase, 'paged', { rootDir: REPO_ROOT })
    expect(resPending.outcome).toBe('recipe-pending')
  }, 60000)

  it('interactive adapter fails when expected fill is mismatched (changed rendered fill fails gate)', async () => {
    const manifest = loadManifest()
    const c = manifest.cases.find(entry => entry.id === 'stage1-cellis-literals')!
    const mutated = {
      ...c,
      status: 'recipe-pending' as const,
      expected: {
        ...c.expected,
        cells: c.expected.cells.map(cell => cell.ref === 'B6' ? { ...cell, fillRgb: 'FF0000FF' } : cell),
      },
    }
    const res = await compareCfCase(mutated, 'interactive', { rootDir: REPO_ROOT })
    expect(res.outcome).toBe('fail')
    expect(res.reason).toContain('fill mismatch: expected FF0000FF, got FFFFC7CE')
  }, 60000)

  it('paged adapter fails when expected fill is mismatched (changed rendered fill fails gate)', async () => {
    const manifest = loadManifest()
    const c = manifest.cases.find(entry => entry.id === 'stage1-cellis-literals')!
    const mutated = {
      ...c,
      status: 'recipe-pending' as const,
      expected: {
        ...c.expected,
        cells: c.expected.cells.map(cell => cell.ref === 'B6' ? { ...cell, fillRgb: 'FF0000FF' } : cell),
      },
    }
    const res = await compareCfCase(mutated, 'paged', { rootDir: REPO_ROOT })
    expect(res.outcome).toBe('fail')
    expect(res.reason).toContain('fill mismatch: expected FF0000FF, got FFFFC7CE')
  }, 60000)

  it('adapters fail when control cell receives conditional formatting', async () => {
    const manifest = loadManifest()
    const c = manifest.cases.find(entry => entry.id === 'stage1-cellis-literals')!
    const mutated = {
      ...c,
      status: 'recipe-pending' as const,
      expected: {
        ...c.expected,
        positiveRefs: c.expected.positiveRefs.filter(r => r !== 'B6'),
        negativeRefs: [...c.expected.negativeRefs, 'B6'],
        cells: c.expected.cells.map(cell => cell.ref === 'B6' ? { ref: 'B6', controlRef: cell.controlRef, matched: false } : cell),
      },
    }
    const resI = await compareCfCase(mutated, 'interactive', { rootDir: REPO_ROOT })
    expect(resI.outcome).toBe('fail')
    expect(resI.reason).toContain('Control cell B6 expected matched: false, but CF formatting applied')

    const resP = await compareCfCase(mutated, 'paged', { rootDir: REPO_ROOT })
    expect(resP.outcome).toBe('fail')
    expect(resP.reason).toContain('Control cell B6 expected matched: false, but CF formatting applied')
  }, 60000)
})

describe('print Paintable carries CF fill', () => {
  it('print-mode rasterization paints the winning CF fill', async () => {
    const { paintCfCapture } = await import('../scripts/cf-compare')
    const { OfficePackage } = await import('../src/core/zip')
    const { parseXlsx } = await import('../src/xlsx/parse')
    const pkg = await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/stage1-cellis-literals.xlsx'))))
    const doc = await parseXlsx(pkg)
    const canvas = await paintCfCapture(doc, 0, 'A1:O26', 'print', 1, 800, 600)
    const data = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data
    let pink = 0
    for (let i = 0; i < data.length; i += 4) {
      if (data[i]! > 200 && data[i + 1]! < 220 && data[i + 2]! > 200) pink++
    }
    // CF pink fill survives print rasterization (not just the notice/PDF shell).
    expect(pink).toBeGreaterThan(500)
  }, 60000)
})

describe('Stage 2 scale cases verify exact Excel-observed fills on interactive and paged adapters', () => {
  for (const mode of ['interactive', 'paged'] as const) {
    it(`${mode}: every measured scale case passes against observed fills`, async () => {
      const manifest = loadManifest()
      for (const c of manifest.cases.filter(entry => entry.stage === 2 && entry.status === 'measured')) {
        const res = await compareCfCase(c, mode, { rootDir: REPO_ROOT })
        expect(res.outcome, `${c.id}: ${res.reason}`).toBe('pass')
        expect(res.outputHash).toMatch(/^[0-9a-f]{64}$/)
      }
    }, 120000)
  }

  it('a wrong painted fill fails against the observed sidecar (mutant)', async () => {
    const { verifyCaseCells } = await import('../scripts/cf-compare')
    const c = loadManifest().cases.find(entry => entry.id === 'stage2-scale-interpolation')!
    const sidecar = JSON.parse(readFileSync(resolve(REPO_ROOT, c.captures[0]!.sidecarPath), 'utf8')) as { observedCells: Array<{ ref: string; fillRgb?: string }> }
    const target = sidecar.observedCells.find(cell => cell.fillRgb && cell.fillRgb !== 'FFFFFFFF')!
    const { parseA1 } = await import('../scripts/cf-refs/verify')
    const pos = parseA1(target.ref)!
    const mismatch = verifyCaseCells(c, REPO_ROOT, (row, col) => row === pos.row && col === pos.col
      ? { applied: true, style: { fillColor: 'FF000000' } }
      : { applied: false, unresolved: true })
    expect(mismatch).toContain(`Cell ${target.ref} observed Excel fill`)
  })
})
