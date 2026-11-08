/** Stage-0 Task 1 harness tests. Run under BOTH backends:
 *  - `bun run s0:test`      → real `canvas` (node-canvas)
 *  - `bun run s0:test:napi` → `canvas` aliased to the @napi-rs/canvas napi shim
 * The shim itself is asserted directly, so both runs prove the same exports. */
import { execSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import { createCanvas, loadImage, registerFont, Canvas, Image } from '../../scripts/stream-s0/napi-shim'
import { listCorpus, loadUnits, renderUnit, type Unit } from '../../scripts/stream-s0/corpus'
import { writeResult } from '../../scripts/stream-s0/results'

// The corpus is a gitignored working directory: when it is absent, the corpus
// tests skip (with a visible message) instead of failing the harness.
const corpusFiles = listCorpus()
it.skipIf(corpusFiles.length === 0)('listCorpus finds the corpus documents', () => {
  // The plan expected 49 files; the owner's corpus dir holds exactly 2 extra
  // Office lock artifacts (`~$docx-having-images.docx`,
  // `~$i-Bug66263-table.docx`) that `listCorpus` correctly excludes — the
  // documents are 11 docx + 9 xlsx + 26 pptx = 46. The test pins THAT so the
  // harness reflects the real inventory; codex-r8's corpus was 49 including
  // those two lock files.
  expect(corpusFiles).toHaveLength(46)
  const by = (kind: string) => corpusFiles.filter((f) => f.kind === kind).length
  expect(by('docx')).toBe(11)
  expect(by('xlsx')).toBe(9)
  expect(by('pptx')).toBe(26)
  for (const f of corpusFiles) expect(f.name.startsWith('~$')).toBe(false)
})

it('shim exposes node-canvas names', () => {
  expect(typeof createCanvas).toBe('function')
  expect(typeof loadImage).toBe('function')
  expect(typeof registerFont).toBe('function')
  expect(typeof Image).toBe('function')
  expect(typeof Canvas).toBe('function')
  const ctx = createCanvas(4, 4).getContext('2d')!
  expect(ctx).toBeTruthy()
  // Shape deliberately matches the plan: full 4×4 RGBA readback.
  expect(ctx.getImageData(0, 0, 4, 4).data.length).toBe(64)
})

// The scale test needs the file list; skip with the corpus when it is empty.
it.skipIf(corpusFiles.length === 0)('renderUnit scales 2× exactly', async () => {
  // Deterministic pick: the first docx whose first unit loads successfully.
  let found: Awaited<ReturnType<typeof loadUnits>> | undefined
  for (const file of corpusFiles) {
    if (file.kind !== 'docx') continue
    try { found = await loadUnits(file) } catch { continue }
    if (found.length > 0) break
    found = undefined
  }
  expect(found, 'at least one docx corpus file must yield units').toBeTruthy()
  const unit: Unit = found![0]
  const at1 = await renderUnit(unit, 1)
  const at2 = await renderUnit(unit, 2)
  if ('skipped' in at1 || 'skipped' in at2) {
    // A docx unit over 8192 px at scale 2 legitimately skips; use a lower
    // scale bound and assert scale consistency another way.
    const atHalf = await renderUnit(unit, 0.5)
    expect('skipped' in atHalf).toBe(false)
    const atQuarter = await renderUnit(unit, 0.25)
    const b1 = atHalf, b2 = atQuarter as typeof atHalf
    expect(Math.abs(b1.width - 2 * b2.width)).toBeLessThanOrEqual(1)
    expect(Math.abs(b1.height - 2 * b2.height)).toBeLessThanOrEqual(1)
    return
  }
  expect(Math.abs(at2.width - 2 * at1.width)).toBeLessThanOrEqual(1)
  expect(Math.abs(at2.height - 2 * at1.height)).toBeLessThanOrEqual(1)
})

it('writeResult writes status, gate and gitSha', () => {
  writeResult('harness-task1', { status: 'REPORTED' })
  const out = JSON.parse(readFileSync(join('docs/validation/streaming-s0', 'harness-task1.json'), 'utf8')) as {
    gate: string
    status: string
    backend: string
    gitSha: string
    generatedAt: string
  }
  expect(out.gate).toBe('harness-task1')
  expect(out.status).toBe('REPORTED')
  expect(['node-canvas', '@napi-rs/canvas']).toContain(out.backend)
  expect(out.gitSha).toMatch(/^[0-9a-f]{40}$/)
  expect(() => execSync('git rev-parse HEAD', { encoding: 'utf8' })).not.toThrow()
})
