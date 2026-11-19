// @vitest-environment node
/**
 * Review fix round 1 (codex, centred glyph key):
 * 1. A glyph key is admitted for Stage 5 only when its own case passes full `verifyCfCase` (lock state, resave receipt,
 *    calibration and the capture checks), the bottom key and the centred key alike.
 * 2. The admitted keys and the restored key agree on what they were derived from: a key input that disappears, changes
 *    or comes back is re-evaluated, never served stale from either cache layer.
 *
 * Every mutant lives in memory: `node:fs` is wrapped so a path can be overridden (new bytes) or hidden (missing) without
 * touching the committed fixtures.
 */
import { resolve } from 'node:path'
import { readFileSync } from 'node:fs'
import { afterEach, describe, expect, it, vi } from 'vitest'

const fsOverrides = vi.hoisted(() => new Map<string, Buffer | null>())

vi.mock('node:fs', async importOriginal => {
  const real = await importOriginal<typeof import('node:fs')>()
  const key = (p: unknown): string => (typeof p === 'string' ? p : p instanceof URL ? p.pathname : String(p))
  const encodingOf = (o: unknown): BufferEncoding | undefined => (typeof o === 'string' ? (o as BufferEncoding) : (o as { encoding?: BufferEncoding } | undefined)?.encoding ?? undefined)
  const existsSync = (p: Parameters<typeof real.existsSync>[0]): boolean => (fsOverrides.has(key(p)) ? fsOverrides.get(key(p)) !== null : real.existsSync(p))
  const readFileSync = ((p: Parameters<typeof real.readFileSync>[0], o?: unknown) => {
    if (fsOverrides.has(key(p))) {
      const b = fsOverrides.get(key(p))
      if (b === null || b === undefined) throw Object.assign(new Error(`ENOENT: ${key(p)}`), { code: 'ENOENT' })
      const enc = encodingOf(o)
      return enc ? b.toString(enc) : Buffer.from(b)
    }
    return real.readFileSync(p as never, o as never)
  }) as typeof real.readFileSync
  const statSync = ((p: Parameters<typeof real.statSync>[0], o?: unknown) => {
    if (fsOverrides.has(key(p))) {
      const b = fsOverrides.get(key(p))
      if (b === null || b === undefined) throw Object.assign(new Error(`ENOENT: ${key(p)}`), { code: 'ENOENT' })
      return { ...real.statSync(p as never), size: b.length, mtimeMs: Date.now() + Math.random() }
    }
    return real.statSync(p as never, o as never)
  }) as typeof real.statSync
  return { ...real, existsSync, readFileSync, statSync, default: { ...real, existsSync, readFileSync, statSync } }
})

const { CF_STAGE5_CASES } = await import('../scripts/cf-refs/cases')
const { auditStage5Manifest, stage5Publication } = await import('../scripts/cf-refs/audit')
const { computeWorksheetRangeGeometry, decodePngSync, evaluateFormulaCase, loadIconGlyphKey, readZipPartsSync, stage5IconKeys, verifyCfCase } = await import('../scripts/cf-refs/verify')

const ROOT = resolve(__dirname, '..')
const CENTRED = 'stage4-icon-glyph-key-centred'
const BOTTOM = 'stage4-icon-glyph-key'
const abs = (rel: string): string => resolve(ROOT, rel)
const sidecarPath = (id: string): string => abs(`tests/fixtures/xlsx-cf/refs/${id}.json`)
const pngPath = (id: string): string => abs(`tests/fixtures/xlsx-cf/refs/${id}.png`)
const manifest = () => JSON.parse(readFileSync(abs('tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as { cases: Array<Parameters<typeof verifyCfCase>[0]> }
const lockSidecar = (id: string): void => {
  const sc = JSON.parse(readFileSync(sidecarPath(id), 'utf8')) as Record<string, unknown>
  fsOverrides.set(sidecarPath(id), Buffer.from(`${JSON.stringify({ ...sc, lockBefore: 'locked' }, null, 2)}\n`))
}

afterEach(() => fsOverrides.clear())
// Each admission runs a full verifyCfCase of one or both key cases: 6-11 s per test on the CI runner (under 1 s locally).
vi.setConfig({ testTimeout: 120_000 })

describe('ruling 1: a key is admitted only when its own case fully verifies', () => {
  it('the committed keys verify and are both admitted (baseline)', () => {
    const m = manifest()
    for (const id of [BOTTOM, CENTRED]) expect(verifyCfCase(m.cases.find(c => c.id === id)!, { rootDir: ROOT }).kind, id).toBe('verified')
    const keys = stage5IconKeys(ROOT)
    expect(keys.bottom).toBeDefined()
    expect(keys.centre).toBeDefined()
  })

  it('a centred sidecar recorded under a locked display (codex mutant) is no key: its glyph table still rebuilds, its case does not verify', () => {
    stage5IconKeys(ROOT) // warm
    lockSidecar(CENTRED)
    const c = manifest().cases.find(x => x.id === CENTRED)!
    expect(verifyCfCase(c, { rootDir: ROOT })).toMatchObject({ kind: 'refusal', code: 'display-locked' })
    // The glyph table alone would still rebuild from the PNG (lockBefore is not one of its checks).
    const table = loadIconGlyphKey(ROOT, CENTRED)
    expect(table && !('kind' in table)).toBe(true)
    const keys = stage5IconKeys(ROOT)
    expect(keys.centre).toBeUndefined()
    expect(keys.bottom).toBeDefined()
  })

  it('with the locked centred key the Stage 5 icon cells are unmeasurable again', () => {
    lockSidecar(CENTRED)
    const t = CF_STAGE5_CASES.find(x => x.id === 'stage5-cfvo-formulas')!
    const m = manifest().cases.find(x => x.id === t.id)!
    const parts = readZipPartsSync(new Uint8Array(readFileSync(abs(m.saved))))!
    const sidecar = JSON.parse(readFileSync(sidecarPath(t.id), 'utf8')) as Record<string, unknown>
    const geom = computeWorksheetRangeGeometry(parts, t.range, { colWidthsPt: sidecar.colWidthsPt as number[], rowHeightsPt: sidecar.rowHeightsPt as number[] })!
    const png = decodePngSync(new Uint8Array(readFileSync(pngPath(t.id))))!
    const ev = evaluateFormulaCase({ ...m, formulaConfigurations: t.formulaConfigurations }, sidecar, png, geom, parts, stage5IconKeys(ROOT))
    expect(ev['icon:thresholds-formula']!.status).not.toBe('discriminating')
    expect(JSON.stringify(ev['icon:thresholds-formula']!.observed ?? {})).not.toMatch(/icon:3Arrows/)
  })

  it('a locked bottom key is no key either', () => {
    lockSidecar(BOTTOM)
    const keys = stage5IconKeys(ROOT)
    expect(keys.bottom).toBeUndefined()
    expect(keys.centre).toBeDefined()
  })

  it('with the locked centred key icon:thresholds-formula is not published', () => {
    lockSidecar(CENTRED)
    const gate = stage5Publication(auditStage5Manifest(manifest() as never, { rootDir: ROOT }))
    expect(gate.published).not.toContain('icon:thresholds-formula')
    // Fail closed, more widely than one id: the committed evaluations of the icon-bearing cases were derived with the
    // centred key, so without it those cases refuse (sidecar-mismatch) and none of their ids publish.
    fsOverrides.clear()
    const committed = new Set(stage5Publication(auditStage5Manifest(manifest() as never, { rootDir: ROOT })).published)
    expect(committed.has('icon:thresholds-formula')).toBe(true)
    // 44 since formula:hoist-aggregate-errors registered its error-propagates reading (it was 43 before), 45 since
    // bar:bounds-formula (stage5-bar-bounds-formula), 46 since formula:cfvo-scope-multi-range (anchor-a1,
    // stage5-cfvo-scope-v2; its icon set is classified against the centred key the case's own placement selects).
    expect(committed.size).toBe(46)
    for (const id of gate.published) expect(committed.has(id), id).toBe(true)
  }, 600_000)
})

describe('ruling 2: both cache layers are invalidated by the same inputs', () => {
  it('a PNG that disappears after a warm cache yields no key (outer and inner layer)', () => {
    expect(stage5IconKeys(ROOT).centre).toBeDefined()
    expect(loadIconGlyphKey(ROOT, CENTRED)).toBeDefined()
    fsOverrides.set(pngPath(CENTRED), null)
    expect(stage5IconKeys(ROOT).centre).toBeUndefined()
    expect(loadIconGlyphKey(ROOT, CENTRED)).toBeUndefined()
  })

  it('a PNG replaced with different bytes is re-evaluated, not served from either cache', () => {
    const warm = stage5IconKeys(ROOT).centre
    expect(warm).toBeDefined()
    fsOverrides.set(pngPath(CENTRED), readFileSync(pngPath(BOTTOM)))
    const after = stage5IconKeys(ROOT)
    expect(after.centre).toBeUndefined()
    const inner = loadIconGlyphKey(ROOT, CENTRED)
    expect(inner && !('kind' in inner) ? inner : undefined).not.toBe(warm)
  })

  it('a PNG restored after a cached miss yields the key again', () => {
    fsOverrides.set(pngPath(CENTRED), null)
    expect(stage5IconKeys(ROOT).centre).toBeUndefined()
    expect(loadIconGlyphKey(ROOT, CENTRED)).toBeUndefined()
    fsOverrides.delete(pngPath(CENTRED))
    expect(stage5IconKeys(ROOT).centre?.classes.size).toBe(54)
    const inner = loadIconGlyphKey(ROOT, CENTRED)
    expect(inner && !('kind' in inner) ? inner.classes.size : 0).toBe(54)
  })
})

describe('codex re-review: admission invalidation covers every file the key case verifies, its input workbook included', () => {
  const inputPath = (id: string): string => abs(manifest().cases.find(c => c.id === id)!.input)
  it('an input replaced after a warm cache is re-verified: hash-mismatch, so no key', () => {
    expect(stage5IconKeys(ROOT).centre).toBeDefined()
    fsOverrides.set(inputPath(CENTRED), readFileSync(inputPath(BOTTOM)))
    expect(verifyCfCase(manifest().cases.find(c => c.id === CENTRED)!, { rootDir: ROOT })).toMatchObject({ kind: 'refusal', code: 'hash-mismatch' })
    expect(stage5IconKeys(ROOT).centre).toBeUndefined()
  })
  it('an input that disappears after a warm cache yields no key', () => {
    expect(stage5IconKeys(ROOT).centre).toBeDefined()
    fsOverrides.set(inputPath(CENTRED), null)
    expect(stage5IconKeys(ROOT).centre).toBeUndefined()
  })
  it('an input restored after a cached miss yields the key again', () => {
    fsOverrides.set(inputPath(CENTRED), null)
    expect(stage5IconKeys(ROOT).centre).toBeUndefined()
    fsOverrides.delete(inputPath(CENTRED))
    expect(stage5IconKeys(ROOT).centre?.classes.size).toBe(54)
  })
})
