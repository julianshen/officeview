/**
 * XLSX Conditional Formatting reference comparator (plan Task 1 & Task 8).
 *
 *   bun scripts/cf-compare.ts [--stage 1] [--mode static|interactive|paged|print|export] [--case <id>] [--out <dir>]
 *
 * Reuses the sparkline comparator primitives (`diffFrames`, `withinTolerance`, `COMPARE_SCALE`, `STRONG_DELTA`,
 * `loadFrame`, `crop`) from `scripts/sparkline-compare.ts`, validates manifest hashes and strict capture
 * dimensions/scale (`1 px/pt` copy-picture screen/printer references; `PRINT_SCALE = 2` for print rasterization),
 * and enforces per-case tolerances (`meanAbs <= 2.0/255`, `rowMax <= 4.0/255`, `strong <= round(0.003 * W * H)`).
 * Before Task 8 wires CF into `renderSheet`, measured cases report `implementation-pending` rather than claiming pass.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, renderSheet } from '../src/xlsx/render'
import type { CfManifest, CfOracleCase } from './cf-refs/cases'
import { CF_MANIFEST_PATH } from './cf-refs/generate'
import { verifyCfCase } from './cf-refs/verify'
import {
  COMPARE_SCALE,
  CSS_PX_PER_PT,
  STRONG_DELTA,
  SUPPORTED_PX_PER_PT,
  crop,
  diffFrames,
  loadFrame,
  withinTolerance,
  type BlockDiff,
  type Frame,
  type Tolerance,
} from './sparkline-compare'

export { COMPARE_SCALE, CSS_PX_PER_PT, STRONG_DELTA, SUPPORTED_PX_PER_PT, crop, diffFrames, loadFrame, withinTolerance }

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
export const PRINT_SCALE = 2

export type CfCompareMode = 'static' | 'interactive' | 'paged' | 'print' | 'export'

export interface CfCompareResult {
  id: string
  mode: CfCompareMode
  outcome: 'pass' | 'fail' | 'recipe-pending' | 'implementation-pending' | 'unsupported'
  diff?: BlockDiff
  tolerance?: Tolerance
  reason?: string
}

export function resolvePixelTolerance(c: CfOracleCase, width: number, height: number): Tolerance {
  const base = c.expected.tolerance ?? { meanAbs: 2.0 / 255, rowMax: 4.0 / 255, strong: 0.003 }
  const strongPixels = base.strong < 1 ? Math.max(1, Math.round(base.strong * width * height)) : Math.round(base.strong)
  return {
    meanAbs: base.meanAbs,
    rowMax: base.rowMax,
    strong: strongPixels,
  }
}

export async function compareCfCase(
  c: CfOracleCase,
  mode: CfCompareMode = 'static',
  options: { rootDir?: string; outDir?: string } = {},
): Promise<CfCompareResult> {
  const rootDir = options.rootDir ?? REPO_ROOT
  const check = verifyCfCase(c, { rootDir })

  if (check.kind === 'refusal') {
    if (check.code === 'unsupported-capability') {
      return { id: c.id, mode, outcome: 'unsupported', reason: check.reason }
    }
    if (check.code === 'semantic-evidence-pending' && check.structuralValid) {
      return { id: c.id, mode, outcome: 'recipe-pending', reason: check.reason }
    }
    return { id: c.id, mode, outcome: 'fail', reason: `${check.code}: ${check.reason}` }
  }

  const appearance = mode === 'print' ? 'printer' : 'screen'
  const cap = c.captures.find(x => x.appearance === appearance) ?? c.captures[0]
  if (!cap) {
    return { id: c.id, mode, outcome: 'fail', reason: `No ${appearance} capture configured for ${c.id}` }
  }

  const sidecarAbs = resolve(rootDir, cap.sidecarPath)
  const sidecar = JSON.parse(readFileSync(sidecarAbs, 'utf8')) as { pxPerPt?: number; scale?: number }
  if (!SUPPORTED_PX_PER_PT.includes(Number(sidecar.pxPerPt))) {
    return {
      id: c.id,
      mode,
      outcome: 'fail',
      reason: `Unsupported capture density pxPerPt=${String(sidecar.pxPerPt)}`,
    }
  }

  const savedAbs = resolve(rootDir, c.saved)
  const pkg = await OfficePackage.load(new Uint8Array(readFileSync(savedAbs)))
  const doc = await parseXlsx(pkg)
  const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]
  if (!sheet) {
    return { id: c.id, mode, outcome: 'fail', reason: `Sheet ${c.sheet} missing in ${c.saved}` }
  }

  // Until Task 3 & Task 8 attach and render `sheet.conditionalFormatting`, report implementation-pending
  // rather than claiming a green pixel match.
  const hasCfRendered = Array.isArray((sheet as unknown as { conditionalFormatting?: unknown[] }).conditionalFormatting)
  if (!hasCfRendered) {
    return {
      id: c.id,
      mode,
      outcome: 'implementation-pending',
      reason: 'Task 3/8 conditionalFormatting parser and renderer not yet wired',
    }
  }

  const refFrame = await loadFrame(resolve(rootDir, cap.pngPath))
  const m = computeMetrics(sheet)
  const scale = Number(sidecar.pxPerPt) / CSS_PX_PER_PT
  const canvas = createCanvas(Math.max(1, Math.round(m.widthPx * scale)), Math.max(1, Math.round(m.heightPx * scale)))
  const ctx = canvas.getContext('2d')
  ctx.scale(scale, scale)
  renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, m)
  const fullData = ctx.getImageData(0, 0, canvas.width, canvas.height).data as unknown as Uint8ClampedArray
  const oursFull: Frame = { width: canvas.width, height: canvas.height, data: fullData }
  const oursCropped = crop(oursFull, 0, 0, refFrame.width, refFrame.height)

  const tol = resolvePixelTolerance(c, refFrame.width, refFrame.height)
  const diff = diffFrames(refFrame, oursCropped)
  const pass = withinTolerance(diff, tol)

  if (options.outDir) {
    mkdirSync(options.outDir, { recursive: true })
    writeFileSync(join(options.outDir, `${c.id}.${mode}.ours.png`), canvas.toBuffer('image/png'))
  }

  return {
    id: c.id,
    mode,
    outcome: pass ? 'pass' : 'fail',
    diff,
    tolerance: tol,
    ...(!pass
      ? {
          reason: `Pixel diff exceeded tolerance: meanAbs=${diff.meanAbs.toFixed(4)} rowMax=${diff.rowMax.toFixed(4)} strong=${diff.strong}`,
        }
      : {}),
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2)
  const stage = Number(args.includes('--stage') ? args[args.indexOf('--stage') + 1] : 1)
  const modeArg = (args.includes('--mode') ? args[args.indexOf('--mode') + 1] : 'static') as CfCompareMode
  const only = args.includes('--case') ? args[args.indexOf('--case') + 1] : undefined
  const outDir = args.includes('--out') ? resolve(args[args.indexOf('--out') + 1]) : undefined

  if (!['static', 'interactive', 'paged', 'print', 'export'].includes(modeArg)) {
    console.error(`Usage: bun scripts/cf-compare.ts [--stage N] [--mode static|interactive|paged|print|export] [--case id] [--out dir]`)
    process.exit(2)
  }

  if (!existsSync(CF_MANIFEST_PATH)) {
    console.error(`Missing manifest at ${CF_MANIFEST_PATH}`)
    process.exit(1)
  }

  const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
  let passCount = 0
  let pendingCount = 0
  let unsupportedCount = 0
  let failCount = 0

  for (const c of manifest.cases) {
    if (c.stage > stage || (only && c.id !== only)) continue
    const res = await compareCfCase(c, modeArg, { rootDir: REPO_ROOT, ...(outDir ? { outDir } : {}) })
    if (res.outcome === 'pass') passCount++
    else if (res.outcome === 'recipe-pending' || res.outcome === 'implementation-pending') pendingCount++
    else if (res.outcome === 'unsupported') unsupportedCount++
    else failCount++

    const stats = res.diff
      ? ` meanAbs=${res.diff.meanAbs.toFixed(4)} rowMax=${res.diff.rowMax.toFixed(4)} strong=${res.diff.strong}`
      : ''
    const note = res.reason ? ` (${res.reason})` : ''
    console.log(`${res.id} [${res.mode}]: ${res.outcome}${stats}${note}`)
  }

  console.log(
    `Total [${modeArg}]: ${passCount} pass, ${pendingCount} pending, ${unsupportedCount} unsupported, ${failCount} fail`,
  )
  process.exit(failCount > 0 ? 1 : 0)
}
