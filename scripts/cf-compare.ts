/**
 * XLSX Conditional Formatting reference comparator (plan Task 1 & Task 8).
 *
 *   bun scripts/cf-compare.ts [--stage 1] [--mode static|interactive|paged|print|export] [--case <id>] [--out <dir>] [--evidence production|audit]
 *
 * `--evidence audit` (Task 22) runs the viewer under the Stage 3 audit's OWN result instead of the published set: the
 * verified capabilities, the selected bar layout and the measured automatic colours, injected exactly on top of
 * production's Stage 1/2 set before any workbook is parsed. Until Task 23 publishes, that is the only way to compare
 * bars against Excel; the default stays `production`.
 *
 * Reuses the sparkline comparator primitives (`diffFrames`, `withinTolerance`, `COMPARE_SCALE`, `STRONG_DELTA`,
 * `loadFrame`, `crop`) from `scripts/sparkline-compare.ts`, validates manifest hashes and strict capture
 * dimensions/scale (`1 px/pt` copy-picture screen/printer references; `PRINT_SCALE = 2` for print rasterization),
 * and enforces per-case tolerances (`meanAbs <= 2.0/255`, `rowMax <= 4.0/255`, `strong <= round(0.003 * W * H)`).
 * Static/print/export use the shared prepared Paintable; unmeasured references remain pending.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { getPaintables } from '../src/render/paint'
import { computeMetrics } from '../src/xlsx/render'
import type { XlsxDocument } from '../src/xlsx/types'
import { parseRangeA1, parseA1, type CfObservedBar, type CfObservedIcon } from './cf-refs/verify'
import { stage5CellAgreement, stage5Conflicts, stage5CounterfactualOf, stage5ScaleToken, stage5VisualDetail, stage5VisualToken, type Stage5Production, type Stage5Truth } from './cf-refs/stage5-withhold'
import type { CfManifest, CfOracleCase } from './cf-refs/cases'
import { CF_MANIFEST_PATH } from './cf-refs/generate'
import { ICON_NO_ICON_MEMBER, iconTwinSetMembers, isCellInSqref, observedIconClass, readZipPartsSync, verifyCfCase } from './cf-refs/verify'
import { STAGE2_WITHHELD_FINDINGS, stage2NumSegmentShape } from './cf-refs/stage2-withhold'
import { CF_ICON_SET_NAMES } from '../src/xlsx/cf/model'
import { caseDocumentEvaluationOf, caseNowOf, caseParseOptionsOf, withCaseZone } from './cf-refs/case-clock'
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

/** Actual print rasterization is 2x; reference-density conversion happens only
 * after painting. Crop at the declared source range, not implicitly at A1. */
export async function paintCfCapture(doc: XlsxDocument, sheetIndex: number, range: string, mode: CfCompareMode, pxPerPt: number, width: number, height: number) {
  const rect = parseRangeA1(range)
  if (!rect) throw new Error('Invalid CF capture range')
  const units = await getPaintables(doc)
  try {
    const unit = units[sheetIndex]
    if (!unit || unit.refusal?.length) throw new Error('CF static render refused: ' + JSON.stringify(unit?.refusal))
    const metrics = computeMetrics(doc.sheets[sheetIndex]!)
    const x = metrics.colWidthsPx.slice(0, rect.minCol).reduce((a, b) => a + b, 0)
    const y = metrics.rowHeightsPx.slice(0, rect.minRow).reduce((a, b) => a + b, 0)
    const density = pxPerPt / CSS_PX_PER_PT, scale = mode === 'print' ? PRINT_SCALE : density
    const full = createCanvas(Math.max(1, Math.round(unit.spec.widthPx * scale)), Math.max(1, Math.round(unit.spec.heightPx * scale)))
    const ctx = full.getContext('2d'); ctx.scale(scale, scale); unit.paint(ctx as unknown as CanvasRenderingContext2D)
    const result = createCanvas(width, height)
    result.getContext('2d').drawImage(full, x * scale, y * scale, width * scale / density, height * scale / density, 0, 0, width, height)
    return result
  } finally { units.dispose() }
}

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
  /** SHA-256 of the deterministic viewport output (states/descriptors). */
  outputHash?: string
  /** Workbook bytes the output was produced from (manifest savedSha256). */
  workbookHash?: string
  /** Oracle case the output is bound to (id + status). */
  oracleRef?: string
  /** SHA-256 of the oracle content (expected + markers) the output is bound to. */
  oracleHash?: string
  /** Source reads issued to produce the output. */
  reads?: number
  /** Task 39 fix round 1 (F2): on a Stage 5 case, the formula cells Excel formats, how many production painted, and how
   * many it withheld with a recorded justification (`stage5WithholdJustification`). */
  stage5?: { formulaMatched: number; painted: number; justifiedWithheld: number }
}

function sha256Hex(lines: string[]): string {
  return createHash('sha256').update(lines.join('\n'), 'utf8').digest('hex')
}

/**
 * Interactive viewport adapter (resident route): prepares the CF text scope
 * over the case range and hashes per-cell descriptors + applied flags. Runs
 * the real resident data path and binds the output to workbook/oracle hashes.
 * Pixel diff stays pending without Excel viewport refs; behavior is covered
 * by tests/browser/xlsx-cf.spec.ts.
 */
export interface ActualCellOutcome {
  applied: boolean
  /** Task 30 fix round: an icon-set rule covers the cell. Then `icon` is production's claim (the family and id it
   * evaluated, or null for no icon) and `iconWithheld` marks a tainted or unresolved icon slot. */
  iconCell?: boolean
  icon?: { set: string; iconId: number } | null
  iconWithheld?: boolean
  unresolved?: boolean
  /** Fix round 2 (R2): the write slots production left unresolved at the cell (compose.ts groups or paths; `ALL` for a
   * cell-wide refusal). Absent: derived from `unresolved` (`ALL` when true). */
  unresolvedSlots?: readonly string[]
  /** Orchestrator ruling (codex r5): every bar and icon descriptor field (`stage5VisualDetail`); required on a Stage 5 cell. */
  visualDetail?: string
  /** Fix round 5 (Stage 5): the leg's resolved bar descriptor (null: none), its resolved DXF properties and its colour-scale
   * colour (null: none), so equality with the counterfactual covers every component. On a Stage 5 cell `icon` is set
   * (null: none) whether or not an icon rule covers it. */
  bar?: Readonly<{ startFraction: number; endFraction: number }> | null
  properties?: Readonly<Record<string, unknown>>
  scaleArgb?: number | null
  style?: {
    fillColor?: string
    color?: string
    bold?: boolean
    underline?: string
    strike?: boolean
    borders?: { top?: string; bottom?: string; left?: string; right?: string }
  }
}

export function verifyCellExpectations(
  c: CfOracleCase,
  actualAt: (row: number, col: number) => ActualCellOutcome,
  /** Task 30 fix round: Excel's observed icon class per ref (`observedIconClass`: a class, a twin set, or null). */
  iconObserved?: ReadonlyMap<string, string | null>,
  /** Fix round 4: on a Stage 5 case, the counterfactual of the real compositor per ref (`stage5CounterfactualOf`). */
  stage5Truths?: ReadonlyMap<string, Stage5Truth>,
): string | undefined {
  for (const exp of c.expected.cells) {
    const pos = parseA1(exp.ref)
    if (!pos) return `Invalid cell reference: ${exp.ref}`
    const actual = actualAt(pos.row, pos.col)
    // Fix round 5: on a measured Stage 5 case (the population the audit rows and conformance compare) every expected cell,
    // resolved or not, must equal the counterfactual run of the real compositor in every component (fill, each font path,
    // colour scale, bar, icon, unresolved writes); a mismatch fails here, before any acceptance path below. An unmeasured
    // case (no Excel evidence; e.g. a hand-authored unpaired x14 twin production refuses structurally) keeps the round-4
    // check on its matched-but-unresolved cells.
    if (c.stage === 5 && (c.status === 'measured' || (exp.matched && actual.unresolved))) {
      const v = stage5CellAgreement(stage5TokenOfActual(actual), undefined, stage5Truths?.get(exp.ref))
      if (!v.agrees) return `Cell ${exp.ref}: ${v.disagreement}`
    }
    // An icon is no cell style (S4-I): an icon cell is compared on whether an icon is drawn (Excel's `iconApplied`), and
    // on its style expectations only when it has some. A withheld icon slot is allowed (fail closed); the glyph class
    // itself is compared in tests/xlsx-cf-oracle-conformance.test.ts.
    // Fix round 2 (R2): on a Stage 5 case the icon slot never excuses the whole cell; its slots are checked one by one
    // in verifyCaseCells against Excel's observation (stage5CellAgreement).
    if (actual.iconCell && c.stage !== 5) {
      if (actual.unresolved || actual.iconWithheld) continue
      const drawn = actual.icon !== undefined && actual.icon !== null && actual.icon.set !== 'NoIcons'
      // Excel's evidence is the sidecar's observed class (the `iconApplied` labels are seeds, unreliable on faint and
      // text-only cells): an icon, no icon (null, only on a classified cell), or a twin set that may hold no icon.
      const obs = iconObserved?.get(exp.ref.toUpperCase())
      const twins = iconTwinSetMembers(obs)
      const decided = obs !== undefined && obs !== 'icon-class-none' && obs !== 'icon-class-ambiguous' && obs !== 'unclassified'
        && !(obs === null && exp.iconEvidence?.bucket !== 'pixel-classified') && !(twins && twins.includes(ICON_NO_ICON_MEMBER))
      if (decided && drawn !== (obs !== null)) return `Cell ${exp.ref}: Excel ${obs === null ? 'draws no icon' : `draws ${obs}`}, but production ${drawn ? 'draws an icon' : 'draws no icon'}`
      const styled = exp.fillRgb !== undefined || exp.fontRgb !== undefined || exp.bold !== undefined || exp.underline !== undefined || exp.strikethrough !== undefined || exp.borderTop !== undefined || exp.borderBottom !== undefined
      if (!styled) continue
    }
    if (exp.matched) {
      // Task 39 (fix rounds 1-5): a Stage 5 unresolved cell was checked against the counterfactual above. A Stage 2 one is
      // justified against Excel's observation in verifyCaseCells, and may never carry a painted fill (codex final P2).
      if (actual.unresolved) {
        if (c.stage === 2 && actual.style?.fillColor) return `Cell ${exp.ref} is reported withheld but paints fill ${String(actual.style.fillColor).toUpperCase()}`
        continue
      }
      if (!actual.applied) {
        return `Cell ${exp.ref} expected matched: true, but no CF formatting applied`
      }
      const style = actual.style ?? {}
      if (exp.fillRgb !== undefined) {
        const fill = style.fillColor ? String(style.fillColor).toUpperCase() : undefined
        if (fill !== exp.fillRgb.toUpperCase()) {
          return `Cell ${exp.ref} fill mismatch: expected ${exp.fillRgb}, got ${fill ?? 'none'}`
        }
      }
      if (exp.fontRgb !== undefined) {
        const font = style.color ? String(style.color).toUpperCase() : undefined
        if (font !== exp.fontRgb.toUpperCase()) {
          return `Cell ${exp.ref} font mismatch: expected ${exp.fontRgb}, got ${font ?? 'none'}`
        }
      }
      if (exp.bold !== undefined && Boolean(style.bold) !== exp.bold) {
        return `Cell ${exp.ref} bold mismatch: expected ${exp.bold}, got ${Boolean(style.bold)}`
      }
      if (exp.underline !== undefined) {
        const actualUnderline = style.underline ?? 'none'
        if (actualUnderline !== exp.underline) {
          return `Cell ${exp.ref} underline mismatch: expected ${exp.underline}, got ${actualUnderline}`
        }
      }
      if (exp.strikethrough !== undefined && Boolean(style.strike) !== exp.strikethrough) {
        return `Cell ${exp.ref} strikethrough mismatch: expected ${exp.strikethrough}, got ${Boolean(style.strike)}`
      }
      if (exp.borderTop !== undefined && Boolean(style.borders?.top) !== exp.borderTop) {
        return `Cell ${exp.ref} borderTop mismatch: expected ${exp.borderTop}, got ${Boolean(style.borders?.top)}`
      }
      if (exp.borderBottom !== undefined && Boolean(style.borders?.bottom) !== exp.borderBottom) {
        return `Cell ${exp.ref} borderBottom mismatch: expected ${exp.borderBottom}, got ${Boolean(style.borders?.bottom)}`
      }
    } else {
      if (actual.applied) {
        return `Control cell ${exp.ref} expected matched: false, but CF formatting applied`
      }
    }
  }
  return undefined
}

/**
 * Stage 2 scale expectations carry seed flags (`matched`/`scaleApplied`) that Excel's
 * display-format measurement may contradict (boolean/text cells and DXF-stopped cells keep
 * the base fill). The measured sidecar `observedCells[].fillRgb` is the oracle, so Stage 2
 * cases are verified against it exactly: an unpainted cell must observe FFFFFFFF.
 * Cells the product withholds (unresolved) are reported, never counted as matches.
 */
/** The numeric cell values of a Stage 2 case's saved worksheet (plain `<v>` numbers; text, booleans and errors omitted). */
export function stage2SavedNumbers(c: CfOracleCase, rootDir: string): Map<string, number> {
  const out = new Map<string, number>()
  // No saved workbook: no value, so the num-segment condition excuses nothing (fail closed).
  if (!existsSync(resolve(rootDir, c.saved))) return out
  const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(rootDir, c.saved))))
  const xml = parts?.get('xl/worksheets/sheet1.xml')
  if (!xml) return out
  for (const m of new TextDecoder().decode(xml).matchAll(/<c r="([A-Z]+\d+)"((?:\s+[a-z]+="[^"]*")*)\s*>(?:<f[^>]*>[^<]*<\/f>|<f[^>]*\/>)?<v>([^<]*)<\/v>/g)) {
    const t = /\st="([^"]*)"/.exec(m[2] ?? '')?.[1]
    if (t !== undefined && t !== 'n') continue
    const v = Number(m[3])
    if (Number.isFinite(v)) out.set(m[1]!, v)
  }
  return out
}

export function stage2ObservedFills(c: CfOracleCase, rootDir: string): Map<string, string> | undefined {
  if (c.stage !== 2) return undefined
  const cap = c.captures.find(x => x.appearance === 'screen') ?? c.captures[0]
  if (!cap?.sidecarPath) return undefined
  const path = resolve(rootDir, cap.sidecarPath)
  if (!existsSync(path)) return undefined
  const sidecar = JSON.parse(readFileSync(path, 'utf8')) as { observedCells?: Array<{ ref: string; fillRgb?: string; fillSource?: string; hidden?: boolean }> }
  const out = new Map<string, string>()
  // Same trust boundary as tests/xlsx-cf-oracle-conformance.test.ts: Excel display-format fills of visible cells only.
  for (const cell of sidecar.observedCells ?? []) {
    if (cell.fillRgb && cell.fillSource === 'excel-display-format' && !cell.hidden) out.set(cell.ref, cell.fillRgb.toUpperCase())
  }
  return out
}

/**
 * Stage 5 observed fills (Ruling S5-A): Excel's DisplayFormat interior colour per designated cell, twin and control, from the
 * sidecar's `formulaObservations`; a cell whose DisplayFormat fill equals its base fill observes FFFFFFFF (no CF fill), as
 * Stage 2 does. Only a Stage 5 case with a readable screen sidecar yields a map. Production's withheld cells are reported by
 * `verifyCaseCells`, never counted as matches.
 */
/** Fix round 2 (R2): Excel's DisplayFormat strikethrough per observed Stage 5 cell. */
export function stage5ObservedStrikes(c: CfOracleCase, rootDir: string): Map<string, boolean> {
  const out = new Map<string, boolean>()
  const cap = c.captures.find(x => x.appearance === 'screen') ?? c.captures[0]
  if (!cap?.sidecarPath || !existsSync(resolve(rootDir, cap.sidecarPath))) return out
  try {
    const sidecar = JSON.parse(readFileSync(resolve(rootDir, cap.sidecarPath), 'utf8')) as { formulaObservations?: Record<string, { strikethrough?: unknown }> }
    for (const [ref, o] of Object.entries(sidecar.formulaObservations ?? {})) out.set(ref.toUpperCase(), o?.strikethrough === true)
  } catch { /* an unreadable sidecar has no strikethrough evidence */ }
  return out
}

export function stage5ObservedFills(c: CfOracleCase, rootDir: string): Map<string, string> | undefined {
  if (c.stage !== 5) return undefined
  const cap = c.captures.find(x => x.appearance === 'screen') ?? c.captures[0]
  if (!cap?.sidecarPath) return undefined
  const path = resolve(rootDir, cap.sidecarPath)
  if (!existsSync(path)) return undefined
  let sidecar: { formulaObservations?: Record<string, { displayFormatFill?: unknown; baseFill?: unknown }> }
  try {
    sidecar = JSON.parse(readFileSync(path, 'utf8')) as typeof sidecar
  } catch {
    return undefined
  }
  const out = new Map<string, string>()
  for (const [ref, o] of Object.entries(sidecar.formulaObservations ?? {})) {
    if (typeof o?.displayFormatFill !== 'string') continue
    const base = typeof o.baseFill === 'string' ? o.baseFill : 'FFFFFFFF'
    out.set(ref.toUpperCase(), o.displayFormatFill.toUpperCase() === base.toUpperCase() ? 'FFFFFFFF' : o.displayFormatFill.toUpperCase())
  }
  return out
}

/**
 * Stage 3 observed bars: the raw px (`observedBars`) the Excel capture recorded per designated cell, never fractions
 * (Ruling S3-A step 5), so conformance compares production layout px at the same scale and no expected value is ever
 * computed by the code under test. Only a Stage 3 case with a readable screen sidecar yields a map; entries with a
 * malformed edge record are dropped rather than guessed.
 */
export function stage3ObservedBars(c: CfOracleCase, rootDir: string): Map<string, CfObservedBar> | undefined {
  if (c.stage !== 3) return undefined
  const cap = c.captures.find(x => x.appearance === 'screen') ?? c.captures[0]
  if (!cap?.sidecarPath) return undefined
  const path = resolve(rootDir, cap.sidecarPath)
  if (!existsSync(path)) return undefined
  let sidecar: { observedBars?: unknown }
  try {
    sidecar = JSON.parse(readFileSync(path, 'utf8')) as { observedBars?: unknown }
  } catch {
    return undefined
  }
  const out = new Map<string, CfObservedBar>()
  const finite = (n: unknown): n is number => typeof n === 'number' && Number.isFinite(n)
  for (const raw of Array.isArray(sidecar.observedBars) ? sidecar.observedBars : []) {
    const b = raw as Partial<CfObservedBar> | null
    if (!b || typeof b.ref !== 'string' || !finite(b.cellLeftPx) || !finite(b.cellWidthPx)) continue
    const edges = b.bar
    if (edges !== null && !(edges && finite(edges.leftPx) && finite(edges.rightPx) && finite(edges.topPx) && finite(edges.bottomPx))) continue
    out.set(b.ref.toUpperCase(), b as CfObservedBar)
  }
  return out
}

/**
 * Stage 4 observed icons: the raw px and glyph class (`observedIcons`) the Excel capture recorded per designated cell
 * (Ruling S4-A), so conformance compares production's icon box and bucket at the same scale and no expected value is
 * computed by the code under test. Only a Stage 4 case with a readable screen sidecar yields a map; an entry with a
 * malformed box or no glyph class is dropped rather than guessed.
 */
export function stage4ObservedIcons(c: CfOracleCase, rootDir: string): Map<string, CfObservedIcon> | undefined {
  if (c.stage !== 4) return undefined
  const cap = c.captures.find(x => x.appearance === 'screen') ?? c.captures[0]
  if (!cap?.sidecarPath) return undefined
  const path = resolve(rootDir, cap.sidecarPath)
  if (!existsSync(path)) return undefined
  let sidecar: { observedIcons?: unknown }
  try {
    sidecar = JSON.parse(readFileSync(path, 'utf8')) as { observedIcons?: unknown }
  } catch {
    return undefined
  }
  const out = new Map<string, CfObservedIcon>()
  const finite = (n: unknown): n is number => typeof n === 'number' && Number.isFinite(n)
  for (const raw of Array.isArray(sidecar.observedIcons) ? sidecar.observedIcons : []) {
    const o = raw as Partial<CfObservedIcon> | null
    if (!o || typeof o.ref !== 'string' || !finite(o.cellLeftPx) || !finite(o.cellTopPx) || !finite(o.cellWidthPx) || !finite(o.cellHeightPx)) continue
    const i = o.icon
    if (i !== null && !(i && finite(i.leftPx) && finite(i.rightPx) && finite(i.topPx) && finite(i.bottomPx) && typeof i.glyphClass === 'string')) continue
    out.set(o.ref.toUpperCase(), { ...(o as CfObservedIcon), ref: o.ref.toUpperCase() })
  }
  return out
}

/** A1 reference of a 0-based cell, multi-letter columns included (fix round 5: AA was named `[`). */
const a1Of = (row: number, col: number): string => { let n = col + 1, l = ''; while (n > 0) { const k = (n - 1) % 26; l = String.fromCharCode(65 + k) + l; n = Math.floor((n - 1) / 26) } return `${l}${row + 1}` }

/** A leg's Stage 5 token, keeping its bar and icon descriptors, resolved properties and colour-scale colour (fix round 5).
 * A component the leg does not report stays undefined, which never agrees with the counterfactual. */
export function stage5TokenOfActual(actual: ActualCellOutcome): Stage5Production {
  const slots = [...(actual.unresolvedSlots ?? (actual.unresolved ? ['ALL'] : []))]
  const visual = 'icon' in actual && 'bar' in actual ? stage5VisualToken(actual.bar, actual.icon, slots) : undefined
  return { fill: actual.applied && actual.style?.fillColor ? String(actual.style.fillColor).toUpperCase() : 'none', strike: actual.applied && actual.style?.strike === true, unresolved: slots,
    ...(visual !== undefined ? { visual } : {}), ...(actual.properties ? { properties: actual.properties } : {}), ...('scaleArgb' in actual ? { colorScale: stage5ScaleToken(actual.scaleArgb) } : {}), ...(actual.visualDetail !== undefined ? { visualDetail: actual.visualDetail } : {}) }
}

/** The refs a Stage 5 case is checked at: its expected cells and Excel's observed ones. */
export function stage5CompareRefs(c: CfOracleCase, rootDir: string): string[] {
  return [...new Set([...c.expected.cells.map(e => e.ref), ...(c.status === 'measured' ? stage5ObservedFills(c, rootDir)?.keys() ?? [] : [])])]
}

export function verifyCaseCells(
  c: CfOracleCase,
  rootDir: string,
  actualAt: (row: number, col: number) => ActualCellOutcome,
  /** Fix round 4: required on a Stage 5 case, the counterfactual per ref (`stage5CounterfactualOf` over `stage5CompareRefs`). */
  stage5Truths?: ReadonlyMap<string, Stage5Truth>,
): string | undefined {
  if (c.stage === 5 && !stage5Truths) return `Stage 5 case ${c.id}: no counterfactual to verify against`
  const observed = stage2ObservedFills(c, rootDir) ?? (c.status === 'measured' ? stage5ObservedFills(c, rootDir) : undefined)
  // Fail closed (final review I4): a measured Stage 2 case is verified against Excel-observed fills only. A missing
  // sidecar, a sidecar with no trusted fills, or a run that compared zero non-withheld cells verified NOTHING and
  // must refuse, never fall through to seed expectations (which carry no fillRgb) or pass vacuously.
  const measuredStage2 = c.stage === 2 && c.status === 'measured'
  if (measuredStage2) {
    const cap = c.captures.find(x => x.appearance === 'screen') ?? c.captures[0]
    if (!observed) return `Stage 2 case ${c.id}: Excel observation sidecar${cap?.sidecarPath ? ` ${cap.sidecarPath}` : ''} is missing; nothing to verify against`
    if (observed.size === 0) return `Stage 2 case ${c.id}: sidecar has no trusted excel-display-format fills; nothing to verify against`
  }
  const icons = c.stage === 4 ? stage4ObservedIcons(c, rootDir) : undefined
  const iconObserved = icons ? new Map([...icons].map(([ref, o]) => [ref.toUpperCase(), observedIconClass(o)] as const)) : undefined
  if (!observed) return verifyCellExpectations(c, actualAt, iconObserved, stage5Truths)
  const contradicted = new Set(
    c.expected.cells.filter(e => e.matched && e.scaleApplied && observed.get(e.ref) === 'FFFFFFFF').map(e => e.ref),
  )
  const mismatch = verifyCellExpectations(
    { ...c, expected: { ...c.expected, cells: c.expected.cells.filter(e => !contradicted.has(e.ref)) } },
    actualAt,
    undefined,
    stage5Truths,
  )
  if (mismatch) return mismatch
  let compared = 0
  if (c.stage === 5) {
    // Fix round 4: every Stage 5 observed cell, never skipped as unresolved, must equal the counterfactual run of the real
    // compositor (the resolved components and the set of unresolved write paths), and each resolved component is Excel's.
    const strikes = stage5ObservedStrikes(c, rootDir)
    for (const [ref, want] of observed) {
      const pos = parseA1(ref)
      if (!pos) return `Invalid observed cell reference: ${ref}`
      const production = stage5TokenOfActual(actualAt(pos.row, pos.col))
      const excel = { fill: want === 'FFFFFFFF' ? 'none' : want, strike: strikes.get(ref) === true }
      const v = stage5CellAgreement(production, excel, stage5Truths!.get(ref))
      if (!v.agrees) return `Cell ${ref}: ${v.disagreement}`
      compared++
    }
    return compared === 0 ? `Stage 5 case ${c.id}: no observed cell was compared` : undefined
  }
  // Codex final review P2: a withheld Stage 2 cell is accepted only when it paints no fill and, where Excel shows a fill,
  // its withholding is justified: a typed finding (STAGE2_WITHHELD_FINDINGS) or the num-segment condition, recomputed from
  // the case's own marker and the saved cell value. Any other unresolved cell is a failure, never a silent skip.
  const values = c.stage === 2 ? stage2SavedNumbers(c, rootDir) : undefined
  for (const [ref, want] of observed) {
    const pos = parseA1(ref)
    if (!pos) return `Invalid observed cell reference: ${ref}`
    const actual = actualAt(pos.row, pos.col)
    if (actual.unresolved) {
      if (c.stage === 2) {
        if (actual.style?.fillColor) return `Cell ${ref} is reported withheld but paints fill ${String(actual.style.fillColor).toUpperCase()}`
        // Where Excel shows no fill the withheld cell renders the same base, so nothing needs justifying there.
        if (want !== 'FFFFFFFF' && STAGE2_WITHHELD_FINDINGS[`${c.id}:${ref}`] === undefined && !stage2NumSegmentShape(c.markers, ref, values?.get(ref))) {
          return `Cell ${ref} is withheld with no recorded justification (Excel fill ${want}): neither a typed Stage 2 finding nor the num-segment condition`
        }
      }
      continue
    }
    compared++
    const got = (actual.applied && actual.style?.fillColor ? String(actual.style.fillColor) : 'FFFFFFFF').toUpperCase()
    if (got !== want) return `Cell ${ref} observed Excel fill ${want}, got ${got}`
  }
  if (measuredStage2 && compared === 0) {
    return `Stage 2 case ${c.id}: every observed cell was withheld; compared zero cells, so nothing was verified`
  }
  return undefined
}

export async function compareInteractiveViewport(c: CfOracleCase, rootDir: string): Promise<CfCompareResult> {
  const oracleRef = `${c.id} [${c.status}]`
  const workbookHash = c.expected.savedSha256
  // Oracle content hash: binds the expected markers/cells (independent
  // fixture design intent), not just the case id/status.
  const oracleHash = sha256Hex([JSON.stringify({ expected: c.expected, markers: c.markers })])
  const fail = (reason: string): CfCompareResult => ({ id: c.id, mode: 'interactive', outcome: 'fail', reason, oracleRef, workbookHash, oracleHash, reads: 1 })
  const { prepareCfTextScope } = await import('../src/xlsx/cf/text')
  const { createTrackingBudget } = await import('./cf-refs/audit')
  const { createCfMemory, CF_LIMITS } = await import('../src/xlsx/cf/limits')
  const { createCfService } = await import('../src/xlsx/cf/service')
  const { cfRuntimeEvidence } = await import('../src/xlsx/cf/predicates')
  const { residentSource } = await import('../src/xlsx/cf/render')
  const { cfRulesVersion } = await import('../src/xlsx/view/pagedSource')
  const savedAbs = resolve(rootDir, c.saved)
  const pkg = await OfficePackage.load(new Uint8Array(readFileSync(savedAbs)))
  // Final review F2: a Stage 5 clock case is parsed and prepared at its anchor instant in its required zone (case-clock.ts).
  const doc = await parseXlsx(pkg, c.stage === 5 ? caseParseOptionsOf(c) : {})
  const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]
  if (!sheet) return fail(`Sheet ${c.sheet} missing`)
  const rect = parseRangeA1(c.range)
  if (!rect) return fail(`Invalid range ${c.range}`)
  // Service composition: full property VALUES per cell (red vs blue detected).
  const memory = createCfMemory()
  if (!sheet.conditionalFormatting) return fail(`No conditional formatting in ${c.saved}`)
  // Task 39 fix round 1 (F2): production's inputs. A Stage 5 case takes the document's evaluation context (its clock at
  // the case's anchor day for a clock case) and the workbook's other sheets, as render.ts passes them; Stages 1-4 keep
  // their Stage 1 clock injection.
  const stage5 = c.stage === 5
  const dateContext = c.clockRequirement && !stage5 ? {
    clockSerial: c.clockRequirement.anchorSerial,
    dateSystem: c.clockRequirement.dateSystem,
  } : undefined
  const { cfWorkbookOfDocument } = await import('../src/xlsx/cf/render')
  const { cfEvaluationContextFor } = await import('../src/xlsx/cf/formula')
  const anchorNow = caseNowOf(c)
  const documentEvaluation = caseDocumentEvaluationOf(c, doc)
  const oracle = createCfService(sheet.conditionalFormatting, memory, { evidence: cfRuntimeEvidence, dateContext,
    ...(stage5 ? { evaluation: cfEvaluationContextFor(documentEvaluation, anchorNow), sheetName: sheet.name, workbook: cfWorkbookOfDocument(doc.sheets, memory, 'engine') } : {}) })
  const gen = cfRulesVersion(sheet.conditionalFormatting)
  const stamp = { owner: 'compare', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: gen, viewGen: 0, structureGen: 0, clockGen: 0 }
  const source = residentSource(sheet, stamp, memory, 'engine')
  // Task 30: the range in 8-row batches, as a viewport asks for it. One batch over a whole case range of 76 icon rules
  // (stage4-icon-glyph-key A1:K42) exceeds one paint batch's grant and refuses atomically (cf-resource-limit).
  const BATCH_ROWS = 8
  const leases: Awaited<ReturnType<typeof oracle.prepareCf>>[] = []
  try {
    for (let r0 = rect.minRow; r0 <= rect.maxRow; r0 += BATCH_ROWS) {
      leases.push(await oracle.prepareCf(
        { stamp, sourceRects: [{ firstRow: r0, lastRow: Math.min(rect.maxRow, r0 + BATCH_ROWS - 1), firstCol: rect.minCol, lastCol: rect.maxCol }], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes },
        source,
        memory,
      ))
    }
  } catch (error) {
    // Task 39: a batch production refuses at its admitted grant withholds atomically (the viewer shows base formatting),
    // as the paged read does; it is recorded per case instead of aborting the whole comparison run.
    for (const l of leases) l.release()
    if ((error as { code?: string }).code !== 'resource-limit') throw error
    return { id: c.id, mode: 'interactive', outcome: 'recipe-pending', reason: `Interactive preparation refused at the admitted batch (${(error as Error).message}); withheld fail-closed, pixel diff awaits Excel viewport refs.`, oracleRef, workbookHash, oracleHash, reads: 1 }
  }
  const lease = {
    value: { at: (row: number, col: number) => leases[Math.floor((row - rect.minRow) / BATCH_ROWS)]!.value.at(row, col) },
    release: () => { for (const l of leases) l.release() },
  }
  const lines: string[] = []
  const outcomes = new Map<string, { kind: string; properties: Record<string, unknown>; writes: string[]; icon?: { set: string; iconId: number } | null; bar?: { startFraction: number; endFraction: number } | null; scaleArgb?: number | null; visualDetail?: string }>()
  try {
    for (let row = rect.minRow; row <= rect.maxRow; row++) {
      for (let col = rect.minCol; col <= rect.maxCol; col++) {
        const outcome = lease.value.at(row, col)
        if (outcome.kind === 'unresolved' && outcome.reason === 'cf-outside-preparation') continue
        if (outcome.kind === 'ready') {
          const props = outcome.value.properties as Record<string, unknown>
          const values = Object.keys(props).sort().map(k => `${k}=${JSON.stringify(props[k])}`).join('+')
          const writes = outcome.value.unresolvedWrites ? [...outcome.value.unresolvedWrites].sort().join('+') : ''
          const cs = outcome.value.visual?.colorScale
          lines.push(`${row},${col}:ready:${values}|${writes}${cs ? `|scale:${cs.ruleId}:${(cs.argb >>> 0).toString(16)}:${cs.fraction ?? ''}` : ''}`)
          const ic = outcome.value.visual?.iconSet
          const bar = outcome.value.visual?.dataBar
          outcomes.set(`${row},${col}`, { kind: 'ready', properties: props, writes: outcome.value.unresolvedWrites ? [...outcome.value.unresolvedWrites] : [], icon: ic ? { set: ic.set, iconId: ic.iconId } : null,
            bar: bar ? { startFraction: bar.startFraction, endFraction: bar.endFraction } : null, scaleArgb: cs ? cs.argb : null, visualDetail: stage5VisualDetail(bar, ic) })
        } else {
          lines.push(`${row},${col}:${outcome.kind}:${outcome.kind === 'unresolved' ? outcome.reason : ''}`)
          outcomes.set(`${row},${col}`, { kind: outcome.kind, properties: {}, writes: [] })
        }
      }
    }
    // Oracle content binding is the SHA-256 of expected+markers (oracleHash)
    // plus the data-path output hash; semantic equality against Excel's
    // rendered output requires captures (recipe-pending) and is gated by
    // stageComplete/blockedCapabilities in the audit, not by this adapter.
  } finally {
    lease.release()
  }
  oracle.dispose()
  if (memory.currentBytes !== 0) return fail(`service composition leaked ${memory.currentBytes} bytes`)
  // Resident presentation scope: descriptors per cell (viewer-facing output).
  const { budget, net, peak } = createTrackingBudget()
  const base = {}
  let scope: Awaited<ReturnType<typeof prepareCfTextScope>> | undefined
  try {
    scope = await prepareCfTextScope(sheet as never, {
      rects: [{ r0: rect.minRow, r1: rect.maxRow, c0: rect.minCol, c1: rect.maxCol }],
      signal: new AbortController().signal,
      locale: 'en-US',
      budget,
      dateContext,
      ...(stage5 ? { evaluation: documentEvaluation, now: anchorNow, documentSheets: doc.sheets } : {}),
    })
  } catch (error) {
    return fail(error instanceof Error ? error.message : String(error))
  }
  if (scope) {
    for (let row = rect.minRow; row <= rect.maxRow; row++) {
      for (let col = rect.minCol; col <= rect.maxCol; col++) {
        const desc = scope.describe(row, col)?.text ?? ''
        const applied = scope.effectiveStyle(row, col, base as never) !== base
        if (desc || applied) lines.push(`scope:${row},${col}:${desc}|${applied ? '1' : '0'}`)
      }
    }
  }
  const peakBytes = peak()
  const stage5Truths = stage5 ? await stage5CounterfactualOf({ c, doc, sheet, refs: stage5CompareRefs(c, rootDir), now: anchorNow }) : undefined
  const mismatch = verifyCaseCells(c, rootDir, (row, col) => {
    const outcome = outcomes.get(`${row},${col}`)
    const style = scope ? scope.effectiveStyle(row, col, base as never) : (base as never)
    const hasStyle = Boolean(style && style !== base && Object.keys(style).length > 0)
    const desc = scope ? scope.describe(row, col) : undefined
    const hasWinningDesc = desc !== undefined && desc.key !== 'cf-unavailable'
    // A data bar is no cell style (Ruling S3-I) and Excel's DisplayFormat records no fill for it, so on a bar cell only
    // the style counts as applied (Task 39 fix round 1, F2: E2/G2/Y2 zero-length bars at the column minimum).
    const barCell = c.markers.some(m => m.ruleType === 'dataBar' && isCellInSqref(a1Of(row, col), m.sqref))
    const applied = hasStyle || (hasWinningDesc && !(barCell && c.stage === 5))
    const unresolved = outcome === undefined
      || outcome.kind === 'unresolved'
      || (outcome.writes && outcome.writes.length > 0)
      || desc?.key === 'cf-unavailable'
    const iconCell = c.markers.some(m => m.ruleType === 'iconSet' && isCellInSqref(a1Of(row, col), m.sqref))
    const iconWithheld = outcome?.kind !== 'ready' && outcome?.kind !== 'no-match' || (outcome?.writes?.length ?? 0) > 0
    const unresolvedSlots = outcome === undefined || outcome.kind === 'unresolved' ? ['ALL'] : [...(outcome.writes ?? [])]
    // Fix round 5: a Stage 5 cell reports every component (the service outcome's descriptors, properties and scale colour).
    const stage5Parts = stage5 ? { icon: outcome?.icon ?? null, bar: outcome?.bar ?? null, properties: outcome?.properties ?? {}, scaleArgb: outcome?.scaleArgb ?? null, visualDetail: outcome?.visualDetail ?? stage5VisualDetail(null, null) } : {}
    return { applied, unresolved, unresolvedSlots, style: hasStyle ? style : undefined, ...(iconCell ? { iconCell, icon: outcome?.icon ?? null, iconWithheld } : {}), ...stage5Parts }
  }, stage5Truths)
  // Task 39 fix round 1 (F2): the Stage 5 accounting of the formula cells Excel formats.
  let stage5Stats: CfCompareResult['stage5']
  if (stage5) {
    stage5Stats = { formulaMatched: 0, painted: 0, justifiedWithheld: 0 }
    for (const exp of c.expected.cells) {
      if (!exp.matched || !c.markers.some(m => m.formula !== undefined && isCellInSqref(exp.ref, m.sqref))) continue
      const pos = parseA1(exp.ref)!
      stage5Stats.formulaMatched++
      const outcome = outcomes.get(`${pos.row},${pos.col}`)
      const style = scope ? scope.effectiveStyle(pos.row, pos.col, base as never) : (base as never)
      if (style && style !== base && Object.keys(style).length > 0 && (outcome?.writes.length ?? 0) === 0) stage5Stats.painted++
      // Withheld under the counterfactual's own fill refusal (a mismatch with it already failed above).
      else if (stage5Truths?.get(exp.ref)?.counterfactual.unresolved.some(u => stage5Conflicts(u, 'fill'))) stage5Stats.justifiedWithheld++
    }
  }
  scope?.dispose()
  if (net() !== 0) return fail(`resident viewport leaked ${net()} bytes`)
  if (scope && peakBytes <= 0) return fail('resident viewport observed no allocation')
  if (mismatch) return fail(mismatch)
  const isMeasured = c.status === 'measured'
  return {
    id: c.id, mode: 'interactive',
    outcome: isMeasured ? 'pass' : 'recipe-pending',
    reason: isMeasured
      ? `Interactive viewport verified: all ${c.expected.cells.length} cell expectations match.`
      : `Viewport cell expectations verified (${c.expected.cells.length} cells checked); pixel diff awaits Excel viewport refs.`,
    outputHash: sha256Hex(lines), workbookHash, oracleRef, oracleHash, reads: 1,
    ...(stage5Stats ? { stage5: stage5Stats } : {}),
  }
}

/**
 * Paged viewport adapter (worker route): reads the case range through a real
 * paged import session and hashes the wire states. Runs the real worker data
 * path and binds the output to workbook/oracle hashes. Pixel diff stays
 * pending without Excel viewport refs; behavior is covered by
 * tests/browser/xlsx-cf.spec.ts (including forced-paged legs).
 */
export async function comparePagedViewport(c: CfOracleCase, rootDir: string): Promise<CfCompareResult> {
  // Final review F2: the worker takes the viewer's ambient zone (production semantics), so a Stage 5 clock case runs as a
  // viewer in its required zone: the zone is set before the session opens (the paged policy is built once per workbook).
  return c.stage === 5 ? withCaseZone(c, () => comparePagedViewportIn(c, rootDir)) : comparePagedViewportIn(c, rootDir)
}
async function comparePagedViewportIn(c: CfOracleCase, rootDir: string): Promise<CfCompareResult> {
  const oracleRef = `${c.id} [${c.status}]`
  const workbookHash = c.expected.savedSha256
  const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
  const { createPagedWorkbookSource } = await import('../src/xlsx/view/pagedSource')
  const { createCfMemory } = await import('../src/xlsx/cf/limits')
  const rect = parseRangeA1(c.range)
  if (!rect) return { id: c.id, mode: 'paged', outcome: 'fail', reason: `Invalid range ${c.range}`, oracleRef, workbookHash }
  const oracleHash = sha256Hex([JSON.stringify({ expected: c.expected, markers: c.markers })])
  const saved = new Uint8Array(readFileSync(resolve(rootDir, c.saved)))
  const session = await openImportSession(
    { kind: 'blob', blob: new Blob([saved.buffer as ArrayBuffer]) },
    { docId: `compare-paged-${c.id}`, profile: 'D' },
  )
  try {
    await session.done
    const source = createPagedWorkbookSource(session)
    const sheetId = source.sheets[0]?.sheetId
    if (!sheetId) return { id: c.id, mode: 'paged', outcome: 'fail', reason: 'No sheets in paged source', oracleRef, workbookHash }
    const store = source.store(sheetId)
    // Warm pages first so the read evaluates states (not pending on import).
    await store.readRows?.(
      { r0: rect.minRow, r1: rect.maxRow, c0: rect.minCol, c1: rect.maxCol }, undefined,
      { signal: new AbortController().signal, stamp: store.stamp as never },
    ).catch(() => undefined)
    const memory = createCfMemory()
    // Task 39 fix round 1 (F2): a Stage 5 clock case reads the worker's own evaluation clock at the case's anchor day
    // (`now`), as production does; a Stage 1 dateContext would disagree with it and drop the evaluation context.
    const dateContext = c.clockRequirement && c.stage !== 5 ? {
      clockSerial: c.clockRequirement.anchorSerial,
      dateSystem: c.clockRequirement.dateSystem,
    } : undefined
    const anchorNow = c.clockRequirement && c.stage === 5 ? caseNowOf(c) : undefined
    const result = await store.readConditionalFormatting?.({
      rects: [{ r0: rect.minRow, r1: rect.maxRow, c0: rect.minCol, c1: rect.maxCol }],
      stamp: { owner: 'compare', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 },
      grantBytes: 2 * 1024 * 1024,
    }, memory, { signal: new AbortController().signal, dateContext, ...(anchorNow ? { now: anchorNow } : {}) })
    if (!result) return { id: c.id, mode: 'paged', outcome: 'fail', reason: 'Paged read returned nothing', oracleRef, workbookHash, reads: 1 }
    if (result.kind === 'pending') {
      return {
        id: c.id, mode: 'paged', outcome: 'recipe-pending',
        reason: `Paged read pending on import (tracked: ${(result.pendingSheetIds ?? []).join(',') || 'none'}); data path verified, pixel diff awaits Excel viewport refs.`,
        outputHash: sha256Hex(['pending']), workbookHash, oracleRef, oracleHash, reads: 1,
      }
    }
    if (result.kind !== 'ready') {
      return {
        id: c.id, mode: 'paged', outcome: 'recipe-pending',
        reason: `Paged read withheld (${result.kind}); data path verified fail-closed, pixel diff awaits Excel viewport refs.`,
        outputHash: sha256Hex([result.kind]), workbookHash, oracleRef, oracleHash, reads: 1,
      }
    }
    // Hash full property VALUES (not just names): a red→blue fill change
    // must change the hash. Values are primitives (colors, codes, flags).
    const lines = [...result.states]
      .sort((a, b) => a.row - b.row || a.col - b.col)
      .map(s => {
        const props = s.properties ?? {}
        const values = Object.keys(props).sort().map(k => `${k}=${JSON.stringify((props as Record<string, unknown>)[k])}`).join('+')
        const scale = s.scale ? `|scale:${s.scale.ruleId}:${(s.scale.argb >>> 0).toString(16)}:${s.scale.fraction ?? ''}` : ''
        const bar = s.bar ? `|bar:${JSON.stringify(s.bar)}` : ''
        return `${s.row},${s.col}:${values}|${[...(s.unresolvedWrites ?? [])].sort().join('+')}|${s.unresolvedReason ?? ''}${scale}${bar}`
      })
    const { resolveCfCellStyle } = await import('../src/xlsx/cf/dxf')
    const pkg = await OfficePackage.load(saved)
    const doc = await parseXlsx(pkg, c.stage === 5 ? caseParseOptionsOf(c) : {})
    const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]
    const colors = { theme: sheet?.drawingTheme?.palette ?? sheet?.tableStyleSource?.colors.theme, indexed: sheet?.tableStyleSource?.colors.indexed }
    const base = {}
    const stateMap = new Map<string, (typeof result.states)[number]>()
    for (const s of result.states) stateMap.set(`${s.row},${s.col}`, s)
    // Fix round 4: the counterfactual on the worker's clock (the anchor day for a clock case).
    const stage5Truths = c.stage === 5 && sheet ? await stage5CounterfactualOf({ c, doc, sheet, refs: stage5CompareRefs(c, rootDir), now: anchorNow ?? Date.now }) : undefined
    const mismatch = verifyCaseCells(c, rootDir, (row, col) => {
      const s = stateMap.get(`${row},${col}`)
      const iconCell = c.markers.some(m => m.ruleType === 'iconSet' && isCellInSqref(a1Of(row, col), m.sqref))
      // Fix round 5: a Stage 5 cell reports every component from the wire state (null/empty: none).
      const stage5Parts = c.stage === 5 ? {
        icon: s?.icon ? { set: CF_ICON_SET_NAMES[s.icon.set] ?? `#${s.icon.set}`, iconId: s.icon.iconId } : null,
        bar: s?.bar ? { startFraction: s.bar.startFraction, endFraction: s.bar.endFraction } : null,
        properties: s?.properties ?? {}, scaleArgb: s?.scale?.argb ?? null,
        visualDetail: stage5VisualDetail(s?.bar, s?.icon ? { ...s.icon, set: CF_ICON_SET_NAMES[s.icon.set] ?? `#${s.icon.set}` } : null),
      } : {}
      // A cell with no state is a no-match: for an icon cell, production claims no icon.
      if (!s) return iconCell ? { applied: false, iconCell, icon: null, unresolvedSlots: [], ...stage5Parts } : { applied: false, unresolved: true, unresolvedSlots: [], ...stage5Parts }
      const iconState = iconCell ? { iconCell, icon: s.icon ? { set: CF_ICON_SET_NAMES[s.icon.set] ?? `#${s.icon.set}`, iconId: s.icon.iconId } : null, iconWithheld: s.unresolvedReason !== undefined || (s.unresolvedWrites?.length ?? 0) > 0 } : {}
      const unresolved = s.unresolvedReason !== undefined || (s.unresolvedWrites !== undefined && s.unresolvedWrites.length > 0)
      const unresolvedSlots = s.unresolvedReason !== undefined ? ['ALL'] : [...(s.unresolvedWrites ?? [])]
      const hasProps = s.properties !== undefined && Object.keys(s.properties).length > 0
      const hasScale = s.scale?.argb !== undefined
      // A wire bar descriptor is an applied outcome (Task 23): the bar paints, though it writes no style property.
      const hasBar = s.bar !== undefined
      // Task 39 fix round 1 (F2): on a Stage 5 case Excel's expectation is its DisplayFormat fill, which a bar never writes.
      if (!hasProps && !hasScale) return { applied: hasBar && c.stage !== 5, unresolved, unresolvedSlots, ...iconState, ...stage5Parts }
      const style = resolveCfCellStyle(base as never, {
        properties: s.properties,
        ...(hasScale ? { visual: { colorScale: { ruleId: s.scale!.ruleId, argb: s.scale!.argb, ...(s.scale!.fraction !== undefined ? { fraction: s.scale!.fraction } : {}) } } } : {}),
      }, colors)
      return { applied: true, unresolved, unresolvedSlots, style, ...iconState, ...stage5Parts }
    }, stage5Truths)
    const outputHash = sha256Hex(lines)
    if (mismatch) {
      return { id: c.id, mode: 'paged', outcome: 'fail', reason: mismatch, outputHash, workbookHash, oracleRef, oracleHash, reads: 1 }
    }
    const isMeasured = c.status === 'measured'
    return {
      id: c.id, mode: 'paged',
      outcome: isMeasured ? 'pass' : 'recipe-pending',
      reason: isMeasured
        ? `Paged viewport verified: all ${c.expected.cells.length} cell expectations match.`
        : `Viewport cell expectations verified (${c.expected.cells.length} cells checked); pixel diff awaits Excel viewport refs.`,
      outputHash, workbookHash, oracleRef, oracleHash, reads: 1,
    }
  } finally {
    await session.release()
  }
}

export function resolvePixelTolerance(c: CfOracleCase, width: number, height: number): Tolerance {
  const base = c.tolerance ?? { meanAbs: 2.0 / 255, rowMax: 4.0 / 255, strong: 0.003 }
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
      // Pending + valid: viewport modes run real adapters (data path + hash
      // binding); static/print/export return pending (Paintable diff needs refs).
      if (mode === 'interactive') return compareInteractiveViewport(c, rootDir)
      if (mode === 'paged') return comparePagedViewport(c, rootDir)
      return { id: c.id, mode, outcome: 'recipe-pending', reason: check.reason }
    }
    return { id: c.id, mode, outcome: 'fail', reason: `${check.code}: ${check.reason}` }
  }

  // Verified: viewport modes still use adapters (no Excel viewport refs exist
  // for pixel diff); static/print/export fall through to pixel comparison.
  if (mode === 'interactive') return compareInteractiveViewport(c, rootDir)
  if (mode === 'paged') return comparePagedViewport(c, rootDir)

  const appearance = mode === 'print' ? 'printer' : 'screen'
  const cap = c.captures.find(x => x.appearance === appearance)
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

  const refFrame = await loadFrame(resolve(rootDir, cap.pngPath))
  let canvas: ReturnType<typeof createCanvas>
  try { canvas = await paintCfCapture(doc, doc.sheets.indexOf(sheet), c.range, mode, Number(sidecar.pxPerPt), refFrame.width, refFrame.height) }
  catch (error) { return { id: c.id, mode, outcome: 'fail', reason: error instanceof Error ? error.message : String(error) } }
  const oursCropped: Frame = { width: canvas.width, height: canvas.height, data: canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data as unknown as Uint8ClampedArray }

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
  const evidenceArg = args.includes('--evidence') ? args[args.indexOf('--evidence') + 1] : 'production'
  if (evidenceArg !== 'production' && evidenceArg !== 'audit') {
    console.error('Usage: --evidence production|audit')
    process.exit(2)
  }

  if (!['static', 'interactive', 'paged', 'print', 'export'].includes(modeArg)) {
    console.error(`Usage: bun scripts/cf-compare.ts [--stage N] [--mode static|interactive|paged|print|export] [--case id] [--out dir] [--evidence production|audit]`)
    process.exit(2)
  }

  if (!existsSync(CF_MANIFEST_PATH)) {
    console.error(`Missing manifest at ${CF_MANIFEST_PATH}`)
    process.exit(1)
  }

  const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
  if (evidenceArg === 'audit') {
    const { auditStage3Manifest } = await import('./cf-refs/audit')
    const evidence = await import('../src/xlsx/cf/evidence')
    const audit = auditStage3Manifest(manifest, { rootDir: REPO_ROOT })
    const verified = audit.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId)
    evidence.injectExactCfEvidence([...evidence.VERIFIED_CAPABILITIES, ...verified])
    evidence.injectCfBarLayout(audit.barLayout.layout)
    evidence.injectCfAutoBarColours(audit.autoBarColours)
    evidence.injectCfBarGradient(audit.barGradient)
    evidence.injectCfBarLengths(audit.barLengths)
    evidence.injectCfBarGradientBorder(audit.barGradientBorder)
    evidence.injectCfBarAutoBounds(audit.barAutoBounds)
    console.log(`Evidence: audit (${verified.length} Stage 3 capabilities verified by the harness, unpublished)`)
  }
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
