/**
 * Captures Excel for Mac 16.114 screen and printer PNGs for Stage 1 Conditional Formatting oracle cases.
 *
 *   bun scripts/cf-refs/capture.ts [--stage 1] [--case <id>]
 *
 * Reuses `capture` from `scripts/xlsx-refs/capture.ts` and `osa` / `openWorkbook` from `scripts/xlsx-refs/excel.ts`.
 * Requires an awake, unlocked macOS display (`CGSSessionScreenIsLocked=false`).
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { capture } from '../xlsx-refs/capture'
import { openWorkbook, osa, q, runScript, workbookNames } from '../xlsx-refs/excel'
import { closeOwnedWorkbook, isDisplayLocked } from './build'
import { CF_CASES, type CfManifest, type CfOracleCase, type Refusal } from './cases'
import { CF_MANIFEST_PATH, CF_REFS_DIR } from './generate'
import {
  type BarSettingsRow,
  type CfObservedBar,
  type CfObservedIcon,
  type IconCapabilityEvaluation,
  type IconGlyphKey,
  buildIconGlyphKey,
  iconGlyphFaintRecord,
  ICON_GEOMETRY_RULES_VERSION,
  measureCaseBarExtents,
  iconGlyphTwinsRecord,
  classifyIconCapture,
  evaluateIconCapabilities,
  loadIconGlyphKey,
  measureCaseIcons,
  type DecodedPng,
  type WorksheetRangeGeometry,
  computeWorksheetRangeGeometry,
  decodePngSync,
  evaluateFormulaCase,
  formulaObservationOf,
  stage5IconKey,
  type BarCapabilityEvaluation,
  evaluateBarCapabilities,
  isCellInSqref,
  measureCaseBars,
  observeCellsFromCapture,
  readZipPartsSync,
  verifyCfCase,
} from './verify'
import { barCapabilityKind } from './cases'
import { parseXml } from '../../src/core/xml'

export { capture }

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const sha256Hex = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex')

function probeLiveExcelClockState(
  c: CfOracleCase,
  savedAbs: string,
  preExisting: ReadonlySet<string>,
): { excelTodaySerial: number; staleTodayCacheRefreshed: boolean } | null {
  const req = c.clockRequirement
  if (!req) return null
  const name = openWorkbook(savedAbs)
  try {
    const out = runScript(buildClockProbeScript(c, name), 20_000)
    const num = Number(out.trim())
    const serial = Number.isFinite(num) ? Math.floor(num) : req.anchorSerial
    return {
      excelTodaySerial: serial,
      staleTodayCacheRefreshed: serial !== req.staleSavedTodaySerial && serial === req.anchorSerial,
    }
  } finally {
    closeOwnedWorkbook(name, preExisting)
  }
}

/**
 * Converts an Excel AppleScript color return value into an ARGB hex string (`FFFFC7CE`).
 *
 * Supported return shapes:
 * 1. AppleScript RGBColor record/string: `{65535, 51118, 52942}` or `"65535, 51118, 52942"` (16-bit 0..65535 channels)
 * 2. 8-bit RGB record/string: `{255, 199, 206}` or `"255, 199, 206"`
 * 3. BGR integer (classic COM/VBA `Interior.Color` Long): `0x00BBGGRR`
 * 4. Hex string: `#FFC7CE` or `FFFFC7CE`
 */
export function parseExcelColorToArgb(raw: unknown): string | null {
  if (raw === null || raw === undefined || raw === 'none' || raw === '') return null

  // 1. Array of numbers [r, g, b]
  if (Array.isArray(raw) && raw.length >= 3) {
    const nums = raw.slice(0, 3).map(Number)
    if (nums.every(n => Number.isFinite(n) && n >= 0)) {
      return channelsToArgb(nums[0]!, nums[1]!, nums[2]!)
    }
  }

  // 2. String representation
  if (typeof raw === 'string') {
    const trimmed = raw.trim()
    if (!trimmed || trimmed === 'none' || trimmed === 'missing value') return null

    // Comma-separated channels: "{65535, 51118, 52942}" or "65535, 51118, 52942"
    const matches = trimmed.match(/\d+/g)
    if (matches && matches.length >= 3) {
      const r = Number(matches[0])
      const g = Number(matches[1])
      const b = Number(matches[2])
      return channelsToArgb(r, g, b)
    }

    // Numeric decimal integer string (e.g. VBA Color Long "13551615" <= 16777215)
    if (/^\d+$/.test(trimmed)) {
      const num = Number(trimmed)
      if (num <= 16777215) return integerToArgb(num)
      return null
    }

    // Hex string (e.g. "#FFC7CE" or "FFFFC7CE")
    const hexMatch = /^#?([0-9a-f]{6}|[0-9a-f]{8})$/i.exec(trimmed)
    if (hexMatch) {
      const h = hexMatch[1]!
      return h.length === 6 ? `FF${h.toUpperCase()}` : h.toUpperCase()
    }
  }

  // 3. Number (integer Long from VBA/COM)
  if (typeof raw === 'number' && Number.isFinite(raw) && raw >= 0) {
    return integerToArgb(raw)
  }

  return null
}


function channelsToArgb(r: number, g: number, b: number): string {
  // Check if channels are 16-bit (AppleScript standard 0..65535)
  const is16Bit = r > 255 || g > 255 || b > 255
  const to8Bit = (val: number) => {
    if (!is16Bit) return Math.max(0, Math.min(255, Math.round(val)))
    return Math.max(0, Math.min(255, Math.round((val * 255) / 65535)))
  }
  const r8 = to8Bit(r).toString(16).padStart(2, '0').toUpperCase()
  const g8 = to8Bit(g).toString(16).padStart(2, '0').toUpperCase()
  const b8 = to8Bit(b).toString(16).padStart(2, '0').toUpperCase()
  return `FF${r8}${g8}${b8}`
}

function integerToArgb(val: number): string {
  // Classic VBA BGR: red is low byte, blue is high byte
  const r = (val & 0xff).toString(16).padStart(2, '0').toUpperCase()
  const g = ((val >> 8) & 0xff).toString(16).padStart(2, '0').toUpperCase()
  const b = ((val >> 16) & 0xff).toString(16).padStart(2, '0').toUpperCase()
  return `FF${r}${g}${b}`
}

/**
 * The per-cell probes are unrolled AppleScript, and osacompile overflows its tables ("Internal table overflow", -2707)
 * near 500 cells (stage3-bar-lengths-default has 510). Each probe therefore runs over chunks of at most this many cells,
 * the results concatenated in order. Every case up to 100 cells (all of them before Task 22b) runs as one chunk,
 * exactly as before.
 */
export const PROBE_CHUNK_CELLS = 100
export function probeRefChunks(refs: readonly string[], size = PROBE_CHUNK_CELLS): string[][] {
  const out: string[][] = []
  for (let i = 0; i < refs.length; i += size) out.push(refs.slice(i, i + size))
  return out.length > 0 ? out : [[]]
}

export function buildTextProbeScript(workbookName: string, sheetName: string, cellRefs: readonly string[]): string[] {
  return [
    `set wb to workbook "${q(workbookName)}"`,
    `set ws to worksheet "${q(sheetName)}" of wb`,
    `set outList to {}`,
    ...cellRefs.flatMap(r => [
      `set tVal to ""`,
      `try`,
      `  set tVal to (string value of range "${q(r)}" of ws as string)`,
      `end try`,
      `set end of outList to tVal`,
    ]),
    `set AppleScript's text item delimiters to "|||"`,
    `return outList as string`,
  ]
}

export function buildDecorationProbeScript(workbookName: string, sheetName: string, cellRefs: readonly string[]): string[] {
  return [
    `set wb to workbook "${q(workbookName)}"`,
    `set ws to worksheet "${q(sheetName)}" of wb`,
    `set outList to {}`,
    ...cellRefs.flatMap(r => [
      `set uVal to "missing value"`,
      `set sVal to "missing value"`,
      `try`,
      `  set uVal to (underline of font object of display format of range "${q(r)}" of ws as string)`,
      `end try`,
      `try`,
      `  set sVal to (strikethrough of font object of display format of range "${q(r)}" of ws as string)`,
      `end try`,
      `set end of outList to (uVal & ":::" & sVal)`,
    ]),
    `set AppleScript's text item delimiters to "|||"`,
    `return outList as string`,
  ]
}

export function buildColorProbeScript(workbookName: string, sheetName: string, cellRefs: readonly string[]): string[] {
  return [
    `set wb to workbook "${q(workbookName)}"`,
    `set ws to worksheet "${q(sheetName)}" of wb`,
    `set outList to {}`,
    ...cellRefs.flatMap(r => [
      `set AppleScript's text item delimiters to ","`,
      `set dfColor to "missing value"`,
      `try`,
      `  set dfColor to (color of interior object of display format of range "${q(r)}" of ws as string)`,
      `end try`,
      `set baseColor to "missing value"`,
      `try`,
      `  set baseColor to (color of interior object of range "${q(r)}" of ws as string)`,
      `end try`,
      `set end of outList to (dfColor & ":::" & baseColor)`,
    ]),
    `set AppleScript's text item delimiters to "|||"`,
    `return outList as string`,
  ]
}

export function buildColorScaleCriteriaProbeScript(workbookName: string, sheetName: string): string[] {
  return [
    `set wb to workbook "${q(workbookName)}"`,
    `set ws to worksheet "${q(sheetName)}" of wb`,
    `set outList to {}`,
    `try`,
    `  repeat with fc in (get format conditions of ws)`,
    `    try`,
    `      set csc to color scale criteria of fc`,
    `      repeat with crit in csc`,
    `        set cType to "missing value"`,
    `        try`,
    `          set cType to (color scale criterion type of crit as string)`,
    `        end try`,
    `        set cVal to ""`,
    `        try`,
    `          set cVal to (color scale criterion value of crit as string)`,
    `        end try`,
    `        if cType is not "missing value" then`,
    `          set end of outList to (cType & ":::" & cVal)`,
    `        end if`,
    `      end repeat`,
    `    end try`,
    `  end repeat`,
    `end try`,
    `set AppleScript's text item delimiters to "|||"`,
    `return outList as string`,
  ]
}

/**
 * Data bar object-model readback (Ruling S3-A, Excel.sdef `databar format condition` :3185-3204, `databar border`
 * :3181-3184, `negative bar format` :3488-3493). One `|||` row per bar format condition on the sheet, `:::` fields in
 * the order priority, axis, fill type, border type, min point percent, max point percent, direction, show value,
 * negative colour type. Every read is its own `try`, so one unreadable property never aborts the probe. Preparation
 * only: this script has not been run against Excel.
 */
export function buildDataBarSettingsProbeScript(workbookName: string, sheetName: string): string[] {
  const guarded = (variable: string, expr: string): string[] => [
    `    set ${variable} to "missing value"`,
    `    try`,
    `      set ${variable} to (${expr} as string)`,
    `    end try`,
  ]
  return [
    `set wb to workbook "${q(workbookName)}"`,
    `set ws to worksheet "${q(sheetName)}" of wb`,
    `set outList to {}`,
    `try`,
    `  repeat with fc in (get format conditions of ws)`,
    ...guarded('axisVal', 'databar axis position of fc'),
    ...guarded('fillVal', 'databar fill type of fc'),
    ...guarded('borderVal', 'databar border type of databar border of fc'),
    ...guarded('minVal', 'min point percent of fc'),
    ...guarded('maxVal', 'max point percent of fc'),
    ...guarded('dirVal', 'databar direction of fc'),
    ...guarded('showVal', 'format condition show value of fc'),
    ...guarded('negVal', 'negative bar color type of negative bar format of fc'),
    ...guarded('prioVal', 'format condition priority of fc'),
    `    if axisVal is not "missing value" then`,
    `      set end of outList to (prioVal & ":::" & axisVal & ":::" & fillVal & ":::" & borderVal & ":::" & minVal & ":::" & maxVal & ":::" & dirVal & ":::" & showVal & ":::" & negVal)`,
    `    end if`,
    `  end repeat`,
    `end try`,
    `set AppleScript's text item delimiters to "|||"`,
    `return outList as string`,
  ]
}

/**
 * Ruling S4-A one-shot probe (`stage4-icon-om-reachability`): is `icon set format condition` (Excel.sdef X315) reachable?
 * Reads format condition 1 of the range, its `properties`, `icon criteria`, `reverse icon set order` and `show icon
 * only`, each inside its own `try`, one `|||` row per read: `ok:::<detail>` or `error:::<detail>`. Record only: whatever it
 * returns lifts no capability, and an unexpected success is an owner question.
 */
export function buildIconOmReachabilityScript(workbookName: string, sheetName: string, rangeRef = 'B2:B4'): string[] {
  const guarded = (label: string, body: string[]): string[] => [
    'try',
    ...body.map(l => `  ${l}`),
    'on error errMsg number errNum',
    `  set end of outList to ("error:::${label} " & errNum & " " & errMsg)`,
    'end try',
  ]
  return [
    `set wb to workbook "${q(workbookName)}"`,
    `set ws to worksheet "${q(sheetName)}" of wb`,
    `set rng to range "${q(rangeRef)}" of ws`,
    'set outList to {}',
    ...guarded('format condition 1 of range', ['set fc to format condition 1 of rng', 'set end of outList to ("ok:::format condition 1 of range is " & ((class of fc) as string))']),
    ...guarded('properties', ['set props to properties of (format condition 1 of rng)', 'set end of outList to "ok:::properties readable"']),
    ...guarded('icon criteria', ['set crit to icon criteria of (format condition 1 of rng)', 'set end of outList to ("ok:::icon criteria " & ((count of crit) as string))']),
    ...guarded('reverse icon set order', ['set end of outList to ("ok:::reverse icon set order " & ((reverse icon set order of (format condition 1 of rng)) as string))']),
    ...guarded('show icon only', ['set end of outList to ("ok:::show icon only " & ((show icon only of (format condition 1 of rng)) as string))']),
    `set AppleScript's text item delimiters to "|||"`,
    'return outList as string',
  ]
}

/** Parses the reachability probe's answer into rows; anything not `ok:::`/`error:::` is dropped. */
export function parseIconOmReachability(out: string): Array<{ status: 'ok' | 'error'; detail: string }> {
  const rows: Array<{ status: 'ok' | 'error'; detail: string }> = []
  for (const line of out.split('|||')) {
    const m = /^(ok|error):::([\s\S]*)$/.exec(line.trim())
    if (m) rows.push({ status: m[1] as 'ok' | 'error', detail: m[2]!.trim() })
  }
  return rows
}

/**
 * Ruling S4-N copy-text observation (`icon:showvalue-copy-text`, record only): copies the range and reads `the clipboard
 * as text`. The AX tree is not observable from the sdef and is recorded as such.
 */
export function buildIconCopyTextProbeScript(workbookName: string, sheetName: string, rangeRef: string): string[] {
  return [
    `set wb to workbook "${q(workbookName)}"`,
    `set ws to worksheet "${q(sheetName)}" of wb`,
    'set copied to "missing value"',
    'try',
    `  copy range (range "${q(rangeRef)}" of ws)`,
    '  delay 0.5',
    '  set copied to (the clipboard as text)',
    'on error errMsg number errNum',
    '  set copied to ("error:::" & errNum & " " & errMsg)',
    'end try',
    'return copied',
  ]
}

/** Everything a Stage 4 icon capture adds to its sidecar: lock states, raw px, classes, the recomputed hypotheses. */
export function buildIconSidecarExtras(
  c: CfOracleCase,
  parts: Map<string, Uint8Array>,
  decoded: DecodedPng,
  geom: WorksheetRangeGeometry,
  opts: { lockBefore: 'locked' | 'unlocked'; lockAfter: 'locked' | 'unlocked'; rootDir: string; omReadback?: ReturnType<typeof parseIconOmReachability>; copyText?: string },
): { extras: Record<string, unknown>; observed: ReadonlyMap<string, CfObservedIcon> | undefined; evaluation?: Record<string, IconCapabilityEvaluation>; refusal?: string } {
  const base: Record<string, unknown> = {
    zoom: 100,
    dpr: 1,
    lockBefore: opts.lockBefore,
    lockAfter: opts.lockAfter,
    ...(opts.omReadback ? { iconOmReachability: opts.omReadback } : {}),
    ...(opts.copyText !== undefined ? { iconCopyText: opts.copyText, accessibilityTree: 'not observable from the sdef (Ruling S4-N)' } : {}),
  }
  let key: IconGlyphKey | undefined
  if (c.markers.some(m => m.capabilityId.startsWith('icon:glyph-'))) {
    const built = buildIconGlyphKey(c, parts, decoded, geom)
    if ('kind' in built) return { extras: base, observed: undefined, refusal: `${built.code}: ${built.reason}` }
    key = built
    base.glyphClasses = Object.fromEntries(built.classOfPair)
    base.glyphTwins = iconGlyphTwinsRecord(built)
    base.glyphFaint = iconGlyphFaintRecord(built)
  } else {
    const loaded = loadIconGlyphKey(opts.rootDir)
    key = loaded && !('kind' in loaded) ? loaded : undefined
  }
  const measured = measureCaseIcons(c, decoded, geom, key)
  if ('kind' in measured) return { extras: base, observed: undefined, refusal: `${(measured as Refusal).code}: ${(measured as Refusal).reason}` }
  const observed = measured as ReadonlyMap<string, CfObservedIcon>
  const evaluation = evaluateIconCapabilities(c, parts, observed, key, { barExtents: measureCaseBarExtents(c, decoded, geom) })
  const hypotheses = Object.fromEntries(Object.entries(evaluation).filter(([, e]) => e.hypotheses).map(([cap, e]) => [cap, e.hypotheses]))
  // Fix round 7: the geometry readings' per-cell predictions, recomputed and compared by the verifier like iconHypotheses.
  const geometry = Object.fromEntries(Object.entries(evaluation).filter(([, e]) => e.geometry).map(([cap, e]) => [cap, e.geometry]))
  // Fix round 10: the designated cells excluded from text detection (matched ink inside a measured bar), and why.
  const exclusions = Object.fromEntries(Object.entries(evaluation).filter(([, e]) => e.textExclusions).map(([cap, e]) => [cap, e.textExclusions]))
  return { extras: { ...base, observedIcons: [...observed.values()], iconHypotheses: hypotheses, ...(Object.keys(geometry).length > 0 ? { iconGeometryHypotheses: geometry, iconGeometryRulesVersion: ICON_GEOMETRY_RULES_VERSION } : {}), ...(Object.keys(exclusions).length > 0 ? { iconTextExclusions: exclusions } : {}) }, observed, evaluation }
}

/** Parses the `buildDataBarSettingsProbeScript` output into rows (an empty or unparseable answer yields no rows). */
export function parseDataBarSettingsReadback(out: string): BarSettingsRow[] {
  const rows: BarSettingsRow[] = []
  for (const line of out.split('|||')) {
    const f = line.split(':::').map(x => x.trim())
    if (f.length !== 9) continue
    const priority = Number(f[0])
    if (!Number.isFinite(priority)) continue
    rows.push({
      priority,
      axis: f[1]!,
      fillType: f[2]!,
      borderType: f[3]!,
      minPointPercent: f[4]!,
      maxPointPercent: f[5]!,
      direction: f[6]!,
      showValue: f[7]!,
      negativeColorType: f[8]!,
    })
  }
  return rows
}

/* ---------------- Stage 5: semantic DisplayFormat route (Ruling S5-A, round-1 #4) ---------------- */

/** One cell's DisplayFormat readback lines: value, DisplayFormat interior colour, base interior colour, DisplayFormat strike. */
function s5CellReadLines(r: string): string[] {
  return [
    `set AppleScript's text item delimiters to ","`,
    `set vVal to ""`,
    `try`,
    `  set vVal to ((value of range "${q(r)}" of ws) as string)`,
    `end try`,
    `set dfColor to "missing value"`,
    `try`,
    `  set dfColor to (color of interior object of display format of range "${q(r)}" of ws as string)`,
    `end try`,
    `set baseColor to "missing value"`,
    `try`,
    `  set baseColor to (color of interior object of range "${q(r)}" of ws as string)`,
    `end try`,
    `set sVal to "missing value"`,
    `try`,
    `  set sVal to (strikethrough of font object of display format of range "${q(r)}" of ws as string)`,
    `end try`,
    `set end of outList to (vVal & ":::" & dfColor & ":::" & baseColor & ":::" & sVal)`,
  ]
}

function s5ObservationScript(workbookName: string, sheetName: string, refs: readonly string[], calculateFirst: boolean): string[] {
  return [
    `set wb to workbook "${q(workbookName)}"`,
    `set ws to worksheet "${q(sheetName)}" of wb`,
    `set calcMode to "unknown"`,
    `try`,
    `  set calcMode to (calculation as string)`,
    `end try`,
    `set wbCount to ((count of workbooks) as string)`,
    ...(calculateFirst ? [`calculate`] : []),
    `set outList to {}`,
    ...refs.flatMap(s5CellReadLines),
    `set AppleScript's text item delimiters to "|||"`,
    `return calcMode & "###" & wbCount & "###" & (outList as string)`,
  ]
}

/**
 * Round-1 #4: the manual-calculation observation, run on the build's input copy right after it opens and before anything
 * else: the application calculation mode, the open workbook count, then each cell's value and DisplayFormat. It carries no
 * `calculate`, so Excel shows what it shows from the saved values.
 */
export function buildManualObservationScript(workbookName: string, sheetName: string, refs: readonly string[]): string[] {
  return s5ObservationScript(workbookName, sheetName, refs, false)
}

/** The separate, later observation: exactly one explicit `calculate`, then the same reads (recorded apart from the first). */
export function buildExplicitCalculationObservationScript(workbookName: string, sheetName: string, refs: readonly string[]): string[] {
  return s5ObservationScript(workbookName, sheetName, refs, true)
}

export interface S5CellRead { value: string; displayFormatFill?: string; baseFill?: string; strikethrough?: boolean }

/** Parses a manual / explicit-calculation / dynamic-refresh observation (refs in script order). */
export function parseS5Observation(out: string, refs: readonly string[]): { calculationMode: string; workbookCount: number; cells: Record<string, S5CellRead> } {
  const [mode = 'unknown', count = '0', rest = ''] = out.trim().split('###')
  const items = rest.split('|||')
  const cells: Record<string, S5CellRead> = {}
  refs.forEach((ref, i) => {
    const [value = '', df, base, strike] = (items[i] ?? '').split(':::')
    const dfArgb = parseExcelColorToArgb(df)
    const baseArgb = parseExcelColorToArgb(base)
    cells[ref] = { value, ...(dfArgb ? { displayFormatFill: dfArgb } : {}), ...(baseArgb ? { baseFill: baseArgb } : {}), ...(strike === 'true' || strike === 'false' ? { strikethrough: strike === 'true' } : {}) }
  })
  return { calculationMode: mode.trim(), workbookCount: Number(count), cells }
}

/** The clock probe (stale TODAY() cell after a calculate); a `pre-calculate` case skips the calculate (round-1 #4). */
export function buildClockProbeScript(c: CfOracleCase, workbookName: string): string[] {
  const req = c.clockRequirement
  if (!req) return []
  return [
    `set wb to workbook "${q(workbookName)}"`,
    `set ws to worksheet "${q(c.sheet)}" of wb`,
    ...(c.calcObservation === 'pre-calculate' ? [] : [`calculate`]),
    `set v to value of range "${q(req.staleCacheCellRef)}" of ws`,
    `return v as string`,
  ]
}

/**
 * The cells a dynamic refresh reads: every `post:` ref, every `post:` literal positive control (fix round 6) and their
 * same-content controls (the evaluator checks all of them).
 */
export function dynamicRefreshRefs(c: CfOracleCase): string[] {
  const out = new Set<string>()
  const controls = (c.formulaConfigurations ?? []).flatMap(cfg => Object.keys(cfg.positiveControls ?? {}).filter(r => r.startsWith('post:')).map(r => r.slice(5)))
  for (const ref of [...(c.formulaDynamicRefresh?.refs ?? []), ...controls]) {
    out.add(ref)
    const ctrl = c.expected.cells.find(e => e.ref === ref)?.controlRef
    if (ctrl) out.add(ctrl)
  }
  return [...out]
}

/** The live edit of a referenced sheet without saving, then the DisplayFormat reads of the refreshed cells and their controls. */
export function buildDynamicRefreshScript(workbookName: string, sheetName: string, spec: NonNullable<CfOracleCase['formulaDynamicRefresh']>, refs: readonly string[] = spec.refs): string[] {
  return [
    `set wb to workbook "${q(workbookName)}"`,
    `set ws to worksheet "${q(sheetName)}" of wb`,
    ...spec.edits.map(e => `set value of range "${q(e.ref)}" of worksheet "${q(spec.sheet)}" of wb to ${e.value}`),
    `set calcMode to "unknown"`,
    `set wbCount to ((count of workbooks) as string)`,
    `set outList to {}`,
    ...refs.flatMap(s5CellReadLines),
    `set AppleScript's text item delimiters to "|||"`,
    `return calcMode & "###" & wbCount & "###" & (outList as string)`,
  ]
}

/** The Stage 5 scripts the capture runs for a case (exported so the no-premature-calculate rule is tested on text). */
export function buildStage5CaptureScripts(c: CfOracleCase, workbookName: string): { clock: string[]; dynamicRefresh: string[][] } {
  return {
    clock: buildClockProbeScript(c, workbookName),
    dynamicRefresh: c.formulaDynamicRefresh ? [buildDynamicRefreshScript(workbookName, c.sheet, c.formulaDynamicRefresh, dynamicRefreshRefs(c))] : [],
  }
}

/**
 * `--verify-only` precondition per case: `skip` when no capture exists yet (nothing claimed, nothing on disk), `verify`
 * when every capture has its PNG and sidecar, and a loud error when a capture claims a `pngSha256` whose files are gone.
 */
export function verifyOnlyCaptureState(c: CfOracleCase, rootDir: string): 'skip' | 'verify' {
  let onDisk = 0
  for (const cap of c.captures) {
    const have = existsSync(resolve(rootDir, cap.pngPath)) && existsSync(resolve(rootDir, cap.sidecarPath))
    if (have) onDisk++
    else if (cap.pngSha256 !== undefined) throw new Error(`${c.id}: capture ${cap.pngPath} claims pngSha256 ${cap.pngSha256} but its PNG or sidecar is missing`)
  }
  return onDisk === 0 ? 'skip' : 'verify'
}

/**
 * Runs one `copy picture` capture of `savedAbs` and then closes the workbook the capture opened (the shared `capture()`
 * opens it by its file name and leaves it open). A workbook that was already open before the harness started is
 * never closed (`assertCanCloseWorkbook`); the close runs even when the capture throws.
 */
export function captureAndCloseOwned(
  savedAbs: string,
  preExisting: ReadonlySet<string>,
  run: () => void,
  close: (name: string, preExisting: ReadonlySet<string>) => void = closeOwnedWorkbook,
): void {
  const name = savedAbs.split('/').pop()!
  try {
    run()
  } finally {
    if (!preExisting.has(name)) close(name, preExisting)
  }
}

/**
 * The order the capture loop visits the manifest's cases in (indices into `cases`): manifest order, except that the
 * Stage 4 glyph key comes first. Every other Stage 4 case is classified with the key its sidecar restores
 * (`loadIconGlyphKey`), so `--verify-only` must re-derive the key sidecar before any dependent case (Ruling
 * S4-NEARTWIN); the dependents are then re-derived unconditionally from their stored PNGs.
 */
export function verifyOnlyCaseOrder(cases: readonly CfOracleCase[], stage: number): number[] {
  const idx = cases.map((_, i) => i)
  if (stage !== 4) return idx
  const isKey = (i: number): boolean => cases[i]!.id === 'stage4-icon-glyph-key'
  return [...idx.filter(isKey), ...idx.filter(i => !isKey(i))]
}

/** `--verify-only`: whether the case's kept sidecar records a different saved workbook than the one on disk. */
export function staleCaptureSidecar(c: CfOracleCase, rootDir: string, savedSha256: string): boolean {
  const cap = c.captures[0]
  if (!cap || !existsSync(resolve(rootDir, cap.sidecarPath))) return false
  try {
    const sc = JSON.parse(readFileSync(resolve(rootDir, cap.sidecarPath), 'utf8')) as Record<string, unknown>
    return typeof sc.savedSha256 === 'string' && sc.savedSha256 !== savedSha256
  } catch {
    return false
  }
}

export type LockGuardResult =
  | { kind: 'captured' }
  | { kind: 'refused'; code: 'display-locked'; phase: 'before' | 'after'; reason: string }

/**
 * Ruling S3-A step 1: the display lock is checked immediately before and after every capture, not only at process
 * start. A locked (or unknown) display refuses before capturing; a lock observed after capturing discards what the
 * capture wrote, and so does a capture that throws.
 */
export function captureUnderLockGuard(o: { isLocked: () => boolean; run: () => void; discard: () => void }): LockGuardResult {
  if (o.isLocked()) {
    return { kind: 'refused', code: 'display-locked', phase: 'before', reason: 'display locked (or unknown) before capture' }
  }
  try {
    o.run()
  } catch (error) {
    o.discard()
    throw error
  }
  if (o.isLocked()) {
    o.discard()
    return { kind: 'refused', code: 'display-locked', phase: 'after', reason: 'display locked (or unknown) after capture; the capture was discarded' }
  }
  return { kind: 'captured' }
}

/** Numeric cell values of `xl/worksheets/sheet1.xml` (cached numbers only; no text, boolean or error cells). */
export function readNumericCellValues(parts: Map<string, Uint8Array>): Map<string, number> {
  const out = new Map<string, number>()
  const raw = parts.get('xl/worksheets/sheet1.xml')
  if (!raw) return out
  const root = parseXml(new TextDecoder().decode(raw)) as Record<string, any>
  const ws = root.worksheet ?? root
  const rows = ([] as any[]).concat(ws?.sheetData?.row ?? [])
  for (const row of rows) {
    for (const cell of ([] as any[]).concat(row?.c ?? [])) {
      const attrs = cell?.['@attrs'] ?? {}
      const t = attrs.t
      if (t !== undefined && t !== 'n') continue
      const v = typeof cell?.v === 'object' ? cell.v?.['#text'] : cell?.v
      const n = Number(v)
      if (attrs.r && v !== undefined && v !== '' && Number.isFinite(n)) out.set(String(attrs.r).toUpperCase(), n)
    }
  }
  return out
}

/**
 * Which evidence the capture established for each designated bar cell: exactly the verifier's own recomputation
 * (`evaluateBarCapabilities`), written into `barEvidence.settings` for a settings capability and `barEvidence.geometry`
 * for a geometry one. A cell under several capabilities keeps a label only where they agree.
 */
export function classifyBarCapture(
  c: CfOracleCase,
  evaluation: Readonly<Record<string, BarCapabilityEvaluation>>,
): Map<string, NonNullable<CfOracleCase['expected']['cells'][number]['barEvidence']>> {
  type Ev = NonNullable<CfOracleCase['expected']['cells'][number]['barEvidence']>
  const votes = new Map<string, { settings: Set<string>; geometry: Set<string> }>()
  for (const cap of new Set(c.markers.map(m => m.capabilityId))) {
    const kind = barCapabilityKind(cap)
    if (kind !== 'settings' && kind !== 'geometry') continue
    const rec = evaluation[cap]
    const label = rec?.status === 'discriminating' ? 'pixel-measured' : rec?.status === 'non-discriminating' ? 'non-discriminating' : 'unverified'
    for (const cell of c.expected.cells.filter(x => x.matched && c.markers.some(m => m.capabilityId === cap && isCellInSqref(x.ref, m.sqref)))) {
      const v = votes.get(cell.ref) ?? { settings: new Set<string>(), geometry: new Set<string>() }
      v[kind].add(label)
      votes.set(cell.ref, v)
    }
  }
  const labels = new Map<string, Ev>()
  for (const [ref, v] of votes) {
    const pick = (set: Set<string>): string => (set.size === 1 ? [...set][0]! : 'unverified')
    labels.set(ref, { settings: pick(v.settings) as Ev['settings'], geometry: pick(v.geometry) as Ev['geometry'] })
  }
  return labels
}

/** Everything a Stage 3 bar capture adds to its sidecar (lock states, raw px, the recomputed hypotheses). */
export function buildBarSidecarExtras(
  c: CfOracleCase,
  parts: Map<string, Uint8Array>,
  decoded: DecodedPng,
  geom: WorksheetRangeGeometry,
  opts: { lockBefore: 'locked' | 'unlocked'; lockAfter: 'locked' | 'unlocked'; readback?: BarSettingsRow[] },
): { extras: Record<string, unknown>; observed: ReadonlyMap<string, CfObservedBar> | undefined; evaluation?: Record<string, BarCapabilityEvaluation>; refusal?: string } {
  const measured = measureCaseBars(c, decoded, geom, parts)
  const base = {
    zoom: 100,
    dpr: 1,
    lockBefore: opts.lockBefore,
    lockAfter: opts.lockAfter,
    ...(opts.readback ? { barSettingsReadback: opts.readback } : {}),
  }
  if ('kind' in measured) return { extras: base, observed: undefined, refusal: `${measured.code}: ${measured.reason}` }
  const evaluation = evaluateBarCapabilities(c, parts, decoded, geom, measured)
  const hypotheses = Object.fromEntries(Object.entries(evaluation).filter(([, e]) => e.hypotheses).map(([cap, e]) => [cap, e.hypotheses]))
  return {
    extras: { ...base, observedBars: [...measured.values()], barHypotheses: hypotheses },
    observed: measured,
    evaluation,
  }
}

export function probeLiveExcelCellObservations(
  c: CfOracleCase,
  savedAbs: string,
  preExisting: ReadonlySet<string>,
): {
  cellTexts: Record<string, string>
  cellDecorations: Record<string, { underline: 'none' | 'single' | 'double'; strikethrough: boolean }>
  cellColors: Record<string, { displayFormatFill?: string; baseFill?: string }>
  geometry: { widthPt: number; heightPt: number; colWidthsPt: number[]; rowHeightsPt: number[] } | null
  colorScaleCriteria?: Record<string, { type: string; value?: string }[]>
  barSettingsReadback?: BarSettingsRow[]
} | null {
  try {
    const name = openWorkbook(savedAbs)
    try {
      const refs = c.expected.cells.map(x => x.ref)
      // Chunked (PROBE_CHUNK_CELLS): each chunk's results, split, then joined in cell order.
      const chunked = (build: (refs: readonly string[]) => string[]): Array<string | undefined> =>
        probeRefChunks(refs).flatMap(chunk => {
          const parts = runScript(build(chunk), 30_000).split('|||')
          return chunk.map((_, i) => parts[i])
        })
      const texts = chunked(chunk => buildTextProbeScript(name, c.sheet, chunk))
      const cellTexts: Record<string, string> = {}
      refs.forEach((ref, idx) => {
        if (texts[idx] !== undefined) cellTexts[ref] = texts[idx]
      })

      const decs = chunked(chunk => buildDecorationProbeScript(name, c.sheet, chunk))
      const cellDecorations: Record<string, { underline: 'none' | 'single' | 'double'; strikethrough: boolean }> = {}
      refs.forEach((ref, idx) => {
        const parts = decs[idx]?.split(':::')
        if (parts && parts.length === 2) {
          const uStr = parts[0]
          const stkStr = parts[1]
          let underline: 'none' | 'single' | 'double' = 'none'
          if (uStr.includes('double')) underline = 'double'
          else if (uStr.includes('single')) underline = 'single'
          const strikethrough = stkStr === 'true'
          cellDecorations[ref] = { underline, strikethrough }
        }
      })

      const cellColors: Record<
        string,
        {
          displayFormatFill?: string
          baseFill?: string
          rawDisplayFormatFill?: string
          rawBaseFill?: string
        }
      > = {}
      try {
        const colorParts = chunked(chunk => buildColorProbeScript(name, c.sheet, chunk))
        refs.forEach((ref, idx) => {
          const parts = colorParts[idx]?.split(':::')
          if (parts && parts.length === 2) {
            const dfArgb = parseExcelColorToArgb(parts[0])
            const baseArgb = parseExcelColorToArgb(parts[1])
            cellColors[ref] = {
              ...(dfArgb ? { displayFormatFill: dfArgb } : {}),
              ...(baseArgb ? { baseFill: baseArgb } : {}),
              rawDisplayFormatFill: parts[0]?.trim(),
              rawBaseFill: parts[1]?.trim(),
            }
          }
        })
      } catch {
        // color query errors leave cellColors empty
      }

      let colorScaleCriteria: Record<string, { type: string; value?: string }[]> | undefined
      try {
        const critScriptLines = buildColorScaleCriteriaProbeScript(name, c.sheet)
        const critOut = runScript(critScriptLines, 30_000)
        if (critOut.trim()) {
          const items = critOut.split('|||')
          colorScaleCriteria = {
            all: items.map(it => {
              const [type, value] = it.split(':::')
              return { type: type ?? '', ...(value ? { value } : {}) }
            }),
          }
        }
      } catch {
        // criteria query errors leave colorScaleCriteria undefined
      }

      let barSettingsReadback: BarSettingsRow[] | undefined
      if (c.stage === 3 && c.markers.some(m => m.ruleType === 'dataBar')) {
        try {
          barSettingsReadback = parseDataBarSettingsReadback(runScript(buildDataBarSettingsProbeScript(name, c.sheet), 30_000))
        } catch {
          // an unreadable object model leaves the readback undefined: settings stay unverified
        }
      }

      let geometry: { widthPt: number; heightPt: number; colWidthsPt: number[]; rowHeightsPt: number[] } | null = null
      try {
        const geomScriptLines = [
          `set wb to workbook "${q(name)}"`,
          `set ws to worksheet "${q(c.sheet)}" of wb`,
          `set r to range "${q(c.range)}" of ws`,
          `set totalW to width of r`,
          `set totalH to height of r`,
          `set AppleScript's text item delimiters to ","`,
          `set colWidths to {}`,
          `repeat with colVar in (get columns of r)`,
          `  set end of colWidths to (get width of colVar)`,
          `end repeat`,
          `set rowHeights to {}`,
          `repeat with rowVar in (get rows of r)`,
          `  set end of rowHeights to (get height of rowVar)`,
          `end repeat`,
          `return (totalW as string) & ":::" & (totalH as string) & ":::" & (colWidths as string) & ":::" & (rowHeights as string)`,
        ]
        const geomOut = runScript(geomScriptLines, 30_000)
        const gParts = geomOut.split(':::')
        if (gParts.length === 4) {
          const totalW = Number(gParts[0])
          const totalH = Number(gParts[1])
          const colWidthsPt = gParts[2].split(',').map(s => Number(s.trim()))
          const rowHeightsPt = gParts[3].split(',').map(s => Number(s.trim()))
          if (
            Number.isFinite(totalW) &&
            Number.isFinite(totalH) &&
            colWidthsPt.length > 0 &&
            colWidthsPt.every(Number.isFinite) &&
            rowHeightsPt.length > 0 &&
            rowHeightsPt.every(Number.isFinite)
          ) {
            const sumW = colWidthsPt.reduce((a, b) => a + b, 0)
            const sumH = rowHeightsPt.reduce((a, b) => a + b, 0)
            if (Math.abs(sumW - totalW) <= 1 && Math.abs(sumH - totalH) <= 1) {
              geometry = { widthPt: totalW, heightPt: totalH, colWidthsPt, rowHeightsPt }
            }
          }
        }
      } catch {
        // geometry query errors leave geometry null
      }

      return { cellTexts, cellDecorations, cellColors, geometry, colorScaleCriteria, ...(barSettingsReadback ? { barSettingsReadback } : {}) }
    } finally {
      closeOwnedWorkbook(name, preExisting)
    }
  } catch {
    return null
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2)
  const stage = Number(args.includes('--stage') ? args[args.indexOf('--stage') + 1] : 1)
  const only = args.includes('--case') ? args[args.indexOf('--case') + 1] : undefined
  const verifyOnly = args.includes('--verify-only')

  if (!verifyOnly && isDisplayLocked()) {
    console.error(
      '[cf-refs/capture] Refusing capture: macOS display is currently locked (CGSSessionScreenIsLocked=true). Run on an awake, unlocked display.',
    )
    process.exit(1)
  }

  mkdirSync(CF_REFS_DIR, { recursive: true })
  const excelVersion = verifyOnly ? '16.114' : osa(['tell application "Microsoft Excel" to get version'])
  const preExisting = verifyOnly ? new Set<string>() : new Set<string>(workbookNames() ?? [])
  const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest

  let capturedCount = 0
  for (const i of verifyOnlyCaseOrder(manifest.cases, stage)) {
    const c = manifest.cases[i]
    if (c.stage !== stage || (only && c.id !== only) || c.status === 'unsupported') continue
    if (verifyOnly && verifyOnlyCaptureState(c, REPO_ROOT) === 'skip') {
      console.log(`[cf-refs/capture] ${c.id}: no capture yet, skipped (stays ${c.status})`)
      continue
    }
    const savedAbs = resolve(REPO_ROOT, c.saved)
    const savedBytes = new Uint8Array(readFileSync(savedAbs))
    const savedSha256 = sha256Hex(savedBytes)
    if (verifyOnly && staleCaptureSidecar(c, REPO_ROOT, savedSha256)) {
      // The kept PNG shows an older saved workbook (the case was regenerated since): re-verifying would bind that
      // picture to the new workbook. Leave sidecar and manifest entry alone; the case needs build + capture.
      console.warn(`[cf-refs/capture] ${c.id}: the kept capture predates the current saved workbook; skipped (rebuild and recapture it)`)
      continue
    }
    const parts = readZipPartsSync(savedBytes)
    const geom = parts ? computeWorksheetRangeGeometry(parts, c.range) : null

    let hasLiveGeom = false
    let activeGeom = geom
    let cellTexts: Record<string, string> | undefined
    let cellDecorations: Record<string, { underline?: 'none' | 'single' | 'double'; strikethrough?: boolean }> | undefined
    let cellColors:
      | Record<
          string,
          {
            displayFormatFill?: string
            baseFill?: string
            rawDisplayFormatFill?: string
            rawBaseFill?: string
          }
        >
      | undefined
    let colorScaleCriteria: Record<string, { type: string; value?: string }[]> | undefined
    let barReadback: BarSettingsRow[] | undefined
    const isBarCase = c.stage === 3 && c.markers.some(m => m.ruleType === 'dataBar')
    // Stage 4: icon cases take the icon branch only (Ruling S4-A); never the Stage 1 contrast observation.
    const isIconCase = c.stage === 4
    // Stage 5: formula cases take the semantic DisplayFormat route only (Ruling S5-A); never the Stage 1 contrast observation.
    const isFormulaCase = c.stage === 5
    let formulaObservations: Record<string, Omit<S5CellRead, 'value'>> | undefined
    let formulaDynamic: Record<string, S5CellRead> | undefined
    let iconObserved: ReadonlyMap<string, CfObservedIcon> | undefined
    let iconEvaluation: Record<string, IconCapabilityEvaluation> | undefined
    let iconRefusal: string | undefined
    let iconOm: ReturnType<typeof parseIconOmReachability> | undefined
    let iconCopyText: string | undefined
    let sidecarLockBefore: 'locked' | 'unlocked' | undefined
    let sidecarLockAfter: 'locked' | 'unlocked' | undefined
    let barObserved: ReadonlyMap<string, CfObservedBar> | undefined
    let barEvaluation: Record<string, BarCapabilityEvaluation> | undefined
    let lockRefusal: string | undefined
    let barRefusal: string | undefined
    let sidecarExcelVersion = excelVersion
    let sidecarCapturedAt = new Date().toISOString()
    let sidecarHostTimeZone = 'Asia/Taipei'
    let clockState: { excelTodaySerial: number; staleTodayCacheRefreshed: boolean } | null = null
    let observedTodayBranchFromSidecar: 'local' | 'utc' | undefined

    if (verifyOnly) {
      const firstCap = c.captures[0]
      if (firstCap && existsSync(resolve(REPO_ROOT, firstCap.sidecarPath))) {
        try {
          const sc = JSON.parse(readFileSync(resolve(REPO_ROOT, firstCap.sidecarPath), 'utf8')) as Record<string, unknown>
          sidecarExcelVersion = String(sc.excelVersion ?? excelVersion)
          sidecarCapturedAt = String(sc.capturedAt ?? sidecarCapturedAt)
          sidecarHostTimeZone = String(sc.hostTimeZone ?? sidecarHostTimeZone)
          observedTodayBranchFromSidecar = sc.observedTodayBranch as 'local' | 'utc' | undefined
          if (Array.isArray(sc.colWidthsPt) && Array.isArray(sc.rowHeightsPt) && parts) {
            hasLiveGeom = sc.geometrySource === 'excel-measured'
            activeGeom =
              computeWorksheetRangeGeometry(parts, c.range, {
                colWidthsPt: sc.colWidthsPt as number[],
                rowHeightsPt: sc.rowHeightsPt as number[],
              }) ?? geom
          }
          if (Array.isArray(sc.barSettingsReadback)) barReadback = sc.barSettingsReadback as BarSettingsRow[]
          if (Array.isArray(sc.iconOmReachability)) iconOm = sc.iconOmReachability as ReturnType<typeof parseIconOmReachability>
          if (typeof sc.iconCopyText === 'string') iconCopyText = sc.iconCopyText
          if (sc.formulaObservations && typeof sc.formulaObservations === 'object') formulaObservations = sc.formulaObservations as Record<string, Omit<S5CellRead, 'value'>>
          if (sc.formulaDynamicRefresh && typeof sc.formulaDynamicRefresh === 'object') formulaDynamic = sc.formulaDynamicRefresh as Record<string, S5CellRead>
          if (sc.lockBefore === 'unlocked' || sc.lockBefore === 'locked') sidecarLockBefore = sc.lockBefore
          if (sc.lockAfter === 'unlocked' || sc.lockAfter === 'locked') sidecarLockAfter = sc.lockAfter
          if (typeof sc.excelTodaySerial === 'number') {
            clockState = {
              excelTodaySerial: sc.excelTodaySerial,
              staleTodayCacheRefreshed: sc.staleTodayCacheRefreshed === true,
            }
          }
          if (Array.isArray(sc.observedCells)) {
            cellTexts = Object.fromEntries(
              sc.observedCells
                .filter((o: any) => typeof o.formattedText === 'string')
                .map((o: any) => [o.ref, o.formattedText]),
            )
            cellDecorations = Object.fromEntries(
              sc.observedCells
                .filter((o: any) => o.decorationSource === 'excel-display-format')
                .map((o: any) => [o.ref, { underline: o.underline, strikethrough: o.strikethrough }]),
            )
            cellColors = Object.fromEntries(
              sc.observedCells
                .filter((o: any) => o.fillSource === 'excel-display-format' && typeof o.fillRgb === 'string')
                .map((o: any) => [o.ref, { displayFormatFill: o.fillRgb }]),
            )
          }
        } catch {
          // ignore parse errors
        }
      }
    } else {
      clockState = c.clockRequirement ? probeLiveExcelClockState(c, savedAbs, preExisting) : null
      const liveObs = probeLiveExcelCellObservations(c, savedAbs, preExisting)
      const liveGeom = liveObs?.geometry
      hasLiveGeom =
        liveGeom !== null &&
        liveGeom !== undefined &&
        liveGeom.colWidthsPt.length > 0 &&
        liveGeom.rowHeightsPt.length > 0
      activeGeom =
        hasLiveGeom && parts && liveGeom
          ? computeWorksheetRangeGeometry(parts, c.range, {
              colWidthsPt: liveGeom.colWidthsPt,
              rowHeightsPt: liveGeom.rowHeightsPt,
            }) ?? geom
          : geom
      cellTexts = liveObs?.cellTexts
      cellDecorations = liveObs?.cellDecorations
      cellColors = liveObs?.cellColors
      colorScaleCriteria = liveObs?.colorScaleCriteria
      barReadback = liveObs?.barSettingsReadback
      if (isFormulaCase) {
        // Ruling S5-A: Excel's own DisplayFormat per designated cell, twin and control (interior colour and strikethrough).
        formulaObservations = Object.fromEntries(c.expected.cells.map(e => {
          const col = cellColors?.[e.ref]
          const dec = cellDecorations?.[e.ref]
          return [e.ref, { ...(col?.displayFormatFill ? { displayFormatFill: col.displayFormatFill } : {}), ...(col?.baseFill ? { baseFill: col.baseFill } : {}), ...(dec?.strikethrough !== undefined ? { strikethrough: dec.strikethrough } : {}) }]
        }))
        if (c.formulaDynamicRefresh) {
          // A live edit of the referenced sheet, never saved (the close discards it), in its own owned open/close.
          const name = openWorkbook(savedAbs)
          try {
            const refs = dynamicRefreshRefs(c)
            formulaDynamic = parseS5Observation(runScript(buildDynamicRefreshScript(name, c.sheet, c.formulaDynamicRefresh, refs), 30_000), refs).cells
          } finally {
            if (!preExisting.has(name)) closeOwnedWorkbook(name, preExisting)
          }
        }
      }
      // Record-only Excel observations of two Stage 4 cases (Rulings S4-A and S4-N), each in its own owned open/close.
      if (c.id === 'stage4-icon-om-reachability' || c.id === 'stage4-icon-showvalue') {
        const name = openWorkbook(savedAbs)
        try {
          if (c.id === 'stage4-icon-om-reachability') iconOm = parseIconOmReachability(runScript(buildIconOmReachabilityScript(name, c.sheet), 30_000))
          else iconCopyText = runScript(buildIconCopyTextProbeScript(name, c.sheet, 'D2:D4'), 30_000)
        } finally {
          if (!preExisting.has(name)) closeOwnedWorkbook(name, preExisting)
        }
      }
    }

    const updatedTemplate = CF_CASES.find(tc => tc.id === c.id)
    const currentExpected = {
      ...c.expected,
      ...(updatedTemplate
        ? {
            cells: updatedTemplate.expected.cells,
            positiveRefs: updatedTemplate.expected.positiveRefs,
            negativeRefs: updatedTemplate.expected.negativeRefs,
          }
        : {}),
      savedSha256,
    }
    const cWithUpdatedExpected: CfOracleCase = {
      ...c,
      expected: currentExpected,
    }

    const updatedCaptures = c.captures.map(cap => {
      const pngAbs = resolve(REPO_ROOT, cap.pngPath)
      const sidecarAbs = resolve(REPO_ROOT, cap.sidecarPath)
      if (!verifyOnly) {
        // Ruling S3-A step 1: the lock is re-checked immediately before and after every capture; a capture taken
        // across a lock is discarded and the case stays recipe-pending.
        const guard = captureUnderLockGuard({
          isLocked: () => isDisplayLocked(),
          run: () => captureAndCloseOwned(savedAbs, preExisting, () => capture(savedAbs, c.sheet, c.range, pngAbs, { appearance: cap.appearance })),
          discard: () => {
            rmSync(pngAbs, { force: true })
            rmSync(sidecarAbs, { force: true })
          },
        })
        if (guard.kind === 'refused') {
          lockRefusal = `display-locked (${guard.phase}): ${guard.reason}`
          return cap
        }
        sidecarLockBefore = 'unlocked'
        sidecarLockAfter = 'unlocked'
      }
      if (lockRefusal) return cap
      const pngBytes = new Uint8Array(readFileSync(pngAbs))
      const pngSha256 = sha256Hex(pngBytes)
      const decoded = decodePngSync(pngBytes)

      let capGeom = activeGeom
      if (cap.appearance === 'printer' && decoded && capGeom && decoded.height < capGeom.expectedHeightPt) {
        const printerRowHeights = capGeom.rowHeightsPt.map(h => (h === 18 ? 16 : Math.round((h * 16) / 18)))
        capGeom =
          computeWorksheetRangeGeometry(parts!, c.range, {
            colWidthsPt: capGeom.colWidthsPt,
            rowHeightsPt: printerRowHeights,
          }) ?? capGeom
      }

      const barExtra =
        isBarCase && decoded && capGeom && sidecarLockBefore && sidecarLockAfter
          ? buildBarSidecarExtras(cWithUpdatedExpected, parts!, decoded, capGeom, { lockBefore: sidecarLockBefore, lockAfter: sidecarLockAfter, ...(barReadback ? { readback: barReadback } : {}) })
          : undefined
      const iconExtra =
        isIconCase && decoded && capGeom && sidecarLockBefore && sidecarLockAfter
          ? buildIconSidecarExtras(cWithUpdatedExpected, parts!, decoded, capGeom, { lockBefore: sidecarLockBefore, lockAfter: sidecarLockAfter, rootDir: REPO_ROOT, ...(iconOm ? { omReadback: iconOm } : {}), ...(iconCopyText !== undefined ? { copyText: iconCopyText } : {}) })
          : undefined
      if (iconExtra) {
        iconObserved = iconExtra.observed
        iconEvaluation = iconExtra.evaluation
        if (iconExtra.refusal) iconRefusal = iconExtra.refusal
      }
      if (barExtra) {
        barObserved = barExtra.observed
        barEvaluation = barExtra.evaluation
        if (barExtra.refusal) barRefusal = barExtra.refusal
      }

      const observedCells =
        isBarCase || isIconCase || isFormulaCase
          ? null
          : decoded && capGeom
          ? observeCellsFromCapture(decoded, capGeom, cWithUpdatedExpected, {
              observedCellTexts: cellTexts,
              observedCellDecorations: cellDecorations,
            })
          : null

      const finalObservedCells = observedCells?.map(obs => {
        const col = cellColors?.[obs.ref]
        return {
          ...obs,
          ...(col?.displayFormatFill ? { fillSource: 'excel-display-format', fillRgb: col.displayFormatFill } : {}),
          ...(col?.baseFill ? { baseFillRgb: col.baseFill } : {}),
          ...(col?.rawDisplayFormatFill !== undefined ? { rawDisplayFormatFill: col.rawDisplayFormatFill } : {}),
          ...(col?.rawBaseFill !== undefined ? { rawBaseFill: col.rawBaseFill } : {}),
        }
      }) ?? observedCells

      let observedTodayBranch = observedTodayBranchFromSidecar
      if (isFormulaCase && c.clockRequirement && clockState) {
        // Stage 5: the branch is Excel's own TODAY() serial (the refreshed stale cell), not a formatted cell.
        observedTodayBranch = clockState.excelTodaySerial === c.clockRequirement.anchorSerial ? 'local' : clockState.excelTodaySerial === c.clockRequirement.anchorSerial - 1 ? 'utc' : undefined
      }
      if (c.clockRequirement && finalObservedCells && !isFormulaCase) {
        const localObs = finalObservedCells.find(
          o => o.ref.toUpperCase() === c.clockRequirement!.localTodayCellRef.toUpperCase(),
        )
        const utcObs = finalObservedCells.find(
          o => o.ref.toUpperCase() === c.clockRequirement!.utcTodayCellRef.toUpperCase(),
        )
        if (localObs && localObs.bgContrastDelta >= 16) observedTodayBranch = 'local'
        else if (utcObs && utcObs.bgContrastDelta >= 16) observedTodayBranch = 'utc'
      }

      writeFileSync(
        sidecarAbs,
          JSON.stringify(
            {
              excelVersion: sidecarExcelVersion,
              route: 'copy-picture',
              appearance: cap.appearance,
              scale: 1,
              pxPerPt: 1,
              width: decoded?.width,
              height: decoded?.height,
              colWidthsPt: capGeom?.colWidthsPt,
              rowHeightsPt: capGeom?.rowHeightsPt,
              geometrySource: hasLiveGeom ? 'excel-measured' : 'unverified',
              resavedByExcel: c.resavedByExcel === true,
              resaveSavedSha256: c.resaveReceipt?.savedSha256 ?? (c.resavedByExcel ? savedSha256 : undefined),
              sheet: c.sheet,
              range: c.range,
              hostTimeZone: sidecarHostTimeZone,
              savedSha256,
              pngSha256,
              capturedAt: sidecarCapturedAt,
              ...(clockState
                ? {
                    excelTodaySerial: clockState.excelTodaySerial,
                    staleTodayCacheRefreshed: clockState.staleTodayCacheRefreshed,
                    observedTodayBranch,
                  }
                : {}),
              ...(finalObservedCells ? { observedCells: finalObservedCells } : {}),
              ...(colorScaleCriteria ? { colorScaleCriteria } : {}),
              ...(barExtra ? barExtra.extras : {}),
              ...(iconExtra ? iconExtra.extras : {}),
              ...(isFormulaCase
                ? (() => {
                    const recorded = { formulaObservations: formulaObservations ?? {}, ...(formulaDynamic ? { formulaDynamicRefresh: formulaDynamic } : {}) }
                    return { lockBefore: sidecarLockBefore, lockAfter: sidecarLockAfter, ...recorded, formulaEvaluation: evaluateFormulaCase(cWithUpdatedExpected, recorded, decoded ?? undefined, capGeom ?? undefined, parts ?? undefined, stage5IconKey(REPO_ROOT)) }
                  })()
                : {}),
            },
            null,
            2,
          ) + '\n',
        )
      capturedCount++
      return { ...cap, pngSha256 }
    })

    if (lockRefusal || barRefusal || iconRefusal) {
      const why = lockRefusal ?? (barRefusal ? `bar measurement refused (${barRefusal})` : `icon measurement refused (${iconRefusal})`)
      manifest.cases[i] = { ...c, status: 'recipe-pending', ...(lockRefusal ? {} : { captures: updatedCaptures }), reason: lockRefusal ? `Capture discarded: ${why}` : `Capture kept but not promoted: ${why}` }
      console.warn(`[cf-refs/capture] ${c.id} remained recipe-pending: ${why}`)
      continue
    }

    const hasResave =
      c.resavedByExcel === true &&
      c.resaveReceipt !== undefined &&
      c.resaveReceipt.savedSha256 === savedSha256 &&
      c.resaveReceipt.openedRepairFree === true

    const barLabels = isBarCase && barObserved ? classifyBarCapture(cWithUpdatedExpected, barEvaluation ?? {}) : undefined
    const iconLabels = isIconCase && iconObserved ? classifyIconCapture(cWithUpdatedExpected, iconEvaluation ?? {}) : undefined
    const formulaFills = isFormulaCase ? formulaObservationOf(cWithUpdatedExpected, { formulaObservations: formulaObservations ?? {} }, undefined, undefined) : undefined
    const populatedExpectedCells = currentExpected.cells.map(exp => {
      if (formulaFills) {
        // Observed truth, never a guess: matched follows Excel's DisplayFormat (a fill other than the base, or a strike).
        const o = formulaFills.fill(exp.ref)
        return o ? { ...exp, matched: o.fill !== 'none' || o.strike, fillSource: 'excel-display-format' as const } : exp
      }
      if (iconLabels && exp.matched && exp.iconApplied !== undefined) {
        // Observed truth, never a guess: iconApplied follows the PNG and iconEvidence the verifier's own rules.
        const measured = iconObserved?.get(exp.ref)
        return { ...exp, ...(measured ? { iconApplied: measured.icon !== null } : {}), ...(exp.iconEvidence ? { iconEvidence: iconLabels.get(exp.ref) ?? { bucket: 'unverified' as const, geometry: 'unverified' as const } } : {}) }
      }
      if (barLabels && exp.matched) {
        // Observed truth, never a guess: barApplied follows the PNG and barEvidence the verifier's own rules.
        const measured = barObserved?.get(exp.ref)
        return { ...exp, ...(measured ? { barApplied: measured.bar !== null } : {}), barEvidence: barLabels.get(exp.ref) ?? { settings: 'unverified' as const, geometry: 'unverified' as const } }
      }
      const col = cellColors?.[exp.ref]
      if (col?.displayFormatFill && exp.matched && exp.scaleApplied) {
        return {
          ...exp,
          fillSource: 'excel-display-format' as const,
        }
      }
      return exp
    })

    const candidate: CfOracleCase = {
      ...cWithUpdatedExpected,
      status: hasResave && hasLiveGeom ? 'measured' : 'recipe-pending',
      captures: updatedCaptures,
      expected: {
        ...currentExpected,
        cells: populatedExpectedCells,
        provenance: hasResave && hasLiveGeom ? 'excel-observed' : 'seed-hypothesis',
      },
      ...(!hasResave
        ? {
            reason:
              'Excel repair-free resave provenance (resavedByExcel: true and valid resaveReceipt bound to savedSha256) required before promotion to measured',
          }
        : !hasLiveGeom
          ? {
              reason: 'Excel-measured range geometry required before promotion to measured',
            }
          : {}),
    }
    const check = verifyCfCase(candidate, { rootDir: REPO_ROOT })
    if (check.kind === 'verified') {
      const { reason: _droppedReason, ...rest } = candidate
      manifest.cases[i] = rest
      console.log(
        `[cf-refs/capture] verified ${c.id} (minContrast=${check.minPositiveContrastDelta}, maxContrast=${check.positiveContrastMaxDelta})`,
      )
    } else {
      manifest.cases[i] = {
        ...candidate,
        status: 'recipe-pending',
        expected: {
          ...candidate.expected,
          provenance: 'seed-hypothesis',
        },
        reason: `Capture verification failed (${check.code}): ${check.reason}`,
      }
      console.warn(`[cf-refs/capture] ${c.id} remained recipe-pending: ${check.reason}`)
    }
  }

  writeFileSync(CF_MANIFEST_PATH, JSON.stringify(manifest, null, 2) + '\n')
  console.log(`[cf-refs/capture] processed ${capturedCount} capture files`)
}
