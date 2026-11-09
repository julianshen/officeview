/**
 * Captures Excel for Mac 16.114 screen and printer PNGs for Stage 1 Conditional Formatting oracle cases.
 *
 *   bun scripts/cf-refs/capture.ts [--stage 1] [--case <id>]
 *
 * Reuses `capture` from `scripts/xlsx-refs/capture.ts` and `osa` / `openWorkbook` from `scripts/xlsx-refs/excel.ts`.
 * Requires an awake, unlocked macOS display (`CGSSessionScreenIsLocked=false`).
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { capture } from '../xlsx-refs/capture'
import { openWorkbook, osa, q, runScript, workbookNames } from '../xlsx-refs/excel'
import { closeOwnedWorkbook, isDisplayLocked } from './build'
import { CF_CASES, type CfManifest, type CfOracleCase } from './cases'
import { CF_MANIFEST_PATH, CF_REFS_DIR } from './generate'
import {
  computeWorksheetRangeGeometry,
  decodePngSync,
  observeCellsFromCapture,
  readZipPartsSync,
  verifyCfCase,
} from './verify'

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
    const out = runScript(
      [
        `set wb to workbook "${q(name)}"`,
        `set ws to worksheet "${q(c.sheet)}" of wb`,
        `calculate`,
        `set v to value of range "${q(req.staleCacheCellRef)}" of ws`,
        `return v as string`,
      ],
      20_000,
    )
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

    // Numeric decimal integer string (e.g. VBA Color Long "13551615" or concatenated 8-bit channels "248105107")
    if (/^\d+$/.test(trimmed)) {
      const num = Number(trimmed)
      if (num > 16777215) {
        const channels = tryDecodeConcatenatedChannels(trimmed)
        if (channels) return channelsToArgb(channels[0], channels[1], channels[2])
      }
      return integerToArgb(num)
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

function tryDecodeConcatenatedChannels(s: string): [number, number, number] | null {
  if (s.length === 9) {
    const r = Number(s.slice(0, 3))
    const g = Number(s.slice(3, 6))
    const b = Number(s.slice(6, 9))
    if (r <= 255 && g <= 255 && b <= 255) return [r, g, b]
  }
  if (s.length === 8) {
    for (const [l1, l2] of [[2, 3], [3, 2], [3, 3]]) {
      const r = Number(s.slice(0, l1))
      const g = Number(s.slice(l1, l1 + l2))
      const b = Number(s.slice(l1 + l2))
      if (r <= 255 && g <= 255 && b <= 255) return [r, g, b]
    }
  }
  if (s.length === 7) {
    for (const [l1, l2] of [[2, 2], [2, 3], [3, 2], [1, 3]]) {
      const r = Number(s.slice(0, l1))
      const g = Number(s.slice(l1, l1 + l2))
      const b = Number(s.slice(l1 + l2))
      if (r <= 255 && g <= 255 && b <= 255) return [r, g, b]
    }
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
} | null {
  try {
    const name = openWorkbook(savedAbs)
    try {
      const refs = c.expected.cells.map(x => x.ref)
      const textScriptLines = buildTextProbeScript(name, c.sheet, refs)
      const textOut = runScript(textScriptLines, 30_000)
      const texts = textOut.split('|||')
      const cellTexts: Record<string, string> = {}
      refs.forEach((ref, idx) => {
        if (texts[idx] !== undefined) cellTexts[ref] = texts[idx]
      })

      const decScriptLines = buildDecorationProbeScript(name, c.sheet, refs)
      const decOut = runScript(decScriptLines, 30_000)
      const decs = decOut.split('|||')
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

      const colorScriptLines = buildColorProbeScript(name, c.sheet, refs)
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
        const colorOut = runScript(colorScriptLines, 30_000)
        const colorParts = colorOut.split('|||')
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

      return { cellTexts, cellDecorations, cellColors, geometry, colorScaleCriteria }
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
  for (let i = 0; i < manifest.cases.length; i++) {
    const c = manifest.cases[i]
    if (c.stage !== stage || (only && c.id !== only) || c.status === 'unsupported') continue
    const savedAbs = resolve(REPO_ROOT, c.saved)
    const savedBytes = new Uint8Array(readFileSync(savedAbs))
    const savedSha256 = sha256Hex(savedBytes)
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
        capture(savedAbs, c.sheet, c.range, pngAbs, { appearance: cap.appearance })
      }
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

      const observedCells =
        decoded && capGeom
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
      if (c.clockRequirement && finalObservedCells) {
        const localObs = finalObservedCells.find(
          o => o.ref.toUpperCase() === c.clockRequirement!.localTodayCellRef.toUpperCase(),
        )
        const utcObs = finalObservedCells.find(
          o => o.ref.toUpperCase() === c.clockRequirement!.utcTodayCellRef.toUpperCase(),
        )
        if (localObs && localObs.bgContrastDelta >= 16) observedTodayBranch = 'local'
        else if (utcObs && utcObs.bgContrastDelta >= 16) observedTodayBranch = 'utc'
      }

      if (!verifyOnly) {
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
            },
            null,
            2,
          ) + '\n',
        )
      }
      capturedCount++
      return { ...cap, pngSha256 }
    })

    const hasResave =
      c.resavedByExcel === true &&
      c.resaveReceipt !== undefined &&
      c.resaveReceipt.savedSha256 === savedSha256 &&
      c.resaveReceipt.openedRepairFree === true

    const populatedExpectedCells = currentExpected.cells.map(exp => {
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

  if (!verifyOnly) {
    writeFileSync(CF_MANIFEST_PATH, JSON.stringify(manifest, null, 2) + '\n')
  }
  console.log(`[cf-refs/capture] processed ${capturedCount} capture files`)
}
