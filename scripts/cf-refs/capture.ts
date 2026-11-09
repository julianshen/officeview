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
import { CF_STAGE1_CASES, type CfManifest, type CfOracleCase } from './cases'
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

export function probeLiveExcelCellObservations(
  c: CfOracleCase,
  savedAbs: string,
  preExisting: ReadonlySet<string>,
): {
  cellTexts: Record<string, string>
  cellDecorations: Record<string, { underline: 'none' | 'single' | 'double'; strikethrough: boolean }>
  geometry: { widthPt: number; heightPt: number; colWidthsPt: number[]; rowHeightsPt: number[] } | null
} | null {
  try {
    const name = openWorkbook(savedAbs)
    try {
      const refs = c.expected.cells.map(x => x.ref)
      const textScriptLines = [
        `set wb to workbook "${q(name)}"`,
        `set ws to worksheet "${q(c.sheet)}" of wb`,
        `set outList to {}`,
        ...refs.map(r => `set end of outList to (string value of range "${q(r)}" of ws as string)`),
        `set AppleScript's text item delimiters to "|||"`,
        `return outList as string`,
      ]
      const textOut = runScript(textScriptLines, 30_000)
      const texts = textOut.split('|||')
      const cellTexts: Record<string, string> = {}
      refs.forEach((ref, idx) => {
        if (texts[idx] !== undefined) cellTexts[ref] = texts[idx]
      })

      const decScriptLines = [
        `set wb to workbook "${q(name)}"`,
        `set ws to worksheet "${q(c.sheet)}" of wb`,
        `set outList to {}`,
        ...refs.map(
          r =>
            `set end of outList to ((underline of font object of display format of range "${q(r)}" of ws as string) & ":::" & (strikethrough of font object of display format of range "${q(r)}" of ws as string))`,
        ),
        `set AppleScript's text item delimiters to "|||"`,
        `return outList as string`,
      ]
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

      return { cellTexts, cellDecorations, geometry }
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
    if (c.stage > stage || (only && c.id !== only) || c.status === 'unsupported') continue
    const savedAbs = resolve(REPO_ROOT, c.saved)
    const savedBytes = new Uint8Array(readFileSync(savedAbs))
    const savedSha256 = sha256Hex(savedBytes)
    const parts = readZipPartsSync(savedBytes)
    const geom = parts ? computeWorksheetRangeGeometry(parts, c.range) : null

    let hasLiveGeom = false
    let activeGeom = geom
    let cellTexts: Record<string, string> | undefined
    let cellDecorations: Record<string, { underline?: 'none' | 'single' | 'double'; strikethrough?: boolean }> | undefined
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
          }
        } catch {
          // ignore parse errors
        }
      }
    } else {
      clockState = c.clockRequirement ? probeLiveExcelClockState(c, savedAbs, preExisting) : null
      const liveObs = probeLiveExcelCellObservations(c, savedAbs, preExisting)
      hasLiveGeom =
        liveObs?.geometry !== null &&
        liveObs?.geometry !== undefined &&
        liveObs.geometry.colWidthsPt.length > 0 &&
        liveObs.geometry.rowHeightsPt.length > 0
      activeGeom =
        hasLiveGeom && parts
          ? computeWorksheetRangeGeometry(parts, c.range, {
              colWidthsPt: liveObs.geometry.colWidthsPt,
              rowHeightsPt: liveObs.geometry.rowHeightsPt,
            }) ?? geom
          : geom
      cellTexts = liveObs?.cellTexts
      cellDecorations = liveObs?.cellDecorations
    }

    const updatedTemplate = CF_STAGE1_CASES.find(tc => tc.id === c.id)
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

      let observedTodayBranch = observedTodayBranchFromSidecar
      if (c.clockRequirement && observedCells) {
        const localObs = observedCells.find(
          o => o.ref.toUpperCase() === c.clockRequirement!.localTodayCellRef.toUpperCase(),
        )
        const utcObs = observedCells.find(
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
              ...(observedCells ? { observedCells } : {}),
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

    const candidate: CfOracleCase = {
      ...cWithUpdatedExpected,
      status: hasResave && hasLiveGeom ? 'measured' : 'recipe-pending',
      captures: updatedCaptures,
      expected: {
        ...currentExpected,
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
