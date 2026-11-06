/**
 * Captures Excel for Mac 16.114 screen and printer PNGs for Stage 1 Conditional Formatting oracle cases.
 *
 *   bun scripts/cf-refs/capture.ts [--stage 1] [--case <id>]
 *
 * Reuses `capture` from `scripts/xlsx-refs/capture.ts` and `osa` / `openWorkbook` from `scripts/xlsx-refs/excel.ts`.
 * Requires an awake, unlocked macOS display (`CGSSessionScreenIsLocked=false`).
 */
import { createHash } from 'node:crypto'
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { capture } from '../xlsx-refs/capture'
import { openWorkbook, osa, q, runScript, workbookNames } from '../xlsx-refs/excel'
import { closeOwnedWorkbook, isDisplayLocked } from './build'
import type { CfManifest, CfOracleCase } from './cases'
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

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2)
  const stage = Number(args.includes('--stage') ? args[args.indexOf('--stage') + 1] : 1)
  const only = args.includes('--case') ? args[args.indexOf('--case') + 1] : undefined

  if (isDisplayLocked()) {
    console.error(
      '[cf-refs/capture] Refusing capture: macOS display is currently locked (CGSSessionScreenIsLocked=true). Run on an awake, unlocked display.',
    )
    process.exit(1)
  }

  mkdirSync(CF_REFS_DIR, { recursive: true })
  const excelVersion = osa(['tell application "Microsoft Excel" to get version'])
  const preExisting = new Set<string>(workbookNames() ?? [])
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

    const clockState = c.clockRequirement ? probeLiveExcelClockState(c, savedAbs, preExisting) : null

    const updatedCaptures = c.captures.map(cap => {
      const pngAbs = resolve(REPO_ROOT, cap.pngPath)
      const sidecarAbs = resolve(REPO_ROOT, cap.sidecarPath)
      capture(savedAbs, c.sheet, c.range, pngAbs, { appearance: cap.appearance })
      const pngBytes = new Uint8Array(readFileSync(pngAbs))
      const pngSha256 = sha256Hex(pngBytes)
      const decoded = decodePngSync(pngBytes)
      const observedCells = decoded && geom ? observeCellsFromCapture(decoded, geom, c) : null

      let observedTodayBranch: 'local' | 'utc' | undefined
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

      writeFileSync(
        sidecarAbs,
        JSON.stringify(
          {
            excelVersion,
            route: 'copy-picture',
            appearance: cap.appearance,
            scale: 1,
            pxPerPt: 1,
            width: decoded?.width,
            height: decoded?.height,
            colWidthsPt: geom?.colWidthsPt,
            rowHeightsPt: geom?.rowHeightsPt,
            resavedByExcel: c.resavedByExcel === true,
            resaveSavedSha256: c.resaveReceipt?.savedSha256 ?? (c.resavedByExcel ? savedSha256 : undefined),
            sheet: c.sheet,
            range: c.range,
            hostTimeZone: 'Asia/Taipei',
            savedSha256,
            pngSha256,
            capturedAt: new Date().toISOString(),
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
      capturedCount++
      return { ...cap, pngSha256 }
    })

    const hasResave =
      c.resavedByExcel === true &&
      c.resaveReceipt !== undefined &&
      c.resaveReceipt.savedSha256 === savedSha256 &&
      c.resaveReceipt.openedRepairFree === true

    const candidate: CfOracleCase = {
      ...c,
      status: hasResave ? 'measured' : 'recipe-pending',
      captures: updatedCaptures,
      expected: {
        ...c.expected,
        savedSha256,
        provenance: hasResave ? 'excel-observed' : 'seed-hypothesis',
      },
      ...(!hasResave
        ? {
            reason:
              'Excel repair-free resave provenance (resavedByExcel: true and valid resaveReceipt bound to savedSha256) required before promotion to measured',
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
        ...c,
        captures: updatedCaptures,
        status: 'recipe-pending',
        expected: {
          ...c.expected,
          savedSha256,
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
