/**
 * Builds Stage 1 Conditional Formatting oracle workbooks through Excel for Mac 16.114.
 *
 *   bun scripts/cf-refs/build.ts --stage 1 [--resume] [--case <id>]
 *
 * Rules:
 * -With `--resume`, skips ONLY cases where `verifyCfCase(c).kind === 'verified'` AND `inputSha256` matches the current input.
 * -Never closes unrelated open workbooks that were open before `build.ts` started.
 * -Preserves rejected Excel output (`<id>.rejected.xlsx`) and repair diagnostics if Excel strips or repairs a rule.
 */
import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { copyFileSync, existsSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { openWorkbook, osa, q, runScript, waitFor, workbookNames } from '../xlsx-refs/excel'
import type { CfManifest, CfOracleCase } from './cases'
import { CF_FIXTURE_DIR, CF_MANIFEST_PATH, generateAllStage1 } from './generate'
import { readZipPartsSync, verifyCfCase } from './verify'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const sha256Hex = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex')

/** True only when the case is already `verified` (measured + captures + markers valid) AND its input hash matches. */
export function shouldSkipCaseOnResume(c: CfOracleCase, currentInputSha256: string, rootDir = REPO_ROOT): boolean {
  if (c.expected.inputSha256 !== currentInputSha256) return false
  const res = verifyCfCase(c, { rootDir })
  return res.kind === 'verified'
}

/** Enforces that `build.ts` never closes a workbook that was already open before the build started. */
export function assertCanCloseWorkbook(targetName: string, preExistingOpenWorkbooks: ReadonlySet<string>): void {
  if (preExistingOpenWorkbooks.has(targetName)) {
    throw new Error(`Refusing to close unrelated pre-existing open workbook: ${targetName}`)
  }
}

/** Safe close that treats `workbookNames() === undefined` ("missing value" when 0 workbooks remain) as closed. */
export function closeOwnedWorkbook(name: string, preExistingOpenWorkbooks: ReadonlySet<string>): void {
  assertCanCloseWorkbook(name, preExistingOpenWorkbooks)
  osa([`tell application "Microsoft Excel" to close workbook "${q(name)}" saving no`], 15_000)
  waitFor(() => {
    const names = workbookNames()
    return names === undefined || !names.includes(name)
  }, `close ${name}`, 15_000)
}

/** Detects whether macOS WindowServer currently has the console screen locked (`CGSSessionScreenIsLocked`). */
export function isDisplayLocked(): boolean {
  try {
    const out = execFileSync('ioreg', ['-n', 'Root', '-d1', '-a'], { encoding: 'utf8', timeout: 5_000 })
    return out.includes('<key>CGSSessionScreenIsLocked</key>')
  } catch {
    return false
  }
}

export interface LiveExcelProbeResult {
  id: string
  openedRepairFree: boolean
  excelSavedFlag: boolean
  firstRangeSqref: string
  formatConditionsCount: number
  firstRuleType?: string
  resaved: boolean
  diagnostic?: string
}

export function probeAndBuildCase(
  c: CfOracleCase,
  preExistingOpenWorkbooks: ReadonlySet<string>,
  displayLocked: boolean,
): LiveExcelProbeResult {
  const inputAbs = resolve(REPO_ROOT, c.input)
  const savedAbs = resolve(REPO_ROOT, c.saved)
  const workName = `.work-${c.id}.xlsx`
  const workAbs = join(CF_FIXTURE_DIR, workName)
  rmSync(workAbs, { force: true })
  copyFileSync(inputAbs, workAbs)

  const firstSqref = (c.markers[0]?.sqref ?? 'B2').split(' ')[0]
  const name = openWorkbook(workAbs)
  try {
    const rawInfo = runScript(
      [
        `set wb to workbook "${q(name)}"`,
        `set ws to worksheet "${q(c.sheet)}" of wb`,
        `set rng to range "${q(firstSqref)}" of ws`,
        `set fcCount to count of format conditions of rng`,
        `set fcType to "none"`,
        `if fcCount > 0 then`,
        `  try`,
        `    set fcType to (format condition type of format condition 1 of rng) as string`,
        `  on error`,
        `    set fcType to "x14-or-special"`,
        `  end try`,
        `end if`,
        `return {(saved of wb) as string, fcCount as string, fcType}`,
      ],
      20_000,
    )
    const parts = rawInfo.split(', ')
    const excelSavedFlag = parts[0] === 'true'
    const formatConditionsCount = Number(parts[1] ?? 0)
    const firstRuleType = parts[2]

    if (displayLocked) {
      closeOwnedWorkbook(name, preExistingOpenWorkbooks)
      return {
        id: c.id,
        openedRepairFree: excelSavedFlag,
        excelSavedFlag,
        firstRangeSqref: firstSqref,
        formatConditionsCount,
        firstRuleType,
        resaved: false,
        diagnostic:
          'Opened repair-free in Excel 16.114; resave/capture deferred because macOS display is locked (CGSSessionScreenIsLocked=true).',
      }
    }

    const candidateAbs = join(CF_FIXTURE_DIR, `${c.id}.candidate.xlsx`)
    rmSync(candidateAbs, { force: true })
    runScript(
      [
        `set wb to workbook "${q(name)}"`,
        `calculate`,
        `save workbook as wb filename (POSIX file "${q(candidateAbs)}" as string) file format Excel XML file format`,
      ],
      20_000,
    )
    closeOwnedWorkbook(`${c.id}.candidate.xlsx`, preExistingOpenWorkbooks)

    if (!existsSync(candidateAbs)) {
      return {
        id: c.id,
        openedRepairFree: excelSavedFlag,
        excelSavedFlag,
        firstRangeSqref: firstSqref,
        formatConditionsCount,
        firstRuleType,
        resaved: false,
        diagnostic: 'Excel saveAs did not produce candidate file (check folder permission grant).',
      }
    }

    const candidateBytes = new Uint8Array(readFileSync(candidateAbs))
    const zipParts = readZipPartsSync(candidateBytes)
    const hasRepair = zipParts ? [...zipParts.keys()].some(k => /repaired/i.test(k)) : true
    const tempCheck = verifyCfCase(
      {
        ...c,
        saved: `tests/fixtures/xlsx-cf/${c.id}.candidate.xlsx`,
        status: 'recipe-pending',
        expected: { ...c.expected, savedSha256: sha256Hex(candidateBytes) },
      },
      { rootDir: REPO_ROOT },
    )
    if (hasRepair || (tempCheck.kind === 'refusal' && !tempCheck.structuralValid)) {
      const rejectedAbs = join(CF_FIXTURE_DIR, `${c.id}.rejected.xlsx`)
      renameSync(candidateAbs, rejectedAbs)
      return {
        id: c.id,
        openedRepairFree: false,
        excelSavedFlag,
        firstRangeSqref: firstSqref,
        formatConditionsCount,
        firstRuleType,
        resaved: false,
        diagnostic: `Preserved rejected Excel output at ${rejectedAbs}: ${tempCheck.kind === 'refusal' ? tempCheck.reason : 'repair detected'}`,
      }
    }

    renameSync(candidateAbs, savedAbs)
    return {
      id: c.id,
      openedRepairFree: true,
      excelSavedFlag,
      firstRangeSqref: firstSqref,
      formatConditionsCount,
      firstRuleType,
      resaved: true,
    }
  } finally {
    rmSync(workAbs, { force: true })
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2)
  const stage = Number(args.includes('--stage') ? args[args.indexOf('--stage') + 1] : 1)
  const resume = args.includes('--resume')
  const only = args.includes('--case') ? args[args.indexOf('--case') + 1] : undefined

  const manifest: CfManifest = await generateAllStage1()
  // Ensure Excel is running before reading pre-existing workbooks
  execFileSync('open', ['-a', 'Microsoft Excel'])
  execFileSync('sleep', ['2'])
  const preExisting = new Set<string>(workbookNames() ?? [])
  const displayLocked = isDisplayLocked()
  console.log(
    `[cf-refs/build] Excel 16.114 pre-existing workbooks: [${[...preExisting].join(', ')}]; displayLocked=${displayLocked}`,
  )

  for (let i = 0; i < manifest.cases.length; i++) {
    const c = manifest.cases[i]
    if (c.stage > stage || (only && c.id !== only)) continue
    const inputBytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, c.input)))
    const currentInputSha256 = sha256Hex(inputBytes)
    if (resume && shouldSkipCaseOnResume(c, currentInputSha256, REPO_ROOT)) {
      console.log(`[cf-refs/build] skip verified ${c.id}`)
      continue
    }
    const probe = probeAndBuildCase(c, preExisting, displayLocked)
    const savedBytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved)))
    manifest.cases[i] = {
      ...c,
      expected: {
        ...c.expected,
        inputSha256: currentInputSha256,
        savedSha256: sha256Hex(savedBytes),
      },
      reason: probe.diagnostic ?? c.reason,
    }
    console.log(
      `[cf-refs/build] ${c.id}: openedRepairFree=${probe.openedRepairFree}, fcCount=${probe.formatConditionsCount}, firstRuleType=${probe.firstRuleType}, resaved=${probe.resaved}`,
    )
  }

  writeFileSync(CF_MANIFEST_PATH, JSON.stringify(manifest, null, 2) + '\n')
}
