/**
 * Builds Stage 1 Conditional Formatting oracle workbooks through Excel for Mac 16.114.
 *
 *   bun scripts/cf-refs/build.ts --stage 1 [--resume] [--case <id>] [--refresh-clock-date]
 *
 * Rules:
 * - With `--resume`, skips ONLY cases where `verifyCfCase(c).kind === 'verified'` AND `inputSha256` matches the current input.
 * - With `--case <id>`, generates and mutates ONLY the selected case; other cases' resaved workbooks and manifest entries are preserved.
 * - Never closes unrelated open workbooks that were open before `build.ts` started.
 * - Preserves rejected Excel output (`<id>.rejected.xlsx`) and repair diagnostics if Excel strips or repairs a rule.
 */
import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { copyFileSync, existsSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { openWorkbook, osa, q, runScript, waitFor, workbookNames } from '../xlsx-refs/excel'
import type { CfManifest, CfOracleCase, S5ManualObservation } from './cases'
import { buildExplicitCalculationObservationScript, buildManualObservationScript, parseS5Observation } from './capture'
import { CF_FIXTURE_DIR, CF_MANIFEST_PATH, generateAllStage1 } from './generate'
import { readZipPartsSync, s5RetainedSqrefs, verifyCfCase } from './verify'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const sha256Hex = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex')

/** True only when the case is already `verified` (measured + captures + markers valid) AND its input hash matches. */
export function shouldSkipCaseOnResume(c: CfOracleCase, currentInputSha256: string, rootDir = REPO_ROOT): boolean {
  if (c.expected.inputSha256 !== currentInputSha256) return false
  const res = verifyCfCase(c, { rootDir })
  return res.kind === 'verified'
}

/** Filters manifest or catalog cases strictly to the targeted stage, excluding unsupported negative cases. */
export function filterCasesForStage(
  cases: readonly CfOracleCase[],
  stage: number,
  onlyCaseId?: string,
): CfOracleCase[] {
  return cases.filter(
    c => c.stage === stage && (!onlyCaseId || c.id === onlyCaseId) && c.status !== 'unsupported',
  )
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

const LOCK_KEY = '<key>CGSSessionScreenIsLocked</key>'

/**
 * Fail-closed parse of `ioreg -n Root -d1 -a` output (Ruling S3-A step 1, after `scripts/harness/webdriver.ts`'s
 * `<true/>`/`<false/>` match). `locked` for `<true/>`, for anything that is not recognisably ioreg plist output and for a
 * `CGSSessionScreenIsLocked` key whose value is neither `<true/>` nor `<false/>`. `unlocked` only for a recognisable
 * plist whose every `CGSSessionScreenIsLocked` is `<false/>`, or that does not carry the key at all.
 */
export function parseScreenLockState(ioregXml: string): 'locked' | 'unlocked' {
  if (typeof ioregXml !== 'string' || !/<plist[\s>]/.test(ioregXml) || !/<dict>/.test(ioregXml)) return 'locked'
  let from = 0
  for (;;) {
    const at = ioregXml.indexOf(LOCK_KEY, from)
    if (at < 0) return 'unlocked'
    const after = ioregXml.slice(at + LOCK_KEY.length)
    const value = /^\s*<(true|false)\s*\/>/.exec(after)
    if (!value || value[1] === 'true') return 'locked'
    from = at + LOCK_KEY.length
  }
}

const readIoreg = (): string => execFileSync('ioreg', ['-n', 'Root', '-d1', '-a'], { encoding: 'utf8', timeout: 5_000 })

/**
 * Whether macOS WindowServer currently has the console screen locked. Fails closed: an `ioreg` failure, a timeout or an
 * unparseable answer all count as locked, so a capture can never proceed on an unknown display. `run` is injectable
 * for tests.
 */
export function isDisplayLocked(run: () => string = readIoreg): boolean {
  try {
    return parseScreenLockState(run()) === 'locked'
  } catch {
    return true
  }
}

/**
 * Refuses a build on a locked (or unknown) display before any workbook is opened (review fix 3, Ruling S3-A step 1).
 * The per-capture re-check in capture.ts stays; this only stops a build from opening Excel workbooks behind a lock.
 */
export function assertBuildDisplayUnlocked(isLocked: () => boolean = () => isDisplayLocked()): void {
  if (isLocked()) throw new Error('display-locked: the macOS display is locked (or its state is unknown); refusing to open any workbook. Unlock it and rerun.')
}

/** Returns current `YYYY-MM-DD` in `Asia/Taipei` (`UTC+08:00`). */
export function currentTaipeiDateIso(nowMs = Date.now()): string {
  return new Date(nowMs + 8 * 3_600_000).toISOString().slice(0, 10)
}

/**
 * The scripts a build runs on its work copy after the open probe, in order (round-1 #4). An ordinary case resaves with the
 * unconditional `calculate`. A `calcObservation: 'pre-calculate'` case first observes with no calculation
 * (`buildManualObservationScript`), then runs one explicit calculation observation as a separate script, and resaves
 * without a `calculate`.
 */
export function buildCaseBuildScripts(
  c: CfOracleCase,
  workbookName: string,
  candidateAbs: string,
  observeRefs: readonly string[],
): Array<{ kind: 'manual-observation' | 'explicit-calculation' | 'resave'; lines: string[] }> {
  const preCalculate = c.calcObservation === 'pre-calculate'
  const resave = {
    kind: 'resave' as const,
    lines: [
      `set wb to workbook "${q(workbookName)}"`,
      ...(preCalculate ? [] : [`calculate`]),
      `save workbook as wb filename (POSIX file "${q(candidateAbs)}" as string) file format Excel XML file format`,
    ],
  }
  if (!preCalculate) return [resave]
  return [
    { kind: 'manual-observation', lines: buildManualObservationScript(workbookName, c.sheet, observeRefs) },
    { kind: 'explicit-calculation', lines: buildExplicitCalculationObservationScript(workbookName, c.sheet, observeRefs) },
    resave,
  ]
}

/** The cells a `pre-calculate` case observes: its stale sources, every `pre:`/`calc:` ref of its configurations and their controls. */
export function manualObservationRefs(c: CfOracleCase): string[] {
  const out = new Set<string>(Object.keys(c.calcObservationSources ?? {}))
  for (const cfg of c.formulaConfigurations ?? []) {
    for (const ref of cfg.refs) {
      if (!/^(pre|calc):/.test(ref)) continue
      const bare = ref.replace(/^(pre|calc):/, '')
      out.add(bare)
      const ctrl = c.expected.cells.find(e => e.ref === bare)?.controlRef
      if (ctrl) out.add(ctrl)
    }
  }
  return [...out]
}

export interface LiveExcelProbeResult {
  id: string
  /** Stage 5 (fix round 1 #2): the sqref order Excel retained per unpinned marker, read from the resaved XML. */
  retainedSqrefs?: Record<string, string>
  /** Stage 5 `pre-calculate` cases: the observations made before (pre) and after one explicit calculation (calc). */
  manualObservation?: Omit<S5ManualObservation, 'inputSha256'>
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
  isLocked: () => boolean = () => isDisplayLocked(),
): LiveExcelProbeResult {
  assertBuildDisplayUnlocked(isLocked)
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
    const observeRefs = manualObservationRefs(c)
    let manualObservation: LiveExcelProbeResult['manualObservation']
    let pre: ReturnType<typeof parseS5Observation> | undefined
    for (const script of buildCaseBuildScripts(c, name, candidateAbs, observeRefs)) {
      const out = runScript(script.lines, 30_000)
      if (script.kind === 'manual-observation') pre = parseS5Observation(out, observeRefs)
      if (script.kind === 'explicit-calculation' && pre) {
        const calc = parseS5Observation(out, observeRefs)
        manualObservation = { calculationMode: pre.calculationMode, workbookCount: pre.workbookCount, pre: pre.cells, calc: calc.cells }
      }
    }
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
    const retainedSqrefs = c.markers.some(m => m.sqrefOrder === 'any') && zipParts ? s5RetainedSqrefs(c, zipParts) : undefined
    return {
      id: c.id,
      ...(retainedSqrefs ? { retainedSqrefs } : {}),
      ...(manualObservation ? { manualObservation } : {}),
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
  const refreshClock = args.includes('--refresh-clock-date')
  const anchorLocalDate = refreshClock
    ? currentTaipeiDateIso()
    : args.includes('--anchor-date')
      ? args[args.indexOf('--anchor-date') + 1]
      : undefined

  const manifest: CfManifest = await generateAllStage1({ stage, onlyCaseId: only, anchorLocalDate })
  assertBuildDisplayUnlocked()
  execFileSync('open', ['-a', 'Microsoft Excel'])
  execFileSync('sleep', ['2'])
  const preExisting = new Set<string>(workbookNames() ?? [])
  const displayLocked = isDisplayLocked()
  console.log(
    `[cf-refs/build] Excel 16.114 pre-existing workbooks: [${[...preExisting].join(', ')}]; displayLocked=${displayLocked}`,
  )

  for (let i = 0; i < manifest.cases.length; i++) {
    const c = manifest.cases[i]
    // An unsupported case is skipped unless it is named with `--case` (stage3-bar-unresolved-shapes records Excel's
    // open/repair/strip outcome that way; a stripped result is preserved as `<id>.rejected.xlsx`, never as the saved file).
    if (c.stage !== stage || (only && c.id !== only) || (c.status === 'unsupported' && only !== c.id)) continue
    const inputBytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, c.input)))
    const currentInputSha256 = sha256Hex(inputBytes)
    if (resume && shouldSkipCaseOnResume(c, currentInputSha256, REPO_ROOT)) {
      console.log(`[cf-refs/build] skip verified ${c.id}`)
      continue
    }
    const probe = probeAndBuildCase(c, preExisting, displayLocked)
    const savedBytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved)))
    const savedSha256 = sha256Hex(savedBytes)
    const isResaved = probe.resaved || Boolean(c.resavedByExcel)
    const resaveReceipt = probe.resaved
      ? {
          resavedAt: new Date().toISOString(),
          excelVersion: osa(['tell application "Microsoft Excel" to get version']).trim(),
          inputSha256: currentInputSha256,
          savedSha256,
          openedRepairFree: true as const,
        }
      : c.resaveReceipt && c.resaveReceipt.savedSha256 === savedSha256
        ? c.resaveReceipt
        : undefined

    manifest.cases[i] = {
      ...c,
      // Stage 5 (round-1 #4): the pre-calculation observation is bound to the input it was made on.
      ...(probe.manualObservation ? { manualObservation: { inputSha256: currentInputSha256, ...probe.manualObservation } } : {}),
      ...(probe.retainedSqrefs ? { retainedSqrefs: probe.retainedSqrefs } : {}),
      resavedByExcel: isResaved,
      ...(resaveReceipt ? { resaveReceipt } : {}),
      expected: {
        ...c.expected,
        inputSha256: currentInputSha256,
        savedSha256,
      },
      reason: probe.diagnostic ?? c.reason,
    }
    console.log(
      `[cf-refs/build] ${c.id}: openedRepairFree=${probe.openedRepairFree}, fcCount=${probe.formatConditionsCount}, firstRuleType=${probe.firstRuleType}, resaved=${probe.resaved}`,
    )
  }

  writeFileSync(CF_MANIFEST_PATH, JSON.stringify(manifest, null, 2) + '\n')
}
