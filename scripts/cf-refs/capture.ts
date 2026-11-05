/**
 * Captures Excel for Mac 16.114 screen and printer PNGs for Stage 1 Conditional Formatting oracle cases.
 *
 *   bun scripts/cf-refs/capture.ts [--stage 1] [--case <id>]
 *
 * Reuses `capture` from `scripts/xlsx-refs/capture.ts` and `osa` from `scripts/xlsx-refs/excel.ts`.
 * Requires an awake, unlocked macOS display (`CGSSessionScreenIsLocked=false`).
 */
import { createHash } from 'node:crypto'
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { capture } from '../xlsx-refs/capture'
import { osa } from '../xlsx-refs/excel'
import { isDisplayLocked } from './build'
import type { CfManifest } from './cases'
import { CF_MANIFEST_PATH, CF_REFS_DIR } from './generate'
import { verifyCfCase } from './verify'

export { capture }

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const sha256Hex = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex')

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
  const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest

  let capturedCount = 0
  for (let i = 0; i < manifest.cases.length; i++) {
    const c = manifest.cases[i]
    if (c.stage > stage || (only && c.id !== only) || c.status === 'unsupported') continue
    const savedAbs = resolve(REPO_ROOT, c.saved)
    const savedSha256 = sha256Hex(new Uint8Array(readFileSync(savedAbs)))

    const updatedCaptures = c.captures.map(cap => {
      const pngAbs = resolve(REPO_ROOT, cap.pngPath)
      const sidecarAbs = resolve(REPO_ROOT, cap.sidecarPath)
      capture(savedAbs, c.sheet, c.range, pngAbs, { appearance: cap.appearance })
      const pngSha256 = sha256Hex(new Uint8Array(readFileSync(pngAbs)))
      writeFileSync(
        sidecarAbs,
        JSON.stringify(
          {
            excelVersion,
            route: 'copy-picture',
            appearance: cap.appearance,
            scale: 1,
            pxPerPt: 1,
            sheet: c.sheet,
            range: c.range,
            savedSha256,
            pngSha256,
            capturedAt: new Date().toISOString(),
          },
          null,
          2,
        ) + '\n',
      )
      capturedCount++
      return { ...cap, pngSha256 }
    })

    const candidate = {
      ...c,
      status: 'measured' as const,
      captures: updatedCaptures,
      expected: { ...c.expected, savedSha256 },
    }
    const check = verifyCfCase(candidate, { rootDir: REPO_ROOT })
    if (check.kind === 'verified') {
      const { reason: _droppedReason, ...rest } = candidate
      manifest.cases[i] = rest
      console.log(`[cf-refs/capture] verified ${c.id} (maxContrast=${check.positiveContrastMaxDelta})`)
    } else {
      manifest.cases[i] = {
        ...c,
        captures: updatedCaptures,
        status: 'recipe-pending',
        reason: `Capture verification failed (${check.code}): ${check.reason}`,
      }
      console.warn(`[cf-refs/capture] ${c.id} remained recipe-pending: ${check.reason}`)
    }
  }

  writeFileSync(CF_MANIFEST_PATH, JSON.stringify(manifest, null, 2) + '\n')
  console.log(`[cf-refs/capture] processed ${capturedCount} capture files`)
}
