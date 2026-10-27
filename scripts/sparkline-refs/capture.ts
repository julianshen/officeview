/**
 * Captures Excel for Mac's own rendering of every sparkline case (spec section 8) through AppleScript
 * `copy picture … appearance screen format bitmap`, read back from the pasteboard as PNG.
 *
 * Probe findings (2026-10-27, Excel for Mac): the bitmap holds the sparklines at a fixed 1 px per point (72 dpi)
 * whatever the sheet zoom, and the pasteboard's PDF only wraps the same bitmaps. `open` returns before the
 * workbook is registered, so each capture waits for the expected workbook name.
 *
 *   bun scripts/sparkline-refs/capture.ts [--stage N] [--case id]
 *
 * Writes tests/fixtures/xlsx-sparklines/refs/<id>[-bordered][-<sheet>].png plus a .json sidecar.
 */
import { execFileSync } from 'node:child_process'
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { SPARKLINE_CASES, colIndex, colName, type CaseSheet } from './cases'

const FIXTURES = resolve(import.meta.dir, '../../tests/fixtures/xlsx-sparklines')
const REFS = join(FIXTURES, 'refs')

/** Columns I (the host's left neighbour, for spill and border context) through the last host column, over the host rows. */
export function captureRange(sheet: CaseSheet): string | undefined {
  const hosts = sheet.groups.flatMap(g => g.sparklines.map(s => s.sqref))
  if (hosts.length === 0) return undefined
  let r0 = Infinity, r1 = 0, c1 = 0
  for (const ref of hosts) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref)
    if (!m) continue
    const row = Number(m[2])
    let col = colIndex(m[1])
    for (const merge of sheet.merges ?? []) {
      const mm = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(merge)
      if (mm && mm[1] === m[1] && Number(mm[2]) === row) col = Math.max(col, colIndex(mm[3]))
    }
    r0 = Math.min(r0, row); r1 = Math.max(r1, row); c1 = Math.max(c1, col)
  }
  return Number.isFinite(r0) ? `I${r0}:${colName(c1)}${r1}` : undefined
}

function osa(lines: string[], timeoutMs = 60_000): string {
  return execFileSync('osascript', lines.flatMap(l => ['-e', l]), { encoding: 'utf8', timeout: timeoutMs }).trim()
}
const q = (s: string) => s.replace(/\\/g, '\\\\').replace(/"/g, '\\"')

/** Open workbook names; undefined while Excel is mid-transition (it then answers "missing value"). */
const workbookNames = (): string[] | undefined => {
  const out = osa(['tell application "Microsoft Excel" to get name of every workbook'])
  if (out === 'missing value') return undefined
  return out === '' ? [] : out.split(', ')
}
function waitFor(done: () => boolean, what: string): void {
  const deadline = Date.now() + 45_000
  while (!done()) {
    if (Date.now() > deadline) throw new Error(`timed out: ${what}`)
    execFileSync('sleep', ['0.25'])
  }
}

/**
 * Excel completes `open` and `close` only after the sending script returns, so each step is its own script and the
 * workbook is addressed by name, never as the active workbook.
 */
function capture(file: string, sheetName: string, range: string, out: string): void {
  const name = file.split('/').pop()!
  // Excel drops an open that arrives while a close is pending, so workbooks stay open for the whole run and an
  // open that has not registered after 10 s is sent again.
  let sent = 0
  waitFor(() => {
    if (workbookNames()?.includes(name)) return true
    if (Date.now() - sent > 10_000) { execFileSync('open', ['-a', 'Microsoft Excel', file]); sent = Date.now() }
    return false
  }, `open ${name}`)
  // Excel draws some sparklines (group axes, measured: up to ~7 s) after the workbook opens, and a failed copy
  // leaves the previous pasteboard image. So the clipboard is cleared before every copy, and a capture counts only
  // once two copies at least 2 s apart, taken no earlier than 3 s after opening, are byte-identical.
  const opened = Date.now()
  let previous: Buffer | undefined
  const deadline = opened + 60_000
  for (;;) {
    const waitMs = previous ? 2_000 : Math.max(0, 3_000 - (Date.now() - opened))
    if (waitMs) execFileSync('sleep', [String(waitMs / 1000)])
    osa(['set the clipboard to ""'])
    osa([
      'tell application "Microsoft Excel"',
      `set wb to workbook "${q(name)}"`,
      'activate object wb',
      `set ws to worksheet "${q(sheetName)}" of wb`,
      'activate object ws',
      'set zoom of active window to 100',
      // A selection border would be captured, so the selection moves far away first (spec section 8).
      'select range "XFD1048576" of ws',
      `copy picture range "${range}" of ws appearance screen format bitmap`,
      'end tell',
    ], 90_000)
    osa([
      `set f to open for access POSIX file "${q(out)}" with write permission`,
      'set eof f to 0',
      'write (the clipboard as «class PNGf») to f',
      'close access f',
    ])
    const current = readFileSync(out)
    if (previous && current.equals(previous)) return
    if (Date.now() > deadline) throw new Error(`capture did not settle: ${name}`)
    previous = current
  }
}

if (import.meta.main) {
  const args = process.argv.slice(2)
  const stage = Number(args[args.indexOf('--stage') + 1] || 3)
  const only = args.includes('--case') ? args[args.indexOf('--case') + 1] : undefined
  mkdirSync(REFS, { recursive: true })
  const excelVersion = osa(['tell application "Microsoft Excel" to get version'])
  let count = 0
  for (const c of SPARKLINE_CASES) {
    if (c.stage > stage || (only && c.id !== only)) continue
    for (const variant of ['', '-bordered']) {
      for (const sheet of c.sheets) {
        const range = captureRange(sheet)
        if (!range) continue
        const base = `${c.id}${variant}${c.sheets.length > 1 ? `-${sheet.name}` : ''}`
        const png = join(REFS, `${base}.png`)
        capture(join(FIXTURES, `${c.id}${variant}.xlsx`), sheet.name, range, png)
        writeFileSync(join(REFS, `${base}.json`), JSON.stringify({ excelVersion, route: 'copy-picture', scale: 1, pxPerPt: 1, sheet: sheet.name, range, capturedAt: new Date().toISOString() }, null, 2) + '\n')
        count++
        console.log(`captured ${base} (${range})`)
      }
    }
  }
  console.log(`${count} references (workbooks left open; close Excel windows when done)`)
}
