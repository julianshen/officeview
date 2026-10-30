/**
 * AppleScript helpers for driving Excel for Mac (reference harness shared by sparklines and pivots).
 *
 * Excel completes `open` and `close` only after the sending script returns, so each step is its own script and a
 * workbook is addressed by name, never as the active workbook.
 */
import { execFileSync } from 'node:child_process'

export function osa(lines: string[], timeoutMs = 60_000): string {
  return execFileSync('osascript', lines.flatMap(l => ['-e', l]), { encoding: 'utf8', timeout: timeoutMs }).trim()
}
export const q = (s: string) => s.replace(/\\/g, '\\\\').replace(/"/g, '\\"')

/** Runs one script inside `tell application "Microsoft Excel"`. */
export function runScript(lines: string[], timeoutMs = 90_000): string {
  return osa(['tell application "Microsoft Excel"', ...lines, 'end tell'], timeoutMs)
}

/** Open workbook names; undefined while Excel is mid-transition (it then answers "missing value"). */
export const workbookNames = (): string[] | undefined => {
  const out = osa(['tell application "Microsoft Excel" to get name of every workbook'])
  if (out === 'missing value') return undefined
  return out === '' ? [] : out.split(', ')
}

export function waitFor(done: () => boolean, what: string, ms = 45_000): void {
  const deadline = Date.now() + ms
  while (!done()) {
    if (Date.now() > deadline) throw new Error(`timed out: ${what}`)
    execFileSync('sleep', ['0.25'])
  }
}

/** Opens a workbook and waits until Excel has registered it. Excel drops an open that arrives while a close is
 * pending, so an open that has not registered after 10 s is sent again. */
export function openWorkbook(file: string): string {
  const name = file.split('/').pop()!
  let sent = 0
  waitFor(() => {
    if (workbookNames()?.includes(name)) return true
    if (Date.now() - sent > 10_000) { execFileSync('open', ['-a', 'Microsoft Excel', file]); sent = Date.now() }
    return false
  }, `open ${name}`)
  return name
}

export function closeWorkbook(name: string, saving: boolean): void {
  osa([`tell application "Microsoft Excel" to close workbook "${q(name)}" saving ${saving ? 'yes' : 'no'}`])
  waitFor(() => !(workbookNames() ?? [name]).includes(name), `close ${name}`)
}

/** Saves an open workbook as xlsx at `path` (Excel then renames the open workbook to the new file name). */
export function saveAs(name: string, path: string): void {
  runScript([`save workbook as workbook "${q(name)}" filename (POSIX file "${q(path)}" as string) file format Excel XML file format`])
}
