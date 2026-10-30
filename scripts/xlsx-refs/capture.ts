/**
 * Captures Excel for Mac's own rendering of a range through AppleScript `copy picture … format bitmap`, read back
 * from the pasteboard as PNG (shared reference harness: sparklines and pivots).
 *
 * Probe findings (2026-10-27, Excel for Mac): the bitmap is a fixed 1 px per point (72 dpi) whatever the sheet
 * zoom, and the pasteboard's PDF only wraps the same bitmaps.
 */
import { execFileSync } from 'node:child_process'
import { readFileSync, renameSync } from 'node:fs'
import { openWorkbook, osa, q } from './excel'

/** Measured: Excel drew group-axis sparklines up to ~7 s after opening (2026-10-28); 10 s leaves margin. */
const FIRST_COPY_MS = 10_000

export interface CaptureOptions {
  /** Excel's on-screen look, or its printed look (pivot spec §7.10). Default screen. */
  appearance?: 'screen' | 'printer'
}

export function capture(file: string, sheetName: string, range: string, out: string, opts: CaptureOptions = {}): void {
  const name = openWorkbook(file)
  // Excel draws some content (group-axis sparklines, measured up to ~7 s) after the workbook opens, and a failed
  // copy leaves the previous pasteboard image. So the clipboard is cleared before every copy, the first copy waits
  // FIRST_COPY_MS, and a capture counts only once two copies at least 2 s apart are byte-identical. Copies go to a
  // temporary file; only a settled image replaces the reference.
  const opened = Date.now()
  const tmp = `${out}.tmp`
  let previous: Buffer | undefined
  const deadline = opened + 60_000
  for (;;) {
    const waitMs = previous ? 2_000 : Math.max(0, FIRST_COPY_MS - (Date.now() - opened))
    if (waitMs) execFileSync('sleep', [String(waitMs / 1000)])
    osa(['set the clipboard to ""'])
    osa([
      'tell application "Microsoft Excel"',
      `set wb to workbook "${q(name)}"`,
      'activate object wb',
      `set ws to worksheet "${q(sheetName)}" of wb`,
      'activate object ws',
      'set zoom of active window to 100',
      // A selection border would be captured, so the selection moves far away first.
      'select range "XFD1048576" of ws',
      `copy picture range "${range}" of ws appearance ${opts.appearance ?? 'screen'} format bitmap`,
      'end tell',
    ], 90_000)
    osa([
      `set f to open for access POSIX file "${q(tmp)}" with write permission`,
      'set eof f to 0',
      'write (the clipboard as «class PNGf») to f',
      'close access f',
    ])
    const current = readFileSync(tmp)
    if (previous && current.equals(previous)) { renameSync(tmp, out); return }
    if (Date.now() > deadline) throw new Error(`capture did not settle: ${name}`)
    previous = current
  }
}
