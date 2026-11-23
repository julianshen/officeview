/**
 * Stage 4a (S4 §6.1, §6.11): a tiles-mode stream session in the full local interactive viewer. Loaded lazily from
 * RemoteDocView when `document.sheets` is set; it mounts `InteractiveSheetDoc` (its chunk is shared with the local
 * viewer) over `streamShellDoc` (sheet names only) and a `StreamWorkbook` (the server's layouts and tiles).
 *
 * - **One shell per stream.** The viewer keys its mount on the identity of both its document and its source, so both
 *   are held per `RemoteOfficeDoc` (a WeakMap): a parent re-render never remounts the viewer or drops its tiles. The
 *   source is closed when the session ends, never on unmount (StrictMode's simulated unmount would close it for good).
 * - **No text, no copy, no find, no selection** (R6): the viewer shows positions only, and says `TEXT_UNAVAILABLE` once.
 * - **Stage 4b** (`manifest.sheetView === 1`): the viewer selects, resizes and fits as local mode does; search is always
 *   on and copy follows `can.copy` (re-read on every manifest change), each answered by the server (S4 §6.7, P15).
 * - **Print** is the server's static per-sheet units through `RemoteOfficeDoc.print` (the Stage 3b flow; the range
 *   picker counts sheets), started from the viewer's toolbar or Ctrl/Cmd+P; the viewer's screen never prints.
 */
import { useEffect, useRef, useState, type CSSProperties, type ReactElement } from 'react'
import InteractiveSheetDoc, { type SheetFrameStats } from '../xlsx/InteractiveSheetDoc'
import { StreamWorkbook, streamShellDoc } from '../xlsx/streamWorkbook'
import type { XlsxDocument } from '../../xlsx/types'
import type { RemoteOfficeDoc } from '../../stream/remote/remote-doc'
import type { ThemeMode } from '../theme'
import { useRemotePrint } from './RemotePrint'
import { TEXT_UNAVAILABLE } from './textUnavailable'
import { copyFailure } from './RemoteText'

export interface RemoteSheetViewProps {
  document: RemoteOfficeDoc
  theme?: ThemeMode
  className?: string
  style?: CSSProperties
  background?: string
  allowPrint?: boolean
  title?: string
  /** @internal The viewer's frame stats (the stream harness page). */
  onFrame?: (stats: SheetFrameStats) => void
}

interface StreamShell { doc: XlsxDocument; source: StreamWorkbook }
const shells = new WeakMap<RemoteOfficeDoc, StreamShell>()
/** The one shell document and source of a stream (created on first use, closed with the session). */
export function streamShellFor(document: RemoteOfficeDoc): StreamShell {
  let shell = shells.get(document)
  if (!shell) {
    const sheets = document.sheets
    if (!sheets) throw new TypeError('not a tiles-mode stream')
    const source = new StreamWorkbook(sheets, document.manifest.docRef)
    shell = { doc: streamShellDoc(sheets.list), source }
    shells.set(document, shell)
    void document.ended.then(() => source.close())
  }
  return shell
}

/** The print controls sit in the viewer's toolbar, laid out inline (the dialog opens under them). */
const INLINE_BAR: CSSProperties = { position: 'relative', display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }

export default function RemoteSheetView(props: RemoteSheetViewProps): ReactElement {
  const { document: doc, theme, className, style, background, allowPrint = true, title, onFrame } = props
  const shell = streamShellFor(doc)
  const root = useRef<HTMLDivElement | null>(null)
  const print = useRemotePrint(doc, { allowPrint, themeRoot: root, noun: 'sheets', bar: INLINE_BAR, ...(title !== undefined ? { title } : {}) })
  // §10: a session that ends says so once; recovery is a host-driven re-open.
  const [ended, setEnded] = useState<string>()
  useEffect(() => {
    let live = true
    void doc.ended.then((e) => { if (live && e) setEnded(e.code) })
    return () => { live = false }
  }, [doc])
  // A resume may change `can.copy` (S3 item 2): the copy controls follow the current manifest.
  const [, setManifestRev] = useState(0)
  useEffect(() => doc.onManifestChange?.(() => setManifestRev((n) => n + 1)), [doc])
  const sheetView = doc.manifest.sheetView === 1
  const allowCopy = doc.manifest.can.copy && sheetView
  return (
    <div ref={root} data-ov-remote-sheet="" className={className} style={{ display: 'flex', flexDirection: 'column', minHeight: 0, ...style }}>
      {ended ? <div role="alert" data-stream-error={ended}>Viewing session ended ({ended}).</div> : null}
      <InteractiveSheetDoc doc={shell.doc} source={shell.source}
        {...(theme !== undefined ? { theme } : {})} {...(background !== undefined ? { background } : {})} {...(onFrame ? { onFrame } : {})}
        allowCopy={allowCopy} showSearch={sheetView}
        remoteChrome={{ textUnavailable: TEXT_UNAVAILABLE, print: print.start, toolbar: print.ui, copyFailure: (e) => copyFailure(e, doc.manifest.caps.copy) }} />
    </div>
  )
}
