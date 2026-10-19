import type { PageSpec, Paintable } from '../../render/paint'

// Same raster scale and axis ceiling as InteractiveSheetDoc.tsx's print path.
export const PRINT_SCALE = 2
export const PRINT_MAX_AXIS = 16384
export const PRINT_MAX_BYTES = 64 * 1024 * 1024
export const PRINT_JOB_MAX_DECODED = 1024 * 1024 * 1024
export const PRINT_JOB_MAX_BLOB_BYTES = 192 * 1024 * 1024
export type PrintPlan = { kind: 'ok'; scale: 1 | 2 } | { kind: 'refused'; reason: 'too-large' }
export function planPrintJob(specs: readonly PageSpec[], budgets = {
  decoded: PRINT_JOB_MAX_DECODED, workspace: PRINT_MAX_BYTES, maxAxis: PRINT_MAX_AXIS,
}): PrintPlan {
  for (const scale of [PRINT_SCALE, 1] as const) {
    let total = 0
    if (specs.every(({ widthPx, heightPx }) => {
      const w = Math.ceil(widthPx * scale), h = Math.ceil(heightPx * scale)
      const bytes = w * h * 4
      total += bytes
      return w > 0 && h > 0 && w <= budgets.maxAxis && h <= budgets.maxAxis &&
        bytes <= budgets.workspace && total <= budgets.decoded
    })) return { kind: 'ok', scale }
  }
  return { kind: 'refused', reason: 'too-large' }
}
function sizes(specs: readonly PageSpec[]): PageSpec[] {
  return [...new Map(specs.map(s => [`${s.widthPx},${s.heightPx}`, s])).values()]
}
/** Older engines without named pages use one enclosing page size, with each
 * smaller page centred inside it rather than stretched or cropped. */
export function pageRules(specs: readonly PageSpec[], namedPages: boolean): string {
  const distinct = sizes(specs)
  if (namedPages) return distinct.map((s, n) => `@page p${n} { size: ${s.widthPx}px ${s.heightPx}px; margin: 0 }
img[data-page="p${n}"] { page: p${n}; display: block; width: ${s.widthPx}px; height: ${s.heightPx}px; break-after: page }`).join('\n')
  const w = distinct.reduce((max, s) => Math.max(max, s.widthPx), 0)
  const h = distinct.reduce((max, s) => Math.max(max, s.heightPx), 0)
  return `@page { size: ${w}px ${h}px; margin: 0 }
img { display: block; width: ${w}px; height: ${h}px; object-fit: contain; break-after: page }`
}
export interface PrintJob { frame: HTMLIFrameElement; dispose(): void }
const TOO_LARGE = 'This document is too large to print from the viewer.'
export async function printDocument(pages: readonly Paintable[], themeRoot: HTMLElement, opts: {
  signal: AbortSignal; title?: string; onNotice(text: string): void
  rasterize?: (page: Paintable, scale: 1 | 2, canvas: HTMLCanvasElement) => Promise<Blob | null>
}): Promise<PrintJob | null> {
  if (opts.signal.aborted || !pages.length) return null
  const specs = pages.map(p => p.spec), plan = planPrintJob(specs)
  if (plan.kind === 'refused') { opts.onNotice(TOO_LARGE); return null }
  const doc = themeRoot.ownerDocument, host = doc.defaultView!
  const canvas = doc.createElement('canvas'), frame = doc.createElement('iframe')
  // A detached canvas otherwise loses the viewer's inherited font language.
  canvas.lang = themeRoot.closest('[lang]')?.getAttribute('lang') ?? ''
  frame.setAttribute('aria-hidden', 'true'); frame.tabIndex = -1
  frame.setAttribute('data-officeview-print-frame', '')
  frame.style.cssText = 'position:fixed;width:1px;height:1px;opacity:0;pointer-events:none;border:0'
  const urls: string[] = []
  let disposed = false, stopWaiting: (() => void) | undefined
  const dispose = () => {
    if (disposed) return
    disposed = true
    stopWaiting?.()
    opts.signal.removeEventListener('abort', dispose)
    host.removeEventListener('afterprint', dispose)
    frame.contentWindow?.removeEventListener('afterprint', dispose)
    frame.remove()
    for (const url of urls) URL.revokeObjectURL(url)
    urls.length = 0
    canvas.width = canvas.height = 0
  }
  opts.signal.addEventListener('abort', dispose, {once: true})
  const job = {frame, dispose}
  try {
    themeRoot.appendChild(frame)
    const inner = frame.contentDocument!, win = frame.contentWindow!
    host.addEventListener('afterprint', dispose)
    win.addEventListener('afterprint', dispose)
    inner.title = opts.title ?? ''
    const style = inner.createElement('style')
    style.textContent = 'html,body { margin:0; padding:0; background:white }\n' +
      pageRules(specs, !!host.CSS?.supports('page', 'p0')) + '\nimg:last-child { break-after:auto }'
    inner.head.append(style)
    const pageNames = new Map(sizes(specs).map((s, n) => [`${s.widthPx},${s.heightPx}`, `p${n}`]))
    let blobBytes = 0
    for (const page of pages) {
      if (disposed || opts.signal.aborted) { dispose(); return null }
      canvas.width = Math.ceil(page.spec.widthPx * plan.scale)
      canvas.height = Math.ceil(page.spec.heightPx * plan.scale)
      let blob: Blob | null
      if (opts.rasterize) blob = await opts.rasterize(page, plan.scale, canvas)
      else {
        const ctx = canvas.getContext('2d')!
        ctx.scale(plan.scale, plan.scale)
        ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, page.spec.widthPx, page.spec.heightPx)
        page.paint(ctx)
        blob = await new Promise<Blob | null>(resolve => canvas.toBlob(resolve, 'image/png'))
      }
      // Reset state and pixels even if a signal-ignoring rasterizer refilled it.
      canvas.width = canvas.height = 0
      if (disposed || opts.signal.aborted) { dispose(); return null }
      if (!blob) throw new Error('Page rasterization failed')
      blobBytes += blob.size
      if (blobBytes > PRINT_JOB_MAX_BLOB_BYTES) { dispose(); opts.onNotice(TOO_LARGE); return null }
      const url = URL.createObjectURL(blob)
      urls.push(url)
      const img = inner.createElement('img')
      img.dataset.page = pageNames.get(`${page.spec.widthPx},${page.spec.heightPx}`)!
      await new Promise<void>((resolve, reject) => {
        const done = (error?: Error) => {
          img.removeEventListener('load', loaded); img.removeEventListener('error', failed)
          stopWaiting = undefined
          error ? reject(error) : resolve()
        }
        const loaded = () => done(), failed = () => done(new Error('Page image failed to load'))
        stopWaiting = () => done()
        img.addEventListener('load', loaded, {once: true}); img.addEventListener('error', failed, {once: true})
        img.src = url
        inner.body.append(img)
      })
      if (disposed || opts.signal.aborted) { dispose(); return null }
    }
    win.print()
    return job
  } catch {
    const cancelled = disposed || opts.signal.aborted
    dispose()
    if (!cancelled) opts.onNotice('The document could not be printed. Please try again.')
    return null
  } finally {
    canvas.width = canvas.height = 0
  }
}
