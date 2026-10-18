/**
 * Round-4 follow-up 2: an unmounted viewer is collectable after GC with no
 * further host render or event. Runs the PRODUCTION React build (what hosts
 * ship): development builds keep owner stacks (_debugStack) on the last event
 * target's fiber, a dev-only retainer outside the viewer's control.
 *
 * jsdom has no ResizeObserver and no rendering steps, so the Chrome retainer
 * (a pane-size update queued for the removed viewer between the removing
 * commit and its passive cleanups) is modelled below; the observer assertion
 * is what fails without the fix. The release suite checks the same in Chrome.
 */
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest'
import v8 from 'node:v8'
import vm from 'node:vm'

process.env.NODE_ENV = 'production'
// Sources are compiled with the dev JSX transform (jsxDEV); the production
// runtime has no jsxDEV, so map it onto the production jsx().
vi.mock('react/jsx-dev-runtime', async () => {
  const runtime = await import('react/jsx-runtime')
  return { Fragment: runtime.Fragment, jsxDEV: (type: unknown, props: unknown, key: unknown) => (runtime.jsx as (t: unknown, p: unknown, k: unknown) => unknown)(type, props, key) }
})
v8.setFlagsFromString('--expose_gc')
const gc = vm.runInNewContext('gc') as () => void
const collect = async () => { for (let i = 0; i < 6; i++) { gc(); await new Promise(resolve => setTimeout(resolve, 10)) } }
const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms))
async function until<T>(probe: () => T | null | undefined, ms = 5000): Promise<T> {
  const end = Date.now() + ms
  for (;;) { const v = probe(); if (v) return v; if (Date.now() > end) throw new Error('timed out'); await sleep(5) }
}

// Force the worker-backed (paged) route for a small workbook.
vi.mock('../src/xlsx/storage/importSession', async original => {
  const actual = await original<typeof import('../src/xlsx/storage/importSession')>()
  return { ...actual, importLimitsFor: (profile: 'D' | 'M', overrides?: Parameters<typeof actual.importLimitsFor>[1]) => actual.importLimitsFor(profile, { ...overrides, object: { ...overrides?.object, rows: 1 } }) }
})

/**
 * Chrome's ResizeObserver reports an observed element that left the document
 * (its size became 0) at the next rendering step, which can come before
 * React's passive effect cleanups (scheduled in a later task) disconnect it.
 * jsdom has no ResizeObserver; this one models that ordering: removals are
 * reported from a MutationObserver microtask, before any scheduled task.
 */
class RemovalReportingResizeObserver {
  static live = new Set<RemovalReportingResizeObserver>()
  static reported = 0
  private readonly targets = new Set<Element>()
  constructor(private readonly callback: ResizeObserverCallback) { RemovalReportingResizeObserver.live.add(this) }
  observe(target: Element) { this.targets.add(target) }
  unobserve(target: Element) { this.targets.delete(target) }
  disconnect() { this.targets.clear(); RemovalReportingResizeObserver.live.delete(this) }
  report() {
    const gone = [...this.targets].filter(target => !target.isConnected)
    if (gone.length) RemovalReportingResizeObserver.reported++
    if (gone.length) this.callback(gone.map(target => ({ target, contentRect: new DOMRectReadOnly(0, 0, 0, 0) }) as unknown as ResizeObserverEntry), this as unknown as ResizeObserver)
  }
}
let removals: MutationObserver | null = null
beforeAll(() => {
  ;(globalThis as { ResizeObserver?: unknown }).ResizeObserver = RemovalReportingResizeObserver
  removals = new MutationObserver(() => { for (const ro of [...RemovalReportingResizeObserver.live]) ro.report() })
  removals.observe(document.body, { childList: true, subtree: true })
})
afterAll(() => { removals?.disconnect(); delete (globalThis as { ResizeObserver?: unknown }).ResizeObserver })

describe('unmounted viewer collection (production React)', () => {
  it('worker-backed: open, header sort click, close from a timer; with the host idle, GC alone ends every ledger charge', async () => {
    const React = await import('react')
    const { createRoot } = await import('react-dom/client')
    // The production build is loaded: its elements carry no dev-only _store.
    expect((React.createElement('div') as unknown as { _store?: unknown })._store).toBeUndefined()
    const { OfficeFile } = await import('../src/components/OfficeFile')
    const { ViewerKernel } = await import('../src/worker/viewerKernel')
    const { OfficePackage } = await import('../src/core/zip')
    const { parseXlsx } = await import('../src/xlsx/parse')
    const { xlsxBytes } = await import('./helpers/xlsx-package')
    const rows = Array.from({ length: 40 }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}" t="inlineStr"><is><t>name ${r}</t></is></c><c r="B${r + 1}"><v>${(r * 7) % 11}</v></c></row>`).join('')
    const held = await (async () => {
      const kernel = new ViewerKernel({ parse: async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)) }, { yield: async () => {} })
      const listeners = new Set<(event: MessageEvent) => void>()
      const port = {
        postMessage(message: unknown, transfer: Transferable[]) {
          const open = message as { kind: string; source?: { kind: string; blob?: Blob } }
          const blob = open.kind === 'open' && open.source?.kind === 'blob' ? open.source.blob : undefined
          const command = structuredClone(blob ? { ...open, source: { kind: 'blob' } } : message, { transfer: transfer as ArrayBuffer[] }) as { kind: string; source?: unknown }
          if (blob) command.source = { kind: 'blob', blob }
          void kernel.handle(command as never, (out: unknown, replyTransfer?: Transferable[]) => {
            const data = structuredClone(out, { transfer: replyTransfer as ArrayBuffer[] })
            for (const listener of [...listeners]) listener({ data } as MessageEvent)
          })
        },
        addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
        removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
        terminate: vi.fn(),
      }
      let found: { snapshot(): { bytes: number; byPool: Record<string, number> } } | null = null
      const container = document.createElement('div')
      document.body.appendChild(container)
      const root = createRoot(container)
      const internal = { onMainLedger: (l: typeof found) => { found = l } } as object
      // A host that closes the document from a timer or a resolved promise (not
      // a discrete event): React commits that removal with deferred passive cleanups.
      let close: () => void = () => {}
      const data = new Blob([xlsxBytes({ sheetData: rows }) as BlobPart])
      function Host() {
        const [open, setOpen] = React.useState(true)
        close = () => setOpen(false)
        return React.createElement('div', { 'data-testid': 'host' }, open ? React.createElement(OfficeFile, { data, interactiveSheets: true, viewerWorker: () => port, ...internal } as never) : 'closed')
      }
      root.render(React.createElement(Host))
      await until(() => found && found.snapshot().byPool.rowCache > 0)
      const sortButton = await until(() => container.querySelector<HTMLButtonElement>('button[aria-label="Cycle sort column B"]:not([disabled])'))
      sortButton.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      await until(() => container.querySelector('[role=columnheader][aria-sort=ascending]'))
      await sleep(20)
      setTimeout(() => close(), 0)
      await until(() => container.textContent === 'closed')
      return { ledger: found!, root, container }
    })()
    // The host stays mounted and idle: no further render, no event.
    const { ledger } = held
    await sleep(50)
    await collect()
    expect(ledger.snapshot().bytes).toBe(0)
    // No observer was still watching an element the close removed: each was
    // disconnected inside the removing commit, before the browser could report.
    expect(RemovalReportingResizeObserver.reported).toBe(0)
    held.root.unmount(); held.container.remove()
  }, 30_000)
})
