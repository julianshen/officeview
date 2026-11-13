import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { PRINT_JOB_MAX_BLOB_BYTES, printRasters } from '../../src/components/paged/printDocument'

const TOO_LARGE = 'This document is too large to print from the viewer.'
const FAILED = 'The document could not be printed. Please try again.'
const letter = {widthPx: 816, heightPx: 1056}
const a5 = {widthPx: 600, heightPx: 800}
const png = (size = 3) => {
  const blob = new Blob(['png'], {type: 'image/png'})
  return size === 3 ? blob : Object.defineProperty(blob, 'size', {value: size})
}

let root: HTMLDivElement
let abort: AbortController
let notice: ReturnType<typeof vi.fn<(text: string) => void>>
let framePrint: ReturnType<typeof vi.fn<() => void>>
let revoke: ReturnType<typeof vi.fn>
let observers: MutationObserver[] = []
let imageEvent: 'load' | 'error' | null = 'load'
const opts = () => ({signal: abort.signal, onNotice: notice})

beforeEach(() => {
  root = document.createElement('div'); document.body.append(root)
  abort = new AbortController(); notice = vi.fn(); framePrint = vi.fn(); imageEvent = 'load'
  let n = 0
  revoke = vi.fn()
  vi.stubGlobal('URL', Object.assign(class extends URL {}, {createObjectURL: () => `blob:raster-${++n}`, revokeObjectURL: revoke}))
  vi.stubGlobal('CSS', {supports: () => true})
  const append = Node.prototype.appendChild
  vi.spyOn(Node.prototype, 'appendChild').mockImplementation(function<T extends Node>(this: Node, node: T): T {
    const result = append.call(this, node) as T
    if (node instanceof HTMLIFrameElement) {
      node.contentWindow!.print = framePrint
      const observer = new MutationObserver(records => {
        for (const r of records) for (const child of r.addedNodes) {
          if (child.nodeName === 'IMG' && imageEvent) child.dispatchEvent(new Event(imageEvent))
        }
      })
      observers.push(observer)
      observer.observe(node.contentDocument!, {childList: true, subtree: true})
    }
    return result
  })
})
afterEach(() => {
  abort.abort(); observers.forEach(o => o.disconnect()); observers = []
  root.remove(); vi.restoreAllMocks(); vi.unstubAllGlobals()
})

describe('printRasters', () => {
  it('builds one img per raster with named-page rules for each size, then prints once', async () => {
    const job = await printRasters([{...letter, png: png()}, {...a5, png: png()}], root, {...opts(), title: 'Deck'})
    expect(job).not.toBeNull()
    const doc = job!.frame.contentDocument!
    expect(job!.frame.hasAttribute('data-officeview-print-frame')).toBe(true)
    expect(doc.title).toBe('Deck')
    const imgs = [...doc.querySelectorAll<HTMLImageElement>('img[data-page]')]
    expect(imgs.map(i => i.dataset.page)).toEqual(['p0', 'p1'])
    const css = doc.querySelector('style')!.textContent!
    expect(css).toContain('@page p0 { size: 816px 1056px; margin: 0 }')
    expect(css).toContain('@page p1 { size: 600px 800px; margin: 0 }')
    expect(framePrint).toHaveBeenCalledTimes(1)
    expect(notice).not.toHaveBeenCalled()
  })

  it('reuses a page name for equal sizes', async () => {
    const job = await printRasters([{...letter, png: png()}, {...a5, png: png()}, {...letter, png: png()}], root, opts())
    expect([...job!.frame.contentDocument!.querySelectorAll<HTMLImageElement>('img')].map(i => i.dataset.page)).toEqual(['p0', 'p1', 'p0'])
  })

  it('returns null without a frame for an empty list or an aborted signal', async () => {
    expect(await printRasters([], root, opts())).toBeNull()
    abort.abort()
    expect(await printRasters([{...letter, png: png()}], root, opts())).toBeNull()
    expect(root.querySelector('iframe')).toBeNull()
    expect(framePrint).not.toHaveBeenCalled()
  })

  it('aborting before the last image loads disposes the frame and revokes every object URL', async () => {
    imageEvent = null
    const pending = printRasters([{...letter, png: png()}, {...letter, png: png()}, {...letter, png: png()}], root, opts())
    await vi.waitFor(() => expect(root.querySelector('iframe')!.contentDocument!.querySelectorAll('img').length).toBe(1))
    root.querySelector('iframe')!.contentDocument!.querySelector('img')!.dispatchEvent(new Event('load'))
    await vi.waitFor(() => expect(root.querySelector('iframe')!.contentDocument!.querySelectorAll('img').length).toBe(2))
    abort.abort()
    expect(await pending).toBeNull()
    expect(root.querySelector('iframe')).toBeNull()
    expect(revoke.mock.calls.map(c => c[0]).sort()).toEqual(['blob:raster-1', 'blob:raster-2'])
    expect(framePrint).not.toHaveBeenCalled()
    expect(notice).not.toHaveBeenCalled()
  })

  it('refuses a total blob size over the ceiling with the local notice and no print', async () => {
    const half = PRINT_JOB_MAX_BLOB_BYTES / 2 + 1
    expect(await printRasters([{...letter, png: png(half)}, {...letter, png: png(half)}], root, opts())).toBeNull()
    expect(notice).toHaveBeenCalledExactlyOnceWith(TOO_LARGE)
    expect(framePrint).not.toHaveBeenCalled()
    expect(root.querySelector('iframe')).toBeNull()
    expect(revoke).toHaveBeenCalledTimes(1)
  })

  it('an image error fails the job with the retry notice and cleans up', async () => {
    imageEvent = 'error'
    expect(await printRasters([{...letter, png: png()}, {...letter, png: png()}], root, opts())).toBeNull()
    expect(notice).toHaveBeenCalledExactlyOnceWith(FAILED)
    expect(framePrint).not.toHaveBeenCalled()
    expect(root.querySelector('iframe')).toBeNull()
    expect(revoke).toHaveBeenCalledTimes(1)
  })

  it('disposes on afterprint from the frame and releases the URLs', async () => {
    const job = await printRasters([{...letter, png: png()}], root, opts())
    job!.frame.contentWindow!.dispatchEvent(new Event('afterprint'))
    expect(root.querySelector('iframe')).toBeNull()
    expect(revoke).toHaveBeenCalledWith('blob:raster-1')
  })
})
