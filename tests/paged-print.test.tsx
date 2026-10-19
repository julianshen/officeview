import { describe, expect, it } from 'vitest'
import { pageRules, planPrintJob } from '../src/components/paged/printDocument'
const letter = { widthPx: 816, heightPx: 1056 }
describe('print planning', () => {
  it('chooses a uniform scale before allocation and refuses oversized jobs', () => {
    expect(planPrintJob(Array(20).fill(letter))).toEqual({ kind: 'ok', scale: 2 })
    expect(planPrintJob(Array(200).fill(letter))).toEqual({ kind: 'ok', scale: 1 })
    expect(planPrintJob(Array(400).fill(letter))).toEqual({ kind: 'refused', reason: 'too-large' })
    expect(planPrintJob([letter, { widthPx: 9000, heightPx: 10 }])).toEqual({ kind: 'ok', scale: 1 })
    expect(planPrintJob([{ widthPx: NaN, heightPx: 10 }])).toEqual({ kind: 'refused', reason: 'too-large' })
  })
  it('emits named sizes once and a bounded single-size fallback', () => {
    const specs = [letter, letter, { widthPx: 600, heightPx: 800 }]
    const named = pageRules(specs, true)
    expect(named.match(/@page /g)).toHaveLength(2)
    expect(named).toContain('size: 816px 1056px; margin: 0')
    expect(named).toContain('size: 600px 800px; margin: 0')
    const fallback = pageRules(specs, false)
    expect(fallback.match(/@page /g)).toHaveLength(1)
    expect(fallback).toContain('size: 816px 1056px; margin: 0')
    expect(fallback).toContain('object-fit: contain')
  })
})

import { afterEach, beforeEach, vi } from 'vitest'
import { printDocument } from '../src/components/paged/printDocument'
import type { Paintable } from '../src/render/paint'
const pages: Paintable[] = Array.from({length: 3}, () => ({spec: letter, paint: vi.fn()}))
let root: HTMLDivElement
let abort: AbortController
let notice: ReturnType<typeof vi.fn<(text: string) => void>>
let framePrint: ReturnType<typeof vi.fn<() => void>>
let urls: ReturnType<typeof vi.fn>
let revoke: ReturnType<typeof vi.fn>
let observer: MutationObserver
let autoLoad = true
beforeEach(() => {
  root = document.createElement('div'); document.body.append(root)
  abort = new AbortController(); notice = vi.fn(); framePrint = vi.fn(); autoLoad = true
  urls = vi.fn(() => `blob:test-${urls.mock.calls.length}`); revoke = vi.fn()
  vi.stubGlobal('URL', Object.assign(class extends URL {}, {createObjectURL: urls, revokeObjectURL: revoke}))
  vi.spyOn(window, 'print').mockImplementation(() => {})
  const append = Node.prototype.appendChild
  vi.spyOn(Node.prototype, 'appendChild').mockImplementation(function<T extends Node>(this: Node, node: T): T {
    const result = append.call(this, node) as T
    if (node instanceof HTMLIFrameElement) {
      node.contentWindow!.print = framePrint
      observer = new MutationObserver(records => {
        for (const r of records) for (const child of r.addedNodes) {
          if (child.nodeName === 'IMG' && autoLoad) child.dispatchEvent(new Event('load'))
        }
      })
      observer.observe(node.contentDocument!, {childList: true, subtree: true})
    }
    return result
  })
})
afterEach(() => { abort.abort(); observer?.disconnect(); root.remove(); vi.restoreAllMocks(); vi.unstubAllGlobals() })
const blob = () => Promise.resolve(new Blob(['png'], {type: 'image/png'}))
const start = (rasterize: NonNullable<Parameters<typeof printDocument>[2]['rasterize']> = vi.fn(blob), ps = pages) => printDocument(ps, root, {signal: abort.signal, onNotice: notice, rasterize})
async function until(check: () => boolean) { await vi.waitFor(() => expect(check()).toBe(true)) }

describe('isolated print ownership', () => {
  it('prints only page images from the frame and disposes exactly once', async () => {
    root.append(document.createTextNode('HOST-SECRET'))
    const job = await start()
    expect(job).not.toBeNull()
    expect(job!.frame.contentDocument!.body.querySelectorAll('img')).toHaveLength(3)
    expect(job!.frame.contentDocument!.body.textContent).not.toContain('HOST-SECRET')
    expect(job!.frame.getAttribute('aria-hidden')).toBe('true')
    expect(job!.frame.tabIndex).toBe(-1)
    expect(framePrint).toHaveBeenCalledTimes(1); expect(window.print).not.toHaveBeenCalled()
    job!.dispose(); job!.dispose()
    expect(root.querySelector('iframe')).toBeNull(); expect(revoke).toHaveBeenCalledTimes(3)
  })
  it('paints the detached workspace with the parent viewer language', async () => {
    root.lang = 'zh-TW'
    const raster = vi.fn(async (_p: Paintable, _s: 1 | 2, canvas: HTMLCanvasElement) => {
      expect(canvas.lang).toBe('zh-TW')
      return blob()
    })
    expect(await start(raster)).not.toBeNull()
    expect(notice).not.toHaveBeenCalled()
  })
  it('refuses before allocating anything', async () => {
    const raster = vi.fn(blob)
    expect(await start(raster, Array(400).fill(pages[0]))).toBeNull()
    expect(raster).not.toHaveBeenCalled(); expect(root.querySelector('iframe')).toBeNull()
    expect(notice).toHaveBeenCalledWith('This document is too large to print from the viewer.')
  })
  it('reuses one workspace serially, clears per page and zeros at completion', async () => {
    const canvases: HTMLCanvasElement[] = []
    let previous: HTMLCanvasElement | undefined
    const raster = vi.fn(async (_p, _s, c: HTMLCanvasElement) => {
      if (previous) expect(c.getContext('2d')!.getImageData(0,0,1,1).data[3]).toBe(0)
      c.getContext('2d')!.fillRect(0,0,1,1)
      previous = c; canvases.push(c)
      return blob()
    })
    await start(raster)
    expect(new Set(canvases).size).toBe(1)
    expect(canvases[0].width).toBe(0); expect(canvases[0].height).toBe(0)
  })
  it('refuses a blob over the running job cap without publishing it', async () => {
    const raster = vi.fn(async () => ({size: 65 * 1024 * 1024}) as Blob)
    expect(await start(raster)).toBeNull()
    expect(urls).toHaveBeenCalledTimes(2); expect(revoke).toHaveBeenCalledTimes(2)
    expect(framePrint).not.toHaveBeenCalled(); expect(root.querySelector('iframe')).toBeNull()
    expect(notice).toHaveBeenCalledWith('This document is too large to print from the viewer.')
  })
  it('never publishes a late blob after abort', async () => {
    let finish!: (blob: Blob) => void
    const raster = vi.fn(() => new Promise<Blob>(r => {finish = r}))
    const pending = start(raster)
    await until(() => raster.mock.calls.length === 1)
    abort.abort(); expect(root.querySelector('iframe')).toBeNull()
    finish(new Blob(['late'])); expect(await pending).toBeNull()
    expect(urls).not.toHaveBeenCalled(); expect(framePrint).not.toHaveBeenCalled()
  })
  it('abort during image load revokes immediately and settles', async () => {
    autoLoad = false
    const pending = start()
    await until(() => urls.mock.calls.length === 1)
    abort.abort()
    expect(await pending).toBeNull(); expect(revoke).toHaveBeenCalledTimes(1)
    expect(framePrint).not.toHaveBeenCalled()
  })
  it.each(['paint', 'image'])('cleans up on %s failure', async failure => {
    autoLoad = false
    const pending = start(vi.fn(failure === 'paint' ? async () => {throw new Error('paint failed')} : blob))
    if (failure === 'image') {
      await until(() => urls.mock.calls.length === 1)
      root.querySelector('iframe')!.contentDocument!.querySelector('img')!.dispatchEvent(new Event('error'))
    }
    expect(await pending).toBeNull(); expect(root.querySelector('iframe')).toBeNull()
    expect(notice).toHaveBeenCalledTimes(1); expect(revoke).toHaveBeenCalledTimes(failure === 'image' ? 1 : 0)
  })
  it('afterprint while preparing cancels the pending image load', async () => {
    autoLoad = false
    const pending = start()
    await until(() => urls.mock.calls.length === 1)
    window.dispatchEvent(new Event('afterprint'))
    expect(root.querySelector('iframe')).toBeNull()
    expect(await pending).toBeNull()
    expect(revoke).toHaveBeenCalledTimes(1)
    expect(framePrint).not.toHaveBeenCalled()
  })
  it.each(['frame', 'parent'])('cleans up afterprint from %s', async source => {
    const job = await start()
    ;(source === 'frame' ? job!.frame.contentWindow! : window).dispatchEvent(new Event('afterprint'))
    expect(root.querySelector('iframe')).toBeNull(); expect(revoke).toHaveBeenCalledTimes(3)
  })
})

import { render, fireEvent, waitFor, cleanup, act } from '@testing-library/react'
import { OfficeDoc } from '../src/components/OfficeDoc'
import { docxModel, fakePaintables } from './helpers/paged-viewer'
import * as paint from '../src/render/paint'
import { printLoader } from '../src/components/paged/PagedChrome'

describe('mounted print lifecycle', () => {
  beforeEach(() => {
    vi.spyOn(paint, 'getPaintables').mockImplementation(async () => fakePaintables(Array(3).fill({w:100,h:100})))
    vi.spyOn(HTMLCanvasElement.prototype, 'toBlob').mockImplementation(cb => cb(new Blob(['png'])))
  })
  afterEach(cleanup)
  it.each(['new print', 'source', 'permission', 'unmount'])('releases images on %s', async reason => {
    const doc = docxModel(['one'])
    const view = render(<OfficeDoc document={doc} />, {container:root})
    fireEvent.keyDown(view.getByRole('region'), {key:'p', ctrlKey:true})
    await waitFor(() => expect(framePrint).toHaveBeenCalledTimes(1))
    const first = root.querySelector('iframe')!
    if (reason === 'new print') {
      fireEvent.click(view.getByRole('button', {name:'Print'}))
      await waitFor(() => expect(framePrint).toHaveBeenCalledTimes(2))
    } else if (reason === 'source') view.rerender(<OfficeDoc document={docxModel(['two'])} />)
    else if (reason === 'permission') view.rerender(<OfficeDoc document={doc} allowPrint={false} />)
    else view.unmount()
    await waitFor(() => expect(first.isConnected).toBe(false))
    expect(revoke).toHaveBeenCalledTimes(3)
    expect(window.print).not.toHaveBeenCalled()
  })
  it('prints with the toolbar hidden and captures shortcuts from toolbar controls', async () => {
    const doc = docxModel(['one'])
    const view = render(<OfficeDoc document={doc} showToolbar={false} />, {container:root})
    fireEvent.keyDown(view.getByRole('region'), {key:'p', metaKey:true})
    await waitFor(() => expect(framePrint).toHaveBeenCalledTimes(1))
    expect(view.queryByRole('toolbar')).toBeNull()
    view.rerender(<OfficeDoc document={doc} />)
    const button = await view.findByRole('button', {name:'Print'})
    fireEvent.keyDown(button, {key:'p', ctrlKey:true})
    await waitFor(() => expect(framePrint).toHaveBeenCalledTimes(2))
  })
  it.each(['permission', 'source', 'unmount'])('does not create a frame when an import completes after %s', async reason => {
    let release!: (m: typeof import('../src/components/paged/printDocument')) => void
    vi.spyOn(printLoader, 'load').mockImplementation(() => new Promise(r => {release = r}))
    const doc = docxModel(['one'])
    const view = render(<OfficeDoc document={doc} />, {container:root})
    fireEvent.click(await view.findByRole('button', {name:'Print'}))
    await waitFor(() => expect(printLoader.load).toHaveBeenCalled())
    if (reason === 'permission') view.rerender(<OfficeDoc document={doc} allowPrint={false} />)
    else if (reason === 'source') view.rerender(<OfficeDoc document={docxModel(['replacement'])} />)
    else view.unmount()
    await act(async () => release(await import('../src/components/paged/printDocument')))
    expect(root.querySelector('iframe')).toBeNull(); expect(urls).not.toHaveBeenCalled()
  })
  it('permission restoration does not replay a cancelled keyboard print', async () => {
    const doc = docxModel(['one'])
    const view = render(<OfficeDoc document={doc} />, {container:root})
    fireEvent.keyDown(view.getByRole('region'), {key:'p', ctrlKey:true})
    await waitFor(() => expect(framePrint).toHaveBeenCalledTimes(1))
    view.rerender(<OfficeDoc document={doc} allowPrint={false} />)
    view.rerender(<OfficeDoc document={doc} />)
    await act(async () => { await new Promise(r => setTimeout(r, 20)) })
    expect(framePrint).toHaveBeenCalledTimes(1)
    expect(root.querySelector('iframe')).toBeNull()
  })
  it('denied print hides only its own instance and swallows Ctrl-P', async () => {
    const doc = docxModel(['one'])
    const view = render(<><OfficeDoc document={doc} allowPrint={false}/><OfficeDoc document={doc}/></>, {container:root})
    const regions = view.getAllByRole('region')
    expect(fireEvent.keyDown(regions[0], {key:'p', ctrlKey:true})).toBe(false)
    await waitFor(() => expect(view.getAllByRole('button', {name:'Print'})).toHaveLength(1))
    const rules = Array.from(document.styleSheets).flatMap(sheet => Array.from(sheet.cssRules)).flatMap(rule =>
      rule instanceof CSSMediaRule && rule.media.mediaText.includes('print') ? Array.from(rule.cssRules) : [])
    const rule = rules.find(rule => rule instanceof CSSStyleRule && rule.selectorText.includes('data-ov-print')) as CSSStyleRule
    expect(rule.style.display).toBe('none')
    expect(rule.style.getPropertyPriority('display')).toBe('important')
    expect(regions[0].closest(rule.selectorText)).not.toBeNull()
    expect(regions[1].closest(rule.selectorText)).toBeNull()
    expect(framePrint).not.toHaveBeenCalled()
  })
})
