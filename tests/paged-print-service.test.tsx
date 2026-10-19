import { afterEach, expect, it, vi } from 'vitest'
import { cleanup, render } from '@testing-library/react'
import PagedChrome, { type PagedChromeProps } from '../src/components/paged/PagedChrome'
import * as tracking from '../src/components/paged/usePageTracking'
import * as words from '../src/components/paged/useWordCount'
import { docxModel } from './helpers/paged-viewer'

const roots: HTMLElement[] = []
const sheets: HTMLStyleElement[] = []
afterEach(() => { cleanup(); roots.forEach(r => r.remove()); roots.length = 0; vi.restoreAllMocks(); sheets.forEach(s => document.head.append(s)); sheets.length = 0 })
function fixture(sentinel = '') {
  const root = document.createElement('div')
  root.setAttribute('data-ov-theme-root', '')
  root.setAttribute('data-ov-instance', ':«quoted"\\id»')
  if (!sentinel) for (const style of document.querySelectorAll('style')) { sheets.push(style); style.remove() }
  else root.style.setProperty('--ov-sentinel', sentinel)
  const container = document.createElement('div')
  root.append(container); document.body.append(root); roots.push(root)
  const props: PagedChromeProps = {
    document: docxModel(['one two']), compact: false, statusSlot: null,
    containerRef: {current: container}, layout: {specs: [], displayWidth: 800, gap: 16},
    transform: {zoom: 1, panX: 0, panY: 0}, setTransform: vi.fn(), zoom: 1,
    onZoom: vi.fn(), onReset: vi.fn(), fitActive: true, onFit: vi.fn(), showZoom: true,
    search: null, allowPrint: false, pages: [], owner: {}, showToolbar: false, focusSearchRequest: 0,
  }
  return {root, props}
}
function guards() { return [...document.querySelectorAll('style[data-ov-print-guard]')] }
it('adds an escaped instance-only fallback without CSS and removes it on permission/unmount', () => {
  const {root, props} = fixture()
  expect(getComputedStyle(root).getPropertyValue('--ov-sentinel').trim()).toBe('')
  const sibling = document.createElement('div'); sibling.dataset.ovInstance = 'other'
  const ui = render(<PagedChrome {...props} />)
  expect(guards()).toHaveLength(1)
  const style = guards()[0] as HTMLStyleElement
  const media = style.sheet!.cssRules[0] as CSSMediaRule
  const rule = media.cssRules[0] as CSSStyleRule
  expect(media.media.mediaText).toBe('print')
  expect(root.matches(rule.selectorText)).toBe(true)
  expect(sibling.matches(rule.selectorText)).toBe(false)
  expect(rule.style.display).toBe('none')
  expect(rule.style.getPropertyPriority('display')).toBe('important')
  ui.rerender(<PagedChrome {...props} allowPrint />)
  expect(guards()).toHaveLength(0)
  ui.rerender(<PagedChrome {...props} />)
  expect(guards()).toHaveLength(1)
  ui.unmount(); expect(guards()).toHaveLength(0)
})
it('uses the stylesheet guard when the sentinel is present', () => {
  const {props} = fixture('1')
  render(<PagedChrome {...props} />)
  expect(guards()).toHaveLength(0)
})
it('never mounts page tracking or word counting while headless, including after toggling', () => {
  const track = vi.spyOn(tracking, 'usePageTracking')
  const count = vi.spyOn(words, 'useWordCount')
  const {props} = fixture('1')
  const ui = render(<PagedChrome {...props} />)
  expect(track).not.toHaveBeenCalled(); expect(count).not.toHaveBeenCalled()
  ui.rerender(<PagedChrome {...props} showToolbar />)
  expect(track).toHaveBeenCalled(); expect(count).toHaveBeenCalled()
  track.mockClear(); count.mockClear()
  ui.rerender(<PagedChrome {...props} />)
  expect(track).not.toHaveBeenCalled(); expect(count).not.toHaveBeenCalled()
})
it('does not replay a Find request made while the toolbar was hidden', () => {
  const {props} = fixture('1')
  const search: NonNullable<PagedChromeProps['search']> = {
    inputRef: {current: null}, query: '', onQuery: vi.fn(), matches: [], active: -1,
    setActive: vi.fn(), indexing: false,
  }
  const ui = render(<PagedChrome {...props} search={search} focusSearchRequest={1} />)
  ui.rerender(<PagedChrome {...props} search={search} focusSearchRequest={1} showToolbar />)
  expect(document.activeElement).not.toBe(ui.getByRole('searchbox'))
})
