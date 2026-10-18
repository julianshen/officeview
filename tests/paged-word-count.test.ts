import { describe, it, expect } from 'vitest'
import { renderHook, waitFor } from '@testing-library/react'
import {
  PARA,
  STEP,
  UNIT_CHARS,
  countDocumentWords,
  countWords,
  createWordCounter,
  wordSourceChunks,
} from '../src/components/paged/wordCount'
import { useWordCount } from '../src/components/paged/useWordCount'
import type { DocxDocument, DocxParagraph, DocxTextRun } from '../src/docx/types'
import type { OfficeDocSource } from '../src/components/OfficeDoc'

const para = (...texts: string[]): DocxParagraph => ({ runs: texts.map((text) => ({ text })), images: [], align: 'left' })
const pBlock = (p: DocxParagraph) => ({ kind: 'p' as const, paragraph: p })
const section = (blocks: unknown[], extra: Record<string, unknown> = {}) => ({
  margins: {}, pageSize: { widthTwips: 12240, heightTwips: 15840, orientation: 'portrait' }, paragraphs: [], blocks, ...extra,
})
const doc = (sections: unknown[], extra: Record<string, unknown> = {}): DocxDocument =>
  ({ sections, defaultFontFamily: 'Arial', defaultFontSizePt: 11, styleDefaults: new Map(), ...extra }) as unknown as DocxDocument

/** A fake slice clock: `now` advances 1 per call and `schedule` runs the next
 * slice synchronously-ish (a microtask), recording the clock at each slice. */
function fakeClock() {
  let t = 0
  const sliceStarts: number[] = []
  const callsPerSlice: number[] = []
  let calls = 0
  return {
    sliceStarts,
    callsPerSlice,
    now: () => { calls++; return t++ },
    schedule: (cb: () => void) => {
      callsPerSlice.push(calls)
      calls = 0
      sliceStarts.push(t)
      let cancelled = false
      queueMicrotask(() => { if (!cancelled) cb() })
      return () => { cancelled = true }
    },
    finish() { callsPerSlice.push(calls); calls = 0 },
  }
}

/** Items handled per slice: each slice reads the clock once at its start
 * and once after every item. */
const itemsPerSlice = (calls: number[]) => calls.filter((c) => c > 0).map((c) => c - 1)

describe('countWords policy', () => {
  it.each([
    ['Hello, world', 2],
    ["don't stop e-mail 3.14 2026", 5],
    ['$50 and 100%', 3],
    ['日本語のテキスト', 8],
    ['Excel表格', 3],
    ['한국어 텍스트', 2],
    ['', 0],
    ['— • —', 0],
    ['https://example.com/a-b/c.d?x=1 ok', 2],
    ['éte', 1],
    ['𠀀𠀁', 2],
  ])('%j → %i', (text, count) => {
    expect(countWords(text)).toBe(count)
  })

  it('a word split across runs counts once', () => {
    const c = createWordCounter()
    c.push('Hel')
    c.push('lo world')
    c.boundary()
    expect(c.total()).toBe(2)
  })

  it('a token crossing chunk, run and slice boundaries counts once', async () => {
    const clock = fakeClock()
    const d = doc([section([pBlock(para('ab', 'cd', 'ef gh'))])])
    expect(await countDocumentWords(d, { signal: new AbortController().signal, now: clock.now, schedule: clock.schedule, sliceMs: 1 })).toBe(2)
    expect(clock.sliceStarts.length).toBeGreaterThanOrEqual(3)
    const c = createWordCounter()
    for (const part of ['ab', 'cd', 'ef gh']) c.push(part)
    expect(c.total()).toBe(2)
  })

  it('surrogate pairs, CJK and combining marks at chunk edges', () => {
    const text = 'x𠀀́y'
    const expected = countWords(text)
    expect(expected).toBe(3)
    for (let i = 0; i <= text.length; i++) {
      const c = createWordCounter()
      c.push(text.slice(0, i))
      c.push(text.slice(i))
      c.boundary()
      expect(c.total(), `split at ${i}`).toBe(expected)
    }
  })
})

describe('document traversal', () => {
  it('counts body, tables, text boxes and notes; skips headers, footers and comments', async () => {
    const table = { kind: 'table', table: { gridColsTwips: [1], cellMargins: {}, rows: [{ cells: [{ gridSpan: 1, paragraphs: [para('cell one')] }, { gridSpan: 1, paragraphs: [para('two')] }] }] } }
    const withBox: DocxParagraph = { ...para('body'), images: [{ data: new Uint8Array(), widthEmu: 1, heightEmu: 1, drawing: { kind: 'textbox', paragraphs: [para('box words here')], fontFamily: 'A', fontSizePt: 11, insets: { left: 0, top: 0, right: 0, bottom: 0 } } } as never] }
    const d = doc([
      section([pBlock(para('Hello world')), table, pBlock(withBox)], {
        header: [para('HEADER TEXT')], footer: [para('FOOTER')], headerBlocks: [pBlock(para('HEADER'))], footerBlocks: [pBlock(para('FOOTER'))],
      }),
    ], {
      footnotes: [{ id: 1, paragraphs: [], blocks: [pBlock(para('foot note'))] }],
      endnotes: [{ id: 2, paragraphs: [], blocks: [pBlock(para('end'))] }],
      comments: [{ id: 3, text: 'comment words', paragraphs: [para('comment words')] }],
    })
    // Hello world (2) + cell one two (3) + body (1) + box words here (3) + foot note (2) + end (1)
    expect(await countDocumentWords(d, { signal: new AbortController().signal })).toBe(12)
  })

  it('never runs more than 8 ms of work per slice', async () => {
    const clock = fakeClock()
    const d = doc([section(Array.from({ length: 500 }, (_, i) => pBlock(para(`w${i}`))))])
    const total = await countDocumentWords(d, { signal: new AbortController().signal, now: clock.now, schedule: clock.schedule })
    clock.finish()
    expect(total).toBe(500)
    const items = itemsPerSlice(clock.callsPerSlice)
    expect(Math.max(...items)).toBeLessThanOrEqual(8)
    expect(items.length).toBeGreaterThan(100)
  })

  it('a giant single token (2,000,000 chars, no whitespace) is bounded and counts once', async () => {
    const giant = 'a'.repeat(2_000_000)
    const d = doc([section([pBlock(para(giant))])])
    let texts = 0
    for (const item of wordSourceChunks(d)) {
      if (typeof item === 'string') { expect(item.length).toBeLessThanOrEqual(UNIT_CHARS); texts++ }
    }
    expect(texts).toBe(1000)
    const clock = fakeClock()
    expect(await countDocumentWords(d, { signal: new AbortController().signal, now: clock.now, schedule: clock.schedule })).toBe(1)
    clock.finish()
    expect(Math.max(...itemsPerSlice(clock.callsPerSlice))).toBeLessThanOrEqual(8)
  })

  it('never splits a surrogate pair across chunks', () => {
    const text = 'a'.repeat(UNIT_CHARS - 1) + '𠀀' + 'b'
    for (const item of wordSourceChunks(doc([section([pBlock(para(text))])]))) {
      if (typeof item !== 'string') continue
      const last = item.charCodeAt(item.length - 1)
      expect(last >= 0xd800 && last <= 0xdbff).toBe(false)
    }
  })

  it('a long URL paragraph and a 100,000-char CJK paragraph are chunked and counted exactly', async () => {
    const url = 'https://example.com/a-b?c=1&d=2/'.repeat(5000)
    const cjk = '漢'.repeat(100_000)
    const signal = new AbortController().signal
    const fast = { now: () => 0, schedule: (cb: () => void) => { queueMicrotask(cb); return () => {} } }
    expect(await countDocumentWords(doc([section([pBlock(para(url))])]), { signal, ...fast })).toBe(1)
    expect(await countDocumentWords(doc([section([pBlock(para(cjk))])]), { signal, ...fast })).toBe(100_000)
  })

  it('a paragraph with 100,000 runs is traversed lazily', () => {
    const runs: DocxTextRun[] = Array.from({ length: 100_000 }, (_, i) => ({ text: `r${i} ` }))
    let maxIndex = -1
    const proxied = new Proxy(runs, {
      get(target, key, receiver) {
        if (typeof key === 'string' && /^\d+$/.test(key)) maxIndex = Math.max(maxIndex, Number(key))
        return Reflect.get(target, key, receiver)
      },
    })
    const gen = wordSourceChunks(doc([section([pBlock({ runs: proxied, images: [], align: 'left' })])]))
    let first: unknown
    for (const item of gen) { if (typeof item === 'string') { first = item; break } }
    expect(first).toBe('r0 ')
    expect(maxIndex).toBeLessThan(5)
  })

  it('100,000 empty runs before one word: every visit is a budgeted item', async () => {
    const runs: DocxTextRun[] = Array.from({ length: 100_000 }, () => ({ text: '' }))
    runs.push({ text: 'word' })
    const clock = fakeClock()
    const total = await countDocumentWords(doc([section([pBlock({ runs, images: [], align: 'left' })])]), { signal: new AbortController().signal, now: clock.now, schedule: clock.schedule })
    clock.finish()
    expect(total).toBe(1)
    const items = itemsPerSlice(clock.callsPerSlice)
    expect(Math.max(...items)).toBeLessThanOrEqual(8)
    expect(items.length).toBeGreaterThanOrEqual(12_500)
  })

  it('deep non-text structure is bounded and does not recurse', async () => {
    // 10,000 tables nested through cells, then 50,000 empty tables, then 'x'
    const emptyTable = () => ({ kind: 'table', table: { gridColsTwips: [1], cellMargins: {}, rows: [] as unknown[] } })
    let inner: unknown = emptyTable()
    for (let i = 0; i < 10_000; i++) {
      inner = { kind: 'table', table: { gridColsTwips: [1], cellMargins: {}, rows: [{ cells: [{ gridSpan: 1, paragraphs: [], blocks: [inner] }] }] } }
    }
    const blocks: unknown[] = [inner]
    for (let i = 0; i < 50_000; i++) blocks.push(emptyTable())
    blocks.push(pBlock(para('x')))
    const clock = fakeClock()
    const total = await countDocumentWords(doc([section(blocks)]), { signal: new AbortController().signal, now: clock.now, schedule: clock.schedule })
    clock.finish()
    expect(total).toBe(1)
    expect(Math.max(...itemsPerSlice(clock.callsPerSlice))).toBeLessThanOrEqual(8)
  })

  it('abort during a long text-free traversal stops within one item', async () => {
    const runs: DocxTextRun[] = Array.from({ length: 100_000 }, () => ({ text: '' }))
    let maxIndex = -1
    const proxied = new Proxy(runs, {
      get(target, key, receiver) {
        if (typeof key === 'string' && /^\d+$/.test(key)) maxIndex = Math.max(maxIndex, Number(key))
        return Reflect.get(target, key, receiver)
      },
    })
    const controller = new AbortController()
    const clock = fakeClock()
    let slices = 0
    let atAbort = -1
    const schedule = (cb: () => void) => {
      slices++
      if (slices === 6) { controller.abort(); atAbort = maxIndex }
      return clock.schedule(cb)
    }
    const run = countDocumentWords(doc([section([pBlock({ runs: proxied, images: [], align: 'left' })])]), { signal: controller.signal, now: clock.now, schedule })
    await expect(run).rejects.toMatchObject({ name: 'AbortError' })
    expect(atAbort).toBeGreaterThan(0)
    expect(maxIndex).toBeLessThanOrEqual(atAbort + 1)
  })

  it('cancels on abort: no result after document change', async () => {
    const controller = new AbortController()
    const d = doc([section(Array.from({ length: 200 }, () => pBlock(para('a b c'))))])
    const clock = fakeClock()
    let resolved = false
    const run = countDocumentWords(d, { signal: controller.signal, now: clock.now, schedule: clock.schedule }).then((n) => { resolved = true; return n })
    controller.abort()
    await expect(run).rejects.toMatchObject({ name: 'AbortError' })
    expect(resolved).toBe(false)
  })

  it('yields a PARA after each paragraph and a STEP for text-free nodes', () => {
    const items = [...wordSourceChunks(doc([section([pBlock(para('a', '')), pBlock(para())])]))]
    expect(items.filter((i) => i === PARA)).toHaveLength(2)
    expect(items.filter((i) => i === STEP).length).toBeGreaterThan(0)
    expect(items.filter((i) => typeof i === 'string')).toEqual(['a'])
  })
})

describe('useWordCount', () => {
  it('useWordCount returns null until counted, never counts during render, null for PPTX and static XLSX', async () => {
    const d = doc([section([pBlock(para('one two three'))])])
    const { result } = renderHook(({ document }) => useWordCount(document), { initialProps: { document: d as OfficeDocSource } })
    // the first render returns before any counting ran
    expect(result.current).toBeNull()
    await waitFor(() => expect(result.current).toBe(3))
    const pptx = { slides: [] } as unknown as OfficeDocSource
    const xlsx = { sheets: [] } as unknown as OfficeDocSource
    expect(renderHook(() => useWordCount(pptx)).result.current).toBeNull()
    expect(renderHook(() => useWordCount(xlsx)).result.current).toBeNull()
    await new Promise((r) => setTimeout(r, 20))
  })

  it('a document change aborts the old count and resets to null', async () => {
    const a = doc([section([pBlock(para('one two'))])])
    const b = doc([section([pBlock(para('one two three four'))])])
    const { result, rerender } = renderHook(({ document }) => useWordCount(document), { initialProps: { document: a as OfficeDocSource } })
    await waitFor(() => expect(result.current).toBe(2))
    rerender({ document: b as OfficeDocSource })
    expect(result.current).toBeNull()
    await waitFor(() => expect(result.current).toBe(4))
  })
})
