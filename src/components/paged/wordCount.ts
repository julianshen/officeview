/**
 * Word count from the document text model, in bounded idle slices (spec
 * §6.11, R8). Lazy: part of the PagedChrome chunk.
 *
 * Policy (modelled on Word's default count; test expectations are derived
 * from this policy, not measured in Word):
 *  - a word is a maximal run of non-whitespace, non-CJK code points holding
 *    at least one letter or digit;
 *  - each Han, Hiragana or Katakana code point (and U+30FC) is one word and
 *    ends any open word; a combining mark continues a token, never starts one;
 *  - paragraph boundaries are whitespace, run boundaries are not;
 *  - counted: body, table cells (recursively), text boxes, footnotes and
 *    endnotes; not counted: headers, footers and comments.
 */
import type { DocxBlock, DocxDocument, DocxNote, DocxParagraph, DocxTable } from '../../docx/types'

export const UNIT_CHARS = 2000

export interface WordCounter { push(chunk: string): void; boundary(): void; total(): number }

const WS = /\s/u
const ALNUM = /[\p{L}\p{N}]/u
const MARK = /\p{M}/u
const CJK = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}ー]/u

/** A resumable tokenizer: a token, surrogate pair or mark spanning any
 * number of push calls is counted exactly once. */
export function createWordCounter(): WordCounter {
  let count = 0
  let inToken = false
  let tokenHasAlnum = false
  let pendingHigh = ''
  const close = () => {
    if (inToken && tokenHasAlnum) count++
    inToken = false
    tokenHasAlnum = false
  }
  const feed = (cp: string) => {
    if (WS.test(cp)) close()
    else if (CJK.test(cp)) { close(); count++ }
    else if (MARK.test(cp)) { /* continues an open token; never starts one */ }
    else {
      inToken = true
      if (!tokenHasAlnum && ALNUM.test(cp)) tokenHasAlnum = true
    }
  }
  return {
    push(chunk) {
      let text = pendingHigh + chunk
      pendingHigh = ''
      const last = text.charCodeAt(text.length - 1)
      if (last >= 0xd800 && last <= 0xdbff) {
        pendingHigh = text[text.length - 1]
        text = text.slice(0, -1)
      }
      for (const cp of text) feed(cp)
    },
    boundary() {
      if (pendingHigh) { feed(pendingHigh); pendingHigh = '' }
      close()
    },
    total() {
      return count + (inToken && tokenHasAlnum ? 1 : 0)
    },
  }
}

export function countWords(text: string): number {
  const c = createWordCounter()
  c.push(text)
  return c.total()
}

/** A traversal checkpoint carrying no text. */
export const STEP: unique symbol = Symbol('officeview.wordCount.step')
/** A paragraph boundary. */
export const PARA: unique symbol = Symbol('officeview.wordCount.para')

type Item = string | typeof PARA | typeof STEP
/** A node's own items, or a child node to descend into. */
type Visit = Generator<Item | Visit, void, undefined>

/** Cells hold paragraphs today; a `blocks` list (nested tables) is read when present. */
type CellLike = { paragraphs: DocxParagraph[]; blocks?: DocxBlock[] }
type ShapeLike = { paragraphs?: Array<{ runs: Array<{ text?: string }> }>; children?: ShapeLike[] }

function* slices(text: string): Generator<string, void, undefined> {
  for (let i = 0; i < text.length;) {
    let end = Math.min(text.length, i + UNIT_CHARS)
    const c = text.charCodeAt(end - 1)
    if (end < text.length && c >= 0xd800 && c <= 0xdbff) end--
    yield text.slice(i, end)
    i = end
  }
}

function* textNode(text: string | undefined): Visit {
  if (text) yield* slices(text)
  else yield STEP
}

function* paragraphNode(p: DocxParagraph): Visit {
  yield STEP
  const runs = p.runs
  for (let i = 0; i < runs.length; i++) yield textNode(runs[i].text)
  // a text box is its own flow: the paragraph ends before it
  yield PARA
  const images = p.images ?? []
  for (let i = 0; i < images.length; i++) {
    const drawing = images[i].drawing
    if (drawing?.kind === 'textbox') yield paragraphsNode(drawing.paragraphs)
    else if (drawing?.kind === 'diagram') yield shapesNode(drawing.shapes as ShapeLike[])
    else yield STEP
  }
}

function* paragraphsNode(paragraphs: readonly DocxParagraph[]): Visit {
  yield STEP
  for (let i = 0; i < paragraphs.length; i++) yield paragraphNode(paragraphs[i])
}

/** Diagram shape text bodies (their paragraphs), descending into groups. */
function* shapesNode(shapes: readonly ShapeLike[]): Visit {
  yield STEP
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i]
    yield STEP // every shape visit is a budgeted item, empty or not
    const paras = shape.paragraphs ?? []
    for (let j = 0; j < paras.length; j++) {
      const runs = paras[j].runs
      yield STEP
      for (let k = 0; k < runs.length; k++) yield textNode(runs[k].text)
      yield PARA
    }
    if (shape.children?.length) yield shapesNode(shape.children)
  }
}

function* tableNode(table: DocxTable): Visit {
  yield STEP
  const rows = table.rows
  for (let r = 0; r < rows.length; r++) {
    yield STEP
    const cells = rows[r].cells as CellLike[]
    for (let c = 0; c < cells.length; c++) {
      const cell = cells[c]
      yield cell.blocks ? blocksNode(cell.blocks) : paragraphsNode(cell.paragraphs)
    }
  }
}

function* blocksNode(blocks: readonly DocxBlock[]): Visit {
  yield STEP
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i]
    yield b.kind === 'p' ? paragraphNode(b.paragraph) : tableNode(b.table)
  }
}

function* notesNode(notes: readonly DocxNote[] | undefined): Visit {
  yield STEP
  if (!notes) return
  for (let i = 0; i < notes.length; i++) {
    const n = notes[i]
    if (n.type && n.type !== 'normal') { yield STEP; continue }
    yield n.blocks?.length ? blocksNode(n.blocks) : paragraphsNode(n.paragraphs ?? [])
  }
}

function* documentNode(doc: DocxDocument): Visit {
  yield STEP
  for (let i = 0; i < doc.sections.length; i++) {
    const s = doc.sections[i]
    yield s.blocks?.length ? blocksNode(s.blocks) : paragraphsNode(s.paragraphs ?? [])
  }
  yield notesNode(doc.footnotes)
  yield notesNode(doc.endnotes)
}

/** The counted text, lazily: run text in ≤ UNIT_CHARS slices (never splitting
 * a surrogate pair), PARA after each paragraph, STEP for text-free visits.
 * Iterative (an explicit stack of child iterators), so deep nesting neither
 * recurses nor runs ahead: between two yields lies one node visit or one
 * text slice. */
export function* wordSourceChunks(doc: DocxDocument): Generator<Item, void, undefined> {
  const stack: Visit[] = [documentNode(doc)]
  while (stack.length > 0) {
    const r = stack[stack.length - 1].next()
    // a finished child returns through its own item, so unwinding deep
    // nesting is spread over pulls too
    if (r.done) { stack.pop(); yield STEP; continue }
    const v = r.value
    if (typeof v === 'object') stack.push(v)
    else yield v
  }
}

function abortError(): Error {
  const e = new Error('The word count was aborted')
  e.name = 'AbortError'
  return e
}

const idle = (cb: () => void): (() => void) => {
  const g = globalThis as { requestIdleCallback?: (cb: () => void) => number; cancelIdleCallback?: (id: number) => void }
  if (g.requestIdleCallback) {
    const id = g.requestIdleCallback(cb)
    return () => g.cancelIdleCallback?.(id)
  }
  const id = setTimeout(cb, 0)
  return () => clearTimeout(id)
}

/** Counts in slices of ≤ sliceMs (overshooting by at most one item), checking
 * the signal and the clock after every item. Abort rejects with AbortError. */
export function countDocumentWords(
  doc: DocxDocument,
  opts: { signal: AbortSignal; now?: () => number; schedule?: (cb: () => void) => () => void; sliceMs?: number },
): Promise<number> {
  const { signal, now = () => performance.now(), schedule = idle, sliceMs = 8 } = opts
  return new Promise<number>((resolve, reject) => {
    const gen = wordSourceChunks(doc)
    const counter = createWordCounter()
    let cancel: (() => void) | undefined
    let settled = false
    const fail = () => {
      if (settled) return
      settled = true
      cancel?.()
      reject(abortError())
    }
    if (signal.aborted) { fail(); return }
    signal.addEventListener('abort', fail, { once: true })
    const slice = () => {
      if (settled) return
      const start = now()
      for (;;) {
        if (signal.aborted) { fail(); return }
        const r = gen.next()
        if (r.done) {
          settled = true
          signal.removeEventListener('abort', fail)
          counter.boundary()
          resolve(counter.total())
          return
        }
        const v = r.value
        if (typeof v === 'string') counter.push(v)
        else if (v === PARA) counter.boundary()
        if (signal.aborted) { fail(); return }
        if (now() - start >= sliceMs) { cancel = schedule(slice); return }
      }
    }
    cancel = schedule(slice)
  })
}
