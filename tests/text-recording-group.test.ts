/**
 * RECORD_GROUP contract (Streaming Stage 3a, Task 3b): DOCX paragraphs and
 * XLSX cells announce one group key per paragraph/cell to a recording context
 * that defines the hook. The library's own recording context does not define
 * it, so local search, selection and copy are unchanged.
 */
import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import { getPaintables } from '../src/render/paint'
import { buildTextIndex, createRecordingContext, type TextSpan } from '../src/core/search'
import { RECORD_GROUP } from '../src/core/text-recording'
import { buildDocx, buildXlsx } from '../src/testdata/ooxml-builders'

type Paintable = { paint: (ctx: CanvasRenderingContext2D) => void }

const LONG = 'The quick brown fox jumps over the lazy dog and keeps running across the wide green field until the sun goes down behind the distant hills. '.repeat(3).trim()

async function docxPaintables(): Promise<Paintable[]> {
  const bytes = await buildDocx(
    [{ runs: [{ text: LONG }] }, { runs: [{ text: 'Second paragraph.' }] }],
    [{ gridCols: ['3000', '3000'], rows: [
      { cells: [{ paragraphs: [{ runs: [{ text: 'c11' }] }] }, { paragraphs: [{ runs: [{ text: 'c12' }] }] }] },
      { cells: [{ paragraphs: [{ runs: [{ text: 'c21' }] }] }, { paragraphs: [{ runs: [{ text: 'c22' }] }] }] },
    ] }],
  )
  return [...await getPaintables(await parseDocx(await OfficePackage.load(bytes)) as never)] as Paintable[]
}

async function xlsxPaintables(): Promise<Paintable[]> {
  const base = await buildXlsx([{ name: 'S', cols: '<col min="1" max="1" width="12" customWidth="1"/>', rows: [
    { r: 1, cells: [{ ref: 'A1', t: 's', v: 0, style: 3 }, { ref: 'B1', t: 's', v: 1 }, { ref: 'C1', t: 's', v: 2 }] },
  ] }], ['alpha beta gamma delta epsilon zeta', 'bee', 'sea'])
  const zip = await JSZip.loadAsync(base)
  const styles = await zip.file('xl/styles.xml')!.async('string')
  zip.file('xl/styles.xml', styles.replace('<cellXfs count="3">', '<cellXfs count="4">').replace('</cellXfs>', '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf></cellXfs>'))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file('xl/worksheets/sheet1.xml', sheet.replace('<row r="1">', '<row r="1" ht="80" customHeight="1">'))
  const doc = await parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
  return [...await getPaintables(doc as never)] as Paintable[]
}

/** Records one paint plainly (as buildTextIndex does). */
function plain(p: Paintable): TextSpan[] {
  const tiny = createCanvas(4, 4).getContext('2d') as unknown as CanvasRenderingContext2D
  const spans: TextSpan[] = []
  tiny.save(); try { p.paint(createRecordingContext(tiny, spans)) } finally { tiny.restore() }
  return spans
}

/** Records one paint through a wrapper that defines RECORD_GROUP. */
function grouped(p: Paintable): { spans: TextSpan[]; keys: Array<object | undefined>; calls: Array<object | undefined> } {
  const tiny = createCanvas(4, 4).getContext('2d') as unknown as CanvasRenderingContext2D
  const spans: TextSpan[] = [], keys: Array<object | undefined> = [], calls: Array<object | undefined> = []
  let current: object | undefined
  const keep = spans.push.bind(spans)
  spans.push = (...items: TextSpan[]) => { for (const s of items) { keep(s); keys.push(current) } return spans.length }
  const rec = createRecordingContext(tiny, spans) as unknown as Record<string | symbol, unknown>
  const hooked = new Proxy(rec, {
    get(t, prop) { return prop === RECORD_GROUP ? (key?: object) => { calls.push(key); current = key } : t[prop] },
    set(t, prop, value) { t[prop] = value; return true },
  }) as unknown as CanvasRenderingContext2D
  tiny.save(); try { p.paint(hooked) } finally { tiny.restore() }
  return { spans: [...spans], keys, calls }
}

const textOf = (spans: TextSpan[], keys: Array<object | undefined>, key: object) => spans.filter((_, i) => keys[i] === key).map((s) => s.text).join('')

describe('RECORD_GROUP', () => {
  test('the library recording context does not define the hook', () => {
    const tiny = createCanvas(4, 4).getContext('2d') as unknown as CanvasRenderingContext2D
    expect((createRecordingContext(tiny, []) as unknown as Record<symbol, unknown>)[RECORD_GROUP]).toBeUndefined()
  })

  test('local buildTextIndex output is unchanged by the hook (docx and xlsx)', async () => {
    for (const ps of [await docxPaintables(), await xlsxPaintables()]) {
      const base = plain(ps[0]!), withHook = grouped(ps[0]!)
      expect(base.length).toBeGreaterThan(0)
      expect(withHook.spans).toEqual(base)
      // Same spans, same order, no field added (a wrapped XLSX cell paints
      // through the text-body painter and carries its own logical source either way).
      for (const s of withHook.spans) expect(Object.keys(s).every((k) => ['text', 'x', 'y', 'width', 'fontSize', 'placement', 'logical'].includes(k))).toBe(true)
      const index = await buildTextIndex(ps as never)
      expect(index.pages[0]!.lines.flatMap((l) => l.spans)).toEqual(expect.arrayContaining(base))
    }
  })

  test('docx: one key per paragraph (wrapped body, body, each table cell); the group ends', async () => {
    const [page] = await docxPaintables()
    const { spans, keys, calls } = grouped(page!)
    const ink = spans.map((s, i) => ({ s, k: keys[i] })).filter(({ s }) => s.text.trim() !== '')
    for (const { k } of ink) expect(k).toBeDefined()
    const distinct = [...new Set(ink.map(({ k }) => k!))]
    expect(distinct.length).toBe(6)
    const texts = distinct.map((k) => textOf(spans, keys, k))
    expect(texts[0]!.replace(/\s+/g, '')).toBe(LONG.replace(/\s+/g, ''))
    // The long paragraph really wraps: its spans lie on more than one baseline.
    expect(new Set(spans.filter((_, i) => keys[i] === distinct[0]).map((s) => Math.round(s.y))).size).toBeGreaterThan(1)
    expect(texts.slice(1).sort()).toEqual(['Second paragraph.', 'c11', 'c12', 'c21', 'c22'].sort())
    expect(calls[calls.length - 1]).toBeUndefined()
  })

  test('xlsx: one key per cell; a wrapped cell\'s lines share it; the group ends', async () => {
    const [unit] = await xlsxPaintables()
    const { spans, keys, calls } = grouped(unit!)
    const ink = spans.map((s, i) => ({ s, k: keys[i] })).filter(({ s }) => s.text.trim() !== '')
    for (const { k } of ink) expect(k).toBeDefined()
    const distinct = [...new Set(ink.map(({ k }) => k!))]
    expect(distinct.length).toBe(3)
    const wrapped = distinct.find((k) => textOf(spans, keys, k).includes('alpha'))!
    expect(new Set(spans.filter((_, i) => keys[i] === wrapped).map((s) => Math.round(s.y))).size).toBeGreaterThan(1)
    expect(calls[calls.length - 1]).toBeUndefined()
  })
})
