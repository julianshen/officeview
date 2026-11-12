/**
 * PAINT_DECORATION contract (Streaming Stage 3a, Task 3c fix rounds 1–3): the
 * DOCX body painter brackets each stretch of consecutive decorated segments
 * on a line, and the XLSX painter each decorated cell, in one scope —
 * 'text' / 'line' alternating, then false — on a context that defines the hook. The library's own
 * recording context does not define it, and paint output is identical with
 * or without it.
 */
import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import { getPaintables } from '../src/render/paint'
import { createRecordingContext } from '../src/core/search'
import { PAINT_DECORATION } from '../src/core/text-recording'
import { buildDocx, buildXlsx } from '../src/testdata/ooxml-builders'

type Paintable = { paint: (ctx: CanvasRenderingContext2D) => void; spec: { widthPx: number; heightPx: number } }

async function docxPaintables(): Promise<Paintable[]> {
  const zip = await JSZip.loadAsync(await buildDocx([{ runs: [{ text: 'plain ' }, { text: 'typing guide' }, { text: ' and ' }, { text: 'some more' }, { text: ' end' }] }]))
  const xml = (await zip.file('word/document.xml')!.async('string'))
    .replace(/<w:r><w:rPr><\/w:rPr><w:t xml:space="preserve">typing guide/, '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">typing guide')
    .replace(/<w:r><w:rPr><\/w:rPr><w:t xml:space="preserve">some more/, '<w:r><w:rPr><w:strike/></w:rPr><w:t xml:space="preserve">some more')
  expect(xml).toContain('<w:u w:val="single"/>')
  expect(xml).toContain('<w:strike/>')
  zip.file('word/document.xml', xml)
  return [...await getPaintables(await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))) as never)] as Paintable[]
}

async function xlsxPaintables(): Promise<Paintable[]> {
  const base = await buildXlsx([{ name: 'S', cols: '<col min="1" max="2" width="30" customWidth="1"/>', rows: [
    { r: 1, cells: [{ ref: 'A1', t: 's', v: 0, style: 3 }, { ref: 'B1', t: 's', v: 1, style: 4 }] },
  ] }], ['typing guide', 'some more'])
  const zip = await JSZip.loadAsync(base)
  const styles = await zip.file('xl/styles.xml')!.async('string')
  zip.file('xl/styles.xml', styles
    .replace('<fonts count="2">', '<fonts count="4">')
    .replace('</fonts>', '<font><u/><sz val="11"/><name val="Calibri"/></font><font><strike/><sz val="11"/><name val="Calibri"/></font></fonts>')
    .replace('<cellXfs count="3">', '<cellXfs count="5">')
    .replace('</cellXfs>', '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'))
  return [...await getPaintables(await parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))) as never)] as Paintable[]
}

function surface(p: Paintable) {
  const canvas = createCanvas(Math.ceil(p.spec.widthPx), Math.ceil(p.spec.heightPx))
  return { canvas, ctx: canvas.getContext('2d') as unknown as CanvasRenderingContext2D }
}

type Phase = 'text' | 'line' | false
/** Paints through a wrapper that defines PAINT_DECORATION; logs the hook calls and the ops made in each phase. */
function hooked(p: Paintable) {
  const { canvas, ctx } = surface(p)
  const calls: Phase[] = [], during: Record<'text' | 'line', string[]> = { text: [], line: [] }
  let on: Phase = false
  const target = ctx as unknown as Record<string | symbol, unknown>
  const wrapped = new Proxy(target, {
    get(t, prop) {
      if (prop === PAINT_DECORATION) return (v: Phase) => { calls.push(v); on = v }
      const v = t[prop]
      if (typeof v !== 'function') return v
      return (...args: unknown[]) => { if (on) during[on].push(String(prop)); return (v as (...a: unknown[]) => unknown).apply(t, args) }
    },
    set(t, prop, value) { t[prop] = value; return true },
  }) as unknown as CanvasRenderingContext2D
  p.paint(wrapped)
  return { pixels: canvas.toBuffer('raw'), calls, during }
}

describe('PAINT_DECORATION', () => {
  test('the library recording context does not define the hook', () => {
    const tiny = createCanvas(4, 4).getContext('2d') as unknown as CanvasRenderingContext2D
    expect((createRecordingContext(tiny, []) as unknown as Record<symbol, unknown>)[PAINT_DECORATION]).toBeUndefined()
  })

  for (const [kind, load] of [['docx', docxPaintables], ['xlsx', xlsxPaintables]] as const) {
    test(`${kind}: each decorated stretch is one scope ('text', 'line' alternating, then false)`, async () => {
      const [p] = await load()
      const r = hooked(p!)
      // Split into scopes at false; each must be text (line text)* line … i.e. start with text, alternate, end with line.
      const scopes: Phase[][] = [[]]
      for (const v of r.calls) { if (v === false) scopes.push([]); else scopes[scopes.length - 1]!.push(v) }
      expect(r.calls[r.calls.length - 1]).toBe(false)
      expect(scopes.pop()).toEqual([])
      expect(scopes.length).toBeGreaterThanOrEqual(2) // an underlined and a struck stretch
      for (const scope of scopes) {
        expect(scope.length % 2).toBe(0)
        scope.forEach((v, i) => expect(v).toBe(i % 2 === 0 ? 'text' : 'line'))
      }
      expect(r.during.text).toContain('fillText')
      expect(r.during.line.length).toBeGreaterThan(0)
      expect(r.during.line.filter((op) => op === 'fillText' || op === 'strokeText' || op === 'drawImage')).toEqual([])
    })

    test(`${kind}: paint output is identical with and without the hook`, async () => {
      const [p] = await load()
      const { canvas, ctx } = surface(p!)
      p!.paint(ctx)
      expect(Buffer.compare(canvas.toBuffer('raw'), hooked(p!).pixels)).toBe(0)
    })
  }
})
