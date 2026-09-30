/**
 * Regressions for bugs the real-file corpus surfaced. Each test reproduces a
 * fixture shape that the synthetic builders never emitted.
 *
 *   bun scripts/fetch-corpus.ts   # required to run the corpus suite
 */
import { describe, test, expect } from 'vitest'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { layoutDocx } from '../src/docx/layout'
import { parsePptx } from '../src/pptx/parse'
import { getPaintables } from '../src/render/paint'
import { buildTextIndex } from '../src/core/search'

const CORPUS_DIR = join(__dirname, '..', 'corpus')
const corpusAvailable = existsSync(CORPUS_DIR)
const corpus = (name: string): Uint8Array => readFileSync(join(CORPUS_DIR, name))

const measureFixed = (text: string, style: { fontSizePt: number }): number =>
  text.length * style.fontSizePt * 0.6 * (96 / 72)

describe.skipIf(!corpusAvailable)('corpus regressions', () => {
  test('SDT-wrapped table rows are unwrapped (Bug66263-table.docx)', async () => {
    const doc = await parseDocx(await OfficePackage.load(corpus('poi-Bug66263-table.docx')))
    // the bug: a body whose only content is a table produced zero sections
    expect(doc.sections.length).toBe(1)
    const tables = doc.sections[0].blocks.filter((b) => b.kind === 'table')
    expect(tables.length).toBe(1)
    const table = (tables[0] as { kind: 'table'; table: { rows: unknown[]; gridColsTwips: number[] } }).table
    // rows live inside w:sdt > w:sdtContent and must still be collected
    expect(table.rows.length).toBeGreaterThan(0)
    const text = doc.sections[0].blocks
      .filter((b) => b.kind === 'table')
      .flatMap((b) => (b as { kind: 'table'; table: { rows: Array<{ cells: Array<{ paragraphs: Array<{ runs: Array<{ text: string }> }> }> }> } }).table.rows)
      .flatMap((r) => r.cells.flatMap((c) => c.paragraphs.map((p) => p.runs.map((run) => run.text).join(''))))
      .join('')
    expect(text).toContain('SDT Cell 1')
  })

  test('a table with no w:tblGrid still lays out and paints its text', async () => {
    const doc = await parseDocx(await OfficePackage.load(corpus('poi-Bug66263-table.docx')))
    const paintables = await getPaintables(doc as never)
    expect(paintables.length).toBeGreaterThan(0)
    const index = await buildTextIndex(paintables as never)
    const painted = index.pages.flatMap((p) => p.lines.map((l) => l.text)).join('')
    expect(painted).toContain('SDT Cell 1')
  })

  test('placeholder shapes inherit geometry from the slide layout (python-pptx)', async () => {
    const doc = await parsePptx(await OfficePackage.load(corpus('pypptx-test.pptx')))
    const shapes = doc.slides[0].shapes
    // these shapes carry only p:ph — geometry lives in the layout
    const placeholders = shapes.filter((s) => s.placeholder)
    expect(placeholders.length).toBeGreaterThan(0)
    for (const ph of placeholders) {
      expect(ph.widthEmu, 'inherited width').toBeGreaterThan(0)
      expect(ph.heightEmu, 'inherited height').toBeGreaterThan(0)
    }
    // and their text actually reaches the canvas
    const paintables = await getPaintables(doc as never)
    const index = await buildTextIndex(paintables as never)
    const painted = index.pages.flatMap((p) => p.lines.map((l) => l.text)).join('')
    expect(painted.trim().length).toBeGreaterThan(0)
  })

  test('every corpus file paints all of its model text', async () => {
    // guards the class of bug "parsed fine, rendered nothing"
    const { readdirSync } = await import('node:fs')
    const files = readdirSync(CORPUS_DIR).filter((f) => /\.(docx|xlsx|pptx)$/i.test(f))
    const problems: string[] = []
    for (const file of files) {
      // malformed-on-purpose fixtures must be refused, not rendered
      if (file.startsWith('poi-crash-')) continue
      try {
        const { loadOfficeFile } = await import('../src/components/OfficeFile')
        const doc = await loadOfficeFile(corpus(file))
        const paintables = await getPaintables(doc as never)
        const index = await buildTextIndex(paintables as never)
        const painted = index.pages.flatMap((p) => p.lines.map((l) => l.text)).join('').trim().length
        if (paintables.length === 0 && painted === 0) {
          // legitimately empty fixture (e.g. a deck with no slides)
          const minimal = file === 'pypptx-minimal.pptx'
          if (!minimal) problems.push(`${file}: rendered nothing`)
        }
      } catch (e) {
        problems.push(`${file}: threw ${e instanceof Error ? e.message : String(e)}`)
      }
    }
    expect(problems).toEqual([])
  }, 120_000)
})

describe('document with only a table produces a section', () => {
  test('synthetic: table-only body is not dropped', async () => {
    const { buildDocx } = await import('../src/testdata/ooxml-builders')
    const doc = await parseDocx(
      await OfficePackage.load(
        await buildDocx([], [
          {
            gridCols: ['2880', '2880'],
            rows: [{ cells: [{ paragraphs: [{ runs: [{ text: 'only' }] }] }, { paragraphs: [{ runs: [{ text: 'cells' }] }] }] }],
          },
        ]),
      ),
    )
    expect(doc.sections).toHaveLength(1)
    expect(layoutDocx(doc, measureFixed as never).length).toBeGreaterThan(0)
  })
})
