import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { layoutDocx, renderPages, type MeasureFn } from '../src/docx/layout'
import { buildDocx, type DocxParaSpec } from '../src/testdata/ooxml-builders'

const measureFixed: MeasureFn = (text, style) => text.length * style.fontSizePt * 0.6 * (96 / 72)

const p = (text: string): DocxParaSpec => ({ runs: [{ text }] })

const MANY_LINES: DocxParaSpec[] = Array.from({ length: 90 }, (_, i) => p(`body line ${i}`))

async function parseWithHf(paras: DocxParaSpec[], hf: Parameters<typeof buildDocx>[2]) {
  return parseDocx(await OfficePackage.load(await buildDocx(paras, [], hf)))
}

describe('docx header/footer parsing', () => {
  test('reads header1.xml/footer1.xml through document rels', async () => {
    const doc = await parseWithHf([p('body')], {
      header: [{ align: 'right', runs: [{ text: 'Confidential', bold: true }] }],
      footer: [{ align: 'center', runs: [{ text: 'Page 1' }] }],
    })
    const section = doc.sections[0]
    expect(section.header).toBeDefined()
    expect(section.footer).toBeDefined()
    expect(section.header![0].align).toBe('right')
    expect(section.header![0].runs[0]).toMatchObject({ text: 'Confidential', bold: true })
    expect(section.footer![0].align).toBe('center')
  })

  test('no header/footer parts leaves the section undefined', async () => {
    const doc = await parseWithHf([p('body')], {})
    expect(doc.sections[0].header).toBeUndefined()
    expect(doc.sections[0].footer).toBeUndefined()
  })

  test('PAGE field runs are tagged, cached text preserved', async () => {
    const doc = await parseWithHf([], { footer: [{ align: 'right', runs: [{ text: '7', field: 'PAGE' }] }] })
    const run = doc.sections[0].footer![0].runs[0]
    expect(run.field).toBe('PAGE')
    expect(run.text).toBe('7') // cached result, replaced at paint time
  })

  test('field runs do not leave the instruction text in the body', async () => {
    const doc = await parseWithHf([{ runs: [{ text: 'x', field: 'PAGE' }] }], {})
    const texts = doc.sections[0].paragraphs[0].runs.map((r) => r.text).join('')
    expect(texts).not.toContain('PAGE')
  })
})

describe('docx header/footer layout and paint', () => {
  test('header/footer attach to every page of the section', async () => {
    const doc = await parseWithHf(MANY_LINES, {
      header: [{ align: 'right', runs: [{ text: 'HDR' }] }],
      footer: [{ align: 'center', runs: [{ text: 'FTR' }] }],
    })
    const pages = layoutDocx(doc, measureFixed)
    expect(pages.length).toBeGreaterThan(1)
    for (const page of pages) {
      expect(page.header).toBeDefined()
      expect(page.footer).toBeDefined()
    }
    // header sits above the top margin (720 twips = 48px), footer below the
    // content area (page 1056 - 48)
    expect(pages[0].header!.yPx).toBe(48)
    expect(pages[0].footer!.yPx).toBe(1056 - 48)
  })

  test('PAGE field resolves per page, NUMPAGES to the total', async () => {
    const doc = await parseWithHf(MANY_LINES, {
      footer: [{
        align: 'right',
        runs: [
          { text: 'x', field: 'PAGE' },
          { text: ' / ' },
          { text: '0', field: 'NUMPAGES' },
        ],
      }],
    })
    const pages = layoutDocx(doc, measureFixed)
    const { createCanvas } = await import('canvas')

    const textsFor = (index: number): string => {
      const page = pages[index]
      const canvas = createCanvas(Math.ceil(page.widthPx), Math.ceil(page.heightPx))
      const ctx = canvas.getContext('2d')!
      ctx.fillStyle = '#fff'
      ctx.fillRect(0, 0, canvas.width, canvas.height)
      // record fillText calls while forwarding everything else to the real ctx
      const calls: string[] = []
      const spy = new Proxy(ctx as unknown as object, {
        get(target, prop) {
          if (prop === 'fillText') return (t: string) => { calls.push(String(t)) }
          const value = (target as Record<string | symbol, unknown>)[prop]
          return typeof value === 'function' ? value.bind(target) : value
        },
        // assign straight to the target: the default trap would pass the proxy
        // as the receiver, which node-canvas rejects ("Invalid argument")
        set(target, prop, value) {
          ;(target as Record<string | symbol, unknown>)[prop] = value
          return true
        },
      }) as unknown as CanvasRenderingContext2D
      renderPages([page], spy, undefined, { pageNumberStart: index + 1, totalPages: pages.length })
      return calls.join('')
    }

    const first = textsFor(0)
    const second = textsFor(1)
    expect(first).toContain('1')
    expect(second).toContain('2')
    // the literal " / " between the fields survives
    expect(first).toContain('/')
  })
})

describe('per-page canvas field numbering', () => {
  test('paintables resolve PAGE/NUMPAGES against the whole document', async () => {
    const { getPaintables } = await import('../src/render/paint')
    const { createCanvas } = await import('canvas')
    const doc = await parseWithHf(MANY_LINES, {
      footer: [{
        align: 'center',
        runs: [{ text: 'x', field: 'PAGE' }, { text: ' of ' }, { text: '0', field: 'NUMPAGES' }],
      }],
    })
    const paintables = await getPaintables(doc as never)
    expect(paintables.length).toBeGreaterThan(1)

    const paintText = (index: number): string => {
      const { spec, paint } = paintables[index]
      const canvas = createCanvas(Math.ceil(spec.widthPx), Math.ceil(spec.heightPx))
      const ctx = canvas.getContext('2d')!
      const calls: string[] = []
      const spy = new Proxy(ctx as unknown as object, {
        get(target, prop) {
          if (prop === 'fillText') return (t: string) => { calls.push(String(t)) }
          const value = (target as Record<string | symbol, unknown>)[prop]
          return typeof value === 'function' ? value.bind(target) : value
        },
        set(target, prop, value) {
          ;(target as Record<string | symbol, unknown>)[prop] = value
          return true
        },
      }) as unknown as CanvasRenderingContext2D
      paint(spy)
      return calls.join('')
    }

    const total = paintables.length
    expect(paintText(0)).toContain('1')
    expect(paintText(1)).toContain('2')
    // every page knows the document total
    for (let i = 0; i < total; i++) expect(paintText(i)).toContain(String(total))
  })
})
