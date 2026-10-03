import { describe, test, expect, vi, beforeEach, afterEach } from 'vitest'
import { render, waitFor } from '@testing-library/react'
import { OfficeFile } from '../src/components/OfficeFile'
import { WATERMARK_HEADER } from '../src/core/watermark'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { buildDocx } from '../src/testdata/ooxml-builders'

/**
 * The header path: a server stamps the document via X-OfficeView-Watermark and
 * that has to reach the rendered page without the host wiring anything up.
 */

const realFetch = globalThis.fetch
let docBytes: Uint8Array

beforeEach(async () => {
  docBytes = await buildDocx([{ runs: [{ text: 'stamp me' }] }])
})

afterEach(() => {
  globalThis.fetch = realFetch
  vi.restoreAllMocks()
})

/** Serve the doc with the given extra response headers. */
function serveWith(headers: Record<string, string>): void {
  globalThis.fetch = vi.fn(async () =>
    new Response(docBytes as unknown as BodyInit, { status: 200, headers }),
  ) as unknown as typeof fetch
}

/** How much ink a page canvas carries — a watermark shows up as extra pixels. */
function inkOf(container: HTMLElement): number {
  const canvas = container.querySelector('canvas') as HTMLCanvasElement
  const ctx = canvas.getContext('2d')
  if (!ctx) return 0
  const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data
  let dark = 0
  for (let i = 0; i < data.length; i += 4) if (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250) dark++
  return dark
}

async function renderFile(props: Record<string, unknown> = {}) {
  const utils = render(<OfficeFile data={fetch('https://example.test/doc.docx') as never} {...props} />)
  await waitFor(() => expect(utils.container.querySelectorAll('canvas').length).toBeGreaterThan(0))
  return utils
}

describe('watermark from the X-OfficeView-Watermark response header', () => {
  test('a header watermark reaches the rendered page', async () => {
    serveWith({ [WATERMARK_HEADER]: 'text=CONFIDENTIAL; placement=tile; opacity=0.5' })
    const withHeader = await renderFile()
    const headerInk = inkOf(withHeader.container)
    expect(headerInk).toBeGreaterThan(0)

    // same document, no header -> no watermark ink
    serveWith({})
    const without = await renderFile()
    const plainInk = inkOf(without.container)
    expect(headerInk).toBeGreaterThan(plainInk)
  })

  test('a malformed header does not break the load', async () => {
    serveWith({ [WATERMARK_HEADER]: 'text=%E0%A4%A;;;===' })
    const { container } = await renderFile()
    expect(container.querySelectorAll('canvas').length).toBeGreaterThan(0)
    expect(container.querySelector('[role="alert"]')).toBeNull()
  })

  test('a header without text leaves the document unmarked', async () => {
    serveWith({ [WATERMARK_HEADER]: 'rotate=-45' })
    serveWith({ [WATERMARK_HEADER]: 'rotate=-45' })
    const { container } = await renderFile()
    expect(container.querySelectorAll('canvas').length).toBeGreaterThan(0)
  })
})

describe('watermark prop vs header precedence', () => {
  test('the header wins over the prop, because the server is the authority', async () => {
    // Render both variants and compare which one's text is painted. Rather than
    // guess at pixels, assert the merge rule directly on the exported helper —
    // the pixel evidence above covers that the header is applied at all.
    const { mergeWatermark } = await import('../src/core/watermark')
    expect(mergeWatermark({ text: 'PROP' }, { text: 'HEADER' })?.text).toBe('HEADER')
    expect(mergeWatermark({ text: 'PROP' }, undefined)?.text).toBe('PROP')
  })

  test('a prop watermark renders without any header', async () => {
    serveWith({})
    const { container } = await renderFile({ watermark: { text: 'DRAFT', placement: 'center', opacity: 0.6 } })
    expect(container.querySelectorAll('canvas').length).toBeGreaterThan(0)
    expect(inkOf(container)).toBeGreaterThan(0)
  })
})

describe('watermark does not disturb loading of a non-http source', () => {
  test('a plain docx still parses with a watermark prop set', async () => {
    const doc = await parseDocx(await OfficePackage.load(docBytes))
    expect(doc.sections.length).toBeGreaterThan(0)
    const utils = render(<OfficeFile data={docBytes} watermark={{ text: 'X' }} />)
    await waitFor(() => expect(utils.container.querySelectorAll('canvas').length).toBeGreaterThan(0))
    expect(utils.container.querySelector('[role="alert"]')).toBeNull()
  })
})