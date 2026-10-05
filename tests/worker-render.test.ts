import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { handleWorkerRequest, isWorkerRenderSupported, renderPageBitmap, type WorkerRenderRequest } from '../src/worker/office-worker'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { getPaintables } from '../src/render/paint'
import { CT_TYPES, ROOT_RELS } from '../src/testdata/ooxml-builders'

const WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
async function tinyDocx(text: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CT_TYPES)
  zip.file('_rels/.rels', ROOT_RELS)
  zip.file('word/document.xml', `<w:document ${WNS}><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`)
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
  return zip.generateAsync({ type: 'uint8array' })
}
// Node test env: real parsers + node-canvas stand-ins for OffscreenCanvas.
const nodeEnv = {
  async createCanvas(width: number, height: number) {
    return createCanvas(Math.max(1, Math.ceil(width)), Math.max(1, Math.ceil(height)))
  },
  async toBitmap(canvas: unknown) {
    return canvas as ImageBitmap
  },
  async parse(_format: 'docx', source: Uint8Array) {
    return parseDocx(await OfficePackage.load(source))
  },
  async paintables(doc: never, options: never) {
    return getPaintables(doc as never, options as never)
  },
}

describe('worker render protocol', () => {
  test('renders a docx page to a bitmap through the handler', async () => {
    const source = await tinyDocx('Hi worker')
    const req: WorkerRenderRequest = { id: 1, kind: 'render-page', format: 'docx', source: source.buffer as ArrayBuffer, index: 0 }
    const res = await handleWorkerRequest(req, nodeEnv as never)
    expect(res.ok).toBe(true)
    if (!res.ok) return
    expect(res.id).toBe(1)
    expect(res.width).toBeGreaterThan(0)
    expect(res.height).toBeGreaterThan(0)
    const canvas = res.bitmap as unknown as { getContext: (k: string) => { getImageData: (x: number, y: number, w: number, h: number) => { data: Uint8ClampedArray } } }
    const data = canvas.getContext('2d').getImageData(0, 0, res.width, res.height).data
    expect(Array.from(data).some(v => v < 250)).toBe(true)
  })
  test('unknown kind fails closed with an error', async () => {
    const res = await handleWorkerRequest({ id: 2, kind: 'nope', format: 'docx', source: new ArrayBuffer(0), index: 0 } as never, nodeEnv as never)
    expect(res.ok).toBe(false)
    if (res.ok) return
    expect(res.id).toBe(2)
    expect(typeof res.error).toBe('string')
  })
  test('out-of-range page fails closed', async () => {
    const source = await tinyDocx('Hi')
    const res = await handleWorkerRequest({ id: 3, kind: 'render-page', format: 'docx', source: source.buffer as ArrayBuffer, index: 99 }, nodeEnv as never)
    expect(res.ok).toBe(false)
  })
})

describe('worker client', () => {
  test('matches responses by id and surfaces worker errors', async () => {
    let onmessage: ((ev: { data: unknown }) => void) | undefined
    const posted: unknown[] = []
    const worker = {
      postMessage: (msg: unknown) => { posted.push(msg) },
      set onmessage(f: NonNullable<typeof onmessage>) { onmessage = f },
      get onmessage(): NonNullable<typeof onmessage> { return onmessage as NonNullable<typeof onmessage> },
      terminate: () => {},
    }
    const pending = renderPageBitmap(worker as never, { id: 7, kind: 'render-page', format: 'docx', source: new ArrayBuffer(8), index: 0 })
    expect(posted).toHaveLength(1)
    onmessage?.({ data: { id: 8, ok: true, bitmap: {}, width: 1, height: 1 } })
    onmessage?.({ data: { id: 7, ok: false, error: 'boom' } })
    await expect(pending).rejects.toThrow('boom')
  })
  test('support probe reflects environment capabilities', () => {
    expect(typeof isWorkerRenderSupported()).toBe('boolean')
  })
})
