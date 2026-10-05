import { readFileSync } from 'node:fs'
import { afterEach, describe, expect, test, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { getPaintables } from '../src/render/paint'
import type { PptxDocument } from '../src/pptx/types'
import type { FontRegistrationRequest, RegisterFont } from '../src/core/fonts/register'
const bytes = new Uint8Array(readFileSync('tests/fixtures/fonts/Liter-Regular.ttf'))
function doc(): PptxDocument {
 return { slideWidthEmu: 9144000, slideHeightEmu: 6858000, images: [], embeddedFonts: ['regular', 'bold', 'italic', 'boldItalic'].map(variant => ({ family: 'Liter', variant: variant as 'regular', partPath: `font/${variant}`, relationshipId: variant, bytes })), slides: [{ index: 0, widthEmu: 9144000, heightEmu: 6858000, shapes: [{ xEmu: 0, yEmu: 0, widthEmu: 5000000, heightEmu: 2000000, geometry: 'rect', textBody: { anchor: 't', insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0, wrap: true, paragraphs: [{ align: 'left', level: 0, runs: [{ text: 'Guide for NATS', fontFamily: 'Liter', bold: true, fontSizePt: 54 }] }] } }] }] }
}
function deferred() { let resolve!: () => void; const promise = new Promise<void>(r => { resolve = r }); return { promise, resolve } }
afterEach(() => { vi.restoreAllMocks(); vi.unstubAllGlobals() })
describe('embedded font pipeline leases', () => {
 test('awaits headless registration, passes real bytes/descriptors and uses alias for measure and paint', async () => {
  const gate = deferred(), requests: FontRegistrationRequest[] = []; const cleanup = vi.fn()
  const registerFont: RegisterFont = async request => { requests.push(request); await gate.promise; return cleanup }
  const source = doc(); let done = false
  const pending = getPaintables(source, { registerFont }).then(p => { done = true; return p })
  await new Promise(r => setTimeout(r, 0)); expect(done).toBe(false); gate.resolve()
  const pages = await pending
  expect(requests).toHaveLength(4)
  expect(requests.map(r => r.descriptors)).toEqual([{ weight: '400', style: 'normal' }, { weight: '700', style: 'normal' }, { weight: '400', style: 'italic' }, { weight: '700', style: 'italic' }])
  expect(requests.every(r => r.bytes === bytes && r.alias !== 'Liter')).toBe(true)
  expect(new Set(requests.map(r => r.alias)).size).toBe(1)
  const ctx = createCanvas(1000, 700).getContext('2d'), fonts: string[] = [], measured: string[] = []
  const prototype = Object.getPrototypeOf(ctx) as { measureText: typeof ctx.measureText }
  const measure = prototype.measureText as typeof ctx.measureText, fill = ctx.fillText.bind(ctx)
  vi.spyOn(prototype, 'measureText').mockImplementation(function(this: typeof ctx, text: string) { measured.push(this.font); return measure.call(this, text) })
  ctx.fillText = (text, x, y) => { fonts.push(ctx.font); fill(text, x, y) }
  pages[0].paint(ctx as unknown as CanvasRenderingContext2D)
  expect(measured.length).toBeGreaterThan(1); expect(measured.every(f => f.includes(requests[0].alias) && f.includes('bold'))).toBe(true)
  expect(fonts.length).toBeGreaterThan(0); expect(fonts.every(f => f.includes(requests[0].alias) && f.includes('bold'))).toBe(true)
  expect(source.slides[0].shapes[0].textBody?.paragraphs[0].runs[0].fontFamily).toBe('Liter')
  pages.dispose(); pages.dispose(); expect(cleanup).toHaveBeenCalledTimes(4)
 })
 test('shares pending/loaded registration per document and adapter until final consumer releases', async () => {
  const gate = deferred(), cleanup = vi.fn(), registerFont = vi.fn(async () => { await gate.promise; return cleanup })
  const source = doc(); const a = getPaintables(source, { registerFont }), b = getPaintables(source, { registerFont })
  gate.resolve(); const [pa, pb] = await Promise.all([a, b]); expect(registerFont).toHaveBeenCalledTimes(4)
  pa.dispose(); expect(cleanup).not.toHaveBeenCalled(); pb.dispose(); expect(cleanup).toHaveBeenCalledTimes(4)
  const pc = await getPaintables(source, { registerFont }); expect(registerFont).toHaveBeenCalledTimes(8); pc.dispose()
 })
 test('isolates same families across documents and different registration adapters', async () => {
  const a: string[] = [], b: string[] = []
  const ra: RegisterFont = async r => { a.push(r.alias) }, rb: RegisterFont = async r => { b.push(r.alias) }
  const source = doc()
  const pages = await Promise.all([getPaintables(source, { registerFont: ra }), getPaintables(doc(), { registerFont: ra }), getPaintables(source, { registerFont: rb })])
  expect(new Set([...a, ...b]).size).toBe(3); pages.forEach(p => p.dispose())
 })
 test('one failed face keeps neighboring faces available with auditable diagnostics', async () => {
  const source = doc(); const cleanup = vi.fn()
  const pages = await getPaintables(source, { registerFont: async r => { if (r.face.variant === 'bold') throw new Error('bad face'); return cleanup } })
  expect(pages).toHaveLength(1)
  expect(pages.fontDiagnostics).toEqual([expect.objectContaining({ kind: 'font-load-failed', family: 'Liter', variant: 'bold', partPath: 'font/bold' })])
  pages.dispose(); expect(cleanup).toHaveBeenCalledTimes(3)
 })
 test('browser FontFace loads before add and final release deletes only loaded faces', async () => {
  const gate = deferred(), add = vi.fn(), remove = vi.fn(), constructed: unknown[][] = []
  class MockFace { constructor(...args: unknown[]) { constructed.push(args) } async load() { await gate.promise; return this } }
  vi.stubGlobal('FontFace', MockFace); Object.defineProperty(document, 'fonts', { configurable: true, value: { add, delete: remove } })
  const source = doc(); const a = getPaintables(source), b = getPaintables(source)
  await new Promise(r => setTimeout(r, 0)); expect(add).not.toHaveBeenCalled(); gate.resolve()
  const [pa, pb] = await Promise.all([a, b]); expect(constructed).toHaveLength(4); expect(add).toHaveBeenCalledTimes(4)
  pa.dispose(); expect(remove).not.toHaveBeenCalled(); pb.dispose(); expect(remove).toHaveBeenCalledTimes(4)
 })
 test('no font handwritten models retain array and synchronous low level compatibility', async () => {
  const source = doc(); delete source.embeddedFonts
  const pages = await getPaintables(source)
  expect(Array.isArray(pages)).toBe(true); expect(pages.fontDiagnostics).toEqual([]); pages.dispose()
 })
 test('whole extraction failure releases successful registrations', async () => {
  const cleanup = vi.fn(), source = doc(); source.images = undefined as never
  await expect(getPaintables(source, { registerFont: async () => cleanup })).rejects.toThrow()
  expect(cleanup).toHaveBeenCalledTimes(4)
 })
})
