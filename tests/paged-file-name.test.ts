import { describe, it, expect } from 'vitest'
import { fileNameFrom } from '../src/components/fileName'

describe('fileNameFrom', () => {
  it('File.name', () => {
    expect(fileNameFrom(new File([], 'Q3.docx'))).toBe('Q3.docx')
  })
  it('HttpSource URL: last segment, query stripped, decoded', () => {
    expect(fileNameFrom({ url: 'https://x/a/b/Plan%20v2.pptx?sig=1' })).toBe('Plan v2.pptx')
  })
  it('Request URL: hash stripped', () => {
    expect(fileNameFrom(new Request('https://x/r.rtf#p'))).toBe('r.rtf')
  })
  it('Response URL', () => {
    const r = new Response('x')
    Object.defineProperty(r, 'url', { value: 'https://x/files/deck.pptx' })
    expect(fileNameFrom(r)).toBe('deck.pptx')
  })
  it('RFC 6266: filename* wins and is decoded', () => {
    expect(fileNameFrom({ url: 'https://x/download' }, "attachment; filename*=UTF-8''r%C3%A9sum%C3%A9.docx")).toBe('résumé.docx')
    expect(fileNameFrom({ url: 'https://x/download' }, 'attachment; filename="plain.docx"; filename*=UTF-8\'\'%E2%82%AC.docx')).toBe('€.docx')
  })
  it('a malformed filename* falls back to filename', () => {
    expect(fileNameFrom({ url: 'https://x/u.docx' }, "attachment; filename*=UTF-8''%E0%A4; filename=\"x.docx\"")).toBe('x.docx')
  })
  it('filename wins over the URL; unquoted filename works', () => {
    expect(fileNameFrom({ url: 'https://x/u.docx' }, 'attachment; filename=y.docx')).toBe('y.docx')
  })
  it('a malformed URL escape keeps the raw segment', () => {
    expect(fileNameFrom({ url: 'https://x/bad%E0.docx' })).toBe('bad%E0.docx')
  })
  it('no name: bytes, an empty path, nothing', () => {
    expect(fileNameFrom(new Uint8Array())).toBeUndefined()
    expect(fileNameFrom({ url: 'https://x/' })).toBeUndefined()
    expect(fileNameFrom(undefined)).toBeUndefined()
    expect(fileNameFrom(new Blob([]))).toBeUndefined()
  })
})
