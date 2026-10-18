import { describe, it, expect, vi, beforeAll, afterEach } from 'vitest'
import { render, waitFor, act } from '@testing-library/react'
import { buildDocx } from '../src/testdata/ooxml-builders'

// The loader is under test, not the viewer: OfficeDoc is replaced by a probe
// that shows the title it was handed.
vi.mock('../src/components/OfficeDoc', () => ({
  OfficeDoc: (props: { title?: string }) => <div data-testid="doc" data-title={props.title ?? ''} data-has-title={String('title' in props && props.title !== undefined)} />,
}))

const { OfficeFile } = await import('../src/components/OfficeFile')

let bytes: Uint8Array
beforeAll(async () => { bytes = await buildDocx([{ runs: [{ text: 'hello' }] }]) })
afterEach(() => { vi.unstubAllGlobals() })

function response(url: string, headers: Record<string, string> = {}): Response {
  const r = new Response(bytes.slice(), { headers })
  Object.defineProperty(r, 'url', { value: url })
  return r
}

const titleOf = async (c: HTMLElement) => {
  await waitFor(() => expect(c.querySelector('[data-testid=doc]')).toBeTruthy(), { timeout: 5000 })
  return c.querySelector('[data-testid=doc]')!.getAttribute('data-title')
}

describe('OfficeFile file name', () => {
  it('OfficeFile passes the File name as title', async () => {
    const { container } = render(<OfficeFile data={new File([bytes.slice()], 'Q3 report.docx')} />)
    expect(await titleOf(container)).toBe('Q3 report.docx')
  })

  it('uses Content-Disposition when exposed, the URL when the header is absent or not exposed', async () => {
    vi.stubGlobal('fetch', vi.fn(async (input: string) => response(input, { 'content-disposition': "attachment; filename*=UTF-8''r%C3%A9sum%C3%A9.docx" })))
    const a = render(<OfficeFile data={{ url: 'https://x/download/abc.docx?sig=1' }} />)
    await waitFor(async () => expect(await titleOf(a.container)).toBe('résumé.docx'))
    a.unmount()
    // header not exposed (CORS) reads as absent
    vi.stubGlobal('fetch', vi.fn(async (input: string) => response(input)))
    const b = render(<OfficeFile data={{ url: 'https://x/download/Plan%20v2.docx?sig=1' }} />)
    expect(await titleOf(b.container)).toBe('Plan v2.docx')
  })

  it('promised one-shot source: name from the resolved Response URL', async () => {
    const promised = Promise.resolve(response('https://cdn.example/files/Deck%201.docx'))
    const { container } = render(<OfficeFile data={promised} />)
    await waitFor(async () => expect(await titleOf(container)).toBe('Deck 1.docx'))
  })

  it('source replaced during fetch: the stale name is never shown', async () => {
    let releaseOld!: (r: Response) => void
    const old = new Promise<Response>((r) => { releaseOld = r })
    const { container, rerender } = render(<OfficeFile data={old} />)
    const fresh = new File([bytes.slice()], 'fresh.docx')
    rerender(<OfficeFile data={fresh} />)
    expect(await titleOf(container)).toBe('fresh.docx')
    await act(async () => { releaseOld(response('https://x/stale.docx', { 'content-disposition': 'attachment; filename="stale.docx"' })) })
    await new Promise((r) => setTimeout(r, 30))
    expect(container.querySelector('[data-testid=doc]')!.getAttribute('data-title')).toBe('fresh.docx')
  })

  it('host title wins', async () => {
    const { container } = render(<OfficeFile data={new File([bytes.slice()], 'Q3.docx')} title="Quarterly" />)
    expect(await titleOf(container)).toBe('Quarterly')
  })

  it('no name hides the slot', async () => {
    const { container } = render(<OfficeFile data={bytes.slice()} />)
    await titleOf(container)
    expect(container.querySelector('[data-testid=doc]')!.getAttribute('data-has-title')).toBe('false')
  })
})
