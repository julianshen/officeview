import { describe, test, expect } from 'vitest'
import { render, waitFor, fireEvent, act } from '@testing-library/react'
import { OfficeDoc } from '../src/components/OfficeDoc'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { buildDocx, type DocxParaSpec } from '../src/testdata/ooxml-builders'

async function docFor(paras: DocxParaSpec[]) {
  return parseDocx(await OfficePackage.load(await buildDocx(paras)))
}

async function mount(paras: DocxParaSpec[] = [{ runs: [{ text: 'alpha beta' }] }, { runs: [{ text: 'beta gamma' }] }]) {
  const doc = await docFor(paras)
  const utils = render(<OfficeDoc document={doc} />)
  await waitFor(() => expect(utils.container.querySelectorAll('canvas').length).toBeGreaterThan(0))
  return utils
}

const input = (c: HTMLElement) => c.querySelector('[aria-label="Search document"]') as HTMLInputElement
const status = (c: HTMLElement) => c.querySelector('[data-testid="officeview-search-status"]') as HTMLElement

describe('<OfficeDoc> search', () => {
  test('renders a search bar and reports match counts', async () => {
    const utils = await mount()
    const box = utils.container.querySelector('[data-testid="officeview-search"]')
    expect(box).toBeTruthy()
    expect(status(utils.container).textContent).toBe('')

    fireEvent.change(input(utils.container), { target: { value: 'beta' } })
    await waitFor(() => expect(status(utils.container).textContent).toBe('1 of 2'))
  })

  test('navigates matches with the buttons, wrapping around', async () => {
    const utils = await mount()
    fireEvent.change(input(utils.container), { target: { value: 'beta' } })
    await waitFor(() => expect(status(utils.container).textContent).toBe('1 of 2'))

    fireEvent.click(utils.container.querySelector('[aria-label="Next match"]') as HTMLButtonElement)
    await waitFor(() => expect(status(utils.container).textContent).toBe('2 of 2'))

    fireEvent.click(utils.container.querySelector('[aria-label="Next match"]') as HTMLButtonElement)
    await waitFor(() => expect(status(utils.container).textContent).toBe('1 of 2'))

    fireEvent.click(utils.container.querySelector('[aria-label="Previous match"]') as HTMLButtonElement)
    await waitFor(() => expect(status(utils.container).textContent).toBe('2 of 2'))
  })

  test('Enter and Shift+Enter step through matches', async () => {
    const utils = await mount()
    const el = input(utils.container)
    fireEvent.change(el, { target: { value: 'beta' } })
    await waitFor(() => expect(status(utils.container).textContent).toBe('1 of 2'))
    fireEvent.keyDown(el, { key: 'Enter' })
    await waitFor(() => expect(status(utils.container).textContent).toBe('2 of 2'))
    fireEvent.keyDown(el, { key: 'Enter', shiftKey: true })
    await waitFor(() => expect(status(utils.container).textContent).toBe('1 of 2'))
  })

  test('reports no matches and disables navigation', async () => {
    const utils = await mount()
    fireEvent.change(input(utils.container), { target: { value: 'zzzz' } })
    await waitFor(() => expect(status(utils.container).textContent).toBe('no matches'))
    expect((utils.container.querySelector('[aria-label="Next match"]') as HTMLButtonElement).disabled).toBe(true)
  })

  test('Escape clears the query and highlights', async () => {
    const utils = await mount()
    const el = input(utils.container)
    fireEvent.change(el, { target: { value: 'beta' } })
    await waitFor(() => expect(status(utils.container).textContent).toBe('1 of 2'))
    fireEvent.keyDown(el, { key: 'Escape' })
    await waitFor(() => expect(el.value).toBe(''))
    expect(status(utils.container).textContent).toBe('')
  })

  test('highlights matches onto the canvas pixels', async () => {
    const utils = await mount()
    fireEvent.change(input(utils.container), { target: { value: 'alpha' } })
    await waitFor(() => expect(status(utils.container).textContent).toBe('1 of 1'))
    // the page canvas should now contain highlight-yellow pixels
    const canvas = utils.container.querySelector('canvas') as HTMLCanvasElement
    const ctx = canvas.getContext('2d')!
    const img = ctx.getImageData(0, 0, canvas.width, canvas.height).data
    let yellow = 0
    for (let i = 0; i < img.length; i += 4) {
      // rgba(255,214,0,0.45) over white lands around (255, 240, 190)
      if (img[i] > 245 && img[i + 1] > 225 && img[i + 1] < 252 && img[i + 2] < 210) yellow++
    }
    expect(yellow).toBeGreaterThan(20)
  })

  test('cmd/ctrl+F focuses the search input', async () => {
    const utils = await mount()
    const el = input(utils.container)
    const region = utils.container.querySelector('[role="region"]') as HTMLElement
    el.blur()
    fireEvent.keyDown(region, { key: 'f', metaKey: true })
    await waitFor(() => expect(document.activeElement).toBe(el))
  })

  test('search can be hidden', async () => {
    const doc = await docFor([{ runs: [{ text: 'x' }] }])
    const utils = render(<OfficeDoc document={doc} showSearch={false} />)
    await waitFor(() => expect(utils.container.querySelectorAll('canvas').length).toBeGreaterThan(0))
    expect(utils.container.querySelector('[data-testid="officeview-search"]')).toBeNull()
    void act
  })
})
