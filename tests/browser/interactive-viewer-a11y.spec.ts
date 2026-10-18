/**
 * S10.3 keyboard, focus and accessibility-tree receipts for the real
 * interactive viewer (spec §9.2–9.3, AC-A11Y-01/02), on the shipped
 * <OfficeFile interactiveSheets> mounted by src/debug/viewer-e2e.html.
 *
 * Keyboard scenarios run on Chromium and Playwright WebKit (`webkit-playwright`,
 * never Safari evidence). The accessibility tree is read from Chromium's own
 * accessibility layer over CDP (Accessibility.getFullAXTree), i.e. what a
 * screen reader is handed, not the DOM; ARIA index attributes, which CDP does
 * not expose as AX properties, are read from the owned DOM. A real screen
 * reader is not driven here (see docs/validation/interactive-viewer.md).
 */
import { expect, test, type Page } from '@playwright/test'
import { activeCell, ensureFixtures, goTo, openViewer, status } from './viewer-e2e'

test.use({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 2 })
test.beforeAll(async () => { await ensureFixtures() })

const grid = (page: Page) => page.getByRole('grid')
const activeName = async (page: Page) => (await activeCell(page))?.label.split(',')[0] ?? ''
const focusedLabel = (page: Page) => page.evaluate(() => {
  const el = document.activeElement
  return el?.getAttribute('role') === 'grid' ? 'grid' : el?.getAttribute('aria-label') ?? el?.tagName ?? ''
})
/** The descendant id names an owned gridcell inside the grid (not a stale id). */
const descendantValid = (page: Page) => page.evaluate(() => {
  const g = document.querySelector('[role=grid]')!, id = g.getAttribute('aria-activedescendant')
  const cell = id ? document.getElementById(id) : null
  return Boolean(cell && g.contains(cell) && cell.getAttribute('role') === 'gridcell' && cell.parentElement?.getAttribute('role') === 'row')
})
const selectedCells = (page: Page) => page.locator('[role=gridcell][aria-selected=true]').count()

test.describe('S10.3 keyboard and focus (real viewer)', () => {
  test('one Tab stop; arrows, Enter, Home, Ctrl edges after recenter; announcements', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    // One grid Tab entry: the toolbar's last stop (the search box; its match
    // buttons are disabled) → grid → out of the grid; Shift+Tab returns to it.
    await page.getByRole('searchbox', { name: 'Search document' }).focus()
    await page.keyboard.press('Tab')
    expect(await focusedLabel(page)).toBe('grid')
    await page.keyboard.press('Tab')
    expect(await page.evaluate(() => { const g = document.querySelector('[role=grid]')!; return g === document.activeElement || g.contains(document.activeElement) })).toBe(false)
    await page.keyboard.press('Shift+Tab')
    expect(await focusedLabel(page)).toBe('grid')

    await goTo(page, 'B2')
    await grid(page).focus()
    const steps: Array<[string, string]> = [['ArrowDown', 'B3'], ['ArrowRight', 'C3'], ['Enter', 'C4'], ['Home', 'A4'], ['ControlOrMeta+Home', 'A1']]
    for (const [key, want] of steps) {
      await page.keyboard.press(key)
      await expect.poll(() => activeName(page)).toBe(want)
      expect(await descendantValid(page)).toBe(true)
      await expect(status(page)).toContainText(want) // live region names the new active cell
    }
    // Edge navigation across recentered segments: last row, last column, then back one.
    await page.keyboard.press('ControlOrMeta+ArrowDown')
    await expect.poll(() => activeName(page)).toBe('A1048576')
    await page.keyboard.press('ArrowUp')
    await expect.poll(() => activeName(page)).toBe('A1048575')
    await page.keyboard.press('ControlOrMeta+ArrowRight')
    await expect.poll(() => activeName(page)).toBe('XFD1048575')
    await page.keyboard.press('ArrowLeft')
    await expect.poll(() => activeName(page)).toBe('XFC1048575')
    await expect(page.getByRole('rowheader', { name: 'Row 1048575', exact: true })).toBeVisible()
    await expect(page.getByRole('columnheader', { name: /^Column XFC\b/ })).toBeVisible()
    expect(await descendantValid(page)).toBe(true)
    await page.keyboard.press('ControlOrMeta+Home')
    await expect.poll(() => activeName(page)).toBe('A1')
    expect(await focusedLabel(page)).toBe('grid') // focus never left the grid wrapper
  })

  test('multi-selection, Esc order (popup, then selection), Alt+Down and focus return', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    await goTo(page, 'B2')
    await grid(page).focus()
    await page.keyboard.press('Shift+ArrowDown')
    await page.keyboard.press('Shift+ArrowDown')
    await page.keyboard.press('Shift+ArrowRight')
    await expect.poll(() => selectedCells(page)).toBe(6) // B2:C4
    // Ctrl/Cmd+click adds a disjoint cell; the grid keeps focus and one active cell.
    const e6 = page.getByRole('gridcell', { name: /^E6, / })
    const box = (await e6.boundingBox())!
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2)
    await page.keyboard.down('ControlOrMeta')
    await page.mouse.down(); await page.mouse.up()
    await page.keyboard.up('ControlOrMeta')
    await expect.poll(() => selectedCells(page)).toBe(7)
    await expect.poll(() => activeName(page)).toBe('E6')
    expect(await focusedLabel(page)).toBe('grid')

    // Alt+Down opens the active column's filter; focus moves into the popup.
    await page.keyboard.press('Alt+ArrowDown')
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    await expect(dialog).toBeVisible()
    await expect.poll(() => dialog.evaluate(d => d.contains(document.activeElement))).toBe(true)
    // First Esc closes the popup only: selection survives, focus returns to the grid.
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    expect(await focusedLabel(page)).toBe('grid')
    expect(await selectedCells(page)).toBe(7)
    expect(await descendantValid(page)).toBe(true)
    // Second Esc clears the grid selection and the descendant.
    await page.keyboard.press('Escape')
    await expect.poll(() => selectedCells(page)).toBe(0)
    await expect(status(page)).toContainText('Selection cleared.')
    expect(await grid(page).getAttribute('aria-activedescendant')).toBeNull()
    expect(await focusedLabel(page)).toBe('grid')
  })

  test('F2 control mode: Enter/Space activate header controls, popup returns focus to its invoker, Esc back to grid', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    await goTo(page, 'D2')
    await grid(page).focus()
    await page.keyboard.press('F2')
    expect(await focusedLabel(page)).toBe('Select Column D')
    await page.keyboard.press('Tab')
    expect(await focusedLabel(page)).toBe('Cycle sort column D')
    await page.keyboard.press('Enter')
    const header = page.getByRole('columnheader', { name: /^Column D/ })
    await expect(header).toHaveAttribute('aria-sort', 'ascending', { timeout: 30_000 })
    await expect(status(page)).toContainText('View-only sort applied')
    // A committed view is a new owner: control mode ends and focus returns to the
    // grid with a valid active descendant (never lost to <body>).
    await expect.poll(() => focusedLabel(page)).toBe('grid')
    expect(await descendantValid(page)).toBe(true)
    // Space on the same native control (re-entered with F2) cycles to descending.
    await page.keyboard.press('F2')
    await page.keyboard.press('Tab')
    expect(await focusedLabel(page)).toBe('Cycle sort column D')
    await page.keyboard.press(' ')
    await expect(header).toHaveAttribute('aria-sort', 'descending', { timeout: 30_000 })
    await expect(header).toHaveAttribute('aria-label', 'Column D; sort priority 1, descending')
    await expect.poll(() => focusedLabel(page)).toBe('grid')

    await page.keyboard.press('F2')
    await page.keyboard.press('Tab')
    await page.keyboard.press('Tab')
    await expect.poll(() => focusedLabel(page)).toMatch(/^Filter column D/)
    await page.keyboard.press('Enter')
    const dialog = page.getByRole('dialog', { name: 'Column filter' })
    await expect(dialog).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    await expect.poll(() => focusedLabel(page)).toMatch(/^Filter column D/) // the invoker, still owned
    await page.keyboard.press('Escape')
    expect(await focusedLabel(page)).toBe('grid')
    expect(await descendantValid(page)).toBe(true)
    expect(await activeName(page)).toBe('D2')
    // Enter on the column's select control selects the whole column (aria-selected on its header).
    await page.keyboard.press('F2')
    expect(await focusedLabel(page)).toBe('Select Column D; sort priority 1, descending') // the name carries sort state
    await page.keyboard.press('Enter')
    await expect(header).toHaveAttribute('aria-selected', 'true')
    await expect(page.locator('[role=columnheader][aria-selected=true]')).toHaveCount(1)
    await page.keyboard.press('Escape')
    expect(await focusedLabel(page)).toBe('grid')
  })

  test('workbook find and editable inputs are never hijacked by grid or zoom keys', async ({ page }) => {
    await openViewer(page, 'f2-1k.xlsx')
    await goTo(page, 'C3')
    await grid(page).focus()
    const zoom = page.getByRole('button', { name: 'Reset zoom' })
    await expect(zoom).toHaveText('100%')
    await page.keyboard.press('ControlOrMeta+f')
    const search = page.getByLabel('Search document')
    await expect(search).toBeFocused()
    // Typing zoom characters into an input edits the input only (find-as-you-type
    // may move the active cell to a hit; grid arrows and zoom keys must not act).
    await page.keyboard.type('-0+=')
    await expect(search).toHaveValue('-0+=')
    await expect(zoom).toHaveText('100%')
    const afterFind = await activeName(page)
    await page.keyboard.press('ArrowDown')
    await page.keyboard.press('Home')
    await expect(search).toBeFocused()
    expect(await activeName(page)).toBe(afterFind)
    await goTo(page, 'C3')
    await expect.poll(() => activeName(page)).toBe('C3')
    const goToBox = page.getByRole('textbox', { name: 'Go To cell' })
    await goToBox.fill('')
    await goToBox.focus()
    await page.keyboard.type('0-')
    await page.keyboard.press('ArrowUp')
    await expect(goToBox).toHaveValue('0-')
    await expect(goToBox).toBeFocused()
    await expect(zoom).toHaveText('100%')
    expect(await activeName(page)).toBe('C3')
    // The same keys on the grid do zoom: '+' zooms in and '0' resets.
    await grid(page).focus()
    await page.keyboard.press('+')
    await expect(zoom).not.toHaveText('100%')
    await page.keyboard.press('0')
    await expect(zoom).toHaveText('100%')
  })
})

interface AXNode {
  nodeId: string
  ignored: boolean
  role?: { value: string }
  name?: { value: string }
  properties?: Array<{ name: string; value: { value?: unknown; relatedNodes?: Array<{ backendDOMNodeId: number }> } }>
  parentId?: string
  childIds?: string[]
  backendDOMNodeId?: number
}
const prop = (node: AXNode, name: string) => node.properties?.find(p => p.name === name)?.value

test.describe('S10.3 accessibility tree (Chromium AX layer over CDP)', () => {
  test('owned grid/row/header/gridcell hierarchy, display-slot indices, active descendant, selection, sort, live region', async ({ page, browserName }) => {
    test.skip(browserName !== 'chromium', 'CDP Accessibility domain is Chromium-only')
    await openViewer(page, 'f7-two-sheets.xlsx')
    // A sort (score ascending: id 40 leads) and an authored-hidden row 8 give
    // display slots that differ from source rows.
    await page.getByRole('button', { name: 'Cycle sort column B' }).click()
    await expect(page.getByRole('columnheader', { name: /^Column B/ })).toHaveAttribute('aria-sort', 'ascending', { timeout: 30_000 })
    await goTo(page, 'A2')
    await grid(page).focus()
    await page.keyboard.press('Shift+ArrowDown')
    await page.keyboard.press('Shift+ArrowRight')
    await expect.poll(() => selectedCells(page)).toBe(4)
    await expect(status(page)).toContainText('B3, value') // the extension target is announced
    const domActive = (await activeCell(page))!

    const cdp = await page.context().newCDPSession(page)
    await cdp.send('Accessibility.enable')
    const { nodes } = await cdp.send('Accessibility.getFullAXTree') as { nodes: AXNode[] }
    const byId = new Map(nodes.map(n => [n.nodeId, n]))
    const role = (n: AXNode | undefined) => n?.role?.value ?? ''
    // Nearest non-generic ancestor: Chromium keeps positioned divs as 'generic'/ignored nodes.
    const semanticParent = (n: AXNode): AXNode | undefined => {
      let p = n.parentId ? byId.get(n.parentId) : undefined
      while (p && (p.ignored || ['generic', 'none', 'GenericContainer'].includes(role(p)))) p = p.parentId ? byId.get(p.parentId) : undefined
      return p
    }
    const grids = nodes.filter(n => !n.ignored && role(n) === 'grid')
    expect(grids).toHaveLength(1)
    const g = grids[0]
    expect(g.name?.value).toBe('Worksheet A')
    expect(prop(g, 'focused')?.value).toBe(true)
    expect(prop(g, 'multiselectable')?.value).toBe(true)
    const of = (r: string) => nodes.filter(n => !n.ignored && role(n) === r)
    const rows = of('row'), cells = of('gridcell'), rowHeaders = of('rowheader'), colHeaders = of('columnheader')
    expect(cells.length).toBeGreaterThan(20)
    expect(cells.length).toBeLessThan(2000) // bounded, viewport-only
    for (const r of rows) expect(role(semanticParent(r))).toBe('grid')
    for (const c of [...cells, ...rowHeaders, ...colHeaders]) expect(role(semanticParent(c))).toBe('row')
    // Every body row owns exactly one rowheader; the header row owns the corner + column headers.
    const childrenOf = (row: AXNode) => nodes.filter(n => !n.ignored && semanticParent(n) === row)
    const headerRows = rows.filter(r => childrenOf(r).some(c => role(c) === 'columnheader'))
    expect(headerRows).toHaveLength(1)
    for (const r of rows.filter(r => r !== headerRows[0])) expect(childrenOf(r).filter(c => role(c) === 'rowheader')).toHaveLength(1)

    // Active descendant: an AX relation from the grid to an owned, named gridcell.
    const related = prop(g, 'activedescendant')?.relatedNodes?.[0]?.backendDOMNodeId
    const activeNode = nodes.find(n => n.backendDOMNodeId === related)
    expect(role(activeNode)).toBe('gridcell')
    expect(activeNode!.name?.value).toBe(domActive.label)
    expect(domActive.label).toMatch(/; original [A-Z]+\d+$/) // display address differs from source after the sort
    expect(role(semanticParent(semanticParent(activeNode!)!))).toBe('grid')
    // Selection is exposed by the AX layer (aria-selected → selected).
    expect(cells.filter(c => prop(c, 'selected')?.value === true)).toHaveLength(4)
    // Live region: the status line is a polite, atomic region carrying the announcement.
    const live = nodes.filter(n => !n.ignored && role(n) === 'status' && prop(n, 'live')?.value === 'polite')
    expect(live.length).toBeGreaterThanOrEqual(1)
    expect(live.some(n => prop(n, 'atomic')?.value === true)).toBe(true)

    // Display-slot indices and sort state from the owned DOM (no AX property for them).
    const dom = await page.evaluate(() => {
      const g = document.querySelector('[role=grid]')!
      const rows = [...g.querySelectorAll('[role=row]')]
      const header = rows.find(r => r.querySelector('[role=columnheader]'))!
      const body = rows.filter(r => r !== header && !r.hasAttribute('data-transient-active'))
      const label = (r: Element) => r.querySelector('[role=rowheader]')?.getAttribute('aria-label')
      return {
        rowcount: g.getAttribute('aria-rowcount'), colcount: g.getAttribute('aria-colcount'),
        headerIndex: header.getAttribute('aria-rowindex'),
        corner: header.querySelector('[role=columnheader]')?.getAttribute('aria-colindex'),
        colIndex: [...header.querySelectorAll('[role=columnheader]')].slice(1, 4).map(h => [h.getAttribute('aria-label')?.split(';')[0], h.getAttribute('aria-colindex')]),
        bodyRows: body.map(r => [r.getAttribute('aria-rowindex'), label(r), r.querySelector('[role=rowheader]')?.getAttribute('aria-colindex'), r.querySelector('[role=gridcell]')?.getAttribute('aria-colindex'), r.querySelector('[role=gridcell][aria-colindex="2"]')?.textContent ?? null]),
        duplicates: body.length - new Set(body.map(r => r.getAttribute('aria-rowindex'))).size,
        sorted: [...header.querySelectorAll('[aria-sort]')].map(h => [h.getAttribute('aria-label'), h.getAttribute('aria-sort')]),
      }
    })
    expect(dom).toMatchObject({ rowcount: '1048577', colcount: '16385', headerIndex: '1', corner: '1', duplicates: 0 })
    expect(dom.colIndex).toEqual([['Column A', '2'], ['Column B', '3'], ['Column C', '4']])
    // Fixture-derived expectations (f7 sheet A, score ascending = id descending).
    // Body ids 1..40 sit in source rows 2..41; source row 8 (id 7) is authored-hidden
    // and keeps its display slot 7, so slot 7 is a gap; slots 1-6 hold ids 40..35 and
    // slots >= 8 hold ids 34, 33, ... The 0.75pt row (id 9) sorts to slot 33, outside
    // this viewport. aria-rowindex = DisplaySlot + 2 and the label Row DisplaySlot + 1
    // — VisibleRank indexing would close the gap and shift every row after it.
    const expected = (slot: number) => {
      // Past slot 34 (id 8) the descending ids skip the pinned id 7 (a viewer
      // that fills its host shows those slots too).
      const id = slot <= 6 ? 41 - slot : slot <= 34 ? 42 - slot : 41 - slot
      // "original" is appended only when the source address differs (spec §9.2).
      return { index: String(slot + 2), label: `Row ${slot + 1}`, cellA: id < 1 ? `A${slot + 1}, value ` /* below the 40-row body: blank */ : `A${slot + 1}, value ${id}${id === slot ? '' : `; original A${id + 1}`}` }
    }
    const indices = dom.bodyRows.map(([index]) => Number(index))
    expect(indices).toContain(8) // slot 6
    expect(indices).toContain(10) // slot 8
    expect(indices).not.toContain(9) // slot 7: the pinned hidden source row 8
    for (const [index, label, headerCol, firstCol, cellA] of dom.bodyRows) {
      const slot = Number(index) - 2
      expect(headerCol).toBe('1')
      expect(firstCol).toBe('2')
      if (slot === 0) { expect(label).toBe('Row 1'); continue } // the protected header row
      const want = expected(slot)
      expect([index, label]).toEqual([want.index, want.label])
      expect(cellA).toBe(want.cellA)
    }
    expect(dom.sorted).toEqual([['Column B; sort priority 1, ascending', 'ascending']])
    console.log(`[a11y-evidence] chromium AX: 1 grid "Worksheet A" (focused, multiselectable), rows=${rows.length}, columnheaders=${colHeaders.length}, rowheaders=${rowHeaders.length}, gridcells=${cells.length}, selected=4, activedescendant="${activeNode!.name?.value}", live=polite/atomic; DOM rowcount/colcount=${dom.rowcount}/${dom.colcount}, sorted=${JSON.stringify(dom.sorted)}`)
  })
})

test.describe('S10.3 touch emulation (Chromium device emulation; not iPad evidence)', () => {
  test.use({ viewport: { width: 1024, height: 768 }, deviceScaleFactor: 2, hasTouch: true, isMobile: true })
  test('coarse pointer: 44px header/row targets, tap selects a cell, row-header taps are owned by their row', async ({ page, browserName }) => {
    test.skip(browserName !== 'chromium', 'isMobile/coarse-pointer emulation is Chromium-only here')
    await openViewer(page, 'f2-1k.xlsx')
    expect(await page.evaluate(() => matchMedia('(pointer: coarse)').matches)).toBe(true)
    const small = await page.evaluate(() => {
      const pane = document.querySelector('[role=grid]')!.getBoundingClientRect()
      return [...document.querySelectorAll<HTMLElement>('[role=grid] button')]
        .map(b => ({ name: b.getAttribute('aria-label'), ...b.getBoundingClientRect().toJSON() }))
        // A column cut by the pane edge is clamped to its visible width by design (§9.3 ownership).
        .filter(b => b.width > 0 && b.right < pane.right - 1 && b.bottom < pane.bottom - 1)
        .filter(b => b.height < 44 - 0.5 || b.width < 44 - 0.5)
    })
    // Every fully visible control keeps the 44px floor (default columns are 64px wide).
    expect(small).toEqual([])
    // The owned cell lies under the painted surface: tap its geometry, as a finger would.
    const cell = (await page.getByRole('gridcell', { name: /^C5, / }).boundingBox())!
    await page.touchscreen.tap(cell.x + cell.width / 2, cell.y + cell.height / 2)
    await expect.poll(() => activeName(page)).toBe('C5')
    // A tap near the top edge of row 6's header selects row 6, never its neighbour.
    const header = (await page.getByRole('rowheader', { name: 'Row 6', exact: true }).boundingBox())!
    await page.touchscreen.tap(header.x + header.width / 2, header.y + 2)
    await expect(page.getByRole('rowheader', { name: 'Row 6', exact: true })).toHaveAttribute('aria-selected', 'true')
    await expect(page.locator('[role=rowheader][aria-selected=true]')).toHaveCount(1)
  })
})
