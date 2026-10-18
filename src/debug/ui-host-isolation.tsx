/**
 * Dev-only page for tests/browser/ui-host-isolation.spec.ts (spec §6.8): the
 * interactive sheet viewer and the paged viewer mounted inside a host
 * `<div class="app">`, optionally under a hostile host stylesheet.
 *
 * Query parameters:
 *   host=first|last  inject the hostile sheet before / after officeview.css
 *                    (absent: a clean page)
 *   mode=pair        two paged viewers instead: viewer 1 light inside
 *                    `.brand-a` (a host accent override), viewer 2 dark
 *   theme=dark|system  theme of the sheet and paged viewers (default light)
 *
 * Page layout uses inline styles only, so the hostile sheet cannot move the
 * viewers; only what reaches viewer elements can differ. Excluded from the
 * library build (tsconfig.build.json excludes src/debug).
 */
import { useEffect, useState, type JSX } from 'react'
import { createRoot } from 'react-dom/client'
import { OfficeDoc, type OfficeDocSource } from '../components/OfficeDoc'
import '../styles/officeview.css'
import { OfficePackage } from '../core/zip'
import { parseDocx } from '../docx/parse'
import { parseXlsx } from '../xlsx/parse'
import { buildDocx, buildXlsx } from '../testdata/ooxml-builders'
// Loaded on demand by the library; the harness only serves modules reached from src/debug.
import '../components/xlsx/InteractiveSheetDoc'
import '../components/paged/PagedChrome'

const HOSTILE_CSS = 'button{all:unset} *{box-sizing:content-box} .flex{display:block} html{font-size:10px} .app button{padding:20px;color:rgb(1,2,3)}'
const BRAND_CSS = '.brand-a [data-ov-theme-root]{--ov-color-accent:#c50f1f}'

const params = new URLSearchParams(location.search)
const host = params.get('host')
const pair = params.get('mode') === 'pair'
const theme = params.get('theme') === 'dark' ? 'dark' : params.get('theme') === 'system' ? 'system' : 'light'

function addStyle(id: string, css: string, first: boolean): void {
  const style = document.createElement('style')
  style.id = id
  style.textContent = css
  if (first) document.head.prepend(style)
  else document.head.append(style)
}
// officeview.css was injected when its import above evaluated, so append lands after it.
if (host === 'first' || host === 'last') addStyle('hostile', HOSTILE_CSS, host === 'first')
if (pair) addStyle('brand', BRAND_CSS, false)

async function load(): Promise<{ docx: OfficeDocSource; xlsx: OfficeDocSource }> {
  const docx = await parseDocx(await OfficePackage.load(await buildDocx([
    { runs: [{ text: 'Host isolation: a paged document with enough text to search.' }] },
  ])))
  const xlsx = await parseXlsx(await OfficePackage.load(await buildXlsx([
    { name: 'Data', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }, { ref: 'B1', v: 2 }] }, { r: 2, cells: [{ ref: 'A2', v: 3 }] }] },
  ])))
  return { docx, xlsx }
}

const box = { height: 360, display: 'flex', flexDirection: 'column' } as const

function App(): JSX.Element {
  const [docs, setDocs] = useState<{ docx: OfficeDocSource; xlsx: OfficeDocSource } | null>(null)
  useEffect(() => { void load().then(setDocs) }, [])
  if (!docs) return <p>loading</p>
  if (pair) {
    return (
      <div className="app" style={{ display: 'flex', flexDirection: 'row', width: 1000 }}>
        <div className="brand-a" data-viewer="1" style={{ ...box, width: 500 }}><OfficeDoc document={docs.docx} theme="light" /></div>
        <div data-viewer="2" style={{ ...box, width: 500 }}><OfficeDoc document={docs.docx} theme="dark" /></div>
      </div>
    )
  }
  return (
    <div className="app" style={{ display: 'block', width: 900 }}>
      <div data-viewer="sheet" style={box}><OfficeDoc document={docs.xlsx} interactiveSheets theme={theme} /></div>
      <div data-viewer="paged" style={box}><OfficeDoc document={docs.docx} theme={theme} /></div>
      <button id="host-button" type="button">Host button</button>
      <div id="host-flex" className="flex">host flex</div>
    </div>
  )
}

createRoot(document.getElementById('root')!).render(<App />)
