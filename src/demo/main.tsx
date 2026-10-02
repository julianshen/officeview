/**
 * Browser demo for officeview: open any .docx/.xlsx/.pptx/.odt, or click a
 * generated sample. Mobile-first: touch scroll, responsive width, safe areas.
 */
import { useState, type ChangeEvent } from 'react'
import { createRoot } from 'react-dom/client'
import { OfficeFile } from '../components/OfficeFile'
import { buildDocx, buildXlsx, buildPptx, type DocxCellSpec, type DocxParaSpec, type DocxTableSpec } from '../testdata/ooxml-builders'
import { buildOdt, odfDecimalListStyle } from '../testdata/odf-builders'

const p = (text: string): DocxParaSpec => ({ runs: [{ text }] })

async function sampleDocx(): Promise<Uint8Array> {
  const table: DocxTableSpec = {
    gridCols: ['4320', '4320'],
    borders: '<w:top w:val="single"/><w:left w:val="single"/><w:bottom w:val="single"/><w:right w:val="single"/><w:insideH w:val="single"/><w:insideV w:val="single"/>',
    rows: [
      { cells: [{ paragraphs: [{ align: 'center', runs: [{ text: 'Quarter', bold: true }] }], fill: 'FFCC00' }, { paragraphs: [{ runs: [{ text: 'Revenue', bold: true }] }], fill: 'D9D9D9' }] },
      { cells: [{ paragraphs: [p('Q1')] }, { paragraphs: [{ runs: [{ text: '$1,204.50', color: '0070C0' }] }] }] },
      { cells: [{ paragraphs: [p('Q2')] }, { paragraphs: [{ runs: [{ text: '$2,918.25', color: 'C00000' }] }] }] },
    ],
  }
  const paras: DocxParaSpec[] = [
    { align: 'center', runs: [{ text: 'Quarterly Report', bold: true, size: 56 }] },
    { runs: [{ text: 'This document is rendered ', }, { text: 'directly on a <canvas>', italic: true }, { text: ' — no HTML translation. Word-wrapped paragraphs, ', }, { text: 'bold runs', bold: true }, { text: ', colors and spacing all flow through the same layout engine.' }] },
    { align: 'right', runs: [{ text: 'Right-aligned paragraph', color: '0070C0' }] },
    { runs: [{ text: 'A longer paragraph that must wrap across several lines because it keeps going well past the width of the printed page area, exercising the line-breaking logic on a mobile-sized viewport where the page is scaled down to fit the container width. Lorem ipsum dolor sit amet, consectetur adipiscing elit.' }] },
    { runs: [{ text: 'A table with shading, borders and centered headers follows:', }] },
  ]
  return buildDocx(paras, [table])
}

async function sampleXlsx(): Promise<Uint8Array> {
  return buildXlsx([
    {
      name: 'Budget',
      rows: [
        { r: 1, cells: [{ ref: 'A1', t: 's', v: 0, style: 1 }, { ref: 'B1', t: 's', v: 1, style: 1 }, { ref: 'C1', t: 's', v: 2, style: 1 }, { ref: 'D1', t: 's', v: 3, style: 1 }] },
        { r: 2, cells: [{ ref: 'A2', t: 's', v: 4 }, { ref: 'B2', v: 1250.5, style: 2 }, { ref: 'C2', v: 0.185, style: 3 }, { ref: 'D2', v: 45000, style: 4 }] },
        { r: 3, cells: [{ ref: 'A3', t: 's', v: 5 }, { ref: 'B3', v: 980.25, style: 2 }, { ref: 'C3', v: 0.42, style: 3 }, { ref: 'D3', v: 46000, style: 4 }] },
        { r: 4, cells: [{ ref: 'A4', t: 's', v: 6 }, { ref: 'B4', v: 2310.75, style: 2 }, { ref: 'C4', v: 0.605, style: 3 }, { ref: 'D4', v: 47000, style: 4 }] },
      ],
      cols: '<col min="1" max="1" width="14" customWidth="1"/><col min="2" max="4" width="12"/>',
    },
  ], ['Category', 'Amount', 'Share', 'Date', 'Hardware', 'Marketing', 'Total'])
}

/** A small checkerboard PNG for the picture sample. */
async function samplePng(): Promise<Uint8Array> {
  const c = document.createElement('canvas')
  c.width = 96
  c.height = 96
  const ctx = c.getContext('2d')!
  for (let y = 0; y < 6; y++) {
    for (let x = 0; x < 6; x++) {
      ctx.fillStyle = (x + y) % 2 === 0 ? '#2f6fed' : '#ffffff'
      ctx.fillRect(x * 16, y * 16, 16, 16)
    }
  }
  ctx.fillStyle = '#ff3b30'
  ctx.beginPath()
  ctx.arc(48, 48, 16, 0, Math.PI * 2)
  ctx.fill()
  const blob = await new Promise<Blob | null>((resolve) => c.toBlob(resolve, 'image/png'))
  return new Uint8Array(await blob!.arrayBuffer())
}

async function samplePptx(): Promise<Uint8Array> {
  const png = await samplePng()
  return buildPptx([
    {
      prst: 'rect',
      off: ['914400', '914400'],
      ext: ['7315200', '1828800'],
      fill: 'FFCC00',
      paragraphs: [{ align: 'ctr', runs: [{ text: 'officeview on canvas', b: true, sz: '4400' }] }],
    },
    {
      prst: 'roundRect',
      off: ['914400', '3200400'],
      ext: ['4572000', '1828800'],
      fill: 'ED7D31',
      paragraphs: [
        { align: 'left', runs: [{ text: 'Shapes with fills, outlines,', b: true, sz: '2000' }] },
        { align: 'left', runs: [{ text: 'and wrapped text bodies', i: true, sz: '2000' }] },
      ],
    },
    { prst: 'ellipse', off: ['5943600', '3200400'], ext: ['1828800', '1828800'], fill: '4472C4' },
    { image: { data: png }, off: ['1600200', '5486400'], ext: ['1828800', '1828800'] },
    {
      off: ['5029200', '5486400'],
      ext: ['4114800', '1143000'],
      table: {
        colWidths: ['2057400', '1028700', '1028700'],
        firstRow: true,
        bandRow: true,
        styleId: 'DemoTableStyle',
        rows: [
          {
            h: '342900',
            cells: [
              { paragraphs: [{ align: 'ctr', runs: [{ text: 'Quarter', b: true, sz: '1400', color: 'FFFFFF' }] }], fill: '4472C4' },
              { paragraphs: [{ align: 'ctr', runs: [{ text: 'Revenue', b: true, sz: '1400', color: 'FFFFFF' }] }], fill: '4472C4' },
              { paragraphs: [{ align: 'ctr', runs: [{ text: 'Share', b: true, sz: '1400', color: 'FFFFFF' }] }], fill: '4472C4' },
            ],
          },
          { cells: [{ paragraphs: [{ runs: [{ text: 'Q1', sz: '1400' }] }] }, { paragraphs: [{ runs: [{ text: '$1,204', sz: '1400' }] }] }, { paragraphs: [{ runs: [{ text: '29%', sz: '1400' }] }] }] },
          { cells: [{ paragraphs: [{ runs: [{ text: 'Q2', sz: '1400' }] }] }, { paragraphs: [{ runs: [{ text: '$2,918', sz: '1400' }] }] }, { paragraphs: [{ runs: [{ text: '71%', sz: '1400' }] }] }] },
        ],
      },
    },
  ])
}

type Sample = { label: string; emoji: string; load: () => Promise<Uint8Array | ReadableStream<Uint8Array>> }
async function sampleDocxLongTable(): Promise<Uint8Array> {
  // 3 merged "group" cells in column A, sized so a group crosses the page break
  const GROUPS = [22, 25, 13]
  const groupCells: DocxCellSpec[] = []
  GROUPS.forEach((size, g) => {
    for (let i = 0; i < size; i++) {
      groupCells.push(
        i === 0
          ? { paragraphs: [p(`Group ${g + 1}`)], vMerge: 'restart', fill: 'F2F2F2' }
          : { vMerge: 'continue' },
      )
    }
  })
  const table: DocxTableSpec = {
    gridCols: ['2880', '3600', '2880'],
    borders: '<w:top w:val="single"/><w:left w:val="single"/><w:bottom w:val="single"/><w:right w:val="single"/><w:insideH w:val="single"/><w:insideV w:val="single"/>',
    rows: [
      {
        isHeader: true,
        cells: [
          { paragraphs: [{ align: 'center', runs: [{ text: 'Group', bold: true }] }], fill: 'D9D9D9' },
          { paragraphs: [{ align: 'center', runs: [{ text: 'Item', bold: true }] }], fill: 'D9D9D9' },
          { paragraphs: [{ align: 'center', runs: [{ text: 'Status', bold: true }] }], fill: 'D9D9D9' },
        ],
      },
      ...groupCells.map((groupCell, i) => ({
        cells: [
          groupCell,
          { paragraphs: [p(`Row ${i + 1}`)] },
          { paragraphs: [{ runs: [{ text: i % 3 === 0 ? 'at risk' : 'on track', color: i % 3 === 0 ? 'C00000' : '0070C0' }] }] },
        ],
      })),
    ],
  }
  return buildDocx(
    [
      { align: 'center', runs: [{ text: 'Multi-page table', bold: true, size: 44 }] },
      p('The header row repeats on each page, and merged group cells (column A) continue across the page break.'),
    ],
    [table],
  )
}

async function sampleDocxHeaderFooter(): Promise<Uint8Array> {
  return buildDocx(
    [
      { align: 'center', runs: [{ text: 'Quarterly Summary', bold: true, size: 40 }] },
      p('Every page below carries the same header and a footer with live page numbers (PAGE / NUMPAGES fields), substituted at paint time.'),
      ...Array.from({ length: 46 }, (_, i) => p(`Section ${Math.floor(i / 12) + 1} detail line ${i + 1}.`)),
    ],
    [],
    {
      header: [{ align: 'right', runs: [{ text: 'ACME Confidential', italic: true, color: '808080', size: 18 }] }],
      footer: [
        {
          align: 'center',
          runs: [
            { text: 'Page ' },
            { text: '1', field: 'PAGE' },
            { text: ' of ' },
            { text: '1', field: 'NUMPAGES' },
          ],
        },
      ],
    },
  )
}

async function sampleDocxLists(): Promise<Uint8Array> {
  return buildDocx(
    [
      { runs: [{ text: 'Numbered list (decimal)', bold: true }] },
      { runs: [{ text: 'Resolve document parts' }], numId: 1 },
      { runs: [{ text: 'Lay out and paginate' }], numId: 1 },
      { runs: [{ text: 'Paint onto the canvas' }], numId: 1 },
      { runs: [{ text: 'Bullets' }], numId: 2 },
      { runs: [{ text: 'First bullet' }], numId: 2 },
      { runs: [{ text: 'Second bullet, long enough to wrap onto a second line so the hanging indent is visible' }], numId: 2 },
      { runs: [{ text: 'Nested list' }], numId: 2 },
      { runs: [{ text: 'Nested item' }], numId: 2, ilvl: 1 },
      { runs: [{ text: 'Nested item two' }], numId: 2, ilvl: 1 },
      { runs: [{ text: 'Alphabetical and roman', bold: true }] },
      { runs: [{ text: 'alpha' }], numId: 3 },
      { runs: [{ text: 'beta' }], numId: 3 },
      { runs: [{ text: 'roman' }], numId: 4 },
    ],
  )
}

async function sampleDocxStreamed(): Promise<ReadableStream<Uint8Array>> {
  const bytes = await sampleDocx()
  let at = 0
  return new ReadableStream<Uint8Array>({
    async pull(controller) {
      await new Promise((r) => setTimeout(r, 40))
      if (at >= bytes.length) {
        controller.close()
        return
      }
      controller.enqueue(bytes.subarray(at, at + 96))
      at += 96
    },
  })
}

async function sampleDocxProtected(): Promise<Uint8Array> {
  return sampleDocx()
}

async function sampleOdt(): Promise<Uint8Array> {
  return buildOdt({
    extraAutoStyles: odfDecimalListStyle('DemoList'),
    paras: [
      { text: 'Writer document (ODF)', heading: 1 },
      { runs: [{ text: 'This .odt renders through the same layout engine as Word files: ' }, { text: 'styles cascade', bold: true }, { text: ' from ODF automatic styles.' }] },
    ],
    lists: [{ styleName: 'DemoList', items: [{ text: 'first item' }, { text: 'second item' }] }],
    tables: [{
      colWidths: ['5cm', '5cm'],
      tableBorders: 'fo:border-top="0.5pt solid #000000" fo:border-bottom="0.5pt solid #000000"',
      rows: [{ cells: [{ text: 'Writer' }, { text: 'renders here' }] }],
    }],
  })
}
  
  async function sampleRtf(): Promise<Uint8Array> {
  // A hand-written RTF exercising the paths the parser supports: run formatting,
  // a colour table, alignment, a table, a list and a picture.
  // Reuse the real checkerboard PNG: RTF stores pictures as a hex payload, so
  // this also proves a genuine (not hand-faked) PNG survives the encode/decode.
  const hex = Array.from(await samplePng())
    .map((b) => b.toString(16).padStart(2, '0'))
    .join('')
  // \pict payloads are conventionally wrapped; whitespace is ignored on read.
  const onePixelPng = hex.replace(/(.{64})/g, '$1\n')
  const rtf = [
    '{\\rtf1\\ansi\\ansicpg1252\\deff0',
    '{\\fonttbl{\\f0\\froman\\fcharset204 Times New Roman;}{\\f1\\fswiss\\fcharset0 Calibri;}{\\f2\\fnil\\fcharset2 Symbol;}}',
    '{\\colortbl ;\\red220\\green38\\blue38;\\red0\\green112\\blue192;\\red242\\green242\\blue242;}',
    '\\paperw12240\\paperh15840\\margl1440\\margr1440\\margt1440\\margb1440',
    '\\qc\\b\\fs36\\cf1 RTF via canvas\\b0\\fs24\\cf0\\par',
    '\\pard\\qc\\i Rendered straight from the RTF byte stream \\u8212 ?no HTML in between.\\i0\\par',
    '\\pard\\sa240 Formatting runs: \\b bold\\b0 , \\i italic\\i0 , \\ul underline\\ulnone , \\strike struck\\strike0 , \\fs32 larger\\fs24 .\\par',
    '\\pard\\b Colour table:\\b0 \\cf1 red\\cf0 , \\cf2 blue\\cf0 .\\par',
    // \\ansicpg1251 + a charset-204 default font: the hex bytes below decode as
    // Cyrillic (Привет), and this exercises the browser's own TextDecoder path.
    String.raw`\pard\b Code page 1251:\b0 \'cf\'f0\'e8\'e2\'e5\'f2 \u8212 ? not mojibake.\par`,
    '\\pard\\b A table:\\b0\\par',
    '\\trowd\\trgaph108\\trleft0',
    '\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrb\\brdrs\\clbrdrr\\brdrs\\cellx3600',
    '\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrb\\brdrs\\clbrdrr\\brdrs\\cellx7200',
    '\\clcbpat3\\intbl \\b Shaded cell\\b0\\par',
    '\\intbl Plain cell\\par \\cell',
    '\\intbl Third column\\par \\cell\\row',
    '\\pard\\b An inline picture:\\b0 {\\pict\\pngblip\\picw16\\pich16\\picwgoal576\\pichgoal576',
    onePixelPng,
    '}\\par',
    '\\pard\\b A list:\\b0\\par',
    '\\li720\\fi-360{\\*\\pn\\pnf2\\pnindent720\\pnstart1\\pndec}{\\listtext\\f2 1.\\tab}\\ilvl0\\pntext First numbered item\\par',
    '\\li720\\fi-360{\\*\\pn\\pnf2\\pnindent720\\pnstart1\\pndec}{\\listtext\\f2 2.\\tab}\\ilvl0\\pntext Second numbered item\\par',
    '}',
  ].join('\n')
  return new TextEncoder().encode(rtf)
}

const SAMPLES: Sample[] = [
  { label: 'Word (.docx)', emoji: '📄', load: sampleDocx },
  { label: 'Writer (.odt)', emoji: '📝', load: sampleOdt },
  { label: 'RTF (.rtf)', emoji: '📃', load: sampleRtf },
  { label: 'Protected', emoji: '🔒', load: sampleDocxProtected },
  { label: 'Streamed', emoji: '🌊', load: sampleDocxStreamed },
  { label: 'Word lists', emoji: '🔢', load: sampleDocxLists },
  { label: 'Word header/footer', emoji: '📄', load: sampleDocxHeaderFooter },
  { label: 'Word long table', emoji: '🧾', load: sampleDocxLongTable },
  { label: 'Excel (.xlsx)', emoji: '📊', load: sampleXlsx },
  { label: 'PowerPoint (.pptx)', emoji: '📽️', load: samplePptx },
]

function App() {
  const [data, setData] = useState<Uint8Array | ReadableStream<Uint8Array> | null>(null)
  const [protectedMode, setProtectedMode] = useState(false)
  const [label, setLabel] = useState('')
  const [busy, setBusy] = useState(false)

  const openSample = async (s: Sample) => {
    setBusy(true)
    try {
      setLabel(`sample ${s.label}`)
      setProtectedMode(s.label === 'Protected')
      setData(await s.load())
    } finally {
      setBusy(false)
    }
  }

  const onFile = (e: ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0]
    if (!f) return
    setBusy(true)
    setLabel(f.name)
    const reader = new FileReader()
    reader.onload = () => {
      setProtectedMode(false)
      setData(new Uint8Array(reader.result as ArrayBuffer))
      setBusy(false)
    }
    reader.readAsArrayBuffer(f)
  }

  return (
    <div style={{ fontFamily: '-apple-system, system-ui, sans-serif', color: '#e8eaf0', minHeight: '100dvh', paddingBottom: 'env(safe-area-inset-bottom)' }}>
      <header style={{ padding: 'calc(env(safe-area-inset-top) + 16px) 16px 8px', position: 'sticky', top: 0, background: 'rgba(26,29,35,0.92)', backdropFilter: 'blur(8px)', zIndex: 10, borderBottom: '1px solid #2c313c' }}>
        <h1 style={{ fontSize: 18, margin: '0 0 8px', fontWeight: 700 }}>officeview <span style={{ color: '#8b93a7', fontWeight: 400 }}>— canvas renderer</span></h1>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
          <label style={{ background: '#2f6fed', color: '#fff', padding: '8px 14px', borderRadius: 8, fontSize: 14, cursor: 'pointer' }}>
            Open file…
            <input type="file" accept=".docx,.xlsx,.pptx,.odt" onChange={onFile} style={{ display: 'none' }} />
          </label>
          {SAMPLES.map((s) => (
            <button key={s.label} onClick={() => void openSample(s)} disabled={busy}
              style={{ background: '#2c313c', color: '#e8eaf0', border: '1px solid #3a4150', padding: '8px 12px', borderRadius: 8, fontSize: 14, cursor: 'pointer' }}>
              {s.emoji} {s.label}
            </button>
          ))}
        </div>
        {label && <p style={{ margin: '8px 0 0', fontSize: 12, color: '#8b93a7' }}>viewing: {label}</p>}
      </header>
      <main style={{ padding: '12px 0' }}>
        {data === null ? (
          <p style={{ textAlign: 'center', color: '#8b93a7', padding: '40px 24px' }}>
            Open a .docx / .xlsx / .pptx / .odt — or tap a sample above.
          </p>
        ) : (
          <OfficeFile
            data={data}
            loading={(p) => (
              <p style={{ textAlign: 'center', color: '#8b93a7' }}>
                {p.total ? `Streaming… ${Math.round((p.loaded / p.total) * 100)}%` : `Received ${p.loaded.toLocaleString()} bytes…`}
              </p>
            )}
            error={(msg) => <p style={{ textAlign: 'center', color: '#ff6b6b' }}>⚠️ {msg}</p>}
            allowCopy={!protectedMode}
            allowPrint={!protectedMode}
            background="#22252d"
            style={{ minHeight: '60dvh' }}
          />
        )}
      </main>
    </div>
  )
}

createRoot(document.getElementById('root')!).render(<App />)
