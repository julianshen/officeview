/**
 * Generates the sample documents the website (site/) shows in its demos and
 * playground, into site/public/samples/. Every file is written by this script
 * (or by the repo's own test-data builders), so the site ships no third-party
 * documents and the samples can be regenerated at any time:
 *
 *   bun scripts/build-site-samples.ts
 *
 * The documents are fictional ("Halcyon Instruments"); the numbers are made up.
 */
import JSZip from 'jszip'
import { deflateSync } from 'node:zlib'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { buildPptxDeck, mainSeqXml, timingXmlOf, type PptxShapeSpec, type PptxSlideSpec } from '../src/testdata/ooxml-builders'
import { buildOdt, odfBulletListStyle, odfWriterHeadingStyles } from '../src/testdata/odf-builders'
import { buildChartDocument, chartExGallery, classicGallery, galleryHierarchyWorkbook } from '../src/testdata/chart-gallery'

const OUT = fileURLToPath(new URL('../site/public/samples/', import.meta.url))

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

// ---------------------------------------------------------------------------
// PNG (for the report's figure and the RTF picture)
// ---------------------------------------------------------------------------

const CRC_TABLE = Array.from({ length: 256 }, (_, n) => {
  let c = n
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
  return c >>> 0
})
function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff
  for (const b of bytes) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}
function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length)
  const view = new DataView(out.buffer)
  view.setUint32(0, data.length)
  const typeAndData = new Uint8Array(4 + data.length)
  typeAndData.set(new TextEncoder().encode(type))
  typeAndData.set(data, 4)
  out.set(typeAndData, 4)
  view.setUint32(8 + data.length, crc32(typeAndData))
  return out
}

/** A tiny RGB raster with rectangle fills, encoded as PNG. */
class Raster {
  readonly px: Uint8Array
  constructor(readonly width: number, readonly height: number, background: string) {
    this.px = new Uint8Array(width * height * 3)
    this.rect(0, 0, width, height, background)
  }
  rect(x: number, y: number, w: number, h: number, hex: string): void {
    const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16))
    for (let yy = Math.max(0, Math.round(y)); yy < Math.min(this.height, Math.round(y + h)); yy++) {
      for (let xx = Math.max(0, Math.round(x)); xx < Math.min(this.width, Math.round(x + w)); xx++) {
        const i = (yy * this.width + xx) * 3
        this.px[i] = r; this.px[i + 1] = g; this.px[i + 2] = b
      }
    }
  }
  png(): Uint8Array {
    const ihdr = new Uint8Array(13)
    const v = new DataView(ihdr.buffer)
    v.setUint32(0, this.width)
    v.setUint32(4, this.height)
    ihdr.set([8, 2, 0, 0, 0], 8) // 8-bit RGB
    const raw = new Uint8Array(this.height * (1 + this.width * 3))
    for (let y = 0; y < this.height; y++) raw.set(this.px.subarray(y * this.width * 3, (y + 1) * this.width * 3), y * (1 + this.width * 3) + 1)
    const parts = [new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]), pngChunk('IHDR', ihdr), pngChunk('IDAT', new Uint8Array(deflateSync(raw))), pngChunk('IEND', new Uint8Array())]
    const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
    let at = 0
    for (const p of parts) { out.set(p, at); at += p.length }
    return out
  }
}

/** Quarterly revenue bars by region (the report's Figure 1). */
function revenueFigure(): Uint8Array {
  const r = new Raster(1200, 520, 'FFFFFF')
  const colors = ['0F4C81', '2E86C1', '85C1E9']
  const data = [[4.2, 3.1, 2.0], [4.8, 3.4, 2.3], [5.6, 3.9, 2.6], [6.1, 4.4, 3.1]]
  for (let g = 0; g <= 4; g++) r.rect(80, 460 - g * 100, 1080, 2, 'E5E8EC')
  data.forEach((quarter, q) => {
    quarter.forEach((value, s) => {
      const h = value * 62
      r.rect(140 + q * 260 + s * 58, 460 - h, 48, h, colors[s])
    })
  })
  r.rect(80, 460, 1080, 3, '8A94A6')
  return r.png()
}

/** A small logo mark for the RTF letterhead. */
function logoMark(): Uint8Array {
  const r = new Raster(96, 96, 'FFFFFF')
  r.rect(8, 8, 80, 80, '0F4C81')
  r.rect(24, 24, 14, 48, 'FFFFFF')
  r.rect(58, 24, 14, 48, 'FFFFFF')
  r.rect(24, 42, 48, 12, 'FFFFFF')
  return r.png()
}

// ---------------------------------------------------------------------------
// Word: a multi-page quarterly business review
// ---------------------------------------------------------------------------

type Run = string | { text: string; b?: boolean; i?: boolean; color?: string; sz?: number; field?: string }

function runXml(run: Run): string {
  const r = typeof run === 'string' ? { text: run } : run
  const rPr = `<w:rPr>${r.b ? '<w:b/>' : ''}${r.i ? '<w:i/>' : ''}${r.color ? `<w:color w:val="${r.color}"/>` : ''}${r.sz ? `<w:sz w:val="${r.sz}"/>` : ''}</w:rPr>`
  if (r.field) {
    return `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rPr}<w:instrText xml:space="preserve"> ${r.field} </w:instrText></w:r>` +
      `<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${rPr}<w:t>${esc(r.text)}</w:t></w:r><w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`
  }
  return `<w:r>${rPr}<w:t xml:space="preserve">${esc(r.text)}</w:t></w:r>`
}

function para(runs: Run[], opts: { style?: string; align?: string; num?: [numId: number, ilvl: number]; after?: number; before?: number; keepNext?: boolean } = {}): string {
  const pPr = [
    opts.style ? `<w:pStyle w:val="${opts.style}"/>` : '',
    opts.keepNext ? '<w:keepNext/>' : '',
    opts.num ? `<w:numPr><w:ilvl w:val="${opts.num[1]}"/><w:numId w:val="${opts.num[0]}"/></w:numPr>` : '',
    opts.after !== undefined || opts.before !== undefined ? `<w:spacing${opts.before !== undefined ? ` w:before="${opts.before}"` : ''}${opts.after !== undefined ? ` w:after="${opts.after}"` : ''}/>` : '',
    opts.align ? `<w:jc w:val="${opts.align}"/>` : '',
  ].join('')
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs.map(runXml).join('')}</w:p>`
}

interface Cell { runs: Run[]; fill?: string; align?: string; span?: number }

function table(widths: number[], rows: Cell[][], opts: { header?: boolean; band?: string } = {}): string {
  const border = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="C9D1DC"/>`
  const tblPr = `<w:tblPr><w:tblW w:w="${widths.reduce((a, b) => a + b, 0)}" w:type="dxa"/>` +
    `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders>` +
    `<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="110" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="110" w:type="dxa"/></w:tblCellMar></w:tblPr>`
  const grid = `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`
  const body = rows.map((cells, ri) => {
    const header = opts.header && ri === 0
    const trPr = header ? '<w:trPr><w:tblHeader/></w:trPr>' : ''
    let col = 0
    const tcs = cells.map((c) => {
      const w = widths.slice(col, col + (c.span ?? 1)).reduce((a, b) => a + b, 0)
      col += c.span ?? 1
      const fill = c.fill ?? (header ? '0F4C81' : opts.band && ri % 2 === 0 ? opts.band : undefined)
      const runs = header ? c.runs.map((r) => (typeof r === 'string' ? { text: r, b: true, color: 'FFFFFF' } : { ...r, b: true, color: 'FFFFFF' })) : c.runs
      return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${c.span ? `<w:gridSpan w:val="${c.span}"/>` : ''}${fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : ''}<w:vAlign w:val="center"/></w:tcPr>` +
        para(runs, { align: c.align, after: 0 }) + '</w:tc>'
    }).join('')
    return `<w:tr>${trPr}${tcs}</w:tr>`
  }).join('')
  return `<w:tbl>${tblPr}${grid}${body}</w:tbl>`
}

function inlinePicture(rid: string, cx: number, cy: number, name: string): string {
  return `<w:p><w:pPr><w:spacing w:before="120" w:after="60"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="${esc(name)}"/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="${esc(name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
    `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
}

async function reportDocx(): Promise<Uint8Array> {
  const H1 = (t: string) => para([t], { style: 'Heading1', keepNext: true })
  const H2 = (t: string) => para([t], { style: 'Heading2', keepNext: true })
  const P = (...runs: Run[]) => para(runs)
  const bullet = (...runs: Run[]) => para(runs, { num: [1, 0], after: 60 })
  const step = (...runs: Run[]) => para(runs, { num: [2, 0], after: 60 })
  const money = (n: number) => `$${n.toFixed(1)}M`

  const regions: Array<[string, number[]]> = [['North America', [4.2, 4.8, 5.6, 6.1]], ['Europe', [3.1, 3.4, 3.9, 4.4]], ['Asia Pacific', [2.0, 2.3, 2.6, 3.1]]]
  const revenueRows: Cell[][] = [
    ['Region', 'Q4 2025', 'Q1 2026', 'Q2 2026', 'Q3 2026', 'Change'].map((t, i) => ({ runs: [t], align: i ? 'right' : undefined })),
    ...regions.map(([name, v]) => [
      { runs: [name] },
      ...v.map((n) => ({ runs: [money(n)], align: 'right' })),
      { runs: [{ text: `+${Math.round(((v[3] - v[0]) / v[0]) * 100)}%`, color: '107C41', b: true }], align: 'right' },
    ]),
    [
      { runs: [{ text: 'Total', b: true }], fill: 'E8EEF6' },
      ...[0, 1, 2, 3].map((q) => ({ runs: [{ text: money(regions.reduce((s, [, v]) => s + v[q], 0)), b: true }], align: 'right', fill: 'E8EEF6' })),
      { runs: [{ text: '+47%', b: true, color: '107C41' }], align: 'right', fill: 'E8EEF6' },
    ],
  ]

  const milestones: Cell[][] = [
    ['Milestone', 'Owner', 'Due', 'Status'].map((t) => ({ runs: [t] })),
    [{ runs: ['Field calibration kit GA'] }, { runs: ['Hardware'] }, { runs: ['Oct 14'] }, { runs: [{ text: 'Shipped', color: '107C41', b: true }] }],
    [{ runs: ['Fleet dashboard 2.0'] }, { runs: ['Platform'] }, { runs: ['Nov 3'] }, { runs: [{ text: 'On track', color: '0F6CBD', b: true }] }],
    [{ runs: ['EU data residency'] }, { runs: ['Infrastructure'] }, { runs: ['Nov 21'] }, { runs: [{ text: 'At risk', color: 'C50F1F', b: true }] }],
    [{ runs: ['Partner API v3'] }, { runs: ['Platform'] }, { runs: ['Dec 9'] }, { runs: [{ text: 'On track', color: '0F6CBD', b: true }] }],
    [{ runs: ['Sensor firmware 5.2'] }, { runs: ['Embedded'] }, { runs: ['Dec 16'] }, { runs: [{ text: 'Planning', color: '605E5C', b: true }] }],
  ]

  const body = [
    para([{ text: 'Halcyon Instruments', color: '0F6CBD', b: true }], { after: 40 }),
    para(['Quarterly Business Review'], { style: 'Title' }),
    para(['Third quarter, fiscal 2026 · Prepared for the board of directors'], { style: 'Subtitle' }),
    table([2340, 2340, 2340, 2340], [[
      { runs: [{ text: 'Revenue', color: '605E5C', sz: 18 }, { text: '' }], fill: 'F3F6FA' },
      { runs: [{ text: 'Gross margin', color: '605E5C', sz: 18 }], fill: 'F3F6FA' },
      { runs: [{ text: 'Net retention', color: '605E5C', sz: 18 }], fill: 'F3F6FA' },
      { runs: [{ text: 'Customers', color: '605E5C', sz: 18 }], fill: 'F3F6FA' },
    ], [
      { runs: [{ text: '$13.6M', b: true, sz: 36, color: '0F4C81' }], fill: 'F3F6FA' },
      { runs: [{ text: '64.2%', b: true, sz: 36, color: '0F4C81' }], fill: 'F3F6FA' },
      { runs: [{ text: '118%', b: true, sz: 36, color: '0F4C81' }], fill: 'F3F6FA' },
      { runs: [{ text: '1,284', b: true, sz: 36, color: '0F4C81' }], fill: 'F3F6FA' },
    ]]),
    H1('Executive summary'),
    P('Revenue grew to ', { text: '$13.6 million', b: true }, ' in the third quarter, up 47% on the same quarter last year and ahead of plan for the second quarter running. Growth was broad: every region expanded, and Asia Pacific, our newest market, grew fastest on a percentage basis.'),
    P('Gross margin improved by 2.1 points as the field calibration kit moved to the new contract manufacturer. Operating expenses grew more slowly than revenue, and we ended the quarter with ', { text: '22 months of runway', i: true }, ' at the current burn rate.'),
    bullet({ text: 'Bookings', b: true }, ' reached $15.2M, a record, driven by three enterprise renewals in Europe.'),
    bullet({ text: 'Churn', b: true }, ' fell to 0.8% of annual recurring revenue, the lowest since launch.'),
    bullet({ text: 'Headcount', b: true }, ' closed at 212, nine below plan; hiring is concentrated in support and platform engineering.'),
    H1('Revenue by region'),
    P('All three regions grew quarter on quarter. North America remains the largest market at 45% of revenue; Europe benefited from the renewals noted above.'),
    table([2600, 1300, 1300, 1300, 1300, 1060], revenueRows, { header: true, band: 'F7F9FC' }),
    inlinePicture('rIdFig1', 5486400, 2377440, 'Figure 1'),
    para([{ text: 'Figure 1. ', b: true }, { text: 'Quarterly revenue by region, in millions of dollars (North America, Europe, Asia Pacific).', i: true, color: '605E5C' }], { align: 'center', after: 240 }),
    H1('Product milestones'),
    P('Five milestones are tracked for the fourth quarter. The EU data residency work is at risk because the second region was provisioned two weeks late; a recovery plan is described below.'),
    table([3600, 2000, 1500, 2260], milestones, { header: true, band: 'F7F9FC' }),
    H2('Recovery plan for EU data residency'),
    step('Move two platform engineers from the partner API team for six weeks.'),
    step('Run the migration rehearsal in staging on November 4 instead of November 18.'),
    step('Hold the customer announcement until the rehearsal has passed its rollback test.'),
    H1('Outlook'),
    P('We are raising the full-year revenue outlook to between $50 million and $52 million. The main risks are component lead times for the next sensor generation and the timing of two public-sector contracts that slipped from the third quarter into the fourth.'),
    H2('Priorities for the fourth quarter'),
    bullet('Ship fleet dashboard 2.0 and the partner API v3 on their current dates.'),
    bullet('Bring EU data residency back on schedule, with weekly status to the board.'),
    bullet('Close the two public-sector contracts and begin onboarding in January.'),
    bullet('Finish the support reorganisation so that every enterprise account has a named engineer.'),
    para([{ text: 'This review was prepared by the finance team. Figures are unaudited.', i: true, color: '605E5C', sz: 18 }], { before: 360 }),
  ].join('')

  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
  const document = `${XML}<w:document ${W} xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${body}` +
    `<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/><w:footerReference w:type="default" r:id="rIdFtr"/>` +
    `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1300" w:right="1440" w:bottom="1300" w:left="1440" w:header="620" w:footer="620" w:gutter="0"/></w:sectPr></w:body></w:document>`
  const header = `${XML}<w:hdr ${W}>${para([{ text: 'Halcyon Instruments  ·  Q3 2026 business review', color: '8A94A6', sz: 17 }], { align: 'right', after: 0 })}</w:hdr>`
  const footer = `${XML}<w:ftr ${W}>${para([{ text: 'Confidential  ·  Page ', color: '8A94A6', sz: 17 }, { text: '1', field: 'PAGE', color: '8A94A6', sz: 17 }, { text: ' of ', color: '8A94A6', sz: 17 }, { text: '1', field: 'NUMPAGES', color: '8A94A6', sz: 17 }], { align: 'center', after: 0 })}</w:ftr>`
  const styles = `${XML}<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="242424"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="288" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="60"/></w:pPr><w:rPr><w:b/><w:color w:val="0F2B46"/><w:sz w:val="56"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="320"/></w:pPr><w:rPr><w:color w:val="605E5C"/><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="400" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="0F4C81"/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="280" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="0F6CBD"/><w:sz w:val="25"/></w:rPr></w:style>
</w:styles>`
  const lvl = (fmt: string, text: string) => `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="540" w:hanging="300"/></w:pPr></w:lvl>`
  const numbering = `${XML}<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:abstractNum w:abstractNumId="0">${lvl('bullet', '•')}</w:abstractNum><w:abstractNum w:abstractNumId="1">${lvl('decimal', '%1.')}</w:abstractNum>` +
    `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>`)
  zip.file('_rels/.rels', `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`)
  zip.file('word/_rels/document.xml.rels', `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rIdNum" Type="${REL}/numbering" Target="numbering.xml"/>
<Relationship Id="rIdHdr" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdFtr" Type="${REL}/footer" Target="footer1.xml"/>
<Relationship Id="rIdFig1" Type="${REL}/image" Target="media/figure1.png"/></Relationships>`)
  zip.file('word/document.xml', document)
  zip.file('word/styles.xml', styles)
  zip.file('word/numbering.xml', numbering)
  zip.file('word/header1.xml', header)
  zip.file('word/footer1.xml', footer)
  zip.file('word/media/figure1.png', revenueFigure())
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

// ---------------------------------------------------------------------------
// Excel
// ---------------------------------------------------------------------------

interface XCell { v: string | number; s?: number; f?: string }
type XRow = Array<XCell | null>

const colName = (i: number): string => (i >= 26 ? colName(Math.floor(i / 26) - 1) : '') + String.fromCharCode(65 + (i % 26))

/** Cell style indices into XLSX_STYLES' cellXfs. */
const S = { title: 1, header: 2, text: 3, money: 4, pct: 5, totalLabel: 6, totalMoney: 7, subtitle: 8, int: 9, date: 10, totalPct: 11, headerLeft: 12 }

const XLSX_STYLES = `${XML}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="3"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0"/><numFmt numFmtId="165" formatCode="0.0%"/><numFmt numFmtId="166" formatCode="d mmm yyyy"/></numFmts>
<fonts count="5"><font><sz val="10"/><color rgb="FF242424"/><name val="Arial"/></font><font><b/><sz val="16"/><color rgb="FF0F2B46"/><name val="Arial"/></font>
<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FF242424"/><name val="Arial"/></font><font><i/><sz val="10"/><color rgb="FF605E5C"/><name val="Arial"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF0F4C81"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8EEF6"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left/><right/><top/><bottom style="thin"><color rgb="FFD5DCE6"/></bottom><diagonal/></border>
<border><left/><right/><top style="medium"><color rgb="FF0F4C81"/></top><bottom style="medium"><color rgb="FF0F4C81"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="13">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="0" fontId="3" fillId="3" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="164" fontId="3" fillId="3" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left"/></xf>
<xf numFmtId="165" fontId="3" fillId="3" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
</cellXfs></styleSheet>`

interface SheetSpec { name: string; rows: XRow[]; widths: number[]; merges?: string[]; freezeRows?: number; autoFilter?: string; rowHeights?: Record<number, number> }

async function writeXlsx(sheets: SheetSpec[]): Promise<Uint8Array> {
  const strings: string[] = []
  const stringIndex = new Map<string, number>()
  const sst = (s: string) => {
    let i = stringIndex.get(s)
    if (i === undefined) { i = strings.length; strings.push(s); stringIndex.set(s, i) }
    return i
  }
  const sheetXml = (sheet: SheetSpec) => {
    const rows = sheet.rows.map((row, ri) => {
      const ht = sheet.rowHeights?.[ri + 1]
      const cells = row.map((c, ci) => {
        if (!c) return ''
        const ref = `${colName(ci)}${ri + 1}`
        const s = c.s !== undefined ? ` s="${c.s}"` : ''
        const f = c.f ? `<f>${esc(c.f)}</f>` : ''
        return typeof c.v === 'string' && !c.f ? `<c r="${ref}"${s} t="s"><v>${sst(c.v)}</v></c>` : `<c r="${ref}"${s}>${f}<v>${c.v}</v></c>`
      }).join('')
      return `<row r="${ri + 1}"${ht ? ` ht="${ht}" customHeight="1"` : ''}>${cells}</row>`
    }).join('')
    const pane = sheet.freezeRows ? `<pane ySplit="${sheet.freezeRows}" topLeftCell="A${sheet.freezeRows + 1}" activePane="bottomLeft" state="frozen"/>` : ''
    return `${XML}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<sheetViews><sheetView workbookViewId="0" showGridLines="1">${pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>` +
      `<cols>${sheet.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>` +
      `<sheetData>${rows}</sheetData>${sheet.autoFilter ? `<autoFilter ref="${sheet.autoFilter}"/>` : ''}` +
      `${sheet.merges?.length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : ''}` +
      `<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>`
  }
  const sheetParts = sheets.map(sheetXml)
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
    sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
    `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`)
  zip.file('_rels/.rels', `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`)
  zip.file('xl/workbook.xml', `${XML}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><sheets>${sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>` +
    `${sheets.some((s) => s.autoFilter) ? `<definedNames>${sheets.map((s, i) => (s.autoFilter ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${esc(s.name)}'!$${s.autoFilter.replace(':', ':$').replace(/([A-Z]+)(\d+)/g, '$1$$$2')}</definedName>` : '')).join('')}</definedNames>` : ''}</workbook>`)
  zip.file('xl/_rels/workbook.xml.rels', `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
    `<Relationship Id="rIdSst" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/></Relationships>`)
  sheetParts.forEach((xml, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, xml))
  zip.file('xl/sharedStrings.xml', `${XML}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">${strings.map((s) => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('')}</sst>`)
  zip.file('xl/styles.xml', XLSX_STYLES)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/** A small, presentation-quality budget workbook: two sheets, formulas with cached values. */
async function budgetXlsx(): Promise<Uint8Array> {
  const lines: Array<[string, number[]]> = [
    ['Salaries and benefits', [1_820_000, 1_870_000, 1_940_000, 2_010_000]],
    ['Cloud infrastructure', [212_000, 228_000, 241_000, 259_000]],
    ['Contract manufacturing', [640_000, 702_000, 755_000, 790_000]],
    ['Marketing programs', [180_000, 240_000, 215_000, 300_000]],
    ['Travel and events', [64_000, 92_000, 71_000, 110_000]],
    ['Facilities', [118_000, 118_000, 121_000, 121_000]],
    ['Software and tools', [46_000, 48_000, 52_000, 55_000]],
    ['Professional services', [72_000, 38_000, 44_000, 90_000]],
  ]
  const first = 5
  const totals = [0, 1, 2, 3].map((q) => lines.reduce((s, [, v]) => s + v[q], 0))
  const grand = totals.reduce((a, b) => a + b, 0)
  const budgetRows: XRow[] = [
    [{ v: 'Operating budget, fiscal 2027', s: S.title }],
    [{ v: 'Halcyon Instruments · all figures in US dollars', s: S.subtitle }],
    [],
    [{ v: 'Cost line', s: S.headerLeft }, ...['Q1', 'Q2', 'Q3', 'Q4', 'Full year', 'Share'].map((v) => ({ v, s: S.header }))],
    ...lines.map(([name, v], i): XRow => {
      const r = first + i
      const total = v.reduce((a, b) => a + b, 0)
      return [{ v: name, s: S.text }, ...v.map((n) => ({ v: n, s: S.money })), { v: total, s: S.money, f: `SUM(B${r}:E${r})` }, { v: total / grand, s: S.pct, f: `F${r}/$F$${first + lines.length}` }]
    }),
    [{ v: 'Total', s: S.totalLabel }, ...totals.map((n, q) => ({ v: n, s: S.totalMoney, f: `SUM(${colName(q + 1)}${first}:${colName(q + 1)}${first + lines.length - 1})` })),
      { v: grand, s: S.totalMoney, f: `SUM(F${first}:F${first + lines.length - 1})` }, { v: 1, s: S.totalPct, f: `SUM(G${first}:G${first + lines.length - 1})` }],
  ]

  const people: Array<[string, number, number, number]> = [
    ['Engineering', 96, 104, 118], ['Hardware', 28, 31, 33], ['Support', 22, 30, 36], ['Sales', 31, 34, 40], ['Marketing', 12, 13, 15], ['Operations', 14, 15, 16], ['Finance and legal', 9, 9, 10],
  ]
  const hcRows: XRow[] = [
    [{ v: 'Headcount plan', s: S.title }],
    [{ v: 'End-of-period employees by team', s: S.subtitle }],
    [],
    [{ v: 'Team', s: S.headerLeft }, ...['Today', 'Mid-year', 'Year end', 'Growth'].map((v) => ({ v, s: S.header }))],
    ...people.map(([team, a, b, c], i): XRow => [{ v: team, s: S.text }, { v: a, s: S.int }, { v: b, s: S.int }, { v: c, s: S.int }, { v: (c - a) / a, s: S.pct, f: `D${first + i}/B${first + i}-1` }]),
    [{ v: 'Total', s: S.totalLabel }, ...[1, 2, 3].map((k) => ({ v: people.reduce((s, p) => s + (p[k] as number), 0), s: S.totalLabel, f: `SUM(${colName(k)}${first}:${colName(k)}${first + people.length - 1})` })),
      { v: people.reduce((s, p) => s + p[3], 0) / people.reduce((s, p) => s + p[1], 0) - 1, s: S.totalPct, f: `D${first + people.length}/B${first + people.length}-1` }],
  ]
  return writeXlsx([
    { name: 'Budget', rows: budgetRows, widths: [28, 13, 13, 13, 13, 14, 9], merges: ['A1:G1', 'A2:G2'], freezeRows: 4, rowHeights: { 1: 26, 4: 20 } },
    { name: 'Headcount', rows: hcRows, widths: [24, 12, 12, 12, 10], merges: ['A1:E1', 'A2:E2'], freezeRows: 4, rowHeights: { 1: 26, 4: 20 } },
  ])
}

/** A 2,500-row sales ledger for the interactive spreadsheet demo (sort, filter, find, freeze panes). */
async function ledgerXlsx(): Promise<Uint8Array> {
  let seed = 20261024
  const rand = () => ((seed = (seed * 1103515245 + 12345) % 2 ** 31) / 2 ** 31)
  const pick = <T,>(xs: T[]) => xs[Math.floor(rand() * xs.length)]
  const regions = ['North America', 'Europe', 'Asia Pacific', 'Latin America']
  const reps = ['A. Okafor', 'B. Lindqvist', 'C. Tanaka', 'D. Moreau', 'E. Novak', 'F. Haddad', 'G. Silva', 'H. Kim', 'I. Rossi', 'J. Patel']
  const products = [['Field calibration kit', 4800], ['Sensor array S5', 1250], ['Fleet dashboard (seat)', 420], ['Gateway G2', 2100], ['Support plan', 3600], ['Partner API (annual)', 9800]] as const
  const statuses = ['Closed', 'Closed', 'Closed', 'Invoiced', 'Pending']
  const day0 = 46023 // 1 Jan 2026 as an Excel serial date
  const rows: XRow[] = [[...['Order', 'Date', 'Region', 'Sales rep', 'Product', 'Units', 'Unit price', 'Amount', 'Status'].map((v, i) => ({ v, s: i < 5 || i === 8 ? S.headerLeft : S.header }))]]
  for (let i = 0; i < 2500; i++) {
    const [product, price] = pick([...products])
    const units = 1 + Math.floor(rand() * (price > 4000 ? 6 : 40))
    const r = i + 2
    rows.push([
      { v: `HI-${String(26000 + i)}`, s: S.text },
      { v: day0 + Math.floor((i / 2500) * 296), s: S.date },
      { v: pick(regions), s: S.text },
      { v: pick(reps), s: S.text },
      { v: product, s: S.text },
      { v: units, s: S.int },
      { v: price, s: S.money },
      { v: units * price, s: S.money, f: `F${r}*G${r}` },
      { v: pick(statuses), s: S.text },
    ])
  }
  return writeXlsx([{ name: 'Orders', rows, widths: [12, 13, 15, 14, 24, 8, 11, 12, 10], freezeRows: 1, autoFilter: `A1:I${rows.length}`, rowHeights: { 1: 20 } }])
}

// ---------------------------------------------------------------------------
// PowerPoint: a deck that exercises the slideshow
// ---------------------------------------------------------------------------

const EMU = 12700 // per point
const pt = (n: number) => String(Math.round(n * EMU))

/** Strip the builder's default black outline from every shape (it is a test-fixture default). */
async function noOutlines(bytes: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(n))) {
    const xml = await zip.file(name)!.async('string')
    zip.file(name, xml
      .replace(/<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"\/><\/a:solidFill><\/a:ln>/g, '<a:ln><a:noFill/></a:ln>')
      // Text boxes: no fill unless one was asked for, and generous insets.
      .replace(/<p:txBody><a:bodyPr\/>/g, '<p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/>'))
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/** One click: fade the shape in (presetID 10, Fade entrance). */
function fadeIn(base: number, spid: number): string {
  return `<p:par><p:cTn id="${base}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${base + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${base + 2}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:set><p:cBhvr><p:cTn id="${base + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
    `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${base + 4}" dur="500"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect>` +
    `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
}

const p14 = (xml: string, fallback = '<p:transition spd="slow"><p:fade/></p:transition>') => ({ xml, wrapInAlternateContent: { requires: 'p14', fallbackXml: fallback } })

async function launchPptx(): Promise<Uint8Array> {
  const W = 12192000, H = 6858000
  const text = (id: number, x: number, y: number, w: number, h: number, paragraphs: PptxShapeSpec['paragraphs'], fill?: string, name?: string, prst = 'rect'): PptxShapeSpec =>
    ({ id, name, prst, off: [pt(x), pt(y)], ext: [pt(w), pt(h)], fill, paragraphs })
  const title = (id: number, s: string, color = '0F2B46') => text(id, 48, 40, 864, 70, [{ align: 'l', runs: [{ text: s, b: true, sz: '3600', color }] }])
  const card = (id: number, x: number, heading: string, body: string, accent: string): PptxShapeSpec =>
    text(id, x, 150, 268, 200, [
      { align: 'l', runs: [{ text: heading, b: true, sz: '2200', color: accent }] },
      { align: 'l', runs: [{ text: body, sz: '1600', color: '3B3A39' }] },
    ], 'F3F6FA')

  const slides: PptxSlideSpec[] = [
    {
      layout: [0, 0],
      shapes: [
        text(2, 64, 170, 832, 90, [{ align: 'l', runs: [{ text: 'Introducing Fleet Dashboard 2.0', b: true, sz: '4800', color: 'FFFFFF' }] }]),
        text(3, 64, 262, 832, 50, [{ align: 'l', runs: [{ text: 'Every sensor, every site, one view. Launch plan for Q4 2026.', sz: '2200', color: 'C7E0F4' }] }]),
        text(4, 64, 470, 600, 30, [{ align: 'l', runs: [{ text: 'Halcyon Instruments  ·  Product marketing', sz: '1400', color: '9FC5E8' }] }]),
      ],
      notes: [['Welcome everyone. This deck shows the officeview slideshow: transitions, click builds, Morph and speaker notes.'], ['Press N or the right arrow to advance.']],
    },
    {
      layout: [0, 1],
      transition: { xml: '<p:transition spd="slow"><p:push dir="u"/></p:transition>' },
      shapes: [
        title(2, 'Why customers are asking for it'),
        card(3, 48, 'Too many tabs', 'Operators watch an average of 6 tools to follow one fleet. Alerts arrive late or twice.', '0F6CBD'),
        card(4, 346, 'Slow answers', 'A question like "which gateways dropped out overnight?" takes a support ticket today.', 'C50F1F'),
        card(5, 644, 'One view', 'Dashboard 2.0 puts live status, history and alerts on one screen, on any device.', '107C41'),
      ],
      timingXml: timingXmlOf(mainSeqXml([fadeIn(3, 3), fadeIn(8, 4), fadeIn(13, 5)])),
      notes: [['Each card fades in on a click. These are click builds from the slide timing.']],
    },
    {
      layout: [0, 1],
      transition: p14('<p:transition spd="slow" p14:dur="1600"><p14:vortex dir="r"/></p:transition>'),
      shapes: [
        title(2, 'Launch timeline'),
        {
          id: 3, off: [pt(48), pt(140)], ext: [pt(864), pt(260)],
          table: {
            colWidths: [pt(330), pt(178), pt(178), pt(178)],
            firstRow: true, bandRow: true, styleId: 'LaunchTable',
            rows: [
              { h: pt(44), cells: ['Workstream', 'October', 'November', 'December'].map((t) => ({ paragraphs: [{ align: 'l', runs: [{ text: t, b: true, sz: '1600', color: 'FFFFFF' }] }] })) },
              ...[['Private beta, 40 customers', 'Running', 'Feedback', '—'], ['Pricing and packaging', 'Approved', '—', '—'], ['Docs and training', 'Drafting', 'Review', 'Published'], ['General availability', '—', '—', 'Dec 9']]
                .map((row) => ({ h: pt(48), cells: row.map((t, i) => ({ paragraphs: [{ align: 'l', runs: [{ text: t, b: i === 0, sz: '1500', color: '242424' }] }] })) })),
            ],
          },
        },
      ],
      notes: [['This slide arrived with the Vortex 3-D transition, drawn as a WebGL2 shader.']],
    },
    {
      layout: [0, 1],
      transition: p14('<p:transition spd="slow" p14:dur="1400"><p14:prism isContent="1"/></p:transition>'),
      shapes: [
        title(2, 'Three layers, one model'),
        text(3, 70, 190, 150, 150, [{ align: 'ctr', runs: [{ text: 'Sensors', b: true, sz: '1800', color: 'FFFFFF' }] }], '0F4C81', '!!sensors', 'ellipse'),
        text(4, 250, 190, 150, 150, [{ align: 'ctr', runs: [{ text: 'Gateways', b: true, sz: '1800', color: 'FFFFFF' }] }], '2E86C1', '!!gateways', 'ellipse'),
        text(5, 430, 190, 150, 150, [{ align: 'ctr', runs: [{ text: 'Cloud', b: true, sz: '1800', color: '0F2B46' }] }], '85C1E9', '!!cloud', 'ellipse'),
        text(6, 620, 170, 300, 200, [{ align: 'l', runs: [{ text: 'Data flows from the edge to the cloud in under two seconds.', sz: '1800', color: '3B3A39' }] }], undefined, '!!caption'),
      ],
      notes: [['Next slide uses Morph: the three circles keep their names and move into a stack.']],
    },
    {
      layout: [0, 1],
      transition: { xml: '<p:transition spd="slow"><p159:morph xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" option="byObject"/></p:transition>', wrapInAlternateContent: { requires: 'p159', fallbackXml: '<p:transition spd="slow"><p:fade/></p:transition>' } },
      shapes: [
        title(2, 'Three layers, one model'),
        text(3, 600, 110, 300, 90, [{ align: 'ctr', runs: [{ text: 'Sensors', b: true, sz: '2000', color: 'FFFFFF' }] }], '0F4C81', '!!sensors', 'roundRect'),
        text(4, 600, 220, 300, 90, [{ align: 'ctr', runs: [{ text: 'Gateways', b: true, sz: '2000', color: 'FFFFFF' }] }], '2E86C1', '!!gateways', 'roundRect'),
        text(5, 600, 330, 300, 90, [{ align: 'ctr', runs: [{ text: 'Cloud', b: true, sz: '2000', color: '0F2B46' }] }], '85C1E9', '!!cloud', 'roundRect'),
        text(6, 60, 170, 480, 200, [{ align: 'l', runs: [{ text: 'Morph moved, resized and reshaped the shapes the two slides share, matched by their !! names.', sz: '2000', color: '3B3A39' }] }], undefined, '!!caption'),
      ],
      notes: [['Morph matches shapes named with a !! prefix and animates position, size and colour.']],
    },
    {
      layout: [0, 0],
      transition: { xml: '<p:transition spd="slow"><p15:prstTrans prst="curtains"/></p:transition>', wrapInAlternateContent: { requires: 'p15', fallbackXml: '<p:transition spd="slow"><p:fade/></p:transition>' } },
      shapes: [
        text(2, 64, 190, 832, 90, [{ align: 'l', runs: [{ text: 'Thank you', b: true, sz: '5400', color: 'FFFFFF' }] }]),
        text(3, 64, 285, 832, 50, [{ align: 'l', runs: [{ text: 'Questions: launch@halcyon.example', sz: '2200', color: 'C7E0F4' }] }]),
      ],
      notes: [['Closing slide, reached with the Curtains transition. Press Esc to leave the show.']],
    },
  ]
  const bytes = await buildPptxDeck({
    slideSize: [W, H],
    masters: [{
      bg: { solid: 'FFFFFF' },
      layouts: [
        { bg: { gradient: { angle: 135, stops: [[0, '0F2B46'], [100000, '0F6CBD']] } }, showMasterSp: false },
        { bg: { solid: 'FFFFFF' }, shapes: [{ prst: 'rect', off: ['0', '0'], ext: [String(W), pt(10)], fill: '0F6CBD' }, { prst: 'rect', off: [pt(48), pt(510)], ext: [pt(864), '12700'], fill: 'E1DFDD' }] },
      ],
    }],
    slides,
  })
  return noOutlines(bytes)
}

// ---------------------------------------------------------------------------
// Writer (.odt) and RTF
// ---------------------------------------------------------------------------

async function notesOdt(): Promise<Uint8Array> {
  return buildOdt({
    extraAutoStyles: odfBulletListStyle('Bullets'),
    extraCommonStyles: odfWriterHeadingStyles() +
      '<style:style style:name="Heading_20_2" style:display-name="Heading 2" style:family="paragraph" style:parent-style-name="Heading" style:default-outline-level="2">' +
      '<style:text-properties fo:font-size="115%" fo:font-weight="bold" fo:color="#0F4C81"/></style:style>',
    pageLayouts: { pm1: { widthCm: 21, heightCm: 29.7, marginCm: 2.2 } },
    masters: {
      Standard: {
        layout: 'pm1',
        footer: [{ align: 'center', runs: [{ text: 'Design review notes · page ', color: '8A94A6', sizePt: 9 }, { text: '1', field: 'PAGE', color: '8A94A6', sizePt: 9 }] }],
      },
    },
    paras: [
      { text: 'Design review: gateway G3 enclosure', heading: 1, style: 'Heading_20_1' },
      { runs: [{ text: 'Date: ', bold: true }, { text: '21 October 2026    ' }, { text: 'Attendees: ', bold: true }, { text: 'hardware, manufacturing, support' }], spacingAfterPt: 10 },
      { text: 'Decisions', heading: 2, style: 'Heading_20_2' },
      { runs: [{ text: 'The team agreed to move from a two-part clamshell to a ' }, { text: 'single extruded body', bold: true }, { text: ' with end caps. Tooling cost goes up, but assembly time drops by about a third and the seal is easier to test.' }], spacingAfterPt: 8 },
      { runs: [{ text: 'This file is an OpenDocument Text document (', italic: true }, { text: '.odt', italic: true, bold: true }, { text: '), rendered by the same layout engine as Word files.', italic: true }], spacingAfterPt: 8 },
      { text: 'Open questions', heading: 2, style: 'Heading_20_2' },
    ],
    lists: [{ styleName: 'Bullets', items: [{ text: 'Can the supplier hold ±0.1 mm on the extrusion wall?' }, { text: 'Do we need a second gasket for the IP67 rating?' }, { text: 'Which colour does marketing want for the end caps?' }] }],
    tables: [{
      colWidths: ['7cm', '5cm', '4.6cm'],
      cellBorders: 'fo:border="0.5pt solid #C9D1DC"',
      headerRows: [{ cells: [{ text: 'Action', fill: 'E8EEF6' }, { text: 'Owner', fill: 'E8EEF6' }, { text: 'Due', fill: 'E8EEF6' }] }],
      rows: [
        { cells: [{ text: 'Request extrusion samples' }, { text: 'Manufacturing' }, { text: '4 November' }] },
        { cells: [{ text: 'Run thermal simulation' }, { text: 'Hardware' }, { text: '11 November' }] },
        { cells: [{ text: 'Update field service guide' }, { text: 'Support' }, { text: '2 December' }] },
      ],
    }],
  })
}

function letterRtf(): Uint8Array {
  const hex = Array.from(logoMark()).map((b) => b.toString(16).padStart(2, '0')).join('').replace(/(.{96})/g, '$1\n')
  const rtf = [
    '{\\rtf1\\ansi\\ansicpg1252\\deff0',
    '{\\fonttbl{\\f0\\fswiss\\fcharset0 Arial;}{\\f1\\froman\\fcharset0 Georgia;}}',
    '{\\colortbl ;\\red15\\green76\\blue129;\\red96\\green94\\blue92;\\red232\\green238\\blue246;\\red16\\green124\\blue65;}',
    '\\paperw12240\\paperh15840\\margl1440\\margr1440\\margt1200\\margb1200',
    `\\pard\\sa120{\\pict\\pngblip\\picw96\\pich96\\picwgoal720\\pichgoal720\n${hex}}\\par`,
    '\\pard\\sa0\\b\\fs28\\cf1 Halcyon Instruments\\b0\\cf0\\fs20\\par',
    '\\pard\\sa360\\cf2 400 Harbor Way, Suite 12  \\u183 ?  Portland, OR 97209\\cf0\\par',
    '\\pard\\sa240\\fs21 24 October 2026\\par',
    '\\pard\\sa240 Dear customer,\\par',
    '\\pard\\sa200\\sl300\\slmult1 Thank you for joining the private beta of \\b Fleet Dashboard 2.0\\b0 . This letter is a Rich Text Format (\\i .rtf\\i0 ) file, parsed straight from its byte stream and laid out by the same engine that renders Word documents.\\par',
    '\\pard\\sa200\\sl300\\slmult1 Your beta account includes the following:\\par',
    '\\trowd\\trgaph108\\trleft0',
    '\\clcbpat3\\clbrdrt\\brdrs\\brdrcf2\\clbrdrl\\brdrs\\brdrcf2\\clbrdrb\\brdrs\\brdrcf2\\clbrdrr\\brdrs\\brdrcf2\\cellx4680',
    '\\clcbpat3\\clbrdrt\\brdrs\\brdrcf2\\clbrdrl\\brdrs\\brdrcf2\\clbrdrb\\brdrs\\brdrcf2\\clbrdrr\\brdrs\\brdrcf2\\cellx9360',
    '\\intbl\\b Item\\b0\\cell\\intbl\\b Included\\b0\\cell\\row',
    '\\trowd\\trgaph108\\trleft0',
    '\\clbrdrt\\brdrs\\brdrcf2\\clbrdrl\\brdrs\\brdrcf2\\clbrdrb\\brdrs\\brdrcf2\\clbrdrr\\brdrs\\brdrcf2\\cellx4680',
    '\\clbrdrt\\brdrs\\brdrcf2\\clbrdrl\\brdrs\\brdrcf2\\clbrdrb\\brdrs\\brdrcf2\\clbrdrr\\brdrs\\brdrcf2\\cellx9360',
    '\\intbl Dashboard seats\\cell\\intbl 25\\cell\\row',
    '\\trowd\\trgaph108\\trleft0',
    '\\clbrdrt\\brdrs\\brdrcf2\\clbrdrl\\brdrs\\brdrcf2\\clbrdrb\\brdrs\\brdrcf2\\clbrdrr\\brdrs\\brdrcf2\\cellx4680',
    '\\clbrdrt\\brdrs\\brdrcf2\\clbrdrl\\brdrs\\brdrcf2\\clbrdrb\\brdrs\\brdrcf2\\clbrdrr\\brdrs\\brdrcf2\\cellx9360',
    '\\intbl Data retention\\cell\\intbl 13 months\\cell\\row',
    '\\trowd\\trgaph108\\trleft0',
    '\\clbrdrt\\brdrs\\brdrcf2\\clbrdrl\\brdrs\\brdrcf2\\clbrdrb\\brdrs\\brdrcf2\\clbrdrr\\brdrs\\brdrcf2\\cellx4680',
    '\\clbrdrt\\brdrs\\brdrcf2\\clbrdrl\\brdrs\\brdrcf2\\clbrdrb\\brdrs\\brdrcf2\\clbrdrr\\brdrs\\brdrcf2\\cellx9360',
    '\\intbl Support\\cell\\intbl \\cf4\\b Named engineer\\b0\\cf0\\cell\\row',
    '\\pard\\sa200\\sb240\\sl300\\slmult1 We will send a short survey at the end of each month. If anything breaks, reply to this letter or write to beta@halcyon.example.\\par',
    '\\pard\\sa0 Kind regards,\\par',
    '\\pard\\sa0\\sb360\\b Mara Ellison\\b0\\par',
    '\\pard\\cf2 Head of Product, Halcyon Instruments\\cf0\\par',
    '}',
  ].join('\n')
  return new TextEncoder().encode(rtf)
}

// ---------------------------------------------------------------------------
// Charts: the repo's synthetic chart gallery, one chart per family
// ---------------------------------------------------------------------------

/**
 * Every natively drawn chart family in one Word document. Data labels are
 * stripped from the classic charts, waterfall and funnel: in browsers they
 * currently hit a measurement bug in the chart engine and fall back to a card.
 */
async function chartsDocx(): Promise<Uint8Array> {
  const workbook = await galleryHierarchyWorkbook()
  const charts = [...classicGallery(), ...chartExGallery(true)].map((chart) => ({
    name: chart.name,
    part: chart.hierarchy ? chart.part : chart.part.replace(/<c:dLbls>[\s\S]*?<\/c:dLbls>/g, '').replace(/<cx:dataLabels\b[\s\S]*?<\/cx:dataLabels>/g, '').replace('<a:t>with labels</a:t>', '<a:t>clustered</a:t>'),
    ...(chart.hierarchy ? { workbook } : {}),
  }))
  return buildChartDocument('docx', charts, { columns: 2, width: 336, height: 228 })
}

// ---------------------------------------------------------------------------

const samples: Array<[string, () => Promise<Uint8Array> | Uint8Array]> = [
  ['quarterly-review.docx', reportDocx],
  ['operating-budget.xlsx', budgetXlsx],
  ['sales-ledger.xlsx', ledgerXlsx],
  ['product-launch.pptx', launchPptx],
  ['chart-gallery.docx', chartsDocx],
  ['design-review.odt', notesOdt],
  ['beta-welcome.rtf', letterRtf],
]

mkdirSync(OUT, { recursive: true })
for (const [name, make] of samples) {
  const bytes = await make()
  writeFileSync(join(OUT, name), bytes)
  console.log(`${name.padEnd(24)} ${(bytes.length / 1024).toFixed(1).padStart(7)} KiB`)
}
