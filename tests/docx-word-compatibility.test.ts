// @vitest-environment node
import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { layoutDocx, type MeasureFn } from '../src/docx/layout'

const measureFixed: MeasureFn = (text, style) => {
  return text.length * style.fontSizePt * 0.6 * (96 / 72)
}

describe('Phase 22: Word Compatibility Gaps', () => {
  test('Footnotes/endnotes parts parse with reference markers and layout (currently absent)', async () => {
    const zip = new JSZip()

    // Package relationships
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
    )

    // Document relationships: point to footnotes and endnotes parts
    zip.file(
      'word/_rels/document.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>`
    )

    // Document body with footnoteReference and endnoteReference
    zip.file(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Statement needing clarification</w:t></w:r>
      <w:r><w:footnoteReference w:id="1"/></w:r>
      <w:r><w:t> and endnote citation</w:t></w:r>
      <w:r><w:endnoteReference w:id="1"/></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`
    )

    // Footnotes part
    zip.file(
      'word/footnotes.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:id="1">
    <w:p>
      <w:r><w:footnoteRef/></w:r>
      <w:r><w:t> Clarification detail in footnote.</w:t></w:r>
    </w:p>
  </w:footnote>
</w:footnotes>`
    )

    // Endnotes part
    zip.file(
      'word/endnotes.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="1">
    <w:p>
      <w:r><w:endnoteRef/></w:r>
      <w:r><w:t> Citation source in endnote.</w:t></w:r>
    </w:p>
  </w:endnote>
</w:endnotes>`
    )

    const buffer = await zip.generateAsync({ type: 'uint8array' })
    const pkg = await OfficePackage.load(buffer)
    const doc = await parseDocx(pkg)

    // 1. Footnotes & Endnotes parts parsed
    expect(doc.footnotes).toBeDefined()
    expect(doc.footnotes?.length).toBeGreaterThanOrEqual(1)
    const fn1 = doc.footnotes?.find((fn) => fn.id === 1)
    expect(fn1).toBeDefined()
    expect(fn1?.paragraphs[0].runs.map((r) => r.text).join('')).toContain('Clarification detail in footnote')

    expect(doc.endnotes).toBeDefined()
    expect(doc.endnotes?.length).toBeGreaterThanOrEqual(1)
    const en1 = doc.endnotes?.find((en) => en.id === 1)
    expect(en1).toBeDefined()
    expect(en1?.paragraphs[0].runs.map((r) => r.text).join('')).toContain('Citation source in endnote')

    // 2. Reference markers in body document runs
    const bodyPara = doc.sections[0].paragraphs[0]
    const fnRefRun = bodyPara.runs.find((r) => r.footnoteReference?.id === 1)
    expect(fnRefRun).toBeDefined()
    expect(fnRefRun?.text).toBe('1')

    const enRefRun = bodyPara.runs.find((r) => r.endnoteReference?.id === 1)
    expect(enRefRun).toBeDefined()
    expect(enRefRun?.text).toBe('1')

    // 3. Layout renders footnote on the page
    const pages = layoutDocx(doc, measureFixed)
    expect(pages.length).toBeGreaterThanOrEqual(1)

    // Verify footnote is laid out on the page (in page.footnotes or page.lines)
    const allText = pages[0].lines.flatMap((l) => l.segs.map((s) => s.text)).join(' ')
    const hasInLines = allText.includes('Clarification detail in footnote')
    const hasInFootnotes = pages[0].footnotes?.lines?.some((l) =>
      l.segs.map((s) => s.text).join('').includes('Clarification detail in footnote')
    )
    expect(hasInLines || hasInFootnotes).toBe(true)
  })

  test('Comment parts parse and render or diagnose explicitly (currently absent)', async () => {
    const zip = new JSZip()

    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
    )

    zip.file(
      'word/_rels/document.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>`
    )

    zip.file(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t>Claimed fact here</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
      <w:r><w:commentReference w:id="0"/></w:r>
    </w:p>
  </w:body>
</w:document>`
    )

    zip.file(
      'word/comments.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Alice Reviewer" w:date="2026-10-09T18:00:00Z" w:initials="AR">
    <w:p>
      <w:r><w:t>Please provide a source for this claim.</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>`
    )

    const buffer = await zip.generateAsync({ type: 'uint8array' })
    const pkg = await OfficePackage.load(buffer)
    const doc = await parseDocx(pkg)

    // 1. Comments part parsed
    expect(doc.comments).toBeDefined()
    expect(doc.comments).toHaveLength(1)
    const c0 = doc.comments![0]
    expect(c0.id).toBe(0)
    expect(c0.author).toBe('Alice Reviewer')
    expect(c0.initials).toBe('AR')
    expect(c0.text).toBe('Please provide a source for this claim.')

    // 2. Explicit diagnostic surfaced
    expect(doc.commentDiagnostics).toBeDefined()
    expect(doc.commentDiagnostics![0]).toEqual({
      id: 0,
      author: 'Alice Reviewer',
      text: 'Please provide a source for this claim.'
    })

    // 3. Comment reference run linked in body paragraph
    const bodyPara = doc.sections[0].paragraphs[0]
    const commentRefRun = bodyPara.runs.find((r) => r.commentReference?.id === 0)
    expect(commentRefRun).toBeDefined()

    // 4. Comment reference carries the id pointer but no invented inline text
    expect(commentRefRun?.text).toBe('')
    const pages = layoutDocx(doc, measureFixed)
    expect(pages.length).toBeGreaterThanOrEqual(1)
    const textRendered = pages[0].lines.flatMap((l) => l.segs.map((s) => s.text)).join('')
    expect(textRendered).toContain('Claimed fact here')
    expect(textRendered).not.toContain('Comment')
  })

  test('BiDi paragraphs (w:bidi) render in visual order', async () => {
    const zip = new JSZip()

    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
    )

    zip.file(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:bidi w:val="1"/>
      </w:pPr>
      <w:r><w:t>RightFirst</w:t></w:r>
      <w:r><w:t> </w:t></w:r>
      <w:r><w:t>LeftSecond</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
    )

    const buffer = await zip.generateAsync({ type: 'uint8array' })
    const pkg = await OfficePackage.load(buffer)
    const doc = await parseDocx(pkg)

    // 1. BiDi paragraph flag and default right alignment
    const para = doc.sections[0].paragraphs[0]
    expect(para.bidi).toBe(true)
    expect(para.align).toBe('right')

    // 2. Layout renders in visual order
    const pages = layoutDocx(doc, measureFixed)
    expect(pages.length).toBeGreaterThanOrEqual(1)
    const line = pages[0].lines[0]
    expect(line).toBeDefined()
    expect(line.bidi).toBe(true)

    // In visual order for RTL/BiDi: the first logical word (RightFirst) has a higher penOffset/x than LeftSecond
    const segRight = line.segs.find((s) => s.text === 'RightFirst')
    const segLeft = line.segs.find((s) => s.text === 'LeftSecond')
    expect(segRight).toBeDefined()
    expect(segLeft).toBeDefined()
    expect(segRight!.penOffset!).toBeGreaterThan(segLeft!.penOffset!)
  })

  test('Complex-script and East-Asia run fonts (w:cs, w:eastAsia, w:hAnsi) resolve instead of falling back to ascii only', async () => {
    const zip = new JSZip()

    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
    )

    zip.file(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <!-- 1. Pure EastAsia without ascii -->
      <w:r>
        <w:rPr><w:rFonts w:eastAsia="SimSun"/></w:rPr>
        <w:t>你好</w:t>
      </w:r>
      <!-- 2. Pure CS without ascii -->
      <w:r>
        <w:rPr><w:rFonts w:cs="Arial Unicode MS"/></w:rPr>
        <w:t>مرحبا</w:t>
      </w:r>
      <!-- 3. Pure hAnsi without ascii -->
      <w:r>
        <w:rPr><w:rFonts w:hAnsi="Times New Roman"/></w:rPr>
        <w:t>Bonjour</w:t>
      </w:r>
      <!-- 4. Mixed ascii + eastAsia on East Asian text -->
      <w:r>
        <w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="MS Gothic"/></w:rPr>
        <w:t>日本語テキスト</w:t>
      </w:r>
      <!-- 5. Mixed ascii + cs on Arabic text with rtl -->
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Calibri" w:cs="Traditional Arabic"/>
          <w:rtl/>
        </w:rPr>
        <w:t>عربي</w:t>
      </w:r>
      <!-- 6. ascii + eastAsia on ASCII text uses ascii -->
      <w:r>
        <w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="MS Gothic"/></w:rPr>
        <w:t>Pure English</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`
    )

    const buffer = await zip.generateAsync({ type: 'uint8array' })
    const pkg = await OfficePackage.load(buffer)
    const doc = await parseDocx(pkg)

    const runs = doc.sections[0].paragraphs[0].runs
    expect(runs).toHaveLength(6)

    // 1. Pure EastAsia
    expect(runs[0].fontFamily).toBe('SimSun')

    // 2. Pure CS
    expect(runs[1].fontFamily).toBe('Arial Unicode MS')

    // 3. Pure hAnsi
    expect(runs[2].fontFamily).toBe('Times New Roman')

    // 4. East Asian text with both ascii and eastAsia resolves to eastAsia font
    expect(runs[3].fontFamily).toBe('MS Gothic')

    // 5. CS/RTL text with both ascii and cs resolves to cs font
    expect(runs[4].fontFamily).toBe('Traditional Arabic')

    // 6. ASCII text with both retains ascii font
    expect(runs[5].fontFamily).toBe('Calibri')
  })

  test('Drop caps and hyphenation degrade gracefully with explicit coverage', async () => {
    const zip = new JSZip()

    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
    )

    zip.file(
      'word/_rels/document.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`
    )

    zip.file(
      'word/settings.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:autoHyphenation w:val="1"/>
</w:settings>`
    )

    zip.file(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <!-- Paragraph 1: Drop cap -->
    <w:p>
      <w:pPr>
        <w:framePr w:dropCap="drop" w:lines="3"/>
      </w:pPr>
      <w:r><w:t>Once upon a time in a faraway kingdom.</w:t></w:r>
    </w:p>
    <!-- Paragraph 2: Soft hyphen and suppressAutoHyphens -->
    <w:p>
      <w:pPr>
        <w:suppressAutoHyphens w:val="1"/>
      </w:pPr>
      <w:r>
        <w:t>super</w:t>
        <w:softHyphen/>
        <w:t>fragilistic</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`
    )

    const buffer = await zip.generateAsync({ type: 'uint8array' })
    const pkg = await OfficePackage.load(buffer)
    const doc = await parseDocx(pkg)

    // 1. Document-level autoHyphenation parsed from settings
    expect(doc.autoHyphenation).toBe(true)

    // 2. Drop cap parsed with explicit diagnostic
    const p1 = doc.sections[0].paragraphs[0]
    expect(p1.dropCap).toBe('drop')
    expect(p1.diagnostics).toBeDefined()
    expect(p1.diagnostics?.some((d) => d.feature === 'drop-cap')).toBe(true)

    // 3. Drop cap paragraph text fully preserved in layout without dropping the first letter
    const pages = layoutDocx(doc, measureFixed)
    expect(pages.length).toBeGreaterThanOrEqual(1)
    const p1Text = pages[0].lines[0].segs.map((s) => s.text).join('')
    expect(p1Text).toContain('Once upon a time')

    // 4. Paragraph 2 suppressAutoHyphens parsed
    const p2 = doc.sections[0].paragraphs[1]
    expect(p2.suppressAutoHyphens).toBe(true)

    // 5. Soft hyphen in unbroken text renders cleanly without rogue hyphen glyph
    const p2Runs = p2.runs.map((r) => r.text).join('')
    expect(p2Runs).toContain('\u00AD')
    const allRenderedText = pages[0].lines.flatMap((l) => l.segs.map((s) => s.text)).join('')
    expect(allRenderedText).toContain('superfragilistic')
  })
})

const REL_HEAD = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

async function buildDocx(files: Record<string, string>): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('_rels/.rels', `${REL_HEAD}<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  for (const [path, body] of Object.entries(files)) zip.file(path, body)
  return zip.generateAsync({ type: 'uint8array' })
}

const relsFor = (...parts: Array<[string, string]>) =>
  `${REL_HEAD}${parts.map(([type, target], i) => `<Relationship Id="rIdX${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`).join('')}</Relationships>`

describe('Phase 22 review remediation', () => {
  test('C1: malformed auxiliary parts do not fail the whole document', async () => {
    const buf = await buildDocx({
      'word/_rels/document.xml.rels': relsFor(['footnotes', 'footnotes.xml'], ['endnotes', 'endnotes.xml'], ['comments', 'comments.xml'], ['settings', 'settings.xml']),
      'word/document.xml': `<w:document ${W_NS}><w:body><w:p><w:r><w:t>Still renders</w:t></w:r></w:p></w:body></w:document>`,
      'word/footnotes.xml': '<w:footnotes <<<broken',
      'word/endnotes.xml': '<<<broken',
      'word/comments.xml': '<<<broken',
      'word/settings.xml': '<<<broken',
    })
    const doc = await parseDocx(await OfficePackage.load(buf))
    expect(doc.sections[0].paragraphs[0].runs[0].text).toBe('Still renders')
    expect(doc.footnotes).toBeUndefined()
    expect(doc.comments).toBeUndefined()
  })

  test('C2/M3: bidi lines paint with explicit left textAlign and restore direction', async () => {
    const { createCanvas } = await import('canvas')
    const { renderPages } = await import('../src/docx/layout')
    const buf = await buildDocx({
      'word/document.xml': `<w:document ${W_NS}><w:body><w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>AAA</w:t></w:r></w:p></w:body></w:document>`,
    })
    const doc = await parseDocx(await OfficePackage.load(buf))
    const pages = layoutDocx(doc, measureFixed)
    const canvas = createCanvas(pages[0].widthPx, pages[0].heightPx)
    const ctx = canvas.getContext('2d')
    const aligns: string[] = []
    const original = ctx.fillText.bind(ctx)
    ctx.fillText = ((text: string, x: number, y: number) => {
      aligns.push(ctx.textAlign)
      original(text, x, y)
    }) as typeof ctx.fillText
    renderPages(pages, ctx as unknown as CanvasRenderingContext2D)
    expect(aligns.length).toBeGreaterThan(0)
    expect(aligns.every((a) => a === 'left')).toBe(true)
    expect(ctx.direction).toBe('ltr')
  })

  test('M1: endnotes are laid out on the last page', async () => {
    const buf = await buildDocx({
      'word/_rels/document.xml.rels': relsFor(['endnotes', 'endnotes.xml']),
      'word/document.xml': `<w:document ${W_NS}><w:body><w:p><w:r><w:t>Body</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p></w:body></w:document>`,
      'word/endnotes.xml': `<w:endnotes ${W_NS}><w:endnote w:id="2"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t> Source text</w:t></w:r></w:p></w:endnote></w:endnotes>`,
    })
    const doc = await parseDocx(await OfficePackage.load(buf))
    const pages = layoutDocx(doc, measureFixed)
    const last = pages[pages.length - 1]
    const text = last.endnotes?.lines.map((l) => l.segs.map((s) => s.text).join('')).join('') ?? ''
    expect(text).toContain('Source text')
  })

  test('m1: separator notes are never laid out as footnotes', async () => {
    const buf = await buildDocx({
      'word/_rels/document.xml.rels': relsFor(['footnotes', 'footnotes.xml']),
      'word/document.xml': `<w:document ${W_NS}><w:body><w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="-1"/></w:r></w:p></w:body></w:document>`,
      'word/footnotes.xml': `<w:footnotes ${W_NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:t>SEPARATOR-CONTENT</w:t></w:r></w:p></w:footnote></w:footnotes>`,
    })
    const doc = await parseDocx(await OfficePackage.load(buf))
    const pages = layoutDocx(doc, measureFixed)
    expect(pages[0].footnotes).toBeUndefined()
  })

  test('m6: note reference markers are superscript and smaller', async () => {
    const buf = await buildDocx({
      'word/document.xml': `<w:document ${W_NS}><w:body><w:p><w:r><w:t>Hi</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p></w:body></w:document>`,
    })
    const doc = await parseDocx(await OfficePackage.load(buf))
    const ref = doc.sections[0].paragraphs[0].runs.find((r) => r.footnoteReference)!
    expect(ref.vertAlign).toBe('superscript')
    const pages = layoutDocx(doc, measureFixed)
    const segs = pages[0].lines[0].segs
    const refSeg = segs.find((s) => s.run.footnoteReference)!
    const textSeg = segs.find((s) => s.text === 'Hi')!
    expect(refSeg.style.fontSizePt).toBeLessThan(textSeg.style.fontSizePt)
  })

  test('m7/m8: invalid note type and fractional ids are rejected', async () => {
    const buf = await buildDocx({
      'word/_rels/document.xml.rels': relsFor(['footnotes', 'footnotes.xml']),
      'word/document.xml': `<w:document ${W_NS}><w:body><w:p><w:r><w:footnoteReference w:id="1.5"/></w:r></w:p></w:body></w:document>`,
      'word/footnotes.xml': `<w:footnotes ${W_NS}><w:footnote w:id="1.5"><w:p><w:r><w:t>x</w:t></w:r></w:p></w:footnote><w:footnote w:id="3" w:type="garbage"><w:p><w:r><w:t>y</w:t></w:r></w:p></w:footnote></w:footnotes>`,
    })
    const doc = await parseDocx(await OfficePackage.load(buf))
    expect(doc.footnotes?.map((n) => n.id)).toEqual([3])
    expect(doc.footnotes?.[0].type).toBe('normal')
    expect(doc.sections[0].paragraphs[0].runs.some((r) => r.footnoteReference)).toBe(false)
  })

  test('m10: explicit rFonts hint overrides script detection', async () => {
    const buf = await buildDocx({
      'word/document.xml': `<w:document ${W_NS}><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="MS Gothic" w:hint="eastAsia"/></w:rPr><w:t>Latin only</w:t></w:r></w:p></w:body></w:document>`,
    })
    const doc = await parseDocx(await OfficePackage.load(buf))
    expect(doc.sections[0].paragraphs[0].runs[0].fontFamily).toBe('MS Gothic')
  })

  test('m12: multiple soft hyphens never paint the soft hyphen glyph', async () => {
    const buf = await buildDocx({
      'word/document.xml': `<w:document ${W_NS}><w:body><w:p><w:r><w:t>aaaa</w:t><w:softHyphen/><w:t>bbbb</w:t><w:softHyphen/><w:t>cccc</w:t></w:r></w:p></w:body></w:document>`,
    })
    const doc = await parseDocx(await OfficePackage.load(buf))
    const pages = layoutDocx(doc, measureFixed)
    const joined = pages[0].lines.map((l) => l.segs.map((s) => s.text).join('')).join('')
    expect(joined).not.toContain('\u00ad')
    expect(joined.replace(/-/g, '')).toBe('aaaabbbbcccc')
  })

  test('noBreakHyphen is parsed as U+2011', async () => {
    const buf = await buildDocx({
      'word/document.xml': `<w:document ${W_NS}><w:body><w:p><w:r><w:t>a</w:t><w:noBreakHyphen/><w:t>b</w:t></w:r></w:p></w:body></w:document>`,
    })
    const doc = await parseDocx(await OfficePackage.load(buf))
    expect(doc.sections[0].paragraphs[0].runs[0].text).toBe('a\u2011b')
  })
})
