/** Stage-0 hidden fixtures, with marker-bearing comment parts added to the two
 * comment cases (the generator's comment fixtures carry only the control). */
import JSZip from 'jszip'
import { hiddenFixtures, type Fixture } from '../../scripts/stream-s0/hidden-fixtures'
import { buildPptxDeck } from '../../src/testdata/ooxml-builders'

/** Three slides; the middle one is hidden by its slide root (`<p:sld show="0">`). No corpus deck has one. */
export const hiddenSlideDeck = (hidden = true) => buildPptxDeck({
  slides: ['One', 'Two (hidden)', 'Three'].map((text, i) => ({
    hidden: hidden && i === 1,
    shapes: [{ off: ['457200', '457200'] as [string, string], ext: ['4572000', '914400'] as [string, string], paragraphs: [{ runs: [{ text }] }] }],
  })),
})

export async function augmentedFixtures(): Promise<Array<Fixture & { withHidden: Uint8Array }>> {
  const out = []
  for (const f of await hiddenFixtures()) {
    const zip = await JSZip.loadAsync(f.withHidden)
    if (f.id === 'xlsx-comment') zip.file('xl/comments1.xml', `<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><commentList><comment ref="A1" authorId="0"><text><t>${f.marker}</t></text></comment></commentList></comments>`)
    if (f.id === 'docx-comment') zip.file('word/comments.xml', `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0"><w:p><w:r><w:t>${f.marker}</w:t></w:r></w:p></w:comment></w:comments>`)
    out.push({ ...f, withHidden: await zip.generateAsync({ type: 'uint8array' }) })
  }
  return out
}

/** The same deck with the middle slide hidden only by its presentation entry
 * (`<p:sldId … show="0"/>`); its slide root carries no `show`. */
export async function entryHiddenSlideDeck(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await hiddenSlideDeck(false))
  const pres = await zip.file('ppt/presentation.xml')!.async('string')
  const marked = pres.replace('<p:sldId id="257" ', '<p:sldId id="257" show="0" ')
  if (marked === pres) throw new Error('fixture: no second sldId')
  zip.file('ppt/presentation.xml', marked)
  return zip.generateAsync({ type: 'uint8array' })
}
