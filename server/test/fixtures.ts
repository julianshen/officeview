/** Stage-0 hidden fixtures, with marker-bearing comment parts added to the two
 * comment cases (the generator's comment fixtures carry only the control). */
import JSZip from 'jszip'
import { hiddenFixtures, type Fixture } from '../../scripts/stream-s0/hidden-fixtures'

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
