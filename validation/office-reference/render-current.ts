import { readFileSync } from 'node:fs'
import { loadOfficeFile } from '../../src/components/OfficeFile'
import { getPaintables } from '../../src/render/paint'
import { renderPaintables, savePng } from '../../src/test/pixel-diff'
const output = import.meta.dir
for (const file of process.argv.slice(2)) {
 const doc = await loadOfficeFile(readFileSync(`${output}/../../corpus/${file}`))
 const units = await getPaintables(doc)
 const bitmaps = await renderPaintables(units)
 for (let i = 0; i < bitmaps.length; i++) await savePng(bitmaps[i], `${output}/${file}.officeview.${i}.png`)
 console.log(JSON.stringify({file,units: units.map(p=>p.spec)}))
}
