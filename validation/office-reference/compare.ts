import {loadPng, diffBitmaps, savePng} from '../../src/test/pixel-diff'
import {writeFileSync} from 'node:fs'
const dir=import.meta.dir
// REPORT.md explains why the DOCX table and XLSX print-page pairs are excluded.
const pairs=[['pydocx-having-images.word.pdf.0.png','pydocx-having-images.docx.officeview.0.png'],['poi-table_test.powerpoint.pdf.0.png','poi-table_test.pptx.officeview.0.png']]
const results=[]
for(const [ref,actual] of pairs){
 const a=await loadPng(`${dir}/${ref}`), b=await loadPng(`${dir}/${actual}`)
 const d=diffBitmaps(a,b,8)
 await savePng({width:d.width,height:d.height,data:d.mask},`${dir}/${actual}.diff.png`)
 results.push({reference:ref,actual,width:a.width,height:a.height,threshold:8,changedPixels:d.changedPixels,totalPixels:d.totalPixels,ratio:d.ratio})
}
writeFileSync(`${dir}/metrics.json`,JSON.stringify(results,null,2)+'\n')
console.log(JSON.stringify(results,null,2))
