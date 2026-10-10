import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import type { XlsxDocument } from '../src/xlsx/types'

const metadata=`<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:da="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><da:dynamicArrayProperties fDynamic="1"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`
async function fixture(meta=metadata,cm='1',formula='SEQUENCE(C1)',manual=false,related=true){
  const zip=await JSZip.loadAsync(await buildXlsx([{name:'S',rows:[{r:1,cells:[{ref:'A1',formula,v:1},{ref:'C1',v:3}]},{r:2,cells:[{ref:'A2',v:2,style:2}]},{r:3,cells:[{ref:'A3',v:3,style:2}]}]}]))
  const path='xl/worksheets/sheet1.xml'
  zip.file(path,(await zip.file(path)!.async('string')).replace('<c r="A1"','<c cm="'+cm+'" r="A1"').replace('<f>','<f t="array" ref="A1:A3">'))
  zip.file('xl/properties/native-metadata.xml',meta)
  if(related)zip.file('xl/_rels/workbook.xml.rels',(await zip.file('xl/_rels/workbook.xml.rels')!.async('string')).replace('</Relationships>','<Relationship Id="md" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="properties/native-metadata.xml"/></Relationships>'))
  if(manual)zip.file('xl/workbook.xml',(await zip.file('xl/workbook.xml')!.async('string')).replace('</workbook>','<calcPr calcMode="manual"/></workbook>'))
  return parseXlsx(await OfficePackage.load(await zip.generateAsync({type:'uint8array'})))
}
function at(doc:XlsxDocument,ref:string){return doc.sheets[0].rows.flatMap(r=>r.cells).find(c=>c.ref===ref)!}
describe('native dynamic array cell metadata',()=>{
  it('shrinks and grows saved spills, removes old results and retains styles',async()=>{
    const doc=await fixture();const style=at(doc,'A3').style
    at(doc,'C1').value=2;evaluateWorkbookFormulas(doc,{forceRecalc:true})
    expect(at(doc,'A2').value).toBe(2);expect(at(doc,'A3').value).toBeNull();expect(at(doc,'A3').style).toBe(style)
    at(doc,'C1').value=4;evaluateWorkbookFormulas(doc,{forceRecalc:true})
    expect(at(doc,'A3').value).toBe(3);expect(at(doc,'A4').value).toBe(4)
    at(doc,'C1').value=1;evaluateWorkbookFormulas(doc,{forceRecalc:true})
    expect(at(doc,'A2').value).toBeNull();expect(at(doc,'A3').value).toBeNull();expect(at(doc,'A4')).toBeUndefined()
  })
  it('resolves a follower dependency through its recalculated owner',async()=>{
    const doc=await fixture();at(doc,'C1').value=2
    doc.sheets[0].rows[0].cells.push({ref:'D1',col:3,row:0,styleIndex:0,value:null,formula:'SUM(A1:A3)'})
    evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(at(doc,'D1').value).toBe(3)
  })
  it('preserves an edited saved follower and reports a blocking spill',async()=>{
    const doc=await fixture();at(doc,'A2').value=99
    evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(at(doc,'A1').value).toBe('#SPILL!');expect(at(doc,'A2').value).toBe(99)
  })
  it('preserves a replacement cell even when its value matches the cache',async()=>{
    const doc=await fixture();const row=doc.sheets[0].rows[1];row.cells[0]={...row.cells[0]}
    evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(at(doc,'A1').value).toBe('#SPILL!');expect(at(doc,'A2').value).toBe(2)
  })
  it('retains a complete cache for unsupported dynamic formulas',async()=>{
    const doc=await fixture(metadata,'1','UnsupportedNativeFunction()');evaluateWorkbookFormulas(doc,{forceRecalc:true})
    expect([at(doc,'A1').value,at(doc,'A2').value,at(doc,'A3').value]).toEqual([1,2,3])
  })
  it('manual file loads retain the full saved spill until explicit recalc',async()=>{
    const doc=await fixture(metadata,'1','SEQUENCE(C1)',true);at(doc,'C1').value=2
    evaluateWorkbookFormulas(doc);expect(at(doc,'A3').value).toBe(3)
    evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(at(doc,'A3').value).toBeNull()
  })
  it.each([
    ['absent relationship',metadata,'1',false],['cm zero',metadata,'0',true],['cm suffix',metadata,'1x',true],
    ['wrong namespace',metadata.replace('2017/dynamicarray','2017/foreign'),'1',true],
    ['fDynamic false',metadata.replace('fDynamic="1"','fDynamic="0"'),'1',true],
    ['invalid type index',metadata.replace('t="1" v="0"','t="2" v="0"'),'1',true],
    ['invalid future index',metadata.replace('v="0"','v="1"'),'1',true],
    ['other metadata type',metadata.replace(/XLDAPR/g,'OTHER'),'1',true],
  ] as const)('keeps fixed legacy arrays with %s',async(_name,meta,cm,related)=>{
    const doc=await fixture(meta,cm,'SEQUENCE(C1)',false,related);at(doc,'C1').value=2
    evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(at(doc,'A3').value).toBe(3)
    expect(doc.diagnostics?.some(d=>d.feature==='legacy-array-output-mismatch')).toBe(true)
  })
  it('discovers a grown saved spill before an earlier reader outside the old range',async()=>{
    const doc=await fixture();at(doc,'C1').value=4
    doc.sheets[0].rows[0].cells.unshift({ref:'D1',col:3,row:0,styleIndex:0,value:null,formula:'A4'})
    evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(at(doc,'D1').value).toBe(4)
  })
  it('does not lease saved results inside a newly authored merge',async()=>{
    const doc=await fixture();doc.sheets[0].mergeRanges.push({minRow:1,maxRow:1,minCol:0,maxCol:1})
    evaluateWorkbookFormulas(doc,{forceRecalc:true});expect(at(doc,'A1').value).toBe('#SPILL!')
  })
})
