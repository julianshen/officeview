import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { buildXlsx } from '../src/testdata/ooxml-builders'

it('decodes saved text once at each text node, without interpreting escaped literal prefixes or formulas',async()=>{
  const zip=await JSZip.loadAsync(await buildXlsx([{name:'S',rows:[{r:1,cells:[
    {ref:'A1',t:'s',v:0},{ref:'B1',t:'s',v:1},{ref:'C1',t:'str',v:'_xDE00_',formula:'MID("😀x",2,1)'},
    {ref:'D1',t:'str',v:'_x005F_x0041_'},{ref:'E1',t:'s',v:2},{ref:'F1',t:'s',v:3},
    {ref:'I1',t:'s',v:5},
  ]}]}],['_xD83D__xDE00_','_x005F_x0041_','_x000A_','_xZZZZ_','unused','_X0041_']))
  const sheet='xl/worksheets/sheet1.xml'
  zip.file(sheet,(await zip.file(sheet)!.async('string')).replace('</row>','<c r="G1" t="inlineStr"><is><r><t>_xD83D_</t></r><r><t>_xDE00_</t></r></is></c></row>'))
  zip.file('xl/sharedStrings.xml',(await zip.file('xl/sharedStrings.xml')!.async('string')).replace('</sst>','<si><r><t>_x00</t></r><r><t>41_</t></r></si></sst>'))
  zip.file(sheet,(await zip.file(sheet)!.async('string')).replace('</row>','<c r="H1" t="s"><v>6</v></c></row>'))
  const doc=await parseXlsx(await OfficePackage.load(await zip.generateAsync({type:'uint8array'})))
  const cells=doc.sheets[0].rows[0].cells
  expect(cells.map(c=>c.value)).toEqual(['😀','_x0041_','\uDE00','_x0041_','\n','_xZZZZ_','_X0041_','😀','_x0041_'])
  expect(cells[2].formula).toBe('MID("😀x",2,1)')
  expect(cells[3].valueIsError).toBe(false)
})
