import { describe, expect, it } from 'vitest'
import { hidePivotFieldItems, pivotFieldXmlScope, replaceXmlExactlyOnce, selectSlicerItems, setSlicerItemStates } from '../scripts/xlsx-refs/control-interaction-xml'

describe('control interaction XML edits', () => {
  it('replaces only the intended part-local match and rejects zero or ambiguous targets', () => {
    const source = '<pivotField axis="axisCol"><items><item x="0"/><item x="1"/></items></pivotField>'
    expect(replaceXmlExactlyOnce(source, /<item x="1"\/>/, '<item h="1" x="1"/>', 'manual filter'))
      .toBe('<pivotField axis="axisCol"><items><item x="0"/><item h="1" x="1"/></items></pivotField>')
    expect(() => replaceXmlExactlyOnce(source, /<item x="9"\/>/, '<item h="1" x="9"/>', 'missing'))
      .toThrow('missing matched 0 XML targets')
    expect(() => replaceXmlExactlyOnce('<items><item x="1"/><item x="1"/></items>', /<item x="1"\/>/, '<item h="1" x="1"/>', 'ambiguous'))
      .toThrow('ambiguous matched 2 XML targets')
  })

  it('serializes exact slicer selection and no-data flags by item id', () => {
    const xml = '<cache><items count="3"><i x="0" s="1"/><i x="1"/><i x="2" nd="1"/></items></cache>'
    expect(setSlicerItemStates(xml, [
      { x: 0, selected: true },
      { x: 1, selected: false },
      { x: 2, selected: true, noData: true },
    ])).toBe('<cache><items count="3"><i x="0" s="1"/><i x="1"/><i x="2" s="1" nd="1"/></items></cache>')
    expect(() => setSlicerItemStates(xml, [{ x: 0, selected: true }])).toThrow('item count 3 does not match state count 1')
  })

  it('applies requested selections while preserving each item no-data marker', () => {
    const xml = '<cache><items count="3"><i x="0" s="1"/><i x="1"/><i x="2" nd="1"/></items></cache>'
    expect(selectSlicerItems(xml, [1, 2])).toBe('<cache><items count="3"><i x="0"/><i x="1" s="1"/><i x="2" s="1" nd="1"/></items></cache>')
  })

  it('hides only the requested item in the uniquely matching pivot axis field', () => {
    const xml = '<pivotFields><pivotField axis="axisRow"><items><item x="2"/></items></pivotField><pivotField axis="axisCol"><items><item x="1"/><item x="2"/></items></pivotField></pivotFields>'
    expect(hidePivotFieldItems(xml, 'axisCol', [2]))
      .toBe('<pivotFields><pivotField axis="axisRow"><items><item x="2"/></items></pivotField><pivotField axis="axisCol"><items><item x="1"/><item h="1" x="2"/></items></pivotField></pivotFields>')
    expect(() => hidePivotFieldItems(xml, 'axisCol', [9])).toThrow('item x=9 matched 0 XML targets')
  })

  it('checks saved item markers only inside their declared pivot axis field', () => {
    const xml = '<pivotFields><pivotField axis="axisRow"><items><item x="2"/></items></pivotField><pivotField axis="axisCol"><items><item x="2"/></items></pivotField></pivotFields>'
    expect(pivotFieldXmlScope(xml, 'axisCol')).toBe('<pivotField axis="axisCol"><items><item x="2"/></items></pivotField>')
    expect(pivotFieldXmlScope(xml, 'axisCol').split('<item x="2"/>').length - 1).toBe(1)
    expect(() => pivotFieldXmlScope(xml, 'axisPage')).toThrow('axisPage pivot field scope matched 0 XML targets')
  })
})
