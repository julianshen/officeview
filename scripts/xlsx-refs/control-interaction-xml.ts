/** Apply an XML mutation only when its source construct occurs exactly once. */
export function replaceXmlExactlyOnce(
  xml: string,
  target: RegExp,
  replacement: string,
  label: string,
): string {
  if (target.global || target.sticky) throw new Error(`${label} matcher must not be global or sticky`)
  const matches = [...xml.matchAll(new RegExp(target.source, `${target.flags}g`))]
  if (matches.length !== 1) throw new Error(`${label} matched ${matches.length} XML targets`)
  const match = matches[0]!
  if (match.index === undefined) throw new Error(`${label} matcher did not capture a range`)
  return xml.slice(0, match.index) + replacement + xml.slice(match.index + match[0].length)
}

export interface SlicerItemState {
  x: number
  selected: boolean
  noData?: boolean
}

/** Return one pivot field's complete XML block, selected by its axis attribute. */
export function pivotFieldXmlScope(xml: string, axis: 'axisRow' | 'axisCol' | 'axisPage'): string {
  const fields = [...xml.matchAll(/<pivotField\b([^>]*)>[\s\S]*?<\/pivotField>/g)]
    .filter(match => new RegExp(`\\baxis="${axis}"`).test(match[1] ?? ''))
  if (fields.length !== 1) throw new Error(`${axis} pivot field scope matched ${fields.length} XML targets`)
  return fields[0]![0]
}

/** Add h="1" to one cache-indexed pivot item in one named axis field. */
export function hidePivotFieldItems(xml: string, axis: 'axisRow' | 'axisCol' | 'axisPage', itemIds: readonly number[]): string {
  const fields = [...xml.matchAll(/<pivotField\b([^>]*)>[\s\S]*?<\/pivotField>/g)]
    .filter(match => new RegExp(`\\baxis="${axis}"`).test(match[1] ?? ''))
  if (fields.length !== 1) throw new Error(`${axis} pivot field matched ${fields.length} XML targets`)
  const field = fields[0]!
  let body = field[0]
  for (const itemId of itemIds) {
    if (!Number.isInteger(itemId) || itemId < 0) throw new Error(`invalid pivot item id ${itemId}`)
    const itemPattern = new RegExp(`<item\\b(?=[^>]*\\bx="${itemId}")[^>]*/>`)
    const matches = [...body.matchAll(new RegExp(itemPattern.source, 'g'))]
    if (matches.length !== 1) throw new Error(`item x=${itemId} matched ${matches.length} XML targets`)
    const item = matches[0]![0]
    if (/\bh="1"/.test(item)) continue
    const replacement = item.replace(/<item\b/, '<item h="1"')
    body = replaceXmlExactlyOnce(body, itemPattern, replacement, `hide pivot item ${itemId}`)
  }
  const start = field.index
  if (start === undefined) throw new Error(`${axis} pivot field has no source range`)
  return xml.slice(0, start) + body + xml.slice(start + field[0].length)
}

export function selectSlicerItems(xml: string, selectedItemIds: readonly number[]): string {
  const blocks = [...xml.matchAll(/<items\b([^>]*)>([\s\S]*?)<\/items>/g)]
  if (blocks.length !== 1) throw new Error(`slicer items block matched ${blocks.length} XML targets`)
  const count = Number(/\bcount="(\d+)"/.exec(blocks[0]![1] ?? '')?.[1])
  const ids = new Set(selectedItemIds)
  if (!Number.isSafeInteger(count) || [...ids].some(x => !Number.isInteger(x) || x < 0 || x >= count)) {
    throw new Error(`selected slicer item id is outside item count ${count}`)
  }
  const noData = new Set<number>()
  for (const match of (blocks[0]![2] ?? '').matchAll(/<i\b([^>]*)\/>/g)) {
    const x = Number(/\bx="(\d+)"/.exec(match[1] ?? '')?.[1])
    if (/\bnd="1"/.test(match[1] ?? '')) noData.add(x)
  }
  return setSlicerItemStates(xml, Array.from({ length: count }, (_, x) => ({ x, selected: ids.has(x), noData: noData.has(x) })))
}

export function clearSlicerItems(xml: string): string {
  const blocks = [...xml.matchAll(/<items\b([^>]*)>([\s\S]*?)<\/items>/g)]
  if (blocks.length !== 1) throw new Error(`slicer items block matched ${blocks.length} XML targets`)
  const count = Number(/\bcount="(\d+)"/.exec(blocks[0]![1] ?? '')?.[1])
  if (!Number.isSafeInteger(count) || count < 0) throw new Error(`invalid slicer item count ${count}`)
  return selectSlicerItems(xml, Array.from({ length: count }, (_, x) => x))
}

/** Replace one complete <items> vector while retaining per-item extension attributes. */
export function setSlicerItemStates(xml: string, states: readonly SlicerItemState[]): string {
  const blocks = [...xml.matchAll(/<items\b([^>]*)>([\s\S]*?)<\/items>/g)]
  if (blocks.length !== 1) throw new Error(`slicer items block matched ${blocks.length} XML targets`)
  const block = blocks[0]!
  const count = Number(/\bcount="(\d+)"/.exec(block[1] ?? '')?.[1])
  if (!Number.isSafeInteger(count) || count !== states.length) {
    throw new Error(`item count ${count} does not match state count ${states.length}`)
  }
  const byId = new Map(states.map(state => [state.x, state]))
  if (byId.size !== count || states.some(state => !Number.isInteger(state.x) || state.x < 0 || state.x >= count)) {
    throw new Error(`slicer state must contain each item id 0..${count - 1} exactly once`)
  }
  const itemMatches = [...(block[2] ?? '').matchAll(/<i\b([^>]*)\/>/g)]
  if (itemMatches.length !== count) throw new Error(`slicer item vector has ${itemMatches.length} entries; expected ${count}`)
  const seen = new Set<number>()
  const items = itemMatches.map(match => {
    const attrs = match[1] ?? ''
    const x = Number(/\bx="(\d+)"/.exec(attrs)?.[1])
    if (!Number.isInteger(x) || seen.has(x)) throw new Error(`slicer item id ${x} is missing or duplicated`)
    seen.add(x)
    const state = byId.get(x)
    if (!state) throw new Error(`slicer item id ${x} has no desired state`)
    const extra = attrs.replace(/\s+(?:x|s|nd)="[^"]*"/g, '')
    return `<i x="${x}"${state.selected ? ' s="1"' : ''}${state.noData ? ' nd="1"' : ''}${extra}/>`
  }).join('')
  if (seen.size !== count) throw new Error(`slicer item vector covers ${seen.size} ids; expected ${count}`)
  const start = block.index!
  const original = block[0]!
  const replacement = `<items count="${count}">${items}</items>`
  return xml.slice(0, start) + replacement + xml.slice(start + original.length)
}
