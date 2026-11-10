import { posix } from 'node:path'
import { all, attr, local, remove, type Context, type Node } from './tree'

const hidden = (node: Node) => /^(1|true)$/i.test(attr(node, 'hidden') ?? '')
const colNumber = (letters: string) => [...letters.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0)
const target = (base: string, value: string) => value.startsWith('/') ? value.slice(1) : posix.normalize(posix.join(posix.dirname(base), value))
function linked(ctx: Context, part: string, id: string | undefined): string | undefined {
  const rels = ctx.parts.get(posix.join(posix.dirname(part), '_rels', posix.basename(part) + '.rels'))
  const rel = rels && all(rels, 'Relationship').find(r => attr(r, 'Id') === id)
  const value = rel && attr(rel, 'Target')
  return value && attr(rel!, 'TargetMode') !== 'External' ? target(part, value) : undefined
}
function diagnostic(ctx: Context, part: string, node: Node, key?: string) {
  ctx.diagnostics.push({ code: 'hidden-stripped', part, element: node.name, ...(key ? { attr: key, value: attr(node, key) } : {}) })
}
interface Sheet { path?: string; hidden: boolean; rows: number[]; cols: Array<[number, number]> }

export function projectXlsx(ctx: Context): void {
  const wb = ctx.parts.get('xl/workbook.xml')
  const sheets = new Map<string, Sheet>()
  const deleted = new Set<number>()
  for (const [index, sheet] of (wb ? all(wb, 'sheet') : []).entries()) {
    const path = linked(ctx, 'xl/workbook.xml', sheet.attrs['r:id'] ?? attr(sheet, 'id'))
    const state = attr(sheet, 'state')
    const isHidden = state !== undefined && state !== 'visible'
    sheets.set(attr(sheet, 'name') ?? '', { path, hidden: isHidden, rows: [], cols: [] })
    if (isHidden) {
      deleted.add(index)
      remove(wb!, sheet)
      if (path) ctx.dropPart(path)
      if (state === 'hidden' || state === 'veryHidden') diagnostic(ctx, 'xl/workbook.xml', sheet, 'state')
      else ctx.diagnostics.push({ code: 'hidden-unrecognised', part: 'xl/workbook.xml', element: sheet.name, attr: 'state', value: state })
    }
  }
  if (wb) for (const name of all(wb, 'definedName')) {
    const id = attr(name, 'localSheetId')
    if (id === undefined) continue
    const index = Number(id)
    if (deleted.has(index)) remove(wb, name)
    else name.attrs.localSheetId = String(index - [...deleted].filter(n => n < index).length)
  }

  for (const [part, root] of [...ctx.parts]) {
    if (/^xl\/(?:comments[^/]*\.xml|threadedComments\/|pivotCache\/pivotCacheRecords[^/]*\.xml)/i.test(part)) {
      diagnostic(ctx, part, root); ctx.dropPart(part); continue
    }
    if (/^xl\/worksheets\/[^/]+\.xml$/.test(part)) {
      const sheet = [...sheets.values()].find(s => s.path === part)
      const ranges: Array<[number, number]> = []
      for (const col of all(root, 'col')) if (hidden(col)) {
        const range: [number, number] = [Number(attr(col, 'min')), Number(attr(col, 'max'))]
        ranges.push(range); sheet?.cols.push(range)
        ctx.handled.add(col); diagnostic(ctx, part, col, 'hidden')
      }
      let rowIndex = 0
      for (const row of all(root, 'row')) {
        rowIndex = Number(attr(row, 'r') ?? rowIndex + 1)
        const rowHidden = hidden(row)
        if (attr(row, 'hidden') !== undefined && !/^(0|false)$/i.test(attr(row, 'hidden')!)) sheet?.rows.push(rowIndex)
        if (rowHidden) { ctx.handled.add(row); diagnostic(ctx, part, row, 'hidden') }
        let column = 0
        row.children = row.children.filter(cell => {
          if (local(cell.name) !== 'c') return true
          const ref = attr(cell, 'r')?.match(/^([A-Z]+)/i)
          column = ref ? colNumber(ref[1]!) : column + 1
          return !rowHidden && !ranges.some(([lo, hi]) => column >= lo && column <= hi)
        })
      }
    }
    if (/^xl\/drawings\/[^/]+\.xml$/.test(part)) {
      for (const anchor of [...all(root, 'twoCellAnchor'), ...all(root, 'oneCellAnchor'), ...all(root, 'absoluteAnchor')]) if (all(anchor, 'cNvPr').some(hidden)) {
        diagnostic(ctx, part, anchor); remove(root, anchor)
      }
    }
  }

  const unsafeCaches = new Map<string, boolean>()
  for (const [part, root] of ctx.parts) if (/^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/i.test(part)) {
    const sources = all(root, 'worksheetSource')
    const unsafe = sources.length !== 1 || sources.some(source => {
      const sheet = sheets.get(attr(source, 'sheet') ?? '')
      const range = attr(source, 'ref')?.match(/^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i)
      if (!sheet || sheet.hidden || !sheet.path || !ctx.parts.has(sheet.path) || !range) return true
      const x1 = colNumber(range[1]!), y1 = Number(range[2]), x2 = colNumber(range[3] ?? range[1]!), y2 = Number(range[4] ?? range[2])
      return x1 > x2 || y1 > y2 || sheet.rows.some(y => y >= y1 && y <= y2) || sheet.cols.some(([lo, hi]) => lo <= x2 && hi >= x1)
    })
    unsafeCaches.set(part, unsafe)
    if (unsafe) emptyItems(ctx, part, root)
  }
  // Controls use the x14 extension pivotCacheId, not workbook cacheId.
  const cacheIds = new Map<string, boolean>()
  for (const [part, unsafe] of unsafeCaches) {
    const root = ctx.parts.get(part)!
    for (const definition of all(root, 'pivotCacheDefinition')) {
      const id = attr(definition, 'pivotCacheId')
      if (id !== undefined) cacheIds.set(id, unsafe || cacheIds.get(id) === true)
    }
  }
  for (const [part, root] of ctx.parts) if (/^xl\/(?:slicerCaches|timelineCaches)\/.*\.xml$/i.test(part)) {
    const ids: string[] = []
    const visit = (n: Node) => { const id = attr(n, 'pivotCacheId'); if (id !== undefined) ids.push(id); n.children.forEach(visit) }
    visit(root)
    if (!ids.length || ids.some(id => cacheIds.get(id) !== false)) emptyItems(ctx, part, root)
  }
  purgeUnusedSharedStrings(ctx)
}

function emptyItems(ctx: Context, part: string, root: Node) {
  // These lists carry source values independently of worksheet cells. Keep field
  // metadata so existing rendered pivot cells retain their authored geometry.
  const lists = new Set(['sharedItems', 'groupItems', 'items', 'levels', 'members', 'selections', 'bounds', 'state', 'timelineState'])
  const visit = (node: Node) => {
    if (lists.has(local(node.name))) {
      node.children = []; delete node.text
      for (const key of Object.keys(node.attrs)) if (local(key) !== 'count' && !key.startsWith('xmlns')) delete node.attrs[key]
      node.attrs.count = '0'
    } else node.children.forEach(visit)
  }
  visit(root); diagnostic(ctx, part, root)
}

/** Keep indices stable: unused shared-string slots become empty strings. */
export function purgeUnusedSharedStrings(ctx: Context): void {
  const sst = ctx.parts.get('xl/sharedStrings.xml')
  if (!sst) return
  const used = new Set<number>()
  for (const [part, root] of ctx.parts) if (/^xl\/worksheets\/[^/]+\.xml$/.test(part)) {
    for (const cell of all(root, 'c')) if (attr(cell, 't') === 's') {
      const value = all(cell, 'v')[0]
      if (value) used.add(Number(value.text ?? value.children.map(n => n.text ?? '').join('')))
    }
  }
  all(sst, 'si').forEach((si, index) => {
    if (!used.has(index) && (si.children.length || si.text)) { si.children = []; delete si.text; diagnostic(ctx, 'xl/sharedStrings.xml', si) }
  })
}
