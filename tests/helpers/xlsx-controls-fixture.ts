import type { FilterRange, SortSpec } from '../../src/xlsx/view/contracts'
import { createControlsClient, type ControlsClientOptions, type ControlSnapshot, type StagedControl } from '../../src/components/xlsx/controlsClient'
import { ValueIndexCache, SOURCE_STRING, SOURCE_NUMBER, SOURCE_ERROR, SOURCE_BLANK, type ValueColumnSource } from '../../src/xlsx/view/valueIndex'
import { MemoryValueRunStore } from '../../src/xlsx/view/valueRuns'
import { FilterEngine, type ColumnClassification } from '../../src/xlsx/view/filter'
import { compileValuesMembership } from '../../src/xlsx/view/valueIndex'
import { createFixture } from './xlsx-viewer-fixtures'
import { StructuralIndex } from '../../src/xlsx/view/structure'

export function controlsFixture(count = 11020, classification: ColumnClassification = { kind: 'text', hasNumeric: true }, table = false) {
  const handle = createFixture('F2', { rowCount: count + 1 })
  const store = handle.createStore({ read: { isPopulated: (row, col) => row <= count && col < 10, cellAt: (row, col) => ({ value: row === 0 ? `Header ${col}` : row + col, hasCachedValue: true, valueIsError: false, styleIndex: 0 }) } })
  const range: FilterRange = { r0: 0, r1: table ? count - 1 : count, c0: 0, c1: 9, headerRows: table ? 0 : 1, totalsRows: 0 }
  const structure = new StructuralIndex({ stamp: store.stamp, completeness: 'complete', tables: table ? [{ id: 'table', range, headerRows: 0, totalsRows: 0 }] : [], merges: [], spills: [], drawings: [], missing: [] })
  let snapshot: ControlSnapshot = { owner: { ...store.stamp, fontGen: 1, viewGen: 1 }, filter: null, sort: null, records: { visible: count, total: count, completeness: 'complete' } }
  const cache = new ValueIndexCache({ limitBytes: 32 * 1024 * 1024 }), runs = new MemoryValueRunStore({ limitBytes: 16 * 1024 * 1024, account: cache.reserve() })
  const context = { locale: 'en-US', dateSystem: '1900' as const, timeZone: 'UTC', nowMs: Date.UTC(2026, 9, 11) }
  let publications = 0
  const valueAt = (ordinal: number): string | number => classification.kind === 'date' ? 46306 + ordinal / 2 : classification.kind === 'number' ? ordinal - 4 : ordinal === count - 2 ? 'beyond-A' : ordinal === count - 1 ? 'beyond-B' : ordinal === 0 ? '<img src=x onerror=evil()>' : `item-${ordinal}`
  const source: ValueColumnSource = { get stamp() { return { docId: snapshot.owner.docId, sheetId: snapshot.owner.sheetId, valueGen: snapshot.owner.valueGen } }, pageCount: Math.ceil(count / 512), async readPage(page, signal) {
    signal.throwIfAborted()
    const n = Math.min(512, count - page * 512), tags = new Uint8Array(n), numbers = new Float64Array(n), texts: (string | undefined)[] = []
    for (let i = 0; i < n; i++) { const ordinal = page * 512 + i, value = valueAt(ordinal); tags[i] = typeof value === 'number' ? SOURCE_NUMBER : SOURCE_STRING; if (typeof value === 'number') numbers[i] = value; else texts[i] = value }
    if (classification.kind !== 'date' && page === 0 && n > 2) { tags[1] = SOURCE_ERROR; texts[1] = '#N/A'; tags[2] = SOURCE_BLANK }
    return { kind: 'values', pageId: `p${page}`, count: n, tags, numbers, texts }
  } }
  const identity = (r: FilterRange, col: number) => ({ ...source.stamp, range: r, col, classification, classificationGen: 1 })
  const options: ControlsClientOptions = { snapshot: () => snapshot, context: () => context, cache, runs,
    rangeRequest: col => ({ store, structure, selection: { active: { row: 1, col: col ?? 0 }, anchor: { row: 1, col: col ?? 0 }, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: count }] }, cols: [{ c0: 0, c1: 9 }] }] } }),
    domain: (_owner, r, col) => ({ source, identity: identity(r, col), label: (_kind, value) => String(value) }),
    async stageFilter(filter, _owner, signal): Promise<StagedControl> {
      let visible = count
      if (filter?.columns.size) {
        const rule = filter.columns.get(0), match = rule?.kind === 'values' ? (await compileValuesMembership(rule, identity(filter.range, 0), runs, { signal })).matches : rule ? new FilterEngine(context).compileRule(rule, classification) : () => true
        visible = 0
        for (let i = 0; i < count; i++) { const value = valueAt(i); if (match(classification.kind !== 'date' && i === 1 ? { kind: 'error', value: '#N/A' } : classification.kind !== 'date' && i === 2 ? { kind: 'blank' } : typeof value === 'number' ? { kind: 'number', value, displayText: String(value), ...(classification.kind === 'date' ? { numFmtId: 14 } : {}) } : { kind: 'string', value })) visible++ }
      }
      return { kind: 'ready', publish() { publications++; snapshot = { ...snapshot, owner: { ...snapshot.owner, viewGen: snapshot.owner.viewGen + 1 }, filter, records: { visible, total: count, completeness: 'complete' } }; return snapshot } }
    },
    async stageSort(sort: SortSpec | null): Promise<StagedControl> { return { kind: 'ready', publish() { publications++; snapshot = { ...snapshot, owner: { ...snapshot.owner, viewGen: snapshot.owner.viewGen + 1 }, sort }; return snapshot } } },
  }
  return { options, client: createControlsClient(options), cache, runs, range, source, context, store, get snapshot() { return snapshot }, set snapshot(value: ControlSnapshot) { snapshot = value }, get publications() { return publications }, valueAt }
}
