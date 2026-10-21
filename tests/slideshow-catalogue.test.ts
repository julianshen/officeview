import { readFileSync } from 'node:fs'
import { describe, it, expect } from 'vitest'

async function modules() {
  const loaded = { ...await import('../src/slideshow/catalogue'), ...await import('../src/slideshow/inventory') }
  expect(loaded).toHaveProperty('CATALOGUE')
  return loaded
}

describe('slideshow catalogue', () => {
  it('every inventory transition, behaviour and preset maps to exactly one catalogue row', async () => {
    const { INVENTORY, CATALOGUE, rowForPreset, rowForTransition } = await modules()
    expect(INVENTORY.transitions).toHaveLength(56) // 21 ECMA + 19 p14 + 12 p15 + morph default/3 options
    expect(INVENTORY.behaviours).toHaveLength(10)
    expect(INVENTORY.presets.filter(p => p.subtype === undefined)).toHaveLength(199)
    expect(INVENTORY.presets.filter(p => p.subtype !== undefined)).toHaveLength(40) // two fly families x 20 source directions
    expect(new Set(CATALOGUE.map(r => r.id)).size).toBe(CATALOGUE.length)
    for (const t of INVENTORY.transitions) {
      const row = rowForTransition(t.ns, t.element, t.preset)
      expect(row, JSON.stringify(t)).toBeDefined()
      expect(CATALOGUE.filter(r => r.id === row!.id)).toHaveLength(1)
    }
    for (const b of INVENTORY.behaviours) expect(CATALOGUE.filter(r => r.kind === 'behaviour' && r.ns === 'p' && r.element === b)).toHaveLength(1)
    for (const p of INVENTORY.presets) {
      const rows = CATALOGUE.filter(r => r.preset?.cls === p.cls && r.preset.id === p.id && r.preset.subtype === p.subtype)
      expect(rows, JSON.stringify(p)).toHaveLength(1)
      expect(rowForPreset(p.cls, p.id, p.subtype)).toBe(rows[0])
    }
  })
  it('every row has an implementation, a status, a tolerance and a stage; verified is false unless a compare record exists', async () => {
    const { CATALOGUE } = await modules()
    for (const row of CATALOGUE) {
      expect(['file', 'preset-table', 'mask', 'transform', 'shader', 're-raster', 'stand-in']).toContain(row.implementation)
      expect(['exact', 'approximate', 'pending-stage-2', 'pending-stage-3', 'never']).toContain(row.status)
      expect([1, 2, 3]).toContain(row.stage)
      expect(row.verified).toBe(false) // no compare records exist in S1-3
      expect(row.uiName.length).toBeGreaterThan(0)
      expect(row.tolerance).toEqual(row.status === 'exact' ? { meanAbs: 0.02, rowMax: 0.10 } : { meanAbs: 0.08, rowMax: 0.30 })
    }
  })
  it('every 3-D row names a lookalike that is itself a 2-D exact row', async () => {
    const { CATALOGUE, rowForTransition } = await modules()
    for (const element of ['vortex', 'prism', 'honeycomb', 'glitter', 'ripple', 'switch', 'flip', 'doors', 'window', 'ferris', 'gallery', 'conveyor', 'shred']) {
      expect(rowForTransition('p14', element)?.status).toBe('pending-stage-3')
    }
    for (const row of CATALOGUE.filter(r => r.kind === 'transition' && r.stage === 3)) {
      const fallback = CATALOGUE.find(r => r.id === row.lookalike)
      expect(fallback, row.id).toMatchObject({ kind: 'transition', status: 'exact', stage: 1 })
    }
    expect(rowForTransition('p14', 'prism')?.lookalike).toBe('transition:p:push')
    expect(rowForTransition('p14', 'vortex')?.lookalike).toBe('transition:p:dissolve')
  })
  it('no row is pending a stage that has already landed', async () => {
    const { CATALOGUE, LANDED_STAGE } = await modules()
    expect(LANDED_STAGE).toBe(1)
    for (const r of CATALOGUE) if (r.status.startsWith('pending-stage-')) expect(Number(r.status.at(-1))).toBeGreaterThan(LANDED_STAGE)
  })
  it('the validation doc catalogue section equals catalogueMarkdown()', async () => {
    const { catalogueMarkdown } = await modules()
    const doc = readFileSync('docs/validation/slideshow.md', 'utf8')
    expect(doc.split('<!-- slideshow-catalogue:start -->\n')[1]?.split('\n<!-- slideshow-catalogue:end -->')[0]).toBe(catalogueMarkdown())
  })
  it('the inventory module imports nothing from src/slideshow/read', async () => {
    await modules()
    const source = readFileSync('src/slideshow/inventory.ts', 'utf8')
    expect(source).not.toMatch(/\bimport\s/)
    expect(source).toContain('2578da962608a2123d637b80499123c258829582')
    expect(source).toContain('MPL-2.0')
    const catalogue = readFileSync('src/slideshow/catalogue.ts', 'utf8')
    expect(catalogue).not.toMatch(/from ['"].*inventory/)
  })
  it('unknown identities do not silently become known effects; namespaces and subtypes stay distinct', async () => {
    const { rowForPreset, rowForTransition } = await modules()
    expect(rowForTransition('p', 'wipe')?.status).toBe('exact')
    expect(rowForTransition('http://schemas.openxmlformats.org/presentationml/2006/main', 'wipe')?.status).toBe('exact')
    expect(rowForTransition('unknown', 'wipe')).toBeUndefined()
    expect(rowForTransition('p14', 'wipe')).toBeUndefined()
    expect(rowForTransition('p', 'wipe', 'bogus')).toBeUndefined()
    expect(rowForTransition('p15', 'prstTrans', 'bogus')).toBeUndefined()
    expect(rowForTransition('p15', 'prstTrans:fallOver')).toBeUndefined()
    expect(rowForTransition('p159', 'morph', 'byWord')?.status).toBe('pending-stage-3')
    expect(rowForPreset('entr', 2, 1)?.uiName).toContain('from-top')
    expect(rowForPreset('entr', 2, 999999)).toBeUndefined()
    expect(rowForPreset('unknown', 2)).toBeUndefined()
    expect(rowForPreset('entr:2', 1)).toBeUndefined()
    expect(rowForPreset('entr', 999999)).toBeUndefined()
    expect(rowForPreset('entr', 2, NaN)).toBeUndefined()
  })
  it('stage-one file behaviours differ from preset reconstruction and media stand-ins', async () => {
    const { CATALOGUE, rowForPreset, catalogueMarkdown } = await modules()
    expect(rowForPreset('entr', 1)).toMatchObject({ implementation: 'preset-table', status: 'approximate', stage: 1 })
    for (const name of ['set','anim','animEffect','animMotion','animRot','animScale','animClr']) expect(CATALOGUE.find(r => r.id === `behaviour:p:${name}`)).toMatchObject({ status: 'exact', stage: 1 })
    for (const name of ['cmd','audio','video']) expect(CATALOGUE.find(r => r.id === `behaviour:p:${name}`)).toMatchObject({ status: 'pending-stage-2', stage: 2, implementation: 'stand-in' })
    expect(catalogueMarkdown()).toContain('trigger beats hyperlink')
    expect(catalogueMarkdown()).toContain('kiosk link inertness')
    expect(catalogueMarkdown()).toContain('pending D4')
  })
})
