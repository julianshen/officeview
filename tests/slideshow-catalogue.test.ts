import { readFileSync } from 'node:fs'
import { describe, it, expect } from 'vitest'

async function modules() {
  const loaded = { ...await import('../src/slideshow/catalogue'), ...await import('../src/slideshow/inventory') }
  return loaded
}

describe('slideshow catalogue', () => {
  it('every inventory transition, behaviour and preset maps to exactly one catalogue row', async () => {
    const { INVENTORY, CATALOGUE, rowForPreset, rowForTransition } = await modules()
    expect(INVENTORY.transitions).toHaveLength(56) // 21 ECMA + 19 p14 + 12 p15 + morph default/3 options
    expect(INVENTORY.behaviours).toHaveLength(10)
    expect(INVENTORY.presets.filter(p => p.subtype === undefined)).toHaveLength(199)
    expect(INVENTORY.presets.filter(p => p.subtype !== undefined)).toHaveLength(2040) // 102 entrance/exit families (excluding wheel) x 20 directions
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
      expect(rowForTransition('p14', element)).toMatchObject({ status: 'approximate', implementation: 'shader', stage: 3 })
    }
    for (const row of CATALOGUE.filter(r => r.kind === 'transition' && r.stage === 3)) {
      const fallback = CATALOGUE.find(r => r.id === row.lookalike)
      expect(fallback, row.id).toMatchObject({ kind: 'transition', status: 'exact', stage: 1 })
    }
    expect(rowForTransition('p14', 'prism')?.lookalike).toBe('transition:p:push')
    expect(rowForTransition('p14', 'vortex')?.lookalike).toBe('transition:p:dissolve')
  })
  it('no pending-stage rows remain', async () => {
    const { CATALOGUE, LANDED_STAGE } = await modules()
    // Stage 3 close (S3-7): the catalogue is complete. Every row is exact, approximate (with its
    // tolerance) or never; none is pending a stage.
    expect(LANDED_STAGE).toBe(3)
    expect(CATALOGUE.filter(r => r.status.startsWith('pending-stage-')).map(r => r.id)).toEqual([])
    for (const r of CATALOGUE) expect(['exact', 'approximate', 'never'], r.id).toContain(r.status)
  })
  it('every 3-D and Morph row is an approximate shader row that keeps its look-alike, unverified until D4 (S3-7)', async () => {
    const { CATALOGUE, TOLERANCE_APPROX } = await modules()
    const stage3 = CATALOGUE.filter(r => r.stage === 3)
    expect(stage3).toHaveLength(30) // 14 p14 + 12 p15 prstTrans + 4 Morph
    for (const r of stage3) {
      expect(r, r.id).toMatchObject({ kind: 'transition', status: 'approximate', implementation: 'shader', verified: false })
      expect(r.tolerance, r.id).toEqual(TOLERANCE_APPROX)
      expect(r.lookalike, r.id).toMatch(/^transition:p:/)
    }
    expect(stage3.filter(r => r.ns === 'p159').map(r => r.id)).toEqual([
      'transition:p159:morph', 'transition:p159:morph:byObject', 'transition:p159:morph:byWord', 'transition:p159:morph:byChar',
    ])
  })
  it('the validation doc catalogue section equals catalogueMarkdown()', async () => {
    const { catalogueMarkdown } = await modules()
    const doc = readFileSync('docs/validation/slideshow.md', 'utf8')
    expect(doc.split('<!-- slideshow-catalogue:start -->\n')[1]?.split('\n<!-- slideshow-catalogue:end -->')[0]).toBe(catalogueMarkdown())
  })
  it('the inventory module imports nothing from src/slideshow/read', async () => {
    await modules()
    const source = readFileSync('src/slideshow/inventory.ts', 'utf8')
    expect(source).not.toMatch(/^\s*import\b|import\(/m)
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
    expect(rowForTransition('p159', 'morph', 'byWord')?.status).toBe('approximate')
    expect(rowForPreset('entr', 2, 1)?.uiName).toContain('from-top')
    expect(rowForPreset('entr', 2, 999999)).toBe(rowForPreset('entr', 2))
    expect(rowForPreset('unknown', 2)).toBeUndefined()
    expect(rowForPreset('entr:2', 1)).toBeUndefined()
    expect(rowForPreset('entr', 999999)).toBeUndefined()
    expect(rowForPreset('entr', 2, NaN)).toBeUndefined()
  })
  it('stage-one file behaviours differ from preset reconstruction; stage-two media plays from the file', async () => {
    const { CATALOGUE, rowForPreset, catalogueMarkdown } = await modules()
    expect(rowForPreset('entr', 1)).toMatchObject({ implementation: 'preset-table', status: 'approximate', stage: 1 })
    for (const name of ['set','anim','animEffect','animMotion','animRot','animScale','animClr']) expect(CATALOGUE.find(r => r.id === `behaviour:p:${name}`)).toMatchObject({ status: 'exact', stage: 1 })
    // S2-7 plays shape media; S2-9 flipped the rows. Video is an overlay that shape effects do not move; cmd covers the call verbs.
    expect(CATALOGUE.find(r => r.id === 'behaviour:p:audio')).toMatchObject({ status: 'exact', stage: 2, implementation: 'file' })
    for (const name of ['cmd','video']) expect(CATALOGUE.find(r => r.id === `behaviour:p:${name}`)).toMatchObject({ status: 'approximate', stage: 2, implementation: 'file' })
    expect(CATALOGUE.find(r => r.id === 'sound:transition')).toMatchObject({ status: 'exact', stage: 2, implementation: 'file' })
    // S2-10: an effect sound (p:audio on a p:sndTgt) plays with its effect, like a transition sound.
    expect(CATALOGUE.find(r => r.id === 'sound:effect')).toMatchObject({ status: 'exact', stage: 2, implementation: 'file' })
    expect(CATALOGUE.find(r => r.id === 'sound:effect')?.lookalike).toBeUndefined()
    // S2-8: text builds by word and letter (p:iterate wd / lt) are clipped per-item layers.
    for (const id of ['build:word','build:letter']) expect(CATALOGUE.find(r => r.id === id)).toMatchObject({ status: 'approximate', stage: 2, implementation: 'mask' })
    expect(catalogueMarkdown()).toContain('trigger beats hyperlink')
    expect(catalogueMarkdown()).toContain('kiosk link inertness')
    expect(catalogueMarkdown()).toContain('pending D4')
  })
})

it('known preset families retain direction parameters, including zero and special mappings', async () => {
  const { rowForPreset, presetDirection } = await modules()
  for (const [id, subtype, direction] of [[22,8,'from-left'],[3,10,'vertical'],[16,21,'vertical-in'],[1,0,'0'],[5,5,'downward'],[5,10,'across'],[17,10,'across'],[18,3,'right-to-top'],[18,6,'right-to-bottom'],[18,9,'left-to-top'],[18,12,'left-to-bottom'],[21,8,'8']] as const) {
    expect(rowForPreset('entr', id, subtype)).toBeDefined()
    expect(presetDirection('entr', id, subtype)).toBe(direction)
    expect(presetDirection('exit', id, subtype)).toBe(direction)
  }
  expect(rowForPreset('entr',1,0)).toBe(rowForPreset('entr',1))
  expect(presetDirection('entr',22,999999)).toBe('999999')
  expect(rowForPreset('entr',999999,8)).toBeUndefined()
})
it('inventory is duplicate-free and covers the catalogue in reverse', async () => {
  const { INVENTORY, CATALOGUE } = await modules()
  expect(new Set(INVENTORY.presets.map(p => JSON.stringify([p.cls,p.id,p.subtype]))).size).toBe(INVENTORY.presets.length)
  expect(new Set(INVENTORY.transitions.map(t => JSON.stringify([t.ns,t.element,t.preset]))).size).toBe(INVENTORY.transitions.length)
  expect(new Set(INVENTORY.behaviours).size).toBe(INVENTORY.behaviours.length)
  for (const r of CATALOGUE) {
    if (r.preset) expect(INVENTORY.presets.some(p => p.cls === r.preset!.cls && p.id === r.preset!.id && p.subtype === r.preset!.subtype)).toBe(true)
    if (r.kind === 'transition') expect(INVENTORY.transitions.some(t => `transition:${t.ns}:${t.element}${t.preset === undefined ? '' : ':' + t.preset}` === r.id)).toBe(true)
    if (r.status.startsWith('pending-stage-')) expect(r.stage).toBe(Number(r.status.at(-1)))
  }
})
it('shared policy rows are frozen and an unavailable codec is an approximate stand-in', async () => {
  const { CATALOGUE } = await modules()
  expect(Object.isFrozen(CATALOGUE)).toBe(true)
  for (const r of CATALOGUE) {
    expect(Object.isFrozen(r)).toBe(true)
    expect(Object.isFrozen(r.tolerance)).toBe(true)
    if (r.preset) expect(Object.isFrozen(r.preset)).toBe(true)
    if (r.lookalike === undefined) expect(Object.hasOwnProperty.call(r,'lookalike')).toBe(false)
  }
  expect(CATALOGUE.find(r => r.id === 'media:codec')).toMatchObject({ status: 'approximate', stage: 2, implementation: 'stand-in' })
})
it('the validation doc lists every Stage 2 diagnostic code with its catalogue class and panel group (S2-9)', async () => {
  const { DIAGNOSTIC_CODES } = await import('../src/slideshow/diagnostics')
  const { diagnosticKind } = await import('../src/slideshow/presenter/diagnostics')
  const doc = readFileSync('docs/validation/slideshow.md', 'utf8')
  const section = doc.split('### Stage 2 diagnostic codes\n')[1]?.split('\n## ')[0] ?? ''
  const rows = section.split('\n').filter(l => /^\| `/.test(l)).map(l => l.split('|').slice(1, -1).map(c => c.trim()))
  const PANEL = { refused: 'Refused', approximate: 'Approximate', pending: 'Pending', other: 'Other notes' } as const
  const listed = new Set<string>()
  for (const [code, cls, panel] of rows) {
    const name = code.replace(/`/g, '')
    listed.add(name)
    expect(DIAGNOSTIC_CODES[name], name).toBe(cls)
    expect(PANEL[diagnosticKind({ code: name, message: '' })], name).toBe(panel)
  }
  // The Stage 2 block of the catalogue (media playback onwards). S2-10 retired media-stand-in: effect sounds play.
  const keys = Object.keys(DIAGNOSTIC_CODES)
  expect(keys).not.toContain('media-stand-in')
  const stage2 = keys.slice(keys.indexOf('media-autoplay'), keys.indexOf('transition-shader-fallback'))
  expect(stage2.filter(c => !listed.has(c))).toEqual([])
  expect(rows.length).toBe(listed.size)
})
it('the validation doc lists every Stage 3 diagnostic code with its catalogue class and panel group (S3-7)', async () => {
  const { DIAGNOSTIC_CODES } = await import('../src/slideshow/diagnostics')
  const { diagnosticKind } = await import('../src/slideshow/presenter/diagnostics')
  const doc = readFileSync('docs/validation/slideshow.md', 'utf8')
  const section = doc.split('### Stage 3 diagnostic codes\n')[1]?.split('\n## ')[0] ?? ''
  const rows = section.split('\n').filter(l => /^\| `/.test(l)).map(l => l.split('|').slice(1, -1).map(c => c.trim()))
  const PANEL = { refused: 'Refused', approximate: 'Approximate', pending: 'Pending', other: 'Other notes' } as const
  const listed = rows.map(([code, cls, panel]) => {
    const name = code.replace(/`/g, '')
    expect(DIAGNOSTIC_CODES[name], name).toBe(cls)
    expect(PANEL[diagnosticKind({ code: name, message: '' })], name).toBe(panel)
    return name
  })
  // The Stage 3 block of the catalogue: transition-shader-fallback onwards, each listed once.
  const keys = Object.keys(DIAGNOSTIC_CODES)
  expect(listed).toEqual(keys.slice(keys.indexOf('transition-shader-fallback')))
  expect(listed).toEqual(['transition-shader-fallback', 'morph-refused', 'morph-units-coarsened', 'morph-unit-bound'])
  // morph-fallback is an effect of transition-approximate, not a code.
  expect(keys).not.toContain('morph-fallback')
})
