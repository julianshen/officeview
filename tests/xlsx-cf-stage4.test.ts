// @vitest-environment node
/**
 * Task 30: Stage 4 (icon sets) validation, resource audit and PR evidence.
 *
 * Verifies that:
 * - every enabled icon capability has verified evidence, and production publishes exactly the Stage 4 publication gate's
 *   result: verified, not withheld by a ruling, Excel's matched reading equal to production's, and (glyphs) a class that
 *   passed the OD-9(b) approximation check;
 * - every unresolved icon shape refuses atomically on the resident and the wire paths;
 * - corpus and oracle workbooks parse to identical CF rules on the resident and the streamed paths;
 * - the Stage 4 pressure scenarios hold their bounds with zero unowned live bytes.
 *
 * Covers only the `icon:*` rows: manifest-format mutations stay in tests/xlsx-cf-oracle.test.ts (Task 24).
 */
import { describe, expect, it } from 'vitest'
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join, resolve } from 'node:path'
import {
  auditStage4Manifest,
  CF_STAGE4_NEVER_PUBLISH,
  CF_STAGE4_PRODUCTION_READINGS,
  stage4Publication,
  type Stage4AuditReport,
} from '../scripts/cf-refs/audit'
import { CF_STAGE4_CAPABILITIES, CF_STAGE4_NONDISCRIMINATING, CF_STAGE4_UNSUPPORTED, type CfManifest } from '../scripts/cf-refs/cases'
import * as verifiedExports from '../src/xlsx/cf/verifiedCapabilities'
import { VERIFIED_CAPABILITIES, VERIFIED_STAGE1_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES, iconSelectionEvidence } from '../src/xlsx/cf/evidence'
import { CF_ICON_PAINT_READINGS } from '../src/xlsx/cf/iconPaint'
import { CF_ICON_GLYPH_CLASSES, CF_ICON_GLYPH_WITHHELD, iconGlyph } from '../src/xlsx/cf/iconGlyphs'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { prepareCfTextScope, prepareCfWireScope } from '../src/xlsx/cf/text'
import { createCfService } from '../src/xlsx/cf/service'
import { createCfMemory, CF_LIMITS } from '../src/xlsx/cf/limits'
import { residentSource } from '../src/xlsx/cf/render'
import { cfRulesVersion } from '../src/xlsx/view/pagedSource'
import type { FormatBudget } from '../src/xlsx/view/formatWorkspace'
import type { XlsxSheet } from '../src/xlsx/types'
import type { CfReadResult } from '../src/xlsx/view/contracts'
import type { CfIconSetName } from '../src/xlsx/cf/model'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const CORPUS = resolve(REPO_ROOT, 'corpus')
const manifest = (): CfManifest => JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
const UNAVAILABLE = 'Conditional formatting unavailable'

function trackingBudget(): { budget: FormatBudget; net: () => number } {
  let held = 0
  const budget: FormatBudget = {
    reserve: bytes => { held += bytes },
    release: bytes => { held -= bytes },
    retain: (_targets, bytes, commit) => { held += bytes; commit() },
    settle: () => {},
  }
  return { budget, net: () => held }
}

async function savedSheet(path: string): Promise<XlsxSheet> {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, path)))))
  return (doc.sheets.find(s => s.name === 'Sheet1') ?? doc.sheets[0]!) as unknown as XlsxSheet
}

const colOf = (letters: string): number => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
/** The single rect of an `A1:B2` sqref. */
function sqrefRect(sqref: string): { firstRow: number; lastRow: number; firstCol: number; lastCol: number } {
  const [a, b = a] = sqref.split(':').map(ref => /^([A-Z]+)(\d+)$/.exec(ref)!)
  return { firstRow: Number(a![2]) - 1, lastRow: Number(b![2]) - 1, firstCol: colOf(a![1]!), lastCol: colOf(b![1]!) }
}

function baseOf(sheet: XlsxSheet, row: number, col: number): unknown {
  for (const band of sheet.rows) {
    const cell = band.cells.find(candidate => candidate.row === row && candidate.col === col)
    if (cell) return cell.style
  }
  return undefined
}

const published = (): ReadonlySet<string> | undefined => Reflect.get(verifiedExports, 'VERIFIED_STAGE4_CAPABILITIES') as ReadonlySet<string> | undefined
const publishedSelections = (): Readonly<Record<string, string>> | undefined => Reflect.get(verifiedExports, 'VERIFIED_STAGE4_ICON_SELECTIONS') as Readonly<Record<string, string>> | undefined

describe('every enabled icon capability has verified evidence', () => {
  const audit = auditStage4Manifest(manifest(), { rootDir: REPO_ROOT })
  const by = (d: string) => audit.perCapability.filter(p => p.disposition === d).map(p => p.capabilityId).sort()
  const gate = stage4Publication(audit)

  it('the Stage 4 audit is complete, with every icon capability accounted for', () => {
    expect(audit.ok).toBe(true)
    expect(audit.stageComplete).toBe(true)
    expect(audit.blockedCapabilities).toEqual([])
    expect(audit.perCapability.map(p => p.capabilityId).sort()).toEqual(CF_STAGE4_CAPABILITIES.map(c => c.capabilityId).sort())
    for (const p of audit.perCapability) {
      expect(p.capabilityId.startsWith('icon:'), p.capabilityId).toBe(true)
      expect(['verified', 'unverified-nondiscriminating', 'refused-unsupported', 'supported-structural'], p.capabilityId).toContain(p.disposition)
    }
    expect(by('unverified-nondiscriminating')).toEqual(expect.arrayContaining(Object.keys(CF_STAGE4_NONDISCRIMINATING)))
    expect(by('refused-unsupported')).toEqual(expect.arrayContaining(Object.keys(CF_STAGE4_UNSUPPORTED)))
    expect(by('supported-structural')).toEqual([])
    expect(by('verified').length).toBeGreaterThan(100)
  })

  it('the publication gate: verified only, never a ruling-withheld id, matched reading = production reading, admitted glyph classes only', () => {
    const verified = new Set(by('verified'))
    // Partition: every Stage 4 id is either published or withheld with a reason.
    expect([...gate.published, ...Object.keys(gate.withheld)].sort()).toEqual(audit.perCapability.map(p => p.capabilityId).sort())
    for (const id of gate.published) expect(verified.has(id), id).toBe(true)
    // The ruling deny-list is verified (Excel answered) yet withheld.
    for (const id of Object.keys(CF_STAGE4_NEVER_PUBLISH)) {
      expect(verified.has(id), `${id} is verified`).toBe(true)
      expect(gate.published, id).not.toContain(id)
      expect(gate.withheld[id], id).toMatch(/^never-publish: /)
    }
    expect(Object.keys(CF_STAGE4_NEVER_PUBLISH).sort()).toEqual(['icon:geometry-short-row-clip', 'icon:inset-overflow', 'icon:scope-merged-cell'])
    // Ruling S4-RIGHTINSET-ZERO supersedes the older "do not publish": its matched reading is production's.
    expect(gate.published).toContain('icon:inset-right-aligned')
    // Non-discriminating by construction (S4-SNAP-SHARED): never published; the paint gate uses the Stage 3 bar one.
    expect(gate.published).not.toContain('icon:geometry-snapping')
    expect(VERIFIED_CAPABILITIES.has('bar:geometry-snapping')).toBe(true)
    // Every published id with a matched reading carries production's reading.
    for (const id of gate.published) {
      const matched = audit.matchedReadings[id]
      if (matched === undefined) continue
      const prod = CF_STAGE4_PRODUCTION_READINGS[id]
      expect(prod, `${id}: production names a reading`).toBeDefined()
      if (typeof prod === 'string') expect(prod, id).toBe(matched)
      else expect(prod, id).toContain(matched)
      if (id in CF_ICON_PAINT_READINGS) expect(CF_ICON_PAINT_READINGS[id], id).toBe(matched)
    }
    // Selections are exactly the published list-valued capabilities, each with its matched reading.
    expect(Object.keys(gate.selections).sort()).toEqual(gate.published.filter(id => Array.isArray(CF_STAGE4_PRODUCTION_READINGS[id])).sort())
    for (const [id, v] of Object.entries(gate.selections)) expect(v, id).toBe(audit.matchedReadings[id])
    // Glyphs: a pair publishes iff its class passed the approximation check (54 classes for 75 pairs).
    const glyphIds = audit.perCapability.filter(p => p.capabilityId.startsWith('icon:glyph-') && p.capabilityId !== 'icon:glyph-key').map(p => p.capabilityId)
    expect(glyphIds).toHaveLength(75)
    expect(new Set(Object.values(CF_ICON_GLYPH_CLASSES)).size).toBe(54)
    expect(CF_ICON_GLYPH_WITHHELD.size).toBe(15)
    let admittedPairs = 0
    for (const id of glyphIds) {
      const [, set, iconId] = /^icon:glyph-(.+)-(\d)$/.exec(id)!
      const cls = CF_ICON_GLYPH_CLASSES[`${set}:${iconId}`]!
      const admitted = !CF_ICON_GLYPH_WITHHELD.has(cls)
      if (admitted) admittedPairs++
      expect(gate.published.includes(id), `${id} (class ${cls})`).toBe(admitted && verified.has(id))
    }
    // The 39 admitted classes cover more than 39 pairs (shared artwork), fewer than all 75.
    expect(admittedPairs).toBeGreaterThan(39)
    expect(admittedPairs).toBeLessThan(75)
    // Non-vacuous: the gate both publishes and withholds verified ids for each reason.
    const reasons = Object.values(gate.withheld)
    for (const prefix of ['never-publish: ', 'glyph-approximation-failed: ', 'unverified-nondiscriminating', 'refused-unsupported']) {
      expect(reasons.some(r => r.startsWith(prefix)), prefix).toBe(true)
    }
  })

  it('the gate withholds a verified id whose matched reading production does not implement (mutations)', () => {
    const mutate = (id: string, reading: string): Stage4AuditReport => ({ ...audit, matchedReadings: { ...audit.matchedReadings, [id]: reading } })
    expect(stage4Publication(audit).published).toContain('icon:inset-left')
    expect(stage4Publication(mutate('icon:inset-left', 'none')).withheld['icon:inset-left']).toMatch(/^reading-mismatch: /)
    expect(stage4Publication(mutate('icon:thresholds-percentile-inc-exc', 'nearest')).withheld['icon:thresholds-percentile-inc-exc']).toMatch(/^reading-unimplemented: /)
    // A verified id with a matched reading that production names nowhere is withheld (fail closed).
    const extra: Stage4AuditReport = { ...audit, perCapability: [...audit.perCapability, { capabilityId: 'icon:invented', caseId: 'x', disposition: 'verified' }], matchedReadings: { ...audit.matchedReadings, 'icon:invented': 'r' } }
    expect(stage4Publication(extra).withheld['icon:invented']).toMatch(/^reading-unimplemented: /)
    // A verified glyph whose class failed the approximation check stays withheld, so publishing the key cannot un-withhold it.
    expect(stage4Publication(audit).withheld['icon:glyph-3Stars-0']).toMatch(/^glyph-approximation-failed: /)
  })

  it('production publishes exactly the gated Stage 4 set and its selections; the Stage 1-3 sets are untouched', () => {
    const s4 = published()
    expect(s4, 'VERIFIED_STAGE4_CAPABILITIES is published').toBeDefined()
    expect([...s4!].sort()).toEqual(gate.published)
    expect(publishedSelections()).toEqual(gate.selections)
    for (const id of gate.published) {
      expect(VERIFIED_CAPABILITIES.has(id), id).toBe(true)
      expect(cfRuntimeEvidence.allows(id), id).toBe(true)
    }
    for (const id of Object.keys(gate.withheld)) {
      expect(VERIFIED_CAPABILITIES.has(id), `${id} must not be production-verified`).toBe(false)
      expect(cfRuntimeEvidence.allows(id), id).toBe(false)
    }
    for (const [id, v] of Object.entries(gate.selections)) expect(iconSelectionEvidence(id), id).toBe(v)
    const s3 = Reflect.get(verifiedExports, 'VERIFIED_STAGE3_CAPABILITIES') as ReadonlySet<string>
    expect([...VERIFIED_CAPABILITIES].sort()).toEqual([...new Set([...VERIFIED_STAGE1_CAPABILITIES, ...VERIFIED_STAGE2_CAPABILITIES, ...s3, ...s4!])].sort())
    // A withheld glyph class never paints, even with its key capability verified.
    expect(iconGlyph('3Stars' as CfIconSetName, 0)).toBeUndefined()
    expect(iconGlyph('3Arrows' as CfIconSetName, 0)).toBeDefined()
  })
})

describe('every unresolved icon shape refuses atomically on both paths', () => {
  it('every parse-marked unresolved icon set withholds on the resident scope: unavailable, base style, no icon', async () => {
    const stage4 = manifest().cases.filter(c => c.stage === 4)
    const reasons = new Set<string>()
    let targetsSeen = 0, probed = 0
    for (const c of stage4) {
      // Excel cannot open or strips the unresolved shapes, so their case is read as authored (the workbook a viewer gets).
      const sheet = await savedSheet(c.status === 'unsupported' ? c.input : c.saved)
      const cf = sheet.conditionalFormatting
      if (!cf) continue
      const targets: { label: string; ranges: readonly { firstRow: number; lastRow: number; firstCol: number; lastCol: number }[] }[] = cf.unresolvedSheet
        ? c.markers.map(m => ({ label: `${m.capabilityId} (sheet: ${cf.unresolvedSheet!.reason})`, ranges: [sqrefRect(m.sqref)] }))
        : cf.rules.filter(rule => {
          const d = rule.definition as { kind?: string; type?: string }
          return d.kind === 'unresolved' && d.type === 'iconSet'
        }).map(rule => ({ label: `pri${rule.priority} ${(rule.definition as { reason: string }).reason}`, ranges: rule.ranges }))
      if (targets.length === 0) continue
      if (cf.unresolvedSheet) reasons.add(cf.unresolvedSheet.reason)
      else for (const rule of cf.rules) if ((rule.definition as { kind?: string }).kind === 'unresolved') reasons.add((rule.definition as { reason: string }).reason)
      const oracleMemory = createCfMemory()
      const oracle = createCfService(cf, oracleMemory, { evidence: cfRuntimeEvidence })
      const stamp = { owner: 'cf-text', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(cf), viewGen: 0, structureGen: 0, clockGen: 0 }
      const source = residentSource(sheet, stamp, oracleMemory, 'engine')
      for (const target of targets) {
        targetsSeen++
        const rects = target.ranges.map(r => ({ r0: r.firstRow, r1: Math.min(r.lastRow, r.firstRow + 4), c0: r.firstCol, c1: r.firstCol }))
        const lease = await oracle.prepareCf(
          { stamp, sourceRects: rects.map(r => ({ firstRow: r.r0, lastRow: r.r1, firstCol: r.c0, lastCol: r.c1 })), signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes },
          source,
          oracleMemory,
        )
        try {
          const { budget, net } = trackingBudget()
          const scope = await prepareCfTextScope(sheet, { rects, signal: new AbortController().signal, locale: 'en-US', budget })
          if (cf.unresolvedSheet) {
            expect(scope, `${c.id} sheet-level refusal builds no scope`).toBeUndefined()
            for (const r of rects) for (let row = r.r0; row <= r.r1; row++) {
              expect(lease.value.at(row, r.c0), `${c.id} ${target.label} ${row},${r.c0}`).toEqual({ kind: 'unresolved', reason: cf.unresolvedSheet.reason })
              probed++
            }
            expect(net()).toBe(0)
            continue
          }
          expect(scope, `${c.id} resident scope`).toBeDefined()
          for (const r of rects) for (let row = r.r0; row <= r.r1; row++) {
            const outcome = lease.value.at(row, r.c0)
            if (outcome.kind === 'unresolved' && outcome.reason === 'cf-outside-preparation') continue
            const where = `${c.id} ${target.label} ${row},${r.c0}`
            // The oracle never yields an icon for an unresolved rule's cell, and the scope paints none.
            expect(outcome.kind === 'ready' && outcome.value.visual?.iconSet !== undefined, where).toBe(false)
            expect(scope!.iconAt?.(row, r.c0), where).toBeUndefined()
            if (outcome.kind === 'unresolved') {
              expect(scope!.describe(row, r.c0)?.text, where).toBe(UNAVAILABLE)
              expect(scope!.effectiveStyle(row, r.c0, baseOf(sheet, row, r.c0) as never), where).toBe(baseOf(sheet, row, r.c0))
            }
            probed++
          }
          scope!.dispose()
          expect(net()).toBe(0)
        } finally {
          lease.release()
        }
      }
      oracle.dispose()
      expect(oracleMemory.currentBytes).toBe(0)
    }
    expect(targetsSeen).toBeGreaterThan(5)
    expect(probed).toBeGreaterThan(20)
    expect(reasons.size).toBeGreaterThan(0)
  })

  it('the unresolved-shapes workbook withholds every icon through the real paged worker route', async () => {
    const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
    const { createPagedWorkbookSource } = await import('../src/xlsx/view/pagedSource')
    const c = manifest().cases.find(x => x.id === 'stage4-icon-unresolved-shapes')!
    const sheet = await savedSheet(c.input)
    const cf = sheet.conditionalFormatting!
    const unresolved = cf.unresolvedSheet
      ? c.markers.map(m => ({ label: m.capabilityId, ranges: [sqrefRect(m.sqref)] }))
      : cf.rules.filter(r => (r.definition as { kind?: string }).kind === 'unresolved').map(r => ({ label: `pri${r.priority}`, ranges: r.ranges }))
    expect(unresolved.length).toBeGreaterThan(5)
    const session = await openImportSession(
      { kind: 'blob', blob: new Blob([new Uint8Array(readFileSync(join(REPO_ROOT, c.input))).buffer as ArrayBuffer]) },
      { docId: 'stage4-wire-unresolved', profile: 'D' },
    )
    let probed = 0
    try {
      await session.done
      const source = createPagedWorkbookSource(session)
      const sheetId = source.sheets[0]!.sheetId
      const store = source.store(sheetId)
      const memory = createCfMemory()
      for (const rule of unresolved) {
        const rects = rule.ranges.map(r => ({ r0: r.firstRow, r1: Math.min(r.lastRow, r.firstRow + 4), c0: r.firstCol, c1: r.firstCol }))
        const result = await store.readConditionalFormatting?.({
          rects,
          stamp: { owner: 'audit', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 },
          grantBytes: 2 * 1024 * 1024,
        }, memory, { signal: new AbortController().signal })
        expect(result, rule.label).toBeDefined()
        // The worker sends no icon for an unresolved rule's cells, whether it refuses the whole read or the rule.
        expect(result!.states.filter(s => s.icon !== undefined), rule.label).toEqual([])
        if (result!.kind === 'unresolved') {
          const { budget, net } = trackingBudget()
          await expect(prepareCfWireScope(async () => ({ result: result as CfReadResult, release: () => {} }), rects, { locale: 'en-US', budget }))
            .rejects.toThrow('CF wire preparation is unresolved')
          expect(net()).toBe(0)
        }
        for (const r of rects) probed += r.r1 - r.r0 + 1
      }
      memory.close()
    } finally {
      await session.release()
    }
    expect(probed).toBeGreaterThan(20)
  }, 120000)
})

// The Stage 4 corpus inventory (docs/validation/xlsx-cf.md §4): no corpus workbook carries an icon set.
const CORPUS_ICONSET_FILES = 0
const corpusFiles = existsSync(CORPUS) ? readdirSync(CORPUS).filter(f => f.endsWith('.xlsx') && !f.startsWith('~$')).sort() : []

describe.skipIf(!existsSync(CORPUS) || corpusFiles.length === 0)('corpus iconSet files parse identically on both paths', () => {
  it('the iconSet file count equals the §4 inventory', async () => {
    const { readZipPartsSync } = await import('../scripts/cf-refs/verify')
    let withIcons = 0
    for (const file of corpusFiles) {
      const parts = readZipPartsSync(new Uint8Array(readFileSync(join(CORPUS, file))))
      expect(parts, file).not.toBeNull()
      const sheets = [...parts!.entries()].filter(([name]) => /^xl\/worksheets\/[^/]+\.xml$/.test(name))
      if (sheets.some(([, xml]) => /<(x14:)?iconSet\b/.test(new TextDecoder().decode(xml)))) withIcons++
    }
    expect(withIcons).toBe(CORPUS_ICONSET_FILES)
  })

  it.each(corpusFiles)('%s: resident and streamed CfSheetRules deep-equal, nothing throws', async file => {
    const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
    const bytes = new Uint8Array(readFileSync(join(CORPUS, file)))
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes.slice().buffer as ArrayBuffer]) }, { docId: `stage4-corpus-${file}`, profile: 'D' })
    try {
      await session.done.catch(() => undefined)
      for (const paged of session.workbook.sheets) {
        if (paged.kind !== 'worksheet') continue
        const resident = doc.sheets.find(s => s.name === paged.name)
        expect(paged.conditionalFormatting, `${file} ${paged.name}`).toEqual(resident?.conditionalFormatting)
      }
    } finally {
      await session.release()
    }
  }, 120000)
})

// The Stage 4 oracle workbooks (as Excel saved them, and as authored) are the icon writer shapes this repository holds.
const ICON_WORKBOOKS = [...new Set(manifest().cases.filter(c => c.stage === 4).flatMap(c => [c.saved, c.input]))].filter(p => existsSync(join(REPO_ROOT, p)))
describe('Stage 4 oracle workbooks parse identically on both paths', () => {
  it.each(ICON_WORKBOOKS)('%s', async path => {
    const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
    const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, path)))
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes.slice().buffer as ArrayBuffer]) }, { docId: `stage4-both-${path}`, profile: 'D' })
    try {
      await session.done
      for (const paged of session.workbook.sheets) {
        if (paged.kind !== 'worksheet') continue
        const resident = doc.sheets.find(s => s.name === paged.name)
        expect(paged.conditionalFormatting, `${path} ${paged.name}`).toEqual(resident?.conditionalFormatting)
      }
    } finally {
      await session.release()
    }
  }, 120000)
})
