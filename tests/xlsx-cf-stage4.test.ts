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
  baselineNoCf,
  checkPressureSample,
  iconRuleOf,
  pressureBarUnitCap,
  pressureCycles,
  pressureIcon1MRowSharedScan,
  pressureIconDistinctRanges,
  pressureIconGlyphAdmission,
  pressureIconWholeColumnExtents,
  pressureRefusedPercentileIconWithNum,
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

/**
 * Carried decision (Task 29 review, parked to Task 30): a NoIcons cell with showValue="0". The canvas hides its value
 * (`icon:noicons-showvalue` reads `text-hidden`, published), nothing paints, so nothing is announced as an icon, and it
 * is not "unavailable" (the outcome is evidenced, not withheld). The value stays in the accessible name and in copy
 * text: the same as a Stage 3 bar with showValue="0" (S3-O), which hides canvas text only.
 */
describe('NoIcons with showValue="0" announces no icon and no unavailability; the value stays in copy text', () => {
  it('stage4-icon-inset G2:G4 under production evidence', async () => {
    const sheet = await savedSheet('tests/fixtures/xlsx-cf/stage4-icon-inset.xlsx')
    const { budget, net } = trackingBudget()
    const rects = [{ r0: 1, r1: 3, c0: 5, c1: 6 }]
    const scope = (await prepareCfTextScope(sheet, { rects, signal: new AbortController().signal, locale: 'en-US', budget }))!
    expect(scope).toBeDefined()
    for (const row of [1, 2, 3]) for (const col of [5, 6]) {
      const where = `${col === 5 ? 'F' : 'G'}${row + 1}`
      expect(scope.iconAt?.(row, col), where).toBeUndefined()
      const text = scope.describe(row, col)?.text
      expect(text, where).not.toBe(UNAVAILABLE)
      expect(text ?? '', where).not.toMatch(/Icon \d of \d|圖示/)
    }
    // The canvas hides G's value (showValue="0") and keeps F's.
    const cell = (row: number, col: number) => sheet.rows.find(r => r.index === row)!.cells.find(c => c.col === col)!
    const host = (row: number, col: number) => ({ x: 0, y: 0, width: 200, height: 20, textWidth: 10, cell: cell(row, col) })
    for (const row of [1, 2, 3]) {
      expect(scope.hidesTextAt?.(row, 6, host(row, 6), { zoom: 1, dpr: 1 }), `G${row + 1}`).toBe(true)
      expect(scope.hidesTextAt?.(row, 5, host(row, 5), { zoom: 1, dpr: 1 }), `F${row + 1}`).toBe(false)
      expect(cell(row, 6).value, `G${row + 1} keeps its value for copy and the accessible name`).toEqual(expect.anything())
    }
    scope.dispose()
    expect(net()).toBe(0)
  })
})

describe('Stage 4 pressure scenarios and resource bounds', () => {
  it('a 1M-row DECLARED A:A scope with 100 icon sets, 100 bars and 100 scales shares one scan and evaluates the first icon', async () => {
    const sample = await pressureIcon1MRowSharedScan()
    checkPressureSample(sample, { expectOutcome: 'admitted' })
    expect(sample.statisticsReads).toBe(100)
    expect(sample.cellsWithResults).toBe(40)
    expect(sample.iconOnlyCells).toBe(40)
    expect(sample.netBytes).toBe(0)
  })

  it('a refused percentile icon set withholds while a num icon set still evaluates (control: with room it evaluates too)', async () => {
    const sample = await pressureRefusedPercentileIconWithNum()
    checkPressureSample(sample, { expectOutcome: 'admitted' })
    expect(sample.cellsWithResults).toBe(10)
    expect(sample.percentileWithheld).toBe(10)
    expect(sample.controlPercentileIcons).toBe(10)
    expect(sample.netBytes).toBe(0)
  })

  it('a whole-column icon set does not extend extents (isolation)', async () => {
    const sample = await pressureIconWholeColumnExtents()
    checkPressureSample(sample, { expectOutcome: 'admitted' })
    expect(sample.cellsWithResults).toBe(10)
  })

  it('glyph admission with every class the sidecar records in one batch: admitted classes decode once, withheld never, refused nothing', async () => {
    const g = await pressureIconGlyphAdmission()
    expect(g.classes).toBe(54)
    expect(g.admitted).toBe(CF_ICON_GLYPH_CLASSES && new Set(Object.values(CF_ICON_GLYPH_CLASSES).filter(c => !CF_ICON_GLYPH_WITHHELD.has(c))).size)
    expect(g.admitted).toBe(39)
    expect(g.decodes).toBe(39)
    expect(g.withheldDecoded).toBe(0)
    expect(g.refusedDecodes).toBe(0)
    expect(g.peakBytes).toBeGreaterThan(0)
    expect(g.netBytes).toBe(0)
  })

  it('distinct-range icon pressure refuses past the rect cap; the shared statistics scan refuses past the 1,000,000-unit cap', async () => {
    checkPressureSample(await pressureIconDistinctRanges(), { expectOutcome: 'refused' })
    const cap = await pressureBarUnitCap()
    expect(CF_LIMITS.workUnits).toBe(1_000_000)
    expect(cap.refusedReason).toBe('cf-resource-limit')
    expect(cap.refusedReads).toBeLessThan(cap.refusedCells)
    expect(cap.netBytes).toBe(0)
  }, 120000)

  it('20 open/read/cancel/close cycles over an icon range return to zero', async () => {
    const cycles = await pressureCycles(20, { fixture: 'tests/fixtures/xlsx-cf/stage4-icon-thresholds.xlsx', rect: { r0: 1, r1: 11, c0: 1, c1: 1 }, requireIcons: true })
    expect(cycles.cycles).toBe(20)
    expect(cycles.iconStates).toBeGreaterThanOrEqual(20)
    expect(cycles.finalNetBytes).toBe(0)
    expect(cycles.maxPeakBytes).toBeGreaterThan(0)
  }, 120000)

  it('zero icon work, glyph decodes or reads on a no-CF baseline', async () => {
    const baseline = await baselineNoCf(iconRuleOf(0, { firstRow: 0, lastRow: 99, firstCol: 0, lastCol: 0 }) as unknown as Record<string, unknown>)
    expect(baseline).toEqual({ paintReserves: 0, sourceReads: 0, scopeBuilt: false, glyphDecodes: 0 })
  })
})

describe('S4-ROW-UNIFORM: the resident sheet and the worker reader agree on each row\'s font uniformity', () => {
  it.each([
    ['stage4-icon-rows-bottom', [] as { start: number; end: number }[]],
    // Rows 2..4 hold the 8 and 20 pt cells of columns B, D, E, G, H, J; row 5 holds the 6 pt short-row case's 11 pt cells only.
    ['stage4-icon-geometry', null],
  ])('%s', async (id, expected) => {
    const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
    const { cfNonUniformFontRows } = await import('../src/xlsx/rowFont')
    const path = `tests/fixtures/xlsx-cf/${id}.xlsx`
    const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, path)))
    const rs = await savedSheet(path); const resident = cfNonUniformFontRows(rs.rows, rs.rowFontsBlocked)
    if (expected) expect(resident).toEqual(expected)
    else expect(resident.some(r => r.start <= 2 && r.end >= 2), 'row 3 (C3, M3..O3) holds a 20 pt cell').toBe(true)
    const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes.slice().buffer as ArrayBuffer]) }, { docId: `row-uniform-${id}`, profile: 'D' })
    try {
      await session.done
      const paged = session.workbook.sheets.find(s => s.kind === 'worksheet') as unknown as { metadata: { nonUniformFontRows?: unknown } }
      expect(paged.metadata.nonUniformFontRows).toEqual(resident)
    } finally {
      await session.release()
    }
  }, 120000)
})

/**
 * Fix round 2 (codex P1 ×2): the effective font of a row's empty cells and of rich-text runs. Codex's variants of
 * stage4-icon-rows-bottom, each on both paths (resident parse and the worker's record reader): a 20 pt column style; a
 * 20 pt Normal style (xf 0) with every stored cell explicitly Calibri 11; a neighbouring inline run in Calibri 20; a
 * neighbouring shared-string run in Calibri 20. Each withholds B2's bottom icon; the unmodified workbook paints it.
 */
describe('S4-ROW-UNIFORM covers column and default styles and rich-text runs on both paths (fix round 2)', () => {
  const BASE = 'tests/fixtures/xlsx-cf/stage4-icon-rows-bottom.xlsx'
  const FONT20 = '<font><sz val="20"/><color rgb="FF000000"/><name val="Calibri"/></font>'
  const FONT11 = '<font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>'
  const RUN20 = '<r><rPr><sz val="20"/><rFont val="Calibri"/></rPr><t>x</t></r>'
  const PHONETIC = '<rPh sb="0" eb="1"><t>p</t></rPh><phoneticPr fontId="2" type="noConversion"/>'
  const edit = (part: string, fn: (xml: string) => string) => ({ part, fn })
  const fonts = (xml: string, extra: string, first?: string) => xml.replace(/<fonts count="\d+"([^>]*)>(<font>.*?<\/font>)/, (_m, a, f0) => `<fonts count="3"${a}>${first ?? f0}`).replace('</fonts>', `${extra}</fonts>`)
  const VARIANTS: Record<string, ReturnType<typeof edit>[]> = {
    'a 20 pt column style': [
      edit('xl/styles.xml', x => fonts(x, FONT20).replace(/<cellXfs count="3">(.*?)<\/cellXfs>/, '<cellXfs count="4">$1<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>')),
      edit('xl/worksheets/sheet1.xml', x => x.replace('</cols>', '<col min="8" max="8" width="10" style="3" customWidth="1"/></cols>')),
    ],
    'a 20 pt Normal style with explicitly Calibri 11 cells': [
      edit('xl/styles.xml', x => fonts(x, FONT11, FONT20).replace(/(<cellXfs count="3"><xf [^>]*>.*?<\/xf>)<xf numFmtId="0" fontId="0"/, '$1<xf numFmtId="0" fontId="2"').replace(/(<cellXfs count="3"><xf [^>]*>.*?<\/xf><xf [^>]*>.*?<\/xf>)<xf numFmtId="0" fontId="0"/, '$1<xf numFmtId="0" fontId="2"')),
      edit('xl/worksheets/sheet1.xml', x => x.replace(/<c r="([A-Z]+\d+)" t="s">/g, '<c r="$1" s="2" t="s">').replace(/<c r="([A-Z]+\d+)">/g, '<c r="$1" s="2">')),
    ],
    'a neighbouring inline run in Calibri 20': [
      edit('xl/worksheets/sheet1.xml', x => x.replace(/(<row r="2"[^>]*>.*?)<\/row>/, `$1<c r="G2" s="2" t="inlineStr"><is>${RUN20}</is></c></row>`)),
    ],
    // Fix round 3 (codex P1): phonetic runs, a phoneticPr naming a 20 pt font, and ph="1".
    'a neighbouring inline phonetic run with a 20 pt phoneticPr font': [
      edit('xl/styles.xml', x => fonts(x, FONT20)),
      edit('xl/worksheets/sheet1.xml', x => x.replace(/(<row r="2"[^>]*>.*?)<\/row>/, `$1<c r="G2" s="2" t="inlineStr" ph="1"><is><t>x</t>${PHONETIC}</is></c></row>`)),
    ],
    'a neighbouring shared-string phonetic run with a 20 pt phoneticPr font': [
      edit('xl/styles.xml', x => fonts(x, FONT20)),
      edit('xl/sharedStrings.xml', x => x.replace('</sst>', `<si><t>y</t>${PHONETIC}</si></sst>`)),
      edit('xl/worksheets/sheet1.xml', x => x.replace(/(<row r="2"[^>]*>.*?)<\/row>/, '$1<c r="G2" s="2" t="s" ph="1"><v>RICH</v></c></row>')),
    ],
    'a neighbouring plain cell showing phonetics (ph="1")': [
      edit('xl/worksheets/sheet1.xml', x => x.replace(/(<row r="2"[^>]*>.*?)<\/row>/, '$1<c r="G2" s="2" ph="1"><v>1</v></c></row>')),
    ],
    'a neighbouring shared-string run in Calibri 20': [
      edit('xl/sharedStrings.xml', x => x.replace('</sst>', `<si>${RUN20}</si></sst>`)),
      edit('xl/worksheets/sheet1.xml', x => x.replace(/(<row r="2"[^>]*>.*?)<\/row>/, '$1<c r="G2" s="2" t="s"><v>RICH</v></c></row>')),
    ],
  }
  async function variant(edits: ReturnType<typeof edit>[]): Promise<Uint8Array> {
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(readFileSync(join(REPO_ROOT, BASE)))
    const sstCount = ((await zip.file('xl/sharedStrings.xml')!.async('string')).match(/<si>/g) ?? []).length
    for (const e of edits) {
      const before = await zip.file(e.part)!.async('string')
      const after = e.fn(before).replace('<v>RICH</v>', `<v>${sstCount}</v>`)
      expect(after, `${e.part} edited`).not.toBe(before)
      zip.file(e.part, after)
    }
    return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
  }
  async function bothPaths(bytes: Uint8Array) {
    const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
    const { cfNonUniformFontRows } = await import('../src/xlsx/rowFont')
    const { prepareCfPaint } = await import('../src/xlsx/cf/render')
    const { createCanvas } = await import('canvas')
    const sheet = (await parseXlsx(await OfficePackage.load(bytes))).sheets[0]! as unknown as XlsxSheet
    const resident = cfNonUniformFontRows(sheet.rows, sheet.rowFontsBlocked)
    const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes.slice().buffer as ArrayBuffer]) }, { docId: `row-font-${resident.length}-${Math.random()}`, profile: 'D' })
    let worker: unknown
    try {
      await session.done
      worker = (session.workbook.sheets.find(s => s.kind === 'worksheet') as unknown as { metadata: { nonUniformFontRows?: unknown } }).metadata.nonUniformFontRows
    } finally { await session.release() }
    const ctx = createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
    const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext: () => ctx }))!
    const b2 = batch.iconAt(1, 1) !== undefined
    batch.dispose()
    return { resident, worker, b2 }
  }

  it('the unmodified workbook: no non-uniform row on either path, B2 paints', async () => {
    const got = await bothPaths(new Uint8Array(readFileSync(join(REPO_ROOT, BASE))))
    expect(got).toEqual({ resident: [], worker: [], b2: true })
  }, 120000)

  it.each(Object.keys(VARIANTS))('%s: row 2 is non-uniform on both paths, and B2 withholds', async name => {
    const got = await bothPaths(await variant(VARIANTS[name]!))
    expect(got.worker, 'worker equals resident').toEqual(got.resident)
    expect(got.resident.some(r => r.start <= 1 && r.end >= 1), 'row 2 is non-uniform').toBe(true)
    if (name.includes('style')) expect(got.resident).toEqual([{ start: 0, end: 1_048_575 }])
    else expect(got.resident).toEqual([{ start: 1, end: 1 }])
    expect(got.b2, 'B2 bottom icon withheld').toBe(false)
  }, 120000)
})

/** Fix round 2 (codex P2): the per-case/variant validation records `audit --stage 4` emits. */
describe('Stage 4 per-case validation records are complete', () => {
  it('every measured Stage 4 case and every designated cell (variant) has a row with measured and produced fields', async () => {
    const { stage4CaseRecords } = await import('../scripts/cf-refs/audit')
    const m = manifest()
    const rows = await stage4CaseRecords(m, REPO_ROOT)
    const measured = m.cases.filter(c => c.stage === 4 && c.status === 'measured')
    expect(measured.length).toBeGreaterThan(10)
    const keys = new Set(rows.map(r => `${r.caseId}:${r.ref}`))
    for (const c of measured) for (const e of c.expected.cells) expect(keys.has(`${c.id}:${e.ref}`), `${c.id}:${e.ref}`).toBe(true)
    expect(rows).toHaveLength(measured.reduce((n, c) => n + c.expected.cells.length, 0))
    for (const r of rows) {
      expect(r.mode).toBe('resident')
      expect(r.stamp.cfGen, r.caseId).toBeGreaterThan(0)
      expect(r.bytePeak, r.caseId).toBeGreaterThan(0)
      expect(r.workUnits, r.caseId).toBeGreaterThan(0)
      expect(['withheld', null].includes(r.produced.glyphClass as never) || typeof r.produced.glyphClass === 'string', r.ref).toBe(true)
      // Every painted icon carries its produced box, within +-1 px of Excel's measured one where Excel measured one.
      if (r.produced.paints) {
        expect(r.produced.box, `${r.caseId}:${r.ref}`).toBeDefined()
        if (r.measured.box) for (const k of ['l', 't', 'r', 'b'] as const) expect(Math.abs(r.produced.box![k] - r.measured.box[k]), `${r.caseId}:${r.ref} ${k}`).toBeLessThanOrEqual(1)
      }
    }
    expect(rows.filter(r => r.produced.paints).length).toBeGreaterThanOrEqual(39)
  }, 300000)
})
