// @vitest-environment node
/**
 * Task 35: every reading production registers (`CF_FORMULA_PRODUCTION_READINGS`) is the reading production implements,
 * shown on the Stage 5 fixtures themselves. Under the registered capabilities (and each listed selection in turn), the
 * fill, strikethrough and visual production shows at each designated cell of each configuration equal that reading's
 * predicted token (a `twin:` token resolved against production's own twin cell), or the cell is withheld (fail closed:
 * an unadmitted function, a whole-rule hoist refusal). Each entry is reproduced in full by at least one configuration,
 * except the named unreachable ones, and a rival reading is contradicted there, so the comparison discriminates.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cfDocumentEvaluationOf, cfEvaluationContextFor } from '../src/xlsx/cf/formula'
import { cfWorkbookOfDocument, residentSource } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { createCfService } from '../src/xlsx/cf/service'
import { CF_LIMITS, createCfMemory } from '../src/xlsx/cf/limits'
import { injectCfFormulaSelections, withInjectedCfEvidence } from '../src/xlsx/cf/evidence'
import { CF_FORMULA_PRODUCTION_READINGS } from '../src/xlsx/cf/formulaReadings'
import { cfRulesVersion } from '../src/xlsx/view/pagedSource'
import { CF_STAGE5_OWNERS, type CfManifest } from '../scripts/cf-refs/cases'
import { CF_STAGE5_UNCONFIRMABLE, stage5Publication, stage5UnconfirmableScope, auditStage5Manifest, type Stage5AuditReport } from '../scripts/cf-refs/audit'
import { rc, shown } from './helpers/cf-formula-rules'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
/** Registered entries no Stage 5 configuration reproduces in full yet, with why (the cells it cannot reach withhold, never paint). */
const UNREACHABLE: Readonly<Record<string, string>> = {
  'formula:truth-error': 'each configuration holds an NA() row, and NA is not on the allow-list (S5-B): those cells withhold, the 1/0 rows show error-no-match',
  'formula:fn-conditional-aggregate': 'its configuration\'s COUNTIF($A:$A,A2) column holds booleans and numeric text, a whole-rule hoist refusal (S5-HOIST); reproduced through the stage5-formula-hoist rows',
}
const REUSE_ONLY = /^formula:composition-stub-successor-/

type Config = NonNullable<CfManifest['cases'][number]['formulaConfigurations']>[number]
/** The refs of a configuration that production shows on open: a bare ref, or `pre:X` (what Excel shows before any calculate,
 * which is what production shows: it never calculates). `calc:` and `post:` observations follow an explicit calculation and
 * are not paint. */
const isPaintRef = (r: string): boolean => !r.includes(':') || r.startsWith('pre:')
const cellOfRef = (r: string): string => r.replace(/^pre:/, '')
/** The instant a clock fixture is evaluated at: 03:00 on its anchor local date in its required zone (inside the capture window, where the
 * local and UTC dates differ), so `TODAY()` is the anchor and the readings discriminate. */
function clockFor(c: CfManifest['cases'][number]): { zone: string; now: number } | undefined {
  const req = (c as unknown as { clockRequirement?: { anchorSerial?: number; requiredTimeZone?: string } }).clockRequirement
  if (!req?.anchorSerial || !req.requiredTimeZone) return undefined
  const midnightUtc = (req.anchorSerial - 25569) * 86_400_000
  const wall = (epoch: number): number => { const p = Object.fromEntries(new Intl.DateTimeFormat('en-GB', { timeZone: req.requiredTimeZone, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' }).formatToParts(new Date(epoch)).map(x => [x.type, Number(x.value)])); return Date.UTC(p.year!, p.month! - 1, p.day!, p.hour!, p.minute!, p.second!) }
  const offset = wall(midnightUtc) - midnightUtc
  return { zone: req.requiredTimeZone, now: midnightUtc + 3 * 3_600_000 - offset }
}
/** What production shows at one cell. */
export type Shown = Readonly<{ fill: string; strike: boolean; visual: string; withheld: boolean
  /** A bar or icon is drawn: a visual descriptor that is not a zero-length bar (S5-WITNESS: a bar at its lower bound is empty). */
  drawn: boolean }>
const tokenOf = (cfg: Config, s: Shown | undefined): string => !s ? 'missing' : s.withheld ? 'withheld' : `${s.fill}${s.strike ? '+strike' : ''}${cfg.observable === 'crop' ? `|${s.visual}` : ''}`
/** The cells a configuration's verdict reads: its designated cells, every twin its hypotheses name, every declared witness. */
const cellsOf = (cfg: Config): string[] => {
  const refs = cfg.refs.filter(isPaintRef).map(cellOfRef)
  const twins = Object.values(cfg.hypotheses).flatMap(h => Object.values(h)).filter(t => t.startsWith('twin:')).map(t => t.slice(5).replace('+strike', ''))
  return [...new Set([...refs, ...twins, ...Object.keys(cfg.twinExpect ?? {})])]
}
/**
 * The verdict of one configuration under one reading, pure in what production showed (`got`). Returns its violations,
 * whether the reading is reproduced in full (no designated cell withheld), and whether a rival reading is contradicted.
 */
export function configurationVerdict(cfg: Config, reading: string, got: ReadonlyMap<string, Shown>): { violations: string[]; unusable: string[]; reproduced: boolean; discriminating: boolean } {
  const refs = cfg.refs.filter(isPaintRef)
  const predicted = cfg.hypotheses[reading] as Readonly<Record<string, string>>
  const violations: string[] = []
  // Ruling S5-WITNESS: every declared witness is validated against its own declaration before any twin observation is
  // used as an expectation. A crop witness is about its bar or icon (`visual`: one is drawn; `empty`: none, or a
  // zero-length bar); a fill witness about its fill or strikethrough. A withheld witness satisfies neither. A configuration
  // whose witnesses do not hold cannot be evidence: it is neither reproduced nor discriminating (`unusable`).
  const unusable: string[] = []
  for (const [ref, expected] of Object.entries(cfg.twinExpect ?? {})) {
    const w = got.get(ref)
    if (expected === 'either') continue
    const renders = w !== undefined && !w.withheld && (cfg.observable === 'crop' ? w.drawn : w.fill !== 'none' || w.strike)
    const empty = w !== undefined && !w.withheld && !renders
    if (expected === 'visual' ? !renders : !empty) unusable.push(`${cfg.id} witness ${ref}: declared ${expected}, production shows ${tokenOf(cfg, w)}`)
  }
  const witnessesHold = unusable.length === 0
  const want = (h: Readonly<Record<string, string>>, ref: string): string => {
    const p = h[ref]!, strike = p.endsWith('+strike'), base = p.replace('+strike', '')
    if (!base.startsWith('twin:')) return `${base}${strike ? '+strike' : ''}`
    const twin = tokenOf(cfg, got.get(base.slice(5)))
    return cfg.observable === 'crop' ? twin : `${twin.replace('+strike', '')}${strike ? '+strike' : ''}`
  }
  const bad = refs.filter(r => tokenOf(cfg, got.get(cellOfRef(r))) !== 'withheld' && tokenOf(cfg, got.get(cellOfRef(r))) !== want(predicted, r))
  for (const r of bad) violations.push(`${cfg.id} ${reading} ${r}: production ${tokenOf(cfg, got.get(cellOfRef(r)))}, reading predicts ${want(predicted, r)}`)
  const reproduced = witnessesHold && violations.length === 0 && refs.every(r => tokenOf(cfg, got.get(cellOfRef(r))) !== 'withheld')
  const discriminating = reproduced && Object.entries(cfg.hypotheses).some(([name, h]) => name !== reading && refs.some(r => want(h, r) !== tokenOf(cfg, got.get(cellOfRef(r)))))
  if (reproduced && !discriminating) violations.push(`${cfg.id}: no rival of ${reading} is contradicted (the comparison does not discriminate)`)
  return { violations, unusable, reproduced, discriminating }
}

/**
 * Fix round 2 (ruling a): a capability whose fixture production cannot confirm (`CF_STAGE5_UNCONFIRMABLE`) is never
 * reproduced, and the audit refuses to publish it. Its configuration is still checked: no violation, no designated cell
 * withheld, and the only failing witnesses are exactly the documented ones. Returns the violations of that narrow check.
 */
export function unconfirmableViolations(cfg: Config, verdict: ReturnType<typeof configurationVerdict>, got: ReadonlyMap<string, Shown>, caseId: string): string[] {
  const scope = stage5UnconfirmableScope(cfg.capabilityId, caseId)
  const allowed = new Set(scope?.witnesses ?? [])
  const out = [...verdict.violations]
  for (const reason of verdict.unusable) {
    const ref = /witness (\S+):/.exec(reason)?.[1]
    if (!ref || !allowed.has(ref)) out.push(`${reason} (not a documented unconfirmable witness)`)
  }
  // Task 37 fix round 1: an entry whose designated cells production withholds (fail closed) accepts them withheld.
  if (!scope?.designatedWithheld) for (const r of cfg.refs.filter(isPaintRef).map(cellOfRef)) if (got.get(r)?.withheld !== false) out.push(`${cfg.id} ${r}: withheld`)
  return out
}

describe('the configuration verdict validates declared witnesses (fix round 1 #4)', () => {
  it('suppressing stop-bar E4 and its required positive twin G4 fails (bar-only twin G since Task 39)', () => {
    const manifest = JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
    const cfg = manifest.cases.find(x => x.id === 'stage5-composition')!.formulaConfigurations!.find(x => x.id === 'stop-bar')!
    expect(cfg.twinExpect?.['G4']).toBe('visual')
    // An in-memory world that satisfies `halts` everywhere: on rows 3/5/7 E shows the red-only twin F, elsewhere the bar-only twin G.
    const red: Shown = { fill: 'FFFFC7CE', strike: false, visual: '', withheld: false, drawn: false }, bar: Shown = { fill: 'none', strike: false, visual: 'dataBar', withheld: false, drawn: true }
    const redBar: Shown = { fill: 'FFFFC7CE', strike: false, visual: 'dataBar', withheld: false, drawn: true }, empty: Shown = { fill: 'none', strike: false, visual: '', withheld: false, drawn: false }
    // G2 holds the column minimum: a zero-length bar, declared empty.
    const zeroBar: Shown = { fill: 'none', strike: false, visual: 'dataBar', withheld: false, drawn: false }
    const world = new Map<string, Shown>()
    for (const r of [2, 3, 4, 5, 6, 7]) {
      const on = r % 2 === 1
      world.set(`E${r}`, on ? red : r === 2 ? zeroBar : bar); world.set(`F${r}`, red); world.set(`H${r}`, redBar); world.set(`G${r}`, r === 2 ? zeroBar : bar)
    }
    expect(configurationVerdict(cfg, 'halts', world)).toMatchObject({ violations: [], unusable: [], reproduced: true, discriminating: true })
    // Codex counterexample: E4 and G4 both suppressed. Production's twin alone would call it reproduced.
    const suppressed = new Map(world)
    suppressed.set('E4', empty); suppressed.set('G4', empty)
    const v = configurationVerdict(cfg, 'halts', suppressed)
    expect(v.reproduced).toBe(false)
    expect(v.discriminating).toBe(false)
    expect(v.unusable).toEqual(['stop-bar witness G4: declared visual, production shows none|'])
  })
  it('stop-icon: only the documented L3/L5/L7 witnesses may fail; suppressing I4 and its positive twin K4 fails (fix round 2; letters since Task 39)', () => {
    const manifest = JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
    const cfg = manifest.cases.find(x => x.id === 'stage5-composition')!.formulaConfigurations!.find(x => x.id === 'stop-icon')!
    expect(CF_STAGE5_UNCONFIRMABLE['formula:composition-stop-icon']?.witnesses).toEqual(['L3', 'L5', 'L7'])
    const red: Shown = { fill: 'FFFFC7CE', strike: false, visual: '', withheld: false, drawn: false }, icon: Shown = { fill: 'none', strike: false, visual: 'iconSet', withheld: false, drawn: true }
    const withheld: Shown = { fill: 'none', strike: false, visual: '', withheld: true, drawn: false }, empty: Shown = { fill: 'none', strike: false, visual: '', withheld: false, drawn: false }
    // Production's world: halts everywhere, the red + icon twins L withheld (icon:composition-dxf-fill unpublished).
    const world = new Map<string, Shown>()
    for (const r of [2, 3, 4, 5, 6, 7]) { const on = r % 2 === 1; world.set(`I${r}`, on ? red : icon); world.set(`J${r}`, red); world.set(`L${r}`, withheld); world.set(`K${r}`, icon) }
    const v = configurationVerdict(cfg, 'halts', world)
    expect(v.reproduced).toBe(false)
    expect(unconfirmableViolations(cfg, v, world, 'stage5-composition')).toEqual([])
    // Codex counterexample: I4 and its required positive twin K4 both suppressed.
    const suppressed = new Map(world)
    suppressed.set('I4', empty); suppressed.set('K4', empty)
    const sv = configurationVerdict(cfg, 'halts', suppressed)
    expect(sv.reproduced).toBe(false)
    expect(unconfirmableViolations(cfg, sv, suppressed, 'stage5-composition')).toEqual([expect.stringMatching(/witness K4: declared visual.*not a documented unconfirmable witness/)])
  })
})

describe('unconfirmable scopes are per case (the bar:bounds-formula follow-up, orchestrator ruling 4)', () => {
  const manifest = JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
  it('only the superseded stage5-cfvo-formulas keeps a scope; the owner stage5-bar-bounds-formula is confirmed in full', () => {
    expect(stage5UnconfirmableScope('bar:bounds-formula', 'stage5-cfvo-formulas')).toMatchObject({ witnesses: ['W2', 'W3', 'W4', 'W5', 'W6', 'W7', 'W8', 'W9'], designatedWithheld: true })
    expect(stage5UnconfirmableScope('bar:bounds-formula', 'stage5-bar-bounds-formula')).toBeUndefined()
    expect(stage5UnconfirmableScope('bar:bounds-formula', 'stage5-composition')).toBeUndefined()
    expect(stage5UnconfirmableScope('formula:composition-stop-icon', 'stage5-composition')?.witnesses).toEqual(['L3', 'L5', 'L7'])
  })
  it('the old scope still refuses its own misreadings', () => {
    const cfg = manifest.cases.find(x => x.id === 'stage5-cfvo-formulas')!.formulaConfigurations!.find(x => x.id === 'cfvo-bar-formula')!
    const bar: Shown = { fill: 'none', strike: false, visual: 'dataBar', withheld: false, drawn: true }
    const withheld: Shown = { fill: 'none', strike: false, visual: '', withheld: true, drawn: false }
    // Production today: I, W and X withhold (beyond bounds, degenerate); Y paints.
    const world = new Map<string, Shown>()
    for (let r = 2; r <= 9; r++) { world.set(`I${r}`, withheld); world.set(`W${r}`, withheld); world.set(`X${r}`, withheld); world.set(`Y${r}`, r === 2 ? { ...bar, drawn: false } : bar) }
    const check = (w: Map<string, Shown>) => unconfirmableViolations(cfg, configurationVerdict(cfg, 'evaluated', w), w, 'stage5-cfvo-formulas')
    expect(check(world)).toEqual([])
    // A painted formula bar that contradicts the evaluated reading's twin is a misreading, not the documented withholding.
    const painted = new Map(world); painted.set('I5', bar)
    expect(check(painted)).toEqual([expect.stringMatching(/^cfvo-bar-formula evaluated I5: production none\|dataBar, reading predicts withheld$/)])
    // A witness outside the documented W2..W9 never fails quietly.
    const yGone = new Map(world); yGone.set('Y5', withheld)
    expect(check(yGone).length).toBeGreaterThan(0)
  })
  it('the superseded scope does not block publication: bar:bounds-formula publishes from its owner', () => {
    const base = auditStage5Manifest({ cases: [] } as never, { rootDir: REPO_ROOT })
    const report: Stage5AuditReport = { ...base, ok: true, stageComplete: true, blocked: [], blockedCapabilities: [],
      perCapability: [{ capabilityId: 'bar:bounds-formula', caseId: 'stage5-bar-bounds-formula', disposition: 'verified' }],
      matchedReadings: { 'bar:bounds-formula': 'evaluated' } }
    expect(stage5Publication(report).published).toEqual(['bar:bounds-formula'])
    const old: Stage5AuditReport = { ...report, perCapability: [{ capabilityId: 'bar:bounds-formula', caseId: 'stage5-cfvo-formulas', disposition: 'verified' }] }
    expect(stage5Publication(old).withheld['bar:bounds-formula']).toMatch(/fixture-unconfirmable: stage5-cfvo-formulas/)
  })
})

describe('an unconfirmable capability is never published (fix round 2)', () => {
  it('a verified halts report for formula:composition-stop-icon still does not publish it', () => {
    const base = auditStage5Manifest({ cases: [] } as never, { rootDir: REPO_ROOT })
    const report: Stage5AuditReport = { ...base, ok: true, stageComplete: true, blocked: [], blockedCapabilities: [],
      perCapability: [{ capabilityId: 'formula:composition-stop-icon', caseId: 'stage5-composition', disposition: 'verified' }, { capabilityId: 'formula:composition-stop-bar', caseId: 'stage5-composition', disposition: 'verified' }],
      matchedReadings: { 'formula:composition-stop-icon': 'halts', 'formula:composition-stop-bar': 'halts' } }
    const pub = stage5Publication(report)
    expect(pub.published).toEqual(['formula:composition-stop-bar'])
    expect(pub.withheld['formula:composition-stop-icon']).toMatch(/fixture-unconfirmable/)
  })
})

/** Runs every non-record configuration of every Stage 5 case under each reading of `registry`, as production evaluates it. */
async function readingsCheck(registry: Readonly<Record<string, string | readonly string[]>>) {
  const manifest = JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
  const lists = Object.entries(registry).filter((e): e is [string, readonly string[]] => typeof e[1] !== 'string')
  const reproduced = new Set<string>(), reproducedAt = new Set<string>(), unconfirmableChecked = new Set<string>(), violations: string[] = []
  let discriminating = 0
  for (const c of manifest.cases.filter(x => x.stage === 5)) {
    for (const cfg of c.formulaConfigurations ?? []) {
      if (cfg.record || cfg.role) continue
      const entry = registry[cfg.capabilityId]
      if (entry === undefined) continue
      for (const reading of typeof entry === 'string' ? [entry] : entry) {
        const predicted = cfg.hypotheses[reading] as Readonly<Record<string, string>> | undefined
        // Owner-aware (the cfvo-relative follow-up): the case owning the capability (CF_STAGE5_OWNERS) must pre-register
        // every registered reading; a superseded case keeps the historical verification of the readings it names, and is
        // never given a hypothesis after its capture.
        const owner = CF_STAGE5_OWNERS[cfg.capabilityId]
        if (predicted === undefined && owner !== undefined && owner !== c.id) continue
        expect(predicted, `${cfg.capabilityId} ${cfg.id}: ${reading} is a hypothesis`).toBeDefined()
        const refs = cfg.refs.filter(isPaintRef)
        if (refs.length === 0) continue
        const selections = Object.fromEntries(lists.map(([k, v]) => [k, v[0]!]))
        selections[cfg.capabilityId] = reading
        const got = await withInjectedCfEvidence(Object.keys(registry), async () => {
          injectCfFormulaSelections(selections)
          try {
            // A `pre:` observation is what Excel shows on opening the authored input, before any calculate; the saved file is Excel's resave
            // (its stale values recomputed), so the input is the workbook production must reproduce it on.
            const file = cfg.refs.some(r => r.startsWith('pre:')) ? (c as unknown as { input: string }).input : c.saved
            const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, file)))))
            const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
            const cells = cellsOf(cfg)
            const clock = clockFor(c)
            const memory = createCfMemory()
            // The document's other sheets (Task 36): cross-sheet configurations read them through the static workbook seam.
            const service = createCfService(sheet.conditionalFormatting!, memory, { evidence: cfRuntimeEvidence, evaluation: cfEvaluationContextFor(clock ? { ...cfDocumentEvaluationOf(doc), timeZone: clock.zone } : cfDocumentEvaluationOf(doc), clock ? () => clock.now : Date.now), sheetName: sheet.name, workbook: cfWorkbookOfDocument(doc.sheets, memory, 'engine') })
            const stamp = { owner: 'readings', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(sheet.conditionalFormatting!), viewGen: 0, structureGen: 0, clockGen: 0 }
            const lease = await service.prepareCf({ stamp, sourceRects: cells.map(r => { const { row, col } = rc(r); return { firstRow: row, lastRow: row, firstCol: col, lastCol: col } }), signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, residentSource(sheet, stamp, memory, 'engine'))
            const out = new Map<string, Shown>()
            for (const r of cells) {
              const { row, col } = rc(r)
              const o = lease.value.at(row, col), s = shown(o)
              const v = o.kind === 'ready' ? o.value.visual : undefined
              // Task 37: an icon's token names the icon it draws, so a crop configuration whose readings differ only in the
              // bucket (stage5-cfvo-scope scope-icon) can discriminate; a twin is still production's own twin cell.
              const drawn = v !== undefined && (v.iconSet !== undefined || (v.dataBar !== undefined && v.dataBar.endFraction !== v.dataBar.startFraction))
              // The cfvo-relative confirmation case: a bar's token names its extent too, so a crop configuration whose
              // readings differ only in bar length (stage5-cfvo-relative-a1 rel-a1-bar) can discriminate.
              const extent = (b: NonNullable<typeof v>['dataBar']): string => (b ? `dataBar:${b.startFraction.toFixed(4)}-${b.endFraction.toFixed(4)}` : 'dataBar')
              out.set(r, { fill: s.fill, strike: s.strike, withheld: s.withheld, drawn, visual: v ? Object.keys(v).sort().map(k => k === 'iconSet' && v.iconSet ? `iconSet:${v.iconSet.set}/${v.iconSet.iconId}` : k === 'dataBar' ? extent(v.dataBar) : k).join('&') : '' })
            }
            lease.release(); service.dispose()
            return out
          } finally { injectCfFormulaSelections(null) }
        })
        const verdict = configurationVerdict(cfg, reading, got)
        if (stage5UnconfirmableScope(cfg.capabilityId, c.id)) {
          // Never reproduced; checked narrowly (only its documented witnesses may fail) and refused publication by the audit.
          for (const v of unconfirmableViolations(cfg, verdict, got, c.id)) violations.push(`${cfg.capabilityId} ${c.id}/${v}`)
          unconfirmableChecked.add(cfg.capabilityId)
          continue
        }
        for (const v of verdict.violations) violations.push(`${cfg.capabilityId} ${c.id}/${v}`)
        if (verdict.reproduced) { reproduced.add(cfg.capabilityId); reproducedAt.add(`${cfg.capabilityId}|${c.id}|${reading}`) }
        if (verdict.discriminating) discriminating++
      }
    }
  }
  return { reproduced, reproducedAt, unconfirmableChecked, violations, discriminating }
}

/** One run of the check over production's registry, shared by the assertions below (it walks every Stage 5 fixture). */
let registryCheck: ReturnType<typeof readingsCheck> | undefined
const productionReadingsCheck = (): ReturnType<typeof readingsCheck> => (registryCheck ??= readingsCheck(CF_FORMULA_PRODUCTION_READINGS))

describe('every registered formula reading is the reading production implements, on the Stage 5 fixtures', () => {
  it('each designated cell is the registered reading\'s prediction or withheld; each entry is reproduced and discriminating', async () => {
    const registry = CF_FORMULA_PRODUCTION_READINGS
    const { reproduced, unconfirmableChecked, violations, discriminating } = await productionReadingsCheck()
    expect(violations).toEqual([])
    const unreproduced = Object.keys(registry).filter(id => !reproduced.has(id) && !REUSE_ONLY.test(id) && UNREACHABLE[id] === undefined && !unconfirmableChecked.has(id))
    expect([...unconfirmableChecked].sort()).toEqual(Object.keys(CF_STAGE5_UNCONFIRMABLE).sort())
    expect(unreproduced).toEqual([])
    for (const id of Object.keys(UNREACHABLE)) expect(reproduced.has(id), `${id} is now reachable: drop it from UNREACHABLE`).toBe(false)
    expect(discriminating).toBeGreaterThan(30)
  }, 300_000)
})

describe('the anchor-a1 registration (stage5-cfvo-relative-a1, f9d5be94)', () => {
  it('both cfvo scope ids name anchor-a1, reproduced and discriminating on the owner case and only where it was pre-registered', async () => {
    expect(CF_FORMULA_PRODUCTION_READINGS['formula:cfvo-relative']).toBe('anchor-a1')
    expect(CF_FORMULA_PRODUCTION_READINGS['formula:cfvo-scope-multi-range']).toEqual(['rule-wide', 'per-range', 'anchor-a1'])
    const { reproducedAt, violations } = await productionReadingsCheck()
    expect(violations).toEqual([])
    // The owner, and stage5-cfvo-scope-v2 (which named anchor-a1 before its capture); stage5-cfvo-formulas K and
    // stage5-cfvo-scope never did and keep their historical verification.
    expect([...reproducedAt].filter(k => k.endsWith('|anchor-a1')).sort()).toEqual([
      'formula:cfvo-relative|stage5-cfvo-relative-a1|anchor-a1',
      'formula:cfvo-scope-multi-range|stage5-cfvo-relative-a1|anchor-a1',
      'formula:cfvo-scope-multi-range|stage5-cfvo-scope-v2|anchor-a1',
    ])
  }, 300_000)
})
