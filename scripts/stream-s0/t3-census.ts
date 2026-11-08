/**
 * Stage-0 Task 3 — T3 census battery: run the census over every corpus unit
 * at zoom 1 / DPR 1 under napi (the §17 row-0 chosen server backend), and
 * write the `notInSubset` RED gate and per-unit face resolution.
 *
 * Produces `docs/validation/streaming-s0/t3.json`:
 *   { gate, status GREEN|RED|REPORTED, backend, gitSha, generatedAt, zoom, dpr,
 *     methods, props, notInSubset, faceResolution, recordText,
 *     classEvidence, unitsByClass, corpusFiles, unitsSeen, unitsFailed,
 *     fontStrings }
 *
 * Also (plan Task 3 + codex-r7 round-9 tail): for Task 2's painter-class
 * subset, the `unitsByClass` list records every (file, unit) whose paint used
 * `shadowBlur > 0`, `setTransform`, `createPattern`, or a canvas-source
 * `drawImage` — Task 2 runs those through the full zoom × DPR grid.
 */
import { execFileSync } from 'node:child_process'
import { createCanvas } from 'canvas'  // resolved by `bun run s0:napi`→preload (napi) or plain (node-canvas)
import { listCorpus, loadUnits, type CorpusFile } from './corpus'
import { wrapCensus, type CensusReport } from './census-proxy'
import { writeResult, resolveBackend } from './results'

const ZOOM = 1
const DPR = 1
/** Hard requirement from plan Task 3. */
const PROBE_STRING = 'Hxg 漢字 مرحبا'

interface FaceResolutionRow {
  font: string
  widthNode: number
  widthNapi: number
  relDiff: number
}

/**
 * napi's createCanvas resolved through the shim (`canvas` spec is rewritten
 * by the chrome preload in bun scripts). The census calls the napi arm
 * directly through the shim's createCanvas.
 */
function measureWithNapi(font: string): number {
  const ctx = createCanvas(128, 64).getContext('2d') as CanvasRenderingContext2D
  ctx.font = font
  return ctx.measureText(PROBE_STRING).width
}

function measureWithNode(font: string): number {
  // Bun preload rewrites 'canvas' → the napi shim in this process, so the
  // node-canvas arm is measured in a child node process (execFileSync,
  // round-10 F2) is handled by the execFileSync child node run below.
  const expr = [
    `const PROBE = ${JSON.stringify(PROBE_STRING)}`,
    `const FONT = ${JSON.stringify(font)}`,
    'const { createCanvas } = await import("canvas")',
    'const ctx = createCanvas(128, 64).getContext("2d")',
    'ctx.font = FONT',
    'console.log(JSON.stringify({ ok: true, width: ctx.measureText(PROBE).width }))',
  ].join('\n')
  const raw = execFileSync('node', ['--input-type=module', "--eval", expr], { encoding: 'utf8', timeout: 10_000 })
  const last = raw.trim().split(/\r?\n/).at(-1) ?? '{}'
  const parsed = JSON.parse(last) as { ok?: boolean; width?: number; error?: string }
  if (!parsed.ok || typeof parsed.width !== 'number') {
    throw new Error(parsed.error ?? 'node-canvas measurement failed')
  }
  return parsed.width
}

const round2 = (n: number): number => Math.round(n * 100) / 100
const relDiff = (a: number, b: number): number => {
  const denom = Math.max(Math.abs(a), Math.abs(b), 1)
  return Math.round((Math.abs(a - b) / denom) * 10_000) / 10_000
}

function round9AggregateFonts(into: CensusReport, from: CensusReport): void {
  if (Array.isArray(from.fontStrings)) {
    for (const f of from.fontStrings) {
      if (typeof f === 'string' && !(into.fontStrings ?? []).includes(f)) (into.fontStrings ??= []).push(f)
    }
  }
}

function mergeCensus(into: CensusReport, from: CensusReport): void {
  for (const [key, count] of Object.entries(from.methods)) into.methods[key] = (into.methods[key] ?? 0) + count
  for (const [name, slot] of Object.entries(from.props)) {
    const to = into.props[name] ?? (into.props[name] = { count: 0, values: [] })
    to.count += slot.count
  }
  for (const key of from.notInSubset) {
    if (!into.notInSubset.includes(key)) into.notInSubset.push(key)
  }
  for (const rec of from.recordText) {
    into.recordText.push({ text: rec.text, scope: rec.scope, start: rec.start, end: rec.end })
  }
  round9AggregateFonts(into, from)
  into.classEvidence.shadowBlurSeen ||= from.classEvidence.shadowBlurSeen
  into.classEvidence.setTransformSeen ||= from.classEvidence.setTransformSeen
  into.classEvidence.createPatternSeen ||= from.classEvidence.createPatternSeen
  into.classEvidence.drawImageCanvasSourceSeen ||= from.classEvidence.drawImageCanvasSourceSeen
}

/** Round-9 face resolution — every distinct `ctx.font` the census saw, measured
 * under BOTH backends and recorded with the relative difference. */
function collectFontResolutions(values: Array<string> | undefined): FaceResolutionRow[] {
  const fonts = new Set<string>()
  // census values are `<valueTag>:<raw>`; font values are `string:<raw>` entries
  for (const entry of values ?? []) {
    if (!entry.startsWith('string:')) continue
    fonts.add(entry.slice('string:'.length))
  }
  const out: FaceResolutionRow[] = []
  for (const font of fonts) {
    try {
      const widthNode = measureWithNode(font)
      const widthNapi = measureWithNapi(font)
      out.push({ font, widthNode: round2(widthNode), widthNapi: round2(widthNapi), relDiff: relDiff(widthNode, widthNapi) })
    } catch (error) {
      console.warn(`[t3] face resolution skipped for ${font}: ${(error as Error).message}`)
    }
  }
  return out
}

async function run(): Promise<number> {
  // Run me with the napi preload: `bun run s0:napi scripts/stream-s0/t3-census.ts`
  // (the bun preload maps `canvas` to the napi shim; running without it
  // measures the node-canvas backend, which the report records via
  // `resolveBackend`).
  const corpus = listCorpus()
  const aggregate: CensusReport = {
    methods: {}, props: {}, notInSubset: [], faceResolution: [], recordText: [],
    classEvidence: { shadowBlurSeen: false, setTransformSeen: false, createPatternSeen: false, drawImageCanvasSourceSeen: false },
    fontStrings: [],
  }
  const unitsByClass: Array<{ file: string; unit: number; classes: ReadonlyArray<string> }> = []
  let unitsSeen = 0
  const unitsFailed: Array<{ file: string; error: string }> = []

  for (const file of corpus as readonly CorpusFile[]) {
    let lease: ReturnType<typeof loadUnits> | undefined
    try {
      lease = await loadUnits(file)
      for (const unit of (lease as readonly Unit[])) {
        unitsSeen++
        const classes: string[] = []
        const real = createCtxFor(unit)
        const { ctx, report, classEvidence } = wrapCensus(real)
        try {
          unit.paint(ctx)
        } finally {
          mergeCensus(aggregate, report())
        }
        const evidence = classEvidence()
        if (evidence.shadowBlurSeen) classes.push('shadowBlur>0')
        if (evidence.setTransformSeen) classes.push('setTransform')
        if (evidence.createPatternSeen) classes.push('createPattern')
        if (evidence.drawImageCanvasSourceSeen) classes.push('drawImage/source=canvas')
        if (process.env.S0_PROBE === '1') {
          console.error('[t3-class]', file.name, unit.index, JSON.stringify(evidence))
        }
        if (classes.length > 0) unitsByClass.push({ file: file.name, unit: unit.index, classes })
      }
    } catch (error) {
      unitsFailed.push({ file: file.name, error: (error as Error).message })
    } finally {
      // Ownership: the lease's `dispose` lives on the returned array as a
      // non-enumerable property (round-10 fix 5) — callers never key on
      // array identity, so copies can't silently leak.
      // Ownership: dispose lives on the returned array as a non-enumerable
      // property (round-10 fix 5), so callers can't leak on array identity.
      if (lease !== undefined) lease.dispose()
    }
  }

  const faceResolution: FaceResolutionRow[] = []
  try {
    const fontsForRes = (aggregate as unknown as { fontStrings?: string[] }).fontStrings ?? []
    for (const font of fontsForRes) {
      // The census's own `font` collection (round-9 requirement): measure the
      // same string under BOTH backends in the same run.
      try {
        const widthNode = measureWithNode(font)
        const widthNapi = measureWithNapi(font)
        faceResolution.push({ font, widthNode: round2(widthNode), widthNapi: round2(widthNapi), relDiff: relDiff(widthNode, widthNapi) })
      } catch (error) {
        console.warn(`[t3] face resolution skipped for ${font}: ${(error as Error).message}`)
      }
    }
  } catch (error) {
    console.warn(`[t3] face resolution failed: ${(error as Error).message}`)
  }

  const failedOnly = unitsFailed.length > 0
  // Plan Task 3: GREEN iff notInSubset is empty. A file that fails to load
  // (e.g. the deliberately corrupt poi-crash-* fuzzer fixtures) is listed in
  // unitsFailed for the report; whether a failure is napi-only is T1's gate.
  void failedOnly
  const status = aggregate.notInSubset.length > 0 ? 'RED' : 'GREEN'
  const backend = resolveBackend()
  writeResult('t3', {
    status,
    backend,
    zoom: ZOOM,
    dpr: DPR,
    corpusFiles: corpus.length,
    unitsSeen,
    unitsFailed,
    methods: aggregate.methods,
    props: aggregate.props,
    notInSubset: aggregate.notInSubset,
    faceResolution,
    recordText: aggregate.recordText,
    classEvidence: aggregate.classEvidence,
    unitsByClass,
    fontStrings: aggregate.fontStrings ?? [],
    note: status === 'RED'
      ? (aggregate.notInSubset.length > 0
        ? `notInSubset is a RED §7.3/§7.5 gap for the owner — do not fix by editing the subset`
        : 'unreachable')
      : `census agrees with §7.3/§7.5 (recorded subset covers everything the corpus paints); ${unitsFailed.length} file(s) failed to load and are listed in unitsFailed (T1 decides whether any failure is napi-only)`,
  })
  console.log(
    `[t3] status ${status} · backend ${backend} · files ${corpus.length}`,
    `· units ${unitsSeen} · failed ${unitsFailed.length}`,
    `· notInSubset ${aggregate.notInSubset.length}`,
    `· faceRows ${faceResolution.length}`,
    `· recordText ${aggregate.recordText.length}`
  )
  return status === 'RED' ? 3 : 0
}

function createCtxFor(unit: Unit): CanvasRenderingContext2D {
  return createCanvas(
    Math.max(1, Math.ceil(unit.widthPx * ZOOM * DPR)),
    Math.max(1, Math.ceil(unit.heightPx * ZOOM * DPR)),
  ).getContext('2d') as unknown as CanvasRenderingContext2D
}


run().then(
  (code) => process.exit(code),
  (error) => { console.error(error); process.exit(1) },
)
