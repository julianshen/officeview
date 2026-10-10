/** Finite B2 conditional handlers; no public ownership/array ABI additions. */
import type { FunctionHandler, EvaluatorFn } from '../functions'
import type { AstNode, ElementValue, EvaluationContext, EvaluationValue, EvaluationError, MatrixValue, ResolvedRef, RefArea, RefCursor, RefEntry } from '../types'
import { coerceToBoolean, coerceToNumber, finiteOrNum, formulaError, isEvaluationError, isMatrixValue, markUnsupported, matrixOf } from '../evaluator'
import { matchWildcard } from './wildcard'
import { currentReferenceGeneration, isReferenceGenerationChanged, referenceGenerationChanged, ownedMemoEpoch, requestOwnedMemoRestart, MAX_REFERENCE_GENERATION_RESTARTS, createReferenceRegion } from '../reference-generation'
import { isReferenceNode } from '../refs'

export type ConditionalOrigin = 'direct' | 'omitted' | 'reference' | 'array'
type PredicateName = 'ISBLANK' | 'ISNUMBER' | 'ISTEXT' | 'ISLOGICAL' | 'ISERROR' | 'ISERR' | 'ISNA' | 'TYPE' | 'ERROR.TYPE' | 'N'

const capabilityErrors = new WeakSet<object>()
function unavailable(ctx: EvaluationContext | undefined, feature: string, message: string): EvaluationError {
  markUnsupported(ctx, feature)
  ctx?.reportFormulaIssue?.({ kind: 'unsupported-reference', feature, message })
  const error = formulaError('#NAME?')
  capabilityErrors.add(error)
  return error
}
/** Broadcast index into a matrix, or undefined when the position is not
 * covered. Undefined is a shape gate, never a padded null/zero. */
function broadcastAt(matrix: MatrixValue, row: number, col: number): ElementValue | undefined {
  const r = matrix.rows === 1 ? 0 : row
  const c = matrix.cols === 1 ? 0 : col
  if (r >= matrix.rows || c >= matrix.cols) return undefined
  return matrix.values[r][c]
}

function matrixInput(value: unknown): boolean {
  if (typeof value !== 'object' || value === null || !('kind' in value) || value.kind !== 'formula-matrix') return false
  if (!('rows' in value) || !('cols' in value) || !('values' in value)) return false
  return typeof value.rows === 'number' && Number.isSafeInteger(value.rows) && value.rows > 0 && typeof value.cols === 'number' && Number.isSafeInteger(value.cols) && value.cols > 0 && Array.isArray(value.values) && value.values.length === value.rows && value.values.every(row => Array.isArray(row) && row.length === value.cols)
}
const ERROR_NUMBERS: Readonly<Record<string, number>> = { '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7 }

/** Private unknown-value adapter: TYPE matrix64 is measured. Literal grammar
 * and elementwise array publication are C2; array input is never flattened. */
export function conditionalPredicate(name: PredicateName, value: unknown, origin: ConditionalOrigin, ctx?: EvaluationContext): EvaluationValue {
  if (isEvaluationError(value) && capabilityErrors.has(value)) return value
  if (matrixInput(value)) {
    // Measured: TYPE(matrix) is the scalar 64. Every other known predicate is
    // documented elementwise (Microsoft array guidelines), so each genuine
    // element is classified independently — error-looking text stays text.
    if (name === 'TYPE') return 64
    const matrix = value as MatrixValue
    const values: ElementValue[][] = []
    for (let row = 0; row < matrix.rows; row++) {
      const out: ElementValue[] = []
      for (let col = 0; col < matrix.cols; col++) {
        const element = conditionalScalar(name, matrix.values[row][col], 'reference', ctx)
        if (isEvaluationError(element) && capabilityErrors.has(element)) return element
        out.push(element)
      }
      values.push(out)
    }
    return { kind: 'formula-matrix', rows: matrix.rows, cols: matrix.cols, values }
  }
  return conditionalScalar(name, value, origin, ctx)
}

function conditionalScalar(name: PredicateName, value: unknown, origin: ConditionalOrigin, ctx?: EvaluationContext): ElementValue {
  if (origin === 'omitted') return unavailable(ctx, 'conditional-omitted-predicate-unverified', 'Omitted predicate input has no accepted function-specific native observation')
  switch (name) {
    case 'ISBLANK': return origin === 'reference' && value === null
    case 'ISNUMBER': return typeof value === 'number'
    case 'ISTEXT': return typeof value === 'string'
    case 'ISLOGICAL': return typeof value === 'boolean'
    case 'ISERROR': return isEvaluationError(value)
    case 'ISERR': return isEvaluationError(value) && value.code !== '#N/A'
    case 'ISNA': return isEvaluationError(value) && value.code === '#N/A'
    case 'TYPE':
      if (isEvaluationError(value)) return 16
      if (typeof value === 'boolean') return 4
      if (typeof value === 'string') return 2
      if (typeof value === 'number' || value === null) return 1
      return unavailable(ctx, 'conditional-value-unavailable', 'TYPE input is outside the finite scalar/matrix value contract')
    case 'ERROR.TYPE': {
      if (!isEvaluationError(value)) return formulaError('#N/A')
      const number = ERROR_NUMBERS[value.code]
      return number ?? unavailable(ctx, 'conditional-modern-error-unverified', 'Modern ERROR.TYPE mapping has no accepted native observation')
    }
    case 'N':
      if (isEvaluationError(value)) return value
      if (typeof value === 'number') return value
      if (typeof value === 'boolean') return value ? 1 : 0
      if (typeof value === 'string' || value === null) return 0
      return unavailable(ctx, 'conditional-value-unavailable', 'N input is outside the finite scalar value contract')
  }
}
function originOf(arg: AstNode): ConditionalOrigin { return arg.type === 'empty' ? 'omitted' : isReferenceNode(arg) ? 'reference' : 'direct' }
function predicate(name: PredicateName): FunctionHandler {
  return (args, ctx, evalNode) => args.length !== 1 ? formulaError('#VALUE!') : conditionalPredicate(name, args[0].type === 'empty' ? null : evalNode(args[0], ctx), originOf(args[0]), ctx)
}
function selected(arg: AstNode, ctx: EvaluationContext | undefined, evalNode: EvaluatorFn): EvaluationValue {
  if (arg.type === 'empty') return 0
  const value = evalNode(arg, ctx)
  return value === null ? 0 : value
}

type AggregateName = 'SUMIF' | 'SUMIFS' | 'COUNTIF' | 'COUNTIFS' | 'AVERAGEIF' | 'AVERAGEIFS'
type Relation = '=' | '<>' | '<' | '>' | '<=' | '>='
type Criterion = { kind: 'number'; number: number; op: Relation } | { kind: 'blank' } | { kind: 'text'; pattern: string; op: '=' | '<>' } | { kind: 'error'; code: '#N/A' }
interface View { ref: ResolvedRef; area: RefArea; prepared: boolean; absent?: number }
interface Pair { rangeArg: AstNode; criterionArg: AstNode; view?: View; criterion?: Criterion }
interface Stream { view: View; cursor: RefCursor; next?: RefEntry; loaded: boolean; done: boolean; criterionIndex?: number; target: boolean }
interface CurrentRow { position: number; values: Array<ElementValue | undefined>; loaded: boolean[]; nextCriterion: number; matches: boolean; target?: ElementValue; targetLoaded: boolean }
const AGGREGATE = Symbol('conditional aggregate')
interface AggregateState {
  tag: typeof AGGREGATE
  family: AggregateName
  pairs: Pair[]
  targetArg?: AstNode
  target?: View
  initialized: boolean
  streams: Stream[]
  row?: CurrentRow
  total: number
  count: number
  baseline: boolean
  generation?: number
  owned: ReturnType<typeof ownedMemoEpoch>
  restarts: number
  result?: EvaluationValue
  complete: boolean
}
const fallbackWork = new WeakMap<EvaluationContext, WeakMap<AstNode[], AggregateState>>()
function isState(value: unknown): value is AggregateState { return typeof value === 'object' && value !== null && 'tag' in value && value.tag === AGGREGATE }
function storeState(args: AstNode[], ctx: EvaluationContext, state: AggregateState): void {
  if (ctx.functionWork) ctx.functionWork.set(args, state)
  else {
    let work = fallbackWork.get(ctx)
    if (!work) { work = new WeakMap(); fallbackWork.set(ctx, work) }
    work.set(args, state)
  }
}
function stateOf(args: AstNode[], ctx: EvaluationContext): AggregateState | undefined {
  const value = ctx.functionWork ? ctx.functionWork.get(args) : fallbackWork.get(ctx)?.get(args)
  return isState(value) ? value : undefined
}
function newState(family: AggregateName, args: AstNode[], ctx: EvaluationContext, restarts = 0): AggregateState | EvaluationError {
  const pairs: Pair[] = []
  let targetArg: AstNode | undefined
  if (family === 'COUNTIF' || family === 'SUMIF' || family === 'AVERAGEIF') {
    const max = family === 'COUNTIF' ? 2 : 3
    if (args.length < 2 || args.length > max) return formulaError('#VALUE!')
    pairs.push({ rangeArg: args[0], criterionArg: args[1] })
    if (args.length === 3) targetArg = args[2]
  } else if (family === 'COUNTIFS') {
    if (args.length < 2 || args.length % 2 !== 0 || args.length > 254) return formulaError('#VALUE!')
    for (let i = 0; i < args.length; i += 2) pairs.push({ rangeArg: args[i], criterionArg: args[i + 1] })
  } else {
    if (args.length < 3 || args.length % 2 !== 1 || args.length > 255) return formulaError('#VALUE!')
    targetArg = args[0]
    for (let i = 1; i < args.length; i += 2) pairs.push({ rangeArg: args[i], criterionArg: args[i + 1] })
  }
  return { tag: AGGREGATE, family, pairs, targetArg, initialized: false, streams: [], total: 0, count: 0, baseline: false, generation: ctx.references ? currentReferenceGeneration(ctx.references) : undefined, owned: ownedMemoEpoch(ctx), restarts, complete: false }
}
function viewOf(arg: AstNode, ctx: EvaluationContext): View | EvaluationError {
  if (!ctx.references) return unavailable(ctx, 'conditional-reference-unavailable', 'Conditional ranges require sparse reference services; legacy clamped ranges are not full geometry')
  if (!isReferenceNode(arg)) return formulaError('#VALUE!')
  const ref = ctx.references.resolve(arg, ctx)
  if (ref === undefined) return formulaError('#VALUE!')
  if (isEvaluationError(ref)) return ref
  if (ref.areas.length !== 1) return unavailable(ctx, 'conditional-reference-shape-unverified', 'Conditional multi-area/3D range pairing has no accepted function-specific profile')
  return { ref, area: ref.areas[0], prepared: false }
}
function criterionOf(arg: AstNode, family: AggregateName, ctx: EvaluationContext, evalNode: EvaluatorFn): Criterion | EvaluationError {
  if (arg.type === 'empty') return unavailable(ctx, 'conditional-omitted-criterion-unverified', 'Explicitly omitted criteria remain separate from empty text and blank references')
  let value = evalNode(arg, ctx)
  if (isEvaluationError(value)) return value
  if (value === null) {
    if (isReferenceNode(arg) && (family === 'COUNTIFS' || family === 'AVERAGEIF')) value = 0
    else return unavailable(ctx, 'conditional-blank-criterion-unverified', 'This function has no accepted blank criterion-input profile')
  }
  if (typeof value === 'boolean') return unavailable(ctx, 'conditional-boolean-criterion-unverified', 'Boolean criteria require a function-specific native profile')
  if (typeof value === 'number') return Number.isFinite(value) ? { kind: 'number', number: value, op: '=' } : formulaError('#VALUE!')
  if (typeof value !== 'string') return unavailable(ctx, 'conditional-criteria-unavailable', 'Criteria must be a finite scalar input; array criteria belong to the later array runtime')
  if (value === '') return { kind: 'blank' }
  if (value.length > 255) return unavailable(ctx, 'conditional-long-criterion-unverified', 'Criteria beyond255 characters have documented function limits but no accepted outcome for this profile')
  const prefix = /^(<=|>=|<>|=|<|>)/.exec(value)
  const op: Relation = prefix ? prefix[1] as Relation : '='
  const body = prefix ? value.slice(prefix[0].length) : value
  if (prefix && /^[<>=]/.test(body)) return unavailable(ctx, 'conditional-criteria-unverified', 'Repeated/malformed comparison prefixes have no accepted native interpretation')
  if (body === '') return unavailable(ctx, 'conditional-empty-comparison-unverified', 'Empty comparison operand is not silently conflated with literal empty text')
  if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(body)) {
    const number = Number(body)
    return Number.isFinite(number) ? { kind: 'number', number, op } : unavailable(ctx, 'conditional-numeric-criterion-unverified', 'Nonfinite numeric criteria have no accepted outcome')
  }
  if (body === '#N/A' && op === '=') return family === 'COUNTIF' ? { kind: 'error', code: '#N/A' } : unavailable(ctx, 'conditional-error-criterion-unverified', 'Error-string criterion matching is measured only for this COUNTIF profile')
  if (op !== '=' && op !== '<>') return unavailable(ctx, 'conditional-text-comparison-unverified', 'Text ordering is outside the documented equality/negative ASCII text profile')
  const check = matchWildcard(body, '')
  if (check.status !== 'matched' && check.status !== 'not-matched') return unavailable(ctx, `conditional-wildcard-${check.status}`, check.reason)
  return { kind: 'text', pattern: body, op }
}
function relation(left: number, right: number, op: Relation): boolean {
  switch (op) { case '=': return left === right; case '<>': return left !== right; case '<': return left < right; case '>': return left > right; case '<=': return left <= right; case '>=': return left >= right }
}
function matches(value: ElementValue, criterion: Criterion, family: AggregateName, ctx: EvaluationContext): boolean | EvaluationError {
  if (criterion.kind === 'blank') return value === null || value === ''
  if (criterion.kind === 'error') {
    if (isEvaluationError(value)) return value.code === criterion.code
    if (typeof value === 'string' && value === criterion.code) return unavailable(ctx, 'conditional-error-text-unverified', 'Error-looking text is not converted into the measured genuine-error criterion profile')
    return false
  }
  if (criterion.kind === 'text') {
    if (criterion.op === '<>' && (typeof value !== 'string' || value === '')) return unavailable(ctx, 'conditional-negative-text-input-unverified', 'Negative text criteria on missing, empty, or nontext cells have no accepted function-specific profile')
    if (typeof value !== 'string') return false
    const result = matchWildcard(criterion.pattern, value)
    if (result.status === 'matched') return criterion.op === '='
    if (result.status === 'not-matched') return criterion.op === '<>'
    return unavailable(ctx, `conditional-wildcard-${result.status}`, result.reason)
  }
  if (value === null) {
    if (family === 'AVERAGEIFS') return relation(0, criterion.number, criterion.op)
    if (criterion.op === '<>') return unavailable(ctx, 'conditional-negative-blank-unverified', 'Negative numeric criteria on structural blanks have no accepted function-specific profile')
    return false
  }
  if (typeof value === 'number') return relation(value, criterion.number, criterion.op)
  if (typeof value === 'boolean') {
    if (family === 'AVERAGEIF') return false // primary documented TRUE/FALSE cells in its range ignored
    if (relation(value ? 1 : 0, criterion.number, criterion.op)) return unavailable(ctx, 'conditional-boolean-range-unverified', 'Matching boolean/numeric range coercion has no accepted function-specific profile')
    return false
  }
  if (typeof value === 'string') {
    if (value !== '') {
      const numeric = coerceToNumber(value)
      if (!isEvaluationError(numeric) && relation(numeric, criterion.number, criterion.op)) return unavailable(ctx, 'conditional-numeric-text-unverified', 'Stored numeric text is not silently coerced to a numeric criteria match')
    }
    if (criterion.op === '<>') return unavailable(ctx, 'conditional-negative-text-unverified', 'Negative numeric criteria against text has no accepted profile')
    return false
  }
  if (criterion.op === '<>') return unavailable(ctx, 'conditional-negative-error-unverified', 'Negative criteria error inclusion has no accepted profile')
  return false
}
function sameShape(a: RefArea, b: RefArea): boolean { return a.cols === b.cols && a.rows === b.rows }
function sameCells(a: RefArea, b: RefArea): boolean { return sameShape(a, b) && a.sheetId === b.sheetId && a.firstCol === b.firstCol && a.firstRow === b.firstRow }
function resizeTarget(view: View, area: RefArea, ctx: EvaluationContext): View | EvaluationError {
  if (sameShape(view.area, area)) return view
  const region = { ...view.area, rows: area.rows, cols: area.cols }
  if (![region.firstCol, region.firstRow, region.cols, region.rows].every(Number.isInteger) || region.firstCol < 0 || region.firstRow < 0 || region.cols < 0 || region.rows < 0 || region.firstCol + region.cols > 16384 || region.firstRow + region.rows > 1048576) return formulaError('#REF!')
  const ref = createReferenceRegion(view.ref, [region])
  if (!ref) return unavailable(ctx, 'conditional-target-resize-unavailable', 'Custom target reference has no private generation-bound region adapter')
  return { ref, area: region, prepared: false }
}
function initialize(state: AggregateState, ctx: EvaluationContext, evalNode: EvaluatorFn): EvaluationError | undefined {
  for (const pair of state.pairs) {
    if (!pair.view) { const view = viewOf(pair.rangeArg, ctx); if (isEvaluationError(view)) return view; pair.view = view }
    if (!pair.criterion) { const criterion = criterionOf(pair.criterionArg, state.family, ctx, evalNode); if (isEvaluationError(criterion)) return criterion; pair.criterion = criterion }
  }
  const first = state.pairs[0].view!
  if (!state.target && !state.family.startsWith('COUNT')) {
    if (state.targetArg?.type === 'empty') return unavailable(ctx, 'conditional-omitted-target-unverified', 'Explicitly omitted target input is not assumed to mean absent optional target')
    if (!state.targetArg) state.target = first
    else {
      const view = viewOf(state.targetArg, ctx)
      if (isEvaluationError(view)) return view
      if (state.family === 'SUMIF' || state.family === 'AVERAGEIF') {
        const resized = resizeTarget(view, first.area, ctx)
        if (isEvaluationError(resized)) return resized
        state.target = resized
      } else state.target = view
    }
  }
  const shape = state.target?.area ?? first.area
  for (const pair of state.pairs) if (!sameShape(pair.view!.area, shape)) return formulaError('#VALUE!')
  const views = state.pairs.map(pair => pair.view!)
  if (state.target && !views.includes(state.target)) views.push(state.target)
  for (const view of views) {
    if (!view.prepared) { ctx.references!.prepare(view.ref, ctx); view.prepared = true }
    if (view.absent === undefined) view.absent = view.ref.countAbsent()
  }
  let restriction = -1
  for (let i = 0; i < state.pairs.length; i++) {
    const pair = state.pairs[i]
    const missing = pair.view!.absent! > 0 ? matches(null, pair.criterion!, state.family, ctx) : false
    if (isEvaluationError(missing)) return missing
    if (!missing && restriction === -1) restriction = i
  }
  const stream = (view: View, criterionIndex?: number, target = false): Stream => ({ view, cursor: view.ref.openCursor('populated'), loaded: false, done: false, criterionIndex, target })
  if (restriction !== -1) state.streams = [stream(state.pairs[restriction].view!, restriction)]
  else if (state.target) state.streams = [stream(state.target, undefined, true)]
  else {
    state.baseline = true
    state.total = first.absent!
    state.streams = state.pairs.map((pair, i) => stream(pair.view!, i))
  }
  state.initialized = true
  return undefined
}
function position(entry: RefEntry, cols: number): number { return entry.offsetRow * cols + entry.offsetCol }
function runRows(state: AggregateState, ctx: EvaluationContext): EvaluationValue {
  const cols = state.pairs[0].view!.area.cols
  for (;;) {
    if (!state.row) {
      for (const stream of state.streams) {
        if (!stream.loaded && !stream.done) { stream.next = stream.view.ref.peek(stream.cursor); stream.loaded = true; if (!stream.next) stream.done = true }
      }
      let next: number | undefined
      for (const stream of state.streams) if (stream.next) { const p = position(stream.next, cols); if (next === undefined || p < next) next = p }
      if (next === undefined) break
      const row: CurrentRow = { position: next, values: [], loaded: [], nextCriterion: 0, matches: true, targetLoaded: false }
      for (const stream of state.streams) if (stream.next && position(stream.next, cols) === next) {
        if (stream.criterionIndex !== undefined) { row.values[stream.criterionIndex] = stream.next.value; row.loaded[stream.criterionIndex] = true }
        if (stream.target) { row.target = stream.next.value; row.targetLoaded = true }
      }
      state.row = row
    }
    const row = state.row
    const offsetRow = Math.floor(row.position / cols), offsetCol = row.position % cols
    while (row.matches && row.nextCriterion < state.pairs.length) {
      const i = row.nextCriterion, pair = state.pairs[i]
      if (!row.loaded[i]) { row.values[i] = pair.view!.ref.readAt(0, offsetRow, offsetCol); row.loaded[i] = true }
      const hit = matches(row.values[i]!, pair.criterion!, state.family, ctx)
      if (isEvaluationError(hit)) return hit
      row.matches = hit
      row.nextCriterion++
    }
    let add = 0, count = 0
    if (state.target && row.matches) {
      if (!row.targetLoaded) {
        const source = state.pairs.findIndex((pair, i) => sameCells(pair.view!.area, state.target!.area) && row.loaded[i])
        if (source >= 0) row.target = row.values[source]
        else row.target = state.target.ref.readAt(0, offsetRow, offsetCol)
        row.targetLoaded = true
      }
      const value = row.target!
      if (isEvaluationError(value)) return unavailable(ctx, 'conditional-selected-target-error-unverified', 'Selected target-error propagation requires an accepted function-specific native profile')
      if (typeof value === 'number') { add = value; count = 1 }
      else if (typeof value === 'boolean') {
        if (state.family === 'SUMIFS') { add = value ? 1 : 0; count = 1 }
        else return unavailable(ctx, 'conditional-boolean-target-unverified', 'This function has no settled boolean target coercion profile')
      }
    } else if (!state.target) {
      add = row.matches ? 1 : 0
      if (state.baseline) {
        const firstStored = state.streams.some(stream => stream.criterionIndex === 0 && stream.next && position(stream.next, cols) === row.position)
        if (!firstStored) add-- // the first structural-blank baseline already included this position
      }
    }
    // All required cell/criterion reads completed before cursor/accumulator commit.
    for (const stream of state.streams) if (stream.next && position(stream.next, cols) === row.position) { stream.view.ref.advance(stream.cursor); stream.next = undefined; stream.loaded = false }
    state.total += add
    state.count += count
    state.row = undefined
  }
  return state.family.startsWith('AVERAGE') ? state.count === 0 ? formulaError('#DIV/0!') : finiteOrNum(state.total / state.count) : finiteOrNum(state.total)
}
function aggregate(family: AggregateName): FunctionHandler {
  return (args, ctx, evalNode) => {
    if (!ctx?.references) return unavailable(ctx, 'conditional-reference-unavailable', 'Conditional aggregates need full sparse reference geometry')
    let state = stateOf(args, ctx)
    if (!state || state.family !== family) { const created = newState(family, args, ctx); if (isEvaluationError(created)) return created; state = created; storeState(args, ctx, state) }
    for (;;) {
      try {
        const owned = ownedMemoEpoch(ctx)
        if (owned && (state.owned?.owner !== owned.owner || state.owned.epoch !== owned.epoch)) {
          if (owned.epoch > 0) { const fresh = newState(family, args, ctx, state.restarts); if (isEvaluationError(fresh)) return fresh; state = fresh; storeState(args, ctx, state) }
          else state.owned = owned
        }
        const current = currentReferenceGeneration(ctx.references)
        if (state.generation !== undefined && current !== undefined && state.generation !== current) throw referenceGenerationChanged(state.generation, current)
        if (state.complete) return state.result!
        if (!state.initialized) { const error = initialize(state, ctx, evalNode); if (error) return error }
        const result = runRows(state, ctx)
        const after = currentReferenceGeneration(ctx.references)
        if (state.generation !== undefined && after !== undefined && state.generation !== after) throw referenceGenerationChanged(state.generation, after)
        state.complete = true
        state.result = result
        if (!ctx.functionWork) fallbackWork.get(ctx)?.delete(args)
        return result
      } catch (error) {
        if (!isReferenceGenerationChanged(error)) throw error
        const restarts = state.restarts + 1
        const fresh = newState(family, args, ctx, restarts)
        if (isEvaluationError(fresh)) return fresh
        state = fresh
        storeState(args, ctx, state)
        if (restarts > MAX_REFERENCE_GENERATION_RESTARTS) return unavailable(ctx, 'reference-generation-unstable', `Conditional footprint changed beyond the ${MAX_REFERENCE_GENERATION_RESTARTS}-restart application budget`)
        requestOwnedMemoRestart(ctx)
      }
    }
  }
}

export const CONDITIONAL_FUNCTIONS: Readonly<Record<string, FunctionHandler>> = {
  SUMIF: aggregate('SUMIF'), SUMIFS: aggregate('SUMIFS'), COUNTIF: aggregate('COUNTIF'), COUNTIFS: aggregate('COUNTIFS'), AVERAGEIF: aggregate('AVERAGEIF'), AVERAGEIFS: aggregate('AVERAGEIFS'),
  IFNA: (args, ctx, evalNode) => {
    if (args.length !== 2) return formulaError('#VALUE!')
    if (args[0].type === 'empty') return unavailable(ctx, 'conditional-omitted-ifna-unverified', 'Omitted IFNA input remains distinct from a documented empty-cell input')
    const value = evalNode(args[0], ctx)
    if (isEvaluationError(value) && value.code === '#N/A') {
      if (args[1].type === 'empty') return unavailable(ctx, 'conditional-omitted-ifna-unverified', 'Omitted IFNA fallback has no accepted native observation')
      const fallback = evalNode(args[1], ctx)
      return fallback === null && isReferenceNode(args[1]) ? '' : fallback
    }
    // Documented IFNA array form (remarks143): replace only genuine #N/A cells;
    // other errors and error-looking text are preserved per cell.
    if (isMatrixValue(value)) {
      const hasNA = value.values.some((row) => row.some((element) => isEvaluationError(element) && element.code === '#N/A'))
      if (!hasNA) return value
      if (args[1].type === 'empty') return unavailable(ctx, 'conditional-omitted-ifna-unverified', 'Omitted IFNA fallback has no accepted native observation')
      const fallback = evalNode(args[1], ctx)
      const fallbackMatrix = isMatrixValue(fallback) ? fallback : undefined
      const fallbackScalar = fallbackMatrix ? undefined : (fallback === null && isReferenceNode(args[1]) ? '' : fallback)
      let gated = false
      const values = value.values.map((row, r) => row.map((element, c): ElementValue => {
        if (!(isEvaluationError(element) && element.code === '#N/A')) return element
        if (fallbackMatrix) {
          const picked = broadcastAt(fallbackMatrix, r, c)
          if (picked === undefined) { gated = true; return null }
          return picked
        }
        return fallbackScalar as ElementValue
      }))
      if (gated) return unavailable(ctx, 'array-shape-mismatch-unverified', 'IFNA fallback array dimensions do not cover the selected NA cells; no null/zero padding is invented')
      return matrixOf(value.rows, value.cols, values)
    }
    return value === null && isReferenceNode(args[0]) ? '' : value
  },
  IFS: (args, ctx, evalNode) => {
    if (args.length < 2 || args.length % 2 !== 0 || args.length > 254) return formulaError('#VALUE!')
    for (let i = 0; i < args.length; i += 2) {
      const value = evalNode(args[i], ctx)
      if (isEvaluationError(value)) return value
      const test = coerceToBoolean(value)
      if (isEvaluationError(test)) return test
      if (test) return selected(args[i + 1], ctx, evalNode)
    }
    return formulaError('#N/A')
  },
  SWITCH: (args, ctx, evalNode) => {
    if (args.length < 3 || args.length > 254) return formulaError('#VALUE!')
    const value = evalNode(args[0], ctx)
    if (isEvaluationError(value)) return value
    const hasDefault = args.length % 2 === 0
    const end = hasDefault ? args.length - 1 : args.length
    for (let i = 1; i < end; i += 2) {
      const candidate = evalNode(args[i], ctx)
      if (isEvaluationError(candidate)) return candidate
      // Exact type equality; unverified cross-type comparison is not coerced.
      if (typeof candidate === typeof value && candidate === value) return selected(args[i + 1], ctx, evalNode)
    }
    return hasDefault ? selected(args[args.length - 1], ctx, evalNode) : formulaError('#N/A')
  },
  ISBLANK: predicate('ISBLANK'), ISNUMBER: predicate('ISNUMBER'), ISTEXT: predicate('ISTEXT'), ISLOGICAL: predicate('ISLOGICAL'), ISERROR: predicate('ISERROR'), ISERR: predicate('ISERR'), ISNA: predicate('ISNA'), TYPE: predicate('TYPE'), 'ERROR.TYPE': predicate('ERROR.TYPE'), N: predicate('N'),
}
