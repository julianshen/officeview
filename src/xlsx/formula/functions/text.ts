/**
 * functions/text.ts — TEXT(value, format) over the one shared finite common
 * formatter (format.ts), the same formatter the numeric cell renderer uses.
 *
 * Unsupported finite-grammar classes are a CAPABILITY gate (frame mark +
 * specific diagnostic + internal #NAME?), never a fabricated Excel error.
 */
import type { FunctionHandler } from '../functions'
import type { EvaluationContext, EvaluationError } from '../types'
import { formulaError, isEvaluationError, isMatrixValue, markUnsupported } from '../evaluator'
import { formatCommon } from '../format'
import { resolveSemantics } from './datetime'

/**
 * Honest gate for an UNVERIFIED matrix origin on either TEXT slot: frame mark +
 * attributable diagnostic + internal #NAME? gate shape. Raised at the FIRST
 * slot that settles the capability so a later argument is never read once the
 * unsupported profile is already known (one-needed-child rule).
 */
function textMatrixGate(ctx: EvaluationContext | undefined): EvaluationError {
  markUnsupported(ctx, 'text-array-input-unverified')
  ctx?.reportFormulaIssue?.({
    kind: 'native-gate',
    gate: 'text-array-input-unverified',
    message: 'TEXT array-input profile is not verified; scalar TEXT only',
  })
  return formulaError('#NAME?')
}

export const textHandler: FunctionHandler = (args, ctx, evalNode) => {
  if (args.length !== 2) return formulaError('#VALUE!')
  const value = evalNode(args[0], ctx)
  if (isEvaluationError(value)) return value
  // Matrix VALUE settles the capability before the format argument is read.
  if (isMatrixValue(value)) return textMatrixGate(ctx)
  const formatValue = evalNode(args[1], ctx)
  if (isEvaluationError(formatValue)) return formatValue
  // Matrix FORMAT gates BEFORE the scalar format type check, so a matrix
  // format never falls through to a counterfeit #VALUE!.
  if (isMatrixValue(formatValue)) return textMatrixGate(ctx)
  if (typeof formatValue !== 'string') return formulaError('#VALUE!')
  const result = formatCommon(value, formatValue, resolveSemantics(ctx))
  if (result.kind === 'text') return result.text
  if (isEvaluationError(result)) return result
  markUnsupported(ctx, 'format-unsupported')
  ctx?.reportFormulaIssue?.({
    kind: 'native-gate',
    gate: 'format-unsupported',
    message: `Format '${formatValue}' is outside the finite common-format grammar: ${result.reason}`,
  })
  return formulaError('#NAME?')
}

export const TEXT_FUNCTIONS: Record<string, FunctionHandler> = {
  TEXT: textHandler,
}
