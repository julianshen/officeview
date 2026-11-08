import { checkBitmap, type Bitmap, type Chip, type MarkParams } from './types'
import { checkParams } from './params'
import { pedestalFor, referenceCell } from './embed'
import { lum, register, type Registration } from './register'
export type Decision = { kind: 'attributed'; index: number; z: number; margin: number } | { kind: 'none'; zBest: number } | { kind: 'abstain'; reason: 'flat' }
/** Immutable handle; privately owned contiguous M×512 signs. */
export interface CandidateMatrix { readonly count: number; readonly bytes: number }
const matrices = new WeakMap<CandidateMatrix, Int8Array>()
export function prepareCandidates(codes: readonly Int8Array[]): CandidateMatrix {
  const bytes = codes.length * 512
  if (!Number.isSafeInteger(bytes) || bytes > 0x7fffffff) throw new Error('Candidate matrix too large')
  const data = new Int8Array(bytes)
  for (let m = 0; m < codes.length; m++) {
    const code = codes[m]!
    if (code.length !== 512) throw new Error('Expected extended BCH sign words')
    for (let n = 0; n < 512; n++) {
      const sign = code[n]!
      if (sign !== 1 && sign !== -1) throw new Error('Expected extended BCH sign words')
      data[m * 512 + n] = sign
    }
  }
  const handle = Object.freeze({ count: codes.length, bytes })
  matrices.set(handle, data)
  return handle
}
type CandidateInput = Int8Array[] | CandidateMatrix
export interface Scores { scores: Float64Array; support: number; registration?: Registration; flat: boolean }
/** Reference subtraction, shared pedestal subtraction, then support-centred
 * residual and RMS sigma. No unmatched spatial band-pass. The template is
 * also centred implicitly because sum(r-mean(r))=0. Original chip indices
 * An unmarked reference still has the hypothetical pedestal subtracted;
 * spatially varying pedestal can therefore inflate null sigma. The detector
 * does not infer whether embedding happened. Calibrate this full pipeline.
 * Original indices survive cropping/masking: never renumber a shortened codeword. */
export function scoreCapture(capture: Bitmap, reference: Bitmap, chips: Chip[], candidates: CandidateInput, params: MarkParams): Scores {
  checkBitmap(capture); checkBitmap(reference); checkParams(params)
  const matrix = Array.isArray(candidates) ? prepareCandidates(candidates) : candidates
  const codeData = matrices.get(matrix)
  if (!codeData) throw new Error('Invalid candidate matrix')
  const registration = register(capture, reference)
  if (!registration) return { scores: new Float64Array(), support: 0, flat: false }
  const { scaleX, scaleY, x: dx, y: dy, reference: scaledReference } = registration
  const values: number[] = [], indices: number[] = []
  for (let n = 0; n < chips.length; n++) {
    const chip = chips[n]!
    if (!Number.isInteger(chip.x) || !Number.isInteger(chip.y) || chip.cell !== params.cellPx || chip.x < 0 || chip.y < 0 || chip.x + chip.cell > reference.width || chip.y + chip.cell > reference.height || ![1, -1].includes(chip.sign0)) throw new Error('Invalid chip support')
    const pedestal = pedestalFor(referenceCell(reference, chip), params)
    if (pedestal === null) continue // shifted reference cannot admit both signs
    const x0 = Math.ceil((chip.x - dx) * scaleX), y0 = Math.ceil((chip.y - dy) * scaleY)
    const x1 = Math.floor((chip.x + chip.cell - dx) * scaleX), y1 = Math.floor((chip.y + chip.cell - dy) * scaleY)
    if (x0 < 0 || y0 < 0 || x1 > capture.width || y1 > capture.height || x1 <= x0 || y1 <= y0) continue
    let raw = 0
    for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) raw += lum(capture, x, y) - lum(scaledReference, x, y)
    raw /= (x1 - x0) * (y1 - y0)
    values.push(raw - pedestal); indices.push(n)
  }
  const L = values.length
  if (L < 2) return { scores: new Float64Array(), support: L, registration, flat: true }
  const mean = values.reduce((a, b) => a + b, 0) / L
  let sumSquares = 0
  for (let i = 0; i < L; i++) { values[i] -= mean; sumSquares += values[i]! ** 2 }
  if (sumSquares < 1e-16) return { scores: new Float64Array(), support: L, registration, flat: true }
  // sigma=sqrt(sumSquares/L), denominator sigma*sqrt(L)=sqrt(sumSquares).
  const denominator = Math.sqrt(sumSquares), scores = new Float64Array(matrix.count)
  for (let m = 0; m < matrix.count; m++) {
    let dot = 0
    for (let j = 0; j < L; j++) { const n = indices[j]!; dot += values[j]! * chips[n]!.sign0 * codeData[m * 512 + n % 512]! }
    scores[m] = dot / denominator
  }
  return { scores, support: L, registration, flat: false }
}
export function detect(capture: Bitmap, reference: Bitmap, chips: Chip[], candidates: CandidateInput, params: MarkParams, tau: { abs: number; margin: number }): Decision {
  if (!Number.isFinite(tau.abs) || !Number.isFinite(tau.margin) || tau.abs < 0 || tau.margin < 0) throw new Error('Invalid decision thresholds')
  const result = scoreCapture(capture, reference, chips, candidates, params)
  if (result.flat) return { kind: 'abstain', reason: 'flat' }
  let best = -Infinity, second = -Infinity, index = -1
  for (let m = 0; m < result.scores.length; m++) { const z = result.scores[m]!; if (z > best) { second = best; best = z; index = m } else if (z > second) second = z }
  if (index >= 0 && best >= tau.abs && best - second >= tau.margin) return { kind: 'attributed', index, z: best, margin: best - second }
  return { kind: 'none', zBest: best }
}
