import { checkBitmap, type Bitmap, type Chip, type MarkParams } from './types'
import { checkParams } from './params'
import { pedestalFor, referenceCell } from './embed'
import { lum, register, type Registration } from './register'
export type Decision = { kind: 'attributed'; index: number; z: number; margin: number } | { kind: 'none'; zBest: number } | { kind: 'abstain'; reason: 'flat' }
export interface Scores { scores: Float64Array; support: number; registration?: Registration; flat: boolean }
/** Reference subtraction, shared pedestal subtraction, then support-centred
 * residual and RMS sigma. No unmatched spatial band-pass. The template is
 * also centred implicitly because sum(r-mean(r))=0. Original chip indices
 * survive cropping/masking: never renumber a shortened codeword. */
export function scoreCapture(capture: Bitmap, reference: Bitmap, chips: Chip[], candidates: Int8Array[], params: MarkParams): Scores {
  checkBitmap(capture); checkBitmap(reference); checkParams(params)
  if (candidates.some(code => code.length !== 512 || code.some(v => v !== 1 && v !== -1))) throw new Error('Expected extended BCH sign words')
  const registration = register(capture, reference)
  if (!registration) return { scores: new Float64Array(), support: 0, flat: false }
  const { factor, x: dx, y: dy, reference: scaledReference } = registration
  const values: number[] = [], indices: number[] = []
  let rawEnergy = 0
  for (let n = 0; n < chips.length; n++) {
    const chip = chips[n]!
    if (!Number.isInteger(chip.x) || !Number.isInteger(chip.y) || chip.cell !== params.cellPx || chip.x < 0 || chip.y < 0 || chip.x + chip.cell > reference.width || chip.y + chip.cell > reference.height || ![1, -1].includes(chip.sign0)) throw new Error('Invalid chip support')
    const pedestal = pedestalFor(referenceCell(reference, chip), params)
    if (pedestal === null) continue // shifted reference cannot admit both signs
    const x0 = Math.ceil(chip.x * factor - dx), y0 = Math.ceil(chip.y * factor - dy)
    const x1 = Math.floor((chip.x + chip.cell) * factor - dx), y1 = Math.floor((chip.y + chip.cell) * factor - dy)
    if (x0 < 0 || y0 < 0 || x1 > capture.width || y1 > capture.height || x1 <= x0 || y1 <= y0) continue
    let raw = 0
    for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) raw += lum(capture, x, y) - lum(scaledReference, x + dx, y + dy)
    raw /= (x1 - x0) * (y1 - y0)
    rawEnergy += raw * raw
    values.push(raw - pedestal); indices.push(n)
  }
  const L = values.length
  if (L < 2 || rawEnergy < 1e-16) return { scores: new Float64Array(), support: L, registration, flat: true }
  const mean = values.reduce((a, b) => a + b, 0) / L
  let sumSquares = 0
  for (let i = 0; i < L; i++) { values[i] -= mean; sumSquares += values[i]! ** 2 }
  if (sumSquares < 1e-16) return { scores: new Float64Array(), support: L, registration, flat: true }
  // sigma=sqrt(sumSquares/L), denominator sigma*sqrt(L)=sqrt(sumSquares).
  const denominator = Math.sqrt(sumSquares), scores = new Float64Array(candidates.length)
  for (let m = 0; m < candidates.length; m++) {
    let dot = 0
    for (let j = 0; j < L; j++) { const n = indices[j]!; dot += values[j]! * chips[n]!.sign0 * candidates[m]![n % 512]! }
    scores[m] = dot / denominator
  }
  return { scores, support: L, registration, flat: false }
}
export function detect(capture: Bitmap, reference: Bitmap, chips: Chip[], candidates: Int8Array[], params: MarkParams, tau: { abs: number; margin: number }): Decision {
  if (!Number.isFinite(tau.abs) || !Number.isFinite(tau.margin) || tau.abs < 0 || tau.margin < 0) throw new Error('Invalid decision thresholds')
  const result = scoreCapture(capture, reference, chips, candidates, params)
  if (result.flat) return { kind: 'abstain', reason: 'flat' }
  let best = -Infinity, second = -Infinity, index = -1
  for (let m = 0; m < result.scores.length; m++) { const z = result.scores[m]!; if (z > best) { second = best; best = z; index = m } else if (z > second) second = z }
  if (index >= 0 && best >= tau.abs && best - second >= tau.margin) return { kind: 'attributed', index, z: best, margin: best - second }
  return { kind: 'none', zBest: best }
}
