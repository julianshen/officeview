// Slideshow frame comparison (S1-21). Compares a rendered frame against a
// PowerPoint reference frame and records the result. The renderer CLI lands
// in S1-22; this module only provides the comparison primitives and record I/O.
import { readdir } from 'node:fs/promises'

export interface Tolerance { meanAbs: number; rowMax: number }

export const TOLERANCE_EXACT: Tolerance = { meanAbs: 2, rowMax: 8 }
export const TOLERANCE_APPROX: Tolerance = { meanAbs: 6, rowMax: 24 }

export interface CompareResult { pass: boolean; meanAbs: number; rowMax: number; worstRow: number }

export function compareFrames(a: ImageData, b: ImageData, tol: Tolerance): CompareResult {
  const n = Math.min(a.data.length, b.data.length)
  const width = a.width
  const rows = a.height
  const perRowAbs: number[] = new Array(rows).fill(0)
  const perRowCount: number[] = new Array(rows).fill(0)
  let totalAbs = 0
  let count = 0
  for (let i = 0; i < n; i += 4) {
    const row = Math.floor((i / 4) / width)
    const d = Math.abs(a.data[i] - b.data[i]) + Math.abs(a.data[i+1] - b.data[i+1]) + Math.abs(a.data[i+2] - b.data[i+2])
    const abs = d / 3
    totalAbs += abs
    count++
    perRowAbs[row] += abs
    perRowCount[row]++
  }
  const meanAbs = count ? totalAbs / count : 0
  let rowMax = 0, worstRow = 0
  for (let r = 0; r < rows; r++) {
    const m = perRowCount[r] ? perRowAbs[r] / perRowCount[r] : 0
    if (m > rowMax) { rowMax = m; worstRow = r }
  }
  const pass = meanAbs <= tol.meanAbs && rowMax <= tol.rowMax
  return { pass, meanAbs, rowMax, worstRow }
}

export interface RefSet { files: string[] }

export async function readRefs(dir: string): Promise<RefSet> {
  try {
    const entries = await readdir(dir, { recursive: true } as any)
    const files = entries.filter(e => typeof e === 'string' && e.endsWith('.png')) as string[]
    return { files }
  } catch {
    return { files: [] }
  }
}

export interface CompareRecord {
  row: string; deck: string; slide: number; step: number; pct: number
  pass: boolean; meanAbs: number; rowMax: number
}

export function writeCompareRecord(path: string, rows: CompareRecord[]): void {
  const fs = require('node:fs')
  fs.writeFileSync(path, JSON.stringify(rows, null, 2))
}

// CLI entry
if (import.meta.url === `file://${process.argv[1]}`) {
  console.log('slideshow-compare: renderer CLI lands in S1-22')
}
