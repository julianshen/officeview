/** JSON + CSV artifact writing for automated runs. */
import { randomBytes } from 'node:crypto'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import type { BaselineComparison } from './baseline'
import type { HarnessMetric, HarnessResult } from '../../src/debug/harness/types'

export interface RunMetadata {
  browser: string
  browserVersion: string
  requestedFixtureId: string
  profile: string
  mode: string
  commit: string
  repeat: number
  warmup: number
  baselineComparison?: BaselineComparison
  baselineError?: string
}

export interface CollectedRun extends RunMetadata {
  results: HarnessResult[]
}

/** RFC 4180 quoting plus neutralization of spreadsheet formula injection. */
export function csvField(value: unknown): string {
  if (value === null || value === undefined) return ''
  if (typeof value !== 'string') return String(value)
  const guarded = /^[=+\-@\t\r]/.test(value) ? `'${value}` : value
  return /[",\r\n]/.test(guarded) ? `"${guarded.replace(/"/g, '""')}"` : guarded
}

export function metricCell(metric: HarnessMetric, field: 'p50' | 'p95' | 'max' | 'samples'): string {
  switch (metric.status) {
    case 'unavailable':
      return ''
    case 'none-observed':
      return field === 'samples' ? '0' : ''
    case 'insufficient-samples':
      return field === 'p95' ? '' : String(metric[field])
    case 'ok':
      return String(metric[field])
  }
}

const CSV_HEADER =
  'record_type,result_index,requested_fixture,fixture_id,evidence,profile,mode,browser,sample_name,unit,value,warmup_excluded,metric_status,p50,p95,max,samples'
const NAME_ATTEMPTS = 5

function csvRows(results: HarnessResult[]): string[] {
  const lines = [CSV_HEADER]
  results.forEach((result, index) => {
    const env = result.environment
    const lead = (type: string) => [type, index, env.requestedFixtureId, env.fixtureId, env.evidence, env.profile, env.mode, env.browser.name]
    for (const sample of result.samples) {
      for (const value of sample.values) {
        lines.push([...lead('sample'), sample.name, sample.unit, value, sample.warmupExcluded, '', '', '', '', ''].map(csvField).join(','))
      }
    }
    for (const [name, metric] of Object.entries(result.metrics)) {
      lines.push(
        [...lead('metric'), name, '', '', '', metric.status, metricCell(metric, 'p50'), metricCell(metric, 'p95'), metricCell(metric, 'max'), metricCell(metric, 'samples')]
          .map(csvField)
          .join(','),
      )
    }
  })
  return lines
}

/**
 * Write `<base>.json` and `<base>.csv` exclusively (never overwriting). The
 * name comes from validated enums, the clock and a random id — not from result
 * content.
 */
export function writeArtifacts(
  run: CollectedRun,
  outputDir: string,
  now: () => Date = () => new Date(),
  newId: () => string = () => randomBytes(4).toString('hex'),
): { json: string; csv: string } {
  mkdirSync(outputDir, { recursive: true })
  const { results, ...metadata } = run
  const stamp = now().toISOString().replace(/[:.]/g, '-')
  let lastBase = ''
  for (let attempt = 0; attempt < NAME_ATTEMPTS; attempt++) {
    lastBase = `viewer-${run.browser}-${run.profile}-${run.requestedFixtureId}-${run.mode}-${stamp}-${newId()}`
    const json = join(outputDir, `${lastBase}.json`)
    const csv = join(outputDir, `${lastBase}.csv`)
    try {
      writeFileSync(json, `${JSON.stringify({ schemaVersion: 1, run: metadata, results }, null, 2)}\n`, { flag: 'wx' })
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'EEXIST') continue
      throw error
    }
    writeFileSync(csv, `${csvRows(results).join('\n')}\n`, { flag: 'wx' })
    return { json, csv }
  }
  throw new Error(`artifact ${lastBase} already exists after ${NAME_ATTEMPTS} attempts (EEXIST); nothing was overwritten`)
}
