import { describe, expect, it } from 'vitest'
import { gateRows, renderReport } from '../../scripts/stream-s0/report'

describe('Stage-0 report', () => {
  it('lists NOT RUN for a missing json', () => {
    const rows = gateRows(new Map([['T1', undefined]]))
    expect(rows.find((r) => r.id === 'T1')?.status).toBe('NOT RUN')
    expect(rows.find((r) => r.id === 'T12')?.status).toBe('NOT RUN')
  })
  it('RED gates appear in the owner section', () => {
    const md = renderReport(new Map([['T8', { status: 'RED', bestOf3P95: 300, sheet: { rows: 500, cols: 100 }, tilePx: 512 }]]))
    const owner = md.split('## RED gates — owner decision needed')[1]!.split('## ')[0]!
    expect(owner).toContain('T8 XLSX tile timing: RED')
    expect(owner).toContain('T1 Backend confirm')
    expect(owner).toContain('NOT RUN')
  })
})
