import { describe, expect, it, vi } from 'vitest'
import { buildCfCandidate, createCfMemory, createCfCaps, createCfDiagnostics, createCfWorkBudget, measureCfPayload } from '../src/xlsx/cf/limits'
import { hasEvidence, withCfEvidence } from '../src/xlsx/cf/evidence'

function ready<T>(r: import('../src/xlsx/cf/model').CfOutcome<T>): T {
  if (r.kind !== 'ready') throw new Error(r.kind)
  return r.value
}
function reserved(r: ReturnType<ReturnType<typeof createCfMemory>['reserve']>) {
  if (r.kind !== 'reserved') throw new Error(r.reason)
  return r
}

describe('CF admission and ownership', () => {
  it('reserves before allocation and refuses one byte below the bound', () => {
    for (const limit of [1023, 1024]) {
      const memory = createCfMemory({ limitBytes: limit })
      const build = vi.fn(() => { expect(memory.currentBytes).toBe(1024); return new Uint8Array(1024) })
      const result = buildCfCandidate(memory, 1024, build)
      expect(result.kind).toBe(limit === 1024 ? 'ready' : 'unresolved')
      expect(build).toHaveBeenCalledTimes(limit === 1024 ? 1 : 0)
      if (result.kind === 'ready') { expect(memory.retainedBytes).toBe(1024); result.value.release(); result.value.release() }
      expect(memory.currentBytes).toBe(0)
    }
  })
  it('includes old retained results in growth peaks and retains them through close', () => {
    const memory = createCfMemory({ limitBytes: 1536 })
    const old = ready(buildCfCandidate(memory, 512, () => ({})))
    const r = reserved(memory.reserve(512))
    expect(r.grow(1024)).toBe(true); expect(r.grow(1024)).toBe(true)
    expect(memory.currentBytes).toBe(1536); expect(memory.constructionBytes).toBe(1024)
    expect(r.grow(1025)).toBe(false)
    const value = {}; const next = r.adopt(value, 768)
    expect(r.adopt(value, 768)).toBe(next)
    r.release() // ownership moved to the lease
    old.release(); expect(memory.currentBytes).toBe(768)
    memory.close(); memory.close()
    expect(memory.reserve(1).kind).toBe('unresolved')
    expect(memory.currentBytes).toBe(768) // consumer still holds it
    next.release(); next.release(); expect(memory.currentBytes).toBe(0)
  })
  it('settles exactly once on exception, refusal, replacement and transfer', () => {
    const a = createCfMemory({ limitBytes: 1024 }), b = createCfMemory({ limitBytes: 511 })
    expect(() => buildCfCandidate(a, 512, () => { throw new Error('builder') })).toThrow('builder')
    expect(a.currentBytes).toBe(0)
    const lease = ready(buildCfCandidate(a, 512, () => ({})))
    expect(lease.transfer(b).kind).toBe('unresolved'); expect(a.currentBytes).toBe(512)
    const dest = createCfMemory({ limitBytes: 512 })
    const moved = lease.transfer(dest); expect(moved.kind).toBe('ready')
    expect(a.currentBytes).toBe(0); expect(dest.currentBytes).toBe(512)
    lease.release(); expect(dest.currentBytes).toBe(512)
    ready(moved).release(); expect(dest.currentBytes).toBe(0)
    const pending = reserved(a.reserve(100)); a.close(); expect(a.currentBytes).toBe(100)
    pending.release(); expect(a.currentBytes).toBe(0)
  })
  it('honours a live external account and keeps the worker grant separate', () => {
    let total = 200, calls = 0
    const context = { limitBytes: 1024, canReserve: (n: number) => total + n <= 1024, owns: () => false,
      reserve: () => { let bytes = 0; return { owner: 'cf', get bytes() { return bytes }, trySet(n: number) { calls++; if (total - bytes + n > 1024) return false; total += n - bytes; bytes = n; return true }, charge() {}, adopt() {}, release() { total -= bytes; bytes = 0 } } } }
    const memory = createCfMemory({ limitBytes: 1000, context })
    expect(buildCfCandidate(memory, 825, () => ({})).kind).toBe('unresolved')
    const lease = ready(buildCfCandidate(memory, 824, () => ({})))
    expect(total).toBe(1024); expect(calls).toBeGreaterThan(0)
    lease.release(); expect(total).toBe(200)
  })
  it('rejects invalid bytes and bounds exact accounting workspace', () => {
    const m = createCfMemory({ limitBytes: 8192 })
    for (const bytes of [-1, NaN, Infinity, 1.5, Number.MAX_SAFE_INTEGER + 1]) expect(m.reserve(bytes).kind).toBe('unresolved')
    const backing = new ArrayBuffer(64), a = new Uint8Array(backing, 0, 1), b = new Uint8Array(backing, 1, 1)
    expect(measureCfPayload([a, b], m)).toEqual({ kind: 'ready', value: 96 }) // array header/slots + one full backing
    expect(m.currentBytes).toBe(0)
    expect(measureCfPayload([a, b], createCfMemory({ limitBytes: 1 })).kind).toBe('unresolved')
  })
})

describe('CF bounded work and metadata', () => {
  it('caps diagnostics at 64 and refuses unknown over-cap rule scope', () => {
    const diagnostics = createCfDiagnostics()
    for (let n = 0; n < 65; n++) diagnostics.add({ kind: 'unsupported', feature: 'cf-test', message: String(n) })
    expect(diagnostics.entries).toHaveLength(64)
    diagnostics.add({ kind: 'unsupported', feature: 'cf-test', message: '0' }); expect(diagnostics.entries).toHaveLength(64)
    const caps = createCfCaps()
    for (let n = 0; n < 4096; n++) expect(caps.addRule(1).kind).toBe('ready')
    expect(caps.addRule(1)).toMatchObject({ kind: 'unresolved', scope: 'sheet', writes: 'ALL' })
    expect(caps.status.kind).toBe('unresolved')
    expect(caps.addRule(0).kind).toBe('unresolved')
  })
  it('enforces exact limits including per-rule ranges and sheet/workbook metadata', () => {
    for (const [name, max] of Object.entries({ textUnits: 8192, formulaTokens: 4096, formulaChildren: 3, cfvo: 5, customIcons: 5, preparedCells: 5000, matchingRules: 4096, descriptorUnits: 128 })) {
      const caps = createCfCaps(); expect(caps.check(name, max).kind).toBe('ready'); expect(caps.check(name, max + 1).kind).toBe('unresolved')
    }
    expect(createCfCaps().addRule(257).kind).toBe('unresolved')
    const caps = createCfCaps(); for (let n = 0; n < 64; n++) expect(caps.addRule(256).kind).toBe('ready')
    expect(caps.addRule(1).kind).toBe('unresolved')
    const book = createCfMemory(); const sheets = Array.from({ length: 4 }, () => createCfCaps(book))
    for (const sheet of sheets) expect(sheet.reserveMetadata(4 * 1024 ** 2).kind).not.toBe('unresolved')
    expect(sheets[0].reserveMetadata(1).kind).toBe('unresolved')
    expect(createCfCaps(book).reserveMetadata(1).kind).toBe('unresolved')
    sheets.forEach(s => s.dispose()); expect(book.currentBytes).toBe(0)
  })
  it('budgets work and active time, excluding the time spent yielded', () => {
    let now = 0
    const budget = createCfWorkBudget({ now: () => now })
    expect(budget.spend(127)).toBe('continue'); expect(budget.spend(1)).toBe('yield')
    budget.resume(); now = 4; expect(budget.spend(1)).toBe('yield')
    now = 10000; budget.resume(); expect(budget.spend(1)).toBe('continue')
    now += 46; expect(budget.spend(128)).toBe('refused')
    expect(budget.spend(0)).toBe('refused')
    const units = createCfWorkBudget({ now: () => 0 })
    for (let n = 0; n < 7812; n++) { expect(units.spend(128)).toBe('yield'); units.resume() }
    expect(units.spend(64)).toBe('continue'); expect(units.spend(1)).toBe('refused')
  })
  it('absent evidence refuses despite matching synthetic values', () => {
    const evaluate = vi.fn(() => true)
    expect(hasEvidence('stage1-cellis-literals', 'greaterThan')).toBe(false)
    expect(withCfEvidence('stage1-cellis-literals', 'greaterThan', evaluate).kind).toBe('unresolved')
    expect(evaluate).not.toHaveBeenCalled()
  })
})

it('metadata and diagnostics admit bounded storage before growth', () => {
  const book = createCfMemory({ limitBytes: 10 }), caps = createCfCaps(book)
  expect(caps.reserveMetadata(8).kind).toBe('ready')
  expect(caps.reserveMetadata(3).kind).toBe('unresolved')
  expect(book.currentBytes).toBe(8)
  expect(caps.check('textUnits', 1).kind).toBe('unresolved') // sheet sentinel is sticky
  caps.dispose(); expect(book.currentBytes).toBe(0)
  const m = createCfMemory({ limitBytes: 1 }), diagnostics = createCfDiagnostics(m)
  diagnostics.add({ kind: 'unsupported', feature: 'cf-test', message: 'large' })
  expect(diagnostics.entries).toHaveLength(0)
  expect(diagnostics.refused).toBe(true)
  expect(m.currentBytes).toBe(0)
})

it('duplicate diagnostics do not consume repeated admission', () => {
 const m = createCfMemory({limitBytes: 256}), d = createCfDiagnostics(m)
 for (let i=0;i<100;i++) d.add({kind:'unsupported',feature:'cf-test',message:'same'})
 expect(d.entries).toHaveLength(1); expect(d.refused).toBe(false)
 expect(m.currentBytes).toBe(214); d.dispose(); expect(m.currentBytes).toBe(0)
})
