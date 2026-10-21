import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { compileFormula, type FormulaEnv } from '../src/slideshow/formula'

function env(over: Partial<FormulaEnv> = {}): FormulaEnv {
  return {
    $: 0.25,
    ppt_x: 0.5,
    ppt_y: 0.4,
    ppt_w: 0.2,
    ppt_h: 0.1,
    current: (name: string) => ({ ppt_x: 0.6, ppt_y: 0.7, ppt_w: 0.3, ppt_h: 0.35 } as Record<string, number>)[name] ?? Number.NaN,
    rand: () => 0.5,
    ...over,
  }
}

function run(src: string, e: FormulaEnv = env()): number {
  const r = compileFormula(src)
  if (!r.ok) throw new Error(`expected ok, got ${r.reason}: ${src.slice(0, 60)}`)
  return r.fn(e)
}

function refusal(src: string): string | null {
  const r = compileFormula(src)
  return r.ok ? null : r.reason
}

/** `n` levels of grouping parentheses around `1`. */
function parens(n: number): string { return '('.repeat(n) + '1' + ')'.repeat(n) }
/** `n` nested one-argument calls around `1`. */
function calls(n: number): string { return 'abs('.repeat(n) + '1' + ')'.repeat(n) }

describe('slideshow formula interpreter', () => {
  it('evaluates arithmetic, precedence, right-associative ^ and unary minus', () => {
    expect(run('1+2*3')).toBe(7)
    expect(run('(1+2)*3')).toBe(9)
    expect(run('10-4-3')).toBe(3)
    expect(run('8/4/2')).toBe(1)
    expect(run('2^3^2')).toBe(512)
    expect(run('-2^2')).toBe(-4)
    expect(run('2^-1')).toBe(0.5)
    expect(run('--3')).toBe(3)
    expect(run('1 - -1')).toBe(2)
    expect(run('.5 + 1.25 + 2.')).toBeCloseTo(3.75, 12)
    expect(run('1e3 + 2.5E-1')).toBeCloseTo(1000.25, 12)
    expect(run('pi')).toBeCloseTo(Math.PI, 12)
    expect(run('e')).toBeCloseTo(Math.E, 12)
    expect(run('max(1, 2) + min(3, -4)')).toBe(-2)
    expect(run('abs(-2) + ceil(1.2) + floor(1.8)')).toBe(5)
    expect(run('sqrt(16) + exp(0) + ln(e)')).toBeCloseTo(6, 12)
    expect(run('sin(pi/2) + cos(0) + tan(0)')).toBeCloseTo(2, 12)
    expect(run('asin(1) + acos(1) + atan(1)')).toBeCloseTo(Math.PI / 2 + Math.PI / 4, 12)
    expect(run('deg(pi) + rad(180)')).toBeCloseTo(180 + Math.PI, 12)
    expect(run('$*4')).toBe(1)
    // Whitespace is free.
    expect(run('  1\t+\n2 ')).toBe(3)
    // Non-finite results collapse to NaN; the caller diagnoses it.
    for (const src of ['1/0', '-1/0', 'sqrt(-1)', 'ln(0)', '9^9^9^9', '0/0', 'acos(2)']) expect(run(src), src).toBeNaN()
    // A non-finite literal is refused at compile time.
    expect(refusal('1e999')).toBe('syntax')
    // Malformed expressions.
    for (const src of ['', '   ', '1+', '*1', '(1', '1)', '()', '1 2', '2e', '2 e', 'abs', 'abs()', 'abs(1,2)', 'max(1)', 'max(1,2,3)',
      'max(1,)', '1..2', '1.2.3', '+1', '1%2', 'a?b:c', '1,2', 'pi(1)', '$(1)', '#', '# ppt_x', '#1'])
      expect(refusal(src), src).not.toBeNull()
    expect(refusal('1 2')).toBe('syntax')
    expect(refusal('abs(1,2)')).toBe('syntax')
  })

  it('#ppt_x and ppt_w read the environment', () => {
    expect(run('ppt_x')).toBe(0.5)
    expect(run('ppt_y')).toBe(0.4)
    expect(run('ppt_w')).toBe(0.2)
    expect(run('ppt_h')).toBe(0.1)
    expect(run('#ppt_x')).toBe(0.6)
    expect(run('#ppt_y')).toBe(0.7)
    expect(run('#ppt_w')).toBe(0.3)
    expect(run('#ppt_h')).toBe(0.35)
    expect(run('#ppt_x-ppt_w/2')).toBeCloseTo(0.5, 12)
    const asked: string[] = []
    run('#ppt_x + #ppt_h', env({ current: n => { asked.push(n); return 1 } }))
    expect(asked).toEqual(['ppt_x', 'ppt_h'])
    // The same formula, a different environment.
    const r = compileFormula('ppt_x + $')
    expect(r.ok).toBe(true)
    if (r.ok) {
      expect(r.fn(env())).toBe(0.75)
      expect(r.fn(env({ ppt_x: 2, $: 1 }))).toBe(3)
    }
    // Hostile environments never throw out of the compiled function.
    expect(run('#ppt_x', env({ current: () => { throw new Error('boom') } }))).toBeNaN()
    expect(run('ppt_x', env({ ppt_x: Number.POSITIVE_INFINITY }))).toBeNaN()
    expect(run('ppt_x', env({ ppt_x: 'x' as unknown as number }))).toBeNaN()
    expect(run('#ppt_y', env({ current: () => Number.NaN }))).toBeNaN()
    // Only the four ppt_* names take a # form.
    for (const src of ['#pi', '#$', '#foo', '#constructor', '#__proto__', '#e', '#abs'])
      expect(refusal(src), src).toBe('unknown-identifier')
  })

  it('rand uses the injected Rng', () => {
    let calls = 0
    const rng = () => { calls++; return [0.25, 0.75][(calls - 1) % 2] }
    const r = compileFormula('rand(4) + rand(8)')
    expect(r.ok).toBe(true)
    if (!r.ok) return
    expect(r.fn(env({ rand: rng }))).toBe(1 + 6)
    expect(calls).toBe(2)
    // A fresh draw on every evaluation.
    expect(r.fn(env({ rand: rng }))).toBe(1 + 6)
    expect(calls).toBe(4)
    // A misbehaving generator yields NaN rather than throwing.
    expect(run('rand(1)', env({ rand: () => Number.NaN }))).toBeNaN()
    expect(run('rand(1)', env({ rand: () => { throw new Error('x') } }))).toBeNaN()
    expect(refusal('rand()')).toBe('syntax')
    expect(refusal('rand')).toBe('syntax')
  })

  it('accepts 1024 bytes, 256 tokens and depth 32; rejects 1025 bytes, 257 tokens, depth 33 and every identifier not in the list', () => {
    // Bytes: a valid expression padded with spaces (spaces are not tokens).
    expect(run('1+1' + ' '.repeat(1021))).toBe(2)
    expect(refusal('1+1' + ' '.repeat(1022))).toBe('too-long')
    // Bytes are UTF-8 bytes, not UTF-16 units: 342 three-byte characters exceed the limit.
    expect(refusal(' '.repeat(342))).toBe('too-long')
    expect(refusal(' '.repeat(341))).toBe('syntax')
    expect(refusal('x'.repeat(100000))).toBe('too-long')
    // Tokens: `-1+1+…+1` with 128 ones is 256 tokens; 129 ones and 128 pluses is 257.
    const t256 = '-1' + '+1'.repeat(127)
    const t257 = '1' + '+1'.repeat(128)
    expect(run(t256)).toBe(126)
    expect(refusal(t257)).toBe('too-many-tokens')
    // `#ppt_x` is one token.
    expect(run('-#ppt_x' + '+#ppt_x'.repeat(127))).toBeCloseTo(126 * 0.6, 9)
    // Depth: nesting of parentheses, grouping or call arguments alike.
    expect(run(parens(32))).toBe(1)
    expect(refusal(parens(33))).toBe('too-deep')
    expect(run(calls(32))).toBe(1)
    expect(refusal(calls(33))).toBe('too-deep')
    expect(run('max(' + parens(31) + ', 2)')).toBe(2)
    expect(refusal('max(' + parens(32) + ', 2)')).toBe('too-deep')
    // Long ^ and unary chains stay within the token cap and evaluate without recursion trouble.
    expect(run('1' + '^1'.repeat(127))).toBe(1)
    expect(run('-'.repeat(255) + '1')).toBe(-1)
    // Every identifier not in the list.
    for (const id of ['x', 'y', 'ppt_r', 'cosh', 'sinh', 'tanh', 'log', 'pow', 'round', 'PI', 'E', 'Abs', 'SIN', 'ppt_X', 'ppt', '_', 'ppt_x2', 'pie', 'math', 'Math',
      'window', 'document', 'alert', 'undefined', 'NaN', 'Infinity', 'null', 'true'])
      expect(refusal(id), id).toBe('unknown-identifier')
    expect(refusal('cosh(1)')).toBe('unknown-identifier')
    expect(refusal('1 + foo(2)')).toBe('unknown-identifier')
  })

  it('rejects constructor, __proto__, this, globalThis, brackets, quotes, semicolons and backticks', () => {
    for (const id of ['constructor', '__proto__', 'prototype', 'this', 'globalThis', 'hasOwnProperty', 'toString', 'valueOf', 'constructor(1)', 'toString(1)', '__proto__(1)'])
      expect(refusal(id), id).toBe('unknown-identifier')
    for (const src of ['[1]', 'abs[1]', '{1}', '"1"', "'1'", '1;2', '1;', '`1`', '`${1}`', 'a.b', 'pi.constructor', '1=1', '1==1', '!1', '1&1', '1|1', '\\1', '@', '1<2', 'x=>1',
      'abs\u0000(1)', ' 1'])
      expect(refusal(src), src).toBe('syntax')
    // Deterministic fuzz: nothing throws out of compile or evaluation, and results are finite or NaN.
    const alphabet = '0123456789.eE+-*/^(),$# _abcdilmnoprstxwhy[]"\';`{} '
    let seed = 12345
    const next = () => { seed = (seed * 1103515245 + 12345) >>> 0; return seed / 2 ** 32 }
    for (let n = 0; n < 3000; n++) {
      let s = ''
      const len = Math.floor(next() * 40)
      for (let i = 0; i < len; i++) s += alphabet[Math.floor(next() * alphabet.length)]
      const r = compileFormula(s)
      if (r.ok) {
        const v = r.fn(env())
        expect(Number.isNaN(v) || Number.isFinite(v), s).toBe(true)
      } else expect(['too-long', 'too-many-tokens', 'too-deep', 'syntax', 'unknown-identifier']).toContain(r.reason)
    }
    expect(refusal(undefined as unknown as string)).toBe('syntax')
  })

  it('the module contains no eval or Function', () => {
    const src = readFileSync(resolve(__dirname, '../src/slideshow/formula.ts'), 'utf8')
    expect(src).not.toMatch(/\beval\b/)
    expect(src).not.toMatch(/\bFunction\b/)
    expect(src).not.toMatch(/\bnew\s+Function\b/)
    expect(src).not.toMatch(/\bimport\s*\(/)
    expect(src).not.toMatch(/\bsetTimeout\b|\bsetInterval\b/)
  })
})
