// Bounded animation formula interpreter (spec 6.1 formulas, spec 7 hostile input).
//
// `p:anim@formula`, `p:tav@fmla` and values such as `#ppt_x` are parsed by a small
// recursive-descent parser into an AST, which is then walked by a switch. Nothing here ever
// hands deck text to a JavaScript evaluator, and identifier lookups go through Maps/Sets, never
// plain objects, so names like `constructor` or `__proto__` cannot reach Object.prototype.
//
// Grammar (whitespace is ASCII space, tab, CR, LF and separates tokens; it is not a token):
//   expr    := term (('+' | '-') term)*
//   term    := unary (('*' | '/' | '%') unary)*
//   unary   := ('-' | '+')* power
//   power   := primary ('^' unary)?           right-associative, binds tighter than unary minus
//   primary := number | var | '#' hashName | func '(' args ')' | '(' expr ')'
//   var     := '$' | 'ppt_x' | 'ppt_y' | 'ppt_w' | 'ppt_h' | 'ppt_r' | 'pi' | 'e'
//   hashName:= 'ppt_x' | 'ppt_y' | 'ppt_w' | 'ppt_h' | 'ppt_r'  (read through env.current)
//   func    := abs acos asin atan ceil cos cosh deg exp floor ln max min rad rand sin sinh sqrt tan tanh
//   number  := digits ['.' digits*] | '.' digits, then an optional exponent [eE][+-]?digits
//
// Decisions (the brief is silent on these; see task-S1-5-report.md for ECMA-376 differences):
//  - Identifiers (functions, variables, constants, `#` names) are case-insensitive: they are
//    lower-cased before lookup, and env.current always receives the lower-case name.
//  - `-2^2` is -4 and `2^-1` is 0.5. Unary plus is accepted and does nothing.
//  - `%` is the remainder (sign of the dividend), at `*` and `/` precedence, left-associative;
//    `x % 0` is NaN. Taken as modulo per review; not verified against the ECMA-376 text.
//  - Arity: max and min take exactly two arguments, every other function exactly one, except
//    `rand`: `rand()` is one draw from the injected Rng and `rand(x)` is x times one draw; a
//    fresh draw is taken on every evaluation.
//  - `ppt_r` (rotation, degrees) is an optional FormulaEnv field; when absent it reads as NaN.
//  - `#` forms are limited to the five ppt_* names; any other `#name` is unknown-identifier.
//  - A semicolon-separated formula list (`p:tav@fmla` per segment) is the caller's to split;
//    compileFormula takes one expression and refuses `;` as syntax.
//  - Limits: 1024 UTF-8 bytes, 256 tokens (`#ppt_x` is one token), and depth 32, where depth is
//    the nesting of parentheses (grouping and call argument lists alike).
//  - A non-finite literal (`1e999`) is refused as syntax. At evaluation time any non-finite
//    intermediate or final value, a non-number environment value, a throwing `current` or
//    `rand`, collapses to NaN: a compiled formula never throws, and callers diagnose NaN.
import type { Rng } from './types'
import { READER_LIMITS } from './read/limits'

export interface FormulaEnv {
  $: number
  ppt_x: number
  ppt_y: number
  ppt_w: number
  ppt_h: number
  /** Rotation in degrees. Optional so the briefed shape stays valid; absent reads as NaN. */
  ppt_r?: number
  current: (name: string) => number
  rand: Rng
}

/** Never throws. Returns NaN for any non-finite outcome; the caller diagnoses `Number.isNaN`. */
export type CompiledFormula = (env: FormulaEnv) => number

export type FormulaRefusal = 'too-long' | 'too-many-tokens' | 'too-deep' | 'syntax' | 'unknown-identifier'

type Tok =
  | { k: 'num'; v: number }
  | { k: 'id'; v: string }
  | { k: 'hash'; v: string }
  | { k: 'op'; v: string }

type EnvVar = '$' | 'ppt_x' | 'ppt_y' | 'ppt_w' | 'ppt_h' | 'ppt_r'

type Node =
  | { t: 'num'; v: number }
  | { t: 'var'; v: EnvVar }
  | { t: 'cur'; v: string }
  | { t: 'neg'; a: Node }
  | { t: 'bin'; op: string; a: Node; b: Node }
  | { t: 'call'; f: (x: number, y: number) => number; args: Node[] }
  | { t: 'rand'; a: Node | undefined }

const ENV_VARS: ReadonlySet<string> = new Set(['$', 'ppt_x', 'ppt_y', 'ppt_w', 'ppt_h', 'ppt_r'])
const HASH_NAMES: ReadonlySet<string> = new Set(['ppt_x', 'ppt_y', 'ppt_w', 'ppt_h', 'ppt_r'])
const CONSTANTS: ReadonlyMap<string, number> = new Map([['pi', Math.PI], ['e', Math.E]])

const FUNCS: ReadonlyMap<string, { arity: 1 | 2; f: (x: number, y: number) => number }> = new Map([
  ['abs', { arity: 1, f: x => Math.abs(x) }],
  ['acos', { arity: 1, f: x => Math.acos(x) }],
  ['asin', { arity: 1, f: x => Math.asin(x) }],
  ['atan', { arity: 1, f: x => Math.atan(x) }],
  ['ceil', { arity: 1, f: x => Math.ceil(x) }],
  ['cos', { arity: 1, f: x => Math.cos(x) }],
  ['cosh', { arity: 1, f: x => Math.cosh(x) }],
  ['deg', { arity: 1, f: x => x * 180 / Math.PI }],
  ['exp', { arity: 1, f: x => Math.exp(x) }],
  ['floor', { arity: 1, f: x => Math.floor(x) }],
  ['ln', { arity: 1, f: x => Math.log(x) }],
  ['max', { arity: 2, f: (x, y) => Math.max(x, y) }],
  ['min', { arity: 2, f: (x, y) => Math.min(x, y) }],
  ['rad', { arity: 1, f: x => x * Math.PI / 180 }],
  ['sin', { arity: 1, f: x => Math.sin(x) }],
  ['sinh', { arity: 1, f: x => Math.sinh(x) }],
  ['sqrt', { arity: 1, f: x => Math.sqrt(x) }],
  ['tan', { arity: 1, f: x => Math.tan(x) }],
  ['tanh', { arity: 1, f: x => Math.tanh(x) }],
] as const)

class Refusal {
  constructor(readonly reason: FormulaRefusal) {}
}

function utf8Length(s: string, cap: number): number {
  let n = 0
  for (let i = 0; i < s.length && n <= cap; i++) {
    const c = s.charCodeAt(i)
    if (c < 0x80) n += 1
    else if (c < 0x800) n += 2
    else if (c >= 0xd800 && c <= 0xdbff) { n += 4; i++ } // a surrogate pair is one 4-byte code point
    else n += 3
  }
  return n
}

const isDigit = (c: string): boolean => c >= '0' && c <= '9'
const isIdStart = (c: string): boolean => (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c === '_'
const isIdChar = (c: string): boolean => isIdStart(c) || isDigit(c)

function tokenize(src: string): Tok[] {
  const toks: Tok[] = []
  const push = (t: Tok): void => {
    if (toks.length >= READER_LIMITS.formulaTokens) throw new Refusal('too-many-tokens')
    toks.push(t)
  }
  let i = 0
  while (i < src.length) {
    const c = src[i]
    if (c === ' ' || c === '\t' || c === '\r' || c === '\n') { i++; continue }
    if (isDigit(c) || (c === '.' && isDigit(src[i + 1] ?? ''))) {
      const start = i
      while (isDigit(src[i] ?? '')) i++
      if (src[i] === '.') { i++; while (isDigit(src[i] ?? '')) i++ }
      if (src[i] === 'e' || src[i] === 'E') {
        let j = i + 1
        if (src[j] === '+' || src[j] === '-') j++
        if (isDigit(src[j] ?? '')) { i = j; while (isDigit(src[i] ?? '')) i++ }
      }
      const v = Number(src.slice(start, i))
      if (!Number.isFinite(v)) throw new Refusal('syntax')
      push({ k: 'num', v })
      continue
    }
    if (isIdStart(c)) {
      const start = i
      while (isIdChar(src[i] ?? '')) i++
      push({ k: 'id', v: src.slice(start, i).toLowerCase() })
      continue
    }
    if (c === '$') { i++; push({ k: 'id', v: '$' }); continue }
    if (c === '#') {
      const start = ++i
      while (isIdChar(src[i] ?? '') || (i === start && src[i] === '$')) i++
      if (i === start) throw new Refusal('syntax')
      push({ k: 'hash', v: src.slice(start, i).toLowerCase() })
      continue
    }
    if ('+-*/%^(),'.includes(c)) { i++; push({ k: 'op', v: c }); continue }
    throw new Refusal('syntax')
  }
  return toks
}

function parse(toks: readonly Tok[]): Node {
  let pos = 0
  let depth = 0
  const peekOp = (v: string): boolean => { const t = toks[pos]; return t !== undefined && t.k === 'op' && t.v === v }
  const expect = (v: string): void => { if (!peekOp(v)) throw new Refusal('syntax'); pos++ }
  const open = (): void => {
    expect('(')
    if (++depth > READER_LIMITS.formulaDepth) throw new Refusal('too-deep')
  }
  const close = (): void => { expect(')'); depth-- }

  function expr(): Node {
    let a = term()
    while (peekOp('+') || peekOp('-')) {
      const op = (toks[pos++] as { v: string }).v
      a = { t: 'bin', op, a, b: term() }
    }
    return a
  }
  function term(): Node {
    let a = unary()
    while (peekOp('*') || peekOp('/') || peekOp('%')) {
      const op = (toks[pos++] as { v: string }).v
      a = { t: 'bin', op, a, b: unary() }
    }
    return a
  }
  function unary(): Node {
    let negs = 0
    for (;;) {
      if (peekOp('-')) negs++
      else if (!peekOp('+')) break
      pos++
    }
    const p = power()
    return negs % 2 === 1 ? { t: 'neg', a: p } : p
  }
  function power(): Node {
    const base = primary()
    if (!peekOp('^')) return base
    pos++
    return { t: 'bin', op: '^', a: base, b: unary() }
  }
  function primary(): Node {
    const t = toks[pos]
    if (t === undefined) throw new Refusal('syntax')
    if (t.k === 'num') { pos++; return { t: 'num', v: t.v } }
    if (t.k === 'hash') {
      if (!HASH_NAMES.has(t.v)) throw new Refusal('unknown-identifier')
      pos++
      return { t: 'cur', v: t.v }
    }
    if (t.k === 'op') {
      if (t.v !== '(') throw new Refusal('syntax')
      open()
      const e = expr()
      close()
      return e
    }
    // Identifier.
    pos++
    if (ENV_VARS.has(t.v)) return { t: 'var', v: t.v as EnvVar }
    const k = CONSTANTS.get(t.v)
    if (k !== undefined) return { t: 'num', v: k }
    const fn = FUNCS.get(t.v)
    if (fn === undefined && t.v !== 'rand') throw new Refusal('unknown-identifier')
    open()
    if (fn === undefined) { // rand() or rand(x)
      const a = peekOp(')') ? undefined : expr()
      close()
      return { t: 'rand', a }
    }
    const args = [expr()]
    if (fn.arity === 2) { expect(','); args.push(expr()) }
    close()
    return { t: 'call', f: fn.f, args }
  }

  const root = expr()
  if (pos !== toks.length) throw new Refusal('syntax')
  return root
}

const fin = (x: unknown): number => (typeof x === 'number' && Number.isFinite(x) ? x : Number.NaN)

function evaluate(n: Node, env: FormulaEnv): number {
  switch (n.t) {
    case 'num': return n.v
    case 'var': return fin(env[n.v])
    case 'cur': return fin(env.current(n.v))
    case 'neg': return fin(-evaluate(n.a, env))
    case 'rand': {
      const scale = n.a === undefined ? 1 : evaluate(n.a, env)
      return fin(scale * fin(env.rand()))
    }
    case 'call': return fin(n.f(evaluate(n.args[0], env), n.args.length > 1 ? evaluate(n.args[1], env) : 0))
    case 'bin': {
      const a = evaluate(n.a, env)
      const b = evaluate(n.b, env)
      switch (n.op) {
        case '+': return fin(a + b)
        case '-': return fin(a - b)
        case '*': return fin(a * b)
        case '/': return fin(a / b)
        case '%': return fin(a % b)
        default: return fin(Math.pow(a, b))
      }
    }
  }
}

/** Compiles one formula. Never throws; every refusal is reported as a reason. */
export function compileFormula(src: string): { ok: true; fn: CompiledFormula } | { ok: false; reason: FormulaRefusal } {
  try {
    if (typeof src !== 'string') return { ok: false, reason: 'syntax' }
    if (src.length > READER_LIMITS.formulaBytes || utf8Length(src, READER_LIMITS.formulaBytes) > READER_LIMITS.formulaBytes)
      return { ok: false, reason: 'too-long' }
    const ast = parse(tokenize(src))
    const fn: CompiledFormula = env => {
      try { return fin(evaluate(ast, env)) } catch { return Number.NaN }
    }
    return { ok: true, fn }
  } catch (err) {
    return { ok: false, reason: err instanceof Refusal ? err.reason : 'syntax' }
  }
}
