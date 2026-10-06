import type { Token } from './types'

const MAX_FORMULA_LENGTH = 8192

/**
 * Tokenize an Excel formula string into tokens.
 * Strips leading '=', skips whitespace, and bounds input length.
 */
export function tokenize(input: string): Token[] {
  let str = input.trim()
  if (str.startsWith('=')) {
    str = str.slice(1).trim()
  }

  if (str.length > MAX_FORMULA_LENGTH) {
    str = str.slice(0, MAX_FORMULA_LENGTH)
  }

  const tokens: Token[] = []
  let i = 0
  const len = str.length

  while (i < len) {
    const ch = str[i]

    // Skip whitespace
    if (/\s/.test(ch)) {
      i++
      continue
    }

    // Number literal: digits with optional decimal point and exponent
    if (/[0-9]/.test(ch) || (ch === '.' && i + 1 < len && /[0-9]/.test(str[i + 1]))) {
      let start = i
      while (i < len && /[0-9]/.test(str[i])) i++
      if (i < len && str[i] === '.') {
        i++
        while (i < len && /[0-9]/.test(str[i])) i++
      }
      if (i < len && (str[i] === 'e' || str[i] === 'E')) {
        let eIdx = i
        i++
        if (i < len && (str[i] === '+' || str[i] === '-')) i++
        const expStart = i
        while (i < len && /[0-9]/.test(str[i])) i++
        if (i === expStart) {
          // invalid exponent, roll back
          i = eIdx
        }
      }
      const numStr = str.slice(start, i)
      const numVal = parseFloat(numStr)
      tokens.push({ type: 'number', value: numStr, numValue: numVal })
      continue
    }

    // Operators and delimiters
    if (ch === '(') {
      tokens.push({ type: 'lparen', value: '(' })
      i++
      continue
    }
    if (ch === ')') {
      tokens.push({ type: 'rparen', value: ')' })
      i++
      continue
    }
    if (ch === ',') {
      tokens.push({ type: 'comma', value: ',' })
      i++
      continue
    }
    if (ch === ':') {
      tokens.push({ type: 'colon', value: ':' })
      i++
      continue
    }

    // Arithmetic operators & percent
    if (ch === '+' || ch === '-' || ch === '*' || ch === '/' || ch === '^' || ch === '%') {
      tokens.push({ type: 'op', value: ch })
      i++
      continue
    }

    // String literal: double-quoted, "" is escaped double quote
    if (ch === '"') {
      i++ // skip opening quote
      let val = ''
      while (i < len) {
        if (str[i] === '"') {
          if (i + 1 < len && str[i + 1] === '"') {
            val += '"'
            i += 2
          } else {
            i++ // skip closing quote
            break
          }
        } else {
          val += str[i]
          i++
        }
      }
      tokens.push({ type: 'string', value: val })
      continue
    }

    // Comparison operators and concat
    if (ch === '<') {
      if (i + 1 < len && str[i + 1] === '>') {
        tokens.push({ type: 'op', value: '<>' })
        i += 2
        continue
      }
      if (i + 1 < len && str[i + 1] === '=') {
        tokens.push({ type: 'op', value: '<=' })
        i += 2
        continue
      }
      tokens.push({ type: 'op', value: '<' })
      i++
      continue
    }
    if (ch === '>') {
      if (i + 1 < len && str[i + 1] === '=') {
        tokens.push({ type: 'op', value: '>=' })
        i += 2
        continue
      }
      tokens.push({ type: 'op', value: '>' })
      i++
      continue
    }
    if (ch === '=') {
      tokens.push({ type: 'op', value: '=' })
      i++
      continue
    }
    if (ch === '&') {
      tokens.push({ type: 'op', value: '&' })
      i++
      continue
    }

    // Unknown single char fallback
    tokens.push({ type: 'op', value: ch })
    i++
  }

  tokens.push({ type: 'eof', value: '' })
  return tokens
}
