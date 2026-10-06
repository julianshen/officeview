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

    // Cell references ($A$1, A1, etc.) or Identifiers
    if (ch === '$' || /[A-Za-z_]/.test(ch)) {
      const rest = str.slice(i)
      const cellMatch = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])/.exec(rest)
      if (cellMatch) {
        const full = cellMatch[0]
        const hasAbsCol = cellMatch[1] === '$'
        const colLetters = cellMatch[2].toUpperCase()
        const hasAbsRow = cellMatch[3] === '$'
        const rowNum = parseInt(cellMatch[4], 10)

        let col = 0
        for (let cIdx = 0; cIdx < colLetters.length; cIdx++) {
          col = col * 26 + (colLetters.charCodeAt(cIdx) - 64)
        }
        col -= 1
        const row = rowNum - 1

        if (col >= 0 && col <= 16383 && row >= 0 && row <= 1048575) {
          tokens.push({
            type: 'cell',
            value: (hasAbsCol ? '$' : '') + colLetters + (hasAbsRow ? '$' : '') + cellMatch[4],
            cellRef: { col, row, absCol: hasAbsCol, absRow: hasAbsRow },
          })
          i += full.length
          continue
        }
      }

      const identMatch = /^([A-Za-z_][A-Za-z0-9_.]*)/.exec(rest)
      if (identMatch) {
        const ident = identMatch[1]
        if (ident.toUpperCase() === 'TRUE') {
          tokens.push({ type: 'boolean', value: 'TRUE' })
        } else if (ident.toUpperCase() === 'FALSE') {
          tokens.push({ type: 'boolean', value: 'FALSE' })
        } else {
          tokens.push({ type: 'ident', value: ident })
        }
        i += ident.length
        continue
      }
    }

    // Unknown single char fallback
    tokens.push({ type: 'op', value: ch })
    i++
  }

  tokens.push({ type: 'eof', value: '' })
  return tokens
}
