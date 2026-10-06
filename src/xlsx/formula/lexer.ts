import type { Token } from './types'

const MAX_FORMULA_LENGTH = 8192

interface ParsedCellCoord {
  col: number
  row: number
  absCol: boolean
  absRow: boolean
  length: number
  text: string
}

function parseCellCoord(s: string): ParsedCellCoord | null {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])(?!\s*\()/.exec(s)
  if (!m) return null
  const hasAbsCol = m[1] === '$'
  const colLetters = m[2].toUpperCase()
  const hasAbsRow = m[3] === '$'
  const rowNum = parseInt(m[4], 10)

  let col = 0
  for (let cIdx = 0; cIdx < colLetters.length; cIdx++) {
    col = col * 26 + (colLetters.charCodeAt(cIdx) - 64)
  }
  col -= 1
  const row = rowNum - 1

  if (col < 0 || col > 16383 || row < 0 || row > 1048575) {
    return null
  }

  return {
    col,
    row,
    absCol: hasAbsCol,
    absRow: hasAbsRow,
    length: m[0].length,
    text: (hasAbsCol ? '$' : '') + colLetters + (hasAbsRow ? '$' : '') + m[4],
  }
}

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
    return [
      { type: 'error', value: `Formula exceeds maximum length of ${MAX_FORMULA_LENGTH} characters` },
      { type: 'eof', value: '' },
    ]
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

    // Excel error literals (#DIV/0!, #REF!, #N/A, #VALUE!, #NAME?, #NUM!, #NULL!, #ERROR)
    if (ch === '#') {
      const rest = str.slice(i)
      const errMatch = /^#(DIV\/0!|REF!|N\/A|VALUE!|NAME\?|NUM!|NULL!|ERROR)/i.exec(rest)
      if (errMatch) {
        tokens.push({ type: 'error', value: errMatch[0].toUpperCase() })
        i += errMatch[0].length
        continue
      }
      tokens.push({ type: 'error', value: ch })
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
      let closed = false
      while (i < len) {
        if (str[i] === '"') {
          if (i + 1 < len && str[i + 1] === '"') {
            val += '"'
            i += 2
          } else {
            i++ // skip closing quote
            closed = true
            break
          }
        } else {
          val += str[i]
          i++
        }
      }
      if (!closed) {
        tokens.push({ type: 'error', value: `Unterminated string literal: "${val}` })
      } else {
        tokens.push({ type: 'string', value: val })
      }
      continue
    }

    // Quoted sheet reference: 'Sheet Name'!A1 or 'Sheet Name'!A1:B10
    if (ch === '\'') {
      let sIdx = i + 1
      let sheetName = ''
      let closed = false
      while (sIdx < len) {
        if (str[sIdx] === '\'') {
          if (sIdx + 1 < len && str[sIdx + 1] === '\'') {
            sheetName += '\''
            sIdx += 2
          } else {
            sIdx++ // skip closing quote
            closed = true
            break
          }
        } else {
          sheetName += str[sIdx]
          sIdx++
        }
      }
      if (closed && sIdx < len && str[sIdx] === '!') {
        sIdx++ // skip '!'
        const c1 = parseCellCoord(str.slice(sIdx))
        if (c1) {
          let consumed = sIdx + c1.length
          if (consumed < len && str[consumed] === ':') {
            const c2 = parseCellCoord(str.slice(consumed + 1))
            if (c2) {
              consumed += 1 + c2.length
              const fullText = str.slice(i, consumed)
              tokens.push({
                type: 'range',
                value: fullText,
                sheet: sheetName,
                rangeRef: {
                  sheet: sheetName,
                  from: { sheet: sheetName, col: c1.col, row: c1.row, absCol: c1.absCol, absRow: c1.absRow },
                  to: { sheet: sheetName, col: c2.col, row: c2.row, absCol: c2.absCol, absRow: c2.absRow },
                },
              })
              i = consumed
              continue
            }
          }
          const fullText = str.slice(i, consumed)
          tokens.push({
            type: 'cell',
            value: fullText,
            sheet: sheetName,
            cellRef: {
              sheet: sheetName,
              col: c1.col,
              row: c1.row,
              absCol: c1.absCol,
              absRow: c1.absRow,
            },
          })
          i = consumed
          continue
        }
      }
      tokens.push({ type: 'error', value: str.slice(i, sIdx) })
      i = sIdx
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

    // Cell references, Range references, Cross-sheet references, or Identifiers
    if (ch === '$' || /[A-Za-z_]/.test(ch)) {
      const rest = str.slice(i)

      // Unquoted cross-sheet reference: Sheet2!A1 or Sheet2!A1:B10
      const sheetMatch = /^([A-Za-z_][A-Za-z0-9_.]*)!/.exec(rest)
      if (sheetMatch) {
        const sheetName = sheetMatch[1]
        const afterBang = rest.slice(sheetMatch[0].length)
        const c1 = parseCellCoord(afterBang)
        if (c1) {
          let consumed = sheetMatch[0].length + c1.length
          if (rest[consumed] === ':') {
            const c2 = parseCellCoord(rest.slice(consumed + 1))
            if (c2) {
              consumed += 1 + c2.length
              const fullText = rest.slice(0, consumed)
              tokens.push({
                type: 'range',
                value: fullText,
                sheet: sheetName,
                rangeRef: {
                  sheet: sheetName,
                  from: { sheet: sheetName, col: c1.col, row: c1.row, absCol: c1.absCol, absRow: c1.absRow },
                  to: { sheet: sheetName, col: c2.col, row: c2.row, absCol: c2.absCol, absRow: c2.absRow },
                },
              })
              i += consumed
              continue
            }
          }
          const fullText = rest.slice(0, consumed)
          tokens.push({
            type: 'cell',
            value: fullText,
            sheet: sheetName,
            cellRef: {
              sheet: sheetName,
              col: c1.col,
              row: c1.row,
              absCol: c1.absCol,
              absRow: c1.absRow,
            },
          })
          i += consumed
          continue
        }
      }

      // Local cell or range reference: A1 or A1:B10
      const c1 = parseCellCoord(rest)
      if (c1) {
        let consumed = c1.length
        if (rest[consumed] === ':') {
          const c2 = parseCellCoord(rest.slice(consumed + 1))
          if (c2) {
            consumed += 1 + c2.length
            const fullText = rest.slice(0, consumed)
            tokens.push({
              type: 'range',
              value: fullText,
              rangeRef: {
                from: { col: c1.col, row: c1.row, absCol: c1.absCol, absRow: c1.absRow },
                to: { col: c2.col, row: c2.row, absCol: c2.absCol, absRow: c2.absRow },
              },
            })
            i += consumed
            continue
          }
        }
        tokens.push({
          type: 'cell',
          value: c1.text,
          cellRef: { col: c1.col, row: c1.row, absCol: c1.absCol, absRow: c1.absRow },
        })
        i += c1.length
        continue
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

    // Unknown single char fallback - fail closed
    tokens.push({ type: 'error', value: ch })
    i++
  }

  tokens.push({ type: 'eof', value: '' })
  return tokens
}
