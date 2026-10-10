import type { AxisEndpoint, Token, WholeColSyntax, WholeRowSyntax } from './types'

const MAX_FORMULA_LENGTH = 8192

interface ParsedCellCoord {
  col: number
  row: number
  absCol: boolean
  absRow: boolean
  length: number
  text: string
}

function colLettersToIndex(letters: string): number {
  let col = 0
  for (let cIdx = 0; cIdx < letters.length; cIdx++) {
    col = col * 26 + (letters.charCodeAt(cIdx) - 64)
  }
  return col - 1
}

interface WholeAxisMatch {
  sheet?: string
  kind: 'col' | 'row'
  from: AxisEndpoint
  to: AxisEndpoint
  text: string
  length: number
}

/**
 * Whole-column (A:C, $A:C) / whole-row (1:3, $1:3) references with optional
 * sheet qualification and per-endpoint absolute anchors. Returns null unless
 * the axis is complete and in bounds; plain cells/ranges never match because
 * the colon must directly follow letters (cols) or digits (rows).
 */
function parseWholeAxis(s: string): WholeAxisMatch | null {
  let sheet: string | undefined
  let rest = s
  if (rest[0] === "'") {
    let j = 1
    let name = ''
    let closed = false
    while (j < rest.length) {
      if (rest[j] === "'") {
        if (rest[j + 1] === "'") { name += "'"; j += 2 } else { j++; closed = true; break }
      } else { name += rest[j]; j++ }
    }
    if (!closed || rest[j] !== '!') return null
    sheet = name
    rest = rest.slice(j + 1)
  } else {
    const q = /^([\p{L}_][\p{L}\p{N}_.]*)!/u.exec(rest)
    if (q) {
      sheet = q[1]
      rest = rest.slice(q[0].length)
    }
  }
  const consumedPrefix = s.length - rest.length
  let m = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\p{L}\p{N}_.!(])/u.exec(rest)
  if (m) {
    const from = colLettersToIndex(m[2].toUpperCase())
    const to = colLettersToIndex(m[4].toUpperCase())
    if (from < 0 || from > 16383 || to < 0 || to > 16383) return null
    return {
      sheet,
      kind: 'col',
      from: { index: from, absolute: m[1] === '$' },
      to: { index: to, absolute: m[3] === '$' },
      text: s.slice(0, consumedPrefix + m[0].length),
      length: consumedPrefix + m[0].length,
    }
  }
  m = /^(\$?)(\d+):(\$?)(\d+)(?![\p{L}\p{N}_.!(])/u.exec(rest)
  if (m) {
    const from = parseInt(m[2], 10)
    const to = parseInt(m[4], 10)
    if (from < 1 || from > 1048576 || to < 1 || to > 1048576) return null
    return {
      sheet,
      kind: 'row',
      from: { index: from - 1, absolute: m[1] === '$' },
      to: { index: to - 1, absolute: m[3] === '$' },
      text: s.slice(0, consumedPrefix + m[0].length),
      length: consumedPrefix + m[0].length,
    }
  }
  return null
}

/** End-exclusive index just past the balanced closing bracket, or null.
 * Escape-aware (group 6): `']` / `'[` are literal characters inside the body,
 * not bracket open/close tokens — they cannot alter balance. */
function scanTableBrackets(s: string, openIdx: number): number | null {
  let depth = 0
  for (let j = openIdx; j < s.length; j++) {
    const ch = s[j]
    if (ch === "'") {
      if (s[j + 1] === "'" || s[j + 1] === '[' || s[j + 1] === ']') {
        j++ // skipped-with-next: escaped literal bracket
        continue
      }
      continue
    }
    if (ch === '[') depth++
    else if (ch === ']') {
      depth--
      if (depth === 0) return j + 1
    }
  }
  return null
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
    text: m[0],
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
      { type: 'error', value: `Formula exceeds maximum length of ${MAX_FORMULA_LENGTH} characters`, start: 0 },
      { type: 'eof', value: '', start: str.length },
    ]
  }

  const tokens: Token[] = []
  let i = 0
  const len = str.length
  // Whitespace trivia is recorded on the following token so the parser can
  // recognize reference intersection operands; the token stream itself is
  // unchanged apart from the additive spaceBefore flag.
  let pendingSpace = false
  const emit = (tok: Token): void => {
    if (pendingSpace) tok.spaceBefore = true
    pendingSpace = false
    tokens.push(tok)
  }

  while (i < len) {
    const ch = str[i]
    const tokenStart = i

    if (/\s/.test(ch)) {
      let j = i + 1
      while (j < len && /\s/.test(str[j])) j++
      pendingSpace = true
      i = j
      continue
    }

    // Whole-column / whole-row references (A:C, $A:C, 1:3, Sheet1!A:C).
    // Plain cells/ranges never match: the colon must directly follow column
    // letters (cols) or row digits (rows).
    const axis = (ch === '$' || ch === "'" || /[\p{L}\p{N}]/u.test(ch)) ? parseWholeAxis(str.slice(i)) : null
    if (axis) {
      if (axis.kind === 'col') {
        const ref: WholeColSyntax = {
          ...(axis.sheet !== undefined ? { sheet: axis.sheet } : {}),
          from: axis.from,
          to: axis.to,
        }
        emit({ type: 'wholeCol', value: axis.text, start: tokenStart, sheet: axis.sheet, wholeColRef: ref })
      } else {
        const ref: WholeRowSyntax = {
          ...(axis.sheet !== undefined ? { sheet: axis.sheet } : {}),
          from: axis.from,
          to: axis.to,
        }
        emit({ type: 'wholeRow', value: axis.text, start: tokenStart, sheet: axis.sheet, wholeRowRef: ref })
      }
      i += axis.length
      continue
    }

    // Excel error literals (#DIV/0!, #REF!, #N/A, #VALUE!, #NAME?, #NUM!, #NULL!)
    if (ch === '#') {
      const rest = str.slice(i)
      const errMatch = /^#(DIV\/0!|REF!|N\/A|VALUE!|NAME\?|NUM!|NULL!|SPILL!|CALC!)/i.exec(rest)
      if (errMatch) {
        emit({ type: 'error', value: errMatch[0].toUpperCase(), start: tokenStart })
        i += errMatch[0].length
        continue
      }
      emit({ type: 'error', value: ch, start: tokenStart })
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
      emit({ type: 'number', value: numStr, start: tokenStart, numValue: numVal })
      continue
    }

    // Operators and delimiters
    if (ch === '(') {
      emit({ type: 'lparen', value: '(', start: tokenStart })
      i++
      continue
    }
    if (ch === ')') {
      emit({ type: 'rparen', value: ')', start: tokenStart })
      i++
      continue
    }
    if (ch === ',') {
      emit({ type: 'comma', value: ',', start: tokenStart })
      i++
      continue
    }
    if (ch === ':') {
      emit({ type: 'colon', value: ':', start: tokenStart })
      i++
      continue
    }
    // Array-literal delimiters: { } open/close the literal, ; starts a new row.
    if (ch === '{') {
      emit({ type: 'lbrace', value: '{', start: tokenStart })
      i++
      continue
    }
    if (ch === '}') {
      emit({ type: 'rbrace', value: '}', start: tokenStart })
      i++
      continue
    }
    if (ch === ';') {
      emit({ type: 'semicolon', value: ';', start: tokenStart })
      i++
      continue
    }

    // Implicit-intersection prefix (@A1) only where a reference operand can
    // follow; a lone @ stays a fail-closed error token as before.
    if (ch === '@') {
      let k = i + 1
      while (k < len && /\s/.test(str[k])) k++
      const nc = str[k]
      if (nc !== undefined && (nc === '(' || nc === '[' || nc === '{' || nc === "'" || nc === '$' || /[\p{L}_\p{N}]/u.test(nc))) {
        emit({ type: 'at', value: '@', start: tokenStart })
      } else {
        emit({ type: 'error', value: '@', start: tokenStart })
      }
      i++
      continue
    }

    // Unqualified structured reference ([...] in the current table context).
    if (ch === '[') {
      const end = scanTableBrackets(str, i)
      if (end !== null) {
        emit({ type: 'table', value: str.slice(i, end), start: tokenStart })
        i = end
        continue
      }
      emit({ type: 'error', value: ch, start: tokenStart })
      i++
      continue
    }

    // Arithmetic operators & percent
    if (ch === '+' || ch === '-' || ch === '*' || ch === '/' || ch === '^' || ch === '%') {
      emit({ type: 'op', value: ch, start: tokenStart })
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
        const preview = val.length > 32 ? val.slice(0, 32) + '...' : val
        emit({ type: 'error', value: `Unterminated string literal: "${preview}`, start: tokenStart })
      } else {
        emit({ type: 'string', value: val, start: tokenStart })
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
        const afterBangQ = str.slice(sIdx)
        const nameAfterQ = /^([\p{L}_][\p{L}\p{N}_.]*)/u.exec(afterBangQ)
        if (nameAfterQ && parseCellCoord(afterBangQ) === null) {
          emit({ type: 'ident', value: nameAfterQ[1].toUpperCase(), raw: nameAfterQ[1], start: tokenStart, sheet: sheetName })
          i = sIdx + nameAfterQ[0].length
          continue
        }
        const c1 = parseCellCoord(str.slice(sIdx))
        if (c1) {
          let consumed = sIdx + c1.length
          if (consumed < len && str[consumed] === ':') {
            const c2 = parseCellCoord(str.slice(consumed + 1))
            if (c2) {
              consumed += 1 + c2.length
              const fullText = str.slice(i, consumed)
              emit({
                type: 'range',
                value: fullText,
                start: tokenStart,
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
          emit({
            type: 'cell',
            value: fullText,
            start: tokenStart,
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
      emit({ type: 'error', value: str.slice(i, sIdx), start: tokenStart })
      i = sIdx
      continue
    }

    // Comparison operators and concat
    if (ch === '<') {
      if (i + 1 < len && str[i + 1] === '>') {
        emit({ type: 'op', value: '<>', start: tokenStart })
        i += 2
        continue
      }
      if (i + 1 < len && str[i + 1] === '=') {
        emit({ type: 'op', value: '<=', start: tokenStart })
        i += 2
        continue
      }
      emit({ type: 'op', value: '<', start: tokenStart })
      i++
      continue
    }
    if (ch === '>') {
      if (i + 1 < len && str[i + 1] === '=') {
        emit({ type: 'op', value: '>=', start: tokenStart })
        i += 2
        continue
      }
      emit({ type: 'op', value: '>', start: tokenStart })
      i++
      continue
    }
    if (ch === '=') {
      emit({ type: 'op', value: '=', start: tokenStart })
      i++
      continue
    }
    if (ch === '&') {
      emit({ type: 'op', value: '&', start: tokenStart })
      i++
      continue
    }

    // Cell references, Range references, Cross-sheet references, or Identifiers
    // Note: [\p{L}_] matches Unicode letters including CJK for sheet names and identifiers
    if (ch === '$' || /[\p{L}_]/u.test(ch)) {
      const rest = str.slice(i)

      // Unquoted cross-sheet reference: Sheet2!A1 or 工作表1!A1:B10
      const sheetMatch = /^([\p{L}_][\p{L}\p{N}_.]*)!/u.exec(rest)
      if (sheetMatch) {
        const sheetName = sheetMatch[1]
        const afterBang = rest.slice(sheetMatch[0].length)
        // Sheet-qualified defined name: Sheet1!MyName (qualification only;
        // scope resolution happens at evaluation, never in grammar).
        const nameAfter = /^([\p{L}_][\p{L}\p{N}_.]*)/u.exec(afterBang)
        if (nameAfter && parseCellCoord(afterBang) === null) {
          const upper = nameAfter[1].toUpperCase()
          emit({ type: 'ident', value: upper, raw: nameAfter[1], start: tokenStart, sheet: sheetName })
          i += sheetMatch[0].length + nameAfter[0].length
          continue
        }
        const c1 = parseCellCoord(afterBang)
        if (c1) {
          let consumed = sheetMatch[0].length + c1.length
          if (rest[consumed] === ':') {
            const c2 = parseCellCoord(rest.slice(consumed + 1))
            if (c2) {
              consumed += 1 + c2.length
              const fullText = rest.slice(0, consumed)
              emit({
                type: 'range',
                value: fullText,
                start: tokenStart,
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
          emit({
            type: 'cell',
            value: fullText,
            start: tokenStart,
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

      // Local cell or range reference: A1 or A1:B10 (columns are strictly ASCII A-Z)
      const c1 = parseCellCoord(rest)
      if (c1) {
        let consumed = c1.length
        if (rest[consumed] === ':') {
          const c2 = parseCellCoord(rest.slice(consumed + 1))
          if (c2) {
            // 3D run guard: S1:S2!A1 is a sheet run, never a range. When the
            // second coordinate is followed by '!'+grid (cell or whole axis), leave the colon for
            // the parser's sheet-run rule instead of forming a range token.
            const afterC2 = rest.slice(consumed + 1 + c2.length)
            if (!(afterC2[0] === '!' && (parseCellCoord(afterC2.slice(1)) !== null || parseWholeAxis(afterC2.slice(1)) !== null))) {
              consumed += 1 + c2.length
              const fullText = rest.slice(0, consumed)
              emit({
                type: 'range',
                value: fullText,
                start: tokenStart,
                rangeRef: {
                  from: { col: c1.col, row: c1.row, absCol: c1.absCol, absRow: c1.absRow },
                  to: { col: c2.col, row: c2.row, absCol: c2.absCol, absRow: c2.absRow },
                },
              })
              i += consumed
              continue
            }
          }
        }
        emit({
          type: 'cell',
          value: c1.text,
          start: tokenStart,
          cellRef: { col: c1.col, row: c1.row, absCol: c1.absCol, absRow: c1.absRow },
        })
        i += c1.length
        continue
      }

      const identMatch = /^([\p{L}_][\p{L}\p{N}_.]*)/u.exec(rest)
      if (identMatch) {
        const fullIdent = identMatch[1]        // Qualified structured reference: Table1[...] (whitespace between the
        // table name and its brackets is allowed). The bracket text is kept
        // verbatim; items/columns decode at parse time from the raw source.
        let bIdx = fullIdent.length
        while (bIdx < rest.length && /\s/.test(rest[bIdx])) bIdx++
        if (rest[bIdx] === '[') {
          const end = scanTableBrackets(rest, bIdx)
          if (end !== null) {
            const raw = rest.slice(0, end)
            emit({ type: 'table', value: raw, start: tokenStart, tableName: fullIdent })
            i += end
            continue
          }
        }
        const cleanIdent = rest[bIdx] === '(' ? fullIdent.replace(/^_xlfn\.(_xlws\.)?/i, '') : fullIdent
        const upper = cleanIdent.toUpperCase()
        if (upper === 'TRUE') {
          emit({ type: 'boolean', value: 'TRUE', start: tokenStart })
        } else if (upper === 'FALSE') {
          emit({ type: 'boolean', value: 'FALSE', start: tokenStart })
        } else {
          emit({ type: 'ident', value: upper, start: tokenStart, raw: fullIdent })
        }
        i += fullIdent.length
        continue
      }
    }

    // Unknown single char fallback - fail closed
    emit({ type: 'error', value: ch, start: tokenStart })
    i++
  }

  // Spill suffix: an immediate '#' after a cell/range/whole-axis reference
  // belongs to the reference (A1#); the '#' token is consumed with it. A
  // spaced '#' or an error literal (#REF!) stays a separate token.
  for (let k = tokens.length - 1; k >= 0; k--) {
    const t = tokens[k]
    if (t.type === 'cell' || t.type === 'range' || t.type === 'wholeCol' || t.type === 'wholeRow') {
      const n = tokens[k + 1]
      if (n && n.type === 'error' && n.value === '#' && n.start === t.start + t.value.length) {
        t.type = 'spill'
        t.value += '#'
        tokens.splice(k + 1, 1)
      }
    }
  }

  emit({ type: 'eof', value: '', start: len })
  return tokens
}
