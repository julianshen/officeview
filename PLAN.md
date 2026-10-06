# PLAN: XLSX Formula Support in OfficeView

## Objective
Support formula parsing, evaluation, and rendering for `.xlsx` spreadsheets in OfficeView. When spreadsheet cells contain formulas (`<f>` in OpenXML) without precalculated cached values (`<v>`), or when explicit recalculation is triggered (`calcPr/@fullCalcOnLoad="1"` or `f/@ca="1"`), evaluate the formulas at parse/model time and render computed values with faithful styling, formatting, and alignment onto the canvas grid.

---

## Scope & File Boundaries
- `src/xlsx/formula/types.ts` (new): AST nodes, Token, FormulaValue, FormulaError
- `src/xlsx/formula/lexer.ts` (new): Tokenizer handling operators, ranges, `_xlfn.` prefix, string escaping
- `src/xlsx/formula/parser.ts` (new): Recursive descent parser with Excel precedence (`-2^2 = 4`)
- `src/xlsx/formula/functions.ts` (new): Deterministic standard library (Math, Logic, Text, Aggregates)
- `src/xlsx/formula/evaluator.ts` (new): Evaluator with cycle detection (returns 0 on cycle), error propagation, and 15-digit precision
- `src/xlsx/formula/shared.ts` (new): Shared formula offset translation (`t="shared"`)
- `src/xlsx/formula/index.ts` (new): `evaluateWorkbookFormulas` & `evaluateSheetFormulas`
- `src/xlsx/types.ts`: Model additions (`formula?: string`, `sharedFormula?: { si: number; ref?: string }`)
- `src/xlsx/parse.ts`: Parse `f/@t`, `f/@si`, `f/@ref`, `f/@ca`, and integrate workbook formula evaluation
- `src/xlsx/render.ts`: Render computed formula values, errors, and formatting
- `src/testdata/ooxml-builders.ts`: Extend `XlsxCellSpec` to emit `<f>`, shared formulas, and omit `<v>`
- `tests/xlsx-formula.test.ts` (new): Comprehensive test suite covering all phases

---

## Clean Architecture & Key Design Invariants

```
src/xlsx/formula/
├── types.ts          # Pure domain types (AST nodes, Token, FormulaValue, FormulaError)
├── lexer.ts          # Pure tokenizer for Excel formulas (strips _xlfn., bounds length/tokens)
├── parser.ts         # Recursive descent AST parser with Excel operator precedence
├── functions.ts      # Deterministic standard function library (no nondeterministic RAND/NOW)
├── evaluator.ts      # Evaluator with cycle guard (returns 0), first-class error propagation
├── shared.ts         # Shared formula reference re-indexing for fill-down/fill-right
└── index.ts          # evaluateWorkbookFormulas (workbook-level context for Sheet2!A1)
```

### Invariants & Non-Goals
1. **Parse-Time Evaluation Invariant**: Formulas are evaluated at model/parse time, NEVER during `renderSheet`. Viewport culling in `renderSheet` must not prevent off-screen cells from resolving dependencies for on-screen cells.
2. **First-Class Error Values**: Errors (`#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, `#NULL!`) are first-class `FormulaError` values that propagate through operations (`1 + #DIV/0! = #DIV/0!`) and are caught by `IFERROR`.
3. **Excel Parity & Cycle Handling**: Circular references return `0` (matching Excel default behavior) and avoid infinite recursion via an evaluation-path set.
4. **Excel Operator Precedence**: Unary minus binds tighter than exponentiation (`-2^2 = 4`).
5. **15-Digit Floating Precision**: Numerical results are rounded to 15 significant digits to avoid floating-point artifacts like `0.1 + 0.2 = 0.30000000000000004`.
6. **Deliberately Out of Scope**: `xl/calcChain.xml` (recalc-order hint), dynamic array spilling/`LAMBDA`, nondeterministic functions (`RAND`, `NOW`, `TODAY` return `#NAME?`), external workbook links (`[1]Sheet1!A1` return `#REF!`).

---

## TDD Implementation Phases

### Phase 0: Test Fixture & Model Enablement
- [x] Test: `XlsxCellSpec` and `buildXlsx` support formula `<f>` and omitting `<v>`
- [x] Test: `XlsxCellSpec` supports shared formula attributes (`t="shared"`, `si`, `ref`)
- [x] Test: `XlsxCellSpec` supports error cells (`t="e"`)

### Phase 1: Formula Tokenizer (Lexer)
- [ ] Test: Tokenizer handles arithmetic operators, unary minus, and percent (`+`, `-`, `*`, `/`, `^`, `%`)
- [ ] Test: Tokenizer handles string literals with escaped quotes and comparison operators (`=`, `<>`, `<`, `<=`, `>`, `>=`)
- [ ] Test: Tokenizer handles cell references (relative `A1`, absolute `$A$1`, mixed `A$1`, `$A1`)
- [ ] Test: Tokenizer handles range references (`A1:B10`) and cross-sheet references (`Sheet2!A1`, `'My Sheet'!A1:B2`)
- [ ] Test: Tokenizer strips `_xlfn.` function prefix and normalizes function names case-insensitively

### Phase 2: Formula Parser (AST)
- [ ] Test: Parser parses literals and respects Excel unary precedence (`-2^2` parses as `(-2)^2`)
- [ ] Test: Parser respects operator precedence (`^` > `*`, `/` > `+`, `-` > `&` > comparisons)
- [ ] Test: Parser parses function calls with multiple arguments and empty arguments
- [ ] Test: Parser parses nested expressions and parenthesized sub-expressions
- [ ] Test: Parser returns `#NAME?` or syntax error node on malformed input without throwing

### Phase 3: Evaluator & Standard Functions
- [ ] Test: Evaluates arithmetic operations with 15-digit rounding and percent (`10 + 50% = 10.5`)
- [ ] Test: Evaluates string concatenation (`&`) and Excel comparison ordering (number < text < FALSE < TRUE)
- [ ] Test: Propagates error values through operations (`1 + #DIV/0! = #DIV/0!`)
- [ ] Test: Evaluates Math functions (SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, ABS, ROUND, INT, MOD, PRODUCT)
- [ ] Test: Evaluates Logic functions with short-circuiting (IF, AND, OR, NOT, IFERROR)
- [ ] Test: Evaluates Text functions (CONCAT, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER)
- [ ] Test: Handles blank cells correctly (0 in math, "" in concat, ignored in SUM)
- [ ] Test: Detects circular references and returns 0 without stack overflow

### Phase 4: Shared Formulas & Workbook Context
- [ ] Test: Translates shared formula relative references by row/col offset (`si` master to dependent cells)
- [ ] Test: Resolves cross-sheet references (`Sheet2!A1`) using workbook-level context
- [ ] Test: Evaluates multi-cell dependency chains across rows and sheets in correct order
- [ ] Test: Preserves cached `<v>` unless missing, `ca="1"`, or `fullCalcOnLoad="1"`

### Phase 5: Integration & Canvas Rendering
- [ ] Test: `parseXlsx` evaluates formula cells when `<v>` is absent
- [ ] Test: Viewport culling does not break off-screen formula dependencies
- [ ] Test: `renderSheet` renders calculated formula cell values onto canvas with correct alignment and styling
- [ ] Test: `renderSheet` applies number format (`numFmtId`) to formula results
- [ ] Test: `renderSheet` renders formula error strings (#DIV/0!) centered per Excel convention
- [ ] Test: End-to-end fixture test parsing and rendering an XLSX file with missing `<v>` formulas
