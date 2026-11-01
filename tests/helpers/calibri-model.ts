/**
 * A Calibri-like text measure for tests (Task 29 autofit): approximate Calibri advance widths in em at 11 pt, a bold
 * factor, and a CJK advance for the fallback font Excel for Mac uses. Calibrated so that fitting each Task 24
 * start file reproduces Excel's saved `<col>` widths (423 of 433 columns within one Excel pixel; the 10 others are
 * widths a setup file kept from an earlier layout, not fits). It stands in for the viewer's canvas measurement,
 * which uses whatever font the browser has.
 */
const EM: Record<string, number> = { ' ': .226, '.': .252, '-': .306, '(': .303, ')': .303, ',': .250, '%': .715, '/': .386, ':': .268,
  '0': .507, '1': .507, '2': .507, '3': .507, '4': .507, '5': .507, '6': .507, '7': .507, '8': .507, '9': .507,
  a: .479, b: .525, c: .423, d: .525, e: .498, f: .305, g: .471, h: .525, i: .229, j: .239, k: .455, l: .229, m: .799, n: .525, o: .527, p: .525, q: .525, r: .349, s: .391, t: .335, u: .525, v: .452, w: .715, x: .433, y: .453, z: .395,
  A: .579, B: .544, C: .533, D: .615, E: .488, F: .459, G: .630, H: .623, I: .252, J: .319, K: .520, L: .420, M: .855, N: .646, O: .662, P: .517, Q: .673, R: .543, S: .459, T: .487, U: .653, V: .567, W: .890, X: .519, Y: .487, Z: .468 }
const CJK_EM = 1.03, BOLD = 1.03, OTHER_EM = .5
export function calibriTextPt(text: string, bold: boolean, sizePt = 11): number {
  let em = 0
  for (const ch of text) em += EM[ch] ?? (ch.charCodeAt(0) > 0x2e00 ? CJK_EM : OTHER_EM)
  return em * sizePt * (bold ? BOLD : 1)
}
