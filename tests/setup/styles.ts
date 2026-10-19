/**
 * Injects the compiled, scoped stylesheet's theme contract into every jsdom
 * test (plan Global Constraints, test split): token values and the
 * --ov-sentinel the viewer's missing-stylesheet check reads come from the
 * real pipeline.
 *
 * Only the contract, by default: the token rules (light, dark, system) and
 * the print guard — the rules product code keys off when the sentinel is
 * present (canvasColors reads the tokens; PagedChrome leaves a denied print
 * to the stylesheet's @media print rule). jsdom resolves getComputedStyle by
 * matching every rule against the element, and Testing Library calls it for
 * every candidate of every role query (accessible names, isInaccessible), so
 * the ~150 utility and base rules made the whole suite about 1.4x slower
 * (full CI-equivalent run in a Linux container: 449 s vs 320 s) and the
 * longest viewer tests ~2.5x slower (policy eviction: 4.7 s vs 1.9 s;
 * getComputedStyle time 3,066 ms vs 428 ms), past vitest's 5 s timeout on
 * GitHub's runner. A file that asserts on utility or base-rule styles calls
 * installFullStylesheet() at module scope.
 */
import css from '../../src/styles/officeview.css?inline'

export const STYLE_ELEMENT_ID = 'officeview-test-styles'

const isContract = (rule: CSSRule): boolean => rule instanceof CSSMediaRule
  ? rule.media.mediaText === 'print' || Array.from(rule.cssRules).every(isContract)
  : rule instanceof CSSStyleRule && rule.selectorText.startsWith(':where([data-ov-theme-root]')

if (typeof document !== 'undefined' && !document.getElementById(STYLE_ELEMENT_ID)) {
  const style = document.createElement('style')
  style.id = STYLE_ELEMENT_ID
  style.textContent = css
  document.head.appendChild(style)
  const sheet = style.sheet!
  for (let i = sheet.cssRules.length - 1; i >= 0; i--) if (!isContract(sheet.cssRules[i])) sheet.deleteRule(i)
}

/** The whole compiled stylesheet (utilities and base rules too), for this file. */
export function installFullStylesheet(): void {
  const style = document.getElementById(STYLE_ELEMENT_ID)
  if (style) style.textContent = css
}
