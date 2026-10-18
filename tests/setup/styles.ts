/**
 * Injects the compiled, scoped stylesheet into every jsdom test (plan Global
 * Constraints, test split): token values and the --ov-sentinel the viewer's
 * missing-stylesheet check reads come from the real pipeline.
 */
import css from '../../src/styles/officeview.css?inline'

export const STYLE_ELEMENT_ID = 'officeview-test-styles'

if (typeof document !== 'undefined' && !document.getElementById(STYLE_ELEMENT_ID)) {
  const style = document.createElement('style')
  style.id = STYLE_ELEMENT_ID
  style.textContent = css
  document.head.appendChild(style)
}
