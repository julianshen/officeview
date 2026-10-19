/**
 * The one registration of officeview's CSS pipeline (plan Global Constraints,
 * "Scoping stage"): `@tailwindcss/vite` generates the CSS, then Vite's
 * `css.postcss.plugins` run ovCssScope on it before the CSS is consumed.
 * vite.config.ts and vitest.config.ts spread this object; the harness config
 * inherits it through mergeConfig and must not add a second copy.
 */
import tailwindcss from '@tailwindcss/vite'
import type { CSSOptions, PluginOption } from 'vite'
import { ovCssScope } from './postcss-ov-scope'

// Vite types PostCSS plugins against its own nested postcss copy; the plugin
// is the same runtime shape, so only the type is bridged.
type VitePostcssPlugin = NonNullable<Exclude<CSSOptions['postcss'], string | undefined>['plugins']>[number]

export const ovCss: { plugins: PluginOption[]; css: CSSOptions } = {
  plugins: [tailwindcss()],
  css: { postcss: { plugins: [ovCssScope() as unknown as VitePostcssPlugin] } },
}
