import type { FontResolver } from './register'

export type CjkFallbackRegion = 'auto' | 'sc' | 'tc' | 'hk' | 'jp' | 'kr'
export interface FallbackFontsOptions {
  /** Explicit ordered fallback chain, applied before the generic. Wins over cjkFallback. */
  fallbackChain?: string[]
  /** Regional CJK families heading the chain. 'auto'/absent adds none (host pins explicitly). */
  cjkFallback?: CjkFallbackRegion
}
/** Regional first-fallback families, most-preferred first (system names; first available wins). */
export const REGIONAL_FALLBACKS: Record<Exclude<CjkFallbackRegion, 'auto'>, string[]> = {
  sc: ['PingFang SC', 'Microsoft YaHei', 'Noto Sans CJK SC'],
  tc: ['PingFang TC', 'Microsoft JhengHei', 'Noto Sans CJK TC'],
  hk: ['PingFang HK', 'Microsoft JhengHei', 'Noto Sans CJK HK'],
  jp: ['Hiragino Kaku Gothic ProN', 'Meiryo', 'Yu Gothic', 'Noto Sans CJK JP'],
  kr: ['Apple SD Gothic Neo', 'Malgun Gothic', 'Noto Sans CJK KR'],
}
const GENERICS = new Set(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui'])
// Serif set mirrors src/docx/styles.ts fontFamilyCss (kept in sync manually;
// fallback.ts must not import format-specific styles).
const SERIF = /^(Cambria(?: Math)?|Constantia|Times New Roman|Georgia|Garamond|Palatino(?: Linotype)?|Book Antiqua)$/i
const quote = (family: string): string => `"${family.replace(/["\\]/g, '\\$&')}"`
const chainOf = (options: FallbackFontsOptions | undefined): string[] => {
  if (options?.fallbackChain?.length) return options.fallbackChain
  if (options?.cjkFallback && options.cjkFallback !== 'auto') return REGIONAL_FALLBACKS[options.cjkFallback]
  return []
}
/**
 * Compose an explicit, ordered fallback chain onto any resolver. Bare families
 * become `"Fam", chain..., generic`; full stacks splice the chain before the
 * final generic. With no options the resolver passes through byte-identical,
 * so default rendering (and goldens) never move.
 */
export function withFallbackFonts(resolve: FontResolver, options?: FallbackFontsOptions): FontResolver {
  const chain = chainOf(options).map(quote)
  if (!chain.length) return resolve
  return (family: string): string => {
    const raw = resolve(family)
    // A trailing generic marks a full stack (quoted or not): splice the chain
    // before it and leave the head exactly as the resolver authored it.
    const tail = /,\s*([A-Za-z-]+)\s*$/.exec(raw)
    if (tail && GENERICS.has(tail[1].toLowerCase())) {
      return `${raw.slice(0, raw.length - tail[0].length)}, ${chain.join(', ')}, ${tail[1]}`
    }
    return `${quote(raw)}, ${chain.join(', ')}, ${SERIF.test(raw) ? 'serif' : 'sans-serif'}`
  }
}
