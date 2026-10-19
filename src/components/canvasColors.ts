/** Canvas UI colours only: paper, authored content and gridlines never use these.
 * Lazy in both viewers. Ancestor-only stylesheet changes are sampled on the next
 * theme change or remount; root style/class/theme changes are observed directly.
 */
import { useEffect, useState, type RefObject } from 'react'
import { CANVAS_TOKENS, type CanvasToken } from '../styles/tokens'
import type { ThemeMode } from './theme'
export type { CanvasToken }
export type CanvasColors = Partial<Record<CanvasToken, string>>
export function readCanvasColors(root: Element): CanvasColors {
  const style = root.ownerDocument.defaultView!.getComputedStyle(root)
  return Object.fromEntries(CANVAS_TOKENS.flatMap(token => {
    const value = style.getPropertyValue(`--ov-color-${token}`).trim()
    return value ? [[token, value]] : []
  }))
}
export function paletteKey(colors: CanvasColors): string {
  return JSON.stringify(CANVAS_TOKENS.map(token => colors[token] ?? ''))
}
export function useCanvasColors(root: RefObject<Element | null>, mode: ThemeMode): { colors: CanvasColors; key: string } {
  const [palette, setPalette] = useState(() => ({colors: {} as CanvasColors, key: paletteKey({})}))
  useEffect(() => {
    const el = root.current
    if (!el) return
    const update = () => {
      const colors = readCanvasColors(el), key = paletteKey(colors)
      setPalette(previous => previous.key === key ? previous : {colors, key})
    }
    update()
    const media = el.ownerDocument.defaultView?.matchMedia?.('(prefers-color-scheme: dark)')
    media?.addEventListener('change', update)
    const observer = new MutationObserver(update)
    observer.observe(el, {attributes:true, attributeFilter:['style','class','data-ov-theme']})
    return () => { observer.disconnect(); media?.removeEventListener('change', update) }
  }, [root, mode])
  return palette
}
