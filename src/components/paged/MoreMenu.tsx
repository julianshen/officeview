/**
 * The compact toolbar's More menu (spec §6.5 keyboard contract): it opens
 * focused on its first item; ArrowUp/ArrowDown (wrapping), Home and End move
 * between items; Escape, Tab-away and outside clicks close it and return
 * focus to the More button. The menu exists only in the compact toolbar,
 * so its items are 44px targets on any pointer.
 */
import { useEffect, useRef, useState, type ReactElement, type ReactNode } from 'react'
import { Ellipsis } from 'lucide-react'
import { Icon } from '../Icon'
import { Pop, popover, Presence } from '../motion/presets'

export interface MoreItem {
  key: string
  label: string
  icon: ReactNode
  onSelect(): void
  /** Set for a toggle (menuitemcheckbox). */
  checked?: boolean
  disabled?: boolean
}

const ITEM = [
  'ov:flex ov:w-full ov:items-center ov:gap-2 ov:py-0 ov:px-3 ov:min-h-11 ov:text-left ov:text-base ov:text-text ov:bg-transparent ov:border-0 ov:cursor-pointer',
  'ov:hover:bg-hover ov:focus-visible:bg-hover ov:focus-visible:outline-2 ov:focus-visible:outline-solid ov:focus-visible:-outline-offset-2 ov:focus-visible:outline-accent',
  'ov:aria-checked:font-[600] ov:aria-disabled:text-muted ov:aria-disabled:cursor-default',
].join(' ')

export function MoreMenu({ items, buttonClass }: { items: readonly MoreItem[]; buttonClass: string }): ReactElement {
  const [open, setOpen] = useState(false)
  const button = useRef<HTMLButtonElement | null>(null)
  const menu = useRef<HTMLDivElement | null>(null)
  const close = () => { setOpen(false); button.current?.focus() }
  const entries = () => [...(menu.current?.querySelectorAll<HTMLElement>('[role^=menuitem]') ?? [])]

  useEffect(() => {
    if (!open) return
    entries()[0]?.focus()
    const outside = (e: Event) => {
      const t = e.target as Node
      if (!menu.current?.contains(t) && !button.current?.contains(t)) close()
    }
    // The owner document: a viewer rendered into a popup window listens there.
    const doc = button.current?.ownerDocument ?? document
    doc.addEventListener('pointerdown', outside, true)
    return () => doc.removeEventListener('pointerdown', outside, true)
  }, [open])

  const onKeyDown = (e: React.KeyboardEvent) => {
    const list = entries()
    const i = list.indexOf(document.activeElement as HTMLElement)
    const go = (n: number) => { e.preventDefault(); list[(n + list.length) % list.length]?.focus() }
    if (e.key === 'ArrowDown') go(i + 1)
    else if (e.key === 'ArrowUp') go(i - 1)
    else if (e.key === 'Home') go(0)
    else if (e.key === 'End') go(list.length - 1)
    else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); close() }
    else if (e.key === 'Tab') setOpen(false)
  }

  return (
    <div className="ov:relative ov:flex ov:shrink-0">
      <button
        ref={button}
        type="button"
        aria-label="More options"
        aria-haspopup="menu"
        aria-expanded={open}
        data-ov-control="more"
        onClick={() => setOpen((o) => !o)}
        className={buttonClass}
      >
        <Icon icon={Ellipsis} name="ellipsis" />
      </button>
      <Presence>{open && (
        <Pop
          key="menu"
          v={popover}
          ref={menu}
          role="menu"
          aria-label="More options"
          onKeyDown={onKeyDown}
          className="ov:absolute ov:right-0 ov:top-full ov:mt-1 ov:z-10 ov:min-w-44 ov:py-1 ov:bg-surface ov:text-text ov:border ov:border-border ov:rounded-panel ov:shadow-popover"
        >
          {items.map((item) => (
            <button
              key={item.key}
              type="button"
              role={item.checked === undefined ? 'menuitem' : 'menuitemcheckbox'}
              aria-checked={item.checked}
              // aria-disabled, not disabled: the item stays reachable by arrows
              aria-disabled={item.disabled || undefined}
              tabIndex={-1}
              onClick={() => { if (item.disabled) return; item.onSelect(); close() }}
              className={ITEM}
            >
              {item.icon}
              {item.label}
            </button>
          ))}
        </Pop>
      )}</Presence>
    </div>
  )
}
