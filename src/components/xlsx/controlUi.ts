import { useEffect, useRef, useState } from 'react'
import type { CacheIdentity } from '../../xlsx/view/contracts'
import { boundedSelectionText } from '../../xlsx/view/selection'
import { sameControlOwner, type ControlResult, type ControlsClient, type ControlSnapshot } from './controlsClient'

export interface ControlProps {
  owner: CacheIdentity
  client: ControlsClient
  open?: boolean
  invoker?: HTMLElement | null
  embedded?: boolean
  fallbackFocus?: () => void
  onClose: () => void
  onNotice: (message: string) => void
  onApplied: (snapshot: ControlSnapshot) => void
}
export const controlTarget = { minHeight: 44, minWidth: 44 } as const
export const CONTROL_CLS = 'ov:rounded-control ov:border ov:border-border ov:bg-field ov:text-text ov:enabled:hover:bg-hover ov:disabled:text-muted ov:disabled:opacity-50 ov:disabled:cursor-default ov:px-2 ov:py-1 ov:text-xs'
export const CONTROL_BTN_CLS = 'ov:rounded-control ov:border ov:border-border ov:bg-field ov:text-text ov:enabled:hover:bg-hover ov:disabled:text-muted ov:disabled:opacity-50 ov:disabled:cursor-default ov:px-2 ov:py-1 ov:text-xs ov:font-medium'
export const CONTROL_PRIMARY_BTN_CLS = 'ov:rounded-control ov:border ov:border-border ov:bg-accent ov:text-on-accent ov:enabled:hover:opacity-90 ov:disabled:opacity-50 ov:disabled:cursor-default ov:px-3 ov:py-1 ov:text-xs ov:font-semibold'
export const CHECKBOX_CLS = 'ov:rounded-control ov:border ov:border-border ov:bg-field ov:text-accent ov:p-0.5'
export function failureText(result: Exclude<ControlResult, { snapshot: ControlSnapshot }>): string {
  if (result.kind === 'refused') return `Operation refused: ${'message' in result.refusal ? result.refusal.message : result.refusal.reason}.`
  return result.kind === 'incomplete' ? 'Source data is incomplete; wait and retry.' : result.kind === 'stale' ? 'Source or view changed; reopen and retry.' : 'Operation cancelled.'
}
export function useControlLifetime(props: ControlProps) {
  const latest = useRef(props); latest.current = props
  const ownerKey = JSON.stringify(props.owner), previous = useRef(ownerKey), active = useRef(false), sequence = useRef(0)
  const controller = useRef<AbortController | null>(null), dispose = useRef<() => void>(() => {})
  const [invalid, setInvalid] = useState(false)
  useEffect(() => {
    if (previous.current !== ownerKey && props.open !== false && active.current) {
      active.current = false; controller.current?.abort(); dispose.current(); setInvalid(true)
      latest.current.onNotice('Source or view changed; reopen and retry.'); latest.current.onClose()
    } else { active.current = props.open !== false; setInvalid(false) }
    previous.current = ownerKey
    return () => { controller.current?.abort(); dispose.current() }
  }, [ownerKey, props.open])
  useEffect(() => () => { active.current = false; controller.current?.abort(); dispose.current() }, [])
  const begin = () => {
    controller.current?.abort()
    const abort = new AbortController(), id = ++sequence.current, owner = { ...latest.current.owner }
    controller.current = abort
    const current = () => active.current && id === sequence.current && !abort.signal.aborted && sameControlOwner(owner, latest.current.owner) && sameControlOwner(owner, latest.current.client.owner())
    const accepts = (result: ControlResult) => current() || (active.current && id === sequence.current && !abort.signal.aborted && 'snapshot' in result
      && sameControlOwner(result.snapshot.owner, latest.current.client.owner()) && (sameControlOwner(latest.current.owner, owner) || sameControlOwner(latest.current.owner, result.snapshot.owner)))
    return { signal: abort.signal, owner, current, accepts }
  }
  const cancel = (restore = true) => {
    active.current = false; sequence.current++; controller.current?.abort(); dispose.current()
    const p = latest.current
    p.onClose()
    if (restore && !p.embedded) {
      if (p.invoker?.isConnected && sameControlOwner(p.owner, p.client.owner())) p.invoker.focus()
      else p.fallbackFocus?.()
    }
  }
  const error = (message: unknown) => boundedSelectionText(message instanceof Error ? message.message : 'Operation unavailable; retry.')
  const result = (outcome: ControlResult, message: string) => {
    if (!('snapshot' in outcome)) { latest.current.onNotice(failureText(outcome)); return false }
    active.current = false; dispose.current()
    const p = latest.current
    p.onApplied(outcome.snapshot)
    if (sameControlOwner(outcome.snapshot.owner, p.client.owner())) p.onNotice(message)
    p.onClose()
    if (!p.embedded) {
      if (p.invoker?.isConnected && p.owner.docId === outcome.snapshot.owner.docId && p.owner.sheetId === outcome.snapshot.owner.sheetId) p.invoker.focus()
      else p.fallbackFocus?.()
    }
    return true
  }
  return { begin, cancel, result, error, invalid, dispose, ownerKey }
}
